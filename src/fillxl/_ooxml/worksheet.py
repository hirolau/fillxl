"""Surgical editing of one worksheet part.

Rows and cells are read out of the raw XML and written back as text.  A row
that is never modified is written back with its original bytes, and a sheet
that is never modified is not written back at all, so a fill cannot change
anything it did not mean to change.

The engine never inserts or removes rows (see ``IMPLEMENTATION-LOG.md``, D3):
growing a table writes cells into rows that were verified empty, which keeps
merges, data validations, drawings, hyperlinks, chart anchors, defined names
and array-formula references valid without touching them.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from attrs import define, field
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from fillxl._formulas import shift_relative
from fillxl._ooxml import xmltext

if TYPE_CHECKING:
    from collections.abc import Iterator

    from fillxl._ooxml.package import OoxmlPackage

_TAG = r"(?:[A-Za-z_][\w.\-]*:)?"
_CELL = re.compile(
    rf"<(?P<prefix>{_TAG})?c\b(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</{_TAG}c>)",
    re.S,
)
_ROW = re.compile(
    rf"<(?P<prefix>{_TAG})?row\b(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</{_TAG}row>)",
    re.S,
)
_SHEET_DATA = re.compile(
    rf"<(?P<prefix>{_TAG})?sheetData\b(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</{_TAG}sheetData>)",
    re.S,
)
_DIMENSION = re.compile(rf"<{xmltext.any_prefix('dimension')}\b[^>]*/>")
_CONDITIONAL_FORMATTING = xmltext.element_re("conditionalFormatting")
_FORMULA = xmltext.element_re("f")
_VALUE_ELEMENT = re.compile(
    rf"<{xmltext.any_prefix('v')}\b[^>]*?(?:/>|>(?P<text>.*?)</{xmltext.any_prefix('v')}>)",
    re.S,
)
_INLINE_STRING = re.compile(rf"<{xmltext.any_prefix('is')}\b")
_MERGE_CELL = xmltext.element_re("mergeCell")
_ROW_REFERENCE = re.compile(r"^\$?(?P<column>[A-Za-z]{1,3})\$?(?P<row>\d+)$")
_HAS_FORMULA = re.compile(rf"<{_TAG}f\b")


def column_index(letters: str) -> int:
    """1-based column number of the Excel column *letters*.

    Examples:
        >>> column_index("A")
        1
        >>> column_index("AB")
        28

    """
    index = 0
    for character in letters.upper():
        index = index * 26 + ord(character) - ord("A") + 1
    return index


def split_reference(reference: str) -> tuple[int, int]:
    """``(row, column)`` of a plain cell reference such as ``B12``.

    Examples:
        >>> split_reference("B12")
        (12, 2)
        >>> split_reference("$AA$3")
        (3, 27)

    """
    match = _ROW_REFERENCE.match(reference)
    if match is None:  # pragma: no cover - callers pass references from the file
        msg = f"Not a cell reference: {reference!r}"
        raise ValueError(msg)
    return int(match.group("row")), column_index(match.group("column"))


def cell_has_value(xml: str) -> bool:
    """Whether a ``<c>`` element holds something a user would call data.

    A cell that only carries a style has no value.  A cell with an *empty*
    ``<v/>`` has none either — which is exactly why an array range has to be
    treated as occupied separately: a spill whose result is blank stores an
    empty value.

    Examples:
        >>> cell_has_value('<c r="A1" s="1"/>')
        False
        >>> cell_has_value('<c r="A1"><v>3</v></c>')
        True
        >>> cell_has_value('<c r="A1" t="str"><v/></c>')
        False
        >>> cell_has_value('<c r="A1"><f>SUM(B1:B2)</f><v>3</v></c>')
        True
        >>> cell_has_value('<c r="A1" t="inlineStr"><is><t>x</t></is></c>')
        True

    """
    if _FORMULA.search(xml) is not None:
        return True
    if _INLINE_STRING.search(xml) is not None:
        return True
    match = _VALUE_ELEMENT.search(xml)
    return match is not None and bool((match.group("text") or "").strip())


def formula_text(xml: str) -> str | None:
    """The formula of a ``<c>`` element, with Excel's leading ``=``.

    Examples:
        >>> formula_text('<c r="A1"><f>SUM(B1:B2)</f><v>3</v></c>')
        '=SUM(B1:B2)'
        >>> formula_text('<c r="A1"><v>3</v></c>') is None
        True

    """
    match = _FORMULA.search(xml)
    if match is None:
        return None
    text = xmltext.element_text(match.group(0), "f")
    return None if text is None else "=" + text


def array_range(xml: str) -> str | None:
    """The ``ref`` of the first array formula in a ``<c>`` element.

    Attributes are read from the start tag, so their order in the file does
    not matter: Excel writes ``t`` first, other writers put ``ref`` first.

    Examples:
        >>> array_range('<c r="F8" cm="1"><f t="array" ref="F8:F12">_xlfn.SEQUENCE(5)</f></c>')
        'F8:F12'
        >>> array_range('<c r="F8" cm="1"><f ref="F8:F12" t="array">_xlfn.SEQUENCE(5)</f></c>')
        'F8:F12'
        >>> array_range('<c r="F8"><f>SUM(A1)</f></c>') is None
        True

    """
    tag = xmltext.start_tag(xml, "f")
    if tag is None or xmltext.attribute(tag, "t") != "array":
        return None
    return xmltext.attribute(tag, "ref")


def _is_array_cell(xml: str | None) -> bool:
    """Whether *xml* is a cell holding an array (spilled) formula."""
    return xml is not None and array_range(xml) is not None


def _plain_formula(xml: str, text: str) -> str:
    """*xml* with its ``<f>`` element rewritten as a plain formula holding *text*.

    The ``t``, ``si`` and ``ref`` attributes describe a shared group; a cell
    that is no longer shared must not keep them.
    """
    match = _FORMULA.search(xml)
    if match is None:  # pragma: no cover - only formula cells are unshared
        return xml
    element = match.group(0)
    opening = element[: element.find(">") + 1]
    name = re.match(rf"<({xmltext.any_prefix('f')})", opening)
    if name is None:  # pragma: no cover - the pattern found an <f> start tag
        return xml
    for attribute in ("t", "si", "ref"):
        opening = xmltext.remove_attribute(opening, attribute)
    if opening.endswith("/>"):
        opening = f"{opening[:-2].rstrip()}>"
    replacement = f"{opening}{xmltext.escape_text(text)}</{name.group(1)}>"
    return f"{xml[: match.start()]}{replacement}{xml[match.end() :]}"


@define
class _Row:
    """One ``<row>`` element, decomposed into its attributes and cells."""

    attributes: str
    prefix: str = ""
    cells: dict[int, str] = field(factory=dict)
    raw: str | None = None

    @property
    def changed(self) -> bool:
        """Whether the row has to be re-rendered."""
        return self.raw is None


class SheetPart:
    """A worksheet part that can be read and edited cell by cell."""

    def __init__(self, package: OoxmlPackage, part: str) -> None:
        """Parse the worksheet *part* of *package*.

        Args:
            package: The open package.
            part: Package part name, e.g. ``xl/worksheets/sheet1.xml``.

        """
        self.package = package
        self.part = part
        text = package.text(part)
        match = _SHEET_DATA.search(text)
        if match is None:
            msg = f"{part} has no <sheetData> element"
            raise ValueError(msg)
        self._head = text[: match.start()]
        self._prefix = match.group("prefix") or ""
        self._attributes = match.group("attrs") or ""
        self._tail = text[match.end() :]
        self._rows = self._parse_rows(match.group("body") or "")
        self._array_cache: list[CellRange] | None = None
        self._shared_master: dict[str, tuple[int, int, str]] = {}
        self._shared_members: dict[str, list[tuple[int, int]]] = {}
        self._cell_group: dict[tuple[int, int], str] = {}
        self._index_shared()
        self._changed = False

    # ── parsing ────────────────────────────────────────────────────────────

    def _parse_rows(self, body: str) -> dict[int, _Row]:
        rows: dict[int, _Row] = {}
        previous = 0
        for match in _ROW.finditer(body):
            attributes = match.group("attrs") or ""
            number = xmltext.attribute(f"<row{attributes}/>", "r")
            row_number = int(number) if number is not None and number.isdigit() else previous + 1
            previous = row_number
            cells: dict[int, str] = {}
            previous_column = 0
            for cell in _CELL.finditer(match.group("body") or ""):
                xml = cell.group(0)
                reference = xmltext.attribute(xml, "r")
                if reference is None:
                    # The reference is optional: a cell without one follows the
                    # previous cell of its row.
                    column = previous_column + 1
                    xml = xmltext.set_tag_attribute(
                        xml,
                        "r",
                        f"{get_column_letter(column)}{row_number}",
                    )
                else:
                    _, column = split_reference(reference)
                previous_column = column
                cells[column] = xml
            rows[row_number] = _Row(
                attributes=attributes,
                prefix=match.group("prefix") or "",
                cells=cells,
                raw=match.group(0),
            )
        return rows

    def _dirty(self) -> None:
        self._changed = True

    def _index_shared(self) -> None:
        """Index the shared-formula groups of the sheet.

        Excel stores a formula filled down or right once, on a master cell,
        plus an empty ``<f t="shared" si="…"/>`` on every other cell of the
        group.  A dependent has no text of its own, so without this index it
        looks like a cell that holds no formula at all.
        """
        for row, column, xml in self.iter_cells():
            tag = xmltext.start_tag(xml, "f")
            if tag is None or xmltext.attribute(tag, "t") != "shared":
                continue
            si = xmltext.attribute(tag, "si")
            if si is None:  # pragma: no cover - Excel always writes si
                continue
            self._cell_group[(row, column)] = si
            self._shared_members.setdefault(si, []).append((row, column))
            text = xmltext.element_text(xml, "f")
            if text is not None:
                self._shared_master.setdefault(si, (row, column, text))

    # ── cells ──────────────────────────────────────────────────────────────

    def cell(self, row: int, column: int) -> str | None:
        """The ``<c>`` XML of one cell, or ``None`` when the cell does not exist."""
        entry = self._rows.get(row)
        return None if entry is None else entry.cells.get(column)

    def style(self, row: int, column: int) -> str | None:
        """The ``s`` (style index) of a cell, or ``None`` for the default style."""
        xml = self.cell(row, column)
        return None if xml is None else xmltext.attribute(xml, "s")

    def reference(self, row: int, column: int) -> str:
        """The A1 reference of one cell, e.g. ``B12``."""
        return f"{get_column_letter(column)}{row}"

    def write_cell(self, row: int, column: int, xml: str) -> None:
        """Store the ``<c>`` element *xml* at *row*, *column*."""
        entry = self._rows.get(row)
        if entry is None:
            entry = _Row(attributes="", prefix=self._prefix)
            self._rows[row] = entry
        previous = entry.cells.get(column)
        entry.cells[column] = xml
        entry.raw = None
        if _is_array_cell(previous) or _is_array_cell(xml):
            # Overwriting or deleting an anchor changes which cells are
            # reserved, so checking only the new XML is not enough.
            self._array_cache = None
        self._dirty()

    def delete_cell(self, row: int, column: int) -> None:
        """Remove the cell at *row*, *column* from the sheet."""
        entry = self._rows.get(row)
        if entry is None or column not in entry.cells:
            return
        if _is_array_cell(entry.cells.pop(column)):
            self._array_cache = None
        entry.raw = None
        self._dirty()

    def set_style(self, row: int, column: int, style: str) -> None:
        """Give the cell at *row*, *column* the style index *style*."""
        xml = self.cell(row, column)
        if xml is None:
            self.write_cell(row, column, f'<c r="{self.reference(row, column)}" s="{style}"/>')
            return
        self.write_cell(row, column, xmltext.set_tag_attribute(xml, "s", style))

    def has_value(self, row: int, column: int) -> bool:
        """Whether the cell at *row*, *column* holds data (not just a style)."""
        xml = self.cell(row, column)
        return xml is not None and cell_has_value(xml)

    def iter_cells(self) -> Iterator[tuple[int, int, str]]:
        """Yield ``(row, column, xml)`` for every cell in the sheet."""
        for number in sorted(self._rows):
            row = self._rows[number]
            for column in sorted(row.cells):
                yield number, column, row.cells[column]

    def row_numbers(self) -> list[int]:
        """The row numbers present in the sheet, ascending."""
        return sorted(self._rows)

    # ── sheet features ─────────────────────────────────────────────────────

    def merged_ranges(self) -> list[str]:
        """Every merged range of the sheet."""
        return [
            xmltext.unescape(xmltext.attribute(match.group(0), "ref") or "")
            for match in _MERGE_CELL.finditer(self._head + self._tail)
        ]

    def array_ranges(self) -> list[CellRange]:
        """The ranges reserved by array (spilled) formulas.

        Excel needs these treated as occupied: a spill can hold blank results,
        which store as an empty value and look like an empty cell.

        The list is built once and cached; every cell write asks for it, so
        re-scanning the sheet per write made filling a table quadratic.
        """
        if self._array_cache is None:
            self._array_cache = self._collect_array_ranges()
        return self._array_cache

    def _collect_array_ranges(self) -> list[CellRange]:
        ranges: list[CellRange] = []
        for _, _, xml in self.iter_cells():
            reference = array_range(xml)
            if reference is None:
                continue
            try:
                ranges.append(CellRange(reference))
            except ValueError:  # pragma: no cover - Excel writes valid ranges
                continue
        return ranges

    def array_range_at(self, row: int, column: int) -> str | None:
        """The array range that covers cell *row*, *column*, if any."""
        for cell_range in self.array_ranges():
            if (
                cell_range.min_row <= row <= cell_range.max_row
                and cell_range.min_col <= column <= cell_range.max_col
            ):
                return cell_range.coord
        return None

    def formula_cells(self) -> Iterator[tuple[int, int, str]]:
        """Yield ``(row, column, formula)`` for every cell that holds a formula."""
        for number in sorted(self._rows):
            for column in sorted(self._rows[number].cells):
                text = self.formula(number, column)
                if text is not None:
                    yield number, column, text

    def formula(self, row: int, column: int) -> str | None:
        """The formula of a cell, with Excel's leading ``=``.

        A cell that belongs to a shared-formula group gets the master's text
        moved to its own position, which is the formula Excel displays.
        """
        xml = self.cell(row, column)
        if xml is None:
            return None
        si = self._cell_group.get((row, column))
        if si is None:
            return formula_text(xml)
        master = self._shared_master.get(si)
        if master is None:  # pragma: no cover - a group without a master is corrupt
            return None
        master_row, master_column, text = master
        return shift_relative(
            f"={text}",
            rows=row - master_row,
            columns=column - master_column,
        )

    def unshare(self, si: str) -> None:
        """Replace the shared-formula group *si* with one formula per cell.

        Editing one member of a group in place would either orphan the others
        or leave a master whose ``ref`` no longer covers them.
        """
        master = self._shared_master.pop(si, None)
        members = self._shared_members.pop(si, [])
        for cell in members:
            self._cell_group.pop(cell, None)
        if master is None:  # pragma: no cover - groups always have a master
            return
        master_row, master_column, text = master
        for row, column in members:
            xml = self.cell(row, column)
            if xml is None:  # pragma: no cover - members always exist
                continue
            moved = shift_relative(
                text,
                rows=row - master_row,
                columns=column - master_column,
            )
            self.write_cell(row, column, _plain_formula(xml, moved))

    def unshare_range(
        self,
        first_row: int,
        last_row: int,
        first_col: int,
        last_col: int,
    ) -> None:
        """Unshare every group with at least one member inside the rectangle."""
        doomed = {
            si
            for (row, column), si in self._cell_group.items()
            if first_row <= row <= last_row and first_col <= column <= last_col
        }
        for si in sorted(doomed):
            self.unshare(si)

    def set_formula(self, row: int, column: int, formula: str) -> None:
        """Replace the formula of an existing cell, keeping everything else."""
        si = self._cell_group.get((row, column))
        if si is not None:
            self.unshare(si)
        xml = self.cell(row, column)
        if xml is None:  # pragma: no cover - callers rewrite existing formulas
            return
        text = formula[1:] if formula.startswith("=") else formula
        self.write_cell(row, column, xmltext.replace_element_text(xml, "f", text))

    def extend_conditional_formatting(
        self,
        *,
        table_start_column: int,
        table_end_column: int,
        table_top_row: int,
        first_data_row: int,
        old_body_last_row: int,
        table_last_row: int,
    ) -> list[tuple[str, str]]:
        """Resize conditional-formatting areas that track the table's height.

        See :func:`fillxl._tables.cf_range_tracks_table` for the rule.  The
        element's ``sqref`` is the only thing rewritten inside it.

        Returns:
            ``(old sqref, new sqref)`` for every changed block.

        """
        from fillxl._tables import cf_range_tracks_table

        changes: list[tuple[str, str]] = []

        def replace(match: re.Match[str]) -> str:
            element = match.group(0)
            sqref = xmltext.attribute(element, "sqref")
            # x14 blocks hold their range in a child <xm:sqref> element.
            in_child = False
            if not sqref:
                sqref = xmltext.element_text(element, "sqref")
                in_child = sqref is not None
            if not sqref:  # pragma: no cover - sqref is required
                return element
            areas = sqref.split()
            resized: list[str] = []
            extended = False
            for area in areas:
                try:
                    cell_range = CellRange(area)
                except ValueError:  # pragma: no cover - Excel writes valid ranges
                    resized.append(area)
                    continue
                if cf_range_tracks_table(
                    cell_range,
                    table_start_column=table_start_column,
                    table_end_column=table_end_column,
                    table_top_row=table_top_row,
                    first_data_row=first_data_row,
                    old_body_last_row=old_body_last_row,
                    table_last_row=table_last_row,
                ):
                    resized.append(
                        CellRange(
                            min_col=cell_range.min_col,
                            min_row=cell_range.min_row,
                            max_col=cell_range.max_col,
                            max_row=table_last_row,
                        ).coord,
                    )
                    extended = True
                else:
                    resized.append(area)
            if not extended:
                return element
            new_sqref = " ".join(resized)
            changes.append((sqref, new_sqref))
            self._dirty()
            if in_child:
                return xmltext.replace_element_text(element, "sqref", new_sqref)
            return xmltext.set_tag_attribute(element, "sqref", new_sqref)

        self._tail = _CONDITIONAL_FORMATTING.sub(replace, self._tail)
        return changes

    def conditional_formatting_sqrefs(self) -> list[str]:
        """The ``sqref`` of every conditional-formatting block, in file order."""
        return [
            xmltext.attribute(match.group(0), "sqref") or ""
            for match in _CONDITIONAL_FORMATTING.finditer(self._head + self._tail)
        ]

    # ── writing the part back ──────────────────────────────────────────────

    def to_xml(self) -> str:
        """The current XML of the whole worksheet part."""
        body = "".join(
            row.raw if row.raw is not None else self._render_row(number, row)
            for number, row in sorted(self._rows.items())
        )
        opening = f"<{self._prefix}sheetData{self._attributes}>"
        closing = f"</{self._prefix}sheetData>"
        return f"{self._head}{opening}{body}{closing}{self._tail}"

    def save(self) -> None:
        """Write the part back when it changed; leave it untouched otherwise."""
        if self._changed:
            self.package.set_text(self.part, self.to_xml())
            self._changed = False

    def _render_row(self, number: int, row: _Row) -> str:
        attributes = xmltext.remove_attribute(row.attributes, "spans")
        tag = xmltext.set_tag_attribute(f"<row{attributes}/>", "r", str(number))
        attributes = tag[len("<row") : -2]
        body = "".join(row.cells[column] for column in sorted(row.cells))
        if not body:
            return f"<{row.prefix}row{attributes}/>"
        return f"<{row.prefix}row{attributes}>{body}</{row.prefix}row>"

    def expand_dimension(self, row: int) -> None:
        """Grow the ``<dimension>`` reference so it still covers *row*."""
        match = _DIMENSION.search(self._head)
        if match is None:
            return
        element = match.group(0)
        reference = xmltext.attribute(element, "ref")
        if reference is None:  # pragma: no cover - dimension always has a ref
            return
        bounds = CellRange(reference)
        if row <= bounds.max_row:
            return
        bounds = CellRange(
            min_col=bounds.min_col,
            min_row=bounds.min_row,
            max_col=bounds.max_col,
            max_row=row,
        )
        self._head = self._head[: match.start()] + xmltext.set_tag_attribute(
            element,
            "ref",
            bounds.coord,
        ) + self._head[match.end() :]
        self._dirty()


class SheetStore:
    """One :class:`SheetPart` per part, so a sheet is never parsed twice.

    Editing a sheet twice through two instances would lose the first set of
    changes when the second one is written back; the store makes that
    impossible.
    """

    def __init__(self, package: OoxmlPackage) -> None:
        """Bind the store to *package*."""
        self.package = package
        self._sheets: dict[str, SheetPart] = {}

    def get(self, part: str) -> SheetPart:
        """The editor for *part*, created on first use."""
        if part not in self._sheets:
            self._sheets[part] = SheetPart(self.package, part)
        return self._sheets[part]

    def formula_parts(self, parts: list[str]) -> list[SheetPart]:
        """The editors of those *parts* that contain at least one formula."""
        return [
            self.get(part)
            for part in parts
            if _HAS_FORMULA.search(self.package.text(part)) is not None
        ]

    def save_all(self) -> None:
        """Write back every sheet that changed."""
        for sheet in self._sheets.values():
            sheet.save()

    def forget(self, part: str) -> None:
        """Drop the cached editor of *part*, so the next :meth:`get` re-reads it.

        An operation that changes a part directly through the package (hiding a
        sheet, deleting one, activating one) leaves a cached editor holding the
        old text; saving it back would undo the change.
        """
        self._sheets.pop(part, None)
