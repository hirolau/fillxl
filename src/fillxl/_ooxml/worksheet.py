"""Surgical editing of one worksheet part.

Rows and cells are read out of the raw XML and written back as text.  A row
that is never modified is written back with its original bytes, and a sheet
that is never modified is not written back at all, so a fill cannot change
anything it did not mean to change.

The engine never inserts or removes rows: growing a table writes cells into
rows that were verified empty, which keeps merges, data validations, drawings,
hyperlinks, chart anchors, defined names and array-formula references valid
without touching them.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING

from attrs import define, field
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from fillxl._formulas import shift_relative
from fillxl._ooxml import xmltext

if TYPE_CHECKING:
    from collections.abc import Callable, Iterator

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
_HIDDEN_ATTRIBUTE = re.compile(r'\s+hidden\s*=\s*"[^"]*"')
_SHEET_VIEWS = xmltext.element_re("sheetViews")
_SHEET_VIEW = xmltext.element_re("sheetView")
# ``\b`` keeps this from matching ``<dataValidations>``, whose container is not
# a validation itself.
DATA_VALIDATION = xmltext.element_re("dataValidation")

# Where a missing <sheetViews> has to be inserted (ECMA-376 element order).
WORKSHEET_VIEW_ANCHORS = ("sheetFormatPr", "cols", "sheetData")
CHARTSHEET_VIEW_ANCHORS = (
    "sheetProtection",
    "customSheetViews",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "drawing",
)

logger = logging.getLogger(__name__)


def root_element_end(text: str) -> int:
    """The index just past the root element's start tag.

    ``text.find(">")`` is not that index: it stops at the end of the XML
    declaration, or of a comment that precedes the root.
    """
    for match in re.finditer(r"<[^>]*>", text):
        if match.group(0)[1:2] in ("?", "!"):
            continue
        return match.end()
    return len(text)  # pragma: no cover - a part always has a root element


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


def clear_cached_value(xml: str) -> str:
    """*xml* without the cached result of its formula, and without its type.

    A cached ``<v>`` is what every reader that cannot calculate sees: the
    template's number, presented as the filled workbook's number.  Dropping it
    makes such a reader see an obviously missing value instead of a plausibly
    wrong one; Excel recalculates on open (the engine sets ``fullCalcOnLoad``).

    Examples:
        >>> clear_cached_value('<c r="B7"><f>SUM(B1:B6)</f><v>42</v></c>')
        '<c r="B7"><f>SUM(B1:B6)</f></c>'
        >>> clear_cached_value('<c r="B7" t="str"><f t="array" ref="B7:B9">X</f><v>x</v></c>')
        '<c r="B7"><f t="array" ref="B7:B9">X</f></c>'
        >>> clear_cached_value('<c r="B7"><v>42</v></c>')
        '<c r="B7"><v>42</v></c>'

    """
    if _FORMULA.search(xml) is None:
        return xml
    return strip_value(xml)


def strip_value(xml: str) -> str:
    """*xml* without a ``<v>`` element that holds a result, and without its type.

    An *empty* ``<v/>`` is not a result — openpyxl writes one on every formula
    cell — so a cell that only carries one is returned unchanged, which keeps
    untouched formula sheets byte-identical.

    Examples:
        >>> strip_value('<c r="F9" t="str"><v>2</v></c>')
        '<c r="F9"></c>'
        >>> strip_value('<c r="F9"><f>1+1</f><v/></c>')
        '<c r="F9"><f>1+1</f><v/></c>'
        >>> strip_value('<c r="F9" t="inlineStr"><is><t>x</t></is></c>')
        '<c r="F9" t="inlineStr"><is><t>x</t></is></c>'

    """
    if not any((match.group("text") or "").strip() for match in _VALUE_ELEMENT.finditer(xml)):
        return xml
    stripped = _VALUE_ELEMENT.sub("", xml)
    opening_end = stripped.find(">")
    if opening_end < 0:  # pragma: no cover - a <c> element always has a start tag
        return stripped
    return f"{xmltext.remove_attribute(stripped[:opening_end], 't')}{stripped[opening_end:]}"


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

    @property
    def tail(self) -> str:
        """Everything after ``</sheetData>``: validations, formatting, extensions."""
        return self._tail

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
        """Store the ``<c>`` element *xml* at *row*, *column*.

        A shared-formula stub is refused: every fill path unshares a group
        before it edits one of its cells, so a stub arriving here was built
        from a snapshot the engine has already superseded.
        """
        self._refuse_shared(xml, row, column)
        self._unshare_around(row, column)
        self._replace_cell_xml(row, column, xml)

    def _refuse_shared(self, xml: str, row: int, column: int) -> None:
        """Raise when *xml* is a shared-formula stub about to overwrite a cell."""
        tag = xmltext.start_tag(xml, "f")
        if tag is None or xmltext.attribute(tag, "t") != "shared":
            return
        msg = (
            f"Cannot write a shared formula to {self.reference(row, column)}: "
            "a shared-formula group has to be unshared before one of its cells is edited."
        )
        raise RuntimeError(msg)

    def _replace_cell_xml(self, row: int, column: int, xml: str) -> None:
        """Store *xml* at *row*, *column* without touching shared-formula groups.

        Removing a cached result never changes which cells belong to a group,
        so a caller that only rewrites a value must not unshare — doing so
        would rewrite XML built from the pre-unshare snapshot.
        """
        entry = self._rows.get(row)
        if entry is None:
            entry = _Row(attributes="", prefix=self._prefix)
            self._rows[row] = entry
        previous = entry.cells.get(column)
        entry.cells[column] = xml
        entry.raw = None
        if _is_array_cell(previous) != _is_array_cell(xml):
            self._array_cache = None
        self._dirty()

    def delete_cell(self, row: int, column: int) -> None:
        """Remove the cell at *row*, *column* from the sheet."""
        self._unshare_around(row, column)
        entry = self._rows.get(row)
        if entry is None or column not in entry.cells:
            return
        if _is_array_cell(entry.cells.pop(column)):
            self._array_cache = None
        entry.raw = None
        self._dirty()

    def _unshare_around(self, row: int, column: int) -> None:
        """Expand the shared-formula group of the cell about to be overwritten.

        A shared group has one master and empty dependents, so replacing any
        member in place would leave the others either without a master or
        pointing at a ``ref`` that no longer covers them.  ``unshare`` drops
        the group from the index before it rewrites the members, so the
        ``write_cell`` calls it makes do not come back here.
        """
        si = self._cell_group.get((row, column))
        if si is not None:
            self.unshare(si)

    def set_style(self, row: int, column: int, style: str) -> None:
        """Give the cell at *row*, *column* the style index *style*."""
        # Unsharing replaces the cell's XML, so read it afterwards.
        self._unshare_around(row, column)
        xml = self.cell(row, column)
        if xml is None:
            self.write_cell(row, column, f'<c r="{self.reference(row, column)}" s="{style}"/>')
            return
        self.write_cell(row, column, xmltext.set_tag_attribute(xml, "s", style))

    def has_value(self, row: int, column: int) -> bool:
        """Whether the cell at *row*, *column* holds data (not just a style)."""
        xml = self.cell(row, column)
        return xml is not None and cell_has_value(xml)

    def has_formula(self) -> bool:
        """Whether the sheet holds at least one formula cell."""
        return any(_HAS_FORMULA.search(xml) for _, _, xml in self.iter_cells())

    def iter_cells(self) -> Iterator[tuple[int, int, str]]:
        """Yield ``(row, column, xml)`` for every cell in the sheet."""
        for number in sorted(self._rows):
            row = self._rows[number]
            for column in sorted(row.cells):
                yield number, column, row.cells[column]

    def row_numbers(self) -> list[int]:
        """The row numbers present in the sheet, ascending."""
        return sorted(self._rows)

    def unhide_rows(self, first_row: int, last_row: int) -> list[int]:
        """Clear the ``hidden`` flag of rows *first_row* through *last_row*.

        A template row can be hidden by the author or by a filter that was
        applied when it was saved.  Data written into it would be invisible,
        and ``SUBTOTAL(109, …)`` would leave it out of a totals row, so the
        rows a fill writes to are always shown.

        Returns:
            The row numbers that were hidden and are now shown.

        """
        shown: list[int] = []
        for number in sorted(self._rows):
            if not first_row <= number <= last_row:
                continue
            row = self._rows[number]
            if xmltext.attribute(f"<row{row.attributes}/>", "hidden") is None:
                continue
            row.attributes = _HIDDEN_ATTRIBUTE.sub("", row.attributes)
            row.raw = None
            self._dirty()
            shown.append(number)
        return shown

    # ── sheet features ─────────────────────────────────────────────────────

    def resize_validation_lists(self, resize: Callable[[str], str | None]) -> int:
        """Rewrite the list source of every data validation of the sheet.

        Both spellings are handled: the classic ``<formula1>text</formula1>``
        and the x14 ``<x14:formula1><xm:f>text</xm:f></x14:formula1>`` that
        Excel writes inside ``extLst``.  *resize* returns the new text for a
        source, or ``None`` to leave it alone.

        Returns:
            How many sources were rewritten.

        """
        changed = 0

        def replace(match: re.Match[str]) -> str:
            nonlocal changed
            element = match.group(0)
            for name in ("formula1", "formula2"):
                block_match = xmltext.element_re(name).search(element)
                if block_match is None:
                    continue
                block = block_match.group(0)
                child = xmltext.element_re("f").search(block)
                holder = "f" if child is not None else name
                source = xmltext.element_text(block, holder)
                if source is None:
                    continue
                resized = resize(source)
                if resized is None:
                    continue
                block = xmltext.replace_element_text(block, holder, resized)
                element = f"{element[: block_match.start()]}{block}{element[block_match.end() :]}"
                changed += 1
            return element

        updated = DATA_VALIDATION.sub(replace, self._tail)
        if updated != self._tail:
            self._tail = updated
            self._dirty()
        return changed

    def set_tab_selected(self, selected: bool) -> None:
        """Set ``tabSelected`` on this sheet's ``<sheetView>``.

        The workbook says which tab is *active*; each sheet also says whether
        it is *selected*, and a selection that disagrees leaves two tabs
        highlighted in Excel.  ``<sheetViews>`` always precedes
        ``<sheetData>``, so this only edits the head of the part.
        """
        value = "1" if selected else "0"
        view = _SHEET_VIEW.search(self._head)
        if view is not None:
            updated = xmltext.set_tag_attribute(view.group(0), "tabSelected", value)
            self._head = f"{self._head[: view.start()]}{updated}{self._head[view.end() :]}"
            self._dirty()
            return
        if not selected:
            return
        # A sheet can be saved without a <sheetView>; insert one where the
        # schema expects it, before the row defaults, the columns and the data.
        new_view = '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>'
        for name in WORKSHEET_VIEW_ANCHORS:
            anchor = xmltext.start_tag_re(name).search(self._head)
            if anchor is not None:
                self._head = (
                    f"{self._head[: anchor.start()]}{new_view}{self._head[anchor.start() :]}"
                )
                self._dirty()
                return
        views = _SHEET_VIEWS.search(self._head)
        if views is not None:  # pragma: no cover - a sheetViews block without a view
            self._head = f"{self._head[: views.start()]}{new_view}{self._head[views.end() :]}"
            self._dirty()
            return
        root_end = root_element_end(self._head)
        self._head = f"{self._head[:root_end]}{new_view}{self._head[root_end:]}"
        self._dirty()

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
        """Yield ``(row, column, formula)`` for every cell that holds a formula.

        A shared member's text comes from ``shift_relative``, which tokenises
        the master's formula.  One formula Excel wrote but openpyxl cannot
        read is skipped rather than failing the whole fill; a group that has
        to be rewritten raises later, in :meth:`unshare`.
        """
        for number in sorted(self._rows):
            for column in sorted(self._rows[number].cells):
                try:
                    text = self.formula(number, column)
                except TokenizerError:
                    logger.debug(
                        "Cannot read the formula in %s; leaving it alone",
                        self.reference(number, column),
                    )
                    continue
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
            try:
                moved = shift_relative(
                    f"={text}",
                    rows=row - master_row,
                    columns=column - master_column,
                )
            except TokenizerError as exc:
                msg = (
                    f"Cannot rewrite the shared formula in {self.reference(row, column)}: "
                    f"the fillxl engine cannot read its master formula {text!r}. "
                    "Use engine='win32com' to fill this template."
                )
                raise ValueError(msg) from exc
            self.write_cell(row, column, _plain_formula(xml, moved.removeprefix("=")))

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
        """Replace the formula of an existing cell, keeping everything else.

        The template's cached result goes with it: it describes a formula that
        no longer exists.
        """
        si = self._cell_group.get((row, column))
        if si is not None:
            self.unshare(si)
        xml = self.cell(row, column)
        if xml is None:  # pragma: no cover - callers rewrite existing formulas
            return
        text = formula[1:] if formula.startswith("=") else formula
        self.write_cell(
            row,
            column,
            clear_cached_value(xmltext.replace_element_text(xml, "f", text)),
        )

    def clear_stale_results(self) -> int:
        """Drop every cached formula result of the sheet, spills included.

        The engine cannot calculate, so a cached value on this sheet would be
        the template's number shown as the filled report's number.  Removing it
        means a reader that cannot recalculate sees a blank rather than a
        plausible lie, which is what the openpyxl engine produces too.

        A spilled array keeps its anchor formula and its ``ref``; only the
        results below the anchor (cells that hold nothing but a ``<v>``) are
        cleared, and Excel refills them when it recalculates.

        Returns:
            How many cells were changed.

        """
        cleared = 0
        for row, column, xml in list(self.iter_cells()):
            stripped = clear_cached_value(xml)
            if stripped != xml:
                self._replace_cell_xml(row, column, stripped)
                cleared += 1
        for cell_range in self.array_ranges():
            cleared += self._clear_spill_results(cell_range)
        return cleared

    def _clear_spill_results(self, cell_range: CellRange) -> int:
        """Clear the non-anchor cells of one spilled array range."""
        cleared = 0
        for row in range(cell_range.min_row, cell_range.max_row + 1):
            for column in range(cell_range.min_col, cell_range.max_col + 1):
                xml = self.cell(row, column)
                if xml is None or _FORMULA.search(xml) is not None:
                    # The anchor holds the formula; a mid-range formula means
                    # the range is not a simple spill.
                    continue
                stripped = strip_value(xml)
                if stripped != xml:
                    self._replace_cell_xml(row, column, stripped)
                    cleared += 1
        return cleared

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

        See :func:`fillxl._tables.range_tracks_table` for the rule.  The
        element's ``sqref`` is the only thing rewritten inside it.

        Returns:
            ``(old sqref, new sqref)`` for every changed block.

        """
        from fillxl._tables import range_tracks_table

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
                if range_tracks_table(
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
            self._assert_shared_formulas_intact()
            self.package.set_text(self.part, self.to_xml())
            self._changed = False

    def _assert_shared_formulas_intact(self) -> None:
        """Refuse to save a sheet whose shared-formula groups lost their master.

        A stub with no covering master still parses, so nothing else would
        catch it: Excel reports "Removed Records: Formula" and silently drops
        the cell's formula when the file is opened.
        """
        masters: dict[str, tuple[int, int, CellRange]] = {}
        for row, column, xml in self.iter_cells():
            tag = xmltext.start_tag(xml, "f")
            if tag is None or xmltext.attribute(tag, "t") != "shared":
                continue
            si = xmltext.attribute(tag, "si")
            ref = xmltext.attribute(tag, "ref")
            if si is None or ref is None or xmltext.element_text(xml, "f") is None:
                continue
            masters.setdefault(si, (row, column, CellRange(ref)))
        for row, column, xml in self.iter_cells():
            tag = xmltext.start_tag(xml, "f")
            if tag is None or xmltext.attribute(tag, "t") != "shared":
                continue
            si = xmltext.attribute(tag, "si")
            if si is None or xmltext.element_text(xml, "f") is not None:
                continue
            master = masters.get(si)
            if (
                master is not None
                and master[2].min_row <= row <= master[2].max_row
                and master[2].min_col <= column <= master[2].max_col
            ):
                continue
            msg = (
                f"The template part {self.part} could not be edited safely: cell "
                f"{self.reference(row, column)} belongs to shared formula {si}, which has "
                "no covering master formula. Save the template in Excel again, or use "
                "engine='win32com'. Please report this with the template."
            )
            raise RuntimeError(msg)

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
        self._head = (
            self._head[: match.start()]
            + xmltext.set_tag_attribute(
                element,
                "ref",
                bounds.coord,
            )
            + self._head[match.end() :]
        )
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
        """The editors of those *parts* that contain at least one formula.

        A cached editor is asked directly, so a part edited during this session
        is judged by its current cells; an editor that has not been needed yet
        is not created for a sheet that holds no formula.
        """
        found: list[SheetPart] = []
        for part in parts:
            cached = self._sheets.get(part)
            if cached is not None:
                has_formula = cached.has_formula()
            else:
                has_formula = _HAS_FORMULA.search(self.package.text(part)) is not None
            if has_formula:
                found.append(self.get(part))
        return found

    def save_all(self) -> None:
        """Write back every sheet that changed.

        Raises:
            RuntimeError: If an editor would write a part the package no longer
                has, which would resurrect a sheet that was deleted.

        """
        for part, sheet in self._sheets.items():
            if not self.package.has(part):
                msg = (
                    f"Cannot write {part}: the part was removed from the workbook. "
                    "Sheets removed by delete_sheets have to be discarded from the store."
                )
                raise RuntimeError(msg)
            sheet.save()

    def discard(self, part: str) -> None:
        """Drop the cached editor of *part*, so its edits are never written back.

        Only a part that the package dropped needs this: a deleted sheet must
        not come back because an editor still held the text it had before.
        """
        self._sheets.pop(part, None)
