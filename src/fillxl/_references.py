"""Ranges that point at a table body without being a cell formula.

A fill resizes the charts and the A1 formulas that track a table, because both
are readable and rewritable.  A pivot cache source, a defined name and a
data-validation list hold the same rectangle in places no engine edits, so
after a fill they still cover the template's rows.  Those are worth a warning:
a pivot marked ``refreshOnLoad`` refreshes onto the old range and shows a total
that quietly leaves the filled rows out.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, NamedTuple

from openpyxl.worksheet.cell_range import CellRange

from ._charts import split_formula_arguments
from ._formulas import _replace_rows, sole_reference
from ._plan import TableGeometry
from ._warnings import UntrackedRangeWarning, warn_user

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping

    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet


BUILT_IN_NAME_PREFIX = "_xlnm."
"""Prefix of the names Excel owns (``Print_Area``, ``_FilterDatabase``, …)."""


class RangeSource(NamedTuple):
    """One reference to a worksheet rectangle, and what holds it.

    Attributes:
        description: Where the reference lives, for the warning — e.g.
            ``"the defined name 'SalesData'"``.
        sheet_title: Title of the sheet the reference points at.
        reference: The A1 rectangle, optionally with ``$`` markers.

    """

    description: str
    sheet_title: str
    reference: str


def warn_about_untracked_ranges(
    sources: Iterable[RangeSource],
    *,
    table_name: str,
    sheet_title: str,
    geometry: TableGeometry,
    new_body_last_row: int,
) -> None:
    """Warn about references that still cover the table's old rows.

    A reference counts when it pointed at exactly the old data body, or at the
    whole old table (a pivot source usually includes the header row), and
    stays inside the table's columns.

    Args:
        sources: The references found in the workbook.
        table_name: Name of the table that changed size.
        sheet_title: Title of the sheet holding the table.
        geometry: Where the table was before the fill.
        new_body_last_row: Last data row of the table after the fill.

    """
    if new_body_last_row == geometry.body_last_row:
        return
    old_table = (geometry.top_row, geometry.last_row)
    old_body = (geometry.first_data_row, geometry.body_last_row)
    for source in sources:
        if not _covers(
            source,
            expected_sheet=sheet_title,
            rows=(old_body, old_table),
            first_col=geometry.first_col,
            last_col=geometry.last_col,
        ):
            continue
        warn_user(
            f"Table {table_name!r} was resized, but {source.description} still covers "
            f"{_display_sheet(source.sheet_title)}!{source.reference} — the table's old "
            "rows. The rows the fill wrote are missing from it, so a pivot that refreshes "
            "on load or a name that is read back will use the old range. Resize that "
            "reference in the template to match the table.",
            UntrackedRangeWarning,
        )


def resized_reference(
    reference: str,
    *,
    sheet_matches: bool,
    geometry: TableGeometry,
    new_body_last_row: int,
) -> str | None:
    """*reference* resized to follow the table, or ``None`` when it does not track it.

    The rule is the one :func:`_covers` applies to the warning: a range over
    exactly the table's old data body becomes a range over its new body, and a
    range over exactly the old whole table keeps its header row and gains the
    totals row, if any.  Either way the range has to stay inside the table's
    columns.  ``$`` markers are kept as they were.

    Examples:
        >>> g = TableGeometry(
        ...     top_row=1, last_row=2, first_col=1, last_col=2, header_rows=1, totals_rows=0
        ... )
        >>> resized_reference("$A$2:$B$2", sheet_matches=True, geometry=g, new_body_last_row=4)
        '$A$2:$B$4'
        >>> resized_reference("$A$1:$B$2", sheet_matches=True, geometry=g, new_body_last_row=4)
        '$A$1:$B$4'
        >>> resized_reference("$A$1:$B$2", sheet_matches=False, geometry=g, new_body_last_row=4)
        >>> resized_reference("$A$2:$B$3", sheet_matches=True, geometry=g, new_body_last_row=4)
        >>> resized_reference("$C$2:$C$2", sheet_matches=True, geometry=g, new_body_last_row=4)

    """
    if not sheet_matches:
        return None
    try:
        bounds = CellRange(reference.replace("$", ""))
    except (ValueError, TypeError):
        return None
    if not geometry.first_col <= bounds.min_col or bounds.max_col > geometry.last_col:
        return None
    rows = (bounds.min_row, bounds.max_row)
    if rows == (geometry.first_data_row, geometry.body_last_row):
        return _replace_rows(reference, geometry.first_data_row, new_body_last_row)
    if rows == (geometry.top_row, geometry.last_row):
        return _replace_rows(
            reference,
            geometry.top_row,
            new_body_last_row + geometry.totals_rows,
        )
    return None


def resize_reference_text(
    text: str,
    *,
    table_sheet: str,
    geometry: TableGeometry,
    new_body_last_row: int,
    default_sheet: str | None = None,
) -> str | None:
    """*text* with every plain range that tracks the table resized, or ``None``.

    A defined name and a validation list are written as a bare reference, and
    a name may hold several comma-separated ones.  *default_sheet* is the sheet
    an unqualified reference belongs to — a validation's own sheet, for
    instance; a defined name always qualifies its ranges.

    Examples:
        >>> geometry = TableGeometry(
        ...     top_row=1, last_row=2, first_col=1, last_col=2, header_rows=1, totals_rows=0
        ... )
        >>> resize_reference_text(
        ...     "Data!$A$2:$B$2", table_sheet="Data", geometry=geometry, new_body_last_row=4
        ... )
        'Data!$A$2:$B$4'
        >>> resize_reference_text(
        ...     "$A$2:$B$2", table_sheet="Data", geometry=geometry,
        ...     new_body_last_row=4, default_sheet="Data",
        ... )
        '$A$2:$B$4'
        >>> resize_reference_text(
        ...     "$A$2:$B$2", table_sheet="Data", geometry=geometry, new_body_last_row=4
        ... ) is None
        True
        >>> resize_reference_text(
        ...     '"a,b"', table_sheet="Data", geometry=geometry, new_body_last_row=4
        ... ) is None
        True

    """
    parts = split_formula_arguments(text.strip().removeprefix("="))
    resized: list[str] = []
    changed = False
    for part in parts:
        parsed = sole_reference(part)
        if parsed is None:
            resized.append(part)
            continue
        sheet_title, reference = parsed
        owner = sheet_title if sheet_title is not None else default_sheet
        new_reference = (
            None
            if owner is None
            else resized_reference(
                reference,
                sheet_matches=_sheet_key(owner) == _sheet_key(table_sheet),
                geometry=geometry,
                new_body_last_row=new_body_last_row,
            )
        )
        if new_reference is None:
            resized.append(part)
            continue
        prefix = part[: len(part) - len(reference)]
        resized.append(f"{prefix}{new_reference}")
        changed = True
    return ",".join(resized) if changed else None


def _covers(
    source: RangeSource,
    *,
    expected_sheet: str,
    rows: tuple[tuple[int, int], ...],
    first_col: int,
    last_col: int,
) -> bool:
    """Whether *source* is a rectangle over one of *rows* inside the table columns."""
    if _sheet_key(source.sheet_title) != _sheet_key(expected_sheet):
        return False
    try:
        bounds = CellRange(source.reference.replace("$", ""))
    except (ValueError, TypeError):
        # A whole-row (``$1:$1``) or whole-column (``$A:$A``) reference, which
        # openpyxl refuses to turn into a rectangle.  It cannot track a body.
        return False
    return (
        (bounds.min_row, bounds.max_row) in rows
        and first_col <= bounds.min_col
        and bounds.max_col <= last_col
    )


def same_sheet(title: str, other: str) -> bool:
    """Whether *title* and *other* name the same sheet.

    Examples:
        >>> same_sheet("Data", "data")
        True
        >>> same_sheet("Bob''s Data", "Bob's Data")
        True

    """
    return _sheet_key(title) == _sheet_key(other)


def _sheet_key(title: str) -> str:
    """*title* in the form two sheet names can be compared by.

    ``DefinedName.destinations`` leaves an apostrophe inside a quoted name
    doubled (``Bob''s Data``), while the workbook's own title is ``Bob's Data``.
    """
    return title.replace("''", "'").casefold()


def _display_sheet(title: str) -> str:
    """*title* with the doubled apostrophes of a quoted name unfolded."""
    return title.replace("''", "'")


def destination_references(refers_to: str) -> list[tuple[str, str]]:
    """``(sheet, reference)`` for each range a defined name points at.

    A name can be a constant (``"text"``), a number, a formula or a whole-row
    reference; only plain, sheet-qualified rectangles are returned, because
    they are the only ones a table resize can leave stale.

    Examples:
        >>> destination_references("Data!$A$2:$B$2")
        [('Data', '$A$2:$B$2')]
        >>> destination_references("Data!$A$1,Other!$B$2:$C$3")
        [('Data', '$A$1'), ('Other', '$B$2:$C$3')]
        >>> destination_references("'Bob''s Data'!$A$1")
        [("Bob's Data", '$A$1')]
        >>> destination_references('"a constant"')
        []
        >>> destination_references("OFFSET(Data!$A$1,0,0)")
        []
        >>> destination_references("Data!$1:$1")
        []

    """
    references: list[tuple[str, str]] = []
    for part in split_formula_arguments(refers_to.removeprefix("=")):
        parsed = sole_reference(part)
        if parsed is None or parsed[0] is None:
            continue
        references.append((parsed[0], parsed[1]))
    return references


def openpyxl_range_sources(workbook: Workbook, worksheet: Worksheet) -> list[RangeSource]:
    """Every defined name, pivot source and validation list that holds a range.

    Args:
        workbook: The workbook being filled.
        worksheet: The sheet holding the table (its own names and validations
            are read first, but every sheet is scanned: a validation may sit
            elsewhere and still point at this table).

    """
    sources = _name_sources(workbook.defined_names)
    sheets = [worksheet]
    sheets.extend(other for other in workbook.worksheets if other is not worksheet)
    for sheet in sheets:
        # A name scoped to a sheet lives in the sheet's own collection, not in
        # the workbook's.
        sources.extend(_name_sources(sheet.defined_names))
        sources.extend(_print_area_sources(sheet))
        sources.extend(_validation_sources(sheet))
        sources.extend(_pivot_sources(sheet))
    return sources


def _print_area_sources(worksheet: Worksheet) -> list[RangeSource]:
    """The print area of *worksheet*, which openpyxl keeps off the sheet's names.

    The template stores it as the built-in name ``_xlnm.Print_Area``.  Opening
    the workbook moves it to ``worksheet.print_area``, so a scan of the names
    alone would miss it and the openpyxl engine would stay silent where the
    fillxl engine warns.
    """
    area = worksheet.print_area
    areas = area if isinstance(area, list) else [area]
    return [
        RangeSource("the print area (_xlnm.Print_Area)", sheet_title, reference)
        for text in areas
        for sheet_title, reference in destination_references(str(text))
    ]


def _name_sources(names: Mapping[str, Any]) -> list[RangeSource]:
    """The ranges held by the defined names in *names*."""
    return [
        RangeSource(f"the defined name {name.name!r}", sheet_title, reference)
        for name in names.values()
        for sheet_title, reference in destination_references(str(name.attr_text or ""))
    ]


def _validation_sources(worksheet: Worksheet) -> list[RangeSource]:
    """The range a data-validation list is read from, per validation.

    A list source is stored as a bare reference, and an unqualified one means
    the sheet the validation sits on.
    """
    sources: list[RangeSource] = []
    own_sheet = str(worksheet.title)
    for validation in worksheet.data_validations.dataValidation:
        description = (
            f"the data validation over {validation.sqref}"
            if validation.sqref
            else "a data validation"
        )
        for formula in (validation.formula1, validation.formula2):
            if not isinstance(formula, str):
                continue
            parsed = sole_reference(formula)
            if parsed is None:
                continue
            sheet_title, reference = parsed
            sources.append(RangeSource(description, sheet_title or own_sheet, reference))
    return sources


def _pivot_sources(worksheet: Worksheet) -> list[RangeSource]:
    """The worksheet source of every pivot cache used on *worksheet*."""
    sources: list[RangeSource] = []
    for pivot in getattr(worksheet, "_pivots", ()):
        cache_source = getattr(getattr(pivot, "cache", None), "cacheSource", None)
        worksheet_source = getattr(cache_source, "worksheetSource", None)
        if worksheet_source is None:
            continue
        reference = getattr(worksheet_source, "ref", None)
        sheet_title = getattr(worksheet_source, "sheet", None)
        if not reference or not sheet_title:
            # A table name ("Sales") is a source Excel restyles itself.
            continue
        name = getattr(pivot, "name", None) or "a pivot table"
        sources.append(RangeSource(f"the pivot source of {name!r}", sheet_title, reference))
    return sources
