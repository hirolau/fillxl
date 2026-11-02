"""Ranges that point at a table body without being a cell formula.

A fill resizes the charts and the A1 formulas that track a table, because both
are readable and rewritable.  A pivot cache source, a defined name and a
data-validation list hold the same rectangle in places no engine edits, so
after a fill they still cover the template's rows.  Those are worth a warning:
a pivot marked ``refreshOnLoad`` refreshes onto the old range and shows a total
that quietly leaves the filled rows out.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, NamedTuple

from openpyxl.worksheet.cell_range import CellRange

from ._cells import name_destinations
from ._warnings import UntrackedRangeWarning, warn_user

if TYPE_CHECKING:
    from collections.abc import Iterable

    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet

    from ._plan import TableGeometry


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
            f"{source.sheet_title}!{source.reference} — the table's old rows. The rows "
            "the fill wrote are missing from it, so a pivot that refreshes on load or a "
            "name that is read back will use the old range. Resize that reference in the "
            "template to match the table.",
            UntrackedRangeWarning,
        )


def _covers(
    source: RangeSource,
    *,
    expected_sheet: str,
    rows: tuple[tuple[int, int], ...],
    first_col: int,
    last_col: int,
) -> bool:
    """Whether *source* is a rectangle over one of *rows* inside the table columns."""
    if source.sheet_title.casefold() != expected_sheet.casefold():
        return False
    try:
        bounds = CellRange(source.reference.replace("$", ""))
    except ValueError:  # pragma: no cover - callers pass ranges from the file
        return False
    return (
        (bounds.min_row, bounds.max_row) in rows
        and first_col <= bounds.min_col
        and bounds.max_col <= last_col
    )


def openpyxl_range_sources(workbook: Workbook, worksheet: Worksheet) -> list[RangeSource]:
    """Every defined name, pivot source and validation list that holds a range.

    Args:
        workbook: The workbook being filled.
        worksheet: The sheet holding the table (its own validations are read
            first, but every sheet is scanned: a validation may sit elsewhere
            and still point at this table).

    """
    sources: list[RangeSource] = []
    for name in workbook.defined_names.values():
        for sheet_title, reference in name.destinations:
            sources.append(RangeSource(f"the defined name {name.name!r}", sheet_title, reference))
    sheets = [worksheet]
    sheets.extend(other for other in workbook.worksheets if other is not worksheet)
    for sheet in sheets:
        sources.extend(_validation_sources(sheet))
        sources.extend(_pivot_sources(sheet))
    return sources


def _validation_sources(worksheet: Worksheet) -> list[RangeSource]:
    """The ranges a data-validation list is read from, per validation."""
    sources: list[RangeSource] = []
    for validation in worksheet.data_validations.dataValidation:
        description = (
            f"the data validation over {validation.sqref}"
            if validation.sqref
            else "a data validation"
        )
        for formula in (validation.formula1, validation.formula2):
            if not isinstance(formula, str):
                continue
            for sheet_title, reference in name_destinations(formula):
                sources.append(RangeSource(description, sheet_title, reference))
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
