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
from ._formulas import sole_reference
from ._warnings import UntrackedRangeWarning, warn_user

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping

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
            f"{_display_sheet(source.sheet_title)}!{source.reference} — the table's old "
            "rows. The rows the fill wrote are missing from it, so a pivot that refreshes "
            "on load or a name that is read back will use the old range. Resize that "
            "reference in the template to match the table.",
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
        sources.extend(_validation_sources(sheet))
        sources.extend(_pivot_sources(sheet))
    return sources


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
