"""Table population logic."""

from __future__ import annotations

import logging
from copy import copy
from typing import TYPE_CHECKING, Any, NamedTuple

from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from ._adapters import TabularAdapter
from ._charts import TableChangeSpecs
from ._formulas import (
    extend_table_ranges,
    openpyxl_formula_columns,
    shift_relative_rows,
    warn_about_ignored_formula_data,
)
from ._normalize import (
    is_text_cell_value,
    prepare_table_cell,
)
from ._plan import (
    BlockedRegion,
    RowSpan,
    TableGeometry,
    TablePlan,
    blocked_area,
    plan_table_fill,
)
from ._references import openpyxl_range_sources, warn_about_untracked_ranges
from ._utils import find_worksheet_and_table, table_column_name, workbook_uses_1904

if TYPE_CHECKING:
    from collections.abc import Sequence

    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.table import Table
    from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


def is_single_cell(cell: object) -> bool:
    """Check if an object is a single Cell instance (not a MergedCell).

    Args:
        cell: Object to check

    Returns:
        True if *cell* is a :class:`~openpyxl.cell.cell.Cell` instance.

    Examples:
        >>> from openpyxl.cell.cell import Cell
        >>> is_single_cell(Cell(None, None, None, None))
        True
        >>> is_single_cell("not a cell")
        False
        >>> from openpyxl.worksheet.cell_range import CellRange
        >>> is_single_cell(CellRange("A1:B2"))
        False

    """
    return isinstance(cell, Cell)


def openpyxl_table_geometry(table: Table) -> TableGeometry:
    """Where *table* sits, read from its reference and header/totals settings."""
    ref = CellRange(table.ref)
    return TableGeometry(
        top_row=ref.min_row,
        last_row=ref.max_row,
        first_col=ref.min_col,
        last_col=ref.max_col,
        header_rows=0 if table.headerRowCount == 0 else 1,
        totals_rows=_totals_row_count(table),
    )


def _clear_rows(
    worksheet: Worksheet,
    rows: RowSpan,
    first_col: int,
    last_col: int,
    *,
    keep: Sequence[int] = (),
) -> None:
    """Clear values (not styles) in *rows* between *first_col* and *last_col*.

    Columns in *keep* are left alone: a calculated column keeps its formula,
    which is what Excel does when it shrinks a table.
    """
    first_row, last_row = rows
    skipped = set(keep)
    for row in range(first_row, last_row + 1):
        for col in range(first_col, last_col + 1):
            if col in skipped:
                continue
            cell = worksheet._cells.get((row, col))
            if cell is not None and is_single_cell(cell):
                cell.value = None


def _show_rows_and_drop_filters(
    worksheet: Worksheet,
    table: Table,
    plan: TablePlan,
    original_autofilter_ref: str | None,
) -> None:
    """Show the rows a fill writes and drop the table's filter criteria.

    Excel does not reapply a filter when a table is resized, and a row hidden
    by the author stays hidden.  Both hide data a fill has just written, and a
    ``SUBTOTAL(109, …)`` totals row silently leaves hidden rows out.  The
    worksheet-level filter is cleared only when it covered the table itself.
    """
    for row in range(plan.first_data_row, plan.new_last_row + 1):
        dimension = worksheet.row_dimensions.get(row)
        if dimension is not None and dimension.hidden:
            dimension.hidden = False
    cleared: list[Any] = []
    table_filter = table.autoFilter
    if table_filter is not None and table_filter.filterColumn:
        table_filter.filterColumn = cleared
    sheet_filter = worksheet.auto_filter
    if (
        sheet_filter is not None
        and sheet_filter.filterColumn
        and original_autofilter_ref is not None
        and sheet_filter.ref == original_autofilter_ref
    ):
        sheet_filter.filterColumn = cleared


def _check_overflow_region(
    worksheet: Worksheet,
    table_name: str,
    table: Table,
    rows: RowSpan,
    first_col: int,
    last_col: int,
    data_rows: int,
) -> None:
    """Refuse to grow a table into another table or into merged cells.

    Those cells can look empty and still break the file: an overlapping table
    makes Excel repair the workbook, and writing to a merged cell raises
    openpyxl's "read-only" ``AttributeError``.
    """
    blocked = blocked_area(
        rows,
        first_col,
        last_col,
        _occupied_regions(worksheet, table),
    )
    if blocked is None:
        return
    msg = (
        f"Table '{table_name}' cannot expand to {data_rows} rows: "
        f"{blocked.description} is in the way. Move or remove the content below "
        "the table first."
    )
    raise ValueError(msg)


def _occupied_regions(worksheet: Worksheet, table: Table) -> list[BlockedRegion]:
    """Everything on *worksheet* that a table must not grow into."""
    regions = [
        BlockedRegion(
            description=f"table {other.name!r} at {other.ref}",
            first_row=CellRange(other.ref).min_row,
            last_row=CellRange(other.ref).max_row,
            first_col=CellRange(other.ref).min_col,
            last_col=CellRange(other.ref).max_col,
        )
        for other in worksheet.tables.values()
        if other is not table
    ]
    regions.extend(
        BlockedRegion(
            description=f"the merged range {merged}",
            first_row=merged.min_row,
            last_row=merged.max_row,
            first_col=merged.min_col,
            last_col=merged.max_col,
        )
        for merged in worksheet.merged_cells.ranges
    )
    return regions


def _check_overflow(
    worksheet: Worksheet,
    table_name: str,
    rows: RowSpan,
    first_col: int,
    last_col: int,
    data_rows: int,
) -> None:
    """Refuse to grow a table over non-empty cells.

    ``_cells`` is read directly: ``worksheet.cell()`` would create the cell.
    """
    first_row, last_row = rows
    for row in range(first_row, last_row + 1):
        for col in range(first_col, last_col + 1):
            cell = worksheet._cells.get((row, col))
            if cell is not None and cell.value is not None:
                msg = (
                    f"Table '{table_name}' cannot expand to {data_rows} rows: "
                    f"cell {get_column_letter(col)}{row} already contains data "
                    f"({cell.value!r}). Move or remove the content below the table first."
                )
                raise ValueError(msg)


def populate_table(
    workbook: Workbook,
    table_name: str,
    adapter: TabularAdapter,
    *,
    copy_style_in_tables: bool = True,
    extend_conditional_formatting: bool = True,
    escape_formulas: bool = False,
) -> TableChangeSpecs:
    """Populate an Excel table with tabular data.

    Args:
        workbook: openpyxl Workbook instance
        table_name: Name of the table to populate
        adapter: Adapter holding the column names and rows to write.  The
            caller builds it (and reuses it for column validation) so that
            the underlying data is not read twice.
        copy_style_in_tables: If True, copy cell styles from first data row to new rows
        extend_conditional_formatting: If True, extend conditional formatting
            rules to cover new rows
        escape_formulas: If True, strings starting with ``=`` are written as
            literal text instead of becoming formulas.

    Returns:
        TableChangeSpecs describing the old and new table dimensions

    Raises:
        KeyError: If table is not found
        ValueError: If the table cannot expand into non-empty cells.

    """
    columns = adapter.columns
    rows = adapter.rows

    worksheet, table = find_worksheet_and_table(workbook, table_name)
    geometry = openpyxl_table_geometry(table)
    plan = plan_table_fill(geometry, len(rows), table_name=table_name)
    first_col, last_col = geometry.first_col, geometry.last_col
    header_rows = geometry.header_rows

    # Calculated columns must be detected before any body cell is cleared.
    formula_columns = openpyxl_formula_columns(
        worksheet, table, geometry.first_data_row, geometry.body_last_row
    )
    warn_about_ignored_formula_data(table_name, columns, rows, formula_columns)

    # openpyxl writes cells in place, so the totals row is captured before
    # clearing and recreated at its new position afterwards.
    totals_row_cells = (
        _read_row(worksheet, geometry.last_row, first_col, last_col)
        if geometry.totals_rows
        else None
    )

    if plan.overflow_rows is not None:
        _check_overflow_region(
            worksheet,
            table_name,
            table,
            plan.overflow_rows,
            first_col,
            last_col,
            len(rows),
        )
        _check_overflow(worksheet, table_name, plan.overflow_rows, first_col, last_col, len(rows))

    # The autoFilter must stay in step with the table, or Excel repairs the file.
    original_ref = table.ref
    original_autofilter_ref = table.autoFilter.ref if table.autoFilter is not None else None
    table.ref = plan.new_ref
    if table.autoFilter is not None:
        table.autoFilter.ref = plan.new_autofilter_ref
    # A template saved with a filter applied (or with rows hidden by hand)
    # would keep some of the filled rows out of sight -- and out of a
    # SUBTOTAL(109, …) totals row.  The fill shows everything it writes.
    _show_rows_and_drop_filters(worksheet, table, plan, original_autofilter_ref)
    logger.debug("Table '%s' resized: %s -> %s", table_name, original_ref, plan.new_ref)

    table_first_cell = worksheet.cell(row=geometry.top_row, column=first_col)
    if not is_single_cell(table_first_cell):
        msg = "Expected a single cell! This should never happen in a table."
        raise ValueError(msg)

    # Calculated columns keep their formulas, including when the data is empty.
    formula_column_offsets = {
        offset
        for offset, column in enumerate(table.tableColumns)
        if table_column_name(column) in formula_columns
    }
    keep_columns = [first_col + offset for offset in formula_column_offsets]
    if plan.unused_body_rows is not None:
        _clear_rows(worksheet, plan.unused_body_rows, first_col, last_col, keep=keep_columns)
    if plan.vacated_rows is not None:
        _clear_rows(worksheet, plan.vacated_rows, first_col, last_col)
    if geometry.totals_rows:
        _clear_rows(worksheet, (geometry.last_row, geometry.last_row), first_col, last_col)
    if plan.data_rows == 0:
        _clear_rows(
            worksheet,
            (plan.first_data_row, plan.first_data_row),
            first_col,
            last_col,
            keep=keep_columns,
        )

    data_column_map = {str(name): idx for idx, name in enumerate(columns)}
    warned: set[tuple[str | None, str | None]] = set()
    date1904 = workbook_uses_1904(workbook)

    for col_ix, column in enumerate(table.tableColumns):
        first_row_style = table_first_cell.offset(column=col_ix, row=header_rows)._style
        column_name = table_column_name(column)

        if column_name in formula_columns:
            formula = formula_columns[column_name]
            for row_ix in range(plan.data_rows):
                cell = table_first_cell.offset(column=col_ix, row=header_rows + row_ix)
                if copy_style_in_tables:
                    cell._style = copy(first_row_style)
                # Excel copies the definition down, which moves its relative
                # row references; writing the same text on every row would make
                # every row compute the first row's result.
                cell.value = shift_relative_rows(formula, row_ix)

        elif column_name in data_column_map:
            data_col_idx = data_column_map[column_name]
            for row_ix, row_values in enumerate(rows):
                cell = table_first_cell.offset(column=col_ix, row=header_rows + row_ix)
                raw_value = row_values[data_col_idx]
                value = prepare_table_cell(
                    raw_value,
                    table=table_name,
                    row=row_ix + 1,
                    column=column_name,
                    warned=warned,
                    date1904=date1904,
                )
                # Style first: a date format must be in place before the value
                # arrives, or openpyxl replaces it with its default date format.
                if copy_style_in_tables:
                    cell._style = copy(first_row_style)
                cell.value = value
                # Force text for strings openpyxl would reinterpret: error
                # values such as "#N/A", and formulas when escaping is on.
                if is_text_cell_value(value, escape_formulas=escape_formulas):
                    cell.data_type = "s"

        else:
            # The data has no such column: clear the template's sample values.
            column_index = first_col + col_ix
            _clear_rows(
                worksheet,
                (plan.first_data_row, plan.new_body_last_row),
                column_index,
                column_index,
            )

    if totals_row_cells is not None:
        _write_row(worksheet, plan.new_last_row, first_col, totals_row_cells)
        if plan.vacated_rows is not None:
            # The old totals row keeps its bold text and rules unless the style
            # goes too; Excel moves the row up and leaves a plain row behind.
            _reset_row_style(
                worksheet,
                geometry.last_row,
                plan.first_data_row,
                first_col,
                last_col,
            )

    if plan.new_body_last_row != geometry.body_last_row:
        _resize_adjacent_formulas(workbook, worksheet, plan)
        warn_about_untracked_ranges(
            openpyxl_range_sources(workbook, worksheet),
            table_name=table_name,
            sheet_title=str(worksheet.title),
            geometry=geometry,
            new_body_last_row=plan.new_body_last_row,
        )

    if extend_conditional_formatting:
        _extend_conditional_formatting(
            worksheet=worksheet,
            table_start_column=first_col,
            table_end_column=last_col,
            table_top_row=geometry.top_row,
            first_data_row=plan.first_data_row,
            old_body_last_row=geometry.body_last_row,
            table_last_row=plan.new_body_last_row,
        )

    return plan.change_spec(str(worksheet.title))


def _resize_adjacent_formulas(
    workbook: Workbook,
    worksheet: Worksheet,
    plan: TablePlan,
) -> None:
    """Resize the formulas that tracked the table body, as Excel does.

    Excel rewrites ``=SUM(B4:B4)`` to ``=SUM(B4:B7)`` when the table grows and
    back again when it shrinks — on the table's own sheet, and on any other
    sheet that refers to it.  openpyxl writes formula text, so the range has to
    be rewritten here or the summary keeps summing the first row only.

    Cells in the table's data body are skipped: their formulas belong to the
    table and are written by :func:`populate_table`.  The totals row is *not*
    skipped -- it moved with the table, and an A1 formula there such as
    ``=SUM(B4:B4)`` has to follow the body like any other.  Ranges that only
    look like they track the table are left alone (see
    :func:`extend_table_ranges`).
    """
    geometry = plan.geometry
    first_col, last_col = geometry.first_col, geometry.last_col
    table_sheet = str(worksheet.title)
    for sheet in workbook.worksheets:
        own_sheet = str(sheet.title) == table_sheet
        for (row, column), cell in list(sheet._cells.items()):
            in_body = (
                own_sheet
                and plan.first_data_row <= row <= plan.new_body_last_row
                and first_col <= column <= last_col
            )
            if in_body or not is_single_cell(cell) or cell.data_type != "f":
                continue
            formula = cell.value
            if not isinstance(formula, str):
                continue
            updated = extend_table_ranges(
                formula,
                sheet_title=table_sheet,
                formula_sheet=str(sheet.title),
                first_col=geometry.first_col,
                last_col=geometry.last_col,
                old_body=(geometry.first_data_row, geometry.body_last_row),
                new_body=(plan.first_data_row, plan.new_body_last_row),
            )
            if updated is not None:
                logger.debug("Resized formula %s -> %s", formula, updated)
                cell.value = updated


def _reset_row_style(
    worksheet: Worksheet,
    row: int,
    template_row: int,
    first_col: int,
    last_col: int,
) -> None:
    """Give *row* the style of *template_row*, so no leftover formatting remains."""
    if row == template_row:
        return
    for col in range(first_col, last_col + 1):
        target = worksheet._cells.get((row, col))
        source = worksheet._cells.get((template_row, col))
        if target is None or not is_single_cell(target):
            continue
        if source is not None and is_single_cell(source):
            target._style = copy(source._style)


def _totals_row_count(table: Table) -> int:
    """Number of totals rows in *table* (zero or one).

    Only ``totalsRowCount`` counts: ``totalsRowShown`` means "was shown at some
    point", not "is shown" (ECMA-376), matching COM's ``ShowTotals``.

    Examples:
        >>> from openpyxl.worksheet.table import Table, TableColumn
        >>> columns = [TableColumn(id=1, name="a"), TableColumn(id=2, name="b")]
        >>> _totals_row_count(Table(displayName="T", ref="A1:B2", tableColumns=columns))
        0
        >>> _totals_row_count(
        ...     Table(displayName="T", ref="A1:B3", tableColumns=columns, totalsRowShown=True)
        ... )
        0
        >>> _totals_row_count(
        ...     Table(displayName="T", ref="A1:B3", tableColumns=columns, totalsRowCount=1)
        ... )
        1

    """
    count = getattr(table, "totalsRowCount", None)
    return int(count) if count else 0


class _CellSnapshot(NamedTuple):
    """Value, style and type of a cell captured before the cell is cleared."""

    value: Any
    style: Any
    data_type: str | None


def _read_row(
    worksheet: Worksheet,
    row: int,
    first_col: int,
    last_col: int,
) -> list[_CellSnapshot]:
    """Capture the cells in *row* between *first_col* and *last_col*."""
    snapshots: list[_CellSnapshot] = []
    for col in range(first_col, last_col + 1):
        cell = worksheet._cells.get((row, col))
        if cell is None or not is_single_cell(cell):
            snapshots.append(_CellSnapshot(None, None, None))
            continue
        snapshots.append(_CellSnapshot(cell.value, copy(cell._style), cell.data_type))
    return snapshots


def _write_row(
    worksheet: Worksheet,
    row: int,
    first_col: int,
    snapshots: list[_CellSnapshot],
) -> None:
    """Write snapshots captured by :func:`_read_row` into *row*."""
    for offset, snapshot in enumerate(snapshots):
        cell = worksheet.cell(row=row, column=first_col + offset)
        if not is_single_cell(cell):
            continue
        if snapshot.style is not None:
            cell._style = copy(snapshot.style)
        cell.value = snapshot.value
        if snapshot.data_type is not None:
            cell.data_type = snapshot.data_type


def _extend_conditional_formatting(
    worksheet: Worksheet,
    table_start_column: int,
    table_end_column: int,
    table_top_row: int,
    first_data_row: int,
    old_body_last_row: int,
    table_last_row: int,
) -> None:
    """Extend conditional formatting rules that track the table's height.

    A rule tracks the table when it starts at the table's top row (its header)
    or at its first data row, or when it ends exactly on the table's last data
    row — and when it lies inside the table's columns.  Such a rule is resized
    to the table's new last row.  Ranges are parsed (never compared as
    strings), so a rule on ``B130:B200`` is not mistaken for one on ``B13``,
    and multi-area ranges such as ``B13:B20 D13:D20`` are handled per area.

    Args:
        worksheet: The worksheet containing the table
        table_start_column: 1-based index of the table's first column
        table_end_column: 1-based index of the table's last column
        table_top_row: Row number of the table's first row (header, if shown)
        first_data_row: Row number of the first data row
        old_body_last_row: Row number of the last data row before the fill
        table_last_row: Row number of the last data row in the table
            (the totals row, if present, is below it)

    """
    entries: list[Any] = list(worksheet.conditional_formatting)
    for conditional_formatting in entries:
        new_ranges: list[str] = []
        extended = False

        cell_ranges: list[Any] = sorted(conditional_formatting.sqref.ranges, key=str)
        for cell_range in cell_ranges:
            if cf_range_tracks_table(
                cell_range,
                table_start_column=table_start_column,
                table_end_column=table_end_column,
                table_top_row=table_top_row,
                first_data_row=first_data_row,
                old_body_last_row=old_body_last_row,
                table_last_row=table_last_row,
            ):
                new_ranges.append(
                    CellRange(
                        min_col=cell_range.min_col,
                        min_row=cell_range.min_row,
                        max_col=cell_range.max_col,
                        max_row=table_last_row,
                    ).coord,
                )
                extended = True
            else:
                new_ranges.append(cell_range.coord)

        if not extended:
            continue

        # The rules are re-added under the new sqref instead of mutating
        # ``sqref`` in place: ConditionalFormattingList keys its entries by
        # the object, whose hash includes the range.
        old_sqref = str(conditional_formatting.sqref)
        new_sqref = " ".join(new_ranges)
        rules = list(conditional_formatting.rules)
        del worksheet.conditional_formatting[old_sqref]
        for rule in rules:
            worksheet.conditional_formatting.add(new_sqref, rule)

        logger.debug("Extended conditional formatting: %s -> %s", old_sqref, new_sqref)


def cf_range_tracks_table(
    cell_range: CellRange,
    *,
    table_start_column: int,
    table_end_column: int,
    table_top_row: int,
    first_data_row: int,
    old_body_last_row: int,
    table_last_row: int,
) -> bool:
    """Whether *cell_range* should be resized to the table's last row.

    The rule matches what Excel does — and what a fillxl user means by "extend
    with the table":

    * the area lies inside the table's columns;
    * it starts at the table's top row or its first data row, or it ends
      exactly on the table's last data row;
    * it does not already end on the table's new last row;
    * it does not reach beyond the table — a rule covering more rows than the
      table is the caller's own area and Excel leaves it alone.

    Examples:
        >>> def tracks(area, new_last, old_last=20):
        ...     return cf_range_tracks_table(
        ...         CellRange(area),
        ...         table_start_column=2,
        ...         table_end_column=13,
        ...         table_top_row=12,
        ...         first_data_row=13,
        ...         old_body_last_row=old_last,
        ...         table_last_row=new_last,
        ...     )
        >>> tracks("B13:B20", 30)
        True
        >>> tracks("B13:B20", 20)
        False
        >>> tracks("B130:B200", 30)
        False
        >>> tracks("B12:B20", 30)   # starts on the header row, as Excel allows
        True
        >>> tracks("B2:B200", 30)   # reaches past the table: not ours to resize
        False

    """
    inside_columns = (
        cell_range.min_col >= table_start_column and cell_range.max_col <= table_end_column
    )
    if not inside_columns or cell_range.max_row == table_last_row:
        return False
    starts_with_table = cell_range.min_row in (table_top_row, first_data_row)
    if not (starts_with_table or cell_range.max_row == old_body_last_row):
        return False
    return not (cell_range.max_row > old_body_last_row and cell_range.max_row > table_last_row)
