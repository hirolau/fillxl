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
from ._formulas import openpyxl_formula_columns, warn_about_ignored_formula_data
from ._normalize import is_text_cell_value, prepare_cell_value
from ._utils import find_worksheet_and_table

if TYPE_CHECKING:
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


def table_body_rows(table: Table) -> tuple[int, int]:
    """First and last body row of *table* (header and totals rows excluded)."""
    ref = CellRange(table.ref)
    header_rows = 0 if table.headerRowCount == 0 else 1
    return ref.min_row + header_rows, ref.max_row - _totals_row_count(table)


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
    num_data_rows = len(rows)

    worksheet: Worksheet
    table: Table
    worksheet, table = find_worksheet_and_table(workbook, table_name)

    # A totals row is part of the table's reference but not of its data body.
    totals_rows = _totals_row_count(table)
    # A table can be shown without a header row; Excel names its columns
    # Column1, Column2, … and the reference starts at the first data row.
    header_rows = 0 if table.headerRowCount == 0 else 1

    # Parse the original table reference (e.g. 'B12:M13')
    original_table_ref: str = table.ref
    original_cell_range = CellRange(original_table_ref)

    table_start_row = original_cell_range.min_row
    table_old_last_row = original_cell_range.max_row
    table_old_body_last_row = table_old_last_row - totals_rows
    table_start_column = original_cell_range.min_col
    table_end_column = original_cell_range.max_col
    first_data_row = table_start_row + header_rows

    # Calculated columns must be detected before any body cell is cleared.
    formula_columns = openpyxl_formula_columns(
        worksheet, table, first_data_row, table_old_body_last_row
    )
    warn_about_ignored_formula_data(table_name, columns, rows, formula_columns)

    # Snapshot the totals row before any clearing: openpyxl writes cells in
    # place, so the row has to be recreated at its new position below.
    totals_row_cells = (
        _read_row(worksheet, table_old_last_row, table_start_column, table_end_column)
        if totals_rows
        else None
    )

    # An Excel table needs at least one body row, so empty data keeps a single
    # (cleared) row — the same thing the win32com backend does.
    new_table_body_last_row = first_data_row + max(num_data_rows, 1) - 1
    new_table_last_row = new_table_body_last_row + totals_rows
    new_table_ref = (
        f"{get_column_letter(table_start_column)}{table_start_row}:"
        f"{get_column_letter(table_end_column)}{new_table_last_row}"
    )
    worksheet_title: str = str(worksheet.title)

    # Check for overflow: if the table is growing, ensure no existing content
    # is overwritten.  ``_cells`` is inspected directly because
    # ``worksheet.cell(...)`` would create the cell (and inflate ``max_row``)
    # just to look at it.
    if new_table_last_row > table_old_last_row:
        for check_row in range(table_old_last_row + 1, new_table_last_row + 1):
            for check_col in range(table_start_column, table_end_column + 1):
                existing_cell = worksheet._cells.get((check_row, check_col))
                if existing_cell is not None and existing_cell.value is not None:
                    cell_ref = f"{get_column_letter(check_col)}{check_row}"
                    msg = (
                        f"Table '{table_name}' cannot expand to {new_table_last_row} rows: "
                        f"cell {cell_ref} already contains data ({existing_cell.value!r}). "
                        f"Move or remove the content below the table first."
                    )
                    raise ValueError(
                        msg,
                    )

    # Update table reference — the autoFilter must stay in step with the table,
    # otherwise Excel reports the workbook as repaired on open.  Excel excludes
    # the totals row from the filter range.
    table.ref = new_table_ref
    if table.autoFilter is not None:
        table.autoFilter.ref = (
            f"{get_column_letter(table_start_column)}{table_start_row}:"
            f"{get_column_letter(table_end_column)}{new_table_body_last_row}"
        )
    logger.debug("Table '%s' resized: %s -> %s", table_name, original_table_ref, new_table_ref)

    # Get the first cell of the table
    table_first_cell = worksheet.cell(row=table_start_row, column=table_start_column)
    if not is_single_cell(table_first_cell):
        msg = "Expected a single cell! This should never happen in a table."
        raise ValueError(msg)

    # Clear any leftover rows the new data does not cover (this also clears the
    # single body row that empty data leaves behind).  The totals row is not
    # part of the body and is cleared only after its contents are captured.
    first_unused_row = first_data_row + num_data_rows
    for clear_row in range(first_unused_row, table_old_body_last_row + 1):
        for clear_col in range(table_start_column, table_end_column + 1):
            stale_cell = worksheet._cells.get((clear_row, clear_col))
            if stale_cell is not None and is_single_cell(stale_cell):
                stale_cell.value = None

    if totals_rows:
        _clear_row(worksheet, table_old_last_row, table_start_column, table_end_column)

    # Create a mapping of data column names to column indices
    data_column_map = {str(col_name): idx for idx, col_name in enumerate(columns)}

    # Populate each table column
    for col_ix, column in enumerate(table.tableColumns):
        first_row_cell = table_first_cell.offset(column=col_ix, row=header_rows)
        first_row_style = first_row_cell._style
        column_name = str(column.name)

        if column_name in formula_columns:
            formula = formula_columns[column_name]
            for row_ix in range(num_data_rows):
                cell_to_fill = table_first_cell.offset(column=col_ix, row=header_rows + row_ix)
                if copy_style_in_tables:
                    cell_to_fill._style = copy(first_row_style)
                cell_to_fill.value = formula

        elif column_name in data_column_map:
            data_col_idx = data_column_map[column_name]

            # Write each row value
            for row_ix, row_values in enumerate(rows):
                cell_to_fill = table_first_cell.offset(column=col_ix, row=header_rows + row_ix)

                # Copy style from the first data row *before* assigning the
                # value: a date format is what tells Excel (and openpyxl on
                # read-back) that a datetime column holds dates rather than
                # serial numbers.
                if copy_style_in_tables:
                    cell_to_fill._style = copy(first_row_style)

                value = prepare_cell_value(row_values[data_col_idx])
                cell_to_fill.value = value

                # Force text for strings openpyxl would reinterpret: error
                # values such as "#N/A", and formulas when escaping is on.
                if is_text_cell_value(value, escape_formulas=escape_formulas):
                    cell_to_fill.data_type = "s"

        else:
            # The column is missing from the data: clear whatever sample values
            # the template left in the rows this fill covers.
            column_index = table_start_column + col_ix
            for clear_row in range(first_data_row, new_table_body_last_row + 1):
                stale_cell = worksheet._cells.get((clear_row, column_index))
                if stale_cell is not None and is_single_cell(stale_cell):
                    stale_cell.value = None

    # Recreate the totals row at the bottom of the resized table.
    if totals_row_cells is not None:
        _write_row(worksheet, new_table_last_row, table_start_column, totals_row_cells)

    # Extend conditional formatting to cover all new data rows (the totals row
    # is deliberately excluded).
    # Guard: skip when there are no data rows (first_data_row > new_table_body_last_row)
    if extend_conditional_formatting and first_data_row <= new_table_body_last_row:
        _extend_conditional_formatting(
            worksheet=worksheet,
            table_start_column=table_start_column,
            table_end_column=table_end_column,
            first_data_row=first_data_row,
            table_last_row=new_table_body_last_row,
        )

    return TableChangeSpecs(
        worksheet_title=worksheet_title,
        header_row=table_start_row,
        first_col=table_start_column,
        last_col=table_end_column,
        old_last_row=table_old_body_last_row,
        new_last_row=new_table_body_last_row,
        header_rows=header_rows,
    )


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


def _clear_row(worksheet: Worksheet, row: int, first_col: int, last_col: int) -> None:
    """Clear the values in *row* between *first_col* and *last_col*."""
    for col in range(first_col, last_col + 1):
        cell = worksheet._cells.get((row, col))
        if cell is not None and is_single_cell(cell):
            cell.value = None


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
    first_data_row: int,
    table_last_row: int,
) -> None:
    """Extend conditional formatting rules from first data row to all rows in the table.

    A rule that tracks the table's height — one of its ranges starts on the
    first data row and lies inside the table's columns — is resized to the
    table's last row.  Ranges are parsed (never compared as strings), so a
    rule on ``B130:B200`` is not mistaken for one on ``B13``, and multi-area
    ranges such as ``B13:B20 D13:D20`` are handled per area.

    Args:
        worksheet: The worksheet containing the table
        table_start_column: 1-based index of the table's first column
        table_end_column: 1-based index of the table's last column
        first_data_row: Row number of the first data row
        table_last_row: Row number of the last data row in the table
            (the totals row, if present, is below it)

    """
    entries: list[Any] = list(worksheet.conditional_formatting)
    for conditional_formatting in entries:
        new_ranges: list[str] = []
        extended = False

        cell_ranges: list[Any] = sorted(conditional_formatting.sqref.ranges, key=str)
        for cell_range in cell_ranges:
            if _cf_range_tracks_table(
                cell_range,
                table_start_column=table_start_column,
                table_end_column=table_end_column,
                first_data_row=first_data_row,
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


def _cf_range_tracks_table(
    cell_range: CellRange,
    *,
    table_start_column: int,
    table_end_column: int,
    first_data_row: int,
    table_last_row: int,
) -> bool:
    """Whether *cell_range* should be resized to the table's last row.

    Such a range starts on the first data row and lies inside the table's
    columns, so it grows and shrinks with the table.

    Examples:
        >>> def tracks(area, last_row):
        ...     return _cf_range_tracks_table(
        ...         CellRange(area),
        ...         table_start_column=2,
        ...         table_end_column=13,
        ...         first_data_row=13,
        ...         table_last_row=last_row,
        ...     )
        >>> tracks("B13:B20", 30)
        True
        >>> tracks("B13:B20", 20)
        False
        >>> tracks("B130:B200", 30)
        False

    """
    return (
        cell_range.min_row == first_data_row
        and cell_range.max_row != table_last_row
        and cell_range.min_col >= table_start_column
        and cell_range.max_col <= table_end_column
    )
