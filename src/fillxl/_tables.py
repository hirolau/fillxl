"""Table population logic."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from ._adapters import TabularAdapter
from ._charts import TableChangeSpecs
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
        ValueError: If the table has a totals row, or the table cannot expand
            into non-empty cells.

    """
    columns = adapter.columns
    rows = adapter.rows
    num_data_rows = len(rows)

    worksheet: Worksheet
    table: Table
    worksheet, table = find_worksheet_and_table(workbook, table_name)

    if _has_totals_row(table):
        msg = (
            f"Table '{table_name}' has a totals row, which fillxl does not "
            f"support: writing the data would overwrite the totals and "
            f"miscount the rows. Remove the totals row from the template."
        )
        raise ValueError(
            msg,
        )

    # Parse the original table reference (e.g. 'B12:M13')
    original_table_ref: str = table.ref
    original_cell_range = CellRange(original_table_ref)

    table_start_row = original_cell_range.min_row
    table_old_last_row = original_cell_range.max_row
    table_start_column = original_cell_range.min_col
    table_end_column = original_cell_range.max_col
    first_data_row = table_start_row + 1

    # An Excel table needs at least one body row, so empty data keeps a single
    # (cleared) row — the same thing the win32com backend does.
    new_table_last_row = table_start_row + max(num_data_rows, 1)
    new_table_ref = (
        f"{get_column_letter(table_start_column)}{table_start_row}:"
        f"{get_column_letter(table_end_column)}{new_table_last_row}"
    )
    worksheet_title: str = str(worksheet.title)

    # Check for overflow: if the table is growing, ensure no existing content
    # is overwritten
    if new_table_last_row > table_old_last_row:
        for check_row in range(table_old_last_row + 1, new_table_last_row + 1):
            for check_col in range(table_start_column, table_end_column + 1):
                existing = worksheet.cell(row=check_row, column=check_col).value
                if existing is not None:
                    cell_ref = f"{get_column_letter(check_col)}{check_row}"
                    msg = (
                        f"Table '{table_name}' cannot expand to {new_table_last_row} rows: "
                        f"cell {cell_ref} already contains data ({existing!r}). "
                        f"Move or remove the content below the table first."
                    )
                    raise ValueError(
                        msg,
                    )

    # Update table reference — the autoFilter must stay in step with the table,
    # otherwise Excel reports the workbook as repaired on open.
    table.ref = new_table_ref
    if table.autoFilter is not None:
        table.autoFilter.ref = new_table_ref
    logger.debug("Table '%s' resized: %s -> %s", table_name, original_table_ref, new_table_ref)

    # Get the first cell of the table
    table_first_cell = worksheet.cell(row=table_start_row, column=table_start_column)
    if not is_single_cell(table_first_cell):
        msg = "Expected a single cell! This should never happen in a table."
        raise ValueError(msg)

    # Clear any leftover rows the new data does not cover (this also clears the
    # single body row that empty data leaves behind).
    first_unused_row = first_data_row + num_data_rows
    for clear_row in range(first_unused_row, table_old_last_row + 1):
        for clear_col in range(table_start_column, table_end_column + 1):
            stale_cell = worksheet.cell(row=clear_row, column=clear_col)
            if is_single_cell(stale_cell):
                stale_cell.value = None

    # Create a mapping of data column names to column indices
    data_column_map = {str(col_name): idx for idx, col_name in enumerate(columns)}

    # Populate each table column
    for col_ix, column in enumerate(table.tableColumns):
        first_row_cell = table_first_cell.offset(column=col_ix, row=1)
        first_row_style = first_row_cell._style

        if column.name in data_column_map:
            data_col_idx = data_column_map[column.name]

            # Write each row value
            for row_ix, row_values in enumerate(rows):
                cell_to_fill = table_first_cell.offset(column=col_ix, row=row_ix + 1)

                # Copy the style from the first data row *before* assigning the
                # value: a date format is what tells Excel (and openpyxl on
                # read-back) that a datetime column holds dates rather than
                # serial numbers.
                if copy_style_in_tables:
                    cell_to_fill._style = first_row_style

                value = prepare_cell_value(row_values[data_col_idx])
                cell_to_fill.value = value

                # Force text for strings openpyxl would reinterpret: error
                # values such as "#N/A", and formulas when escaping is on.
                if is_text_cell_value(value, escape_formulas=escape_formulas):
                    cell_to_fill.data_type = "s"

        # Handle calculated columns (formulas)
        elif column.calculatedColumnFormula:
            formula = column.calculatedColumnFormula.attr_text or ""
            for row_ix in range(num_data_rows):
                cell_to_fill = table_first_cell.offset(column=col_ix, row=row_ix + 1)
                if copy_style_in_tables:
                    cell_to_fill._style = first_row_style
                cell_to_fill.value = "=" + formula

        else:
            # The column is missing from the data: clear whatever sample values
            # the template left in the rows this fill covers.
            column_index = table_start_column + col_ix
            for clear_row in range(first_data_row, new_table_last_row + 1):
                stale_cell = worksheet.cell(row=clear_row, column=column_index)
                if is_single_cell(stale_cell):
                    stale_cell.value = None

    # Extend conditional formatting to cover all new rows
    # Guard: skip when there are no data rows (first_data_row > new_table_last_row)
    if extend_conditional_formatting and first_data_row <= new_table_last_row:
        _extend_conditional_formatting(
            worksheet=worksheet,
            table_start_column=table_start_column,
            table_end_column=table_end_column,
            first_data_row=first_data_row,
            table_last_row=new_table_last_row,
        )

    return TableChangeSpecs(
        worksheet_title=worksheet_title,
        header_row=table_start_row,
        first_col=table_start_column,
        last_col=table_end_column,
        old_last_row=table_old_last_row,
        new_last_row=new_table_last_row,
    )


def _has_totals_row(table: Table) -> bool:
    """Whether *table* is configured with a totals row.

    Examples:
        >>> from openpyxl.worksheet.table import Table, TableColumn
        >>> columns = [TableColumn(id=1, name="a"), TableColumn(id=2, name="b")]
        >>> _has_totals_row(Table(displayName="T", ref="A1:B2", tableColumns=columns))
        False

    """
    return bool(getattr(table, "totalsRowCount", 0)) or bool(table.totalsRowShown)


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
        table_last_row: Row number of the last row in the table

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
