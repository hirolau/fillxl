"""Table population logic."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from openpyxl.cell.cell import Cell
from openpyxl.utils import column_index_from_string

from ._adapters import TabularAdapter
from ._charts import TableChangeSpecs
from ._normalize import is_text_cell_value, prepare_cell_value
from ._utils import col_to_excel, find_worksheet_and_table

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
    data: Any,  # noqa: ANN401
    *,
    copy_style_in_tables: bool = True,
    extend_conditional_formatting: bool = True,
) -> TableChangeSpecs:
    """Populate an Excel table with tabular data.

    Args:
        workbook: openpyxl Workbook instance
        table_name: Name of the table to populate
        data: Tabular data (DataFrame, list[dict], dict[str, list], etc.)
        copy_style_in_tables: If True, copy cell styles from first data row to new rows
        extend_conditional_formatting: If True, extend conditional formatting
            rules to cover new rows

    Returns:
        TableChangeSpecs describing the old and new table references

    Raises:
        KeyError: If table is not found
        ValueError: If table structure is invalid

    """
    # Normalize data using adapter
    adapter = TabularAdapter(data)
    columns = adapter.columns
    rows = list(adapter.rows)  # Materialize rows for length calculation

    worksheet: Worksheet
    table: Table
    worksheet, table = find_worksheet_and_table(workbook, table_name)

    # Parse original table reference (e.g., 'B12:M13')
    original_table_ref: str = table.ref
    table_start, table_end = original_table_ref.split(":")

    # Calculate dimensions
    table_start_row = int("".join(x for x in table_start if x.isdigit()))
    table_old_end_row = int("".join(x for x in table_end if x.isdigit()))
    table_end_row = table_start_row + len(rows)
    table_start_column_letter = "".join(x for x in table_start if not x.isdigit())
    table_start_column = column_index_from_string(table_start_column_letter)
    table_end_column_letter = "".join(x for x in table_end if not x.isdigit())
    first_data_row = table_start_row + 1
    worksheet_title: str = str(worksheet.title)

    # Check for overflow: if the table is growing, ensure no existing content
    # is overwritten
    if table_end_row > table_old_end_row:
        end_col = column_index_from_string(table_end_column_letter)
        for check_row in range(table_old_end_row + 1, table_end_row + 1):
            for check_col in range(table_start_column, end_col + 1):
                existing = worksheet.cell(row=check_row, column=check_col).value
                if existing is not None:
                    cell_ref = f"{col_to_excel(check_col)}{check_row}"
                    msg = (
                        f"Table '{table_name}' cannot expand to {table_end_row} rows: "
                        f"cell {cell_ref} already contains data ({existing!r}). "
                        f"Move or remove the content below the table first."
                    )
                    raise ValueError(
                        msg,
                    )

    # Update table reference
    new_table_ref = table_start + ":" + table_end_column_letter + str(table_start_row + len(rows))
    table.ref = new_table_ref
    logger.debug("Table '%s' resized: %s -> %s", table_name, original_table_ref, new_table_ref)

    # Get the first cell of the table
    table_first_cell = worksheet.cell(row=table_start_row, column=table_start_column)
    if not is_single_cell(table_first_cell):
        msg = "Expected a single cell! This should never happen in a table."
        raise ValueError(msg)

    # Clear any leftover rows if the new data is shorter than the old table
    if table_old_end_row > table_end_row:
        range_to_clear = worksheet[
            f"{table_start_column_letter}{table_end_row + 1}:"
            f"{table_end_column_letter}{table_old_end_row}"
        ]
        for row in range_to_clear:
            for cell in row:
                if is_single_cell(cell):
                    cell.value = None

    # Create a mapping of data column names to column indices
    data_column_map = {col_name: idx for idx, col_name in enumerate(columns)}

    # Populate each table column
    for col_ix, column in enumerate(table.tableColumns):
        first_row_cell = table_first_cell.offset(column=col_ix, row=1)
        first_row_style = first_row_cell._style

        # Check if this column exists in the data
        if column.name in data_column_map:
            data_col_idx = data_column_map[column.name]

            # Write each row value
            for row_ix, row_values in enumerate(rows):
                cell_to_fill = table_first_cell.offset(column=col_ix, row=row_ix + 1)

                value = prepare_cell_value(row_values[data_col_idx])
                cell_to_fill.value = value

                # Force text for strings openpyxl would reinterpret: error
                # values such as "#N/A".
                if is_text_cell_value(value):
                    cell_to_fill.data_type = "s"

                # Copy style from first row
                if copy_style_in_tables:
                    cell_to_fill._style = first_row_style

        # Handle calculated columns (formulas)
        elif column.calculatedColumnFormula:
            formula = column.calculatedColumnFormula.attr_text or ""
            for row_ix in range(len(rows)):
                cell_to_fill = table_first_cell.offset(column=col_ix, row=row_ix + 1)
                cell_to_fill.value = "=" + formula
                if copy_style_in_tables:
                    cell_to_fill._style = first_row_style

    # Extend conditional formatting to cover all new rows
    # Guard: skip when there are no data rows (first_data_row > table_end_row)
    if extend_conditional_formatting and first_data_row <= table_end_row:
        _extend_conditional_formatting(
            worksheet=worksheet,
            table_start_column_letter=table_start_column_letter,
            table_end_column_letter=table_end_column_letter,
            first_data_row=first_data_row,
            table_end_row=table_end_row,
        )

    return TableChangeSpecs(
        old_ref=original_table_ref,
        new_ref=new_table_ref,
        worksheet_title=worksheet_title,
    )


def _extend_conditional_formatting(
    worksheet: Worksheet,
    table_start_column_letter: str,
    table_end_column_letter: str,
    first_data_row: int,
    table_end_row: int,
) -> None:
    """Extend conditional formatting rules from first data row to all rows in the table.

    Args:
        worksheet: The worksheet containing the table
        table_start_column_letter: Starting column letter (e.g., 'B')
        table_end_column_letter: Ending column letter (e.g., 'M')
        first_data_row: Row number of the first data row
        table_end_row: Row number of the last row in the table

    """
    conditional_formattings = [str(c.sqref) for c in worksheet.conditional_formatting]
    start_col_idx = column_index_from_string(table_start_column_letter)
    end_col_idx = column_index_from_string(table_end_column_letter)

    for col_idx in range(start_col_idx, end_col_idx + 1):
        # Check each column (B13, C13, ...)
        range_to_check = f"{col_to_excel(col_idx)}{first_data_row}"

        for conditional_formatting in conditional_formattings:
            if conditional_formatting.startswith(range_to_check):
                # Parse the current range
                if ":" not in conditional_formatting:
                    current_end = conditional_formatting
                else:
                    _, current_end = conditional_formatting.split(":")

                end_letter = "".join(x for x in current_end if not x.isdigit())

                # Get the rules for this range
                rules = worksheet.conditional_formatting[conditional_formatting]
                del worksheet.conditional_formatting[conditional_formatting]

                # Create new extended range
                new_range = (
                    conditional_formatting.split(":")[0] + ":" + f"{end_letter}{table_end_row}"
                )

                # Re-add all rules with the new range
                for rule in rules:
                    worksheet.conditional_formatting.add(new_range, rule)

                logger.debug(
                    "Extended conditional formatting: %s -> %s",
                    conditional_formatting,
                    new_range,
                )
