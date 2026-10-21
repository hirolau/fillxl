"""Utility functions for Excel manipulation."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.table import Table, TableColumn
    from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


def table_column_name(column: TableColumn) -> str:
    """The name of a table column as Excel shows it.

    The file format escapes characters such as line breaks (``_x000a_``).
    COM returns them unescaped, so openpyxl names are unescaped to match.

    Examples:
        >>> from openpyxl.worksheet.table import TableColumn
        >>> table_column_name(TableColumn(id=1, name="NAV_x000a_SEK"))
        'NAV\\nSEK'

    """
    from openpyxl.utils.escape import unescape

    return unescape(str(column.name))


def col_to_excel(col: int) -> str:
    """Convert 1-based column number to Excel column letter(s).

    Args:
        col: 1-based column number

    Returns:
        Excel column letter(s)

    Examples:
        >>> col_to_excel(1)
        'A'
        >>> col_to_excel(26)
        'Z'
        >>> col_to_excel(27)
        'AA'

    """
    result = ""
    while col > 0:
        col -= 1
        result = chr(col % 26 + ord("A")) + result
        col //= 26
    return result


def find_worksheet_and_table(workbook: Workbook, table_name: str) -> tuple[Worksheet, Table]:
    """Find a worksheet and table by table name.

    Args:
        workbook: openpyxl Workbook instance
        table_name: Name of the table to find

    Returns:
        Tuple of (Worksheet, Table)

    Raises:
        KeyError: If table is not found in any worksheet

    """
    from openpyxl.worksheet.worksheet import Worksheet as Ws

    for worksheet in workbook.worksheets:
        if isinstance(worksheet, Ws) and hasattr(worksheet, "tables"):
            for table in worksheet.tables.values():
                if table.name == table_name:
                    logger.debug(
                        "Found table '%s' in worksheet '%s'",
                        table_name,
                        worksheet.title,
                    )
                    return worksheet, table

    msg = f"Table '{table_name}' not found in workbook"
    raise KeyError(msg)
