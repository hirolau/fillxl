"""Utility functions for Excel manipulation."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.table import Table
    from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


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
