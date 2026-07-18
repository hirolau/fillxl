"""Utility functions for Excel manipulation."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.table import Table
    from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


def col2num(col: str) -> int:
    """Convert Excel column letter(s) to 1-based column number.

    .. deprecated::
        Use ``openpyxl.utils.column_index_from_string`` instead —
        this function is kept for backward compatibility only.

    Args:
        col: Column letter(s) (e.g., 'A', 'Z', 'AA')

    Returns:
        1-based column number

    Examples:
        >>> col2num('A')
        1
        >>> col2num('Z')
        26
        >>> col2num('AA')
        27

    """
    num = 0
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord("A")) + 1
    return num


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


def get_defined_names(workbook: Workbook) -> list[str]:
    """Get all defined names (named cells + table names) from a workbook.

    Args:
        workbook: openpyxl Workbook instance

    Returns:
        List of all defined name strings

    """
    names = {str(k) for k in workbook.defined_names}

    # Add all table names
    for worksheet in workbook.worksheets:
        if hasattr(worksheet, "tables"):
            for table in worksheet.tables.values():
                names.add(table.name)

    return sorted(names)
