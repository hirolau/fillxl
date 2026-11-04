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


def unquote_sheet_name(text: str) -> str:
    """*text* without the quotes Excel puts around a sheet name.

    Excel quotes a name that is not a plain word and doubles a quote inside
    it, so ``'Bob''s Data'`` is the sheet ``Bob's Data``.

    Examples:
        >>> unquote_sheet_name("'My Sheet'")
        'My Sheet'
        >>> unquote_sheet_name("'It''s'")
        "It's"
        >>> unquote_sheet_name("Data")
        'Data'

    """
    if len(text) >= 2 and text.startswith("'") and text.endswith("'"):
        return text[1:-1].replace("''", "'")
    return text


def workbook_uses_1904(workbook: Workbook) -> bool:
    """Whether *workbook* counts days from 1904-01-01 (the Mac date system).

    Examples:
        >>> from openpyxl import Workbook
        >>> workbook_uses_1904(Workbook())
        False

    """
    from openpyxl.utils.datetime import MAC_EPOCH

    return bool(workbook.epoch == MAC_EPOCH)


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
                # A table part has both a name and a display name; Excel shows
                # the display name, so that is what a user asks to fill.
                if (table.displayName or table.name) == table_name:
                    logger.debug(
                        "Found table '%s' in worksheet '%s'",
                        table_name,
                        worksheet.title,
                    )
                    return worksheet, table

    msg = f"Table '{table_name}' not found in workbook"
    raise KeyError(msg)
