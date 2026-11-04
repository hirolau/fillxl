"""Unit tests for _utils.py helper functions."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from fillxl._utils import find_worksheet_and_table

TEMPLATE = Path(__file__).parent.parent / "assets" / "test_template_1.xlsx"


class TestUtils:
    """Unit tests for _utils.py helper functions."""

    def test_find_worksheet_and_table_finds_the_worksheet_holding_the_table(self) -> None:
        """find_worksheet_and_table returns the worksheet and table by name."""
        wb = openpyxl.load_workbook(TEMPLATE)
        try:
            worksheet, table = find_worksheet_and_table(wb, "TableTest")
            assert worksheet.title == "TestTable"
            assert table.displayName == "TableTest"
        finally:
            wb.close()

    def test_find_worksheet_and_table_raises_keyerror_for_unknown_table(self) -> None:
        """find_worksheet_and_table raises KeyError for unknown table."""
        wb = openpyxl.load_workbook(TEMPLATE)
        try:
            with pytest.raises(KeyError):
                find_worksheet_and_table(wb, "NonexistentTable")
        finally:
            wb.close()
