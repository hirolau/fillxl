"""Unit tests for _utils.py helper functions."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils import column_index_from_string

from fillxl._utils import col_to_excel, find_worksheet_and_table

TEMPLATE = Path(__file__).parent.parent / "assets" / "test_template_1.xlsx"


class TestUtils:
    """Unit tests for _utils.py helper functions."""

    def test_col_to_excel_is_the_inverse_of_column_index_from_string(self) -> None:
        """col_to_excel(n) is the inverse of openpyxl's column_index_from_string."""
        test_numbers = [1, 2, 3, 10, 26, 27, 52, 100, 702, 1000, 16384]
        for num in test_numbers:
            assert column_index_from_string(col_to_excel(num)) == num

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
