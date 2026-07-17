"""Unit tests for _utils.py helper functions."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from fillxl._utils import col2num, col_to_excel, find_worksheet_and_table, get_defined_names

TEMPLATE = Path(__file__).parent.parent / "assets" / "test_template_1.xlsx"


class TestUtils:
    """Unit tests for _utils.py helper functions."""

    def test_col_to_excel_and_col2num_are_inverse_operations(self) -> None:
        """col_to_excel and col2num are inverse operations."""
        test_numbers = [1, 2, 3, 10, 26, 27, 52, 100, 702, 1000, 16384]
        for num in test_numbers:
            col_letter = col_to_excel(num)
            assert col2num(col_letter) == num

    def test_get_defined_names_returns_both_named_cells_and_table_names(self) -> None:
        """get_defined_names returns both named cells and table names."""
        wb = openpyxl.load_workbook(TEMPLATE)
        try:
            names = get_defined_names(wb)
            assert "single_cell" in names
            assert "TableTest" in names
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
