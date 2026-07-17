"""Table filling integration tests — unique coverage not in story suite.

Tests style copying behaviour that the story tests don't cover:
font/fill inheritance and the copy_table_styles=False config flag.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl

from fillxl import FillXLConfig, fill
from tests.conftest import TEMPLATE

if TYPE_CHECKING:
    from pathlib import Path


def _data_rows(count: int) -> list[dict[str, object]]:
    """Return standard data rows."""
    return [
        {
            "Col1": f"row{i}",
            "Col2": i,
            "ColNotFilled": f"nf{i}",
            "ColFormula": None,
            "TestFormattedPct": 0.0,
        }
        for i in range(1, count + 1)
    ]


def test_new_rows_inherit_font_and_fill_from_first_data_row(output_path: Path) -> None:
    """New rows should have the same font, fill, and number format as row 1."""
    fill(
        template=TEMPLATE,
        output=output_path,
        tables={"TableTest": _data_rows(5)},
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["TestTable"]
        first_row_cells = [ws.cell(row=2, column=c) for c in range(1, 6)]

        for row in range(3, 7):
            for col in range(1, 6):
                first = first_row_cells[col - 1]
                cell = ws.cell(row=row, column=col)
                assert cell.font.name == first.font.name, (
                    f"Font name mismatch at {col}{row}: {cell.font.name} != {first.font.name}"
                )
                assert cell.font.size == first.font.size, f"Font size mismatch at {col}{row}"
                assert cell.fill.fgColor.rgb == first.fill.fgColor.rgb, (
                    f"Fill color mismatch at {col}{row}"
                )
    finally:
        wb.close()


def test_new_rows_do_not_inherit_styles_when_copy_table_styles_is_false(output_path: Path) -> None:
    """copy_table_styles=False should not copy styles to new rows."""
    fill(
        template=TEMPLATE,
        output=output_path,
        tables={"TableTest": _data_rows(5)},
        config=FillXLConfig(copy_table_styles=False),
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["TestTable"]
        assert ws.cell(row=6, column=1).value == "row5"
        assert ws.cell(row=6, column=2).value == 5

        new_row_fill = ws.cell(row=6, column=1).fill
        assert new_row_fill.fgColor.rgb == "00000000", (
            f"Expected default fill on new row, got {new_row_fill.fgColor.rgb}"
        )
        assert new_row_fill.patternType is None, (
            f"Expected no fill pattern on new row, got {new_row_fill.patternType}"
        )
    finally:
        wb.close()
