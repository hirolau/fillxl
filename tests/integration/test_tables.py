"""Table filling integration tests — unique coverage not in story suite.

Tests style copying behaviour that the story tests don't cover:
font/fill inheritance and the copy_table_styles=False config flag.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
import pytest

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


def test_a_table_with_a_totals_row_is_rejected(output_path: Path) -> None:
    """Totals rows are not supported: they would be overwritten and miscounted.

    The error message has to say so, rather than silently producing a
    workbook where the totals row contains a data row.
    """
    from openpyxl.worksheet.table import Table, TableColumn

    template = output_path.with_name("totals.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Name"
    ws["B1"] = "Value"
    ws["A2"] = "a"
    ws["B2"] = 1
    ws["A3"] = "Total"
    ws["B3"] = "=SUBTOTAL(109,B2:B2)"
    ws.add_table(
        Table(
            displayName="Totals",
            ref="A1:B3",
            tableColumns=[
                TableColumn(id=1, name="Name", totalsRowLabel="Total"),
                TableColumn(id=2, name="Value", totalsRowFunction="sum"),
            ],
            totalsRowCount=1,
            totalsRowShown=True,
        ),
    )
    wb.save(template)
    wb.close()

    with pytest.raises(ValueError, match="totals row"):
        fill(
            template=template,
            output=output_path,
            tables={"Totals": [{"Name": "b", "Value": 2}]},
        )


def test_an_empty_dataframe_leaves_one_cleared_body_row(output_path: Path) -> None:
    """Empty data keeps a single empty row so the table stays valid."""
    import pandas as pd

    columns = ["Col1", "Col2", "ColNotFilled", "TestFormattedPct"]
    fill(
        template=TEMPLATE,
        output=output_path,
        tables={
            "TableTest": pd.DataFrame(columns=columns),
        },
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["TestTable"]
        table = next(iter(ws.tables.values()))
        assert table.ref == "A1:E2"
        assert table.autoFilter.ref == "A1:E2"
        assert [ws.cell(row=2, column=col).value for col in range(1, 6)] == [
            None,
            None,
            None,
            None,
            None,
        ]
    finally:
        wb.close()
