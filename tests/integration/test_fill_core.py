"""Core fill() integration tests — unique coverage not in story suite.

Tests behaviours that the story tests don't cover: idempotency.
Nested output directory creation is covered by ``tests/story/test_01_basics.py``.
"""

from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from fillxl import fill
from tests.conftest import TEMPLATE

if TYPE_CHECKING:
    from pathlib import Path


class TestNamedCellScope:
    """Defined names are resolved the same way in both engines."""

    def test_a_sheet_scoped_defined_name_can_be_filled(self, output_path: Path) -> None:
        """Sheet-scoped names live on the worksheet, not on the workbook.

        win32com resolves them through ``Workbook.Names``; the openpyxl
        backend has to look at ``worksheet.defined_names`` as well.
        """
        template = output_path.with_name("sheet_scoped.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["B2"] = "old"
        ws.defined_names.add(DefinedName("scoped_title", attr_text="Data!$B$2"))
        wb.save(template)
        wb.close()

        fill(template=template, output=output_path, values={"scoped_title": "Scoped"})

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Data"]["B2"].value == "Scoped"
        finally:
            wb.close()


def test_fill_is_idempotent_when_applied_twice(output_path: Path) -> None:
    """Applying fill() twice with the same data should produce the same result."""
    data = [
        {
            "Col1": f"row{i}",
            "Col2": i,
            "ColNotFilled": f"nf{i}",
            "ColFormula": None,
            "TestFormattedPct": 0.0,
        }
        for i in range(1, 4)
    ]

    first_output = output_path.with_name("first.xlsx")
    fill(
        template=TEMPLATE,
        output=first_output,
        tables={"TableTest": deepcopy(data)},
    )

    second_output = output_path.with_name("second.xlsx")
    fill(
        template=TEMPLATE,
        output=second_output,
        tables={"TableTest": deepcopy(data)},
    )

    wb1 = openpyxl.load_workbook(first_output)
    wb2 = openpyxl.load_workbook(second_output)
    try:
        ws1 = wb1["TestTable"]
        ws2 = wb2["TestTable"]
        for row in range(1, 7):
            for col in range(1, 6):
                assert ws1.cell(row=row, column=col).value == ws2.cell(row=row, column=col).value, (
                    f"Mismatch at ({col}, {row}): {ws1.cell(row=row, column=col).value!r} vs "
                    f"{ws2.cell(row=row, column=col).value!r}"
                )
    finally:
        wb1.close()
        wb2.close()
