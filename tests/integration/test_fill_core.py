"""Core fill() integration tests — unique coverage not in story suite.

Tests behaviours that the story tests don't cover: idempotency.
Nested output directory creation is covered by ``tests/story/test_01_basics.py``.
"""

from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

import openpyxl

from fillxl import fill
from tests.conftest import TEMPLATE

if TYPE_CHECKING:
    from pathlib import Path


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
