"""The win32com engine must not silently truncate a save to ``.xls``.

BIFF8 holds 65 536 rows by 256 columns, and Excel's compatibility checker —
the only thing that would say so — is disabled while a fill runs.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from fillxl._backends._win32com import _refuse_truncation_to_xls


class _FakeRange:
    def __init__(self, row: int, rows: int, column: int, columns: int) -> None:
        self.Row = row
        self.Rows = _FakeCount(rows)
        self.Column = column
        self.Columns = _FakeCount(columns)


class _FakeCount:
    def __init__(self, count: int) -> None:
        self.Count = count


class _FakeWorksheet:
    def __init__(self, name: str, used: _FakeRange) -> None:
        self.Name = name
        self.UsedRange = used


class _FakeWorkbook:
    def __init__(self, *sheets: _FakeWorksheet) -> None:
        self.Worksheets: Any = list(sheets)


def test_a_workbook_that_fits_is_saved() -> None:
    workbook = _FakeWorkbook(_FakeWorksheet("Data", _FakeRange(1, 65_536, 1, 256)))

    _refuse_truncation_to_xls(workbook, Path("out.xls"))


def test_too_many_rows_is_refused_before_saving() -> None:
    workbook = _FakeWorkbook(_FakeWorksheet("Data", _FakeRange(10, 65_528, 1, 3)))

    with pytest.raises(ValueError, match="65536 rows"):
        _refuse_truncation_to_xls(workbook, Path("out.xls"))


def test_too_many_columns_is_refused_before_saving() -> None:
    workbook = _FakeWorkbook(_FakeWorksheet("Data", _FakeRange(1, 10, 200, 58)))

    with pytest.raises(ValueError, match="256 columns"):
        _refuse_truncation_to_xls(workbook, Path("out.xls"))


def test_the_message_names_the_sheet_and_the_alternative() -> None:
    workbook = _FakeWorkbook(_FakeWorksheet("Pivot", _FakeRange(1, 70_000, 1, 3)))

    with pytest.raises(ValueError, match="Pivot") as caught:
        _refuse_truncation_to_xls(workbook, Path("out.xls"))

    assert ".xlsx" in str(caught.value)
