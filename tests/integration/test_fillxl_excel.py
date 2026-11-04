"""Validate the ``fillxl`` engine against real Excel.

These tests author the template **in Excel**, so the dynamic-array markup is the
markup Excel really writes, then fill it with ``engine="fillxl"`` and open the
result in Excel again.  A file Excel refuses to open is a hard failure — that is
exactly what a user would see.

They need Windows with Microsoft Excel installed, and are skipped otherwise
(``-m "not win32com"`` keeps them out of the CI run).
"""

from __future__ import annotations

import zipfile
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from fillxl import FillXLConfig, fill
from tests._engines import _win32com_available

pytestmark = [pytest.mark.excel, pytest.mark.win32com]

if not _win32com_available():  # pragma: no cover - depends on the machine
    pytest.skip("needs Windows with Microsoft Excel", allow_module_level=True)

# A real Excel-authored, macro-enabled template from the repository's assets:
# table ``TableTest`` and a spilled ``=TableTest[Col1]`` on ``SpilledRange``.
XLSM_TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "test_template_1_slsm.xlsm"

ROWS = [
    {"Region": "East", "Amount": 30},
    {"Region": "West", "Amount": 40},
    {"Region": "North", "Amount": 55},
    {"Region": "South", "Amount": 60},
]

_XL_SRC_RANGE = 1
_XL_YES = 1
_XL_OPENXML_WORKBOOK = 51


@pytest.fixture(scope="module")
def excel() -> Iterator[Any]:
    """A hidden Excel instance that never asks the user anything."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    try:
        yield app
    finally:
        app.Quit()
        pythoncom.CoUninitialize()


def _build_template_in_excel(excel: Any, path: Path) -> Path:  # noqa: ANN401
    """A workbook Excel authored: totals-row table plus three dynamic arrays."""
    workbook = excel.Workbooks.Add()
    try:
        sheet = workbook.Worksheets(1)
        sheet.Name = "Data"
        sheet.Range("B2:D4").Value = (
            ("Region", "Amount", "Double"),
            ("North", 10, None),
            ("South", 20, None),
        )
        table = sheet.ListObjects.Add(_XL_SRC_RANGE, sheet.Range("B2:D4"), None, _XL_YES)
        table.Name = "Tbl"
        table.ShowTotals = True
        table.ListColumns("Double").DataBodyRange.Formula = "=[@Amount]*2"
        # Spills that do not depend on the table, and one that does.
        sheet.Range("F8").Formula2 = "=SEQUENCE(5)"
        sheet.Range("H8").Formula2 = "=SUM(F8#)"
        sheet.Range("J3").Formula2 = "=SORT(UNIQUE(Tbl[Region]))"
        excel.CalculateFullRebuild()
        workbook.SaveAs(str(path), FileFormat=_XL_OPENXML_WORKBOOK)
    finally:
        workbook.Close(False)
    return path


def test_the_template_from_excel_really_has_spills(tmp_path: Path, excel: Any) -> None:  # noqa: ANN401
    """Guard for the other tests: the fixture has to contain what they assert."""
    template = _build_template_in_excel(excel, tmp_path / "template.xlsx")

    with zipfile.ZipFile(template) as archive:
        assert "xl/metadata.xml" in archive.namelist()
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    assert 't="array"' in sheet


def test_a_fill_keeps_the_spills_alive_in_excel(tmp_path: Path, excel: Any) -> None:  # noqa: ANN401
    """The spill that is not touched keeps working, values and all."""
    template = _build_template_in_excel(excel, tmp_path / "template.xlsx")
    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})

    workbook = excel.Workbooks.Open(str(output), 0, True)
    try:
        sheet = workbook.Worksheets("Data")
        spill = sheet.Range("F8")
        assert spill.HasSpill, "the spilled array was lost"
        assert str(spill.SpillingToRange.Address) == "$F$8:$F$12"
        assert [spill.Cells(index).Value for index in range(1, 6)] == [1, 2, 3, 4, 5]
        assert sheet.Range("H8").Value == 15
        table = sheet.ListObjects("Tbl")
        assert str(table.Range.Address) == "$B$2:$D$7"
        assert str(table.TotalsRowRange.Address) == "$B$7:$D$7"
    finally:
        workbook.Close(False)


def test_a_spill_that_depends_on_the_table_grows_on_recalculation(
    tmp_path: Path,
    excel: Any,  # noqa: ANN401
) -> None:
    """The engine leaves the workbook consistent; Excel resizes the spill itself.

    This is why the editor does not need to understand ``SORT``/``UNIQUE``:
    it only has to keep ``xl/metadata.xml``, the array markup and the table
    definition intact and ask for a recalculation.
    """
    template = _build_template_in_excel(excel, tmp_path / "template.xlsx")
    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})

    workbook = excel.Workbooks.Open(str(output), 0, True)
    try:
        excel.CalculateFullRebuild()
        sheet = workbook.Worksheets("Data")
        spill = sheet.Range("J3")
        assert spill.HasSpill, "the spilled SORT/UNIQUE array was lost"
        assert str(spill.SpillingToRange.Address) == "$J$3:$J$6"
        values = [spill.Cells(index).Value for index in range(1, 5)]
        assert sorted(str(value) for value in values) == ["East", "North", "South", "West"]
        # The calculated column was filled into the new rows by the table.
        assert all(
            sheet.Cells(row, 4).Formula == "=[@Amount]*2" or sheet.Cells(row, 4).Value is not None
            for row in range(3, 7)
        )
    finally:
        workbook.Close(False)


def test_excel_reopens_the_output_without_repairing_it(tmp_path: Path, excel: Any) -> None:  # noqa: ANN401
    """A repaired workbook loses parts; a clean open is what a user needs."""
    template = _build_template_in_excel(excel, tmp_path / "template.xlsx")
    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})

    workbook = excel.Workbooks.Open(str(output), 0, True)
    try:
        # Re-saving through Excel proves the package is coherent: a workbook
        # that had to be repaired loses parts as soon as Excel writes it again.
        resaved = tmp_path / "resaved.xlsx"
        workbook.SaveCopyAs(str(resaved))
    finally:
        workbook.Close(False)

    with zipfile.ZipFile(resaved) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    assert 't="array"' in sheet


def test_a_real_macro_template_opens_with_its_spill_alive(tmp_path: Path, excel: Any) -> None:  # noqa: ANN401
    """The end-to-end goal: fill rows into a table, and Excel's spill follows.

    ``SpilledRange!B2`` holds ``=TableTest[Col1]``, so its size depends on the
    table.  The engine never touches that sheet; it only keeps the package
    coherent and asks for a recalculation, and Excel grows the spill from
    ``B2:B5`` to ``B2:B7`` by itself.
    """
    rows = [
        {"Col1": index, "Col2": index * 2, "ColNotFilled": None, "TestFormattedPct": 0.1}
        for index in range(1, 7)
    ]
    output = fill(
        XLSM_TEMPLATE,
        tmp_path / "out.xlsm",
        engine="fillxl",
        tables={"TableTest": rows},
        config=FillXLConfig(),
    )

    workbook = excel.Workbooks.Open(str(output), 0, True)
    try:
        sheet = workbook.Worksheets("SpilledRange")
        spill = sheet.Range("B2")
        assert spill.HasSpill, "the spilled array was lost"
        excel.CalculateFullRebuild()
        assert str(spill.SpillingToRange.Address) == "$B$2:$B$7"
        assert [spill.Cells(index).Value for index in range(1, 7)] == [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]
    finally:
        workbook.Close(False)
