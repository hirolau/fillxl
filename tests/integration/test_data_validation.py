"""Data-validation ranges that track a table, as Excel extends them.

A dropdown applied to a table column stops at the template's last row unless
the fill extends it; Excel extends it on resize, and the offline engines have
to do the same.
"""

from __future__ import annotations

import contextlib
import importlib
from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import FillXLConfig
from tests._engines import _fill_output, _win32com_available

if TYPE_CHECKING:
    from collections.abc import Iterator
    from pathlib import Path
    from typing import Any

    from fillxl import Engine

ROWS = [{"Name": f"n{index}", "Value": index} for index in range(1, 5)]
OFFLINE_ENGINES = ["openpyxl", "fillxl"]


def _table_with_dropdown_template(path: Path, *, dropdown_cell: str = "B2") -> Path:
    """A one-row table with a list dropdown on the ``Value`` column."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"], worksheet["B1"] = "Name", "Value"
    worksheet["A2"], worksheet["B2"] = "row2", 2
    worksheet.add_table(
        Table(
            displayName="Sales",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
        ),
    )
    validation = DataValidation(type="list", formula1='"one,two,three"', allow_blank=True)
    worksheet.add_data_validation(validation)
    validation.add(dropdown_cell)
    workbook.save(path)
    workbook.close()
    return path


def _dropdown_ranges(path: Path) -> list[str]:
    """The ``sqref`` of every data validation of the sheet, as written."""
    workbook = openpyxl.load_workbook(path)
    try:
        return [
            " ".join(sorted(str(area) for area in validation.sqref.ranges))
            for validation in workbook["Data"].data_validations.dataValidation
        ]
    finally:
        workbook.close()


@pytest.mark.parametrize("engine", OFFLINE_ENGINES)
def test_a_dropdown_on_a_table_column_reaches_every_filled_row(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """A dropdown over the last template row has to cover the rows the fill adds."""
    template = _table_with_dropdown_template(tmp_path / "template.xlsx")

    output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    assert _dropdown_ranges(output) == ["B2:B5"]


@pytest.mark.parametrize("engine", OFFLINE_ENGINES)
def test_a_dropdown_outside_the_table_columns_is_left_alone(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """A dropdown that only happens to sit on the table's rows is the author's."""
    template = _table_with_dropdown_template(tmp_path / "template.xlsx", dropdown_cell="D2")

    output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    assert _dropdown_ranges(output) == ["D2"]


@pytest.mark.parametrize("engine", OFFLINE_ENGINES)
def test_turning_off_data_validation_extension_leaves_dropdowns_as_they_were(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """The flag opts a template out of the extension entirely."""
    template = _table_with_dropdown_template(tmp_path / "template.xlsx")

    output = _fill_output(
        template,
        tmp_path / "out.xlsx",
        engine,
        tables={"Sales": ROWS},
        config=FillXLConfig(extend_data_validation=False),
    )

    assert _dropdown_ranges(output) == ["B2"]


@contextlib.contextmanager
def _excel() -> Iterator[Any]:
    """A hidden Excel instance for the duration of one test."""
    if not _win32com_available():
        pytest.skip("win32com backend requires Windows and Excel")
    pythoncom = importlib.import_module("pythoncom")
    win32com_client = importlib.import_module("win32com.client")
    pythoncom.CoInitialize()
    app = win32com_client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        yield app
    finally:
        with contextlib.suppress(Exception):
            app.Quit()


@pytest.mark.win32com
def test_excel_extends_a_dropdown_to_the_filled_rows(tmp_path: Path) -> None:
    """If Excel did not extend it, the offline engines must not either."""
    template = _table_with_dropdown_template(tmp_path / "template.xlsx")
    output = _fill_output(template, tmp_path / "out.xlsx", "win32com", tables={"Sales": ROWS})

    with _excel() as app:
        workbook = app.Workbooks.Open(str(output))
        try:
            validation = workbook.Worksheets("Data").Range("B5").Validation
            # xlValidateList; xlValidateInputOnly (0) means "no validation here".
            assert validation.Type == 3
        finally:
            workbook.Close(SaveChanges=False)
