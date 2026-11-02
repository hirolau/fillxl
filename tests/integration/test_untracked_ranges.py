"""A fill warns about references it cannot resize.

Charts and A1 formulas that track a table body are rewritten.  A defined name,
a data-validation list and a pivot cache source are not — they live outside
cell formulas — so a fill reports them instead of letting a report quietly
show the template's rows.
"""

from __future__ import annotations

import warnings
from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import Engine, UntrackedRangeWarning
from tests._engines import _fill_output

if TYPE_CHECKING:
    from pathlib import Path

ROWS = [{"Name": "north", "Value": 1}, {"Name": "south", "Value": 2}]
ALL_ENGINES = ["openpyxl", "win32com", "fillxl"]


def _table_template(path: Path) -> Path:
    """A one-row table at ``A1:B2`` on sheet ``Data``."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "Name"
    worksheet["B1"] = "Value"
    worksheet["A2"] = "row2"
    worksheet["B2"] = 2
    table = Table(
        displayName="Sales",
        ref="A1:B2",
        tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
    )
    worksheet.add_table(table)
    workbook.save(path)
    return path


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_defined_name_that_covers_the_body_is_reported(engine: Engine, tmp_path: Path) -> None:
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    workbook.defined_names.add(DefinedName("SalesBody", attr_text="Data!$A$2:$B$2"))
    workbook.save(template)
    workbook.close()

    with pytest.warns(UntrackedRangeWarning, match="SalesBody"):
        _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})


def test_the_excel_engine_resizes_the_name_itself(tmp_path: Path) -> None:
    """Excel inserts rows, so it adjusts the name; the warning would be false there."""
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    workbook.defined_names.add(DefinedName("SalesBody", attr_text="Data!$A$2:$B$2"))
    workbook.save(template)
    workbook.close()

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        output = _fill_output(template, tmp_path / "out.xlsx", "win32com", tables={"Sales": ROWS})

    assert not [record for record in caught if record.category is UntrackedRangeWarning]
    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.defined_names["SalesBody"].attr_text.endswith("$A$2:$B$3")
    finally:
        workbook.close()


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_validation_list_over_the_body_is_reported(engine: Engine, tmp_path: Path) -> None:
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    worksheet = workbook["Data"]
    validation = DataValidation(type="list", formula1="Data!$A$2:$A$2", allow_blank=True)
    worksheet.add_data_validation(validation)
    validation.add("C2")
    workbook.save(template)
    workbook.close()

    with pytest.warns(UntrackedRangeWarning, match="data validation"):
        _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})


@pytest.mark.parametrize("engine", ALL_ENGINES)
def test_a_name_inside_the_table_is_left_alone(engine: Engine, tmp_path: Path) -> None:
    """A single named cell is not a body-wide range, so there is nothing stale."""
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    workbook.defined_names.add(DefinedName("SalesHeader", attr_text="Data!$A$1"))
    workbook.save(template)
    workbook.close()

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    assert not [record for record in caught if record.category is UntrackedRangeWarning]
