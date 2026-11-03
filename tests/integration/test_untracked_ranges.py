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
from tests.ooxml._patch import patch_part

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
    """Excel rewrites it when the ListObject is resized; the warning would be false."""
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
def test_a_sheet_scoped_name_over_the_body_is_reported(engine: Engine, tmp_path: Path) -> None:
    """A name scoped to a sheet is not in the workbook's own collection."""
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    worksheet = workbook["Data"]
    worksheet.defined_names.add(
        DefinedName("LocalBody", attr_text="Data!$A$2:$B$2", localSheetId=0),
    )
    workbook.save(template)
    workbook.close()

    with pytest.warns(UntrackedRangeWarning, match="LocalBody"):
        _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_name_on_a_sheet_with_an_apostrophe_is_reported(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """Excel doubles the quote inside a quoted name; the sheet title has one."""
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    worksheet = workbook.active
    worksheet.title = "Bob's Data"
    workbook.defined_names.add(
        DefinedName("SalesBody", attr_text="'Bob''s Data'!$A$2:$B$2"),
    )
    workbook.save(template)
    workbook.close()

    with pytest.warns(UntrackedRangeWarning, match="Bob's Data"):
        _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_whole_row_and_whole_column_names_do_not_break_a_resize(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """``Print_Titles`` is ``$1:$1`` and ``Print_Area`` can be ``$A:$A``.

    Neither is a rectangle a table resize can leave stale, and neither may be
    turned into one by mistake.
    """
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    workbook.defined_names.add(DefinedName("_xlnm.Print_Titles", attr_text="Data!$1:$1"))
    workbook.defined_names.add(DefinedName("_xlnm.Print_Area", attr_text="Data!$A:$A"))
    workbook.defined_names.add(DefinedName("Constant", attr_text='"hello"'))
    workbook.save(template)
    workbook.close()

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    assert not [record for record in caught if record.category is UntrackedRangeWarning]
    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook["Data"]["A3"].value == "south"
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


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_same_sheet_validation_list_is_reported_without_crashing(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """A list source on the validation's own sheet is stored unqualified.

    Regression: parsing it as a defined name raised ``AttributeError`` from
    inside the warning, so every resize of such a workbook failed.
    """
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    worksheet = workbook["Data"]
    validation = DataValidation(type="list", formula1="$A$2:$A$2", allow_blank=True)
    worksheet.add_data_validation(validation)
    validation.add("C2")
    workbook.save(template)
    workbook.close()

    with pytest.warns(UntrackedRangeWarning, match="data validation"):
        output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook["Data"]["A3"].value == "south"
    finally:
        workbook.close()


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_custom_validation_rule_is_not_mistaken_for_a_range(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """``A1>0`` is an expression, not a rectangle that can go stale."""
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    worksheet = workbook["Data"]
    validation = DataValidation(type="custom", formula1="A1>0", allow_blank=True)
    worksheet.add_data_validation(validation)
    validation.add("C2")
    workbook.save(template)
    workbook.close()

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    assert not [record for record in caught if record.category is UntrackedRangeWarning]


SHEET_PART = "xl/worksheets/sheet1.xml"

# What Excel writes when a cross-sheet list source is added through the newer
# data-validation UI: the reference lives in an ``xm:f`` child inside ``extLst``.
X14_VALIDATION = (
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" '
    'count="1"><x14:dataValidation type="list" allowBlank="1">'
    "<x14:formula1><xm:f>Data!$A$2:$B$2</xm:f></x14:formula1>"
    "<xm:sqref>C2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>"
)


def test_an_x14_validation_list_is_reported_by_fillxl(tmp_path: Path) -> None:
    """The reference of an ``extLst`` validation is markup, not plain text."""
    template = _table_template(tmp_path / "template.xlsx")
    patch_part(
        template,
        SHEET_PART,
        lambda xml: xml.replace("</worksheet>", f"{X14_VALIDATION}</worksheet>"),
    )

    with pytest.warns(UntrackedRangeWarning, match="data validation"):
        _fill_output(template, tmp_path / "out.xlsx", "fillxl", tables={"Sales": ROWS})


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
