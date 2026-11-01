"""Guards that keep a fill from producing a workbook Excel cannot open.

Every test here fills a template with both engines, because the defects it
covers were either reported by one engine only (openpyxl wrote an unparseable
package) or silently produced different numbers per engine (an A1-style
calculated column reused the first row's formula).  Where the fix is a
warning, both engines must raise it; where it is an error, both must refuse and
leave the destination alone.

The win32com cases skip themselves when Excel is not available.
"""

from __future__ import annotations

import contextlib
from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn, TableFormula

from fillxl import CellValueError, FillXLWarning, fill
from tests._engines import ENGINES, _fill_output, _win32com_available

if TYPE_CHECKING:
    from pathlib import Path

    from fillxl import Engine


def _previous_output(path: Path) -> bytes:
    """Write a valid workbook where the next fill will write, and return its bytes."""
    wb = openpyxl.Workbook()
    wb.active.title = "Yesterday"
    wb.active["A1"] = "last week's report"
    wb.save(path)
    wb.close()
    return path.read_bytes()


def _staging_files(directory: Path) -> list[Path]:
    return sorted(directory.glob(".*.fillxl*"))


# ═══════════════════════════════════════════════════════════════════════════════
# Calculated columns: every row gets its own row references
# ═══════════════════════════════════════════════════════════════════════════════


def _calculated_column_template(path: Path, definition: str = "B5*C5") -> None:
    """``Sales!A4:D7`` whose ``Total`` column is the calculated column *definition*.

    The body holds what Excel itself writes for such a column: the definition
    spelled out for each row (``=B5*C5``, ``=B6*C6``, …), never the same text
    three times.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    for index, header in enumerate(["Name", "Qty", "Price", "Total"], start=1):
        ws.cell(row=4, column=index, value=header)
    for row in range(5, 8):
        ws.cell(row=row, column=1, value=f"old{row}")
        ws.cell(row=row, column=2, value=row)
        ws.cell(row=row, column=3, value=row)
        ws.cell(row=row, column=4, value=f"={definition.replace('5', str(row))}")
    ws.add_table(
        Table(
            displayName="Sales",
            ref="A4:D7",
            tableColumns=[
                TableColumn(id=1, name="Name"),
                TableColumn(id=2, name="Qty"),
                TableColumn(id=3, name="Price"),
                TableColumn(
                    id=4,
                    name="Total",
                    calculatedColumnFormula=TableFormula(attr_text=definition),
                ),
            ],
        ),
    )
    wb.save(path)
    wb.close()


_ORDERS = [
    {"Name": "Widget", "Qty": 10, "Price": 4.99},
    {"Name": "Gadget", "Qty": 3, "Price": 9.99},
    {"Name": "Doohickey", "Qty": 7, "Price": 0.5},
]


@pytest.mark.parametrize("engine", ENGINES)
def test_a_calculated_column_writes_each_row_its_own_formula(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """An A1-style definition must move down with the row, as Excel's fill does.

    openpyxl used to copy ``=B5*C5`` into every row, so every row computed the
    first row's total; win32com was right, and the docs promise both engines
    write the same workbook.
    """
    template = tmp_path / "template.xlsx"
    _calculated_column_template(template)

    output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": _ORDERS})

    wb = openpyxl.load_workbook(output)
    try:
        totals = [wb["Sales"].cell(row=row, column=4).value for row in (5, 6, 7)]
    finally:
        wb.close()

    assert totals == ["=B5*C5", "=B6*C6", "=B7*C7"]


def _excel_computed_values(path: Path, column: str, first_row: int, count: int) -> list[object]:
    """Open *path* in Excel, calculate it, and read *count* cells of *column*."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.DisplayAlerts = False
    try:
        workbook = app.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        try:
            app.Calculate()
            sheet = workbook.Worksheets(1)
            return [
                sheet.Range(f"{column}{row}").Value for row in range(first_row, first_row + count)
            ]
        finally:
            workbook.Close(SaveChanges=False)
    finally:
        with contextlib.suppress(Exception):
            app.Quit()


@pytest.mark.excel
def test_excel_computes_a_different_total_for_every_row(tmp_path: Path) -> None:
    """The formulas must be right, not merely different: check what Excel computes."""
    if not _win32com_available():  # pragma: no cover - the excel marker skips this
        pytest.skip("needs Windows with Microsoft Excel")

    template = tmp_path / "template.xlsx"
    _calculated_column_template(template)
    output = tmp_path / "out.xlsx"
    fill(template=template, output=output, engine="openpyxl", tables={"Sales": _ORDERS})

    values = _excel_computed_values(output, "D", 5, 3)

    assert values == pytest.approx([49.9, 29.97, 3.5])


def _excel_template_with_structured_column(path: Path) -> None:
    """Build the structured-reference template in Excel itself.

    openpyxl cannot author a structured calculated column that Excel accepts,
    so the template comes from Excel and only the *fill* is under test.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.DisplayAlerts = False
    try:
        workbook = app.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        sheet.Name = "Sales"
        for index, header in enumerate(["Name", "Qty", "Price", "Total"], start=1):
            sheet.Cells(4, index).Value = header
        for row in range(5, 8):
            sheet.Cells(row, 1).Value = f"old{row}"
            sheet.Cells(row, 2).Value = row
            sheet.Cells(row, 3).Value = row
        table = sheet.ListObjects.Add(1, sheet.Range("A4:D7"))
        table.Name = "Sales"
        table.ListColumns("Total").DataBodyRange.FormulaR1C1 = "=[@Qty]*[@Price]"
        workbook.SaveAs(str(path.resolve()), 51)
        workbook.Close(False)
    finally:
        with contextlib.suppress(Exception):
            app.Quit()


@pytest.mark.excel
def test_a_structured_calculated_column_is_written_unchanged(tmp_path: Path) -> None:
    """Structured references already resolve per row, so they must not be shifted."""
    if not _win32com_available():  # pragma: no cover - the excel marker skips this
        pytest.skip("needs Windows with Microsoft Excel")

    template = tmp_path / "template.xlsx"
    _excel_template_with_structured_column(template)

    output = _fill_output(template, tmp_path / "out.xlsx", "openpyxl", tables={"Sales": _ORDERS})

    wb = openpyxl.load_workbook(output)
    try:
        totals = [wb["Sales"].cell(row=row, column=4).value for row in (5, 6, 7)]
    finally:
        wb.close()

    assert totals == ["=Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]"] * 3


# ═══════════════════════════════════════════════════════════════════════════════
# Formulas next to the table follow it, as they do in Excel
# ═══════════════════════════════════════════════════════════════════════════════


def _summary_template(path: Path) -> None:
    """A one-body-row table plus formulas that do and do not track it."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A3"] = "Name"
    ws["B3"] = "Value"
    ws["A4"] = "one"
    ws["B4"] = 1
    ws["D3"] = "=SUM(B4:B4)"  # tracks the body
    ws["E3"] = "=SUM($B$4:$B$4)"  # tracks the body, anchored
    ws["F3"] = "=SUM(B4:B5)"  # one row too many
    ws["G3"] = "=B4*2"  # a single cell is not a range
    ws["H3"] = "=SUM(Z1:Z1)"  # nothing to do with the table
    ws["I3"] = '=IF(A1>0,"B4:B4",SUM(B4:B4))'  # a string that looks like a range
    ws.add_table(
        Table(
            displayName="T",
            ref="A3:B4",
            tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
        ),
    )
    other = wb.create_sheet("Other")
    other["A1"] = "=SUM(Data!B4:B4)"  # on another sheet, but it tracks the body
    wb.save(path)
    wb.close()


def _formulas(path: Path, cells: tuple[str, ...]) -> dict[str, object]:
    wb = openpyxl.load_workbook(path)
    try:
        sheet = wb["Data"]
        found = {cell: sheet[cell].value for cell in cells}
        found["Other!A1"] = wb["Other"]["A1"].value
    finally:
        wb.close()
    return found


@pytest.mark.parametrize("engine", ENGINES)
def test_a_formula_that_tracks_the_table_grows_with_it(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """Excel rewrites the range on resize; the openpyxl engine has to match."""
    template = tmp_path / "template.xlsx"
    _summary_template(template)

    output = _fill_output(
        template,
        tmp_path / "out.xlsx",
        engine,
        tables={"T": [{"Name": f"n{index}", "Value": index} for index in range(1, 5)]},
    )

    assert _formulas(output, ("D3", "E3", "F3", "G3", "H3", "I3")) == {
        "D3": "=SUM(B4:B7)",
        "E3": "=SUM($B$4:$B$7)",
        "F3": "=SUM(B4:B5)",
        "G3": "=B4*2",
        "H3": "=SUM(Z1:Z1)",
        "I3": '=IF(A1>0,"B4:B4",SUM(B4:B7))',
        "Other!A1": "=SUM(Data!B4:B7)",
    }


@pytest.mark.parametrize("engine", ENGINES)
def test_a_formula_that_tracks_the_table_shrinks_with_it(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """The same rule in reverse: the range goes back to the remaining body rows."""
    template = tmp_path / "template.xlsx"
    _summary_template(template)
    grown = _fill_output(
        template,
        tmp_path / "grown.xlsx",
        engine,
        tables={"T": [{"Name": f"n{index}", "Value": index} for index in range(1, 5)]},
    )

    shrunk = _fill_output(
        grown,
        tmp_path / "shrunk.xlsx",
        engine,
        tables={"T": [{"Name": "only", "Value": 1}]},
    )

    assert _formulas(shrunk, ("D3", "E3", "F3", "G3", "H3")) == {
        "D3": "=SUM(B4:B4)",
        "E3": "=SUM($B$4:$B$4)",
        "F3": "=SUM(B4:B5)",
        "G3": "=B4*2",
        "H3": "=SUM(Z1:Z1)",
        "Other!A1": "=SUM(Data!B4:B4)",
    }


# ═══════════════════════════════════════════════════════════════════════════════
# A table header cannot be renamed through a defined name
# ═══════════════════════════════════════════════════════════════════════════════


def _header_name_template(path: Path) -> None:
    """A table whose ``Name`` header is also reachable as the name ``HeaderCell``."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "Name"
    ws["B1"] = "Value"
    ws["A2"] = "sample"
    ws["B2"] = 1
    ws.add_table(
        Table(
            displayName="T",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
        ),
    )
    wb.defined_names.add(DefinedName("HeaderCell", attr_text="Report!$A$1"))
    wb.save(path)
    wb.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_renaming_a_table_header_through_a_name_is_refused(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """The cell and the table's column name would disagree, and Excel repairs that.

    openpyxl wrote the header text and left ``tableColumns`` behind, which made
    the file unopenable; Excel renames the column, which then invalidates the
    data-to-column mapping the same fill is using.  Both engines refuse.
    """
    template = tmp_path / "template.xlsx"
    _header_name_template(template)
    output = tmp_path / "out.xlsx"
    before = _previous_output(output)

    with pytest.raises(ValueError, match="header cell of table 'T'"):
        fill(template=template, output=output, engine=engine, values={"HeaderCell": "Renamed"})

    assert output.read_bytes() == before
    assert _staging_files(tmp_path) == []
    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Yesterday"]["A1"].value == "last week's report"
    finally:
        wb.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_a_data_cell_of_a_table_can_still_be_filled_through_a_name(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """Only the header row is protected; the body of a table is ordinary cells."""
    template = tmp_path / "template.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "Name"
    ws["A2"] = "sample"
    ws.add_table(
        Table(displayName="T", ref="A1:A2", tableColumns=[TableColumn(id=1, name="Name")]),
    )
    wb.defined_names.add(DefinedName("BodyCell", attr_text="Report!$A$2"))
    wb.save(template)
    wb.close()

    output = _fill_output(template, tmp_path / "out.xlsx", engine, values={"BodyCell": "filled"})

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Report"]["A2"].value == "filled"
    finally:
        wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# Text XML cannot hold, and numbers Excel cannot hold
# ═══════════════════════════════════════════════════════════════════════════════


def _table_template(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Name"
    ws["B1"] = "Value"
    ws["A2"] = "placeholder"
    ws["B2"] = 0
    ws.add_table(
        Table(
            displayName="T",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
        ),
    )
    wb.save(path)
    wb.close()


@pytest.mark.parametrize("engine", ENGINES)
@pytest.mark.parametrize("character", ["\ufffe", "\uffff"])
def test_a_character_xml_forbids_fails_the_fill_and_keeps_the_output(
    tmp_path: Path,
    engine: Engine,
    character: str,
) -> None:
    """openpyxl wrote the raw character and the workbook could not be parsed again.

    Excel does not store the character either — it rewrites it as the literal
    text ``_xFFFE_`` — so neither engine can round-trip it.  Both refuse.
    """
    template = tmp_path / "template.xlsx"
    _table_template(template)
    output = tmp_path / "out.xlsx"
    before = _previous_output(output)

    with pytest.raises(CellValueError, match=r"U\+FFF[EF]"):
        fill(
            template=template,
            output=output,
            engine=engine,
            tables={"T": [{"Name": f"a{character}b", "Value": 1}]},
        )

    assert output.read_bytes() == before
    assert _staging_files(tmp_path) == []
    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Yesterday"]["A1"].value == "last week's report"
    finally:
        wb.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_other_noncharacters_are_written_and_reload(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """The fix must not reject code points XML allows, which openpyxl round-trips."""
    template = tmp_path / "template.xlsx"
    _table_template(template)
    text = "a\ufdd0 b\ufdef c\U0001fffe d\U0010ffff e\u2028"

    output = _fill_output(
        template,
        tmp_path / "out.xlsx",
        engine,
        tables={"T": [{"Name": text, "Value": 1}]},
    )

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Data"]["A2"].value == text
    finally:
        wb.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_an_infinity_is_reported_instead_of_vanishing(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """An empty cell is the only thing Excel can hold, but it is still data loss."""
    template = tmp_path / "template.xlsx"
    _table_template(template)

    with pytest.warns(FillXLWarning, match="Excel cannot store infinity") as caught:
        output = _fill_output(
            template,
            tmp_path / "out.xlsx",
            engine,
            tables={"T": [{"Name": "x", "Value": float("-inf")}]},
        )

    assert len(caught) == 1
    assert "column 'Value'" in str(caught[0].message)
    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Data"]["B2"].value is None
    finally:
        wb.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_nan_is_still_a_missing_value_without_a_warning(
    tmp_path: Path,
    engine: Engine,
    recwarn: pytest.WarningsRecorder,
) -> None:
    """Every data library spells "missing" as NaN, so it must stay silent."""
    template = tmp_path / "template.xlsx"
    _table_template(template)

    output = _fill_output(
        template,
        tmp_path / "out.xlsx",
        engine,
        tables={"T": [{"Name": "x", "Value": float("nan")}]},
    )

    assert [w for w in recwarn if issubclass(w.category, FillXLWarning)] == []
    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Data"]["B2"].value is None
    finally:
        wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# Defined names resolve like Excel's do
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("engine", ENGINES)
def test_a_defined_name_is_found_whatever_its_case(tmp_path: Path, engine: Engine) -> None:
    """Excel and COM resolve names case-insensitively; openpyxl used to raise."""
    template = tmp_path / "template.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Report"
    wb.active["A1"] = "before"
    wb.defined_names.add(DefinedName("ReportTitle", attr_text="Report!$A$1"))
    wb.save(template)
    wb.close()

    output = _fill_output(template, tmp_path / "out.xlsx", engine, values={"reporttitle": "Q4"})

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Report"]["A1"].value == "Q4"
    finally:
        wb.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_a_sheet_scoped_name_still_wins_over_a_global_name_of_another_case(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """An exact match keeps its meaning; the case-insensitive walk is only a fallback."""
    template = tmp_path / "template.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "global"
    ws["A2"] = "scoped"
    wb.defined_names.add(DefinedName("Title", attr_text="Report!$A$1"))
    ws.defined_names.add(DefinedName("title", attr_text="Report!$A$2"))
    wb.save(template)
    wb.close()

    output = _fill_output(template, tmp_path / "out.xlsx", engine, values={"title": "scoped"})

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Report"]["A2"].value == "scoped"
    finally:
        wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# Conditional formatting over the header row
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("engine", ENGINES)
def test_a_conditional_formatting_rule_covering_the_header_row_follows_the_table(
    tmp_path: Path,
    engine: Engine,
) -> None:
    """Excel extends a rule that ends on the table's last row, even from row 1."""
    template = tmp_path / "template.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Name"
    ws["B1"] = "Value"
    ws["A2"] = "one"
    ws["B2"] = 1
    ws.conditional_formatting.add(
        "B1:B2",
        CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill(fill_type="solid")),
    )
    ws.add_table(
        Table(
            displayName="T",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
        ),
    )
    wb.save(template)
    wb.close()

    output = _fill_output(
        template,
        tmp_path / "out.xlsx",
        engine,
        tables={"T": [{"Name": f"n{index}", "Value": index} for index in range(4)]},
    )

    wb = openpyxl.load_workbook(output)
    try:
        ranges = sorted(str(entry.sqref) for entry in wb["Data"].conditional_formatting)
    finally:
        wb.close()

    assert ranges == ["B1:B5"]
