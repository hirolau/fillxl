"""The ``fillxl`` engine: no Excel, and nothing lost.

These tests are about what makes this engine different from the others: it
copies the package and rewrites only what a fill has to change, so spilled
dynamic arrays, cell metadata and everything else the openpyxl object model
cannot represent survive untouched.  No Excel is needed to verify that — the
package itself, and openpyxl as an independent reader, say it.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import FillXLConfig, fill
from fillxl._lossy import find_lost_features
from tests._spills import (
    assert_well_formed,
    build_spill_in_growth_path,
    build_spill_template,
    read_package,
)
from tests.parity.builders import chart_sheet_first_template, sheets_template

# A real Excel-authored, macro-enabled template from the repository's assets:
# table ``TableTest``, a spilled array on ``SpilledRange`` and a VBA project.
XLSM_TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "test_template_1_slsm.xlsm"

ROWS = [
    {"Region": "East", "Amount": 30},
    {"Region": "West", "Amount": 40},
    {"Region": "North", "Amount": 55},
    {"Region": "South", "Amount": 60},
]


@pytest.fixture
def spill_template(tmp_path: Path) -> Path:
    """A totals-row table plus two spilled dynamic arrays."""
    return build_spill_template(tmp_path / "spill.xlsx")


def _sheet_xml(path: Path) -> str:
    return read_package(path)["xl/worksheets/sheet1.xml"].decode()


# ═══════════════════════════════════════════════════════════════════════════════
# Spilled dynamic arrays
# ═══════════════════════════════════════════════════════════════════════════════


def test_a_fill_keeps_the_spilled_arrays(spill_template: Path, tmp_path: Path) -> None:
    """The engine never parses what it does not change, so a spill cannot be lost."""
    output = fill(
        spill_template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Tbl": ROWS},
    )

    assert_well_formed(output)
    before = read_package(spill_template)
    after = read_package(output)
    assert after["xl/metadata.xml"] == before["xl/metadata.xml"]

    sheet = _sheet_xml(output)
    assert '<f t="array" ref="F8:F12">_xlfn.SEQUENCE(5)</f>' in sheet
    assert 'ref="J3:J4"' in sheet
    assert 'cm="1"' in sheet

    workbook = openpyxl.load_workbook(output)
    try:
        worksheet = workbook["Data"]
        master = worksheet.cell(row=8, column=6).value
        assert getattr(master, "text", master) == "=_xlfn.SEQUENCE(5)"
        # The spill itself is untouched, but the results the template cached
        # are gone: after the fill they describe a different table, and a
        # reader that cannot recalculate must see blank rather than the
        # template's numbers.  Excel refills F8:F12 and J3:J4 on open.
        assert [worksheet.cell(row=row, column=6).value for row in range(9, 13)] == [
            None,
            None,
            None,
            None,
        ]
        assert worksheet["H8"].value == "=SUM(_xlfn.ANCHORARRAY(F8))"
    finally:
        workbook.close()


def test_the_fill_resizes_the_table_and_keeps_its_totals_metadata(
    spill_template: Path,
    tmp_path: Path,
) -> None:
    output = fill(spill_template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})

    workbook = openpyxl.load_workbook(output)
    try:
        worksheet = workbook["Data"]
        table = worksheet.tables["Tbl"]
        assert table.ref == "B2:D7"
        assert table.totalsRowCount == 1
        assert [column.name for column in table.tableColumns] == ["Region", "Amount", "Double"]
        assert table.tableColumns[0].totalsRowLabel == "Total"
        assert table.tableColumns[1].totalsRowFunction == "sum"
        assert table.tableColumns[2].calculatedColumnFormula is not None
        # The totals row moved down with the table, its formulas intact.
        assert worksheet["B7"].value == "Total"
        assert worksheet["C7"].value == "=SUBTOTAL(109,Tbl[Amount])"
        # Data landed in rows 3-6 and the calculated column copied down.
        assert [worksheet.cell(row=row, column=2).value for row in range(3, 7)] == [
            "East",
            "West",
            "North",
            "South",
        ]
        assert worksheet["D6"].value == "=[@Amount]*2"
    finally:
        workbook.close()


def test_only_the_parts_that_have_to_change_are_rewritten(
    spill_template: Path,
    tmp_path: Path,
) -> None:
    """The whole point of a surgical editor: untouched bytes stay untouched."""
    output = fill(spill_template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})

    before = read_package(spill_template)
    after = read_package(output)
    assert sorted(before) == sorted(after)
    changed = {name for name in before if name in after and before[name] != after[name]}
    assert changed == {"xl/worksheets/sheet1.xml", "xl/tables/table1.xml"}
    # Everything a reader of the template would notice is still there.
    assert find_lost_features(spill_template, output) == []


def test_a_stale_calculation_chain_is_dropped(spill_template: Path, tmp_path: Path) -> None:
    """A chain that disagrees with the workbook makes Excel refuse to open it."""
    parts = read_package(spill_template)
    parts["xl/calcChain.xml"] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<c r="D3" i="1"/></calcChain>'
    )
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/xl/calcChain.xml" ContentType="application/'
        b'vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
    )
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    parts["xl/_rels/workbook.xml.rels"] = rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdChain" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>'
        "</Relationships>",
    ).encode()
    with zipfile.ZipFile(spill_template, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)

    output = fill(spill_template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})

    after = read_package(output)
    assert "xl/calcChain.xml" not in after
    assert b"calcChain" not in after["[Content_Types].xml"]
    assert b"calcChain" not in after["xl/_rels/workbook.xml.rels"]


def test_growing_into_a_spill_is_refused(tmp_path: Path) -> None:
    """A blank spill result still reserves its cells; writing there breaks Excel."""
    template = build_spill_in_growth_path(tmp_path / "spill_growth.xlsx")

    with pytest.raises(ValueError, match=r"cannot expand.*array formula B6:B9"):
        fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})


# ═══════════════════════════════════════════════════════════════════════════════
# Values, formats and calculation
# ═══════════════════════════════════════════════════════════════════════════════


def _named_cell_template(path: Path) -> Path:
    """One named cell, one table header, one sheet-scoped name."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"], worksheet["B1"] = "amount", "note"
    worksheet["A2"], worksheet["B2"] = 1, "sample"
    worksheet.add_table(
        Table(
            displayName="T",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="amount"), TableColumn(id=2, name="note")],
        ),
    )
    workbook.defined_names.add(DefinedName("report_date", attr_text="Data!$D$1"))
    worksheet.defined_names["scoped_amount"] = DefinedName(
        "scoped_amount",
        attr_text="Data!$D$2",
    )
    workbook.save(path)
    workbook.close()
    return path


def test_named_cells_are_written_with_dates_formatted(tmp_path: Path) -> None:
    import datetime as dt

    template = _named_cell_template(tmp_path / "names.xlsx")
    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        values={"REPORT_DATE": dt.date(2025, 6, 15), "scoped_amount": 12},
    )

    workbook = openpyxl.load_workbook(output)
    try:
        worksheet = workbook["Data"]
        assert worksheet["D1"].value == dt.datetime(2025, 6, 15)
        assert worksheet["D1"].number_format == "yyyy-mm-dd"
        assert worksheet["D2"].value == 12
        # Numeric cells keep the template's General format.
        assert worksheet["D2"].number_format == "General"
    finally:
        workbook.close()


def test_a_table_header_cannot_be_written_through_a_name(tmp_path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"], worksheet["B1"] = "amount", "note"
    worksheet["A2"], worksheet["B2"] = 1, "sample"
    worksheet.add_table(
        Table(
            displayName="T",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="amount"), TableColumn(id=2, name="note")],
        ),
    )
    pointer = tmp_path / "header.xlsx"
    workbook.defined_names.add(DefinedName("the_header", attr_text="Data!$A$1"))
    workbook.save(pointer)
    workbook.close()

    with pytest.raises(ValueError, match="is a header cell of table 'T'"):
        fill(pointer, tmp_path / "out.xlsx", engine="fillxl", values={"the_header": "renamed"})


def test_formulas_are_written_with_full_calc_on_load(tmp_path: Path) -> None:
    template = _named_cell_template(tmp_path / "formulas.xlsx")
    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        values={"report_date": "=1+1"},
        config=FillXLConfig(escape_formulas=False),
    )

    sheet = _sheet_xml(output)
    assert "<f>1+1</f>" in sheet
    workbook_xml = read_package(output)["xl/workbook.xml"].decode()
    assert 'fullCalcOnLoad="1"' in workbook_xml


def test_formula_like_text_is_escaped_by_default(tmp_path: Path) -> None:
    template = _named_cell_template(tmp_path / "escaped.xlsx")
    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"report_date": "=1+1"})

    sheet = _sheet_xml(output)
    assert "<f>1+1</f>" not in sheet
    assert "=1+1" in sheet


# ═══════════════════════════════════════════════════════════════════════════════
# Limits, reported clearly
# ═══════════════════════════════════════════════════════════════════════════════


def test_a_binary_template_is_refused(tmp_path: Path) -> None:
    binary = tmp_path / "template.xlsb"
    binary.write_bytes(b"not a zip")

    with pytest.raises(ValueError, match=r"fillxl engine cannot read a '.xlsb' template"):
        fill(binary, tmp_path / "out.xlsb", engine="fillxl")


def test_format_conversion_is_refused(tmp_path: Path) -> None:
    template = _named_cell_template(tmp_path / "template.xlsx")

    with pytest.raises(ValueError, match=r"cannot write a '.xlsb' file"):
        fill(template, tmp_path / "out.xlsb", engine="fillxl")


def test_a_password_protected_workbook_is_refused(tmp_path: Path) -> None:
    """An encrypted workbook is a CFB container, not a zip package."""
    encrypted = tmp_path / "encrypted.xlsx"
    encrypted.write_bytes(bytes.fromhex("d0cf11e0a1b11ae1") + b"\x00" * 512)

    with pytest.raises(ValueError, match="not an OOXML"):
        fill(encrypted, tmp_path / "out.xlsx", engine="fillxl")


# ???????????????????????????????????????????????????????????????????????????????
# Sheets
# ???????????????????????????????????????????????????????????????????????????????


def test_an_underscore_sheet_is_hidden_and_the_visible_one_becomes_active(
    tmp_path: Path,
) -> None:
    template = sheets_template(tmp_path / "template.xlsx", "_data", "Report")

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl")

    assert_well_formed(output)
    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook["_data"].sheet_state == "hidden"
        assert workbook.active is not None
        assert workbook.active.title == "Report"
        selected = [name for name in workbook.sheetnames if workbook[name].sheet_view.tabSelected]
        assert selected == ["Report"]
    finally:
        workbook.close()


def test_deleting_a_sheet_removes_the_parts_only_it_used(tmp_path: Path) -> None:
    """A deleted chart sheet takes its chart with it instead of leaving it orphaned."""
    template = chart_sheet_first_template(tmp_path / "template.xlsx")
    before = read_package(template)

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        config=FillXLConfig(hide_underscore_sheets=False, delete_sheets=["Chart"]),
    )

    assert_well_formed(output)
    after = read_package(output)
    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.sheetnames == ["Report"]
    finally:
        workbook.close()
    assert len(after) < len(before)
    assert not [name for name in after if name.startswith("xl/charts/chart")]


def test_deleting_a_referenced_sheet_warns_about_what_still_points_at_it(
    tmp_path: Path,
) -> None:
    """A formula, a chart and a name that stay behind would show #REF!."""
    from fillxl import DanglingReferenceWarning

    workbook = openpyxl.Workbook()
    report = workbook.active
    assert report is not None
    report.title = "Report"
    report["A1"] = "=SUM(RawData!B2:B2)"
    raw = workbook.create_sheet("RawData")
    raw.append(["k", "v"])
    raw.append(["a", 1])
    raw.append(["b", 2])
    workbook.defined_names["raw_total"] = DefinedName("raw_total", attr_text="RawData!$B$2")
    template = tmp_path / "template.xlsx"
    workbook.save(template)
    workbook.close()

    with pytest.warns(DanglingReferenceWarning, match="Report!A1"):
        output = fill(
            template,
            tmp_path / "out.xlsx",
            engine="fillxl",
            config=FillXLConfig(hide_underscore_sheets=False, delete_sheets=["RawData"]),
        )

    assert "xl/worksheets/sheet2.xml" not in read_package(output)


def test_on_lossy_raise_accepts_a_fillxl_output(tmp_path: Path) -> None:
    """Nothing is lost, so the strictest policy passes."""
    template = _named_cell_template(tmp_path / "strict.xlsx")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        values={"report_date": "x"},
        config=FillXLConfig(on_lossy="raise"),
    )

    assert output.exists()


def test_a_chart_follows_a_resized_table(spill_template: Path, tmp_path: Path) -> None:
    """A bar chart over the table body grows with it, cache included."""
    from openpyxl.chart import BarChart, Reference

    workbook = openpyxl.load_workbook(spill_template)
    worksheet = workbook["Data"]
    chart = BarChart()
    chart.add_data(
        Reference(worksheet, min_col=3, min_row=2, max_row=4),
        titles_from_data=True,
    )
    worksheet.add_chart(chart, "B12")
    workbook.save(spill_template)
    workbook.close()

    output = fill(spill_template, tmp_path / "out.xlsx", engine="fillxl", tables={"Tbl": ROWS})

    assert_well_formed(output)
    workbook = openpyxl.load_workbook(output)
    try:
        reference = next(iter(workbook["Data"]._charts)).series[0].val.numRef
        assert reference.f.replace("'", "") == "Data!$C$3:$C$6"
        # openpyxl writes charts without point caches; a cache that is there
        # (Excel-authored templates) is brought in step ? see test_ooxml.py.
    finally:
        workbook.close()


# ???????????????????????????????????????????????????????????????????????????????
# Styling
# ???????????????????????????????????????????????????????????????????????????????


def _styled_table_template(path: Path) -> Path:
    """A one-row table whose first data row is bold on a grey fill."""
    from openpyxl.styles import Font, PatternFill

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"], worksheet["B1"] = "Name", "Value"
    worksheet["A2"], worksheet["B2"] = "a", 1
    for cell in (worksheet["A2"], worksheet["B2"]):
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", start_color="FFD9D9D9")
    worksheet.add_table(
        Table(
            displayName="Styled",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
        ),
    )
    workbook.save(path)
    workbook.close()
    return path


def test_new_rows_copy_the_first_data_row_style(tmp_path: Path) -> None:
    template = _styled_table_template(tmp_path / "styled.xlsx")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Styled": [{"Name": name, "Value": index} for index, name in enumerate("abc", 1)]},
    )

    workbook = openpyxl.load_workbook(output)
    try:
        worksheet = workbook["Data"]
        for row in (2, 3, 4):
            assert worksheet.cell(row=row, column=1).font.b is True
            assert worksheet.cell(row=row, column=1).fill.fgColor.rgb == "FFD9D9D9"
    finally:
        workbook.close()


def test_copy_table_styles_false_leaves_new_rows_unformatted(tmp_path: Path) -> None:
    """The engine must not style anything the caller asked it not to touch."""
    template = _styled_table_template(tmp_path / "styled.xlsx")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Styled": [{"Name": name, "Value": index} for index, name in enumerate("abc", 1)]},
        config=FillXLConfig(copy_table_styles=False),
    )

    workbook = openpyxl.load_workbook(output)
    try:
        worksheet = workbook["Data"]
        assert worksheet["A2"].font.b is True  # the template row keeps its style
        assert worksheet["A3"].font.b is not True
        assert worksheet["A3"].fill.fill_type is None
    finally:
        workbook.close()


# ???????????????????????????????????????????????????????????????????????????????
# A real Excel-authored template
# ???????????????????????????????????????????????????????????????????????????????


def _xlsm_rows() -> list[dict[str, object]]:
    return [
        {"Col1": index, "Col2": index * 2, "ColNotFilled": None, "TestFormattedPct": 0.1}
        for index in range(1, 7)
    ]


def test_a_real_macro_template_keeps_its_spill_metadata_and_vba(tmp_path: Path) -> None:
    """``test_template_1_slsm.xlsm`` is authored by Excel, spills and all.

    The spilled array sits on another sheet, so it is not even touched ? but
    ``xl/metadata.xml`` is what makes it a spill after the fill, and it has to
    come through byte for byte.  The spill's *results* go, like every other
    cached result: they were computed from the template's data.
    """
    output = fill(
        XLSM_TEMPLATE,
        tmp_path / "out.xlsm",
        engine="fillxl",
        tables={"TableTest": _xlsm_rows()},
    )

    assert_well_formed(output)
    before = read_package(XLSM_TEMPLATE)
    after = read_package(output)
    assert after["xl/metadata.xml"] == before["xl/metadata.xml"]
    assert after["xl/vbaProject.bin"] == before["xl/vbaProject.bin"]
    spill = after["xl/worksheets/sheet3.xml"].decode()
    assert '<f t="array" ref="B2:B5">TableTest[Col1]</f>' in spill
    assert 'cm="1"' in spill
    assert "<v>" not in spill
    assert "xl/calcChain.xml" not in after

    changed = {name for name in before if name in after and before[name] != after[name]}
    assert changed == {
        "[Content_Types].xml",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/sheet3.xml",
        "xl/tables/table1.xml",
    }

    workbook = openpyxl.load_workbook(output, keep_vba=True)
    try:
        worksheet = workbook["TestTable"]
        table = worksheet.tables["TableTest"]
        assert table.ref == "A1:E7"
        assert [worksheet.cell(row=row, column=1).value for row in range(2, 8)] == [
            1,
            2,
            3,
            4,
            5,
            6,
        ]
        assert worksheet["D7"].value == (
            "=TableTest[[#This Row],[Col1]]+TableTest[[#This Row],[Col2]]"
        )
    finally:
        workbook.close()
