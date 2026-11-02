"""Unit tests for the OOXML primitives behind ``engine="fillxl"``.

These test the surgical editors on their own: text-level XML helpers, the
package container, the worksheet/table/chart/pivot editors.  The end-to-end
behaviour lives in ``tests/integration/test_fillxl_engine.py``.
"""

from __future__ import annotations

import xml.etree.ElementTree as ElementTree
import zipfile
from typing import TYPE_CHECKING

import pytest

from fillxl._cells import name_destinations
from fillxl._charts import TableChangeSpecs
from fillxl._ooxml import xmltext
from fillxl._ooxml.charts import update_chart_references
from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.pivots import set_refresh_on_load
from fillxl._ooxml.styles import StylesPart
from fillxl._ooxml.tables import TablePart
from fillxl._ooxml.workbook import WorkbookPart
from fillxl._ooxml.worksheet import SheetPart, cell_has_value, column_index, split_reference

if TYPE_CHECKING:
    from pathlib import Path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

WORKBOOK_XML = (
    f'<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
    '<bookViews><workbookView activeTab="0"/></bookViews>'
    '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/>'
    '<sheet name="_Hidden" sheetId="2" r:id="rId2"/></sheets>'
    '<definedNames><definedName name="global_name">Data!$A$1</definedName>'
    '<definedName name="scoped" localSheetId="0">Data!$B$2</definedName></definedNames>'
    '<calcPr calcId="191029"/>'
    "</workbook>"
)

WORKBOOK_RELS = (
    f'<Relationships xmlns="{PKG_NS}">'
    f'<Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>'
    f'<Relationship Id="rId2" Type="{REL_NS}/worksheet" Target="worksheets/sheet2.xml"/>'
    f'<Relationship Id="rId3" Type="{REL_NS}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId4" Type="{REL_NS}/calcChain" Target="calcChain.xml"/>'
    "</Relationships>"
)

SHEET_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}" xmlns:mc="http://schemas.'
    'openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" '
    'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
    '<dimension ref="A1:D6"/>'
    '<sheetViews><sheetView workbookViewId="0" tabSelected="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15"/>'
    "<sheetData>"
    '<row r="1" spans="1:3"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>1</v></c></row>'
    '<row r="2" spans="1:3"><c r="A2" s="2"><v>2</v></c></row>'
    '<row r="8"><c r="F8" cm="1"><f t="array" ref="F8:F10">_xlfn.SEQUENCE(3)</f>'
    "<v>1</v></c></row>"
    '<row r="9"><c r="F9"><v>2</v></c></row>'
    '<row r="10"><c r="F10"><v>3</v></c></row>'
    "</sheetData>"
    '<mergeCells count="1"><mergeCell ref="C4:D4"/></mergeCells>'
    '<conditionalFormatting sqref="B2:B2">'
    '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula>'
    "</cfRule></conditionalFormatting>"
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    "</worksheet>"
)

STYLES_XML = (
    f'<styleSheet xmlns="{MAIN_NS}">'
    '<numFmts count="0"/>'
    '<fonts count="1"><font><sz val="11"/></font></fonts>'
    '<fills count="1"><fill><patternFill/></fill></fills>'
    '<borders count="1"><border/></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyBorder="1"/></cellXfs>'
    "</styleSheet>"
)

CHART_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    "<c:chart><c:plotArea><c:barChart><c:ser>"
    '<c:idx val="0"/><c:val><c:numRef><c:f>Data!$B$3:$B$4</c:f>'
    '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/>'
    '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt>'
    "</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart>"
    "</c:chartSpace>"
)

PIVOT_CACHE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<pivotCacheDefinition xmlns="{MAIN_NS}" r:id="rId1" recordCount="2">'
    '<cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Data"/>'
    "</cacheSource></pivotCacheDefinition>"
)


def _content_types(parts: list[str]) -> str:
    overrides = "".join(
        f'<Override PartName="/{name}" ContentType="application/xml"/>' for name in parts
    )
    return f'<Types xmlns="{CT_NS}">{overrides}</Types>'


def build_package(path: Path, extra: dict[str, str] | None = None) -> Path:
    """Write a small but structurally valid workbook package."""
    parts: dict[str, bytes] = {
        "[Content_Types].xml": _content_types(
            [
                "xl/workbook.xml",
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/styles.xml",
            ],
        ).encode(),
        "_rels/.rels": (
            f'<Relationships xmlns="{PKG_NS}"><Relationship Id="rId1" Type="{REL_NS}/'
            'officeDocument" Target="xl/workbook.xml"/></Relationships>'
        ).encode(),
        "xl/workbook.xml": WORKBOOK_XML.encode(),
        "xl/_rels/workbook.xml.rels": WORKBOOK_RELS.encode(),
        "xl/worksheets/sheet1.xml": SHEET_XML.encode(),
        "xl/worksheets/sheet2.xml": (
            f'<worksheet xmlns="{MAIN_NS}"><sheetData/></worksheet>'.encode()
        ),
        "xl/styles.xml": STYLES_XML.encode(),
    }
    for name, text in (extra or {}).items():
        parts[name] = text.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


@pytest.fixture
def package_path(tmp_path: Path) -> Path:
    return build_package(tmp_path / "book.xlsx")


# ═══════════════════════════════════════════════════════════════════════════════
# Text-level helpers
# ═══════════════════════════════════════════════════════════════════════════════


def test_entities_are_escaped_and_unescaped() -> None:
    assert xmltext.unescape("a &amp; b &lt;c&gt; &#65;&#x42;") == "a & b <c> AB"
    assert xmltext.escape_text("a & b <c>") == "a &amp; b &lt;c&gt;"
    assert xmltext.escape_attribute('say "hi"') == "say &quot;hi&quot;"


def test_attributes_are_added_replaced_and_removed() -> None:
    assert xmltext.set_attribute('<c r="A1"/>', "s", "2") == '<c r="A1" s="2"/>'
    assert xmltext.set_attribute('<c r="A1" s="2"/>', "s", "3") == '<c r="A1" s="3"/>'
    assert xmltext.remove_attribute('<c r="A1" s="2"/>', "s") == '<c r="A1"/>'
    assert xmltext.attribute('<c r="A1" s="2"/>', "s") == "2"
    assert xmltext.attribute('<c r="A1"/>', "s") is None


def test_element_text_round_trips() -> None:
    assert xmltext.element_text("<v>42</v>", "v") == "42"
    assert xmltext.element_text("<v/>", "v") is None
    assert xmltext.replace_element_text("<c><f>A1</f></c>", "f", "B2*2") == "<c><f>B2*2</f></c>"
    assert xmltext.replace_element_text("<c><f/></c>", "f", "B2") == "<c><f/></c>"


def test_split_reference_and_column_index() -> None:
    assert split_reference("$AB$12") == (12, 28)
    assert column_index("AA") == 27
    with pytest.raises(ValueError, match="Not a cell reference"):
        split_reference("nonsense")


# ═══════════════════════════════════════════════════════════════════════════════
# Package
# ═══════════════════════════════════════════════════════════════════════════════


def test_untouched_parts_are_copied_byte_for_byte(package_path: Path, tmp_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = package.text("xl/worksheets/sheet1.xml")
    package.set_text("xl/worksheets/sheet1.xml", sheet.replace("A2", "A3"))

    output = tmp_path / "out.xlsx"
    package.save(output)

    with zipfile.ZipFile(package_path) as before, zipfile.ZipFile(output) as after:
        assert sorted(before.namelist()) == sorted(after.namelist())
        for name in before.namelist():
            if name == "xl/worksheets/sheet1.xml":
                continue
            assert before.read(name) == after.read(name), name


def test_a_non_zip_package_is_refused(tmp_path: Path) -> None:
    fake = tmp_path / "encrypted.xlsx"
    fake.write_bytes(b"not a zip at all")

    with pytest.raises(ValueError, match="not an OOXML"):
        OoxmlPackage(fake)


def test_relationships_resolve_relative_and_absolute_targets(package_path: Path) -> None:
    package = OoxmlPackage(package_path)

    assert package.target_part("xl/workbook.xml", "rId1") == "xl/worksheets/sheet1.xml"
    assert package.related_parts("xl/workbook.xml", "/worksheet") == [
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
    ]
    assert package.target_part("", "rId1") == "xl/workbook.xml"


def test_overrides_and_parts_can_be_dropped(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    package.drop("xl/styles.xml")

    assert not package.has("xl/styles.xml")
    assert b"styles.xml" not in package.raw("[Content_Types].xml")


# ═══════════════════════════════════════════════════════════════════════════════
# Worksheet editor
# ═══════════════════════════════════════════════════════════════════════════════


def test_an_untouched_sheet_round_trips_exactly(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")

    assert sheet.to_xml() == SHEET_XML


def test_writing_a_cell_keeps_every_other_row_intact(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")
    sheet.write_cell(2, 3, '<c r="C2"><v>5</v></c>')
    sheet.save()

    text = package.text("xl/worksheets/sheet1.xml")
    # Row 1 is untouched, so it is written back exactly as it was read.
    assert (
        '<row r="1" spans="1:3"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>1</v></c></row>'
    ) in text
    # The edited row is re-rendered: the advisory `spans` hint goes away.
    assert '<row r="2"><c r="A2" s="2"><v>2</v></c><c r="C2"><v>5</v></c></row>' in text
    assert 'mc:Ignorable="x14ac"' in text
    assert '<mergeCell ref="C4:D4"/>' in text


def test_a_new_row_is_inserted_in_order(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")
    sheet.write_cell(5, 1, '<c r="A5"><v>9</v></c>')
    sheet.save()

    rows = [
        line.split('"')[1] for line in package.text("xl/worksheets/sheet1.xml").split("<row r=")[1:]
    ]
    assert rows == ["1", "2", "5", "8", "9", "10"]


def test_a_style_can_be_added_to_a_cell_that_had_none(package_path: Path) -> None:
    """The `s` attribute lands on the start tag even when there was none."""
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")
    sheet.write_cell(2, 1, '<c r="A2"><v>2</v></c>')
    sheet.set_style(2, 1, "3")
    sheet.save()

    root = ElementTree.fromstring(package.text("xl/worksheets/sheet1.xml"))
    row = root.find(f"{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row[@r='2']")
    assert row is not None
    cell = row.find(f"{{{MAIN_NS}}}c[@r='A2']")
    assert cell is not None
    assert cell.get("s") == "3"
    value = cell.find(f"{{{MAIN_NS}}}v")
    assert value is not None
    assert value.text == "2"


def test_clearing_a_cell_keeps_its_style(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")
    sheet.write_cell(2, 1, '<c r="A2" s="2"/>')

    assert sheet.cell(2, 1) == '<c r="A2" s="2"/>'
    assert not sheet.has_value(2, 1)


def test_value_and_formula_detection() -> None:
    assert not cell_has_value('<c r="A1" s="1"/>')
    assert cell_has_value('<c r="A1"><v>0</v></c>')
    assert not cell_has_value('<c r="A1" t="str"><v/></c>')
    assert cell_has_value('<c r="A1"><f>1+1</f></c>')
    assert cell_has_value('<c r="A1" t="inlineStr"><is><t>x</t></is></c>')


def test_array_ranges_are_reported(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")

    assert [str(cell_range) for cell_range in sheet.array_ranges()] == ["F8:F10"]
    assert sheet.array_range_at(9, 6) == "F8:F10"
    assert sheet.array_range_at(9, 5) is None


def test_merged_ranges_are_read(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")

    assert sheet.merged_ranges() == ["C4:D4"]


def test_conditional_formatting_that_tracks_the_table_is_extended(
    package_path: Path,
) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")

    changes = sheet.extend_conditional_formatting(
        table_start_column=1,
        table_end_column=2,
        table_top_row=1,
        first_data_row=2,
        old_body_last_row=2,
        table_last_row=5,
    )

    assert changes == [("B2:B2", "B2:B5")]
    # The attribute has to land on the start tag: a conditionalFormatting
    # element has children, so appending to its closing tag corrupts the part.
    root = ElementTree.fromstring(sheet.to_xml())
    element = root.find(f"{{{MAIN_NS}}}conditionalFormatting")
    assert element is not None
    assert element.get("sqref") == "B2:B5"
    assert element.find(f"{{{MAIN_NS}}}cfRule") is not None


def test_the_dimension_grows_to_cover_new_rows(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    sheet = SheetPart(package, "xl/worksheets/sheet1.xml")

    sheet.expand_dimension(12)
    assert '<dimension ref="A1:D12"/>' in sheet.to_xml()
    sheet.expand_dimension(3)
    assert '<dimension ref="A1:D12"/>' in sheet.to_xml()


# ═══════════════════════════════════════════════════════════════════════════════
# Styles
# ═══════════════════════════════════════════════════════════════════════════════


def test_a_date_format_is_added_once(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    styles = StylesPart(package)

    first = styles.date_style("1", "yyyy-mm-dd")
    second = styles.date_style("1", "yyyy-mm-dd")
    other = styles.date_style("0", "yyyy-mm-dd")

    assert first == second == "2"
    assert other == "3"
    styles.save()

    text = package.text("xl/styles.xml")
    assert '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>' in text
    assert '<cellXfs count="4">' in text


def test_a_complete_format_is_not_duplicated(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    styles = StylesPart(package)

    assert styles.number_format("0") == "General"
    assert styles.number_format("99") is None
    assert styles.date_style("0", "General") == "0"


def test_styles_are_untouched_when_no_date_arrives(package_path: Path, tmp_path: Path) -> None:
    package = OoxmlPackage(package_path)
    styles = StylesPart(package)
    styles.save()

    output = tmp_path / "out.xlsx"
    package.save(output)
    with zipfile.ZipFile(output) as archive:
        assert archive.read("xl/styles.xml") == STYLES_XML.encode()


# ═══════════════════════════════════════════════════════════════════════════════
# Tables
# ═══════════════════════════════════════════════════════════════════════════════


TABLE_XML = (
    f'<table xmlns="{MAIN_NS}" id="1" name="Tbl" displayName="Tbl" ref="B2:D4" '
    'totalsRowCount="1">'
    '<autoFilter ref="B2:D3"/>'
    '<tableColumns count="3">'
    '<tableColumn id="1" name="Region" totalsRowLabel="Total"/>'
    '<tableColumn id="2" name="Amount" totalsRowFunction="sum"/>'
    '<tableColumn id="3" name="Double"><calculatedColumnFormula>[@Amount]*2'
    "</calculatedColumnFormula></tableColumn>"
    "</tableColumns></table>"
)


def _table_package(tmp_path: Path) -> OoxmlPackage:
    path = build_package(tmp_path / "tables.xlsx", {"xl/tables/table1.xml": TABLE_XML})
    return OoxmlPackage(path)


def test_a_table_reports_its_geometry_and_columns(tmp_path: Path) -> None:
    package = _table_package(tmp_path)
    from fillxl._ooxml.workbook import SheetEntry

    sheet = SheetEntry("Data", "1", "rId1", "visible", "xl/worksheets/sheet1.xml", 0)
    table = TablePart(package, "xl/tables/table1.xml", sheet)

    geometry = table.geometry()
    assert (geometry.top_row, geometry.last_row, geometry.first_col, geometry.last_col) == (
        2,
        4,
        2,
        4,
    )
    assert geometry.header_rows == 1
    assert geometry.totals_rows == 1
    columns = table.columns()
    assert [column.name for column in columns] == ["Region", "Amount", "Double"]
    assert columns[0].totals_label == "Total"
    assert columns[1].totals_function == "sum"
    assert columns[2].calculated_formula == "[@Amount]*2"


def test_resizing_a_table_keeps_column_metadata(tmp_path: Path) -> None:
    package = _table_package(tmp_path)
    from fillxl._ooxml.workbook import SheetEntry

    sheet = SheetEntry("Data", "1", "rId1", "visible", "xl/worksheets/sheet1.xml", 0)
    table = TablePart(package, "xl/tables/table1.xml", sheet)
    table.set_refs("B2:D9", "B2:D8")

    text = package.text("xl/tables/table1.xml")
    assert 'ref="B2:D9"' in text
    assert 'totalsRowCount="1"' in text
    assert '<autoFilter ref="B2:D8"/>' in text
    assert "<calculatedColumnFormula>[@Amount]*2</calculatedColumnFormula>" in text


# ═══════════════════════════════════════════════════════════════════════════════
# Charts and pivot caches
# ═══════════════════════════════════════════════════════════════════════════════


def test_chart_references_and_point_counts_follow_the_table(tmp_path: Path) -> None:
    path = build_package(tmp_path / "chart.xlsx")
    package = OoxmlPackage(path)
    package.set_text("xl/charts/chart1.xml", CHART_XML)

    spec = TableChangeSpecs("Data", 2, 2, 2, 4, 6)
    changed = update_chart_references(package, spec)

    assert changed == 1
    text = package.text("xl/charts/chart1.xml")
    assert "<c:f>Data!$B$3:$B$6</c:f>" in text
    assert '<c:ptCount val="4"/>' in text


def test_a_chart_that_shows_something_else_is_untouched(tmp_path: Path) -> None:
    path = build_package(tmp_path / "chart.xlsx")
    package = OoxmlPackage(path)
    package.set_text("xl/charts/chart1.xml", CHART_XML)

    spec = TableChangeSpecs("Other", 2, 2, 2, 4, 6)
    assert update_chart_references(package, spec) == 0
    assert package.text("xl/charts/chart1.xml") == CHART_XML


def test_pivot_caches_are_marked_for_refresh(tmp_path: Path) -> None:
    path = build_package(tmp_path / "pivot.xlsx")
    package = OoxmlPackage(path)
    package.set_text("xl/pivotCache/pivotCacheDefinition1.xml", PIVOT_CACHE_XML)

    assert set_refresh_on_load(package) == 1
    assert 'refreshOnLoad="1"' in package.text("xl/pivotCache/pivotCacheDefinition1.xml")


# ═══════════════════════════════════════════════════════════════════════════════
# Workbook
# ═══════════════════════════════════════════════════════════════════════════════


def test_sheets_are_listed_with_their_parts(package_path: Path) -> None:
    workbook = WorkbookPart(OoxmlPackage(package_path))

    sheets = workbook.sheets()
    assert [sheet.name for sheet in sheets] == ["Data", "_Hidden"]
    assert sheets[0].part == "xl/worksheets/sheet1.xml"
    assert sheets[0].state == "visible"


def test_defined_names_are_found_ignoring_case(package_path: Path) -> None:
    workbook = WorkbookPart(OoxmlPackage(package_path))

    assert workbook.find_defined_name("global_name") is not None
    assert workbook.find_defined_name("GLOBAL_NAME") is not None
    scoped = workbook.find_defined_name("scoped")
    assert scoped is not None
    assert scoped.local_sheet_id == 0
    assert workbook.find_defined_name("missing") is None
    assert workbook.defined_names()[0].text == "Data!$A$1"


def test_name_destinations_match_the_openpyxl_shape() -> None:
    assert name_destinations("Data!$A$1") == [("Data", "$A$1")]
    assert name_destinations("'Bob''s Data'!$A$1") == [("Bob''s Data", "$A$1")]
    assert name_destinations("Data!$A$1,Other!$B$2:$C$3") == [
        ("Data", "$A$1"),
        ("Other", "$B$2:$C$3"),
    ]
    assert name_destinations("OFFSET(Data!$A$1,0,0)") == []


def test_calculation_settings_and_the_calc_chain(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    workbook = WorkbookPart(package)
    package.set_text(
        "xl/calcChain.xml",
        f'<calcChain xmlns="{MAIN_NS}"><c r="D3" i="1"/></calcChain>',
    )

    workbook.set_full_calc_on_load()
    workbook.remove_calc_chain()

    assert 'fullCalcOnLoad="1"' in package.text("xl/workbook.xml")
    assert not package.has("xl/calcChain.xml")
    assert b"calcChain" not in package.raw("xl/_rels/workbook.xml.rels")


def test_a_sheet_can_be_hidden_and_activated(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    workbook = WorkbookPart(package)

    workbook.set_sheet_state("_Hidden", "hidden")
    workbook.set_active_sheet("_Hidden")

    hidden = workbook.sheet("_hidden")
    assert hidden is not None
    assert hidden.state == "hidden"
    assert workbook.active_sheet_name() == "_Hidden"
    assert 'tabSelected="1"' in package.text("xl/worksheets/sheet2.xml")
    assert 'tabSelected="0"' in package.text("xl/worksheets/sheet1.xml")


def test_removing_a_sheet_removes_its_part_and_relationship(package_path: Path) -> None:
    package = OoxmlPackage(package_path)
    workbook = WorkbookPart(package)

    removed = workbook.remove_sheets(["_Hidden"])

    assert [sheet.name for sheet in removed] == ["_Hidden"]
    assert [sheet.name for sheet in workbook.sheets()] == ["Data"]
    assert not package.has("xl/worksheets/sheet2.xml")
    assert b"sheet2.xml" not in package.raw("xl/_rels/workbook.xml.rels")
