"""Modern charts (``chartEx``) and sparklines follow a resized table.

A waterfall chart stores its sources in ``<cx:f>`` elements of its own part,
and a sparkline in ``<xm:f>`` of the sheet's ``extLst``.  Neither is a cell
formula, so neither was touched by a fill: they kept showing the template's
rows with no warning at all.

The fixtures were made in Excel.  ``waterfall.xlsx`` is a waterfall chart over
the table's ``Amount`` column with the ranges written directly into the chart
part, which is the layout Excel wrote before it started keeping them in hidden
``_xlchart`` names; ``sparkline.xlsx`` holds one sparkline per cell.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from fillxl import fill
from tests._spills import assert_well_formed, read_package
from tests.ooxml._patch import part_text, patch_part

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
WATERFALL = FIXTURES / "waterfall.xlsx"
SPARKLINE = FIXTURES / "sparkline.xlsx"

CHART_EX = "xl/charts/chartEx1.xml"
SPARKLINE_SHEET = "xl/worksheets/sheet2.xml"

GROWTH = [{"Step": f"s{index}", "Amount": index * 10} for index in range(5)]
SHRINK = [{"Step": "only", "Amount": 5}]


def _chart_xml(path: Path) -> str:
    return read_package(path)[CHART_EX].decode()


def _dimension(xml: str, kind: str) -> str:
    """The ``<cx:strDim>`` or ``<cx:numDim>`` block of the chart part."""
    match = re.search(rf"<cx:{kind}\b.*?</cx:{kind}>", xml, re.S)
    assert match is not None, f"no <cx:{kind}> in {xml}"
    return match.group(0)


def _formula(block: str) -> str | None:
    """The text of the ``<cx:f>`` element of a dimension block."""
    match = re.search(r"<cx:f>(.*?)</cx:f>", block, re.S)
    return None if match is None else match.group(1)


def _pt_counts(block: str) -> list[int]:
    """The ``ptCount`` of every ``<cx:lvl>`` of a dimension block."""
    return [int(value) for value in re.findall(r'<cx:lvl[^>]*\bptCount="(\d+)"', block)]


def test_a_waterfall_chart_follows_its_table_when_the_table_grows(tmp_path: Path) -> None:
    """The categories and values gain the rows the fill wrote."""
    output = fill(WATERFALL, tmp_path / "out.xlsx", engine="fillxl", tables={"Sales": GROWTH})

    assert_well_formed(output)
    chart = _chart_xml(output)
    categories = _dimension(chart, "strDim")
    values = _dimension(chart, "numDim")
    assert _formula(categories) == "Data!$A$2:$A$6"
    assert _formula(values) == "Data!$B$2:$B$6"
    assert _pt_counts(categories) == [5]
    assert _pt_counts(values) == [5]


def test_a_waterfall_chart_follows_its_table_when_the_table_shrinks(tmp_path: Path) -> None:
    """The same rule in reverse: the chart covers the remaining body rows."""
    output = fill(WATERFALL, tmp_path / "out.xlsx", engine="fillxl", tables={"Sales": SHRINK})

    assert_well_formed(output)
    chart = _chart_xml(output)
    assert _formula(_dimension(chart, "strDim")) == "Data!$A$2:$A$2"
    assert _formula(_dimension(chart, "numDim")) == "Data!$B$2:$B$2"
    assert _pt_counts(_dimension(chart, "numDim")) == [1]


def test_a_modern_chart_series_name_is_left_alone(tmp_path: Path) -> None:
    """The name reads the header cell next to the body; it is not a data range."""
    output = fill(WATERFALL, tmp_path / "out.xlsx", engine="fillxl", tables={"Sales": GROWTH})

    chart = _chart_xml(output)
    match = re.search(r"<cx:tx>.*?</cx:tx>", chart, re.S)
    assert match is not None
    assert "<cx:f>Data!$B$1</cx:f>" in match.group(0)


def test_a_sparkline_over_a_table_column_follows_the_table(tmp_path: Path) -> None:
    """The sparkline's source is markup in the sheet's ``extLst``."""
    output = fill(SPARKLINE, tmp_path / "out.xlsx", engine="fillxl", tables={"Sales": GROWTH})

    assert_well_formed(output)
    sheet = read_package(output)[SPARKLINE_SHEET].decode()
    assert "<xm:f>Data!B2:B6</xm:f>" in sheet


@pytest.mark.excel
def test_excel_resizes_a_modern_chart_itself(tmp_path: Path) -> None:
    """Excel keeps a chartEx source in hidden ``_xlchart`` names and updates them."""
    output = fill(WATERFALL, tmp_path / "out.xlsx", engine="win32com", tables={"Sales": GROWTH})

    workbook = read_package(output)["xl/workbook.xml"].decode()
    sources = re.findall(r'<definedName name="_xlchart[^"]*"[^>]*>(.*?)</definedName>', workbook)
    assert "Data!$A$2:$A$6" in sources
    assert "Data!$B$2:$B$6" in sources


def test_a_sparkline_that_reads_one_row_per_cell_is_left_alone(tmp_path: Path) -> None:
    """A member covering one row of its own is not the table's body."""
    template = tmp_path / "template.xlsx"
    template.write_bytes(SPARKLINE.read_bytes())
    patch_part(
        template,
        SPARKLINE_SHEET,
        lambda xml: xml.replace(
            "<xm:f>Data!B2:B4</xm:f><xm:sqref>A2</xm:sqref>",
            "<xm:f>Data!B2:B2</xm:f><xm:sqref>A2</xm:sqref>",
        ),
    )
    assert "Data!B2:B2" in part_text(template, SPARKLINE_SHEET)

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Sales": GROWTH})

    assert "<xm:f>Data!B2:B2</xm:f>" in read_package(output)[SPARKLINE_SHEET].decode()
