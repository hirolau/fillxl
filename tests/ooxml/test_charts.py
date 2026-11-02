"""Chart data references: the sources move with the table, the names do not.

A chart stores a series name in ``<c:tx>`` and its data in ``<c:val>``, both as
the same kind of reference block.  A series name that points at the first data
row looks exactly like a data range, so rewriting every reference in the part
turns a name into a range.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.data_source import NumDataSource, NumRef, StrRef
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.series import SeriesLabel

from fillxl import fill
from tests.ooxml._build import body_rows, table_template
from tests.ooxml._patch import part_text

if TYPE_CHECKING:
    from pathlib import Path

CHART_PART = "xl/charts/chart1.xml"


def test_a_series_name_pointing_at_the_only_data_row_stays_a_single_cell(
    tmp_path: Path,
) -> None:
    """The name is a label, not a data source: growing the table must not widen it."""
    template = _chart_template(tmp_path / "template.xlsx")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(4)},
    )

    chart = part_text(output, CHART_PART)
    assert _reference(chart, "tx") == "Data!$C$2"
    assert _reference(chart, "val") == "'Data'!$C$2:$C$5"


def test_custom_error_bars_follow_the_table_when_it_grows(tmp_path: Path) -> None:
    """Error-bar bounds are data sources, so they grow with the body."""
    template = _chart_template(tmp_path / "template.xlsx")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(4)},
    )

    chart = part_text(output, CHART_PART)
    assert _reference(chart, "plus") == "Data!$C$2:$C$5"
    assert _reference(chart, "minus") == "Data!$C$2:$C$5"


def _reference(chart_xml: str, element: str) -> str:
    """The ``<f>`` text inside the chart's *element* block."""
    block = re.search(rf"<{element}>(.*?)</{element}>", chart_xml, re.S)
    assert block is not None, f"no <{element}> in the chart"
    formula = re.search(r"<f>(.*?)</f>", block.group(1), re.S)
    assert formula is not None, f"no <f> in <{element}>"
    return formula.group(1)


def _chart_template(path: Path) -> Path:
    """A one-row table whose chart series is named after its only data row."""
    table_template(path, rows=1)
    workbook = openpyxl.load_workbook(path)
    try:
        worksheet = workbook["Data"]
        chart = BarChart()
        chart.add_data(
            Reference(worksheet, min_col=3, min_row=1, max_row=2),
            titles_from_data=True,
        )
        series = chart.series[0]
        series.tx = SeriesLabel(strRef=StrRef(f="Data!$C$2"))
        series.errBars = ErrorBars(
            errDir="y",
            errBarType="both",
            errValType="cust",
            plus=NumDataSource(numRef=NumRef(f="Data!$C$2:$C$2")),
            minus=NumDataSource(numRef=NumRef(f="Data!$C$2:$C$2")),
        )
        workbook.create_sheet("Charts").add_chart(chart, "A1")
        workbook.save(path)
    finally:
        workbook.close()
    return path
