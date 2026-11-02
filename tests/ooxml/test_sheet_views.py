"""Activating a sheet that has no ``<sheetViews>`` block.

Excel tolerates a sheet without views; the schema does not tolerate them
anywhere.  ``<sheetViews>`` belongs after ``<sheetPr>`` and ``<dimension>`` and
before the row defaults, the columns and the data — inserting it at the first
``>`` in the part lands inside the XML declaration.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.chart import BarChart

from fillxl import FillXLConfig, fill
from tests.ooxml._build import body_rows, table_template
from tests.ooxml._patch import part_text, patch_part

if TYPE_CHECKING:
    from collections.abc import Callable
    from pathlib import Path

CHART_SHEET_PART = "xl/chartsheets/sheet1.xml"
SHEET_PART = "xl/worksheets/sheet1.xml"


def test_activating_a_chartsheet_without_views_keeps_the_part_well_formed(
    tmp_path: Path,
    assert_well_formed: Callable[[Path], None],
) -> None:
    """The block goes before ``<drawing>``, where the schema wants it."""
    template = _workbook_with_chartsheet(tmp_path / "template.xlsx")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(2)},
        config=FillXLConfig(active_sheet="Chart"),
    )

    assert_well_formed(output)
    chart_sheet = part_text(output, CHART_SHEET_PART)
    assert "<sheetViews>" in chart_sheet
    assert chart_sheet.index("<sheetViews>") < chart_sheet.index("<drawing")


def test_activating_a_worksheet_without_views_or_format_properties_keeps_the_part_well_formed(
    tmp_path: Path,
    assert_well_formed: Callable[[Path], None],
) -> None:
    """With no ``<sheetFormatPr>`` the block goes before ``<sheetData>``."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    patch_part(template, SHEET_PART, _without_views_and_format)

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(2)},
        config=FillXLConfig(active_sheet="Data"),
    )

    assert_well_formed(output)
    sheet = part_text(output, SHEET_PART)
    assert "<sheetViews>" in sheet
    assert sheet.index("<sheetViews>") < sheet.index("<sheetData>")


def test_activating_a_sheet_with_no_anchor_at_all_lands_after_the_root_element(
    tmp_path: Path,
    assert_well_formed: Callable[[Path], None],
) -> None:
    """The last resort must not append the block to the XML declaration."""
    template = _workbook_with_chartsheet(tmp_path / "template.xlsx")
    patch_part(template, CHART_SHEET_PART, _without_views_and_drawing)

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(2)},
        config=FillXLConfig(active_sheet="Chart"),
    )

    assert_well_formed(output)
    chart_sheet = part_text(output, CHART_SHEET_PART)
    assert re.match(r"^<\?xml[^>]*\?>\s*<chartsheet\b[^>]*><sheetViews>", chart_sheet)


def _workbook_with_chartsheet(path: Path) -> Path:
    """A table plus a chart sheet that has no ``<sheetViews>``."""
    table_template(path, rows=1)
    workbook = openpyxl.load_workbook(path)
    try:
        chart_sheet = workbook.create_chartsheet("Chart")
        chart_sheet.add_chart(BarChart())
        workbook.save(path)
    finally:
        workbook.close()
    patch_part(
        path,
        CHART_SHEET_PART,
        lambda xml: re.sub(r"<sheetViews>.*?</sheetViews>", "", xml, flags=re.S),
    )
    return path


def _without_views_and_format(sheet_xml: str) -> str:
    without_views = re.sub(r"<sheetViews>.*?</sheetViews>", "", sheet_xml, flags=re.S)
    return re.sub(r"<sheetFormatPr\b[^>]*/>", "", without_views)


def _without_views_and_drawing(chart_sheet_xml: str) -> str:
    """A chartsheet the schema allows but that has no element to anchor to."""
    without_views = re.sub(r"<sheetViews>.*?</sheetViews>", "", chart_sheet_xml, flags=re.S)
    stripped = re.sub(r"<drawing\b[^>]*/>", "", without_views)
    if stripped.startswith("<?xml"):
        return stripped
    declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return f"{declaration}{stripped}"
