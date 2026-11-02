"""Conditional formatting that follows the table body.

A data bar authored in Excel is stored as an ``x14`` extension block, and an
extension block keeps its range in a child ``<xm:sqref>`` element rather than
in an attribute.  Only looking at the attribute leaves the bar on the old body
rows, so the new rows are formatted differently from the rest of the table.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from fillxl import fill
from tests.ooxml._build import body_rows, table_template
from tests.ooxml._patch import part_text, patch_part

if TYPE_CHECKING:
    from pathlib import Path

SHEET_PART = "xl/worksheets/sheet1.xml"

X14_DATA_BAR = (
    '<extLst><ext uri="{fillxl-test}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    "<x14:conditionalFormattings><x14:conditionalFormatting "
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    '<x14:cfRule type="dataBar" id="{fillxl-test-rule}">'
    '<x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/>'
    '<x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>'
    "<xm:sqref>A2:E2</xm:sqref></x14:conditionalFormatting>"
    "</x14:conditionalFormattings></ext></extLst>"
)


def test_a_data_bar_keeps_its_extended_styling_on_new_rows(tmp_path: Path) -> None:
    """The child ``<xm:sqref>`` grows with the body, like the attribute form."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    patch_part(
        template,
        SHEET_PART,
        lambda xml: xml.replace("</worksheet>", f"{X14_DATA_BAR}</worksheet>"),
    )

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(4)},
    )

    sheet = part_text(output, SHEET_PART)
    assert "<xm:sqref>A2:E5</xm:sqref>" in sheet
