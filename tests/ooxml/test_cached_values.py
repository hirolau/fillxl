"""A cached formula result must not survive the formula it belonged to.

The fillxl engine cannot calculate, so it writes no results at all.  A result
the template cached would otherwise be read back as the filled report's number
by every reader that cannot recalculate — pandas, ``data_only=True``, previews.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl

from fillxl import fill
from tests.ooxml._build import body_rows, table_with_sum_below_template
from tests.ooxml._patch import part_text, patch_part, put_cell

if TYPE_CHECKING:
    from pathlib import Path

SHEET_PART = "xl/worksheets/sheet1.xml"


def _rows(count: int) -> list[dict[str, object]]:
    return list(body_rows(count))


def test_a_rewritten_formula_loses_the_template_result(tmp_path: Path) -> None:
    """``=SUM(B2:B2)`` becomes ``=SUM(B2:B4)``; the cached 5 belongs to neither."""
    template = table_with_sum_below_template(tmp_path / "template.xlsx", rows=1)
    patch_part(
        template,
        SHEET_PART,
        lambda xml: put_cell(xml, '<c r="G2"><f>SUM(B2:B2)</f><v>5</v></c>'),
    )

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Data": _rows(3)})

    sheet = part_text(output, SHEET_PART)
    assert '<c r="G2"><f>SUM(B2:B4)</f></c>' in sheet


def test_a_reader_that_cannot_calculate_sees_blank_instead_of_a_stale_number(
    tmp_path: Path,
) -> None:
    template = table_with_sum_below_template(tmp_path / "template.xlsx", rows=1)
    patch_part(
        template,
        SHEET_PART,
        lambda xml: put_cell(xml, '<c r="G2"><f>SUM(B2:B2)</f><v>5</v></c>'),
    )

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Data": _rows(3)})

    workbook = openpyxl.load_workbook(output, data_only=True)
    try:
        assert workbook["Data"]["G2"].value is None
    finally:
        workbook.close()


def test_an_untouched_sheet_keeps_its_cached_results(tmp_path: Path) -> None:
    """Only the sheet a fill writes is recalculated, as the README says."""
    template = table_with_sum_below_template(
        tmp_path / "template.xlsx",
        rows=1,
        other_sheet="Report",
    )
    report = openpyxl.load_workbook(template)
    report["Report"]["B1"] = "=1+1"
    report.save(template)
    report.close()
    patch_part(
        template,
        "xl/worksheets/sheet2.xml",
        lambda xml: put_cell(xml, '<c r="B1"><f>1+1</f><v>2</v></c>'),
    )

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", tables={"Data": _rows(3)})

    assert "<v>2</v>" in part_text(output, "xl/worksheets/sheet2.xml")
