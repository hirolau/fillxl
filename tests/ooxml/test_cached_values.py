"""A cached formula result must not survive the formula it belonged to.

The fillxl engine cannot calculate, so it writes no results at all.  A result
the template cached would otherwise be read back as the filled report's number
by every reader that cannot recalculate — pandas, ``data_only=True``, previews.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl

from fillxl import fill
from tests.ooxml._build import body_rows, named_cell_template, table_with_sum_below_template
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


def test_another_sheets_cached_result_is_cleared_as_well(tmp_path: Path) -> None:
    """A ``Report`` sheet's ``=B1*.25`` is as stale as the table's own numbers.

    There is no dependency graph here to tell which cached result is still
    valid, so every formula in the workbook loses its result, exactly as with
    the openpyxl engine.
    """
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

    assert part_text(output, "xl/worksheets/sheet2.xml").count("<v>") == 0
    workbook = openpyxl.load_workbook(output, data_only=True)
    try:
        assert workbook["Report"]["B1"].value is None
    finally:
        workbook.close()


def test_a_named_cell_fill_clears_cached_results_too(tmp_path: Path) -> None:
    """The filled value is an input the report's formulas read."""
    template = named_cell_template(
        tmp_path / "template.xlsx",
        cells={"report_year": ("B2", 2024)},
    )
    patch_part(
        template,
        SHEET_PART,
        lambda xml: put_cell(xml, '<c r="D5"><f>YEAR(B2)</f><v>2024</v></c>'),
    )

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"report_year": 2025})

    assert '<c r="D5"><f>YEAR(B2)</f></c>' in part_text(output, SHEET_PART)


def test_a_sheet_without_formulas_is_copied_byte_for_byte(tmp_path: Path) -> None:
    """Clearing results must not rewrite parts that hold no formula."""
    template = named_cell_template(
        tmp_path / "template.xlsx",
        cells={"report_year": ("B2", 2024)},
        other_sheets=("Notes",),
    )

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"report_year": 2025})

    assert part_text(output, "xl/worksheets/sheet2.xml") == part_text(
        template,
        "xl/worksheets/sheet2.xml",
    )
