"""Partly formatted text in a template stays formatted."""

from __future__ import annotations

import zipfile
from pathlib import Path

import openpyxl
import pytest
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from fillxl import fill
from tests._assets import SPACER_RICH_TEXT_TEMPLATE
from tests.parity.builders import table_template

_ROWS = [{"fund": f"F{i}", "nav": i} for i in range(4)]


def _rich_text() -> CellRichText:
    return CellRichText([TextBlock(InlineFont(b=True), "Bold"), " plain"])


def test_partly_bold_text_in_the_template_survives_the_fill(tmp_path: Path) -> None:
    """openpyxl flattens rich text to a plain string unless asked to keep it."""
    template = table_template(tmp_path / "template.xlsx", top_row=3)
    workbook = openpyxl.load_workbook(template, rich_text=True)
    workbook["Data"]["A1"] = _rich_text()
    workbook.save(template)
    workbook.close()

    output = tmp_path / "out.xlsx"
    fill(template=template, output=output, engine="openpyxl", tables={"Sales": _ROWS})

    workbook = openpyxl.load_workbook(output, rich_text=True)
    try:
        value = workbook["Data"]["A1"].value
        assert isinstance(value, CellRichText)
        assert value[0].font.b  # type: ignore[union-attr]
        assert str(value) == "Bold plain"
    finally:
        workbook.close()


def test_rich_text_in_a_totals_row_survives_a_resize(tmp_path: Path) -> None:
    """The totals row is recreated below the new body, formatting included."""
    template = table_template(tmp_path / "template.xlsx", body=(("A", 1), ("B", 2)), totals=True)
    workbook = openpyxl.load_workbook(template, rich_text=True)
    worksheet = workbook["Data"]
    worksheet["A4"] = _rich_text()
    workbook.save(template)
    workbook.close()

    output = tmp_path / "out.xlsx"
    fill(template=template, output=output, engine="openpyxl", tables={"Sales": _ROWS})

    workbook = openpyxl.load_workbook(output, rich_text=True)
    try:
        worksheet = workbook["Data"]
        table = worksheet.tables["Sales"]
        assert table.ref == "A1:B6"
        value = worksheet["A6"].value
        assert isinstance(value, CellRichText)
        assert value[0].font.b  # type: ignore[union-attr]
    finally:
        workbook.close()


def spacer_rich_text() -> CellRichText:
    """Header text padded with whitespace-only runs, as Excel writes them."""
    spacer = InlineFont(sz=2.95)
    return CellRichText(
        [
            TextBlock(spacer, " \n"),
            TextBlock(InlineFont(b=True), "Header"),
            TextBlock(spacer, "\n "),
        ],
    )


def test_whitespace_only_runs_are_written_with_xml_space_preserve(tmp_path: Path) -> None:
    """Excel refuses ("corrupt") a bare whitespace-only <t>, which openpyxl writes."""
    template = table_template(tmp_path / "template.xlsx", top_row=3)
    workbook = openpyxl.load_workbook(template, rich_text=True)
    workbook["Data"]["A1"] = spacer_rich_text()
    workbook.save(template)
    workbook.close()

    output = tmp_path / "out.xlsx"
    fill(template=template, output=output, engine="openpyxl", tables={"Sales": _ROWS})

    with zipfile.ZipFile(output) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode().replace("\r\n", "\n")
    assert '<t xml:space="preserve"> \n</t>' in sheet
    assert '<t xml:space="preserve">\n </t>' in sheet
    assert "<t> \n</t>" not in sheet

    workbook = openpyxl.load_workbook(output, rich_text=True)
    try:
        assert str(workbook["Data"]["A1"].value) == " \nHeader\n "
    finally:
        workbook.close()


def test_a_workbook_written_by_excel_keeps_its_spacer_runs_preserved(tmp_path: Path) -> None:
    """Excel pads every header cell with whitespace-only runs; a fill must keep them intact."""
    output = tmp_path / "out.xlsx"
    fill(template=SPACER_RICH_TEXT_TEMPLATE, output=output, engine="openpyxl")

    with zipfile.ZipFile(output) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode().replace("\r\n", "\n")
    assert '<t xml:space="preserve"> \n</t>' in sheet
    assert "<t> \n</t>" not in sheet

    workbook = openpyxl.load_workbook(output, rich_text=True)
    try:
        value = workbook["Portfolio Holding"]["A13"].value
        assert isinstance(value, CellRichText)
        assert str(value) == " \nInstrument Type\n "
    finally:
        workbook.close()


@pytest.mark.parametrize("column", [1, 2])
def test_a_cell_outside_the_table_is_never_touched(tmp_path: Path, column: int) -> None:
    """Only the table body is rewritten."""
    template = table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template, rich_text=True)
    workbook["Data"].cell(row=10, column=column, value=_rich_text())
    workbook.save(template)
    workbook.close()

    output = tmp_path / "out.xlsx"
    fill(template=template, output=output, engine="openpyxl", tables={"Sales": _ROWS})

    workbook = openpyxl.load_workbook(output, rich_text=True)
    try:
        assert isinstance(workbook["Data"].cell(row=10, column=column).value, CellRichText)
    finally:
        workbook.close()
