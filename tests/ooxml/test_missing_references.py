"""Cells that carry no ``r`` attribute.

The reference is optional in the file format: a cell without one is simply the
next cell of its row.  Missing that rule loses the cell's position — and a
rewritten row is re-rendered from the parsed cells, so the value would move to
column A.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import fill
from tests.ooxml._build import body_rows, table_template
from tests.ooxml._patch import patch_part

if TYPE_CHECKING:
    from pathlib import Path

SHEET_PART = "xl/worksheets/sheet1.xml"


def test_cells_without_references_survive_when_their_row_is_rewritten(
    tmp_path: Path,
) -> None:
    """A cell is the next one of its row, not the first column."""
    template = table_template(tmp_path / "template.xlsx", rows=2)
    patch_part(template, SHEET_PART, _strip_body_references)

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(2)},
    )

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["Data"]
        assert sheet["A2"].value == "R0"
        assert sheet["B2"].value == 1
        assert sheet["D2"].value == 2
        assert sheet["E2"].value == "n0"
        assert sheet["A3"].value == "R1"
        assert sheet["B3"].value == 2
    finally:
        workbook.close()


def test_cells_without_references_count_as_occupied_when_the_table_grows(
    tmp_path: Path,
) -> None:
    """A value in the growth path blocks the fill, with or without a reference."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    patch_part(template, SHEET_PART, _blocker_without_reference)

    with pytest.raises(ValueError, match="A3 already contains data"):
        fill(
            template,
            tmp_path / "out.xlsx",
            engine="fillxl",
            tables={"Data": body_rows(3)},
        )


def _strip_body_references(sheet_xml: str) -> str:
    """Remove the ``r`` attribute of every cell in rows 2 and 3."""
    import re

    for row in (2, 3):
        start = sheet_xml.find(f'<row r="{row}"')
        end = sheet_xml.find("</row>", start)
        body = sheet_xml[start:end]
        stripped = re.sub(r' r="[A-Z]+\d+"', "", body)
        sheet_xml = f"{sheet_xml[:start]}{stripped}{sheet_xml[end:]}"
    return sheet_xml


def _blocker_without_reference(sheet_xml: str) -> str:
    """Add row 3 holding one value-carrying cell that has no ``r``."""
    row = '<row r="3"><c t="inlineStr"><is><t>blocker</t></is></c></row>'
    closing = sheet_xml.find("</sheetData>")
    return f"{sheet_xml[:closing]}{row}{sheet_xml[closing:]}"
