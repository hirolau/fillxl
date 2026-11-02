"""The worksheet editor: spill detection and the cache behind it.

A spilled formula reserves cells that look empty, so every write has to ask
"is this cell inside a spill?".  Answering that by scanning the whole sheet on
every write made filling a table quadratic; these tests pin the answer and the
cost.
"""

from __future__ import annotations

import time
from typing import TYPE_CHECKING

from fillxl import fill
from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.worksheet import SheetPart
from tests.ooxml._build import body_rows, table_template
from tests.ooxml._patch import patch_part, put_cell

if TYPE_CHECKING:
    from pathlib import Path

SHEET_PART = "xl/worksheets/sheet1.xml"
SPILL = '<c r="H2" cm="1"><f t="array" ref="H2:H6">_xlfn.SEQUENCE(5)</f><v>1</v></c>'


def _sheet_with_spill(template: Path) -> SheetPart:
    patch_part(template, SHEET_PART, lambda xml: put_cell(xml, SPILL))
    return SheetPart(OoxmlPackage(template), SHEET_PART)


def test_filling_a_table_scales_linearly_with_its_row_count(tmp_path: Path) -> None:
    """Four times the rows must not cost nineteen times the time.

    The engine asks "is this cell inside a spill?" before every write.  A
    cached range list makes that a lookup; re-scanning the sheet made a
    4,000-row fill take two minutes.
    """
    template = table_template(tmp_path / "template.xlsx", rows=1)
    patch_part(template, SHEET_PART, lambda xml: put_cell(xml, SPILL))

    def elapsed(rows: int) -> float:
        start = time.perf_counter()
        fill(
            template,
            tmp_path / f"out{rows}.xlsx",
            engine="fillxl",
            tables={"Data": body_rows(rows)},
        )
        return time.perf_counter() - start

    thousand = elapsed(1000)
    four_thousand = elapsed(4000)

    # Linear growth is about 4; absolute times are flaky on CI, so only the
    # ratio is asserted.
    assert four_thousand / thousand < 8


def test_overwriting_a_spill_anchor_frees_its_range_for_later_writes(tmp_path: Path) -> None:
    """A cached range list must not outlive the formula that reserved it."""
    sheet = _sheet_with_spill(table_template(tmp_path / "template.xlsx", rows=1))
    assert sheet.array_range_at(3, 8) == "H2:H6"

    sheet.write_cell(2, 8, '<c r="H2"><v>1</v></c>')

    assert sheet.array_range_at(3, 8) is None


def test_writing_an_array_formula_reserves_its_range_immediately(tmp_path: Path) -> None:
    """A spill written through the editor is protected without re-reading the sheet."""
    sheet = SheetPart(
        OoxmlPackage(table_template(tmp_path / "template.xlsx", rows=1)),
        SHEET_PART,
    )
    assert sheet.array_range_at(5, 8) is None

    sheet.write_cell(2, 8, SPILL)

    assert sheet.array_range_at(5, 8) == "H2:H6"
