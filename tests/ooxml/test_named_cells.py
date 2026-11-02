"""Named cells use the same destination parser and the same write guards.

A defined name can point anywhere — at a formula, at a spilled range, at a
multi-cell range.  The fillxl engine has to answer those the way the other
engines do, or a fill succeeds on one engine and fails on another.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from fillxl import fill
from fillxl._warnings import FillXLWarning
from tests.ooxml._build import table_template
from tests.ooxml._patch import patch_part, put_cell

if TYPE_CHECKING:
    from pathlib import Path

SHEET_PART = "xl/worksheets/sheet1.xml"
SPILL = '<c r="H2" cm="1"><f t="array" ref="H2:H6">_xlfn.SEQUENCE(5)</f><v>1</v></c>'


def test_a_name_defined_by_a_formula_warns_instead_of_raising_a_key_error(
    tmp_path: Path,
) -> None:
    """Both engines say the same thing when a name has no cell destination."""
    template = _named_cell_template(
        tmp_path / "template.xlsx",
        "formula_name",
        "OFFSET(Data!$A$1,0,0)",
    )

    messages = {}
    for engine in ("openpyxl", "fillxl"):
        with pytest.warns(FillXLWarning) as caught:
            fill(
                template,
                tmp_path / f"{engine}.xlsx",
                engine=engine,
                values={"formula_name": 1},
            )
        messages[engine] = " ".join(str(warning.message) for warning in caught)

    assert "has no destinations, so nothing was written" in messages["fillxl"]
    assert messages["fillxl"] == messages["openpyxl"]


def test_a_named_cell_inside_a_spill_is_refused(tmp_path: Path) -> None:
    """Writing into a spilled range would destroy somebody's formula."""
    template = _named_cell_template(
        tmp_path / "template.xlsx",
        "in_spill",
        "Data!$H$3",
    )
    patch_part(template, SHEET_PART, lambda xml: put_cell(xml, SPILL))

    with pytest.raises(ValueError, match="part of the array formula H2:H6"):
        fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"in_spill": 1})


def _named_cell_template(path: Path, name: str, refers_to: str) -> Path:
    """A one-table workbook with a single defined name pointing anywhere."""
    table_template(path, rows=1)
    workbook = openpyxl.load_workbook(path)
    try:
        workbook.defined_names.add(DefinedName(name, attr_text=refers_to))
        workbook.save(path)
    finally:
        workbook.close()
    return path
