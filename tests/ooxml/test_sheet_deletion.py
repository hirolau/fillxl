"""Deleting a sheet has to renumber everything that indexes sheets by position.

``localSheetId`` on a defined name and ``firstSheet`` on the workbook view are
positions in ``<sheets>``, not sheet ids.  Removing a sheet above them shifts
every one of them, and a stale value points a print area or a filter at another
sheet — silently, because the numbers still parse.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

import openpyxl

from fillxl import FillXLConfig, fill
from tests.ooxml._patch import part_text, patch_part

if TYPE_CHECKING:
    from pathlib import Path

WORKBOOK_PART = "xl/workbook.xml"

DEFINED_NAMES = (
    "<definedNames>"
    '<definedName name="_xlnm.Print_Area" localSheetId="0">First!$A$1:$B$2</definedName>'
    '<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'
    "Second!$A$1:$B$2</definedName>"
    '<definedName name="_xlnm.Print_Area" localSheetId="2">Third!$A$1:$B$2</definedName>'
    "</definedNames>"
)


def _delete(path: Path, tmp_path: Path, *names: str) -> Path:
    return fill(
        path,
        tmp_path / "out.xlsx",
        engine="fillxl",
        config=FillXLConfig(delete_sheets=list(names)),
    )


def test_deleting_a_sheet_moves_later_sheets_print_areas_with_them(tmp_path: Path) -> None:
    """``Third`` keeps its print area after ``First`` is gone."""
    template = _three_sheets(tmp_path / "template.xlsx")
    patch_part(template, WORKBOOK_PART, _with_defined_names)

    output = _delete(template, tmp_path, "First")

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.sheetnames == ["Second", "Third"]
        assert workbook["Third"].print_area == "'Third'!$A$1:$B$2"
        assert workbook["Second"].print_area in (None, "")
    finally:
        workbook.close()


def test_deleting_a_sheet_removes_its_own_filter_database_name(tmp_path: Path) -> None:
    """A name scoped to the deleted sheet has nowhere left to point."""
    template = _three_sheets(tmp_path / "template.xlsx")
    patch_part(template, WORKBOOK_PART, _with_defined_names)

    output = _delete(template, tmp_path, "Second")

    workbook_xml = part_text(output, WORKBOOK_PART)
    assert "_xlnm._FilterDatabase" not in workbook_xml
    assert 'localSheetId="0">First!$A$1:$B$2' in workbook_xml
    assert 'localSheetId="1">Third!$A$1:$B$2' in workbook_xml


def test_deleting_the_first_visible_tab_resets_the_first_sheet_shown(tmp_path: Path) -> None:
    """``firstSheet`` past the last remaining sheet makes Excel scroll to nowhere."""
    template = _three_sheets(tmp_path / "template.xlsx")
    patch_part(
        template,
        WORKBOOK_PART,
        lambda xml: re.sub(r'firstSheet="\d+"', 'firstSheet="2"', xml, count=1),
    )

    output = _delete(template, tmp_path, "First")

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.views[0].firstSheet == 0
    finally:
        workbook.close()


def _three_sheets(path: Path) -> Path:
    workbook = openpyxl.Workbook()
    workbook.active.title = "First"
    workbook.active["A1"] = "first"
    workbook.create_sheet("Second")["A1"] = "second"
    workbook.create_sheet("Third")["A1"] = "third"
    workbook.save(path)
    workbook.close()
    return path


def _with_defined_names(workbook_xml: str) -> str:
    """Replace openpyxl's empty block with sheet-scoped names for all three sheets."""
    return re.sub(r"<definedNames\s*/>", DEFINED_NAMES, workbook_xml, count=1)
