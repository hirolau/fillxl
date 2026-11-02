"""Shared formulas: one master cell and empty dependents.

Excel stores a filled-down or filled-right formula once, as
``<f t="shared" ref="C2:C10" si="0">A2*2</f>`` on a master, plus
``<f t="shared" si="0"/>`` on every other cell of the group.  A dependent has
no text, so an editor that only looks at ``<f>`` content sees a cell without a
formula and overwrites a live formula with a value.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import FillXLConfig, fill
from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.worksheet import SheetPart
from fillxl._warnings import DanglingReferenceWarning
from tests._spills import assert_well_formed
from tests.ooxml._build import (
    HEADERS,
    body_rows,
    table_template,
    table_with_calculated_column_template,
)
from tests.ooxml._patch import part_text, patch_part, put_cell, share_formulas
from tests.ooxml.test_parity import fill_both

if TYPE_CHECKING:
    from pathlib import Path

SHEET_PART = "xl/worksheets/sheet1.xml"
TRACKER_PART = "xl/worksheets/sheet2.xml"


def _data_without_total(rows: int) -> list[dict[str, object]]:
    return [{key: value for key, value in row.items() if key != "Total"} for row in body_rows(rows)]


def test_a_calculated_column_stored_as_shared_formulas_is_still_recognised_as_calculated(
    tmp_path: Path,
) -> None:
    """The definition is in the table part; the body only has dependents."""
    template = table_with_calculated_column_template(tmp_path / "template.xlsx", rows=3)
    patch_part(template, SHEET_PART, lambda xml: share_formulas(xml, "D2", "D4"))

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": _data_without_total(4)},
    )

    assert_well_formed(output)
    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["Data"]
        assert [sheet.cell(row=row, column=4).value for row in range(2, 6)] == [
            "=Data[[#This Row],[Amount]]*2"
        ] * 4
    finally:
        workbook.close()


def test_shrinking_a_table_leaves_no_orphaned_shared_formulas_behind(
    tmp_path: Path,
) -> None:
    """Every cell of the group has to leave with its master."""
    template = table_with_calculated_column_template(tmp_path / "template.xlsx", rows=4)
    patch_part(template, SHEET_PART, lambda xml: share_formulas(xml, "D2", "D5"))

    expected, actual = fill_both(
        template,
        tmp_path,
        tables={"Data": _data_without_total(2)},
    )

    assert expected == actual
    sheet_xml = part_text(tmp_path / "fillxl" / template.name, SHEET_PART)
    assert 't="shared"' not in sheet_xml


def test_a_shared_sum_below_the_table_is_resized_without_changing_its_neighbours(
    tmp_path: Path,
) -> None:
    """Each member of a filled-right group sums its own column."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    patch_part(template, SHEET_PART, _add_shared_sum)

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(4)},
    )

    assert_well_formed(output)
    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["Data"]
        assert sheet["B20"].value == "=SUM(B2:B5)"
        assert sheet["C20"].value == "=SUM(C2:C5)"
        assert sheet["D20"].value == "=SUM(D2:D5)"
    finally:
        workbook.close()


def _add_shared_sum(xml: str) -> str:
    """Put a shared ``=SUM`` under the table, filled right across B:D."""
    for column in ("B", "C", "D"):
        xml = put_cell(xml, f'<c r="{column}20"><f>SUM({column}2:{column}2)</f></c>')
    return share_formulas(xml, "B20", "D20")


def test_a_deleted_sheet_referenced_only_by_shared_dependents_is_still_reported(
    tmp_path: Path,
) -> None:
    """The dependents carry no text; the reference still has to be found."""
    template = _template_with_shared_reference(tmp_path / "template.xlsx")

    with pytest.warns(DanglingReferenceWarning) as caught:
        fill(
            template,
            tmp_path / "out.xlsx",
            engine="fillxl",
            tables={"Data": body_rows(2)},
            config=FillXLConfig(delete_sheets=["Old"]),
        )

    messages = " ".join(str(warning.message) for warning in caught)
    assert "Tracker!A2" in messages or "Tracker!A3" in messages


def test_a_shared_formula_group_outside_the_table_is_left_byte_identical(
    tmp_path: Path,
) -> None:
    """A group the fill does not touch keeps its exact markup."""
    template = _template_with_shared_reference(tmp_path / "template.xlsx")
    before = part_text(template, TRACKER_PART)
    assert 't="shared"' in before

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(3)},
    )

    assert part_text(output, TRACKER_PART) == before


def test_a_dependent_is_read_as_the_master_moved_to_its_own_cell(tmp_path: Path) -> None:
    """The index resolves a dependent to the formula Excel would show."""
    template = table_template(tmp_path / "template.xlsx", rows=1)

    def add_group(xml: str) -> str:
        xml = put_cell(xml, '<c r="B20"><f>SUM(B2:B2)</f></c>')
        xml = put_cell(xml, '<c r="C20"><f>SUM(C2:C2)</f></c>')
        return share_formulas(xml, "B20", "C20")

    patch_part(template, SHEET_PART, add_group)

    sheet = SheetPart(OoxmlPackage(template), SHEET_PART)

    assert sheet.formula(20, 2) == "=SUM(B2:B2)"
    assert sheet.formula(20, 3) == "=SUM(C2:C2)"


def _template_with_shared_reference(path: Path) -> Path:
    """``Tracker`` holds a shared group pointing at the sheet ``Old``."""
    workbook = openpyxl.Workbook()
    data = workbook.active
    data.title = "Data"
    for offset, name in enumerate(HEADERS, start=1):
        data.cell(row=1, column=offset, value=name)
    data.cell(row=2, column=1, value="north")
    data.add_table(
        Table(
            displayName="Data",
            ref="A1:E2",
            tableColumns=[
                TableColumn(id=index, name=name) for index, name in enumerate(HEADERS, start=1)
            ],
        ),
    )
    tracker = workbook.create_sheet("Tracker")
    for row in (1, 2, 3):
        tracker.cell(row=row, column=1, value=f"='Old'!A{row}+1")
    workbook.create_sheet("Old")
    workbook.save(path)
    workbook.close()

    patch_part(path, TRACKER_PART, lambda xml: share_formulas(xml, "A1", "A3"))
    return path
