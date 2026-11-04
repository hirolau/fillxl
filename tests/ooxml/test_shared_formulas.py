"""Shared formulas: one master cell and empty dependents.

Excel stores a filled-down or filled-right formula once, as
``<f t="shared" ref="C2:C10" si="0">A2*2</f>`` on a master, plus
``<f t="shared" si="0"/>`` on every other cell of the group.  A dependent has
no text, so an editor that only looks at ``<f>`` content sees a cell without a
formula and overwrites a live formula with a value.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import FillXLConfig, fill
from fillxl._ooxml import xmltext
from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.worksheet import SheetPart, column_index
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


def _share_with_cached(dependants: str, master: str, sheet_xml: str) -> str:
    """One shared group in column C, with cached results Windows could freeze."""
    sheet_xml = put_cell(sheet_xml, f'<c r="C2">{master}</c>')
    sheet_xml = put_cell(sheet_xml, f'<c r="C3"><f t="shared" si="0"/><v>{dependants}</v></c>')
    return put_cell(sheet_xml, f'<c r="C4"><f t="shared" si="0"/><v>{dependants}</v></c>')


def _masters_by_si(sheet_xml: str) -> dict[str, tuple[int, int, str]]:
    """``si`` → (row, column, ref) of every shared master in *sheet_xml*."""
    pattern = re.compile(r"<c r=\"(?P<cell>[A-Z]+)(?P<row>\d+)\">(?P<rest>.*?)</c>", re.S)
    masters: dict[str, tuple[int, int, str]] = {}
    for match in pattern.finditer(sheet_xml):
        cell_xml = match.group(0)
        tag = xmltext.start_tag(cell_xml, "f")
        if tag is None or xmltext.attribute(tag, "t") != "shared":
            continue
        si = xmltext.attribute(tag, "si")
        ref = xmltext.attribute(tag, "ref")
        text = xmltext.element_text(cell_xml, "f")
        if si is None or ref is None or text is None:
            continue
        column = column_index(match.group("cell"))
        masters[si] = (int(match.group("row")), column, ref)
    return masters


def _assert_every_stub_has_a_covering_master(sheet_xml: str) -> None:
    """Every ``<f t="shared" si="N"/>`` must be covered by a master of its ``si``."""
    masters = _masters_by_si(sheet_xml)
    stub = re.compile(
        r"<c r=\"(?P<cell>[A-Z]+)(?P<row>\d+)\"><f t=\"shared\" si=\"(?P<si>\d+)\"/></c>",
    )
    for match in stub.finditer(sheet_xml):
        si = match.group("si")
        assert si in masters, f"{match.group('cell')}{match.group('row')} has no master"
        row = int(match.group("row"))
        column = column_index(match.group("cell"))
        _, _, ref = masters[si]
        bounds = CellRange(ref)
        assert bounds.min_row <= row <= bounds.max_row
        assert bounds.min_col <= column <= bounds.max_col


def _isolated_shared_template(path: Path, dependants: str, master: str) -> Path:
    """A group in ``C2:C4`` next to an unrelated named cell ``filler``."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in (2, 3, 4):
        sheet.cell(row=row, column=1, value=row)
        sheet.cell(row=row, column=3, value=f'=IF(A{row}>0,A{row},"")')
    sheet.cell(row=10, column=1, value=0)
    workbook.defined_names.add(DefinedName("filler", attr_text="Data!$A$10"))
    workbook.save(path)
    workbook.close()
    patch_part(path, SHEET_PART, lambda xml: _share_with_cached(dependants, master, xml))
    return path


def test_a_shared_formula_whose_first_cell_shows_an_empty_string_survives_saving(
    tmp_path: Path,
) -> None:
    """An empty cached result must not leave the dependants without a master."""
    template = _isolated_shared_template(
        tmp_path / "template.xlsx",
        dependants="5",
        master='<f t="shared" ref="C2:C4" si="0">IF(A2&gt;0,A2,"")</f><v></v>',
    )

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"filler": 42})

    assert_well_formed(output)
    sheet_xml = part_text(output, SHEET_PART)
    assert 't="shared"' in sheet_xml
    _assert_every_stub_has_a_covering_master(sheet_xml)


def test_a_shared_formula_with_cached_results_stays_shared_after_saving(
    tmp_path: Path,
) -> None:
    """The group keeps one master and its stubs; only the results go."""
    template = _isolated_shared_template(
        tmp_path / "template.xlsx",
        dependants="5",
        master='<f t="shared" ref="C2:C4" si="0">IF(A2&gt;0,A2,"")</f><v>7</v>',
    )

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"filler": 42})

    assert_well_formed(output)
    sheet_xml = part_text(output, SHEET_PART)
    _assert_every_stub_has_a_covering_master(sheet_xml)
    assert sheet_xml.count('<f t="shared" si="0"/>') == 2
    group = re.findall(r'<c r="C[234]">.*?</c>', sheet_xml)
    assert len(group) == 3
    assert all("<v" not in cell for cell in group)


def test_restyling_a_cell_of_a_shared_formula_keeps_its_formula(tmp_path: Path) -> None:
    """Restyling reads the cell after the group has been unshared."""
    template = _isolated_shared_template(
        tmp_path / "template.xlsx",
        dependants="5",
        master='<f t="shared" ref="C2:C4" si="0">IF(A2&gt;0,A2,"")</f><v>7</v>',
    )
    sheet = SheetPart(OoxmlPackage(template), SHEET_PART)

    sheet.set_style(3, 3, "4")

    xml = sheet.cell(3, 3)
    assert xml is not None
    assert 't="shared"' not in xml
    assert sheet.formula(3, 3) == '=IF(A3>0,A3,"")'
    assert sheet.formula(4, 3) == '=IF(A4>0,A4,"")'


def test_writing_stale_shared_formula_xml_is_refused(tmp_path: Path) -> None:
    """A shared stub is never what a caller of ``write_cell`` means to store."""
    template = _isolated_shared_template(
        tmp_path / "template.xlsx",
        dependants="5",
        master='<f t="shared" ref="C2:C4" si="0">IF(A2&gt;0,A2,"")</f><v>7</v>',
    )
    sheet = SheetPart(OoxmlPackage(template), SHEET_PART)

    with pytest.raises(RuntimeError, match="C3"):
        sheet.write_cell(3, 3, '<c r="C3"><f t="shared" si="0"/></c>')


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


def test_a_named_cell_written_over_a_shared_master_keeps_the_group_readable(
    tmp_path: Path,
) -> None:
    """Writing the master must not leave its dependents without a formula."""
    template = _named_shared_template(tmp_path / "template.xlsx", "over_master")

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"over_master": 99})

    assert_well_formed(output)
    sheet_xml = part_text(output, SHEET_PART)
    assert 't="shared"' not in sheet_xml
    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["Data"]
        assert sheet["B2"].value == 99
        assert sheet["B3"].value == "=A3*2"
        assert sheet["B4"].value == "=A4*2"
    finally:
        workbook.close()


def test_a_named_cell_written_over_a_shared_dependent_keeps_the_group_readable(
    tmp_path: Path,
) -> None:
    """Writing a dependent must not leave the others pointing at a stale group."""
    template = _named_shared_template(tmp_path / "template.xlsx", "over_dependent")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        values={"over_dependent": 99},
    )

    assert_well_formed(output)
    sheet_xml = part_text(output, SHEET_PART)
    assert 't="shared"' not in sheet_xml
    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["Data"]
        assert sheet["B2"].value == "=A2*2"
        assert sheet["B3"].value == 99
        assert sheet["B4"].value == "=A4*2"
    finally:
        workbook.close()


def _named_shared_template(path: Path, name: str) -> Path:
    """``B2:B4`` is one shared group; *name* points at ``B3``."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in (2, 3, 4):
        sheet.cell(row=row, column=1, value=row)
        sheet.cell(row=row, column=2, value=f"=A{row}*2")
    target = "B3" if name == "over_dependent" else "B2"
    workbook.defined_names.add(DefinedName(name, attr_text=f"Data!${target[0]}${target[1]}"))
    workbook.save(path)
    workbook.close()

    patch_part(path, SHEET_PART, lambda xml: share_formulas(xml, "B2", "B4"))
    return path


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
