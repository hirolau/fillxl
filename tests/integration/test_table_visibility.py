"""Rows a fill writes must be visible, and totals must count them.

A template is often saved with a filter applied or with rows hidden by hand.
Excel does not reapply a filter when a table is resized, so without help the
filled rows stay hidden — and ``SUBTOTAL(109, …)``, which the default totals
row uses, leaves hidden rows out of the total it shows.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.worksheet.filters import AutoFilter, FilterColumn, Filters
from openpyxl.worksheet.table import Table, TableColumn, TableFormula

from fillxl import Engine
from tests._engines import ENGINES, _fill_output

if TYPE_CHECKING:
    from pathlib import Path


def _table_template(
    path: Path,
    *,
    body_rows: int = 1,
    totals: bool = False,
) -> Path:
    """A two-column table at ``A1:B*`` with *body_rows* body rows."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "Name"
    worksheet["B1"] = "Value"
    for row in range(2, 2 + body_rows):
        worksheet.cell(row=row, column=1, value=f"row{row}")
        worksheet.cell(row=row, column=2, value=row)
    last_row = 1 + body_rows
    columns = [TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")]
    if totals:
        last_row += 1
        worksheet.cell(row=last_row, column=1, value="Total")
        # An A1 formula, not a structured SUBTOTAL: this is the case that only
        # follows the body when the fill rewrites it.  Excel insists that a
        # totals row holds either a label or a declared totals formula, so the
        # column metadata has to describe it too, or Excel refuses the file.
        worksheet.cell(row=last_row, column=2, value=f"=SUM(B2:B{1 + body_rows})")
        columns[0].totalsRowLabel = "Total"
        columns[1].totalsRowFunction = "custom"
        columns[1].totalsRowFormula = TableFormula(attr_text=f"SUM(B2:B{1 + body_rows})")
    table = Table(
        displayName="Sales",
        ref=f"A1:B{last_row}",
        tableColumns=columns,
        totalsRowCount=1 if totals else None,
    )
    table.autoFilter = AutoFilter(ref=f"A1:B{1 + body_rows}")
    worksheet.add_table(table)
    workbook.save(path)
    return path


ROWS = [{"Name": "north", "Value": 1}, {"Name": "south", "Value": 2}]


@pytest.mark.parametrize("engine", ENGINES)
def test_a_body_row_hidden_in_the_template_is_shown_by_the_fill(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """Data written into a hidden row would exist and stay invisible."""
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    workbook["Data"].row_dimensions[2].hidden = True
    workbook.save(template)
    workbook.close()

    output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    workbook = openpyxl.load_workbook(output)
    try:
        worksheet = workbook["Data"]
        assert worksheet["A2"].value == "north"
        assert worksheet["A3"].value == "south"
        assert not worksheet.row_dimensions[2].hidden
        assert not worksheet.row_dimensions[3].hidden
    finally:
        workbook.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_a_filter_applied_in_the_template_is_cleared_by_the_fill(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """A filter that no longer matches the data hides rows and skews SUBTOTAL."""
    template = _table_template(tmp_path / "template.xlsx")
    workbook = openpyxl.load_workbook(template)
    table = workbook["Data"].tables["Sales"]
    table.autoFilter.filterColumn = [
        FilterColumn(colId=0, filters=Filters(filter=["missing"])),
    ]
    workbook.save(template)
    workbook.close()

    output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    workbook = openpyxl.load_workbook(output)
    try:
        table = workbook["Data"].tables["Sales"]
        assert not table.autoFilter.filterColumn
    finally:
        workbook.close()


@pytest.mark.parametrize("engine", ENGINES)
def test_the_totals_row_a1_formula_follows_the_body(engine: Engine, tmp_path: Path) -> None:
    """``=SUM(B2:B2)`` in a totals row has to become ``=SUM(B2:B3)``, in every engine."""
    template = _table_template(tmp_path / "template.xlsx", totals=True)

    output = _fill_output(template, tmp_path / "out.xlsx", engine, tables={"Sales": ROWS})

    workbook = openpyxl.load_workbook(output)
    try:
        worksheet = workbook["Data"]
        assert worksheet["A4"].value == "Total"
        assert worksheet["B4"].value == "=SUM(B2:B3)"
    finally:
        workbook.close()
