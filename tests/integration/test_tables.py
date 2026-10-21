"""Table filling integration tests — unique coverage not in story suite.

Tests style copying behaviour that the story tests don't cover:
font/fill inheritance and the copy_table_styles=False config flag.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import FillXLConfig, fill
from tests._assets import TEMPLATE

if TYPE_CHECKING:
    from pathlib import Path


def _data_rows(count: int) -> list[dict[str, object]]:
    """Return standard data rows."""
    return [
        {
            "Col1": f"row{i}",
            "Col2": i,
            "ColNotFilled": f"nf{i}",
            "ColFormula": None,
            "TestFormattedPct": 0.0,
        }
        for i in range(1, count + 1)
    ]


def test_new_rows_inherit_font_and_fill_from_first_data_row(output_path: Path) -> None:
    """New rows should have the same font, fill, and number format as row 1."""
    fill(
        template=TEMPLATE,
        output=output_path,
        tables={"TableTest": _data_rows(5)},
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["TestTable"]
        first_row_cells = [ws.cell(row=2, column=c) for c in range(1, 6)]

        for row in range(3, 7):
            for col in range(1, 6):
                first = first_row_cells[col - 1]
                cell = ws.cell(row=row, column=col)
                assert cell.font.name == first.font.name, (
                    f"Font name mismatch at {col}{row}: {cell.font.name} != {first.font.name}"
                )
                assert cell.font.size == first.font.size, f"Font size mismatch at {col}{row}"
                assert cell.fill.fgColor.rgb == first.fill.fgColor.rgb, (
                    f"Fill color mismatch at {col}{row}"
                )
    finally:
        wb.close()


def test_new_rows_do_not_inherit_styles_when_copy_table_styles_is_false(output_path: Path) -> None:
    """copy_table_styles=False should not copy styles to new rows."""
    fill(
        template=TEMPLATE,
        output=output_path,
        tables={"TableTest": _data_rows(5)},
        config=FillXLConfig(copy_table_styles=False),
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["TestTable"]
        assert ws.cell(row=6, column=1).value == "row5"
        assert ws.cell(row=6, column=2).value == 5

        new_row_fill = ws.cell(row=6, column=1).fill
        assert new_row_fill.fgColor.rgb == "00000000", (
            f"Expected default fill on new row, got {new_row_fill.fgColor.rgb}"
        )
        assert new_row_fill.patternType is None, (
            f"Expected no fill pattern on new row, got {new_row_fill.patternType}"
        )
    finally:
        wb.close()


def _build_totals_template(path: Path, *, data_rows: int) -> None:
    """Create a one-table workbook whose table has a totals row."""
    from openpyxl.styles import Font
    from openpyxl.worksheet.filters import AutoFilter
    from openpyxl.worksheet.table import Table, TableColumn

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Name"
    ws["B1"] = "Value"
    for offset in range(data_rows):
        ws.cell(row=2 + offset, column=1, value=f"old{offset}")
        ws.cell(row=2 + offset, column=2, value=offset)
    totals_row = data_rows + 2
    ws.cell(row=totals_row, column=1, value="Total").font = Font(bold=True)
    ws.cell(row=totals_row, column=2, value="=SUBTOTAL(109,Totals[Value])")
    ws.add_table(
        Table(
            displayName="Totals",
            ref=f"A1:B{totals_row}",
            autoFilter=AutoFilter(ref=f"A1:B{totals_row - 1}"),
            tableColumns=[
                TableColumn(id=1, name="Name", totalsRowLabel="Total"),
                TableColumn(id=2, name="Value", totalsRowFunction="sum"),
            ],
            totalsRowCount=1,
            totalsRowShown=True,
        ),
    )
    wb.save(path)
    wb.close()


def test_a_totals_row_stays_below_the_new_data(output_path: Path) -> None:
    """The data replaces the body and the totals row moves below it intact."""
    template = output_path.with_name("totals.xlsx")
    _build_totals_template(template, data_rows=1)

    fill(
        template=template,
        output=output_path,
        tables={"Totals": [{"Name": f"row{i}", "Value": i} for i in range(1, 4)]},
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["Data"]
        table = ws.tables["Totals"]

        assert table.ref == "A1:B5"
        assert table.totalsRowCount == 1
        # Excel excludes the totals row from the auto-filter range.
        assert table.autoFilter.ref == "A1:B4"
        assert [col.totalsRowLabel for col in table.tableColumns] == ["Total", None]
        assert [col.totalsRowFunction for col in table.tableColumns] == [None, "sum"]

        assert [ws.cell(row=r, column=1).value for r in range(2, 5)] == [
            "row1",
            "row2",
            "row3",
        ]
        assert [ws.cell(row=r, column=2).value for r in range(2, 5)] == [1, 2, 3]

        assert ws.cell(row=5, column=1).value == "Total"
        assert ws.cell(row=5, column=2).value == "=SUBTOTAL(109,Totals[Value])"
        assert ws.cell(row=5, column=1).font.bold is True
    finally:
        wb.close()


def test_a_totals_row_moves_up_when_the_data_shrinks(output_path: Path) -> None:
    """Shrinking clears the old body rows and moves the totals row up."""
    template = output_path.with_name("totals_shrink.xlsx")
    _build_totals_template(template, data_rows=3)

    fill(
        template=template,
        output=output_path,
        tables={"Totals": [{"Name": "only", "Value": 7}]},
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["Data"]
        table = ws.tables["Totals"]

        assert table.ref == "A1:B3"
        assert table.autoFilter.ref == "A1:B2"
        assert ws.cell(row=2, column=1).value == "only"
        assert ws.cell(row=3, column=1).value == "Total"
        assert ws.cell(row=3, column=2).value == "=SUBTOTAL(109,Totals[Value])"
        assert ws.cell(row=4, column=1).value is None
        assert ws.cell(row=5, column=1).value is None
    finally:
        wb.close()


def test_growing_a_totals_table_by_one_row_keeps_the_totals_row(output_path: Path) -> None:
    """The existing totals row is not mistaken for content below the table."""
    template = output_path.with_name("totals_grow_one.xlsx")
    _build_totals_template(template, data_rows=1)

    fill(
        template=template,
        output=output_path,
        tables={"Totals": [{"Name": "a", "Value": 1}, {"Name": "b", "Value": 2}]},
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["Data"]
        table = ws.tables["Totals"]

        assert table.ref == "A1:B4"
        assert [ws.cell(row=r, column=2).value for r in range(2, 4)] == [1, 2]
        assert ws.cell(row=4, column=1).value == "Total"
    finally:
        wb.close()


def test_content_below_a_totals_table_is_still_rejected(output_path: Path) -> None:
    """Expanding past the totals row must not overwrite unrelated content."""
    template = output_path.with_name("totals_overflow.xlsx")
    _build_totals_template(template, data_rows=1)

    wb = openpyxl.load_workbook(template)
    wb["Data"]["A4"] = "DO NOT OVERWRITE"
    wb.save(template)
    wb.close()

    with pytest.raises(ValueError, match="cannot expand"):
        fill(
            template=template,
            output=output_path,
            tables={"Totals": [{"Name": f"row{i}", "Value": i} for i in range(1, 4)]},
        )


def test_empty_data_keeps_one_body_row_and_the_totals_row(output_path: Path) -> None:
    """Empty data leaves a single cleared body row above the totals row."""
    template = output_path.with_name("totals_empty.xlsx")
    _build_totals_template(template, data_rows=2)

    fill(
        template=template,
        output=output_path,
        tables={"Totals": {"Name": [], "Value": []}},
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["Data"]
        table = ws.tables["Totals"]

        assert table.ref == "A1:B3"
        assert table.autoFilter.ref == "A1:B2"
        assert ws.cell(row=2, column=1).value is None
        assert ws.cell(row=2, column=2).value is None
        assert ws.cell(row=3, column=1).value == "Total"
    finally:
        wb.close()


def _build_headerless_template(path: Path, *, data_rows: int) -> None:
    """A table with its header row hidden, so Excel names columns Column1/2."""
    from openpyxl.worksheet.table import Table, TableColumn

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for offset in range(data_rows):
        ws.cell(row=3 + offset, column=2, value=100 + offset)
        ws.cell(row=3 + offset, column=3, value=offset)
    ws.add_table(
        Table(
            displayName="NoHeaders",
            ref=f"B3:C{2 + data_rows}",
            tableColumns=[
                TableColumn(id=1, name="Column1"),
                TableColumn(id=2, name="Column2"),
            ],
            headerRowCount=0,
        ),
    )
    wb.save(path)
    wb.close()


def test_a_table_without_a_header_row_is_filled_by_column_name(output_path: Path) -> None:
    """A headerless table is addressed as Column1/Column2 and grows from row 1."""
    template = output_path.with_name("headerless.xlsx")
    _build_headerless_template(template, data_rows=2)

    fill(
        template=template,
        output=output_path,
        tables={
            "NoHeaders": [
                {"Column1": 10, "Column2": 60},
                {"Column1": 20, "Column2": 70},
                {"Column1": 30, "Column2": 80},
            ],
        },
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["Data"]
        table = ws.tables["NoHeaders"]

        assert table.headerRowCount == 0
        assert table.ref == "B3:C5"
        assert [ws.cell(row=r, column=2).value for r in (3, 4, 5)] == [10, 20, 30]
        assert [ws.cell(row=r, column=3).value for r in (3, 4, 5)] == [60, 70, 80]
    finally:
        wb.close()


def test_a_headerless_table_writes_the_first_row_into_the_table(output_path: Path) -> None:
    """The old first row is data, not a header, so it must not survive."""
    template = output_path.with_name("headerless_shrink.xlsx")
    _build_headerless_template(template, data_rows=3)

    fill(
        template=template,
        output=output_path,
        tables={"NoHeaders": [{"Column1": 7, "Column2": 8}]},
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["Data"]
        table = ws.tables["NoHeaders"]

        assert table.ref == "B3:C3"
        assert ws.cell(row=3, column=2).value == 7
        assert ws.cell(row=3, column=3).value == 8
        assert ws.cell(row=4, column=2).value is None
        assert ws.cell(row=5, column=2).value is None
    finally:
        wb.close()


def test_empty_data_leaves_one_blank_row_in_a_headerless_table(output_path: Path) -> None:
    """Empty data keeps a single cleared row so the table stays valid."""
    template = output_path.with_name("headerless_empty.xlsx")
    _build_headerless_template(template, data_rows=2)

    fill(template=template, output=output_path, tables={"NoHeaders": []})

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["Data"]
        table = ws.tables["NoHeaders"]

        assert table.headerRowCount == 0
        assert table.ref == "B3:C3"
        assert ws.cell(row=3, column=2).value is None
        assert ws.cell(row=3, column=3).value is None
    finally:
        wb.close()


def test_an_empty_dataframe_leaves_one_cleared_body_row(output_path: Path) -> None:
    """Empty data keeps a single empty row so the table stays valid."""
    import pandas as pd

    columns = ["Col1", "Col2", "ColNotFilled", "TestFormattedPct"]
    fill(
        template=TEMPLATE,
        output=output_path,
        tables={
            "TableTest": pd.DataFrame(columns=columns),
        },
    )

    wb = openpyxl.load_workbook(output_path)
    try:
        ws = wb["TestTable"]
        table = next(iter(ws.tables.values()))
        assert table.ref == "A1:E2"
        assert table.autoFilter.ref == "A1:E2"
        assert [ws.cell(row=2, column=col).value for col in range(1, 6)] == [
            None,
            None,
            None,
            None,
            None,
        ]
    finally:
        wb.close()


def _sales_template(path: Path) -> Path:
    """A workbook with a table called ``Sales`` on a sheet called ``Data``."""
    from tests.parity.builders import table_template

    return table_template(path, name="Sales", sheet="Data")


def test_a_column_mismatch_names_the_table_it_happened_in(tmp_path: Path) -> None:
    """The error says which table has the mismatch, not just that one exists."""
    template = _sales_template(tmp_path / "template.xlsx")

    with pytest.raises(ValueError, match=r"table 'Sales'"):
        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            engine="openpyxl",
            tables={"Sales": [{"fund": "A"}]},
            config=FillXLConfig(column_validation_strategy="strict"),
        )


def test_dropping_columns_logs_which_table_they_came_from(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """The drop strategy says which table lost which columns."""
    template = _sales_template(tmp_path / "template.xlsx")

    with caplog.at_level("INFO", logger="fillxl._validation"):
        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            engine="openpyxl",
            tables={"Sales": [{"fund": "A", "region": "EU"}]},
            config=FillXLConfig(column_validation_strategy="drop"),
        )

    assert "Table 'Sales'" in caplog.text


def test_a_header_with_a_line_break_matches_the_data_column_of_the_same_name(
    tmp_path: Path,
) -> None:
    """The file format stores ``_x000a_``; the data column is named as shown."""
    from openpyxl.utils.escape import unescape

    from tests.parity.builders import table_template

    template = table_template(
        tmp_path / "template.xlsx",
        header=("fund", "NAV_x000a_SEK"),
        header_cells=("fund", "NAV\nSEK"),
    )
    output = tmp_path / "out.xlsx"

    fill(
        template=template,
        output=output,
        engine="openpyxl",
        tables={"Sales": [{"fund": "A", "NAV\nSEK": 42}]},
        config=FillXLConfig(column_validation_strategy="strict"),
    )

    wb = openpyxl.load_workbook(output)
    try:
        ws = wb["Data"]
        table = ws.tables["Sales"]
        assert unescape(str(table.tableColumns[1].name)) == "NAV\nSEK"
        assert ws["B1"].value == "NAV\nSEK"
        assert ws["B2"].value == 42
    finally:
        wb.close()
