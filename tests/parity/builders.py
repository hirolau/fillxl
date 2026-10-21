"""Workbook builders for the parity tests."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.data_source import NumData
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo


def table_template(
    path: Path,
    *,
    header: tuple[str, ...] = ("fund", "nav"),
    header_cells: tuple[Any, ...] | None = None,
    body: tuple[tuple[Any, ...], ...] = (("A", 1),),
    sheet: str = "Data",
    name: str = "Sales",
    top_row: int = 1,
    totals: bool = False,
    headerless: bool = False,
    number_formats: dict[int, str] | None = None,
    other_sheets: tuple[str, ...] = (),
    active_sheet: str | None = None,
) -> Path:
    """A workbook with one table. *header* holds the table column names as stored in the file;
    *header_cells* (default: *header*) holds what the header cells display."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet
    for other in other_sheets:
        workbook.create_sheet(other)

    row = top_row
    if not headerless:
        for col, title in enumerate(header_cells or header, start=1):
            worksheet.cell(row=row, column=col, value=title)
        row += 1
    for values in body:
        for col, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row, column=col, value=value)
            if number_formats and col in number_formats:
                cell.number_format = number_formats[col]
        row += 1
    body_last_row = row - 1

    columns = [TableColumn(id=i, name=title) for i, title in enumerate(header, start=1)]
    last_row = body_last_row
    last_col = get_column_letter(len(header))
    if totals:
        last_row += 1
        worksheet.cell(row=last_row, column=1, value="Total")
        worksheet.cell(
            row=last_row, column=len(header), value=f"=SUBTOTAL(109,{name}[{header[-1]}])"
        )
        columns[0].totalsRowLabel = "Total"
        columns[-1].totalsRowFunction = "sum"

    table = Table(
        displayName=name,
        ref=f"A{top_row}:{last_col}{last_row}",
        tableColumns=columns,
        headerRowCount=0 if headerless else 1,
        totalsRowCount=1 if totals else None,
        tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True),
    )
    if not headerless:
        table.autoFilter = AutoFilter(ref=f"A{top_row}:{last_col}{body_last_row}")
    worksheet.add_table(table)

    if active_sheet is not None:
        index = workbook.sheetnames.index(active_sheet)
        workbook.active = index
        for position, sheet_name in enumerate(workbook.sheetnames):
            workbook[sheet_name].sheet_view.tabSelected = position == index
    workbook.save(path)
    return path


def named_cells_template(
    path: Path,
    *,
    cells: dict[str, tuple[str, Any]],
    sheet: str = "Data",
) -> Path:
    """A workbook whose *cells* map a defined name to ``(coordinate, value)``."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet
    for name, (coordinate, value) in cells.items():
        worksheet[coordinate] = value
        workbook.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!{coordinate}"))
    workbook.save(path)
    return path


def chart_template(
    path: Path,
    *,
    sheet: str = "Data",
    header: tuple[str, ...] = ("fund", "nav"),
    body: tuple[tuple[Any, ...], ...] = (("A", 1),),
    reference: str = "Data!$B$2",
    chart_sheet: bool = False,
    cache_points: int | None = None,
) -> Path:
    """A workbook with one table and a bar chart whose values point at *reference*."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet
    for col, title in enumerate(header, start=1):
        worksheet.cell(row=1, column=col, value=title)
    for row, values in enumerate(body, start=2):
        for col, value in enumerate(values, start=1):
            worksheet.cell(row=row, column=col, value=value)
    last_row = 1 + len(body)
    last_col = get_column_letter(len(header))

    table = Table(
        displayName="Sales",
        ref=f"A1:{last_col}{last_row}",
        tableColumns=[TableColumn(id=i, name=title) for i, title in enumerate(header, start=1)],
        tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True),
    )
    table.autoFilter = AutoFilter(ref=f"A1:{last_col}{last_row}")
    worksheet.add_table(table)

    chart = BarChart()
    chart.title = "Chart"
    chart.add_data(
        Reference(worksheet, min_col=2, min_row=1, max_row=last_row), titles_from_data=True
    )
    chart.series[0].val.numRef.f = reference
    if cache_points is not None:
        chart.series[0].val.numRef.numCache = NumData(ptCount=cache_points)
    if chart_sheet:
        workbook.create_chartsheet("Chart Sheet").add_chart(chart)
    else:
        workbook.create_sheet("Charts").add_chart(chart, "A1")
    workbook.save(path)
    return path


def sheets_template(
    path: Path,
    *sheets: str,
    active: str | None = None,
    very_hidden: tuple[str, ...] = (),
) -> Path:
    """A workbook with *sheets*; the first one is active unless *active* says otherwise."""
    workbook = Workbook()
    workbook.active.title = sheets[0]
    for name in sheets[1:]:
        workbook.create_sheet(name)
    active_title = active or sheets[0]
    workbook.active = workbook.sheetnames.index(active_title)
    for position, name in enumerate(workbook.sheetnames):
        workbook[name].sheet_view.tabSelected = position == workbook.sheetnames.index(active_title)
    for name in very_hidden:
        workbook[name].sheet_state = "veryHidden"
    workbook.save(path)
    return path


def chart_sheet_first_template(path: Path) -> Path:
    """A workbook whose first sheet is a chart sheet, followed by a worksheet."""
    workbook = Workbook()
    workbook.active.title = "Report"
    chart_sheet = workbook.create_chartsheet("Chart")
    workbook.move_sheet("Chart", offset=-1)
    chart = BarChart()
    chart.title = "Chart"
    chart_sheet.add_chart(chart)
    workbook.save(path)
    return path
