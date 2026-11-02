"""Templates the OOXML hardening tests start from.

Every builder returns the path it wrote, so a test can chain a build with a
:func:`tests.ooxml._patch.patch_part` call to reach a shape openpyxl cannot
write.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import (
    Table,
    TableColumn,
    TableFormula,
    TableStyleInfo,
)

if TYPE_CHECKING:
    from collections.abc import Mapping, Sequence
    from pathlib import Path

    from openpyxl.worksheet.worksheet import Worksheet

HEADERS = ("Region", "Amount", "Tax", "Total", "Note")


def _placeholder(column: str, index: int) -> Any:  # noqa: ANN401
    """A sample value for *column*, so the template looks like a real one."""
    if column == "Region":
        return ("North", "South", "East", "West")[index % 4]
    if column in ("Tax", "Total"):
        return 10.0 + index
    if column == "Note":
        return f"note {index}"
    return index + 1


def _build_table(
    workbook: Workbook,
    *,
    rows: int,
    header: Sequence[str],
    sheet: str,
    table: str,
    top_row: int,
    totals: bool,
    calculated: str | None,
    number_formats: Mapping[str, str],
    headerless: bool = False,
) -> tuple[Worksheet, int]:
    """Write the table into a fresh sheet; returns the sheet and its last row."""
    worksheet = workbook.active
    worksheet.title = sheet
    if not headerless:
        for offset, name in enumerate(header, start=1):
            worksheet.cell(row=top_row, column=offset, value=name)
    first_body_row = top_row if headerless else top_row + 1
    for index in range(rows):
        row = first_body_row + index
        for offset, name in enumerate(header, start=1):
            cell = worksheet.cell(row=row, column=offset)
            cell.value = f"=[@{header[1]}]*2" if name == calculated else _placeholder(name, index)
            if name in number_formats:
                cell.number_format = number_formats[name]

    body_last_row = first_body_row + rows - 1
    last_row = body_last_row
    columns = [TableColumn(id=index, name=name) for index, name in enumerate(header, start=1)]
    if totals:
        last_row += 1
        worksheet.cell(row=last_row, column=1, value="Total")
        columns[0].totalsRowLabel = "Total"
        for offset, name in enumerate(header[1:], start=2):
            if name == calculated:
                continue
            worksheet.cell(row=last_row, column=offset, value=f"=SUBTOTAL(109,{table}[{name}])")
            columns[offset - 1].totalsRowFunction = "sum"
    if calculated is not None:
        columns[list(header).index(calculated)].calculatedColumnFormula = TableFormula(
            attr_text=f"[@{header[1]}]*2",
        )

    definition = Table(
        displayName=table,
        ref=f"A{top_row}:{get_column_letter(len(header))}{last_row}",
        tableColumns=columns,
        headerRowCount=0 if headerless else None,
        totalsRowCount=1 if totals else None,
        tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True),
    )
    definition.autoFilter = None
    worksheet.add_table(definition)
    return worksheet, last_row


def table_template(
    path: Path,
    *,
    rows: int = 1,
    header: Sequence[str] = HEADERS,
    sheet: str = "Data",
    table: str = "Data",
    top_row: int = 1,
    totals: bool = False,
    calculated: str | None = None,
    sum_below: bool = False,
    sum_sheet: str | None = None,
    number_formats: Mapping[str, str] | None = None,
    headerless: bool = False,
) -> Path:
    """A workbook with one table of *rows* body rows.

    Args:
        path: Output path.
        rows: Number of body rows.
        header: Table column names.
        sheet: Name of the sheet holding the table.
        table: Table name.
        top_row: First row of the table.
        totals: Whether the table has a totals row.
        calculated: Name of a calculated column (its cells hold a formula).
        sum_below: Add ``=SUM(...)`` below the table on the same sheet.
        sum_sheet: Add a second sheet whose ``=SUM(...)`` refers to the table.
        number_formats: Per-column number format, by column name.
        headerless: Table without a header row.

    Returns:
        *path*.

    """
    workbook = Workbook()
    worksheet, last_row = _build_table(
        workbook,
        rows=rows,
        header=header,
        sheet=sheet,
        table=table,
        top_row=top_row,
        totals=totals,
        calculated=calculated,
        number_formats=dict(number_formats or {}),
        headerless=headerless,
    )
    first_body_row = top_row if headerless else top_row + 1
    body_last_row = first_body_row + rows - 1
    if sum_below:
        # Outside the table's columns: a summary directly under the body would
        # be in the growth path, and both engines refuse to overwrite it.
        worksheet.cell(
            row=last_row,
            column=len(header) + 2,
            value=f"=SUM(B{first_body_row}:B{body_last_row})",
        )
    if sum_sheet is not None:
        workbook.create_sheet(sum_sheet)["A1"] = (
            f"=SUM({sheet}!B{first_body_row}:B{body_last_row})"
        )
    workbook.save(path)
    workbook.close()
    return path


def table_with_totals_template(
    path: Path,
    *,
    rows: int = 1,
    **kwargs: Any,  # noqa: ANN401
) -> Path:
    """The table template with a totals row."""
    return table_template(path, rows=rows, totals=True, **kwargs)


def table_with_calculated_column_template(
    path: Path,
    *,
    rows: int = 1,
    **kwargs: Any,  # noqa: ANN401
) -> Path:
    """The table template with a calculated column."""
    return table_template(path, rows=rows, calculated="Total", **kwargs)


def table_with_sum_below_template(
    path: Path,
    *,
    rows: int = 1,
    other_sheet: str | None = None,
    **kwargs: Any,  # noqa: ANN401
) -> Path:
    """The table template with a ``=SUM(...)`` that tracks the body."""
    return table_template(
        path,
        rows=rows,
        sum_below=other_sheet is None,
        sum_sheet=other_sheet,
        **kwargs,
    )


def named_cell_template(
    path: Path,
    *,
    cells: Mapping[str, tuple[str, Any]] | None = None,
    sheet: str = "Data",
    other_sheets: Sequence[str] = (),
) -> Path:
    """A workbook whose defined names point at single cells of *sheet*.

    Args:
        path: Output path.
        cells: ``name -> (coordinate, value)``; defaults to one number.
        sheet: Sheet holding the cells.
        other_sheets: Extra empty sheets.

    Returns:
        *path*.

    """
    values = dict(cells or {"report_title": ("B2", 1)})
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet
    for name in other_sheets:
        workbook.create_sheet(name)
    for name, (coordinate, value) in values.items():
        worksheet[coordinate] = value
        workbook.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!{coordinate}"))
    workbook.save(path)
    workbook.close()
    return path


def underscore_sheet_template(path: Path, *, rows: int = 1) -> Path:
    """A table on ``Data`` plus a ``_Helper`` sheet a fill would hide."""
    workbook = Workbook()
    _build_table(
        workbook,
        rows=rows,
        header=HEADERS,
        sheet="Data",
        table="Data",
        top_row=1,
        totals=False,
        calculated=None,
        number_formats={},
    )
    workbook.create_sheet("_Helper")["A1"] = "helper"
    workbook.save(path)
    workbook.close()
    return path


def chart_template(path: Path, *, rows: int = 1, header: Sequence[str] = HEADERS) -> Path:
    """A table with a bar chart over its second column, on a ``Charts`` sheet."""
    workbook = Workbook()
    worksheet, last_row = _build_table(
        workbook,
        rows=rows,
        header=header,
        sheet="Data",
        table="Data",
        top_row=1,
        totals=False,
        calculated=None,
        number_formats={},
    )
    chart = BarChart()
    chart.add_data(
        Reference(worksheet, min_col=2, min_row=1, max_row=last_row),
        titles_from_data=True,
    )
    workbook.create_sheet("Charts").add_chart(chart, "A1")
    workbook.save(path)
    workbook.close()
    return path


def body_rows(count: int, *, first_day: dt.date | None = None) -> list[dict[str, Any]]:
    """*count* rows of fill data shaped like :data:`HEADERS`."""
    start = first_day or dt.date(2024, 1, 1)
    return [
        {
            "Region": f"R{index}",
            "Amount": index + 1,
            "Tax": start + dt.timedelta(days=index),
            "Total": (index + 1) * 2,
            "Note": f"n{index}",
        }
        for index in range(count)
    ]
