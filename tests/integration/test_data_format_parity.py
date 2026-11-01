"""Every supported data format feeds the same table.

``TabularAdapter`` is the only place a data format is interpreted, so the
workbook must not depend on which format the caller happened to use.  Each
case below fills the same template with the same logical rows and compares the
result with the ``list[dict]`` baseline, cell by cell and byte for byte.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Callable
from decimal import Decimal
from typing import TYPE_CHECKING, Any

import openpyxl
import pytest
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import fill

if TYPE_CHECKING:
    from pathlib import Path

COLUMNS = ["Name", "Count", "Ratio", "Active", "Day", "Note"]

ROWS: list[dict[str, Any]] = [
    {
        "Name": "Alice",
        "Count": 3,
        "Ratio": 0.5,
        "Active": True,
        "Day": dt.date(2025, 1, 2),
        "Note": "first",
    },
    {
        "Name": "Bob",
        "Count": 7,
        "Ratio": Decimal("1.25"),
        "Active": False,
        "Day": dt.date(2025, 3, 4),
        "Note": None,
    },
]

Convert = Callable[[list[dict[str, Any]]], Any]


def _list_of_dicts(rows: list[dict[str, Any]]) -> Any:  # noqa: ANN401
    return rows


def _dict_of_lists(rows: list[dict[str, Any]]) -> Any:  # noqa: ANN401
    return {column: [row[column] for row in rows] for column in COLUMNS}


def _pandas_frame(rows: list[dict[str, Any]]) -> Any:  # noqa: ANN401
    pd = pytest.importorskip("pandas")
    return pd.DataFrame(rows, columns=COLUMNS)


def _polars_frame(rows: list[dict[str, Any]]) -> Any:  # noqa: ANN401
    pl = pytest.importorskip("polars")
    return pl.DataFrame(rows)


def _sql_literal(value: Any) -> str:  # noqa: ANN401
    """Render *value* as a duckdb literal so the SQL carries the same logical data."""
    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, dt.date):
        return f"DATE '{value.isoformat()}'"
    if isinstance(value, Decimal):
        return f"CAST({value} AS DECIMAL(10, 2))"
    if isinstance(value, str):
        return "'" + value.replace("'", "''") + "'"
    return repr(value)


def _duckdb_relation(rows: list[dict[str, Any]]) -> Any:  # noqa: ANN401
    duckdb = pytest.importorskip("duckdb")
    tuples = ", ".join(
        "(" + ", ".join(_sql_literal(row[column]) for column in COLUMNS) + ")" for row in rows
    )
    return duckdb.sql(f"SELECT * FROM (VALUES {tuples}) AS t({', '.join(COLUMNS)}) ORDER BY Count")


FORMATS = [
    pytest.param(_list_of_dicts, id="list-dict"),
    pytest.param(_dict_of_lists, id="dict-list"),
    pytest.param(_pandas_frame, id="pandas"),
    pytest.param(_polars_frame, id="polars", marks=pytest.mark.polars),
    pytest.param(_duckdb_relation, id="duckdb", marks=pytest.mark.duckdb),
]


@pytest.fixture
def parity_template(tmp_path: Path) -> Path:
    """A template holding one empty table of the shared columns."""
    path = tmp_path / "parity.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for column_index, name in enumerate(COLUMNS, start=1):
        ws.cell(row=1, column=column_index, value=name)
        ws.cell(row=2, column=column_index, value="sample")
    ws.add_table(
        Table(
            displayName="Items",
            ref=f"A1:{get_column_letter(len(COLUMNS))}2",
            tableColumns=[
                TableColumn(id=index + 1, name=name) for index, name in enumerate(COLUMNS)
            ],
        ),
    )
    wb.save(path)
    wb.close()
    return path


def _filled(template: Path, output: Path, data: Any) -> Path:  # noqa: ANN401
    fill(template=template, output=output, tables={"Items": data})
    return output


def _cells(path: Path) -> list[list[Any]]:
    wb = openpyxl.load_workbook(path)
    try:
        ws = wb["Data"]
        return [
            [ws.cell(row=row, column=column).value for column in range(1, len(COLUMNS) + 1)]
            for row in range(1, len(ROWS) + 2)
        ]
    finally:
        wb.close()


@pytest.mark.parametrize("convert", FORMATS)
def test_every_format_writes_the_same_cells(
    convert: Convert,
    parity_template: Path,
    tmp_path: Path,
) -> None:
    """A format is a way of saying the same rows, so the sheet must agree."""
    baseline = _cells(_filled(parity_template, tmp_path / "baseline.xlsx", _list_of_dicts(ROWS)))
    produced = _cells(_filled(parity_template, tmp_path / "format.xlsx", convert(ROWS)))

    assert produced == baseline


@pytest.mark.parametrize("convert", FORMATS)
@pytest.mark.usefixtures("frozen_clock")
def test_every_format_produces_the_same_bytes(
    convert: Convert,
    parity_template: Path,
    tmp_path: Path,
) -> None:
    """Style, number formats and geometry must not leak the input format."""
    baseline = _filled(parity_template, tmp_path / "baseline.xlsx", _list_of_dicts(ROWS))
    produced = _filled(parity_template, tmp_path / "format.xlsx", convert(ROWS))

    assert produced.read_bytes() == baseline.read_bytes()
