"""Cross-backend equivalence.

The promise of fillxl is that the openpyxl and win32com engines produce the
same workbook for the same input.  Every test here fills one template with
both engines, reads the results back with openpyxl and compares them.

The win32com cases skip themselves when Excel is not available (the Linux CI
runners), so the same suite protects local runs and CI pipelines with Excel.
"""

from __future__ import annotations

import datetime as dt
from decimal import Decimal
from typing import TYPE_CHECKING, Any

import numpy as np
import openpyxl
import pandas as pd
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import Engine, FillXLConfig, fill

if TYPE_CHECKING:
    from pathlib import Path

COLUMNS = ["Name", "Value", "Notes"]

# One value per interesting dtype.  Both engines must agree on all of them.
VALUES = [
    pytest.param(42, id="int"),
    pytest.param(np.int64(42), id="numpy_int64"),
    pytest.param(np.float32(1.5), id="numpy_float32"),
    pytest.param(Decimal("12.34"), id="decimal"),
    pytest.param(dt.date(2025, 12, 25), id="date"),
    pytest.param(dt.time(9, 15, 30), id="time"),
    pytest.param(dt.datetime(2025, 6, 15, 14, 30), id="datetime"),
    pytest.param(
        dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.timezone(dt.timedelta(hours=2))),
        id="tz_datetime",
    ),
    pytest.param(np.datetime64("2025-06-15T14:30"), id="numpy_datetime64"),
    pytest.param(dt.timedelta(hours=6), id="timedelta"),
    pytest.param(float("nan"), id="nan"),
    pytest.param(True, id="bool"),
    pytest.param(None, id="none"),
    pytest.param("hello", id="text"),
    pytest.param("=1+1", id="formula_string"),
    pytest.param("00123", id="numeric_string"),
    pytest.param("1/2", id="slash_string"),
    pytest.param("#N/A", id="error_string"),
    pytest.param("", id="empty_string"),
]

ENGINES = [pytest.param("openpyxl", id="openpyxl"), pytest.param("win32com", id="win32com")]


@pytest.fixture(scope="module")
def equivalence_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """A template with one named cell and a one-row, three-column table."""
    from openpyxl.worksheet.filters import AutoFilter

    path = tmp_path_factory.mktemp("equivalence") / "equivalence.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for col_idx, name in enumerate(COLUMNS, 1):
        ws.cell(row=1, column=col_idx, value=name)
    ws.cell(row=2, column=1, value="sample")
    ws.cell(row=2, column=2, value=0)
    ws.cell(row=2, column=3, value="template note")
    wb.defined_names.add(DefinedName("probe_cell", attr_text="Data!$E$1"))
    ws.add_table(
        Table(
            displayName="Items",
            ref="A1:C2",
            tableColumns=[TableColumn(id=i + 1, name=name) for i, name in enumerate(COLUMNS)],
            autoFilter=AutoFilter(ref="A1:C2"),
        ),
    )
    wb.save(path)
    wb.close()
    return path


class FilledWorkbook:
    """The read-back of one filled workbook."""

    def __init__(self, worksheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        self.worksheet = worksheet
        self.named_cell = worksheet["E1"].value
        self.first_row = [worksheet.cell(row=2, column=col).value for col in range(1, 4)]
        self.table = next(iter(worksheet.tables.values()))


def _fill_and_read(
    template: Path,
    output_path: Path,
    engine: Engine,
    *,
    values: dict[str, Any] | None = None,
    tables: dict[str, Any] | None = None,
    config: FillXLConfig | None = None,
) -> FilledWorkbook:
    """Fill *template* with *engine* and return the openpyxl read-back."""
    output = output_path.with_name(f"{output_path.stem}_{engine}{output_path.suffix}")
    try:
        fill(
            template=template,
            output=output,
            values=values,
            tables=tables,
            engine=engine,
            config=config,
        )
    except Exception as exc:
        message = str(exc).lower()
        if engine == "win32com" and ("win32com" in message or "excel" in message):
            pytest.skip(f"win32com backend not available: {exc}")
        raise

    wb = openpyxl.load_workbook(output)
    try:
        return FilledWorkbook(wb["Data"])
    finally:
        wb.close()


def _table_rows(value: object) -> list[dict[str, object]]:
    """One table row carrying *value* in the Value column."""
    return [{"Name": "row", "Value": value, "Notes": "note"}]


# ═══════════════════════════════════════════════════════════════════════════════
# Same value, both engines
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("value", VALUES)
def test_both_engines_store_the_same_named_cell_value(
    equivalence_template: Path,
    output_path: Path,
    value: object,
) -> None:
    """A named cell receives the same value from both engines."""
    opened = _fill_and_read(
        equivalence_template,
        output_path,
        "openpyxl",
        values={"probe_cell": value},
    )
    closed = _fill_and_read(
        equivalence_template,
        output_path,
        "win32com",
        values={"probe_cell": value},
    )

    if isinstance(value, float) and value != value:  # NaN
        assert opened.named_cell is None
        assert closed.named_cell is None
    elif isinstance(value, Decimal):
        assert isinstance(opened.named_cell, (int, float))
        assert isinstance(closed.named_cell, (int, float))
        assert float(opened.named_cell) == pytest.approx(float(closed.named_cell))
    else:
        assert opened.named_cell == closed.named_cell


@pytest.mark.parametrize("value", VALUES)
def test_both_engines_store_the_same_table_value(
    equivalence_template: Path,
    output_path: Path,
    value: object,
) -> None:
    """A table cell receives the same value from both engines."""
    opened = _fill_and_read(
        equivalence_template,
        output_path,
        "openpyxl",
        tables={"Items": _table_rows(value)},
    )
    closed = _fill_and_read(
        equivalence_template,
        output_path,
        "win32com",
        tables={"Items": _table_rows(value)},
    )

    if isinstance(value, float) and value != value:  # NaN
        assert opened.first_row[1] is None
        assert closed.first_row[1] is None
    elif isinstance(value, Decimal):
        assert isinstance(opened.first_row[1], (int, float))
        assert isinstance(closed.first_row[1], (int, float))
        assert float(opened.first_row[1]) == pytest.approx(float(closed.first_row[1]))
    else:
        assert opened.first_row[1] == closed.first_row[1]


def test_both_engines_write_the_same_datetime_to_a_named_cell(
    equivalence_template: Path,
    output_path: Path,
) -> None:
    """A named-cell datetime is not shifted by the local timezone offset.

    win32com converts naive datetimes with the local offset, so the value has
    to be marked as UTC before it is written — otherwise the two engines store
    timestamps hours apart.
    """
    moment = dt.datetime(2025, 6, 15, 14, 30, 0)

    opened = _fill_and_read(
        equivalence_template,
        output_path,
        "openpyxl",
        values={"probe_cell": moment},
    )
    closed = _fill_and_read(
        equivalence_template,
        output_path,
        "win32com",
        values={"probe_cell": moment},
    )

    assert opened.named_cell == moment
    assert closed.named_cell == moment


@pytest.mark.parametrize("engine", ENGINES)
def test_numpy_integers_land_in_excel_as_numbers_not_text(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """np.int64 values must not be written as strings."""
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        tables={"Items": [{"Name": "a", "Value": np.int64(7), "Notes": "n"}]},
    )

    value = filled.first_row[1]
    assert value == 7
    assert not isinstance(value, str)
    assert filled.worksheet["B2"].data_type == "n"


# ═══════════════════════════════════════════════════════════════════════════════
# Tables: shape, autoFilter, empty data, missing columns
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("engine", ENGINES)
def test_a_resized_table_keeps_its_autofilter_in_step(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """table.ref and table.autoFilter.ref must stay identical.

    A mismatch is a classic cause of Excel's "we found a problem with some
    content" repair prompt.
    """
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        tables={"Items": [{"Name": f"row{i}", "Value": i, "Notes": "n"} for i in range(3)]},
    )

    assert filled.table.ref == "A1:C4"
    assert filled.table.autoFilter.ref == filled.table.ref


@pytest.mark.parametrize("engine", ENGINES)
def test_an_empty_dataframe_leaves_a_valid_one_row_table(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """An Excel table needs one body row, so empty data leaves one cleared row."""
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        tables={"Items": pd.DataFrame(columns=COLUMNS)},
    )

    assert filled.table.ref == "A1:C2"
    assert filled.table.autoFilter.ref == "A1:C2"
    assert filled.first_row == [None, None, None]


@pytest.mark.parametrize("engine", ENGINES)
def test_a_column_missing_from_the_data_does_not_leak_template_values(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """Sample values in the template are cleared for columns the data omits."""
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        tables={"Items": [{"Name": f"row{i}", "Value": i} for i in range(3)]},
        config=FillXLConfig(column_validation_strategy="warn"),
    )

    assert [filled.worksheet.cell(row=row, column=3).value for row in (2, 3, 4)] == [
        None,
        None,
        None,
    ]


# ═══════════════════════════════════════════════════════════════════════════════
# Formula escaping
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("engine", ENGINES)
def test_formula_like_text_becomes_a_formula_by_default(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """A string starting with "=" is a formula in both engines."""
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        tables={"Items": _table_rows("=1+1")},
    )

    assert filled.first_row[1] == "=1+1"
    assert filled.worksheet["B2"].data_type == "f"


@pytest.mark.parametrize("engine", ENGINES)
def test_escape_formulas_stores_literal_text_in_both_engines(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """With escape_formulas=True the "=" string stays text, not a formula."""
    config = FillXLConfig(escape_formulas=True)
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        values={"probe_cell": "=1+1"},
        tables={"Items": _table_rows("=1+1")},
        config=config,
    )

    assert filled.named_cell == "=1+1"
    assert filled.worksheet["E1"].data_type == "s"
    assert filled.first_row[1] == "=1+1"
    assert filled.worksheet["B2"].data_type == "s"


@pytest.mark.parametrize("engine", ENGINES)
def test_numeric_looking_strings_stay_text_in_both_engines(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """COM parses "00123" as typed input; the apostrophe prefix prevents it."""
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        values={"probe_cell": "00123"},
        tables={"Items": _table_rows("00123")},
    )

    assert filled.named_cell == "00123"
    assert filled.first_row[1] == "00123"
    assert filled.worksheet["B2"].data_type == "s"


@pytest.mark.parametrize("engine", ENGINES)
def test_error_looking_strings_are_stored_as_text_not_errors(
    equivalence_template: Path,
    output_path: Path,
    engine: Engine,
) -> None:
    """An error-looking string is data, not an Excel error, in both engines.

    openpyxl turns strings that match an error code into error cells on its
    own, so the cell type has to be forced to text.
    """
    filled = _fill_and_read(
        equivalence_template,
        output_path,
        engine,
        values={"probe_cell": "#N/A"},
        tables={"Items": _table_rows("#N/A")},
    )

    assert filled.named_cell == "#N/A"
    assert filled.worksheet["E1"].data_type == "s"
    assert filled.first_row[1] == "#N/A"
    assert filled.worksheet["B2"].data_type == "s"
