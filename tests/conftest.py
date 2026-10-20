"""Shared fixtures for fillxl tests."""

from __future__ import annotations

import os
import time
import zipfile
from collections.abc import Generator
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest
import xlwings as xw
from freezegun import freeze_time
from openpyxl.chart import (
    AreaChart,
    BarChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
)
from openpyxl.worksheet.table import (
    Table,
    TableColumn,
    TableFormula,
)

TEMPLATE = Path(__file__).parent / "assets" / "test_template_1.xlsx"
CHARTS_TEMPLATE_2 = Path(__file__).parent / "assets" / "test_template_2_charts.xlsx"
PIVOT_TEMPLATE = Path(__file__).parent / "assets" / "test_template_3_pivot.xlsx"
CHARTS_TEMPLATE_4 = Path(__file__).parent / "assets" / "test_template_4_charts_edge_cases.xlsx"


# ── Helpers ─────────────────────────────────────────────────────────────────


def read_output(path: Path) -> openpyxl.Workbook:
    """Open an output ``.xlsx`` with openpyxl for verification.

    Args:
        path: Path to the output file.

    Returns:
        An openpyxl :class:`~openpyxl.Workbook` ready for inspection.

    """
    return openpyxl.load_workbook(path)


# ── Standard fixtures ───────────────────────────────────────────────────────


@pytest.fixture
def output_path(tmp_path: Path) -> Path:
    """Return a temporary output path for each test."""
    return tmp_path / "output.xlsx"


@pytest.fixture
def frozen_clock(monkeypatch: pytest.MonkeyPatch) -> Generator[datetime, None, None]:
    """Freeze the clock *and* pin every zip-entry timestamp for one test.

    Writing a workbook records a timestamp in every zip entry, which makes
    "the same fill twice produces the same bytes" flaky: zip timestamps have
    two-second resolution, so two fills straddling a boundary differ.

    ``freezegun`` covers ``ZipFile.writestr``, which stamps entries from
    ``time.localtime(time.time())``.  It does *not* cover ``ZipFile.write``,
    which openpyxl uses for worksheet XML: that path stamps the entry from the
    modification time of a temporary file (``ZipInfo.from_file``), so the real
    clock leaks in through a second door.  Both sources are pinned here.

    Yields:
        The frozen wall-clock time.  The instant is deliberately far in the
        past, so a real-clock leak cannot coincidentally match it.

    """
    moment = datetime(2001, 2, 3, 4, 5, 6)
    with freeze_time(moment):
        # Whatever the machine's UTC offset, `writestr` will stamp entries with
        # the frozen instant rendered in local time — so pin `from_file` to the
        # very same value and every entry in the archive agrees.
        local = time.localtime(time.time())
        stamp = (
            local.tm_year,
            local.tm_mon,
            local.tm_mday,
            local.tm_hour,
            local.tm_min,
            local.tm_sec,
        )
        original_from_file = zipfile.ZipInfo.__dict__["from_file"]

        def frozen_from_file(
            filename: str | os.PathLike[str],
            arcname: str | None = None,
            strict_timestamps: bool = True,
        ) -> zipfile.ZipInfo:
            """Like ``ZipInfo.from_file``, but stamped from the frozen clock."""
            zinfo: zipfile.ZipInfo = original_from_file.__func__(
                zipfile.ZipInfo,
                filename,
                arcname,
                strict_timestamps=strict_timestamps,
            )
            zinfo.date_time = stamp
            return zinfo

        monkeypatch.setattr(zipfile.ZipInfo, "from_file", frozen_from_file)
        yield moment


@pytest.fixture
def xl_app() -> Generator[xw.App, None, None]:
    """Launch a hidden Excel instance for COM-based verification, quit after test."""
    app = xw.App(visible=False)
    yield app
    app.quit()


# ── Programmatic workbook fixtures ──────────────────────────────────────────


def _build_charts_workbook(path: Path) -> None:
    """Build a workbook with BarChart, LineChart, PieChart, ScatterChart,
    and AreaChart all bound to the TableTest data range.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestTable"

    # Headers
    headers = ["Col1", "Col2", "ColNotFilled", "ColFormula", "TestFormattedPct"]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)

    # Data rows (4 rows)
    data = [
        ["a", 1, "x", None, 0.0],
        ["b", 2, "y", None, 0.0],
        ["c", 7, "z", None, 0.0],
        ["d", 4, "w", None, 0.0],
    ]
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # Table
    table = Table(
        displayName="TableTest",
        ref="A1:E5",
        tableColumns=[
            TableColumn(id=1, name="Col1"),
            TableColumn(id=2, name="Col2"),
            TableColumn(id=3, name="ColNotFilled"),
            TableColumn(id=4, name="ColFormula"),
            TableColumn(id=5, name="TestFormattedPct"),
        ],
    )
    ws.add_table(table)
    ws.cell(row=40, column=1, value="TEST NOT OVERWRITE")

    chart_ws = wb.create_sheet("Charts")

    # BarChart (ColumnChart) — val = Col2, cat = Col1
    bar_chart = BarChart()
    bar_chart.title = "Bar Chart"
    data_ref = Reference(ws, min_col=2, min_row=1, max_row=5)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=5)
    bar_chart.add_data(data_ref, titles_from_data=True)
    bar_chart.set_categories(cats_ref)
    chart_ws.add_chart(bar_chart, "A1")

    # LineChart — val = Col2
    line_chart = LineChart()
    line_chart.title = "Line Chart"
    data_ref = Reference(ws, min_col=2, min_row=2, max_row=5)
    line_chart.add_data(data_ref)
    chart_ws.add_chart(line_chart, "A20")

    # PieChart — val = Col2, cat = Col1
    pie_chart = PieChart()
    pie_chart.title = "Pie Chart"
    data_ref = Reference(ws, min_col=2, min_row=2, max_row=5)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=5)
    pie_chart.add_data(data_ref)
    pie_chart.set_categories(cats_ref)
    chart_ws.add_chart(pie_chart, "A40")

    # ScatterChart — xVal = Col2, yVal = Col1
    scatter_chart = ScatterChart()
    scatter_chart.title = "Scatter Chart"
    x_ref = Reference(ws, min_col=2, min_row=2, max_row=5)
    y_ref = Reference(ws, min_col=1, min_row=2, max_row=5)
    scatter_chart.add_data(y_ref)
    scatter_chart.ser[0].xValues = x_ref
    chart_ws.add_chart(scatter_chart, "A60")

    # AreaChart — val = Col2
    area_chart = AreaChart()
    area_chart.title = "Area Chart"
    data_ref = Reference(ws, min_col=2, min_row=2, max_row=5)
    area_chart.add_data(data_ref)
    chart_ws.add_chart(area_chart, "A80")

    wb.save(path)
    wb.close()


def _build_multi_table_workbook(path: Path) -> None:
    """Build a workbook with two tables (TableTest + TableTest2)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestTables"

    # ── Table 1: TableTest ──
    headers = ["Col1", "Col2", "ColNotFilled", "ColFormula", "TestFormattedPct"]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)

    data = [["a", 1, "x", None, 0.0], ["b", 2, "y", None, 0.0]]
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    table1 = Table(
        displayName="TableTest",
        ref="A1:E3",
        tableColumns=[
            TableColumn(id=1, name="Col1"),
            TableColumn(id=2, name="Col2"),
            TableColumn(id=3, name="ColNotFilled"),
            TableColumn(id=4, name="ColFormula"),
            TableColumn(id=5, name="TestFormattedPct"),
        ],
    )
    ws.add_table(table1)

    # ── Table 2: TableTest2 ──
    headers2 = ["Name", "Value"]
    for col_idx, header in enumerate(headers2, 7):
        ws.cell(row=1, column=col_idx, value=header)

    data2 = [["Alice", 100], ["Bob", 200]]
    for row_idx, row_data in enumerate(data2, 2):
        for col_idx, value in enumerate(row_data, 7):
            ws.cell(row=row_idx, column=col_idx, value=value)

    table2 = Table(
        displayName="TableTest2",
        ref="G1:H3",
        tableColumns=[
            TableColumn(id=1, name="Name"),
            TableColumn(id=2, name="Value"),
        ],
    )
    ws.add_table(table2)

    wb.save(path)
    wb.close()


def _build_hidden_sheets_workbook(path: Path) -> None:
    """Build a workbook with ``_``-prefixed sheets for hide tests."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestSingleCell"
    ws["A1"] = "Single Cell Before"

    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names.add(DefinedName("single_cell", attr_text="TestSingleCell!$A$1"))

    wb.create_sheet("_HiddenSheet").cell(row=1, column=1, value="Should be hidden")
    wb.create_sheet("NormalSheet").cell(row=1, column=1, value="Should be visible")

    ws_table = wb.create_sheet("TestTable")
    for col_idx, header in enumerate(["Col1", "Col2"], 1):
        ws_table.cell(row=1, column=col_idx, value=header)
    ws_table.cell(row=2, column=1, value="a")
    ws_table.cell(row=2, column=2, value=1)

    ws_table.add_table(
        Table(
            displayName="TableTest",
            ref="A1:B2",
            tableColumns=[
                TableColumn(id=1, name="Col1"),
                TableColumn(id=2, name="Col2"),
            ],
        ),
    )

    wb.save(path)
    wb.close()


def _build_delete_sheets_workbook(path: Path) -> None:
    """Build a workbook with a deletable sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestSheet"
    ws["A1"] = "Keep me"

    wb.create_sheet("ToDelete").cell(row=1, column=1, value="Delete me")

    ws_table = wb.create_sheet("TestTable")
    for col_idx, header in enumerate(["Col1", "Col2"], 1):
        ws_table.cell(row=1, column=col_idx, value=header)
    ws_table.cell(row=2, column=1, value="a")
    ws_table.cell(row=2, column=2, value=1)

    ws_table.add_table(
        Table(
            displayName="TableTest",
            ref="A1:B2",
            tableColumns=[
                TableColumn(id=1, name="Col1"),
                TableColumn(id=2, name="Col2"),
            ],
        ),
    )

    wb.save(path)
    wb.close()


def _build_multi_named_cells_workbook(path: Path) -> None:
    """Build a workbook with several regular named cells."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestSingleCell"
    ws["A1"] = "Initial"
    ws["B1"] = "Date cell"
    ws["C1"] = "Bool cell"
    ws["D1"] = "Pct cell"
    ws["D1"].number_format = "0.00%"

    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names.add(DefinedName("single_cell", attr_text="TestSingleCell!$A$1"))
    wb.defined_names.add(DefinedName("date_cell", attr_text="TestSingleCell!$B$1"))
    wb.defined_names.add(DefinedName("bool_cell", attr_text="TestSingleCell!$C$1"))
    wb.defined_names.add(DefinedName("pct_cell", attr_text="TestSingleCell!$D$1"))

    wb.save(path)
    wb.close()


def _build_multi_dest_name_workbook(path: Path) -> None:
    """Build a workbook with a named cell pointing to multiple destinations."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["A1"] = "First dest"

    ws2 = wb.create_sheet("Sheet2")
    ws2["B2"] = "Second dest"

    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names.add(DefinedName("multi_dest", attr_text="Sheet1!$A$1,Sheet2!$B$2"))
    wb.defined_names.add(DefinedName("single_cell", attr_text="Sheet1!$A$1"))

    wb.save(path)
    wb.close()


def _build_calculated_column_workbook(path: Path) -> None:
    """Build a workbook with a table that has a calculated column formula."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestTable"

    headers = ["Col1", "Col2", "ColNotFilled", "ColFormula", "TestFormattedPct"]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)

    ws.cell(row=2, column=1, value="a")
    ws.cell(row=2, column=2, value=1)
    ws.cell(row=2, column=3, value="x")
    ws.cell(row=2, column=5, value=0.0)

    # Column D: calculated column formula = Col1 & Col2 (or similar)
    col_formula = TableFormula(attr_text='Col1&" "&Col2')

    table = Table(
        displayName="TableTest",
        ref="A1:E2",
        tableColumns=[
            TableColumn(id=1, name="Col1"),
            TableColumn(id=2, name="Col2"),
            TableColumn(id=3, name="ColNotFilled"),
            TableColumn(id=4, name="ColFormula", calculatedColumnFormula=col_formula),
            TableColumn(id=5, name="TestFormattedPct"),
        ],
    )
    ws.add_table(table)
    ws.cell(row=40, column=1, value="TEST NOT OVERWRITE")

    wb.save(path)
    wb.close()


def _build_chart_sheet_workbook(path: Path) -> None:
    """Build a workbook with a chart sheet (non-regular worksheet)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DataSheet"
    ws["A1"] = "Value"
    ws["A2"] = 42

    wb.create_chartsheet(title="ChartSheet")

    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names.add(DefinedName("chart_sheet_name", attr_text="ChartSheet!$A$1"))

    wb.save(path)
    wb.close()


@pytest.fixture(scope="session")
def charts_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with BarChart, LineChart, PieChart, ScatterChart, AreaChart.

    Created once per test session and cached.
    """
    path = tmp_path_factory.mktemp("charts") / "charts_template.xlsx"
    _build_charts_workbook(path)
    return path


@pytest.fixture(scope="session")
def multi_table_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with two tables (TableTest + TableTest2)."""
    path = tmp_path_factory.mktemp("multi_table") / "multi_table_template.xlsx"
    _build_multi_table_workbook(path)
    return path


@pytest.fixture(scope="session")
def hidden_sheets_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with ``_HiddenSheet``, ``NormalSheet``, and a table."""
    path = tmp_path_factory.mktemp("hidden_sheets") / "hidden_sheets_template.xlsx"
    _build_hidden_sheets_workbook(path)
    return path


@pytest.fixture(scope="session")
def delete_sheets_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with a ``ToDelete`` sheet plus a normal sheet and table."""
    path = tmp_path_factory.mktemp("delete_sheets") / "delete_sheets_template.xlsx"
    _build_delete_sheets_workbook(path)
    return path


@pytest.fixture(scope="session")
def multi_named_cells_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with several regular named cells (text, date, bool, pct)."""
    path = tmp_path_factory.mktemp("named_cells") / "multi_named_cells_template.xlsx"
    _build_multi_named_cells_workbook(path)
    return path


@pytest.fixture(scope="session")
def multi_dest_name_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with a named cell that has multiple destination ranges."""
    path = tmp_path_factory.mktemp("multi_dest") / "multi_dest_name_template.xlsx"
    _build_multi_dest_name_workbook(path)
    return path


@pytest.fixture(scope="session")
def calculated_column_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with a table that has a calculated column formula in ColFormula."""
    path = tmp_path_factory.mktemp("calc_col") / "calculated_column_template.xlsx"
    _build_calculated_column_workbook(path)
    return path


@pytest.fixture(scope="session")
def chart_sheet_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with a chart sheet (non-regular worksheet) for skip tests."""
    path = tmp_path_factory.mktemp("chart_sheet") / "chart_sheet_template.xlsx"
    _build_chart_sheet_workbook(path)
    return path


# ── Dtype test template ──────────────────────────────────────────────────────


def _build_dtype_template(path: Path) -> None:
    """Build a workbook with named cells and a table covering common Python/NumPy datatypes.

    Named cell sheet: str_cell, int_cell, float_cell, bool_cell, none_cell,
                      datetime_cell, date_cell, time_cell, decimal_cell
    Table sheet:      StrCol, IntCol, FloatCol, BoolCol, NoneCol, DatetimeCol,
                      DateCol, TimeCol, DecimalCol
    """
    from datetime import date, datetime, time
    from decimal import Decimal

    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()

    # ── Sheet 1: Named cells ──────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "DtypeNamedCells"
    ws1["A1"] = "str placeholder"
    ws1["B1"] = "int placeholder"
    ws1["C1"] = "float placeholder"
    ws1["D1"] = "bool placeholder"
    ws1["E1"] = "none placeholder"
    ws1["F1"] = "datetime placeholder"
    ws1["G1"] = "date placeholder"
    ws1["H1"] = "time placeholder"
    ws1["H1"].number_format = "hh:mm:ss"
    ws1["I1"] = "decimal placeholder"

    wb.defined_names.add(DefinedName("str_cell", attr_text="DtypeNamedCells!$A$1"))
    wb.defined_names.add(DefinedName("int_cell", attr_text="DtypeNamedCells!$B$1"))
    wb.defined_names.add(DefinedName("float_cell", attr_text="DtypeNamedCells!$C$1"))
    wb.defined_names.add(DefinedName("bool_cell", attr_text="DtypeNamedCells!$D$1"))
    wb.defined_names.add(DefinedName("none_cell", attr_text="DtypeNamedCells!$E$1"))
    wb.defined_names.add(DefinedName("datetime_cell", attr_text="DtypeNamedCells!$F$1"))
    wb.defined_names.add(DefinedName("date_cell", attr_text="DtypeNamedCells!$G$1"))
    wb.defined_names.add(DefinedName("time_cell", attr_text="DtypeNamedCells!$H$1"))
    wb.defined_names.add(DefinedName("decimal_cell", attr_text="DtypeNamedCells!$I$1"))

    # ── Sheet 2: Table ────────────────────────────────────────────────────
    ws2 = wb.create_sheet("DtypeTable")
    headers = [
        "StrCol",
        "IntCol",
        "FloatCol",
        "BoolCol",
        "NoneCol",
        "DatetimeCol",
        "DateCol",
        "TimeCol",
        "DecimalCol",
    ]
    for col_idx, header in enumerate(headers, 1):
        ws2.cell(row=1, column=col_idx, value=header)

    # One row of placeholder data so the table exists
    ws2.cell(row=2, column=1, value="placeholder")
    ws2.cell(row=2, column=2, value=0)
    ws2.cell(row=2, column=3, value=0.0)
    ws2.cell(row=2, column=4, value=False)
    ws2.cell(row=2, column=5, value=None)
    ws2.cell(row=2, column=6, value=datetime(2025, 1, 1))
    ws2.cell(row=2, column=7, value=date(2025, 1, 1))
    ws2.cell(row=2, column=8, value=time(12, 0)).number_format = "hh:mm:ss"
    ws2.cell(row=2, column=9, value=Decimal(0))

    table = openpyxl.worksheet.table.Table(
        displayName="DtypeTable",
        ref="A1:I2",
        tableColumns=[
            openpyxl.worksheet.table.TableColumn(id=1, name="StrCol"),
            openpyxl.worksheet.table.TableColumn(id=2, name="IntCol"),
            openpyxl.worksheet.table.TableColumn(id=3, name="FloatCol"),
            openpyxl.worksheet.table.TableColumn(id=4, name="BoolCol"),
            openpyxl.worksheet.table.TableColumn(id=5, name="NoneCol"),
            openpyxl.worksheet.table.TableColumn(id=6, name="DatetimeCol"),
            openpyxl.worksheet.table.TableColumn(id=7, name="DateCol"),
            openpyxl.worksheet.table.TableColumn(id=8, name="TimeCol"),
            openpyxl.worksheet.table.TableColumn(id=9, name="DecimalCol"),
        ],
    )
    ws2.add_table(table)

    wb.save(path)
    wb.close()


@pytest.fixture(scope="session")
def dtype_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with named cells and a table covering common datatypes."""
    path = tmp_path_factory.mktemp("dtype") / "dtype_template.xlsx"
    _build_dtype_template(path)
    return path


# ── Real-file templates (from attachments) ──────────────────────────────────


@pytest.fixture(scope="session")
def charts_template_2() -> Path:
    """Workbook with 10 chart sheets, each with a table bound to a chart.

    Comes from ``tests/assets/test_template_2_charts.xlsx``.
    """
    return CHARTS_TEMPLATE_2


@pytest.fixture(scope="session")
def pivot_template() -> Path:
    """Workbook with a Sales Data table and a Pivot Analysis pivot table.

    Comes from ``tests/assets/test_template_3_pivot.xlsx``.
    """
    return PIVOT_TEMPLATE


# ── Cross-backend dtype test template ──────────────────────────────────────


def _build_cross_backend_dtype_template(path: Path) -> None:
    """Build a workbook with named cells and a table covering ALL common dtypes
    for cross-backend comparison testing.

    Extends the basic dtype template with additional columns for edge cases:
    NaN, pd.NA, pd.NaT, pd.Timestamp, timezone-aware datetime, large numbers,
    and numpy-specific types.

    Named cell sheet adds: nan_cell, na_cell, nat_cell, timestamp_cell,
                           tz_datetime_cell, large_int_cell, large_float_cell
    Table sheet adds:      FloatNanCol, NaCol, NaTCol, TimestampCol,
                           TzDatetimeCol, LargeIntCol, LargeFloatCol
    """
    from datetime import date, datetime, time
    from decimal import Decimal

    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()

    # ── Sheet 1: Named cells ──────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "CrossDtypeNamedCells"
    ws1["A1"] = "str placeholder"
    ws1["B1"] = "int placeholder"
    ws1["C1"] = "float placeholder"
    ws1["D1"] = "bool placeholder"
    ws1["E1"] = "none placeholder"
    ws1["F1"] = "datetime placeholder"
    ws1["G1"] = "date placeholder"
    ws1["H1"] = "time placeholder"
    ws1["H1"].number_format = "hh:mm:ss"
    ws1["I1"] = "decimal placeholder"
    ws1["J1"] = "nan placeholder"
    ws1["K1"] = "timestamp placeholder"
    ws1["L1"] = "tz_datetime placeholder"
    ws1["M1"] = "large_int placeholder"
    ws1["N1"] = "large_float placeholder"

    # Basic types
    wb.defined_names.add(DefinedName("cd_str_cell", attr_text="CrossDtypeNamedCells!$A$1"))
    wb.defined_names.add(DefinedName("cd_int_cell", attr_text="CrossDtypeNamedCells!$B$1"))
    wb.defined_names.add(DefinedName("cd_float_cell", attr_text="CrossDtypeNamedCells!$C$1"))
    wb.defined_names.add(DefinedName("cd_bool_cell", attr_text="CrossDtypeNamedCells!$D$1"))
    wb.defined_names.add(DefinedName("cd_none_cell", attr_text="CrossDtypeNamedCells!$E$1"))
    wb.defined_names.add(DefinedName("cd_datetime_cell", attr_text="CrossDtypeNamedCells!$F$1"))
    wb.defined_names.add(DefinedName("cd_date_cell", attr_text="CrossDtypeNamedCells!$G$1"))
    wb.defined_names.add(DefinedName("cd_time_cell", attr_text="CrossDtypeNamedCells!$H$1"))
    wb.defined_names.add(DefinedName("cd_decimal_cell", attr_text="CrossDtypeNamedCells!$I$1"))
    # Edge-case types
    wb.defined_names.add(DefinedName("cd_nan_cell", attr_text="CrossDtypeNamedCells!$J$1"))
    wb.defined_names.add(DefinedName("cd_timestamp_cell", attr_text="CrossDtypeNamedCells!$K$1"))
    wb.defined_names.add(DefinedName("cd_tz_datetime_cell", attr_text="CrossDtypeNamedCells!$L$1"))
    wb.defined_names.add(DefinedName("cd_large_int_cell", attr_text="CrossDtypeNamedCells!$M$1"))
    wb.defined_names.add(DefinedName("cd_large_float_cell", attr_text="CrossDtypeNamedCells!$N$1"))

    # ── Sheet 2: Table ────────────────────────────────────────────────────
    ws2 = wb.create_sheet("CrossDtypeTable")
    headers = [
        "StrCol",
        "IntCol",
        "FloatCol",
        "BoolCol",
        "NoneCol",
        "DatetimeCol",
        "DateCol",
        "TimeCol",
        "DecimalCol",
        "FloatNanCol",
        "NaCol",
        "NaTCol",
        "TimestampCol",
        "TzDatetimeCol",
        "LargeIntCol",
        "LargeFloatCol",
    ]
    for col_idx, header in enumerate(headers, 1):
        ws2.cell(row=1, column=col_idx, value=header)

    # One row of placeholder data so the table exists
    ws2.cell(row=2, column=1, value="placeholder")
    ws2.cell(row=2, column=2, value=0)
    ws2.cell(row=2, column=3, value=0.0)
    ws2.cell(row=2, column=4, value=False)
    ws2.cell(row=2, column=5, value=None)
    ws2.cell(row=2, column=6, value=datetime(2025, 1, 1))
    ws2.cell(row=2, column=7, value=date(2025, 1, 1))
    ws2.cell(row=2, column=8, value=time(12, 0)).number_format = "hh:mm:ss"
    ws2.cell(row=2, column=9, value=Decimal(0))
    ws2.cell(row=2, column=10, value=None)  # FloatNanCol
    ws2.cell(row=2, column=11, value=None)  # NaCol
    ws2.cell(row=2, column=12, value=None)  # NaTCol
    ws2.cell(row=2, column=13, value=datetime(2025, 1, 1))  # TimestampCol
    ws2.cell(row=2, column=14, value=datetime(2025, 1, 1))  # TzDatetimeCol
    ws2.cell(row=2, column=15, value=0)  # LargeIntCol
    ws2.cell(row=2, column=16, value=0.0)  # LargeFloatCol

    table = openpyxl.worksheet.table.Table(
        displayName="CrossDtypeTable",
        ref="A1:P2",
        tableColumns=[
            openpyxl.worksheet.table.TableColumn(id=1, name="StrCol"),
            openpyxl.worksheet.table.TableColumn(id=2, name="IntCol"),
            openpyxl.worksheet.table.TableColumn(id=3, name="FloatCol"),
            openpyxl.worksheet.table.TableColumn(id=4, name="BoolCol"),
            openpyxl.worksheet.table.TableColumn(id=5, name="NoneCol"),
            openpyxl.worksheet.table.TableColumn(id=6, name="DatetimeCol"),
            openpyxl.worksheet.table.TableColumn(id=7, name="DateCol"),
            openpyxl.worksheet.table.TableColumn(id=8, name="TimeCol"),
            openpyxl.worksheet.table.TableColumn(id=9, name="DecimalCol"),
            openpyxl.worksheet.table.TableColumn(id=10, name="FloatNanCol"),
            openpyxl.worksheet.table.TableColumn(id=11, name="NaCol"),
            openpyxl.worksheet.table.TableColumn(id=12, name="NaTCol"),
            openpyxl.worksheet.table.TableColumn(id=13, name="TimestampCol"),
            openpyxl.worksheet.table.TableColumn(id=14, name="TzDatetimeCol"),
            openpyxl.worksheet.table.TableColumn(id=15, name="LargeIntCol"),
            openpyxl.worksheet.table.TableColumn(id=16, name="LargeFloatCol"),
        ],
    )
    ws2.add_table(table)

    wb.save(path)
    wb.close()


@pytest.fixture(scope="session")
def cross_backend_dtype_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Workbook with all dtypes covered for cross-backend comparison tests.

    Includes named cells and a table with columns for basic types + edge cases
    (NaN, pd.NA, pd.NaT, pd.Timestamp, tz-aware datetime, large numbers).
    """
    path = tmp_path_factory.mktemp("cross_dtype") / "cross_backend_dtype_template.xlsx"
    _build_cross_backend_dtype_template(path)
    return path
