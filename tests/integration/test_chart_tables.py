"""Fill each table in test_template_2_charts.xlsx and verify via openpyxl.

One test per table — 10 tests total.  Each test provides fresh data, fills
the corresponding table, then reads back the output to confirm the values
landed in the right cells.
"""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import fill
from tests.conftest import CHARTS_TEMPLATE_2

if TYPE_CHECKING:
    from pathlib import Path

# All tests in this module use CHARTS_TEMPLATE_2 which contains a "10-Waterfall"
# chartEx sheet.  openpyxl cannot preserve chartEx charts, so both load() and
# save() emit a UserWarning.  These warnings are expected and documented —
# suppress them to keep output clean.
pytestmark = pytest.mark.filterwarnings("ignore:.*chartEx.*:UserWarning")

# ── Helpers ──────────────────────────────────────────────────────────────────


def _read_rows(template: Path, table_name: str) -> list[list[object]]:
    """Read the data rows of *table_name* from *template*."""
    wb = openpyxl.load_workbook(template)
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if table_name in ws.tables:
                table = ws.tables[table_name]
                start, end = table.ref.split(":")
                start_row = int("".join(ch for ch in start if ch.isdigit()))
                end_row = int("".join(ch for ch in end if ch.isdigit()))
                ncols = len(table.tableColumns)
                rows: list[list[object]] = []
                for r in range(start_row + 1, end_row + 1):
                    row: list[object] = [ws.cell(row=r, column=c + 1).value for c in range(ncols)]
                    rows.append(row)
                return rows
        pytest.fail(f"Table {table_name!r} not found in workbook")
    finally:
        wb.close()


# ── 1-Column Chart: ColumnChartTable ─────────────────────────────────────────


class TestColumnChartTable:
    """ColumnChartTable - Region / Q1-Q4 quarterly sales data."""

    def test_fill_column_chart_table_with_fewer_rows(self, output_path: Path) -> None:
        """Fill ColumnChartTable with 3 rows (original has 4)."""
        data = [
            {"Region": "Nordic", "Q1": 10000, "Q2": 15000, "Q3": 14000, "Q4": 18000},
            {"Region": "Baltic", "Q1": 8000, "Q2": 9000, "Q3": 9500, "Q4": 11000},
            {"Region": "Balkans", "Q1": 6000, "Q2": 7000, "Q3": 7500, "Q4": 8000},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"ColumnChartTable": data})

        rows = _read_rows(output_path, "ColumnChartTable")
        assert len(rows) == 3
        assert rows[0] == ["Nordic", 10000, 15000, 14000, 18000]
        assert rows[-1] == ["Balkans", 6000, 7000, 7500, 8000]

    def test_fill_column_chart_table_with_more_rows(self, output_path: Path) -> None:
        """Fill ColumnChartTable with 6 rows (original has 4)."""
        data = [
            {"Region": f"R{i}", "Q1": i * 1000, "Q2": i * 1100, "Q3": i * 1200, "Q4": i * 1300}
            for i in range(1, 7)
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"ColumnChartTable": data})

        rows = _read_rows(output_path, "ColumnChartTable")
        assert len(rows) == 6
        assert rows[0] == ["R1", 1000, 1100, 1200, 1300]
        assert rows[-1] == ["R6", 6000, 6600, 7200, 7800]


# ── 2-Line Chart: LineChartTable ─────────────────────────────────────────────


class TestLineChartTable:
    """LineChartTable — Date / 2024 / 2025 monthly data."""

    def test_fill_line_chart_table(self, output_path: Path) -> None:
        """Fill LineChartTable with 3 rows of monthly data."""
        data = [
            {"Date": datetime(2026, 1, 1), "2024": 200000, "2025": 220000},
            {"Date": datetime(2026, 2, 1), "2024": 210000, "2025": 230000},
            {"Date": datetime(2026, 3, 1), "2024": 205000, "2025": 240000},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"LineChartTable": data})

        rows = _read_rows(output_path, "LineChartTable")
        assert len(rows) == 3
        assert rows[0][0] == datetime(2026, 1, 1)
        assert rows[0][1] == 200000
        assert rows[0][2] == 220000


# ── 3-Pie Chart: PieChartTable ───────────────────────────────────────────────


class TestPieChartTable:
    """PieChartTable — Company / Market Share."""

    def test_fill_pie_chart_table(self, output_path: Path) -> None:
        """Fill PieChartTable with 4 companies."""
        data = [
            {"Company": "Alpha", "Market Share": 0.35},
            {"Company": "Beta", "Market Share": 0.25},
            {"Company": "Gamma", "Market Share": 0.20},
            {"Company": "Delta", "Market Share": 0.20},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"PieChartTable": data})

        rows = _read_rows(output_path, "PieChartTable")
        assert len(rows) == 4
        assert rows[0] == ["Alpha", 0.35]
        assert rows[-1] == ["Delta", 0.20]


# ── 4-Bar Chart: BarChartTable ───────────────────────────────────────────────


class TestBarChartTable:
    """BarChartTable — Department / Employees."""

    def test_fill_bar_chart_table(self, output_path: Path) -> None:
        """Fill BarChartTable with 3 departments."""
        data = [
            {"Department": "R&D", "Employees": 50},
            {"Department": "Support", "Employees": 30},
            {"Department": "Admin", "Employees": 15},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"BarChartTable": data})

        rows = _read_rows(output_path, "BarChartTable")
        assert len(rows) == 3
        assert rows[0] == ["R&D", 50]
        assert rows[-1] == ["Admin", 15]


# ── 5-Area Chart: AreaChartTable ─────────────────────────────────────────────


class TestAreaChartTable:
    """AreaChartTable — Week / Desktop / Mobile / Tablet."""

    def test_fill_area_chart_table(self, output_path: Path) -> None:
        """Fill AreaChartTable with 4 weeks of traffic data."""
        data = [
            {"Week": "W1", "Desktop": 30000, "Mobile": 50000, "Tablet": 8000},
            {"Week": "W2", "Desktop": 32000, "Mobile": 55000, "Tablet": 9000},
            {"Week": "W3", "Desktop": 35000, "Mobile": 60000, "Tablet": 10000},
            {"Week": "W4", "Desktop": 38000, "Mobile": 65000, "Tablet": 11000},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"AreaChartTable": data})

        rows = _read_rows(output_path, "AreaChartTable")
        assert len(rows) == 4
        assert rows[0] == ["W1", 30000, 50000, 8000]
        assert rows[-1] == ["W4", 38000, 65000, 11000]


# ── 6-Scatter Chart: ScatterChartTable ───────────────────────────────────────


class TestScatterChartTable:
    """ScatterChartTable — Product / Price ($) / Units Sold."""

    def test_fill_scatter_chart_table(self, output_path: Path) -> None:
        """Fill ScatterChartTable with 4 products."""
        data = [
            {"Product": "Widget Pro", "Price ($)": 39, "Units Sold": 2000},
            {"Product": "Widget Lite", "Price ($)": 19, "Units Sold": 4500},
            {"Product": "Gadget X", "Price ($)": 59, "Units Sold": 1200},
            {"Product": "Gadget Mini", "Price ($)": 29, "Units Sold": 3100},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"ScatterChartTable": data})

        rows = _read_rows(output_path, "ScatterChartTable")
        assert len(rows) == 4
        assert rows[0] == ["Widget Pro", 39, 2000]
        assert rows[-1] == ["Gadget Mini", 29, 3100]


# ── 7-Doughnut Chart: DoughnutChartTable ─────────────────────────────────────


class TestDoughnutChartTable:
    """DoughnutChartTable — Category / Budget."""

    def test_fill_doughnut_chart_table(self, output_path: Path) -> None:
        """Fill DoughnutChartTable with 4 budget categories."""
        data = [
            {"Category": "Salaries", "Budget": 500000},
            {"Category": "Marketing", "Budget": 100000},
            {"Category": "R&D", "Budget": 200000},
            {"Category": "Infra", "Budget": 75000},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"DoughnutChartTable": data})

        rows = _read_rows(output_path, "DoughnutChartTable")
        assert len(rows) == 4
        assert rows[0] == ["Salaries", 500000]
        assert rows[-1] == ["Infra", 75000]


# ── 8-Stacked Column: StackedColumnChartTable ────────────────────────────────


class TestStackedColumnChartTable:
    """StackedColumnChartTable — Quarter / Electronics / Clothing / Food / Home."""

    def test_fill_stacked_column_chart_table(self, output_path: Path) -> None:
        """Fill StackedColumnChartTable with 2 quarters."""
        data = [
            {
                "Quarter": "Q1 2026",
                "Electronics": 90000,
                "Clothing": 45000,
                "Food": 40000,
                "Home": 30000,
            },
            {
                "Quarter": "Q2 2026",
                "Electronics": 95000,
                "Clothing": 50000,
                "Food": 42000,
                "Home": 32000,
            },
        ]
        fill(
            template=CHARTS_TEMPLATE_2,
            output=output_path,
            tables={"StackedColumnChartTable": data},
        )

        rows = _read_rows(output_path, "StackedColumnChartTable")
        assert len(rows) == 2
        assert rows[0] == ["Q1 2026", 90000, 45000, 40000, 30000]
        assert rows[-1] == ["Q2 2026", 95000, 50000, 42000, 32000]


# ── 9-Combo Chart: ComboChartTable ───────────────────────────────────────────


class TestComboChartTable:
    """ComboChartTable — Date / Revenue / Growth Rate."""

    def test_fill_combo_chart_table(self, output_path: Path) -> None:
        """Fill ComboChartTable with 3 months of data."""
        data = [
            {"Date": datetime(2027, 1, 1), "Revenue": 300000, "Growth Rate": 0.04},
            {"Date": datetime(2027, 2, 1), "Revenue": 315000, "Growth Rate": 0.05},
            {"Date": datetime(2027, 3, 1), "Revenue": 330000, "Growth Rate": 0.05},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"ComboChartTable": data})

        rows = _read_rows(output_path, "ComboChartTable")
        assert len(rows) == 3
        assert rows[0][0] == datetime(2027, 1, 1)
        assert rows[0][1] == 300000
        assert rows[0][2] == 0.04


# ── 10-Waterfall: WaterfallChartTable ────────────────────────────────────────


class TestWaterfallChartTable:
    """WaterfallChartTable — Category / Amount."""

    def test_fill_waterfall_chart_table(self, output_path: Path) -> None:
        """Fill WaterfallChartTable with 4 P&L items."""
        data = [
            {"Category": "Revenue", "Amount": 600000},
            {"Category": "COGS", "Amount": -250000},
            {"Category": "OpEx", "Amount": -180000},
            {"Category": "Net Income", "Amount": 170000},
        ]
        fill(template=CHARTS_TEMPLATE_2, output=output_path, tables={"WaterfallChartTable": data})

        rows = _read_rows(output_path, "WaterfallChartTable")
        assert len(rows) == 4
        assert rows[0] == ["Revenue", 600000]
        assert rows[-1] == ["Net Income", 170000]
