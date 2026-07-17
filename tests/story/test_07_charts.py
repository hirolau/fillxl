"""Chapter 7 — Charts and Table References.
=========================================

When you fill a table with more or fewer rows, any **charts** that plot
data from that table need their data ranges updated.  fillxl does this
automatically for all standard chart types (bar, line, pie, scatter, area).

How it works under the hood:

1. After filling a table, fillxl records the old and new table dimensions
   in a :class:`~fillxl._charts.TableChangeSpecs`.
2. It scans every chart series in the workbook.
3. For each series that references the changed table, it updates the
   formula reference and the cached point count.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
from openpyxl.chart import (  # ty:ignore[unresolved-import]
    AreaChart,
    BarChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
)
from openpyxl.worksheet.table import Table, TableColumn  # ty:ignore[unresolved-import]

from fillxl import fill

if TYPE_CHECKING:
    from pathlib import Path

# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_chart_template(path: Path, chart_type: str = "bar", data_rows: int = 4) -> None:
    """Build a workbook with a table and a chart bound to it."""
    wb = openpyxl.Workbook()
    ws_data = wb.active
    ws_data.title = "Data"

    headers = ["Category", "Value"]
    for ci, h in enumerate(headers, 1):
        ws_data.cell(row=1, column=ci, value=h)

    for ri in range(data_rows):
        ws_data.cell(row=ri + 2, column=1, value=f"Cat{ri + 1}")
        ws_data.cell(row=ri + 2, column=2, value=(ri + 1) * 10)

    last_col = openpyxl.utils.get_column_letter(len(headers))  # ty:ignore[possibly-missing-submodule]
    table = Table(
        displayName="ChartData",
        ref=f"A1:{last_col}{data_rows + 1}",
        tableColumns=[
            TableColumn(id=1, name="Category"),
            TableColumn(id=2, name="Value"),
        ],
    )
    ws_data.add_table(table)

    chart_ws = wb.create_sheet("Charts")

    chart_types = {
        "bar": BarChart,
        "line": LineChart,
        "pie": PieChart,
        "scatter": ScatterChart,
        "area": AreaChart,
    }
    chart_cls = chart_types.get(chart_type, BarChart)
    chart = chart_cls()
    chart.title = f"{chart_type.title()} Chart"

    data_ref = Reference(ws_data, min_col=2, min_row=1, max_row=data_rows + 1)
    cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=data_rows + 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    chart_ws.add_chart(chart, "A1")
    wb.save(path)
    wb.close()


def _get_chart_end_rows(output_path: Path) -> list[int]:
    """Extract end-row numbers from all chart series references."""
    wb = openpyxl.load_workbook(output_path)
    try:
        end_rows: list[int] = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for chart in getattr(ws, "_charts", []):
                for ser in chart.ser:
                    for label in ("cat", "val", "xVal", "yVal"):
                        item = getattr(ser, label, None)
                        if item is None:
                            continue
                        for ref_type in ("strRef", "numRef"):
                            ref = getattr(item, ref_type, None)
                            if ref is not None and ref.f:
                                end_rows.append(
                                    int("".join(ch for ch in ref.f.split("$")[-1] if ch.isdigit())),
                                )
        return end_rows
    finally:
        wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 7.1 Chart references update when tables resize
# ═══════════════════════════════════════════════════════════════════════════════


class TestChartReferenceUpdates:
    """Chart data ranges are updated when tables grow or shrink."""

    def test_bar_chart_reference_expands_when_table_grows(
        self,
        output_path: Path,
    ) -> None:
        """Fill 10 rows into a 4-row table — chart refs must point to row 11."""
        _make_chart_template(output_path.with_name("tmpl.xlsx"), "bar", data_rows=4)

        data = [{"Category": f"C{i}", "Value": i * 10} for i in range(1, 11)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"ChartData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        assert all(r == 11 for r in end_rows), f"Expected all refs to end at row 11, got {end_rows}"

    def test_chart_reference_contracts_when_table_shrinks(
        self,
        output_path: Path,
    ) -> None:
        """Fill 2 rows into a 4-row table — chart refs must shrink to row 3."""
        _make_chart_template(output_path.with_name("tmpl.xlsx"), "bar", data_rows=4)

        data = [{"Category": "A", "Value": 10}, {"Category": "B", "Value": 20}]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"ChartData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        assert all(r == 3 for r in end_rows), f"Expected all refs to end at row 3, got {end_rows}"

    def test_line_chart_references_update_when_table_grows(
        self,
        output_path: Path,
    ) -> None:
        """Line charts are updated just like bar charts."""
        _make_chart_template(output_path.with_name("tmpl.xlsx"), "line", data_rows=4)

        data = [{"Category": f"C{i}", "Value": i} for i in range(1, 8)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"ChartData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        # 7 data rows + 1 header = row 8
        assert all(r == 8 for r in end_rows), f"Expected all refs to end at row 8, got {end_rows}"

    def test_pie_chart_references_update_when_table_grows(
        self,
        output_path: Path,
    ) -> None:
        """Pie chart category and value references are both updated."""
        _make_chart_template(output_path.with_name("tmpl.xlsx"), "pie", data_rows=3)

        data = [{"Category": f"C{i}", "Value": i * 5} for i in range(1, 7)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"ChartData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        # 6 data rows + 1 header = row 7
        assert all(r == 7 for r in end_rows), f"Expected all refs to end at row 7, got {end_rows}"

    def test_scatter_chart_xval_and_yval_references_both_update(
        self,
        output_path: Path,
    ) -> None:
        """Scatter charts have separate xVal and yVal — both are updated."""
        _make_chart_template(output_path.with_name("tmpl.xlsx"), "scatter", data_rows=4)

        data = [{"Category": f"C{i}", "Value": i * 3} for i in range(1, 6)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"ChartData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        assert all(r == 7 for r in end_rows), f"Expected all refs to end at row 7, got {end_rows}"

    def test_area_chart_references_update_when_table_grows(
        self,
        output_path: Path,
    ) -> None:
        """Area charts are also kept in sync."""
        _make_chart_template(output_path.with_name("tmpl.xlsx"), "area", data_rows=4)

        data = [{"Category": f"C{i}", "Value": i * 2} for i in range(1, 9)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"ChartData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        # 8 data rows + 1 header = row 9
        assert all(r == 9 for r in end_rows), f"Expected all refs to end at row 9, got {end_rows}"


# ═══════════════════════════════════════════════════════════════════════════════
# 7.2 The TableChangeSpecs record
# ═══════════════════════════════════════════════════════════════════════════════


class TestTableChangeSpecs:
    """The data structure that tracks table dimension changes."""

    def test_table_change_specs_captures_old_and_new_references(self) -> None:
        """A TableChangeSpecs holds the old ref, new ref, and worksheet title."""
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs(
            old_ref="A1:B5",
            new_ref="A1:B10",
            worksheet_title="Data",
        )
        assert spec.old_ref == "A1:B5"
        assert spec.new_ref == "A1:B10"
        assert spec.worksheet_title == "Data"


# ═══════════════════════════════════════════════════════════════════════════════
# 7.3 Charts Following Tables — real-world scenarios
# ═══════════════════════════════════════════════════════════════════════════════


def _make_charts_following_template(
    path: Path,
    *,
    num_charts: int = 1,
    num_series: int = 1,
    data_rows: int = 4,
) -> None:
    """Build a workbook with a table and one or more charts tracking it.

    Args:
        path: Output path for the template.
        num_charts: Number of separate chart objects (each on Charts sheet).
        num_series: Number of data series per chart.
        data_rows: Number of initial data rows in the table.

    """
    wb = openpyxl.Workbook()
    ws_data = wb.active
    ws_data.title = "Data"

    # Headers: Category + Value1 … ValueN
    headers = ["Category"] + [f"Value{s}" for s in range(1, num_series + 1)]
    for ci, h in enumerate(headers, 1):
        ws_data.cell(row=1, column=ci, value=h)

    for ri in range(data_rows):
        ws_data.cell(row=ri + 2, column=1, value=f"Cat{ri + 1}")
        for si in range(num_series):
            ws_data.cell(row=ri + 2, column=si + 2, value=(ri + 1) * (si + 1) * 10)

    last_col = openpyxl.utils.get_column_letter(len(headers))  # ty:ignore[possibly-missing-submodule]
    table = Table(
        displayName="TrackedData",
        ref=f"A1:{last_col}{data_rows + 1}",
        tableColumns=[
            TableColumn(id=1, name="Category"),
            *(TableColumn(id=si + 2, name=f"Value{si + 1}") for si in range(num_series)),
        ],
    )
    ws_data.add_table(table)

    chart_ws = wb.create_sheet("Charts")

    for ci in range(num_charts):
        chart = BarChart()
        chart.title = f"Chart {ci + 1}"

        for si in range(num_series):
            data_ref = Reference(ws_data, min_col=si + 2, min_row=1, max_row=data_rows + 1)
            chart.add_data(data_ref, titles_from_data=True)

        cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=data_rows + 1)
        chart.set_categories(cats_ref)

        chart_ws.add_chart(chart, f"A{(ci * 18) + 1}")

    wb.save(path)
    wb.close()


class TestChartsFollowingTables:
    """Charts adapt their data ranges when the underlying table changes."""

    def test_chart_expands_when_rows_are_added(
        self,
        output_path: Path,
    ) -> None:
        """Adding rows to the table expands chart data ranges."""
        _make_charts_following_template(
            output_path.with_name("tmpl.xlsx"),
            num_charts=1,
            num_series=1,
            data_rows=3,
        )

        data = [{"Category": f"C{i}", "Value1": i * 10} for i in range(1, 8)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"TrackedData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        # 7 data rows + 1 header = row 8
        assert all(r == 8 for r in end_rows), f"Expected refs at row 8, got {end_rows}"

    def test_chart_contracts_when_rows_are_removed(
        self,
        output_path: Path,
    ) -> None:
        """Removing rows from the table shrinks chart data ranges."""
        _make_charts_following_template(
            output_path.with_name("tmpl.xlsx"),
            num_charts=1,
            num_series=1,
            data_rows=6,
        )

        data = [{"Category": "C1", "Value1": 10}, {"Category": "C2", "Value1": 20}]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"TrackedData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        # 2 data rows + 1 header = row 3
        assert all(r == 3 for r in end_rows), f"Expected refs at row 3, got {end_rows}"

    def test_multiple_charts_can_track_the_same_table(
        self,
        output_path: Path,
    ) -> None:
        """Two independent charts both update when the table resizes."""
        _make_charts_following_template(
            output_path.with_name("tmpl.xlsx"),
            num_charts=2,
            num_series=1,
            data_rows=3,
        )

        data = [{"Category": f"C{i}", "Value1": i * 5} for i in range(1, 6)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"TrackedData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        # 5 data rows + 1 header = row 6
        assert all(r == 6 for r in end_rows), f"Expected refs at row 6, got {end_rows}"

    def test_multiple_series_update_together(
        self,
        output_path: Path,
    ) -> None:
        """A chart with two data series has both series references updated."""
        _make_charts_following_template(
            output_path.with_name("tmpl.xlsx"),
            num_charts=1,
            num_series=2,
            data_rows=2,
        )

        data = [
            {"Category": "A", "Value1": 10, "Value2": 100},
            {"Category": "B", "Value1": 20, "Value2": 200},
            {"Category": "C", "Value1": 30, "Value2": 300},
            {"Category": "D", "Value1": 40, "Value2": 400},
        ]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"TrackedData": data},
        )

        end_rows = _get_chart_end_rows(output_path)
        # 4 data rows + 1 header = row 5 -> all 4 refs (cats + 2 vals + 1?)
        # Actually cats + val1 + val2 = 3 refs per chart, all point to row 5
        assert all(r == 5 for r in end_rows), f"Expected all refs at row 5, got {end_rows}"

    def test_charts_survive_save_and_reload(
        self,
        output_path: Path,
    ) -> None:
        """Chart references remain correct after an openpyxl save/reload cycle."""
        _make_charts_following_template(
            output_path.with_name("tmpl.xlsx"),
            num_charts=1,
            num_series=1,
            data_rows=2,
        )

        data = [{"Category": f"C{i}", "Value1": i * 10} for i in range(1, 6)]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"TrackedData": data},
        )

        # First verification: chart refs point to row 6 (5 data + 1 header)
        end_rows_first = _get_chart_end_rows(output_path)
        assert all(r == 6 for r in end_rows_first), (
            f"Expected refs at row 6 on first check, got {end_rows_first}"
        )

        # Simulate save/reload cycle via openpyxl
        wb = openpyxl.load_workbook(output_path)
        wb.save(output_path)
        wb.close()

        # Re-verify after reload
        end_rows_after = _get_chart_end_rows(output_path)
        assert all(r == 6 for r in end_rows_after), (
            f"Expected refs at row 6 after reload, got {end_rows_after}"
        )
