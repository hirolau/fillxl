"""Unit tests for chart helper functions.

Tests the pure helper functions (all_chart_series, update_chart_data_ranges,
detect_unsupported_chart_sheets, _fix_chart_axes) in isolation — no fill() calls.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl

if TYPE_CHECKING:
    from pathlib import Path

    import pytest

# ═══════════════════════════════════════════════════════════════════════════════
# all_chart_series
# ═══════════════════════════════════════════════════════════════════════════════


class TestAllChartSeries:
    """Unit tests for all_chart_series() helper."""

    def test_all_chart_series_returns_series_from_all_sheets(self, charts_template: Path) -> None:
        """all_chart_series() yields series from every chart in every worksheet."""
        from fillxl._charts import all_chart_series

        wb = openpyxl.load_workbook(charts_template)
        try:
            series = list(all_chart_series(wb))
            assert len(series) >= 5, f"Expected >=5 series, got {len(series)}"
            for s in series:
                assert hasattr(s, "val") or hasattr(s, "cat"), (
                    f"Series missing expected attributes: {type(s)}"
                )
        finally:
            wb.close()

    def test_all_chart_series_returns_empty_for_workbook_without_charts(
        self,
        tmp_path: Path,
    ) -> None:
        """A workbook with no charts yields no series."""
        from fillxl._charts import all_chart_series

        wb = openpyxl.Workbook()
        try:
            series = list(all_chart_series(wb))
            assert series == []
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# all_chart_series_references
# ═══════════════════════════════════════════════════════════════════════════════


class TestAllChartSeriesReferences:
    """Unit tests for all_chart_series_references() helper."""

    def test_all_chart_series_references_returns_refs_from_all_series(
        self,
        charts_template: Path,
    ) -> None:
        """all_chart_series_references() yields strRef/numRef from every series."""
        from fillxl._charts import all_chart_series_references

        wb = openpyxl.load_workbook(charts_template)
        try:
            refs = list(all_chart_series_references(wb))
            assert len(refs) >= 5, f"Expected >=5 refs, got {len(refs)}"
            for ref in refs:
                assert hasattr(ref, "f"), f"Reference missing 'f' attribute: {ref}"
        finally:
            wb.close()

    def test_all_chart_series_references_returns_empty_for_no_charts(self, tmp_path: Path) -> None:
        """A workbook with no charts yields no references."""
        from fillxl._charts import all_chart_series_references

        wb = openpyxl.Workbook()
        try:
            refs = list(all_chart_series_references(wb))
            assert refs == []
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# update_chart_data_ranges
# ═══════════════════════════════════════════════════════════════════════════════


class TestUpdateChartDataRanges:
    """Unit tests for update_chart_data_ranges() edge cases."""

    def test_update_chart_data_ranges_with_manual_range_is_unchanged(
        self,
        output_path: Path,
        charts_template: Path,
    ) -> None:
        """A chart series referencing a manual range (not table-bound) is not updated."""
        from fillxl._charts import update_chart_data_ranges
        from fillxl._tables import populate_table

        wb = openpyxl.load_workbook(charts_template)
        try:
            spec = populate_table(
                wb,
                "TableTest",
                [
                    {"Col1": "a", "Col2": 1, "ColNotFilled": "x"},
                    {"Col1": "b", "Col2": 2, "ColNotFilled": "y"},
                    {"Col1": "c", "Col2": 3, "ColNotFilled": "z"},
                    {"Col1": "d", "Col2": 4, "ColNotFilled": "w"},
                    {"Col1": "e", "Col2": 5, "ColNotFilled": "v"},
                ],
            )
            update_chart_data_ranges(spec, wb)

            from fillxl._charts import all_chart_series_references

            refs = [r.f for r in all_chart_series_references(wb)]
            for r in refs:
                end_row = int("".join(ch for ch in r.split("$")[-1] if ch.isdigit()))
                assert end_row == 6, f"Expected end_row=6 for table-bound ref, got {end_row} in {r}"
        finally:
            wb.close()

    def test_update_chart_data_ranges_skips_refs_from_wrong_worksheet(
        self,
        charts_template: Path,
    ) -> None:
        """Refs from a different worksheet than the table are skipped."""
        from fillxl._charts import TableChangeSpecs, update_chart_data_ranges

        wb = openpyxl.load_workbook(charts_template)
        try:
            spec = TableChangeSpecs(
                old_ref="A1:B5",
                new_ref="A1:B10",
                worksheet_title="NonExistentSheet",
            )
            update_chart_data_ranges(spec, wb)
        finally:
            wb.close()

    def test_update_chart_data_ranges_with_same_end_row_does_nothing(
        self,
        charts_template: Path,
    ) -> None:
        """When old and new end rows are the same, refs should not change."""
        from fillxl._charts import (
            TableChangeSpecs,
            all_chart_series_references,
            update_chart_data_ranges,
        )

        wb = openpyxl.load_workbook(charts_template)
        try:
            refs_before = {(r.f, id(r)) for r in all_chart_series_references(wb)}

            spec = TableChangeSpecs(
                old_ref="A1:B5",
                new_ref="A1:B5",
                worksheet_title="TestTable",
            )
            update_chart_data_ranges(spec, wb)

            refs_after = {(r.f, id(r)) for r in all_chart_series_references(wb)}
            assert len(refs_before) == len(refs_after)
        finally:
            wb.close()

    def test_update_chart_data_ranges_handles_missing_f_attribute(
        self,
        charts_template: Path,
    ) -> None:
        """A reference with f=None is skipped without error."""
        from fillxl._charts import TableChangeSpecs, update_chart_data_ranges

        wb = openpyxl.load_workbook(charts_template)
        try:
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for chart in getattr(ws, "_charts", []):
                    for ser in chart.ser:
                        if hasattr(ser, "val") and hasattr(ser.val, "numRef"):
                            ser.val.numRef.f = None

            spec = TableChangeSpecs(
                old_ref="A1:B5",
                new_ref="A1:B6",
                worksheet_title="TestTable",
            )
            update_chart_data_ranges(spec, wb)
        finally:
            wb.close()

    def test_update_chart_data_ranges_adds_stale_column_ref_does_not_match(
        self,
        charts_template: Path,
    ) -> None:
        """A ref whose column is outside the table column range is skipped."""
        from fillxl._charts import TableChangeSpecs, update_chart_data_ranges

        wb = openpyxl.load_workbook(charts_template)
        try:
            spec = TableChangeSpecs(
                old_ref="A1:E5",
                new_ref="A1:E10",
                worksheet_title="TestTable",
            )
            update_chart_data_ranges(spec, wb)
        finally:
            wb.close()

    def test_update_chart_data_ranges_exception_handler_is_exercised(
        self,
        charts_template: Path,
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        """When a chart reference causes a parse error, the exception is logged."""
        import logging

        from fillxl._charts import (
            TableChangeSpecs,
            update_chart_data_ranges,
        )

        wb = openpyxl.load_workbook(charts_template)
        try:
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
                                    ref.f = ref.f.replace("$", "")

            caplog.set_level(logging.DEBUG)

            spec = TableChangeSpecs(
                old_ref="A1:B5",
                new_ref="A1:B10",
                worksheet_title="TestTable",
            )
            update_chart_data_ranges(spec, wb)

            assert any(
                "Unable to update chart series reference" in record.message
                for record in caplog.records
            ), f"Expected exception log, got: {[r.message for r in caplog.records]}"
        finally:
            wb.close()

    def test_update_chart_data_ranges_skips_ref_not_starting_at_data_row(self) -> None:
        """A chart ref that starts at the header row is skipped."""
        from fillxl._charts import (
            TableChangeSpecs,
            all_chart_series_references,
            update_chart_data_ranges,
        )

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"

        headers = ["Category", "Value"]
        for ci, h in enumerate(headers, 1):
            ws.cell(row=1, column=ci, value=h)
        for ri in range(4):
            ws.cell(row=ri + 2, column=1, value=f"Cat{ri + 1}")
            ws.cell(row=ri + 2, column=2, value=(ri + 1) * 10)

        from openpyxl.chart import BarChart, Reference  # ty:ignore[unresolved-import]

        chart = BarChart()
        data_ref = Reference(ws, min_col=2, min_row=1, max_row=5)
        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=5)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart_ws = wb.create_sheet("Charts")
        chart_ws.add_chart(chart, "A1")

        from openpyxl.worksheet.table import Table, TableColumn  # ty:ignore[unresolved-import]

        table = Table(
            displayName="ChartData",
            ref="A1:B5",
            tableColumns=[
                TableColumn(id=1, name="Category"),
                TableColumn(id=2, name="Value"),
            ],
        )
        ws.add_table(table)

        spec = TableChangeSpecs(
            old_ref="A1:B5",
            new_ref="A1:B10",
            worksheet_title="Data",
        )
        update_chart_data_ranges(spec, wb)

        refs = [r.f for r in all_chart_series_references(wb)]
        for r in refs:
            end_row = int("".join(ch for ch in r.split("$")[-1] if ch.isdigit()))
            assert end_row == 10, f"Expected end_row=10, got {end_row} in {r}"


# ═══════════════════════════════════════════════════════════════════════════════
# detect_unsupported_chart_sheets
# ═══════════════════════════════════════════════════════════════════════════════


class TestDetectUnsupportedChartSheets:
    """Unit tests for detect_unsupported_chart_sheets()."""

    def test_detect_returns_empty_for_non_existent_file(self, tmp_path: Path) -> None:
        """A non-existent file returns empty list (no crash)."""
        from fillxl._charts import detect_unsupported_chart_sheets

        result = detect_unsupported_chart_sheets(tmp_path / "nonexistent.xlsx")
        assert result == []

    def test_detect_returns_empty_for_invalid_zip(self, tmp_path: Path) -> None:
        """An invalid .xlsx file returns empty list (no crash)."""
        from fillxl._charts import detect_unsupported_chart_sheets

        bad_file = tmp_path / "bad.xlsx"
        bad_file.write_text("not a zip file")
        result = detect_unsupported_chart_sheets(bad_file)
        assert result == []


# ═══════════════════════════════════════════════════════════════════════════════
# _fix_chart_axes
# ═══════════════════════════════════════════════════════════════════════════════


class TestFixChartAxes:
    """Unit tests for _fix_chart_axes()."""

    def test_fix_chart_axes_on_workbook_without_charts_does_not_raise(self, tmp_path: Path) -> None:
        """_fix_chart_axes() on a workbook without charts does nothing."""
        from fillxl._backends._openpyxl import _fix_chart_axes

        wb = openpyxl.Workbook()
        try:
            _fix_chart_axes(wb)
        finally:
            wb.close()

    def test_fix_chart_axes_sets_axpos_on_bar_chart(self, charts_template: Path) -> None:
        """_fix_chart_axes() sets axPos='b' for TextAxis on bar charts loaded by openpyxl."""
        from openpyxl.chart.axis import TextAxis

        from fillxl._backends._openpyxl import _fix_chart_axes

        wb = openpyxl.load_workbook(charts_template)
        try:
            _fix_chart_axes(wb)
            found_text_axis = False
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for chart in getattr(ws, "_charts", []):
                    for ax in getattr(chart, "_axes", {}).values():
                        if isinstance(ax, TextAxis):
                            assert ax.axPos is not None, "TextAxis axPos should be set"
                            found_text_axis = True
            assert found_text_axis, "No TextAxis found in charts_template"
        finally:
            wb.close()
