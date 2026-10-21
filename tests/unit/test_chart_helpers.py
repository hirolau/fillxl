"""Unit tests for chart helper functions.

Tests the pure helper functions (all_chart_series, update_reference_text,
update_chart_data_ranges, _fix_chart_axes) in isolation — no fill() calls.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, TypeVar

import openpyxl
import pytest

T = TypeVar("T")

if TYPE_CHECKING:
    from pathlib import Path

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
# update_reference_text
# ═══════════════════════════════════════════════════════════════════════════════


class TestUpdateReferenceText:
    """One parser for chart references, shared by both engines."""

    @pytest.mark.parametrize(
        ("text", "expected"),
        [
            ("Data!$B$13:$B$13", "Data!$B$13:$B$20"),
            ("Data!$B$13", "Data!$B$13:$B$20"),
            ("(Data!$B$13,Data!$D$13)", "(Data!$B$13:$B$20,Data!$D$13:$D$20)"),
            ("Data2!$B$13:$B$13", None),
            ("Data!$B$12:$B$13", None),
            ("[1]Data!$B$13", None),
            ("Data!B13:B13", "Data!B13:B20"),
        ],
    )
    def test_only_the_table_data_body_is_rewritten(self, text: str, expected: str | None) -> None:
        """A reference follows the table only when it covers the body's rows."""
        from fillxl._charts import TableChangeSpecs, update_reference_text

        spec = TableChangeSpecs("Data", 12, 2, 13, 13, 20)
        assert update_reference_text(text, spec) == expected

    def test_a_row_below_the_old_last_row_is_not_mistaken_for_it(self) -> None:
        """``$B$134`` must not be read as ``$B$13`` with a stray digit."""
        from fillxl._charts import TableChangeSpecs, update_reference_text

        spec = TableChangeSpecs("Data", 12, 2, 13, 13, 20)
        assert update_reference_text("Data!$B$134", spec) is None

    def test_a_reference_beyond_the_last_table_column_is_left_alone(self) -> None:
        """A chart series to the right of the table does not follow it."""
        from fillxl._charts import TableChangeSpecs, update_reference_text

        spec = TableChangeSpecs("Data", 12, 2, 13, 13, 20)
        assert update_reference_text("Data!$N$13", spec) is None

    def test_a_table_that_keeps_its_size_gives_no_change(self) -> None:
        """Nothing is rewritten when the last row does not move."""
        from fillxl._charts import TableChangeSpecs, update_reference_text

        spec = TableChangeSpecs("Data", 12, 2, 13, 13, 13)
        assert update_reference_text("Data!$B$13:$B$13", spec) is None

    def test_a_sheet_name_with_an_apostrophe_is_unquoted_before_comparison(self) -> None:
        """Excel writes ``'Bob''s Data'``; the stored name is ``Bob's Data``."""
        from fillxl._charts import TableChangeSpecs, update_reference_text

        spec = TableChangeSpecs("Bob's Data", 1, 1, 2, 2, 5)
        assert update_reference_text("'Bob''s Data'!$B$2", spec) == "'Bob''s Data'!$B$2:$B$5"


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
        from fillxl._adapters import TabularAdapter
        from fillxl._charts import update_chart_data_ranges
        from fillxl._tables import populate_table

        wb = openpyxl.load_workbook(charts_template)
        try:
            spec = populate_table(
                wb,
                "TableTest",
                TabularAdapter(
                    [
                        {"Col1": "a", "Col2": 1, "ColNotFilled": "x"},
                        {"Col1": "b", "Col2": 2, "ColNotFilled": "y"},
                        {"Col1": "c", "Col2": 3, "ColNotFilled": "z"},
                        {"Col1": "d", "Col2": 4, "ColNotFilled": "w"},
                        {"Col1": "e", "Col2": 5, "ColNotFilled": "v"},
                    ]
                ),
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
                worksheet_title="NonExistentSheet",
                header_row=1,
                first_col=1,
                last_col=2,
                old_last_row=5,
                new_last_row=10,
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
                worksheet_title="TestTable",
                header_row=1,
                first_col=1,
                last_col=2,
                old_last_row=5,
                new_last_row=5,
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
                worksheet_title="TestTable",
                header_row=1,
                first_col=1,
                last_col=2,
                old_last_row=5,
                new_last_row=6,
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
                worksheet_title="TestTable",
                header_row=1,
                first_col=1,
                last_col=5,
                old_last_row=5,
                new_last_row=10,
            )
            update_chart_data_ranges(spec, wb)
        finally:
            wb.close()

    def test_update_chart_data_ranges_handles_refs_without_dollar_signs(
        self,
        charts_template: Path,
    ) -> None:
        """Refs without "$" separators are parsed and updated, not skipped."""
        from fillxl._charts import (
            TableChangeSpecs,
            all_chart_series_references,
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

            spec = TableChangeSpecs(
                worksheet_title="TestTable",
                header_row=1,
                first_col=1,
                last_col=2,
                old_last_row=5,
                new_last_row=10,
            )
            update_chart_data_ranges(spec, wb)

            updated_refs = [r.f for r in all_chart_series_references(wb) if r.f]
            assert any(r.endswith("10") for r in updated_refs), updated_refs
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

        from openpyxl.chart import BarChart, Reference

        chart = BarChart()
        data_ref = Reference(ws, min_col=2, min_row=1, max_row=5)
        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=5)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart_ws = wb.create_sheet("Charts")
        chart_ws.add_chart(chart, "A1")

        from openpyxl.worksheet.table import Table, TableColumn

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
            worksheet_title="Data",
            header_row=1,
            first_col=1,
            last_col=2,
            old_last_row=5,
            new_last_row=10,
        )
        update_chart_data_ranges(spec, wb)

        refs = [r.f for r in all_chart_series_references(wb)]
        for r in refs:
            end_row = int("".join(ch for ch in r.split("$")[-1] if ch.isdigit()))
            assert end_row == 10, f"Expected end_row=10, got {end_row} in {r}"


# ═══════════════════════════════════════════════════════════════════════════════
# win32com series formulas
# ═══════════════════════════════════════════════════════════════════════════════


class TestWin32ComSeriesFormula:
    """The COM chart updater parses references instead of matching substrings."""

    def test_a_sheet_whose_name_starts_with_the_table_sheet_is_untouched(self) -> None:
        """A chart on "Data2" must not follow a fill of "Data"."""
        from fillxl._backends._win32com import _update_series_formula
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs("Data", 12, 2, 3, 13, 20)
        assert _update_series_formula("=SERIES(T,'Data2'!$B$13:$B$13,1)", spec) is None

    def test_a_reference_to_the_table_body_is_updated(self) -> None:
        """Data-body references grow to the new last row."""
        from fillxl._backends._win32com import _update_series_formula
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs("Data", 12, 2, 3, 13, 20)
        formula = "=SERIES(Title,'Data'!$B$13:$B$13,'Data'!$C$13:$C$13,1)"
        assert (
            _update_series_formula(formula, spec)
            == "=SERIES(Title,'Data'!$B$13:$B$20,'Data'!$C$13:$C$20,1)"
        )

    def test_a_reference_to_the_header_is_not_updated(self) -> None:
        """A series-name reference to the header stays put; data refs grow."""
        from fillxl._backends._win32com import _update_series_formula
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs("Data", 12, 2, 3, 13, 20)
        assert (
            _update_series_formula("=SERIES('Data'!$C$12,'Data'!$B$13:$B$13,1)", spec)
            == "=SERIES('Data'!$C$12,'Data'!$B$13:$B$20,1)"
        )

    def test_a_reference_ending_elsewhere_is_not_updated(self) -> None:
        """Only references that end at the old last row are resized."""
        from fillxl._backends._win32com import _update_series_formula
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs("Data", 12, 2, 3, 13, 20)
        assert _update_series_formula("=SERIES(T,'Data'!$B$13:$B$40,1)", spec) is None

    def test_references_without_dollar_signs_are_still_updated(self) -> None:
        """Refs are parsed, so "$"-less references work too."""
        from fillxl._backends._win32com import _update_series_formula
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs("Data", 12, 2, 3, 13, 20)
        updated = _update_series_formula("=SERIES(T,Data!B13:B13,1)", spec)
        assert updated == "=SERIES(T,Data!B13:B20,1)"

    def test_a_quoted_sheet_name_with_spaces_is_matched_exactly(self) -> None:
        """Quoted sheet names are unquoted before comparing."""
        from fillxl._backends._win32com import _update_series_formula
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs("My Sheet", 12, 2, 3, 13, 20)
        assert (
            _update_series_formula("=SERIES(T,'My Sheet'!$B$13:$B$13,1)", spec)
            == "=SERIES(T,'My Sheet'!$B$13:$B$20,1)"
        )

    def test_an_unparseable_reference_is_skipped(self) -> None:
        """A formula with a broken range reference is returned unchanged."""
        from fillxl._backends._win32com import _update_series_formula
        from fillxl._charts import TableChangeSpecs

        spec = TableChangeSpecs("Data", 12, 2, 3, 13, 20)
        assert _update_series_formula("=SERIES(T,'Data'!$B$:B$13,1)", spec) is None


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

    @staticmethod
    def _fixed_chart(chart: T) -> T:
        from fillxl._backends._openpyxl import _fix_chart_axes

        wb = openpyxl.Workbook()
        try:
            wb.active.add_chart(chart, "A1")
            _fix_chart_axes(wb)
        finally:
            wb.close()
        return chart

    def test_fix_chart_axes_puts_the_scatter_x_axis_at_the_bottom(self) -> None:
        from openpyxl.chart import ScatterChart

        chart = self._fixed_chart(ScatterChart())

        assert chart.x_axis.axPos == "b"
        assert chart.y_axis.axPos == "l"

    def test_fix_chart_axes_keeps_a_right_hand_value_axis(self) -> None:
        from openpyxl.chart import LineChart

        chart = LineChart()
        chart.y_axis.axPos = "r"

        self._fixed_chart(chart)

        assert chart.y_axis.axPos == "r"

    def test_fix_chart_axes_keeps_a_secondary_axis(self) -> None:
        from openpyxl.chart import LineChart

        chart = LineChart()
        chart.y_axis.crosses = "max"
        chart.y_axis.axPos = "r"

        self._fixed_chart(chart)

        assert chart.y_axis.axPos == "r"

    def test_fix_chart_axes_keeps_a_top_category_axis(self) -> None:
        from openpyxl.chart import LineChart

        chart = LineChart()
        chart.x_axis.axPos = "t"

        self._fixed_chart(chart)

        assert chart.x_axis.axPos == "t"

    def test_fix_chart_axes_still_repairs_a_defaulted_line_chart(self) -> None:
        from openpyxl.chart import LineChart

        chart = LineChart()
        chart.x_axis.axPos = "l"

        self._fixed_chart(chart)

        assert chart.x_axis.axPos == "b"
        assert chart.y_axis.axPos == "l"

    def test_fix_chart_axes_horizontal_bar_chart(self) -> None:
        from openpyxl.chart import BarChart

        chart = BarChart()
        chart.barDir = "bar"

        self._fixed_chart(chart)

        assert chart.x_axis.axPos == "l"
        assert chart.y_axis.axPos == "b"
