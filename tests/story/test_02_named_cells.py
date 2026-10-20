"""Chapter 2 — Filling Named Cells.
================================

Excel lets you assign **names** to cells (via the Name Manager or by
typing a name into the box left of the formula bar).  Instead of
referring to ``Sheet1!$A$1``, you can call it ``"revenue"`` or
``"report_date"``.

fillxl uses these names as the bridge between Python and Excel.  You
pass a ``values`` dictionary to :func:`fillxl.fill` where each key is a
named-cell name and each value is the scalar to write.  fillxl finds
the cell behind the name and replaces its content — no cell coordinates
needed.

This chapter covers basic named-cell filling, error handling, formatting
preservation, formula strings, merged cells, and scalar data types across backends.
"""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, cast

import openpyxl
import pytest

from fillxl import FillXLConfig, fill
from fillxl._cells import populate_named_cells
from fillxl._types import Engine

if TYPE_CHECKING:
    from pathlib import Path

# ═══════════════════════════════════════════════════════════════════════════════
# 2.1 Basic named cell filling
# ═══════════════════════════════════════════════════════════════════════════════


class TestFillingNamedCells:
    """Filling single and multiple named cells with scalar values."""

    def test_a_named_cell_is_updated_with_the_given_value(
        self,
        output_path: Path,
        named_cell_template: Path,
    ) -> None:
        """Pass a value for 'my_cell' and it replaces the old content."""
        fill(
            template=named_cell_template,
            output=output_path,
            values={"my_cell": "world"},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Sheet1"]["A1"].value == "world"
        finally:
            wb.close()

    def test_multiple_named_cells_can_be_filled_in_one_call(
        self,
        output_path: Path,
    ) -> None:
        """You can fill as many named cells as you want at once."""
        # Build a template with three named cells
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"
        ws["A1"] = "old_a"
        ws["B1"] = "old_b"
        ws["C1"] = "old_c"

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("cell_a", attr_text="Report!$A$1"))
        wb.defined_names.add(DefinedName("cell_b", attr_text="Report!$B$1"))
        wb.defined_names.add(DefinedName("cell_c", attr_text="Report!$C$1"))

        tmpl = output_path.with_name("multi_cell_template.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"cell_a": "Alpha", "cell_b": "Bravo", "cell_c": "Charlie"},
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Report"]
            assert ws2["A1"].value == "Alpha"
            assert ws2["B1"].value == "Bravo"
            assert ws2["C1"].value == "Charlie"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 2.2 Error handling for named cells
# ═══════════════════════════════════════════════════════════════════════════════


class TestNamedCellErrors:
    """Error handling for missing names, case sensitivity, and edge cases."""

    def test_a_keyerror_is_raised_when_the_named_cell_does_not_exist(
        self,
        output_path: Path,
        named_cell_template: Path,
    ) -> None:
        """If you reference a name that isn't in the workbook, fill() tells you."""
        with pytest.raises(KeyError, match="not found"):
            fill(
                template=named_cell_template,
                output=output_path,
                values={"nonexistent": "value"},
            )

    def test_named_cell_names_are_case_sensitive(
        self,
        output_path: Path,
        named_cell_template: Path,
    ) -> None:
        """'My_Cell' and 'my_cell' are different names."""
        with pytest.raises(KeyError):
            fill(
                template=named_cell_template,
                output=output_path,
                values={"MY_CELL": "value"},
            )

    def test_unknown_named_cells_are_skipped_when_allow_missing_is_true(
        self,
        named_cell_template: Path,
    ) -> None:
        """The internal function can be told to skip missing names gracefully."""
        wb = openpyxl.load_workbook(named_cell_template)
        try:
            # Should not raise
            populate_named_cells(
                wb,
                {"nonexistent": "value", "my_cell": "updated"},
                allow_missing=True,
            )
            assert wb["Sheet1"]["A1"].value == "updated"
        finally:
            wb.close()

    def test_a_warning_is_logged_when_a_named_cell_has_no_destinations(
        self,
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        """A defined name holding a constant (not a range) has no cell destination."""
        from openpyxl.workbook.defined_name import DefinedName

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "ok"

        # A named constant (not a range reference)
        dn = DefinedName("pi_value", attr_text="3.14159")
        wb.defined_names.add(dn)

        caplog.set_level("WARNING")
        populate_named_cells(wb, {"pi_value": "ignored"})

        assert any("no destinations" in r.message for r in caplog.records)

    def test_a_warning_is_logged_when_a_named_cell_points_to_multiple_ranges(
        self,
        output_path: Path,
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        """Only the first destination is updated; the rest are skipped with a warning."""
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Sheet1"
        ws1["A1"] = "first"
        ws2 = wb.create_sheet("Sheet2")
        ws2["B2"] = "second"

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("multi", attr_text="Sheet1!$A$1,Sheet2!$B$2"))

        tmpl = output_path.with_name("multi_dest.xlsx")
        wb.save(tmpl)
        wb.close()

        with caplog.at_level("WARNING", logger="fillxl._cells"):
            fill(template=tmpl, output=output_path, values={"multi": "UPDATED"})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            assert wb2["Sheet1"]["A1"].value == "UPDATED"
            assert wb2["Sheet2"]["B2"].value == "second"
        finally:
            wb2.close()
        assert "multiple destinations" in caplog.text


# ═══════════════════════════════════════════════════════════════════════════════
# 2.3 Named cell formatting preservation
# ═══════════════════════════════════════════════════════════════════════════════


class TestNamedCellFormatting:
    """Existing cell formatting is preserved when values are written."""

    def test_the_cell_number_format_is_preserved_after_filling(
        self,
        output_path: Path,
    ) -> None:
        """A cell formatted as percentage stays as percentage."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = 0.5
        ws["A1"].number_format = "0.00%"

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("pct", attr_text="Data!$A$1"))

        tmpl = output_path.with_name("pct_template.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(template=tmpl, output=output_path, values={"pct": 0.75})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Data"]
            assert ws2["A1"].value == 0.75
            assert ws2["A1"].number_format == "0.00%"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 2.4 Formula strings in named cells
# ═══════════════════════════════════════════════════════════════════════════════


class TestFormulaStringsInNamedCells:
    """Strings starting with '=' are written as Excel formulas when unescaped."""

    def test_a_formula_string_is_written_as_an_excel_formula(
        self,
        output_path: Path,
    ) -> None:
        """Pass '=1+2' with escaping disabled and the cell stores a formula."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = 10
        ws["B1"] = 20

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("total", attr_text="Data!$C$1"))
        tmpl = output_path.with_name("formula_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"total": "=A1+B1"},
            config=FillXLConfig(escape_formulas=False),
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            val = wb2["Data"]["C1"].value
            assert isinstance(val, str) and val.startswith("="), (
                f"Expected a formula string, got {val!r}"
            )
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 2.5 Merged cells and multi-cell ranges
# ═══════════════════════════════════════════════════════════════════════════════


class TestMergedAndMultiCellRanges:
    """Named cells that point into merged or multi-cell ranges."""

    def test_writing_to_a_merged_cell_range_writes_to_the_top_left_cell(
        self,
        output_path: Path,
    ) -> None:
        """A named cell pointing at the top-left of a merged range writes
        successfully — the value lands in the anchor cell (B8).
        """
        from openpyxl.workbook.defined_name import DefinedName

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"
        ws["B8"] = "old merged value"
        ws.merge_cells("B8:C9")

        wb.defined_names.add(DefinedName("merged_cell", attr_text="Report!$B$8"))
        tmpl = output_path.with_name("merged_template.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(template=tmpl, output=output_path, values={"merged_cell": "Updated"})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            assert wb2["Report"]["B8"].value == "Updated"
        finally:
            wb2.close()

    def test_writing_to_a_multi_cell_range_raises_value_error(
        self,
        output_path: Path,
    ) -> None:
        """A named range spanning multiple cells (e.g. E2:F3) cannot be
        filled with a single scalar — fill() raises ValueError.
        """
        from openpyxl.workbook.defined_name import DefinedName

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"
        ws["E2"] = "cell 1"
        ws["F2"] = "cell 2"
        ws["E3"] = "cell 3"
        ws["F3"] = "cell 4"

        wb.defined_names.add(DefinedName("multi_range", attr_text="Report!$E$2:$F$3"))
        tmpl = output_path.with_name("multi_range_template.xlsx")
        wb.save(tmpl)
        wb.close()

        with pytest.raises(ValueError, match=r"multi.cell|range|single cell"):
            fill(
                template=tmpl,
                output=output_path,
                values={"multi_range": "cannot write to a range"},
            )


# ═══════════════════════════════════════════════════════════════════════════════
# 2.6 Scalar Data Types Across Backends
# ═══════════════════════════════════════════════════════════════════════════════


class TestScalarDataTypes:
    """Scalar types (str, int, float, bool, None, datetime, timedelta) work
    consistently regardless of which backend is used.  Custom objects are
    converted to their ``str()`` representation (via ``__str__``).
    """

    @pytest.mark.parametrize(
        "engine",
        [
            "openpyxl",
            pytest.param("win32com", marks=pytest.mark.win32com),
        ],
    )
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            ("Hello", "Hello"),
            (42, 42),
            (3.14, 3.14),
            (True, True),
            (None, None),
        ],
        ids=["str", "int", "float", "bool", "None"],
    )
    def test_scalar_types_are_written_correctly(
        self,
        output_path: Path,
        engine: str,
        value: object,
        expected: object,
    ) -> None:
        """str, int, float, bool, and None are written correctly by every backend."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "placeholder"

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("val", attr_text="Data!$A$1"))
        tmpl = output_path.with_name("types_template.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(template=tmpl, output=output_path, values={"val": value}, engine=cast(Engine, engine))

        wb2 = openpyxl.load_workbook(output_path)
        try:
            assert wb2["Data"]["A1"].value == expected
        finally:
            wb2.close()

    @pytest.mark.parametrize(
        "engine",
        [
            "openpyxl",
            pytest.param("win32com", marks=pytest.mark.win32com),
        ],
    )
    def test_datetime_is_written_as_excel_date(
        self,
        output_path: Path,
        engine: str,
    ) -> None:
        """Datetime objects are stored as Excel date-time values by every backend."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = ""

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("date_cell", attr_text="Data!$A$1"))
        tmpl = output_path.with_name("date_template.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"date_cell": datetime(2025, 6, 15)},
            engine=cast(Engine, engine),
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            val = wb2["Data"]["A1"].value
            assert hasattr(val, "year")
            assert val.year == 2025
        finally:
            wb2.close()

    @pytest.mark.parametrize(
        "engine",
        [
            "openpyxl",
            pytest.param("win32com", marks=pytest.mark.win32com),
        ],
    )
    def test_timedelta_is_written_as_excel_duration(
        self,
        output_path: Path,
        engine: str,
    ) -> None:
        """Timedelta values are written as floats (fraction of a day).

        Both backends normalise timedelta to a float (total_seconds/86400)
        for cross-backend consistency.  The win32com backend has always
        done this; openpyxl now matches.
        """
        from datetime import timedelta

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = ""

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("dur_cell", attr_text="Data!$A$1"))
        tmpl = output_path.with_name("td_template.xlsx")
        wb.save(tmpl)
        wb.close()

        td = timedelta(hours=12, minutes=30)  # 12.5 hours
        expected_float = td.total_seconds() / 86400.0  # 0.520833...

        fill(
            template=tmpl,
            output=output_path,
            values={"dur_cell": td},
            engine=cast(Engine, engine),
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            val = wb2["Data"]["A1"].value
            # Both backends now store timedelta as float (fraction of day)
            assert isinstance(val, (int, float)), f"Expected float, got {type(val).__name__}"
            assert abs(float(val) - expected_float) < 1e-9
        finally:
            wb2.close()

    @pytest.mark.parametrize(
        "engine",
        [
            "openpyxl",
            pytest.param("win32com", marks=pytest.mark.win32com),
        ],
    )
    def test_custom_object_is_converted_to_string_via_str(
        self,
        output_path: Path,
        engine: str,
    ) -> None:
        """A custom object with a __str__ method is converted to its string
        representation — str(obj) is called, which invokes __str__.
        """

        class Greeting:
            def __str__(self) -> str:
                return "Hello from Greeting!"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = ""

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("custom", attr_text="Data!$A$1"))
        tmpl = output_path.with_name("custom_template.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"custom": Greeting()},
            engine=cast(Engine, engine),
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            val = wb2["Data"]["A1"].value
            assert val == "Hello from Greeting!"
        finally:
            wb2.close()
