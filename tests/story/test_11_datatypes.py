"""Chapter 11 — Data Types.
=========================

fillxl handles a wide range of Python data types when writing to Excel.
This chapter documents what types are supported, how they map to Excel
types, and how edge cases like null values, timezone-aware datetimes,
and unsupported complex types are handled.

All tests in this chapter use the :func:`fillxl.fill` entry point and
verify results by reading back with openpyxl.
"""

from __future__ import annotations

from datetime import UTC, date, datetime, time
from decimal import Decimal
from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import FillXLConfig, fill

if TYPE_CHECKING:
    from pathlib import Path

_PERMISSIVE = FillXLConfig(column_validation_strategy="none")


# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_dtype_template(path: Path) -> None:
    """Workbook with typed named cells and a small table for dtype tests."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Types"

    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.table import Table, TableColumn  # ty:ignore[unresolved-import]

    # Named cells for scalar dtype tests
    ws["A1"] = ""  # str_cell
    ws["B1"] = 0  # int_cell
    ws["C1"] = 0.0  # float_cell
    ws["D1"] = False  # bool_cell
    ws["E1"] = None  # none_cell  # ty:ignore[invalid-assignment]
    ws["F1"] = None  # datetime_cell  # ty:ignore[invalid-assignment]
    ws["G1"] = None  # date_cell  # ty:ignore[invalid-assignment]
    ws["H1"] = None  # time_cell  # ty:ignore[invalid-assignment]
    ws["I1"] = 0  # decimal_cell
    ws["J1"] = None  # nan_cell  # ty:ignore[invalid-assignment]

    for col_letter in "ABCDEFGHIJ":
        wb.defined_names.add(
            DefinedName(f"nc_{col_letter.lower()}", attr_text=f"Types!${col_letter}$1"),
        )

    # Table for tabular dtype tests
    ws["A3"] = "Name"
    ws["B3"] = "Value"
    ws["A4"] = "placeholder"
    ws["B4"] = 0
    table = Table(
        displayName="DtypeTable",
        ref="A3:B4",
        tableColumns=[
            TableColumn(id=1, name="Name"),
            TableColumn(id=2, name="Value"),
        ],
    )
    ws.add_table(table)

    wb.save(path)
    wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 11.1 Nullable values
# ═══════════════════════════════════════════════════════════════════════════════


class TestNullableValues:
    """None and NaN-like sentinels become blank cells in Excel."""

    def test_none_becomes_blank_in_a_named_cell(
        self,
        output_path: Path,
    ) -> None:
        """Writing None to a named cell leaves it blank."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_e": None},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["E1"].value is None
        finally:
            wb.close()

    def test_none_becomes_blank_in_a_table(
        self,
        output_path: Path,
    ) -> None:
        """A table cell with None is written as blank."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"DtypeTable": [{"Name": None, "Value": 42}]},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["A4"].value is None
            assert wb["Types"]["B4"].value == 42
        finally:
            wb.close()

    def test_float_nan_becomes_blank_in_a_named_cell(
        self,
        output_path: Path,
    ) -> None:
        """float('nan') is normalised to None for Excel."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_j": float("nan")},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["J1"].value is None
        finally:
            wb.close()

    def test_float_nan_becomes_blank_in_a_table(
        self,
        output_path: Path,
    ) -> None:
        """A table cell with float('nan') is written as blank."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"DtypeTable": [{"Name": "NaN", "Value": float("nan")}]},
            config=_PERMISSIVE,
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["B4"].value is None
        finally:
            wb.close()

    def test_pandas_na_becomes_blank_in_a_table(
        self,
        output_path: Path,
    ) -> None:
        """pd.NA in a DataFrame is written as blank."""
        pd = pytest.importorskip("pandas")
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        df = pd.DataFrame({"Name": ["pd.NA"], "Value": [pd.NA]})

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"DtypeTable": df},
            config=_PERMISSIVE,
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["B4"].value is None
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 11.2 Dates and times
# ═══════════════════════════════════════════════════════════════════════════════


class TestDatesAndTimes:
    """Datetime, date, time, and timezone handling."""

    def test_datetime_is_written_as_an_excel_date(
        self,
        output_path: Path,
    ) -> None:
        """A Python datetime becomes an Excel datetime serial."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        dt_val = datetime(2025, 12, 25, 14, 30, 0)

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_f": dt_val},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["F1"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025  # type: ignore[attr-defined]
            assert val.month == 12  # type: ignore[attr-defined]
            assert val.day == 25  # type: ignore[attr-defined]
            assert val.hour == 14  # type: ignore[attr-defined]
            assert val.minute == 30  # type: ignore[attr-defined]
        finally:
            wb.close()

    def test_date_is_written_as_an_excel_date(
        self,
        output_path: Path,
    ) -> None:
        """A Python date is preserved."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        d_val = date(2025, 6, 15)

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_g": d_val},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["G1"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025  # type: ignore[attr-defined]
            assert val.month == 6  # type: ignore[attr-defined]
            assert val.day == 15  # type: ignore[attr-defined]
        finally:
            wb.close()

    def test_time_is_written_as_an_excel_time(
        self,
        output_path: Path,
    ) -> None:
        """A Python time is preserved."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        t_val = time(9, 15, 30)

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_h": t_val},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["H1"].value
            assert val is not None
            assert hasattr(val, "hour")
            assert val.hour == 9  # type: ignore[attr-defined]
            assert val.minute == 15  # type: ignore[attr-defined]
            assert val.second == 30  # type: ignore[attr-defined]
        finally:
            wb.close()

    def test_timezone_aware_datetime_is_normalised(
        self,
        output_path: Path,
    ) -> None:
        """A tz-aware datetime has its timezone stripped (wall-clock preserved)."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        aware = datetime(2025, 6, 15, 14, 30, tzinfo=UTC)

        with pytest.warns(UserWarning, match="Timezone-aware"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                values={"nc_f": aware},
            )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["F1"].value
            assert val is not None
            # tzinfo should be stripped
            assert getattr(val, "tzinfo", None) is None
            # Wall-clock time preserved
            assert val.hour == 14  # type: ignore[attr-defined]
            assert val.minute == 30  # type: ignore[attr-defined]
        finally:
            wb.close()

    def test_timezone_aware_datetime_triggers_a_warning(
        self,
        output_path: Path,
    ) -> None:
        """The first tz-aware datetime in a fill session issues a UserWarning."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        aware = datetime(2025, 1, 1, 0, 0, tzinfo=UTC)

        with pytest.warns(UserWarning, match="Timezone-aware"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                values={"nc_f": aware},
            )


# ═══════════════════════════════════════════════════════════════════════════════
# 11.3 Numeric types
# ═══════════════════════════════════════════════════════════════════════════════


class TestNumericTypes:
    """Integer, float, Decimal, and special numeric values."""

    def test_decimal_is_converted_to_float(
        self,
        output_path: Path,
    ) -> None:
        """Decimal values become floats in Excel (no native Decimal support)."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_i": Decimal("12.34")},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["I1"].value
            assert isinstance(val, float)
            assert abs(val - 12.34) < 1e-10
        finally:
            wb.close()

    def test_large_integers_are_preserved(
        self,
        output_path: Path,
    ) -> None:
        """Integers larger than 32-bit are preserved."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        large = 2**53 - 1  # max safe integer for float but int in Python

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_b": large},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["B1"].value == large
        finally:
            wb.close()

    def test_negative_numbers_are_preserved(
        self,
        output_path: Path,
    ) -> None:
        """Negative integers and floats are written correctly."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_b": -42, "nc_c": -3.14},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["B1"].value == -42
            assert wb["Types"]["C1"].value == -3.14
        finally:
            wb.close()

    def test_boolean_values_are_written_as_bool(
        self,
        output_path: Path,
    ) -> None:
        """True and False are written as Excel booleans."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_d": True},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["D1"].value is True
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 11.4 String and text values
# ═══════════════════════════════════════════════════════════════════════════════


class TestStringValues:
    """Text values, Unicode, and formula strings."""

    def test_unicode_text_is_preserved(
        self,
        output_path: Path,
    ) -> None:
        """Unicode characters (including emoji) are written correctly."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        text = "Hello, 世界 🌍"

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_a": text},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Types"]["A1"].value == text
        finally:
            wb.close()

    def test_formula_strings_are_written_as_excel_formulas(
        self,
        output_path: Path,
    ) -> None:
        """A string starting with '=' becomes a formula."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={"nc_c": 10, "nc_i": 20},
        )

        # Write a formula via table to test formula formatting
        from openpyxl.workbook.defined_name import DefinedName

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Formulas"
        ws["A1"] = 0
        ws["A2"] = 0
        wb.defined_names.add(DefinedName("formula_result", attr_text="Formulas!$A$1"))
        tmpl2 = output_path.with_name("formula_tmpl.xlsx")
        wb.save(tmpl2)
        wb.close()

        # String starting with = should be written as formula
        fill(
            template=tmpl2,
            output=output_path,
            values={"formula_result": "=1+1"},
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            val = wb2["Formulas"]["A1"].value
            assert val is not None
            assert str(val).startswith("="), f"Expected formula, got {val!r}"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 11.5 Unsupported complex types
# ═══════════════════════════════════════════════════════════════════════════════


class TestUnsupportedTypes:
    """Complex types that cannot be written to Excel are rejected."""

    def test_polars_list_column_is_converted_to_string(
        self,
        output_path: Path,
    ) -> None:
        """A polars DataFrame with a List column is written as string."""
        pl = pytest.importorskip("polars")
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        df = pl.DataFrame({"Name": ["list"], "Value": [[1, 2, 3]]})

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"DtypeTable": df},
            config=_PERMISSIVE,
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["B4"].value
            # List values are converted to their string representation
            assert isinstance(val, str)
            assert "1" in val
            assert "2" in val
            assert "3" in val
        finally:
            wb.close()

    def test_polars_struct_column_is_converted_to_string(
        self,
        output_path: Path,
    ) -> None:
        """A polars DataFrame with a Struct column is written as string."""
        pl = pytest.importorskip("polars")
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        df = pl.DataFrame({"Name": ["struct"], "Value": [{"a": 1, "b": 2}]})

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"DtypeTable": df},
            config=_PERMISSIVE,
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["B4"].value
            # Struct values are converted to their string representation
            assert isinstance(val, str)
            assert "a" in val
            assert "1" in val
        finally:
            wb.close()

    def test_nested_dict_in_list_of_dict_is_converted_to_string(
        self,
        output_path: Path,
    ) -> None:
        """A nested dict value in list[dict] is converted to its string repr."""
        _make_dtype_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"DtypeTable": [{"Name": "nested", "Value": {"key": "val"}}]},
            config=_PERMISSIVE,
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            val = wb["Types"]["B4"].value
            assert isinstance(val, str)
            assert "key" in val
        finally:
            wb.close()
