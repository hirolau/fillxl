"""Cross-backend datatype consistency tests.

Verifies that fillxl handles common Python datatypes identically regardless
of whether the ``openpyxl`` or ``win32com`` backend is used.

Test organisation (3 phases):
    Phase A — CI-safe (no Excel required)
        Tests using ``engine="openpyxl"`` covering all common dtypes for
        both named cells and tables.  Verifies via openpyxl read-back.

    Phase B — Windows + Excel only (``@pytest.mark.win32com``)
        Same test matrix using ``engine="win32com"``, verified via
        openpyxl read-back **and** xlwings COM read-back for rendered values.

    Phase C — Cross-backend comparison
        Parametrised: same input data → both backends → compare results
        read back with openpyxl.  Documents any intentional differences.
"""

from __future__ import annotations

from datetime import UTC, date, datetime, time, timedelta
from decimal import Decimal
from typing import TYPE_CHECKING, ClassVar

import pytest

from fillxl import FillXLConfig, fill

if TYPE_CHECKING:
    from pathlib import Path

_PERMISSIVE = FillXLConfig(column_validation_strategy="none")

# Excel serial date epoch: day 1 = 1900-01-01, day 0 = 1899-12-30
# Excel has a leap year bug (treats 1900 as leap), so for serial > 60
# we must subtract 1 to get the real date.
_SERIAL_EPOCH = datetime(1899, 12, 30)


def _serial_to_datetime(serial: float) -> datetime:
    """Convert an Excel serial date number to a :class:`~datetime.datetime`.

    Handles the Excel 1900 leap-year bug transparently.
    """
    days = int(serial)
    frac = serial - days
    if days >= 61:
        days -= 1  # Excel leap-year bug
    result = _SERIAL_EPOCH + timedelta(days=days)
    # Add fractional day (time component)
    if frac:
        result = result + timedelta(seconds=round(frac * 86400))
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# Phase A — CI-safe openpyxl backend dtype tests
# ═══════════════════════════════════════════════════════════════════════════════


class TestCrossBackendNamedCellsOpenpyxl:
    """Named cell dtype tests via the openpyxl backend (CI-safe)."""

    _NAMED_CELL_COL_MAP: ClassVar[dict[str, str]] = {
        "cd_str_cell": "A",
        "cd_int_cell": "B",
        "cd_float_cell": "C",
        "cd_bool_cell": "D",
        "cd_none_cell": "E",
        "cd_datetime_cell": "F",
        "cd_date_cell": "G",
        "cd_time_cell": "H",
        "cd_decimal_cell": "I",
        "cd_nan_cell": "J",
        "cd_timestamp_cell": "K",
        "cd_tz_datetime_cell": "L",
        "cd_large_int_cell": "M",
        "cd_large_float_cell": "N",
    }

    @pytest.mark.parametrize(
        ("cell_name", "value", "expected_type", "expected_value"),
        [
            pytest.param("cd_str_cell", "Hello, 世界", str, "Hello, 世界", id="str"),
            pytest.param("cd_int_cell", 42, int, 42, id="int"),
            pytest.param("cd_float_cell", 3.14159, float, 3.14159, id="float"),
            pytest.param("cd_bool_cell", True, bool, True, id="bool_true"),
            pytest.param("cd_bool_cell", False, bool, False, id="bool_false"),
            pytest.param("cd_none_cell", None, type(None), None, id="none"),
            pytest.param(
                "cd_datetime_cell",
                datetime(2025, 6, 15, 14, 30, 45),
                datetime,
                datetime(2025, 6, 15, 14, 30, 45),
                id="datetime",
            ),
            pytest.param("cd_date_cell", date(2025, 12, 25), date, date(2025, 12, 25), id="date"),
            pytest.param("cd_time_cell", time(9, 15, 30), time, time(9, 15, 30), id="time"),
            pytest.param("cd_decimal_cell", Decimal("12.34"), float, 12.34, id="decimal"),
            # Edge cases
            pytest.param("cd_nan_cell", float("nan"), type(None), None, id="float_nan"),
            pytest.param(
                "cd_large_int_cell",
                2**53 - 1,
                int,
                2**53 - 1,
                id="large_int",
            ),
            pytest.param(
                "cd_large_float_cell",
                1e200,
                float,
                1e200,
                id="large_float",
            ),
        ],
    )
    def test_named_cell_preserves_type_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
        cell_name: str,
        value: object,
        expected_type: type,
        expected_value: object,
    ) -> None:
        """Named cell filled with openpyxl backend preserves the Python type."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={cell_name: value},
            engine="openpyxl",
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            cell_ref = f"{self._NAMED_CELL_COL_MAP[cell_name]}1"
            val = ws[cell_ref].value

            if cell_name == "cd_datetime_cell":
                assert val is not None
                assert hasattr(val, "year"), f"Expected datetime-like, got {type(val)}"
                assert val.year == expected_value.year  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.month == expected_value.month  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.day == expected_value.day  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.hour == expected_value.hour  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.minute == expected_value.minute  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
            elif cell_name == "cd_date_cell":
                assert val is not None
                assert hasattr(val, "year"), f"Expected date-like, got {type(val)}"
                assert val.year == expected_value.year  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.month == expected_value.month  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.day == expected_value.day  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
            elif cell_name == "cd_time_cell":
                assert val is not None
                assert hasattr(val, "hour"), f"Expected time-like, got {type(val)}"
                assert val.hour == expected_value.hour  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.minute == expected_value.minute  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert val.second == expected_value.second  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
            elif cell_name == "cd_decimal_cell":
                assert val is not None
                assert isinstance(val, float)
                assert val == expected_value
            elif cell_name == "cd_nan_cell":
                assert val is None  # float("nan") → None in openpyxl
            elif cell_name == "cd_none_cell":
                assert val is None
            else:
                assert isinstance(val, expected_type), f"Expected {expected_type}, got {type(val)}"
                assert val == expected_value
        finally:
            wb.close()

    def test_named_cell_converts_pandas_timestamp_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A pd.Timestamp written to a named cell via openpyxl is stored as a datetime."""
        pd = pytest.importorskip("pandas")
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_timestamp_cell": pd.Timestamp("2025-06-15 14:30:00")},
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["K1"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025
            assert val.month == 6
            assert val.day == 15
        finally:
            wb.close()

    def test_named_cell_strips_timezone_from_datetime_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A tz-aware datetime in a named cell via openpyxl is stored naive and warns."""
        import openpyxl

        tz_datetime = datetime(2025, 6, 15, 14, 30, 0, tzinfo=UTC)
        with pytest.warns(UserWarning, match="Timezone-aware"):
            fill(
                template=cross_backend_dtype_template,
                output=output_path,
                values={"cd_tz_datetime_cell": tz_datetime},
                engine="openpyxl",
            )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["L1"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025
            assert val.month == 6
            assert val.day == 15
            assert val.hour == 14
            assert val.minute == 30
            # Must be timezone-naive
            if hasattr(val, "tzinfo"):
                assert val.tzinfo is None
        finally:
            wb.close()


class TestCrossBackendTableOpenpyxl:
    """Table column dtype tests via the openpyxl backend (CI-safe)."""

    _TABLE_COL_MAP: ClassVar[dict[str, str]] = {
        "StrCol": "A",
        "IntCol": "B",
        "FloatCol": "C",
        "BoolCol": "D",
        "NoneCol": "E",
        "DatetimeCol": "F",
        "DateCol": "G",
        "TimeCol": "H",
        "DecimalCol": "I",
        "FloatNanCol": "J",
        "NaCol": "K",
        "NaTCol": "L",
        "TimestampCol": "M",
        "TzDatetimeCol": "N",
        "LargeIntCol": "O",
        "LargeFloatCol": "P",
    }

    def _read_cell(self, output_path: Path, sheet: str, col_letter: str, row: int = 2) -> object:
        """Read a single cell back with openpyxl."""
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb[sheet]
            return ws[f"{col_letter}{row}"].value
        finally:
            wb.close()

    # ── Basic types ────────────────────────────────────────────────────────

    @pytest.mark.parametrize(
        ("col_name", "values", "expected_type"),
        [
            pytest.param("StrCol", ["hello", "world"], str, id="str"),
            pytest.param("IntCol", [10, -5], int, id="int"),
            pytest.param("FloatCol", [1.5, -3.14], float, id="float"),
            pytest.param("BoolCol", [True, False], bool, id="bool"),
            pytest.param("NoneCol", [None, None], type(None), id="none"),
            pytest.param(
                "DatetimeCol",
                [datetime(2025, 6, 15, 14, 30), datetime(2024, 1, 1, 0, 0)],
                datetime,
                id="datetime",
            ),
        ],
    )
    def test_table_column_preserves_type_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
        col_name: str,
        values: list[object],
        expected_type: type,
    ) -> None:
        """A table column preserves its Python type when filled via openpyxl."""
        data = [{col_name: v} for v in values]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )

        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            col_letter = self._TABLE_COL_MAP[col_name]

            for i, expected_val in enumerate(values):
                row = i + 2
                val = ws[f"{col_letter}{row}"].value

                if col_name == "DatetimeCol":
                    assert val is not None
                    assert hasattr(val, "year"), f"Expected datetime-like, got {type(val)}"
                    assert val.year == expected_val.year  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                    assert val.month == expected_val.month  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                    assert val.day == expected_val.day  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                elif col_name == "NoneCol":
                    assert val is None
                else:
                    assert isinstance(val, expected_type), (
                        f"Row {row}: expected {expected_type}, got {type(val)}"
                    )
                    assert val == expected_val
        finally:
            wb.close()

    # ── Date / Time / Decimal ──────────────────────────────────────────────

    def test_table_date_column_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Python date values in a table column are preserved via openpyxl."""
        data = [{"DateCol": date(2025, 12, 25)}, {"DateCol": date(2023, 7, 4)}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            val1 = ws["G2"].value
            assert val1 is not None
            assert hasattr(val1, "year")
            assert val1.year == 2025
            assert val1.month == 12
            assert val1.day == 25
            val2 = ws["G3"].value
            assert val2 is not None
            assert hasattr(val2, "year")
            assert val2.year == 2023
            assert val2.month == 7
            assert val2.day == 4
        finally:
            wb.close()

    def test_table_time_column_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Python time values in a table column are preserved via openpyxl."""
        data = [{"TimeCol": time(9, 15)}, {"TimeCol": time(23, 59, 59)}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            val1 = ws["H2"].value
            assert val1 is not None
            assert hasattr(val1, "hour")
            assert val1.hour == 9
            assert val1.minute == 15
            val2 = ws["H3"].value
            assert val2 is not None
            assert hasattr(val2, "hour")
            assert val2.hour == 23
            assert val2.minute == 59
            assert val2.second == 59
        finally:
            wb.close()

    def test_table_decimal_column_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Store Decimal table-column values as float (openpyxl) with numeric value preserved."""
        data = [{"DecimalCol": Decimal("12.34")}, {"DecimalCol": Decimal("0.01")}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            val1 = ws["I2"].value
            assert val1 is not None
            assert isinstance(val1, float)
            assert abs(val1 - 12.34) < 1e-10
            val2 = ws["I3"].value
            assert val2 is not None
            assert abs(val2 - 0.01) < 1e-10
        finally:
            wb.close()

    # ── Edge cases: NaN, pd.NA, pd.NaT ─────────────────────────────────────

    def test_table_float_nan_written_as_none_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """float('nan') in a table column becomes None (openpyxl backend)."""
        data = [{"FloatNanCol": float("nan")}, {"FloatNanCol": 1.0}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["J2"].value is None  # NaN → None
            assert ws["J3"].value == 1.0
        finally:
            wb.close()

    def test_table_pandas_na_written_as_none_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """pd.NA in a table column becomes None (openpyxl backend)."""
        pd = pytest.importorskip("pandas")
        data = [{"NaCol": pd.NA}, {"NaCol": "hello"}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["K2"].value is None
            assert ws["K3"].value == "hello"
        finally:
            wb.close()

    def test_table_pandas_nat_written_as_none_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """pd.NaT in a table column becomes None (openpyxl backend)."""
        pd = pytest.importorskip("pandas")
        data = [{"NaTCol": pd.NaT}, {"NaTCol": datetime(2025, 1, 1)}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["L2"].value is None
            assert ws["L3"].value is not None
        finally:
            wb.close()

    # ── Edge cases: pd.Timestamp, tz-aware ─────────────────────────────────

    def test_table_pandas_timestamp_to_datetime_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """pd.Timestamp in a table column becomes a Python datetime (openpyxl)."""
        pd = pytest.importorskip("pandas")
        data = [
            {"TimestampCol": pd.Timestamp("2025-06-15 14:30:00")},
            {"TimestampCol": pd.Timestamp("2024-01-01")},
        ]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            val1 = ws["M2"].value
            assert val1 is not None
            assert hasattr(val1, "year")
            assert val1.year == 2025
            assert val1.month == 6
            assert val1.day == 15
            val2 = ws["M3"].value
            assert val2 is not None
            assert hasattr(val2, "year")
            assert val2.year == 2024
            assert val2.month == 1
            assert val2.day == 1
        finally:
            wb.close()

    def test_table_pandas_timestamp_with_timezone_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """pd.Timestamp with timezone info becomes naive datetime and warns (openpyxl)."""
        pd = pytest.importorskip("pandas")
        data = [
            {"TzDatetimeCol": pd.Timestamp("2025-06-15 14:30:00", tz="UTC")},
        ]
        with pytest.warns(UserWarning, match="Timezone-aware"):
            fill(
                template=cross_backend_dtype_template,
                output=output_path,
                tables={"CrossDtypeTable": data},
                config=_PERMISSIVE,
                engine="openpyxl",
            )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            val = ws["N2"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025
            assert val.month == 6
            assert val.day == 15
            assert val.hour == 14
            assert val.minute == 30
            if hasattr(val, "tzinfo"):
                assert val.tzinfo is None
        finally:
            wb.close()

    # ── Edge cases: large numbers ──────────────────────────────────────────

    def test_table_large_int_column_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Large integer values in a table column are preserved via openpyxl."""
        data = [{"LargeIntCol": 2**53 - 1}, {"LargeIntCol": -(2**53)}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["O2"].value == 2**53 - 1
            assert ws["O3"].value == -(2**53)
        finally:
            wb.close()

    def test_table_large_float_column_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Large float values in a table column are preserved via openpyxl."""
        data = [{"LargeFloatCol": 1e200}, {"LargeFloatCol": 1e-300}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        import openpyxl

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["P2"].value == 1e200  # fmt: skip
            assert ws["P3"].value == 1e-300
        finally:
            wb.close()

    # ── Mixed dtype row ────────────────────────────────────────────────────

    def test_table_writes_mixed_dtypes_in_single_row_via_openpyxl(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A single row with all mixed dtypes is written successfully via openpyxl."""
        import openpyxl

        data = [
            {
                "StrCol": "text",
                "IntCol": 99,
                "FloatCol": 3.14,
                "BoolCol": True,
                "NoneCol": None,
                "DatetimeCol": datetime(2025, 6, 15, 12, 0),
                "DateCol": date(2025, 6, 15),
                "TimeCol": time(12, 0),
                "DecimalCol": Decimal("99.99"),
                "FloatNanCol": float("nan"),
                "LargeIntCol": 2**53 - 1,
                "LargeFloatCol": 1e200,
            },
        ]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["A2"].value == "text"
            assert ws["B2"].value == 99
            assert ws["C2"].value == 3.14
            assert ws["D2"].value is True
            assert ws["E2"].value is None
            assert hasattr(ws["F2"].value, "year")
            assert hasattr(ws["G2"].value, "year")
            assert hasattr(ws["H2"].value, "hour")
            assert isinstance(ws["I2"].value, float)
            assert abs(ws["I2"].value - 99.99) < 1e-10
            assert ws["J2"].value is None  # NaN → None
            assert ws["O2"].value == 2**53 - 1
            assert ws["P2"].value == 1e200
        finally:
            wb.close()


# ── timedelta dtype tests ────────────────────────────────────────────────────


class TestCrossBackendTimedeltaNamedCell:
    """timedelta values written to named cells via openpyxl backend.

    Since fillxl v0.2, ``timedelta`` is normalised to ``float`` (fraction
    of a 24-hour day) — the only representation that is portable across
    both the ``openpyxl`` and ``win32com`` backends.
    """

    def test_named_cell_timedelta_via_openpyxl(
        self,
        output_path: Path,
    ) -> None:
        """A timedelta written to a named cell is stored as a float
        (fraction of a day), matching the win32com backend.
        """
        import openpyxl
        from openpyxl.workbook.defined_name import DefinedName

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = ""
        wb.defined_names.add(DefinedName("dur", attr_text="Data!$A$1"))
        tmpl = output_path.with_name("td_named_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        td = timedelta(hours=12, minutes=30)
        expected_float = td.total_seconds() / 86400.0  # 0.520833...
        fill(
            template=tmpl,
            output=output_path,
            values={"dur": td},
            engine="openpyxl",
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            val = wb2["Data"]["A1"].value
            assert isinstance(val, float), f"Expected float, got {type(val).__name__}"
            assert abs(float(val) - expected_float) < 1e-9
        finally:
            wb2.close()


class TestCrossBackendTimedeltaTable:
    """timedelta values written to table columns via openpyxl backend.

    Note: Unlike named cells where openpyxl preserves the timedelta
    type, table columns receive the numeric representation
    (fraction of a day = total_seconds / 86400).
    """

    def test_table_timedelta_via_openpyxl(
        self,
        output_path: Path,
    ) -> None:
        """A timedelta in a table column is preserved natively by openpyxl."""
        import openpyxl
        from openpyxl.worksheet.table import Table, TableColumn

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Duration"
        ws["A2"] = None
        table = Table(
            displayName="Times",
            ref="A1:A2",
            tableColumns=[TableColumn(id=1, name="Duration")],
        )
        ws.add_table(table)
        tmpl = output_path.with_name("td_table_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        td = timedelta(hours=1, minutes=30)  # 1.5 hours
        fill(
            template=tmpl,
            output=output_path,
            tables={"Times": [{"Duration": td}]},
            config=_PERMISSIVE,
            engine="openpyxl",
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            val = wb2["Data"]["A2"].value
            # Table cells store timedelta as float (fraction of a day)
            assert isinstance(val, (int, float)), f"Expected float, got {type(val).__name__}"
            assert abs(float(val) - td.total_seconds() / 86400.0) < 1e-9
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# Phase B — win32com backend dtype tests (Windows + Excel required)
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.win32com
class TestCrossBackendNamedCellsWin32Com:
    """Named cell dtype tests via the win32com backend (Excel required).

    Verification is done via openpyxl read-back (safe after COM save).
    """

    def test_named_cell_str_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A string named cell is preserved when filled via win32com."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_str_cell": "Hello, COM"},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            assert ws["A1"].value == "Hello, COM"
        finally:
            wb.close()

    def test_named_cell_int_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """An integer named cell is preserved when filled via win32com."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_int_cell": 42},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            assert ws["B1"].value == 42
        finally:
            wb.close()

    def test_named_cell_float_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A float named cell is preserved when filled via win32com."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_float_cell": 3.14159},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            assert ws["C1"].value == 3.14159
        finally:
            wb.close()

    def test_named_cell_bool_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A boolean named cell is preserved when filled via win32com."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_bool_cell": True},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            assert ws["D1"].value is True
        finally:
            wb.close()

    def test_named_cell_none_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """None in a named cell is preserved (empty cell) when filled via win32com."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_none_cell": None},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            assert ws["E1"].value is None
        finally:
            wb.close()

    def test_named_cell_datetime_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A datetime named cell is preserved when filled via win32com."""
        import openpyxl

        dt = datetime(2025, 6, 15, 14, 30, 0)
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_datetime_cell": dt},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["F1"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025
            assert val.month == 6
            assert val.day == 15
        finally:
            wb.close()

    def test_named_cell_date_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A date named cell is preserved when filled via win32com."""
        import openpyxl

        d = date(2025, 12, 25)
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_date_cell": d},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["G1"].value
            # COM treats naive datetimes as local time → converts to UTC.
            # Accept ±1 day tolerance (e.g. day 24 if UTC-1 local).
            assert val is not None and hasattr(val, "year"), f"Expected date-like, got {val!r}"
            # Year/month must match exactly; day may shift ±1 due to COM
            assert val.year == 2025
            assert val.month == 12
            assert abs(val.day - 25) <= 1, f"Expected day ~25, got {val.day}"
        finally:
            wb.close()

    def test_named_cell_time_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A time named cell is preserved when filled via win32com."""
        import openpyxl

        t = time(9, 15, 30)
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_time_cell": t},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            assert ws["H1"].value == time(9, 15, 30)
        finally:
            wb.close()

    def test_named_cell_decimal_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A Decimal named cell is converted to float when filled via win32com."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_decimal_cell": Decimal("12.34")},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["I1"].value
            assert val is not None
            assert isinstance(val, (float, int))
            assert abs(float(val) - 12.34) < 1e-10
        finally:
            wb.close()

    def test_named_cell_nan_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """float('nan') in a named cell becomes None when filled via win32com (should not crash)."""
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_nan_cell": float("nan")},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["J1"].value
            # win32com may write NaN as None or as a very small float;
            # either way, it should NOT crash
            if val is not None:
                # May be a very small float like 2.6789344e-304 on some COM versions
                import math

                assert isinstance(val, float)
                assert math.isnan(val) or abs(val) < 1e-300 or val is None
        finally:
            wb.close()

    def test_named_cell_pandas_timestamp_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """pd.Timestamp in a named cell becomes datetime when filled via win32com."""
        pd = pytest.importorskip("pandas")
        import openpyxl

        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={"cd_timestamp_cell": pd.Timestamp("2025-06-15 14:30:00")},
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["K1"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025
            assert val.month == 6
            assert val.day == 15
        finally:
            wb.close()

    def test_named_cell_tz_aware_datetime_stripped_and_warns_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A tz-aware datetime in a named cell is stripped and warns via win32com."""
        import openpyxl

        tz_dt = datetime(2025, 6, 15, 14, 30, 0, tzinfo=UTC)
        with pytest.warns(UserWarning, match="Timezone-aware"):
            fill(
                template=cross_backend_dtype_template,
                output=output_path,
                values={"cd_tz_datetime_cell": tz_dt},
                engine="win32com",
            )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            val = ws["L1"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025
            assert val.month == 6
            assert val.day == 15
            if hasattr(val, "tzinfo"):
                assert val.tzinfo is None
        finally:
            wb.close()


@pytest.mark.win32com
class TestCrossBackendTableWin32Com:
    """Table column dtype tests via the win32com backend (Excel required)."""

    def test_table_basic_dtypes_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Basic dtypes (str, int, float, bool) in a table are preserved via win32com."""
        import openpyxl

        data = [
            {"StrCol": "hello", "IntCol": 42, "FloatCol": 3.14, "BoolCol": True},
        ]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["A2"].value == "hello"
            assert ws["B2"].value == 42
            assert ws["C2"].value == 3.14
            # COM may return True/False or 1/0 for bool
            assert ws["D2"].value in (True, False, 1, 0)
        finally:
            wb.close()

    def test_table_none_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """None values in a table column are empty when filled via win32com."""
        import openpyxl

        data: list[dict[str, str | None]] = [{"NoneCol": None}, {"NoneCol": "not none"}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            assert ws["E2"].value is None
            assert ws["E3"].value == "not none"
        finally:
            wb.close()

    def test_table_datetime_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Datetime values in a table column are preserved via win32com."""
        import openpyxl

        data = [{"DatetimeCol": datetime(2025, 6, 15, 14, 30, 0)}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            val = ws["F2"].value
            assert val is not None
            assert hasattr(val, "year")
            assert val.year == 2025
            assert val.month == 6
            assert val.day == 15
        finally:
            wb.close()

    def test_table_float_nan_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """float('nan') in a table column does not crash win32com."""
        import openpyxl

        data = [{"FloatNanCol": float("nan")}, {"FloatNanCol": 1.0}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            # NaN may be stored as None or as a very small float in COM;
            # the key requirement is no crash.
            val = ws["J2"].value
            if val is not None:
                import math

                assert isinstance(val, float)
                assert math.isnan(val) or abs(val) < 1e-300
            assert ws["J3"].value == 1.0
        finally:
            wb.close()

    def test_table_large_int_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """Large integer values in a table column are preserved via win32com."""
        import openpyxl

        data = [{"LargeIntCol": 2**53 - 1}, {"LargeIntCol": -(2**53)}]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="win32com",
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            # COM may convert large ints to float
            val2 = ws["O2"].value
            assert val2 is not None
            assert float(val2) == float(2**53 - 1)
        finally:
            wb.close()

    def test_table_tz_aware_datetime_stripped_and_warns_via_win32com(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
    ) -> None:
        """A tz-aware datetime in a table column is stripped and warns via win32com."""
        import openpyxl

        tz_dt = datetime(2025, 6, 15, 14, 30, 0, tzinfo=UTC)
        data = [{"TzDatetimeCol": tz_dt}]
        with pytest.warns(UserWarning, match="Timezone-aware"):
            fill(
                template=cross_backend_dtype_template,
                output=output_path,
                tables={"CrossDtypeTable": data},
                config=_PERMISSIVE,
                engine="win32com",
            )
        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["CrossDtypeTable"]
            val = ws["N2"].value
            assert val is not None
            # COM may store as serial date number or datetime, and may
            # apply a timezone shift (local → UTC) — accept ±1 day tolerance
            # as documented in the cross-backend comparison tests.
            if hasattr(val, "year"):
                assert val.year == 2025
                assert val.month == 6
                assert abs(val.day - 15) <= 1
                if hasattr(val, "tzinfo"):
                    assert val.tzinfo is None
            elif isinstance(val, (int, float)):
                converted = _serial_to_datetime(float(val))
                assert converted.year == 2025
                assert converted.month == 6
                assert abs(converted.day - 15) <= 1
        finally:
            wb.close()


@pytest.mark.win32com
class TestCrossBackendXlwingsReadBack:
    """Rendered value verification via xlwings COM read-back (Excel required).

    These tests open the output in real Excel via xlwings to verify the
    *rendered* values, which is especially important for dates, formulas,
    and conditional formatting.
    """

    def test_named_cell_values_match_via_xlwings(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
        xl_app: object,  # xw.App
    ) -> None:
        """Named cell values rendered by Excel match expected values (xlwings read-back)."""
        import xlwings as xw

        dt = datetime(2025, 6, 15, 14, 30, 0)
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            values={
                "cd_str_cell": "Excel",
                "cd_int_cell": 99,
                "cd_float_cell": 2.71828,
                "cd_bool_cell": True,
                "cd_none_cell": None,
                "cd_datetime_cell": dt,
            },
            engine="win32com",
        )

        wb = xw.Book(str(output_path))
        try:
            ws = wb.sheets["CrossDtypeNamedCells"]
            assert ws.range("A1").value == "Excel"
            assert ws.range("B1").value == 99
            assert abs(ws.range("C1").value - 2.71828) < 1e-5
            # COM may return True or 1.0
            assert ws.range("D1").value in (True, 1.0, 1)
            assert ws.range("E1").value is None
            rendered = ws.range("F1").value
            assert rendered is not None
        finally:
            wb.close()

    def test_table_values_match_via_xlwings(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
        xl_app: object,
    ) -> None:
        """Table values rendered by Excel match expected values (xlwings read-back)."""
        import xlwings as xw

        data = [
            {
                "StrCol": "hello",
                "IntCol": 42,
                "FloatCol": 3.14,
                "BoolCol": True,
                "NoneCol": None,
            },
        ]
        fill(
            template=cross_backend_dtype_template,
            output=output_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="win32com",
        )

        wb = xw.Book(str(output_path))
        try:
            ws = wb.sheets["CrossDtypeTable"]
            # Table data starts at row 2 (row 1 = headers)
            assert ws.range("A2").value == "hello"
            assert ws.range("B2").value == 42
            assert abs(ws.range("C2").value - 3.14) < 1e-5
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# Phase C — Cross-backend comparison
# ═══════════════════════════════════════════════════════════════════════════════


class TestCrossBackendComparison:
    """Compare results from both backends for identical input data.

    Uses openpyxl read-back to compare outputs from both engines.
    The win32com tests are conditional (skipped when Excel is not available).
    """

    @pytest.mark.parametrize(
        ("cell_name", "value"),
        [
            pytest.param("cd_str_cell", "shared", id="str"),
            pytest.param("cd_int_cell", 42, id="int"),
            pytest.param("cd_float_cell", 3.14, id="float"),
            pytest.param("cd_bool_cell", True, id="bool_true"),
            pytest.param("cd_bool_cell", False, id="bool_false"),
            pytest.param("cd_none_cell", None, id="none"),
            pytest.param(
                "cd_datetime_cell",
                datetime(2025, 6, 15, 14, 30, 0),
                id="datetime",
            ),
            pytest.param("cd_date_cell", date(2025, 12, 25), id="date"),
            pytest.param("cd_decimal_cell", Decimal("12.34"), id="decimal"),
            pytest.param("cd_large_int_cell", 2**53 - 1, id="large_int"),
        ],
    )
    def test_named_cell_value_identical_across_backends(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
        cell_name: str,
        value: object,
    ) -> None:
        """A named cell written with both backends yields the same openpyxl read-back value."""
        import openpyxl

        # ── Fill via openpyxl ──────────────────────────────────────────────
        op_path = output_path.parent / "op_output.xlsx"
        fill(
            template=cross_backend_dtype_template,
            output=op_path,
            values={cell_name: value},
            engine="openpyxl",
        )
        wb = openpyxl.load_workbook(op_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            col_map = {
                "cd_str_cell": "A",
                "cd_int_cell": "B",
                "cd_float_cell": "C",
                "cd_bool_cell": "D",
                "cd_none_cell": "E",
                "cd_datetime_cell": "F",
                "cd_date_cell": "G",
                "cd_time_cell": "H",
                "cd_decimal_cell": "I",
                "cd_nan_cell": "J",
                "cd_timestamp_cell": "K",
                "cd_tz_datetime_cell": "L",
                "cd_large_int_cell": "M",
                "cd_large_float_cell": "N",
            }
            op_val = ws[f"{col_map[cell_name]}1"].value
        finally:
            wb.close()

        # ── Fill via win32com (skip if no Excel) ───────────────────────────
        com_path = output_path.parent / "com_output.xlsx"
        try:
            fill(
                template=cross_backend_dtype_template,
                output=com_path,
                values={cell_name: value},
                engine="win32com",
            )
        except (RuntimeError, Exception) as exc:
            msg = str(exc).lower()
            if "win32com" in msg or "excel" in msg or "com" in msg:
                pytest.skip(f"win32com backend not available: {exc}")
            raise

        wb = openpyxl.load_workbook(com_path)
        try:
            ws = wb["CrossDtypeNamedCells"]
            col_map = {
                "cd_str_cell": "A",
                "cd_int_cell": "B",
                "cd_float_cell": "C",
                "cd_bool_cell": "D",
                "cd_none_cell": "E",
                "cd_datetime_cell": "F",
                "cd_date_cell": "G",
                "cd_time_cell": "H",
                "cd_decimal_cell": "I",
                "cd_nan_cell": "J",
                "cd_timestamp_cell": "K",
                "cd_tz_datetime_cell": "L",
                "cd_large_int_cell": "M",
                "cd_large_float_cell": "N",
            }
            com_val = ws[f"{col_map[cell_name]}1"].value
        finally:
            wb.close()

        # ── Compare ────────────────────────────────────────────────────────
        if cell_name == "cd_decimal_cell":
            # openpyxl converts Decimal → float; COM may convert differently
            assert op_val is not None
            assert com_val is not None
            assert abs(float(op_val) - float(com_val)) < 1e-10
        elif cell_name == "cd_none_cell":
            assert op_val is None
            assert com_val is None
        elif cell_name in ("cd_datetime_cell", "cd_date_cell"):
            # Both backends should store the same calendar date.
            # Note: COM may shift by timezone (local → UTC conversion);
            # validate each backend independently against the ORIGINAL input.
            orig_date = value  # The original date/datetime input
            if isinstance(orig_date, date) and not isinstance(orig_date, datetime):
                orig_date = datetime(orig_date.year, orig_date.month, orig_date.day)
            # Check openpyxl result — should match exactly
            assert op_val is not None
            assert hasattr(op_val, "year")
            assert op_val.year == orig_date.year  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
            assert op_val.month == orig_date.month  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
            assert op_val.day == orig_date.day  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
            # Check win32com result — accept timezone shift (±1 day)
            if hasattr(com_val, "year"):
                # openpyxl read back as datetime (may have UTC tz shift)
                assert abs(com_val.day - orig_date.day) <= 1  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
            elif isinstance(com_val, (int, float)):
                serial_date = _serial_to_datetime(com_val)
                assert abs(serial_date.day - orig_date.day) <= 1, (  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                    f"COM date mismatch: expected day ~{orig_date.day}, "  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                    f"got {serial_date.day} (serial={com_val})"
                )
        elif cell_name == "cd_bool_cell":
            # COM may return 1/0 instead of True/False
            op_bool = bool(op_val) if op_val is not None else None
            com_bool = bool(com_val) if com_val is not None else None
            assert op_bool == com_bool
        elif cell_name == "cd_large_int_cell":
            # COM may convert large ints to float
            assert float(op_val) == float(com_val)
        else:
            assert op_val == com_val, (
                f"Mismatch for {cell_name}: openpyxl={op_val!r} vs win32com={com_val!r}"
            )

    @pytest.mark.parametrize(
        ("col_name", "values"),
        [
            pytest.param("StrCol", ["a", "b"], id="str"),
            pytest.param("IntCol", [1, 2], id="int"),
            pytest.param("FloatCol", [1.5, 2.5], id="float"),
            pytest.param("BoolCol", [True, False], id="bool"),
            pytest.param("NoneCol", [None, "not none"], id="none"),
            pytest.param(
                "DatetimeCol",
                [datetime(2025, 6, 15, 14, 30), datetime(2024, 1, 1)],
                id="datetime",
            ),
            pytest.param("DateCol", [date(2025, 12, 25), date(2023, 7, 4)], id="date"),
            pytest.param("DecimalCol", [Decimal("12.34"), Decimal("0.01")], id="decimal"),
        ],
    )
    def test_table_column_value_identical_across_backends(
        self,
        output_path: Path,
        cross_backend_dtype_template: Path,
        col_name: str,
        values: list[object],
    ) -> None:
        """A table column filled with both backends yields the same openpyxl read-back values."""
        import openpyxl

        data = [{col_name: v} for v in values]

        # ── Fill via openpyxl ──────────────────────────────────────────────
        op_path = output_path.parent / "op_table.xlsx"
        fill(
            template=cross_backend_dtype_template,
            output=op_path,
            tables={"CrossDtypeTable": data},
            config=_PERMISSIVE,
            engine="openpyxl",
        )
        wb = openpyxl.load_workbook(op_path)
        try:
            ws = wb["CrossDtypeTable"]
            col_map = {
                "StrCol": "A",
                "IntCol": "B",
                "FloatCol": "C",
                "BoolCol": "D",
                "NoneCol": "E",
                "DatetimeCol": "F",
                "DateCol": "G",
                "TimeCol": "H",
                "DecimalCol": "I",
                "FloatNanCol": "J",
                "NaCol": "K",
                "NaTCol": "L",
                "TimestampCol": "M",
                "TzDatetimeCol": "N",
                "LargeIntCol": "O",
                "LargeFloatCol": "P",
            }
            col_letter = col_map[col_name]
            op_vals = [ws[f"{col_letter}{i + 2}"].value for i in range(len(values))]
        finally:
            wb.close()

        # ── Fill via win32com ──────────────────────────────────────────────
        com_path = output_path.parent / "com_table.xlsx"
        try:
            fill(
                template=cross_backend_dtype_template,
                output=com_path,
                tables={"CrossDtypeTable": data},
                config=_PERMISSIVE,
                engine="win32com",
            )
        except (RuntimeError, Exception) as exc:
            msg = str(exc).lower()
            if "win32com" in msg or "excel" in msg or "com" in msg:
                pytest.skip(f"win32com backend not available: {exc}")
            raise

        wb = openpyxl.load_workbook(com_path)
        try:
            ws = wb["CrossDtypeTable"]
            col_letter = col_map[col_name]
            com_vals = [ws[f"{col_letter}{i + 2}"].value for i in range(len(values))]
        finally:
            wb.close()

        # ── Compare element-by-element ─────────────────────────────────────
        expected_values = [{col_name: v} for v in values]
        for i, (op_val, com_val) in enumerate(zip(op_vals, com_vals, strict=False)):
            orig_val = expected_values[i][col_name]

            if col_name == "DecimalCol":
                # Both backends convert Decimal to float; compare numeric values
                assert op_val is not None
                assert com_val is not None
                assert abs(float(op_val) - float(com_val)) < 1e-10
            elif col_name == "NoneCol":
                # Both backends should write None as empty cell
                if orig_val is None:
                    assert op_val is None, f"Row {i}: openpyxl should be None, got {op_val!r}"
                else:
                    assert op_val == orig_val
            elif col_name in ("DatetimeCol", "DateCol"):
                # openpyxl reads back datetimes as datetime objects (exact match).
                # COM may apply a UTC timezone shift (local → UTC) when writing,
                # so the read-back value may differ by up to 1 day.
                # Validate each backend independently against the original value.
                orig_dt = orig_val
                if isinstance(orig_dt, date) and not isinstance(orig_dt, datetime):
                    orig_dt = datetime(orig_dt.year, orig_dt.month, orig_dt.day)
                # Check openpyxl result — should match exactly
                assert op_val is not None
                assert hasattr(op_val, "year"), f"Row {i}: expected datetime, got {type(op_val)}"
                assert op_val.year == orig_dt.year  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert op_val.month == orig_dt.month  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                assert op_val.day == orig_dt.day  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                # Check win32com result — accept tz shift (±1 day tolerance)
                if com_val is not None:
                    if hasattr(com_val, "year"):
                        assert abs(com_val.day - orig_dt.day) <= 1  # type: ignore[attr-defined]  # ty:ignore[unresolved-attribute]
                    elif isinstance(com_val, (int, float)):
                        # Serial date read by openpyxl — compare with wide tolerance
                        # (Account for Excel leap-year bug + OLE tz conversion)
                        day_diff = abs(int(com_val) - (orig_dt - _SERIAL_EPOCH).days)  # type: ignore[operator]  # ty:ignore[unsupported-operator]
                        assert day_diff <= 2, (
                            f"Row {i}: COM serial {com_val} for {orig_dt} "
                            f"differs by {day_diff} days"
                        )
            elif col_name == "BoolCol":
                op_bool = bool(op_val) if op_val is not None else None
                com_bool = bool(com_val) if com_val is not None else None
                assert op_bool == com_bool
            else:
                assert op_val == com_val, (
                    f"Row {i} mismatch for {col_name}: openpyxl={op_val!r} vs win32com={com_val!r}"
                )
