"""Unit tests for value normalisation.

Every cell value in both backends goes through :mod:`fillxl._normalize`, so
these tests pin down the values Excel actually ends up with — including the
COM-specific adaptations that keep the two engines equivalent.
"""

from __future__ import annotations

import datetime as dt
import warnings
from decimal import Decimal

import numpy as np
import pandas as pd
import pytest

from fillxl import FillXLWarning, TimezoneWarning
from fillxl._normalize import (
    check_cell_value,
    check_named_cell_scalar,
    check_table_cell_scalar,
    is_formula_string,
    normalize_value,
    prepare_cell_value,
    prepare_com_value,
    reset_tz_warning_state,
)


class TestNumpyScalars:
    """NumPy scalars are unwrapped to native Python values, never to text."""

    @pytest.mark.parametrize(
        ("value", "expected_type"),
        [
            pytest.param(np.int64(7), int, id="int64"),
            pytest.param(np.int32(7), int, id="int32"),
            pytest.param(np.uint16(7), int, id="uint16"),
            pytest.param(np.float32(1.5), float, id="float32"),
            pytest.param(np.float64(1.5), float, id="float64"),
            pytest.param(np.bool_(True), bool, id="bool_"),
            pytest.param(np.str_("hi"), str, id="str_"),
        ],
    )
    def test_a_numpy_scalar_becomes_the_native_python_equivalent(
        self,
        value: object,
        expected_type: type,
    ) -> None:
        """np.int64 → int, np.float32 → float, … so Excel stores a number."""
        result = normalize_value(value)
        assert type(result) is expected_type
        assert result == value

    def test_a_numpy_integer_is_not_normalised_to_a_string(self) -> None:
        """The bug that made Excel flag cells with "number stored as text"."""
        assert not isinstance(normalize_value(np.int64(7)), str)

    def test_a_numpy_nan_becomes_none(self) -> None:
        """Excel cannot represent NaN, so it is written as an empty cell."""
        assert normalize_value(np.float32("nan")) is None

    def test_a_numpy_datetime64_becomes_a_datetime(self) -> None:
        """Microsecond datetime64 scalars unwrap to datetime."""
        assert normalize_value(np.datetime64("2025-06-15T14:30")) == dt.datetime(
            2025,
            6,
            15,
            14,
            30,
        )

    def test_a_nanosecond_datetime64_becomes_a_datetime(self) -> None:
        """Nanosecond datetime64 unpacks to an int, so it needs its own branch."""
        nano = np.datetime64("2025-06-15T14:30:00.123456789", "ns")
        # Truncated to the microsecond — Excel's serial numbers cannot
        # represent more precision than that anyway.
        assert normalize_value(nano) == dt.datetime(2025, 6, 15, 14, 30, 0, 123456)

    def test_a_picosecond_datetime64_is_not_treated_as_nanoseconds(self) -> None:
        """Every sub-microsecond unit is converted, not just ns.

        ``.item()`` returns a raw count for ps/fs/as too; assuming the count
        is in nanoseconds would turn 1 ps into 1 ns.
        """
        assert normalize_value(np.datetime64(1, "ps")) == dt.datetime(1970, 1, 1)

    def test_a_numpy_timedelta64_becomes_a_day_fraction(self) -> None:
        """timedelta64 → float, like Python timedelta."""
        assert normalize_value(np.timedelta64(6, "h")) == 0.25

    def test_a_zero_dimensional_array_behaves_like_a_scalar(self) -> None:
        """np.array(5) is a scalar in all but name."""
        assert normalize_value(np.array(5)) == 5

    def test_a_numpy_nat_becomes_none(self) -> None:
        """Missing datetime64 values become empty cells."""
        assert normalize_value(np.datetime64("NaT", "ns")) is None

    def test_numpy_day_precision_dates_arrive_as_dates(self) -> None:
        """A day-precision datetime64 is a date, so it gets a date-only format."""
        result = normalize_value(np.datetime64("2025-06-15"))
        assert result == dt.date(2025, 6, 15)
        assert not isinstance(result, dt.datetime)

    def test_a_longdouble_becomes_a_float_not_text(self) -> None:
        """np.longdouble.item() returns another NumPy scalar, not a float."""
        result = normalize_value(np.longdouble("1.25"))
        assert result == 1.25
        assert not isinstance(result, str)

    def test_a_half_precision_float_becomes_a_float(self) -> None:
        assert normalize_value(np.float16(1.5)) == 1.5

    def test_an_unsigned_byte_becomes_an_int(self) -> None:
        assert normalize_value(np.uint8(200)) == 200


class TestPandasSentinels:
    """pandas' missing-value sentinels become empty cells."""

    def test_pd_na_becomes_none(self) -> None:
        assert normalize_value(pd.NA) is None

    def test_pd_nat_becomes_none(self) -> None:
        assert normalize_value(pd.NaT) is None

    def test_a_pandas_timestamp_is_prepared_for_com_as_a_datetime(self) -> None:
        """COM cannot marshal pd.Timestamp, so it is converted."""
        prepared = prepare_com_value(pd.Timestamp("2025-06-15 14:30"))
        assert prepared == dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.UTC)


class TestValueNormalisation:
    """Date, time, duration and decimal handling."""

    def test_a_date_stays_a_date_until_it_reaches_excel(self) -> None:
        """Excel has no date type, but the two engines need to tell them apart."""
        from fillxl._normalize import prepare_com_value

        result = normalize_value(dt.date(2025, 12, 25))

        assert result == dt.date(2025, 12, 25)
        assert not isinstance(result, dt.datetime)
        assert prepare_com_value(result).isoformat() == "2025-12-25T00:00:00+00:00"

    def test_time_becomes_a_fraction_of_a_day(self) -> None:
        """Excel stores a time as the part of a day that has elapsed."""
        assert normalize_value(dt.time(9, 15, 30)) == pytest.approx(33330 / 86400)

    def test_midnight_is_zero(self) -> None:
        assert normalize_value(dt.time(0, 0)) == 0.0

    def test_time_keeps_microseconds(self) -> None:
        assert normalize_value(dt.time(9, 15, 30, 500_000)) == pytest.approx((33330.5) / 86400)

    def test_time_is_not_a_datetime_any_more(self) -> None:
        assert not isinstance(normalize_value(dt.time(9, 15, 30)), dt.datetime)

    def test_a_timedelta_becomes_a_fraction_of_a_day(self) -> None:
        assert normalize_value(dt.timedelta(hours=6)) == 0.25

    def test_a_decimal_becomes_a_float(self) -> None:
        assert normalize_value(Decimal("12.34")) == 12.34

    def test_infinity_becomes_none(self) -> None:
        assert normalize_value(float("inf")) is None

    def test_an_empty_string_becomes_none(self) -> None:
        """openpyxl reads an empty text cell back as None, so both engines
        write a blank cell instead.
        """
        assert normalize_value("") is None

    def test_whitespace_is_still_text(self) -> None:
        """Only a truly empty string is normalised away."""
        assert normalize_value(" ") == " "

    def test_an_unknown_object_falls_back_to_its_string_form(self) -> None:
        class Greeting:
            def __str__(self) -> str:
                return "Hello!"

        assert normalize_value(Greeting()) == "Hello!"

    def test_cell_value_error_is_a_value_error(self) -> None:
        """Callers can catch the documented Value error."""
        from fillxl import CellValueError

        error = CellValueError("too long", table="Sales", row=2, column="Notes")

        assert isinstance(error, ValueError)
        assert str(error) == "Table 'Sales', data row 2, column 'Notes': too long"

    def test_dates_before_march_1900_are_refused_because_excel_miscounts_them(self) -> None:
        """Excel counts a 29 February 1900, so the engines disagree before March."""
        from fillxl import CellValueError

        for value in (dt.date(1900, 2, 28), dt.datetime(1900, 2, 28, 12)):
            with pytest.raises(CellValueError, match="before 1900-03-01"):
                check_cell_value(value)

        check_cell_value(dt.date(1900, 3, 1))


class TestTimezoneHandling:
    """tz-aware datetimes lose their offset, keeping the wall-clock time."""

    def test_an_aware_datetime_loses_its_timezone(self) -> None:
        reset_tz_warning_state()
        aware = dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.timezone(dt.timedelta(hours=2)))
        with pytest.warns(TimezoneWarning, match="Timezone-aware datetimes"):
            result = normalize_value(aware)
        assert result == dt.datetime(2025, 6, 15, 14, 30)
        assert result.tzinfo is None

    def test_the_warning_is_issued_once_per_fill_session(self) -> None:
        """Each fill session warns once; a new session starts over."""
        reset_tz_warning_state()
        aware = dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.UTC)
        with pytest.warns(TimezoneWarning):
            normalize_value(aware)

        with warnings.catch_warnings():
            warnings.simplefilter("error")
            normalize_value(aware)  # already warned in this session

        reset_tz_warning_state()
        with pytest.warns(TimezoneWarning):
            normalize_value(aware)  # a new session warns again


class TestComValuePreparation:
    """COM reinterprets values the way typed input is interpreted.

    The preparation step neutralises that so both backends store the same
    thing.
    """

    @pytest.mark.parametrize(
        "text",
        [
            pytest.param("00123", id="leading_zeros"),
            pytest.param("1/2", id="fraction"),
            pytest.param("12:30", id="time"),
            pytest.param("2025-01-05", id="date"),
            pytest.param("2025-01-05 10:00", id="date_and_time"),
            pytest.param("15 Jun 2025", id="textual_date"),
            pytest.param("$1,234.50", id="currency"),
            pytest.param("12%", id="percent"),
            pytest.param("TRUE", id="boolean_word"),
            pytest.param("hello", id="plain_word"),
            pytest.param("it's", id="embedded_apostrophe"),
        ],
    )
    def test_text_is_marked_as_literal_with_an_apostrophe(self, text: str) -> None:
        """Excel's apostrophe prefix keeps the string exactly as provided."""
        assert prepare_com_value(text) == "'" + text

    def test_a_formula_string_stays_a_formula_by_default(self) -> None:
        """Both backends turn "=1+1" into a formula unless escaping is on."""
        assert prepare_com_value("=1+1") == "=1+1"

    def test_a_formula_string_is_escaped_when_requested(self) -> None:
        assert prepare_com_value("=1+1", escape_formulas=True) == "'=1+1"

    def test_numbers_none_and_booleans_are_left_alone(self) -> None:
        assert prepare_com_value(42) == 42
        assert prepare_com_value(3.5) == 3.5
        assert prepare_com_value(True) is True
        assert prepare_com_value(None) is None

    def test_datetimes_are_marked_as_utc(self) -> None:
        """COM would otherwise shift the serial number by the local offset."""
        prepared = prepare_com_value(dt.datetime(2025, 6, 15, 14, 30))
        assert prepared == dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.UTC)

    def test_the_openpyxl_preparation_leaves_text_and_datetimes_alone(self) -> None:
        """openpyxl stores text as text; no prefix or tz marking is needed."""
        assert prepare_cell_value("00123") == "00123"
        assert prepare_cell_value("=1+1") == "=1+1"
        assert prepare_cell_value(dt.datetime(2025, 6, 15, 14, 30)) == dt.datetime(
            2025,
            6,
            15,
            14,
            30,
        )


class TestFormulaDetection:
    """Only strings that Excel would treat as formulas count as formulas."""

    def test_a_leading_equals_sign_marks_a_formula(self) -> None:
        assert is_formula_string("=SUM(A1:A2)")

    def test_plain_strings_are_not_formulas(self) -> None:
        assert not is_formula_string("00123")
        assert not is_formula_string("a=b")
        assert not is_formula_string(123)
        assert not is_formula_string(None)


class TestDecimalSpecialValues:
    """A Decimal that Excel cannot represent becomes an empty cell.

    ``float()`` runs before the IEEE check, so a NaN, an infinity or a value
    too large for a double never reaches the file as ``nan``/``inf`` ? which
    is what makes Excel offer to repair the workbook.
    """

    @pytest.mark.parametrize(
        "value",
        [
            pytest.param(Decimal("NaN"), id="nan"),
            pytest.param(Decimal("sNaN"), id="signal_nan"),
            pytest.param(Decimal("Infinity"), id="infinity"),
            pytest.param(Decimal("-Infinity"), id="minus_infinity"),
            pytest.param(Decimal("1e400"), id="too_large_for_a_float"),
        ],
    )
    def test_a_decimal_excel_cannot_store_becomes_an_empty_cell(self, value: Decimal) -> None:
        assert normalize_value(value) is None

    def test_an_ordinary_decimal_keeps_its_value(self) -> None:
        assert normalize_value(Decimal("0.1")) == 0.1


class TestValuesExcelCannotStore:
    """Some values cannot go into a cell at all; they are refused with a location."""

    def test_text_with_a_lone_surrogate_is_refused_with_its_position(self) -> None:
        """A surrogate survives every control-character check but not UTF-8."""
        from fillxl import CellValueError

        with pytest.raises(CellValueError, match="lone surrogate") as caught:
            check_cell_value("bad \ud800 text", table="Sales", row=2, column="Notes")

        assert "position 4" in str(caught.value)
        assert "Table 'Sales', data row 2, column 'Notes'" in str(caught.value)

    def test_text_with_a_tab_is_still_allowed(self) -> None:
        check_cell_value("a\tb")

    def test_an_integer_beyond_the_double_range_is_refused(self) -> None:
        """Trailing zeros used to hide the size: 10**400 counted as one digit."""
        from fillxl import CellValueError

        with pytest.raises(CellValueError, match="too large for an Excel cell"):
            check_cell_value(10**400)

    def test_a_large_integer_that_excel_can_hold_only_warns(
        self, recwarn: pytest.WarningsRecorder
    ) -> None:
        check_cell_value(1234567890123456, table="Sales", row=1, column="Id")

        precision = [w for w in recwarn if issubclass(w.category, FillXLWarning)]
        assert len(precision) == 1
        assert "significant digits" in str(precision[0].message)

    def test_the_precision_warning_is_raised_once_per_column(self) -> None:
        warned: set[tuple[str | None, str | None]] = set()

        with pytest.warns(FillXLWarning) as caught:
            for row in (1, 2, 3):
                check_cell_value(
                    1234567890123456, table="Sales", row=row, column="Id", warned=warned
                )

        assert len(caught) == 1


class TestContainerValues:
    """A cell holds one value; containers are refused, not stringified."""

    def test_a_named_cell_refuses_a_list_and_suggests_tables(self) -> None:
        from fillxl import CellValueError

        with pytest.raises(CellValueError, match="tables="):
            check_named_cell_scalar("report", [1, 2])

    def test_a_named_cell_refuses_bytes(self) -> None:
        from fillxl import CellValueError

        with pytest.raises(CellValueError, match="decode it to text"):
            check_named_cell_scalar("report", b"abc")

    def test_a_named_cell_refuses_a_dataframe(self) -> None:
        from fillxl import CellValueError

        with pytest.raises(CellValueError, match="tables="):
            check_named_cell_scalar("report", pd.DataFrame({"a": [1]}))

    def test_a_named_cell_accepts_a_scalar(self) -> None:
        check_named_cell_scalar("report", "text")
        check_named_cell_scalar("report", 3)
        check_named_cell_scalar("report", None)

    def test_a_table_cell_refuses_a_dict_with_the_column_name(self) -> None:
        from fillxl import CellValueError

        with pytest.raises(CellValueError, match=r"column 'Notes'"):
            check_table_cell_scalar({"a": 1}, table="Sales", row=3, column="Notes")

    def test_a_zero_dimensional_array_is_still_a_scalar(self) -> None:
        """``np.array(5)`` is a number, not a container."""
        assert normalize_value(np.array(5)) == 5


class TestSubclassesAndEnums:
    """Values that are nearly a basic type are flattened to exactly one."""

    def test_a_naive_datetime_subclass_is_written_as_a_datetime(self) -> None:
        """openpyxl looks number formats up by exact type, so subclasses crash it."""

        class Clock(dt.datetime):
            pass

        result = normalize_value(Clock(2025, 6, 15, 14, 30))

        assert type(result) is dt.datetime
        assert result == dt.datetime(2025, 6, 15, 14, 30)

    def test_a_date_subclass_stays_a_date(self) -> None:
        class Day(dt.date):
            pass

        result = normalize_value(Day(2025, 6, 15))

        assert type(result) is dt.date

    def test_an_enum_member_is_written_as_its_value(self) -> None:
        """``str(Color.RED)`` is "Color.RED", which is never what the caller meant."""
        import enum

        class Color(enum.Enum):
            RED = "red"

        assert normalize_value(Color.RED) == "red"

    def test_an_integer_enum_member_is_written_as_an_integer(self) -> None:
        import enum

        class Level(enum.IntEnum):
            HIGH = 3

        result = normalize_value(Level.HIGH)

        assert result == 3
        assert isinstance(result, int)

    def test_an_enum_of_unsupported_values_still_normalises(self) -> None:
        import enum

        class When(enum.Enum):
            NOW = dt.date(2025, 6, 15)

        assert normalize_value(When.NOW) == dt.date(2025, 6, 15)


class TestComIntegerWidth:
    """pywin32 cannot marshal a Python int wider than 32 bits reliably."""

    def test_an_integer_beyond_int32_is_handed_over_as_a_float(self) -> None:
        assert prepare_com_value(5_000_000_000) == 5e9

    def test_an_integer_that_fits_stays_an_integer(self) -> None:
        assert prepare_com_value(2**31 - 1) == 2**31 - 1
