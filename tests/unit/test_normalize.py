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

from fillxl._normalize import (
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

    def test_a_date_unit_numpy_datetime64_becomes_a_datetime(self) -> None:
        assert normalize_value(np.datetime64("2025-06-15")) == dt.datetime(2025, 6, 15, 0, 0)

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

    def test_a_date_becomes_a_midnight_datetime(self) -> None:
        """Excel has no date type; everything is a serial number."""
        assert normalize_value(dt.date(2025, 12, 25)) == dt.datetime(2025, 12, 25, 0, 0)

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


class TestTimezoneHandling:
    """tz-aware datetimes lose their offset, keeping the wall-clock time."""

    def test_an_aware_datetime_loses_its_timezone(self) -> None:
        reset_tz_warning_state()
        aware = dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.timezone(dt.timedelta(hours=2)))
        with pytest.warns(UserWarning, match="Timezone-aware datetimes"):
            result = normalize_value(aware)
        assert result == dt.datetime(2025, 6, 15, 14, 30)
        assert result.tzinfo is None

    def test_the_warning_is_issued_once_per_fill_session(self) -> None:
        """Each fill session warns once; a new session starts over."""
        reset_tz_warning_state()
        aware = dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.UTC)
        with pytest.warns(UserWarning):
            normalize_value(aware)

        with warnings.catch_warnings():
            warnings.simplefilter("error")
            normalize_value(aware)  # already warned in this session

        reset_tz_warning_state()
        with pytest.warns(UserWarning):
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
