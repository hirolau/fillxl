"""Unit tests for column validation logic (pure function tests only).

Integration tests that exercise validation via fill() live in tests/integration/
or in the story suite (tests_story/).
"""

from __future__ import annotations

import warnings

import pytest

from fillxl._validation import validate_columns
from fillxl._warnings import FillXLWarning


def test_validate_columns_accepts_exact_match() -> None:
    """Exact column match passes validation."""
    validate_columns(["A", "B"], ["A", "B"], "strict")  # does not raise


def test_validate_columns_reports_without_returning_a_column_list() -> None:
    """The function only validates; backends match columns themselves."""
    assert validate_columns(["A"], ["A"], "strict") is None  # type: ignore[func-returns-value]


def test_validate_columns_strict_raises_on_mismatch() -> None:
    """Strict strategy raises on mismatch."""
    with pytest.raises(ValueError, match="Column mismatch"):
        validate_columns(["A", "B"], ["A", "C"], "strict")


def test_the_warn_strategy_emits_one_warning_that_names_both_column_sets() -> None:
    """A mismatch is the caller's to act on, so it is a warning and not a log line."""
    with pytest.warns(FillXLWarning) as caught:
        validate_columns(["A", "B"], ["A", "C"], "warn", table_name="T")

    assert len(caught) == 1
    message = str(caught[0].message)
    assert "missing columns: ['B']" in message
    assert "extra columns: ['C']" in message


def test_a_column_mismatch_becomes_an_error_under_the_warning_filter() -> None:
    """``simplefilter('error', FillXLWarning)`` has to catch a column mismatch."""
    with warnings.catch_warnings():
        warnings.simplefilter("error", FillXLWarning)
        with pytest.raises(FillXLWarning):
            validate_columns(["A", "B"], ["A", "C"], "warn")


def test_validate_columns_all_but_formulas_allows_extra_and_omits_formula() -> None:
    """Non-formula columns all present, extra allowed, formula omitted."""
    validate_columns(
        ["A", "B", "FormulaCol"],
        ["A", "B", "Extra"],
        "all_but_formulas",
        formula_columns=["FormulaCol"],
    )


def test_validate_columns_all_but_formulas_raises_when_non_formula_column_missing() -> None:
    """Missing a non-formula column raises."""
    with pytest.raises(ValueError, match="no data column"):
        validate_columns(["A", "B"], ["A"], "all_but_formulas")


def test_validate_columns_all_but_formulas_strict_passes_with_exact_non_formula() -> None:
    """Exact non-formula columns passes."""
    validate_columns(
        ["A", "B", "FormulaCol"],
        ["A", "B"],
        "all_but_formulas_strict",
        formula_columns=["FormulaCol"],
    )


def test_validate_columns_all_but_formulas_strict_raises_for_extra_column() -> None:
    """Extra column in strict mode raises."""
    with pytest.raises(ValueError, match="extra columns"):
        validate_columns(
            ["A", "B", "FormulaCol"],
            ["A", "B", "Extra"],
            "all_but_formulas_strict",
            formula_columns=["FormulaCol"],
        )


def test_validate_columns_all_but_formulas_strict_raises_for_missing_column() -> None:
    """Missing a non-formula column in strict mode raises."""
    with pytest.raises(ValueError, match="missing columns"):
        validate_columns(
            ["A", "B", "FormulaCol"],
            ["A"],
            "all_but_formulas_strict",
            formula_columns=["FormulaCol"],
        )


def test_validate_columns_none_skips_validation_entirely() -> None:
    """'none' skips validation regardless of mismatch."""
    validate_columns(["A", "B"], ["X", "Y", "Z"], "none")


def test_validate_columns_still_checks_data_with_rows_but_no_columns() -> None:
    """A DataFrame with an index and no columns resizes the table to blank rows."""
    with pytest.raises(ValueError, match="missing columns"):
        validate_columns(["A", "B"], [], "strict", data_rows=5)


def test_validate_columns_still_checks_all_but_formulas_with_rows_but_no_columns() -> None:
    with pytest.raises(ValueError, match="no data column"):
        validate_columns(["A", "B"], [], "all_but_formulas", data_rows=3)


def test_validate_columns_accepts_data_with_no_rows_and_no_columns() -> None:
    """Empty data is the "clear the table" case, which stays allowed."""
    validate_columns(["A", "B"], [], "strict")
    validate_columns(["A", "B"], [], "strict", data_rows=0)
