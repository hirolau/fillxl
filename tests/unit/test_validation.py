"""Unit tests for column validation logic (pure function tests only).

Integration tests that exercise validation via fill() live in tests/integration/
or in the story suite (tests_story/).
"""

from __future__ import annotations

import pytest

from fillxl._validation import validate_columns


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


def test_validate_columns_drop_accepts_extra_and_missing_columns() -> None:
    """Drop strategy accepts extra and missing columns without raising."""
    validate_columns(["A", "B"], ["A", "C", "D"], "drop")


def test_validate_columns_warn_accepts_extra_and_missing_columns() -> None:
    """Warn strategy accepts extra and missing columns without raising."""
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


def test_validate_columns_drop_logs_info_about_dropped_columns(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Drop strategy logs extra/missing column info."""
    caplog.set_level("INFO")
    validate_columns(["A", "B"], ["A", "C", "D"], "drop")
    assert any(
        record.levelname == "INFO" and "dropping extra data columns" in record.message
        for record in caplog.records
    )


def test_validate_columns_warn_logs_warning_about_extra_and_missing(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Warn strategy logs warnings about extra/missing columns."""
    caplog.set_level("WARNING")
    validate_columns(["A", "B"], ["A", "C"], "warn")
    assert any(
        record.levelname == "WARNING" and "extra" in record.message for record in caplog.records
    )
    assert any(
        record.levelname == "WARNING" and "missing" in record.message for record in caplog.records
    )
