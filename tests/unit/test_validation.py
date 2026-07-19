"""Unit tests for column validation logic (pure function tests only).

Integration tests that exercise validation via fill() live in tests/integration/
or in the story suite (tests_story/).
"""

from __future__ import annotations

import pytest

from fillxl._validation import validate_columns


def test_validate_columns_returns_all_columns_on_exact_match() -> None:
    """Exact column match returns all columns unchanged."""
    result = validate_columns(["A", "B"], ["A", "B"], "strict")
    assert result == ["A", "B"]


def test_validate_columns_strict_raises_on_mismatch() -> None:
    """Strict strategy raises on mismatch."""
    with pytest.raises(ValueError, match="Column mismatch"):
        validate_columns(["A", "B"], ["A", "C"], "strict")


def test_validate_columns_drop_filters_to_intersection() -> None:
    """Drop strategy filters to intersection."""
    result = validate_columns(["A", "B"], ["A", "C", "D"], "drop")
    assert result == ["A"]


def test_validate_columns_warn_returns_all_original_columns() -> None:
    """Warn strategy returns all original columns."""
    result = validate_columns(["A", "B"], ["A", "C"], "warn")
    assert result == ["A", "C"]


def test_validate_columns_all_but_formulas_allows_extra_and_omits_formula() -> None:
    """Non-formula columns all present, extra allowed, formula omitted."""
    result = validate_columns(
        ["A", "B", "FormulaCol"],
        ["A", "B", "Extra"],
        "all_but_formulas",
        formula_columns=["FormulaCol"],
    )
    assert result == ["A", "B", "Extra"]


def test_validate_columns_all_but_formulas_raises_when_non_formula_column_missing() -> None:
    """Missing a non-formula column raises."""
    with pytest.raises(ValueError, match="Missing columns"):
        validate_columns(["A", "B"], ["A"], "all_but_formulas")


def test_validate_columns_all_but_formulas_strict_passes_with_exact_non_formula() -> None:
    """Exact non-formula columns passes."""
    result = validate_columns(
        ["A", "B", "FormulaCol"],
        ["A", "B"],
        "all_but_formulas_strict",
        formula_columns=["FormulaCol"],
    )
    assert result == ["A", "B"]


def test_validate_columns_all_but_formulas_strict_raises_for_extra_column() -> None:
    """Extra column in strict mode raises."""
    with pytest.raises(ValueError, match="Extra columns"):
        validate_columns(
            ["A", "B", "FormulaCol"],
            ["A", "B", "Extra"],
            "all_but_formulas_strict",
            formula_columns=["FormulaCol"],
        )


def test_validate_columns_all_but_formulas_strict_raises_for_missing_column() -> None:
    """Missing a non-formula column in strict mode raises."""
    with pytest.raises(ValueError, match="Missing columns"):
        validate_columns(
            ["A", "B", "FormulaCol"],
            ["A"],
            "all_but_formulas_strict",
            formula_columns=["FormulaCol"],
        )


def test_validate_columns_none_returns_data_columns_unchanged() -> None:
    """'none' returns data columns unchanged regardless of mismatch."""
    result = validate_columns(["A", "B"], ["X", "Y", "Z"], "none")
    assert result == ["X", "Y", "Z"]


def test_validate_columns_drop_logs_info_about_dropped_columns(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Drop strategy logs extra/missing column info."""
    caplog.set_level("INFO")
    validate_columns(["A", "B"], ["A", "C", "D"], "drop")
    assert any(
        record.levelname == "INFO" and "Dropping" in record.message for record in caplog.records
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
