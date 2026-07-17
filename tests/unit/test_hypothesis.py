"""Property-based (Hypothesis) tests for fillxl.

Tests invariants of pure functions under random inputs.
"""

from __future__ import annotations

from hypothesis import given
from hypothesis import strategies as st

from fillxl._utils import col2num, col_to_excel
from fillxl._validation import validate_columns

# ── Column conversion round-trip ─────────────────────────────────────────────


@given(col=st.integers(min_value=1, max_value=16384))  # XFD = col 16384
def test_col_to_excel_round_trips(col: int) -> None:
    """col_to_excel(n) → col2num(result) should give back n for all valid Excel columns."""
    assert col2num(col_to_excel(col)) == col


# ── Column validation strategy invariants ────────────────────────────────────


@given(
    table_cols=st.lists(st.text(min_size=1), min_size=1, max_size=10, unique=True),
    data_cols=st.lists(st.text(min_size=1), min_size=0, max_size=15, unique=True),
)
def test_validate_columns_none_returns_data_unchanged(
    table_cols: list[str],
    data_cols: list[str],
) -> None:
    """The 'none' strategy should always return data_columns exactly as provided."""
    result = validate_columns(table_cols, data_cols, strategy="none")
    assert result == data_cols


@given(
    table_cols=st.lists(st.text(min_size=1), min_size=1, max_size=10, unique=True),
    data_cols=st.lists(st.text(min_size=1), min_size=1, max_size=15, unique=True),
)
def test_validate_columns_drop_returns_subset_of_data(
    table_cols: list[str],
    data_cols: list[str],
) -> None:
    """The 'drop' strategy should only return columns present in both table and data."""
    result = validate_columns(table_cols, data_cols, strategy="drop")
    assert set(result) <= set(data_cols)
    assert set(result) <= set(table_cols)


@given(
    table_cols=st.lists(st.text(min_size=1), min_size=1, max_size=10, unique=True),
    data_cols=st.lists(st.text(min_size=1), min_size=1, max_size=15, unique=True),
)
def test_validate_columns_warn_returns_all_original_columns(
    table_cols: list[str],
    data_cols: list[str],
) -> None:
    """The 'warn' strategy should always return data_columns unchanged."""
    result = validate_columns(table_cols, data_cols, strategy="warn")
    assert result == data_cols
