"""Property-based (Hypothesis) tests for fillxl.

Tests invariants of pure functions under random inputs.
"""

from __future__ import annotations

import pytest
from hypothesis import given
from hypothesis import strategies as st

from fillxl._validation import validate_columns


@given(
    table_cols=st.lists(st.text(min_size=1), min_size=1, max_size=10, unique=True),
    data_cols=st.lists(st.text(min_size=1), min_size=0, max_size=15, unique=True),
)
def test_validate_columns_none_accepts_any_column_combination(
    table_cols: list[str],
    data_cols: list[str],
) -> None:
    """The 'none' strategy accepts any combination of table and data columns."""
    validate_columns(table_cols, data_cols, strategy="none")


@pytest.mark.filterwarnings("ignore::fillxl._warnings.FillXLWarning")
@given(
    table_cols=st.lists(st.text(min_size=1), min_size=1, max_size=10, unique=True),
    data_cols=st.lists(st.text(min_size=1), min_size=1, max_size=15, unique=True),
)
def test_validate_columns_warn_accepts_any_column_combination(
    table_cols: list[str],
    data_cols: list[str],
) -> None:
    """The 'warn' strategy never raises, whatever the column mismatch."""
    validate_columns(table_cols, data_cols, strategy="warn")


@given(
    table_cols=st.lists(st.text(min_size=1), min_size=1, max_size=10, unique=True),
    data_cols=st.lists(st.text(min_size=1), min_size=1, max_size=15, unique=True),
)
def test_validate_columns_strict_raises_exactly_on_mismatch(
    table_cols: list[str],
    data_cols: list[str],
) -> None:
    """The 'strict' strategy only raises when the column sets differ."""
    if set(table_cols) == set(data_cols):
        validate_columns(table_cols, data_cols, strategy="strict")
    else:
        with pytest.raises(ValueError, match="Column mismatch"):
            validate_columns(table_cols, data_cols, strategy="strict")
