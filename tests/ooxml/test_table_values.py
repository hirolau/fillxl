"""Every engine validates a table cell through one function.

Three copies of the same four checks had to stay in step by hand; a rule added
to one copy quietly made an engine accept what another rejected.
"""

from __future__ import annotations

import datetime as dt

import pytest

from fillxl._normalize import CellValueError, prepare_table_cell


def test_prepare_table_cell_refuses_a_container() -> None:
    """A list is not a single Excel value, whatever the engine."""
    with pytest.raises(CellValueError):
        prepare_table_cell([1, 2], table="T", row=1, column="Amount")


def test_prepare_table_cell_returns_the_normalised_value() -> None:
    """The same call validates, warns and normalises."""
    assert (
        prepare_table_cell(
            dt.date(2025, 1, 2),
            table="T",
            row=1,
            column="Booked",
        )
        == dt.date(2025, 1, 2)
    )
