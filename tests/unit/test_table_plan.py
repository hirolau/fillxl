"""The row arithmetic both engines share, tested without Excel."""

from __future__ import annotations

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from fillxl._plan import TableGeometry, plan_table_fill


def _geometry(
    *,
    top_row: int = 1,
    last_row: int = 2,
    header_rows: int = 1,
    totals_rows: int = 0,
    first_col: int = 1,
    last_col: int = 2,
) -> TableGeometry:
    return TableGeometry(
        top_row=top_row,
        last_row=last_row,
        first_col=first_col,
        last_col=last_col,
        header_rows=header_rows,
        totals_rows=totals_rows,
    )


def test_a_table_that_grows_checks_every_row_it_grows_into() -> None:
    """The rows below the table must be empty before the table can use them."""
    plan = plan_table_fill(_geometry(top_row=1, last_row=2), 5)

    assert plan.overflow_rows == (3, 6)
    assert plan.unused_body_rows is None
    assert plan.vacated_rows is None
    assert plan.new_ref == "A1:B6"
    assert plan.resized


def test_a_table_that_shrinks_gives_back_the_rows_below_it() -> None:
    """Rows the table no longer covers are cleared."""
    plan = plan_table_fill(_geometry(top_row=1, last_row=6), 2)

    assert plan.overflow_rows is None
    assert plan.vacated_rows == (4, 6)
    assert plan.unused_body_rows == (4, 6)
    assert plan.new_ref == "A1:B3"


def test_an_empty_payload_keeps_one_blank_body_row() -> None:
    """An Excel table needs a body row, so empty data keeps a cleared one."""
    plan = plan_table_fill(_geometry(top_row=1, last_row=4), 0)

    assert plan.new_body_last_row == 2
    assert plan.unused_body_rows == (2, 4)
    assert plan.vacated_rows == (3, 4)


def test_a_totals_row_moves_with_the_end_of_the_body() -> None:
    """The totals row is part of the table but not of the data body."""
    plan = plan_table_fill(_geometry(top_row=1, last_row=7, totals_rows=1), 3)

    assert plan.new_body_last_row == 4
    assert plan.new_last_row == 5
    assert plan.new_ref == "A1:B5"
    assert plan.new_autofilter_ref == "A1:B4"
    assert plan.unused_body_rows == (5, 6)
    assert plan.vacated_rows == (6, 7)


def test_a_headerless_table_starts_its_body_on_its_first_row() -> None:
    """Without a header row the table's first row is data."""
    geometry = _geometry(top_row=3, last_row=4, header_rows=0)

    assert geometry.first_data_row == 3

    plan = plan_table_fill(geometry, 2)

    assert plan.first_data_row == 3
    assert plan.new_ref == "A3:B4"
    assert not plan.resized


def test_a_table_that_keeps_its_size_is_not_resized() -> None:
    """Nothing moves, so there is nothing to check, clear or resize."""
    plan = plan_table_fill(_geometry(top_row=1, last_row=3), 2)

    assert not plan.resized
    assert plan.overflow_rows is None
    assert plan.unused_body_rows is None
    assert plan.vacated_rows is None


def test_a_negative_row_count_is_a_programming_error() -> None:
    """A negative count cannot describe a fill."""
    with pytest.raises(ValueError, match="must not be negative"):
        plan_table_fill(_geometry(), -1)


def test_the_change_spec_describes_the_data_body_only() -> None:
    """Chart references follow the body, never the header or the totals row."""
    plan = plan_table_fill(_geometry(top_row=12, last_row=20, totals_rows=1), 3)

    spec = plan.change_spec("Data")

    assert spec.worksheet_title == "Data"
    assert spec.first_data_row == 13
    assert spec.old_last_row == 19
    assert spec.new_last_row == 15


@settings(max_examples=200)
@given(
    top_row=st.integers(min_value=1, max_value=40),
    header_rows=st.integers(min_value=0, max_value=1),
    old_body_rows=st.integers(min_value=1, max_value=20),
    totals_rows=st.integers(min_value=0, max_value=1),
    first_col=st.integers(min_value=1, max_value=8),
    width=st.integers(min_value=1, max_value=8),
    data_rows=st.integers(min_value=0, max_value=30),
)
def test_the_plan_stays_inside_and_below_the_old_table(
    top_row: int,
    header_rows: int,
    old_body_rows: int,
    totals_rows: int,
    first_col: int,
    width: int,
    data_rows: int,
) -> None:
    """Unused and vacated rows belong to the old table; overflow is below it."""
    geometry = TableGeometry(
        top_row=top_row,
        last_row=top_row + header_rows + old_body_rows - 1 + totals_rows,
        first_col=first_col,
        last_col=first_col + width - 1,
        header_rows=header_rows,
        totals_rows=totals_rows,
    )

    plan = plan_table_fill(geometry, data_rows)

    for span in (plan.unused_body_rows, plan.vacated_rows):
        if span is not None:
            assert geometry.first_data_row <= span[0] <= span[1] <= geometry.last_row
    if plan.overflow_rows is not None:
        assert plan.overflow_rows[0] == geometry.last_row + 1
        assert plan.overflow_rows[1] == plan.new_last_row
    assert plan.new_body_last_row - plan.first_data_row + 1 == max(data_rows, 1)


class TestBlockedRegions:
    """A growing table must not write into another table or a merged range."""

    def test_a_table_refuses_to_grow_into_a_merged_range(self) -> None:
        from fillxl._plan import BlockedRegion, blocked_area

        merged = BlockedRegion(
            description="the merged range C15:D16",
            first_row=15,
            last_row=16,
            first_col=3,
            last_col=4,
        )

        assert blocked_area((14, 20), 1, 13, [merged]) is merged

    def test_a_table_may_grow_beside_a_merged_range(self) -> None:
        from fillxl._plan import BlockedRegion, blocked_area

        merged = BlockedRegion(
            description="the merged range C15:D16",
            first_row=15,
            last_row=16,
            first_col=3,
            last_col=4,
        )

        assert blocked_area((14, 20), 5, 6, [merged]) is None
