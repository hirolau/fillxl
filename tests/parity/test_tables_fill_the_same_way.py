"""The same template and data must give the same workbook in both engines."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from fillxl import FillXLConfig
from tests.parity.builders import named_cells_template, table_template
from tests.parity.conftest import Parity

pytestmark = pytest.mark.excel


def test_a_growing_table_ends_up_the_same_whichever_engine_fills_it(
    tmp_path: Path, parity: Parity
) -> None:
    template = table_template(tmp_path / "template.xlsx")
    rows = [{"fund": f"F{i}", "nav": i * 1.5} for i in range(10)]
    parity.assert_same(*parity.fill_both(template, tables={"Sales": rows}))


def test_a_shrinking_table_leaves_nothing_behind_in_either_engine(
    tmp_path: Path, parity: Parity
) -> None:
    body = tuple((f"F{i}", i) for i in range(8))
    template = table_template(tmp_path / "template.xlsx", body=body)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": [{"fund": "A", "nav": 1}]}))


def test_an_empty_payload_clears_the_table_the_same_way_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    body = tuple((f"F{i}", i) for i in range(4))
    template = table_template(tmp_path / "template.xlsx", body=body)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": []}))


def test_a_table_with_a_totals_row_grows_the_same_way_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = table_template(tmp_path / "template.xlsx", totals=True)
    rows = [{"fund": f"F{i}", "nav": i + 1} for i in range(6)]
    parity.assert_same(*parity.fill_both(template, tables={"Sales": rows}))


def test_a_table_with_a_totals_row_shrinks_the_same_way_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    body = tuple((f"F{i}", i + 1) for i in range(6))
    template = table_template(tmp_path / "template.xlsx", body=body, totals=True)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": [{"fund": "A", "nav": 3}]}))


def test_a_headerless_table_grows_the_same_way_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = table_template(tmp_path / "template.xlsx", headerless=True)
    rows = [{"fund": f"F{i}", "nav": i * 2} for i in range(5)]
    parity.assert_same(*parity.fill_both(template, tables={"Sales": rows}))


def test_a_headerless_table_shrinks_the_same_way_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    body = tuple((f"F{i}", i) for i in range(6))
    template = table_template(tmp_path / "template.xlsx", body=body, headerless=True)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": [{"fund": "A", "nav": 9}]}))


@pytest.mark.parametrize("escape_formulas", [True, False])
def test_strings_excel_likes_to_reinterpret_are_stored_the_same_way(
    tmp_path: Path, parity: Parity, escape_formulas: bool
) -> None:
    template = table_template(tmp_path / "template.xlsx", header=("text",), body=(("x",),))
    values = ["00123", "1/2", "#N/A", "TRUE", "=1+1"]
    rows = [{"text": value} for value in values]
    parity.assert_same(
        *parity.fill_both(
            template,
            tables={"Sales": rows},
            config=FillXLConfig(escape_formulas=escape_formulas),
        )
    )


@pytest.mark.xfail(strict=True, reason="fixed in WP4")
def test_named_cells_of_every_scalar_type_end_up_the_same_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = named_cells_template(
        tmp_path / "template.xlsx",
        cells={
            "fund_name": ("A1", "placeholder"),
            "fund_nav": ("A2", 0),
            "report_date": ("A3", None),
        },
    )
    values = {"fund_name": "NAE", "fund_nav": 12.5, "report_date": datetime(2025, 1, 2, 9, 30)}
    parity.assert_same(*parity.fill_both(template, values=values))
