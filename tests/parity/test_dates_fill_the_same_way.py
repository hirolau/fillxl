"""Dates must look the same whichever engine fills the template."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from fillxl import FillXLConfig
from tests.parity.builders import named_cells_template, table_template
from tests.parity.conftest import Parity

pytestmark = pytest.mark.excel

_DATES = [{"fund": f"F{i}", "when": dt.date(2025, 1, i + 1)} for i in range(3)]
_DATETIMES = [{"fund": f"F{i}", "when": dt.datetime(2025, 1, i + 1, 9, 30)} for i in range(3)]
_MIXED = [
    {"fund": "F0", "when": dt.date(2025, 1, 1)},
    {"fund": "F1", "when": dt.datetime(2025, 1, 2, 9, 30)},
    {"fund": "F2", "when": None},
    {"fund": "F3", "when": 5},
]


def _date_template(tmp_path: Path) -> Path:
    return table_template(tmp_path / "template.xlsx", header=("fund", "when"))


def test_a_column_of_dates_looks_the_same_in_both_engines(tmp_path: Path, parity: Parity) -> None:
    template = _date_template(tmp_path)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _DATES}))


def test_a_column_of_datetimes_looks_the_same_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = _date_template(tmp_path)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _DATETIMES}))


def test_a_column_mixing_dates_datetimes_blanks_and_numbers_looks_the_same(
    tmp_path: Path, parity: Parity
) -> None:
    template = _date_template(tmp_path)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _MIXED}))


def test_dates_look_the_same_without_copying_table_styles(tmp_path: Path, parity: Parity) -> None:
    template = _date_template(tmp_path)
    parity.assert_same(
        *parity.fill_both(
            template,
            tables={"Sales": _MIXED},
            config=FillXLConfig(copy_table_styles=False),
        )
    )


def test_a_named_cell_holding_a_date_looks_the_same_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = named_cells_template(tmp_path / "template.xlsx", cells={"report_date": ("A1", None)})
    parity.assert_same(*parity.fill_both(template, values={"report_date": dt.date(2025, 1, 2)}))
