"""Chart references must follow the table the same way in both engines."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.parity.builders import chart_template
from tests.parity.conftest import Parity

pytestmark = pytest.mark.excel

_ROWS = [{"fund": f"F{i}", "nav": i + 1, "alt": i + 2} for i in range(5)]


def test_a_chart_on_a_one_row_table_follows_it_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = chart_template(tmp_path / "template.xlsx", reference="Data!$B$2")
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _ROWS}))


def test_a_non_contiguous_series_follows_its_table_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = chart_template(
        tmp_path / "template.xlsx",
        header=("fund", "nav", "alt"),
        body=(("A", 1, 2),),
        reference="(Data!$B$2,Data!$C$2)",
    )
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _ROWS}))


def test_a_chart_on_a_swedish_sheet_follows_its_table_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = chart_template(
        tmp_path / "template.xlsx",
        sheet="Försäljning",
        reference="Försäljning!$B$2",
    )
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _ROWS}))


def test_a_chart_on_a_sheet_with_an_apostrophe_follows_its_table_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = chart_template(
        tmp_path / "template.xlsx",
        sheet="Bob's Data",
        reference="'Bob''s Data'!$B$2",
    )
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _ROWS}))


def test_a_chart_on_its_own_chart_sheet_follows_its_table_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = chart_template(tmp_path / "template.xlsx", chart_sheet=True)
    parity.assert_same(*parity.fill_both(template, tables={"Sales": _ROWS}))
