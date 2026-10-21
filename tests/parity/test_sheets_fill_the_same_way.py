"""Sheets must end up in the same state in both engines."""

from __future__ import annotations

from pathlib import Path

import pytest

from fillxl import FillXLConfig
from tests.parity.builders import chart_sheet_first_template, sheets_template
from tests.parity.conftest import Parity

pytestmark = pytest.mark.excel


def test_hiding_the_active_underscore_sheet_makes_a_visible_sheet_active(
    tmp_path: Path, parity: Parity
) -> None:
    template = sheets_template(tmp_path / "template.xlsx", "_data", "Report")
    parity.assert_same(*parity.fill_both(template))


def test_a_very_hidden_sheet_stays_very_hidden_in_both_engines(
    tmp_path: Path, parity: Parity
) -> None:
    template = sheets_template(
        tmp_path / "template.xlsx", "_data", "Report", very_hidden=("_data",)
    )
    parity.assert_same(*parity.fill_both(template))


def test_deleting_the_active_sheet_leaves_the_same_sheet_active(
    tmp_path: Path, parity: Parity
) -> None:
    template = sheets_template(tmp_path / "template.xlsx", "Report", "Summary", active="Report")
    parity.assert_same(
        *parity.fill_both(
            template,
            config=FillXLConfig(hide_underscore_sheets=False, delete_sheets=("Report",)),
        )
    )


def test_deleting_a_sheet_before_the_active_one_keeps_it_active(
    tmp_path: Path, parity: Parity
) -> None:
    template = sheets_template(tmp_path / "template.xlsx", "Report", "Summary", active="Summary")
    parity.assert_same(
        *parity.fill_both(
            template,
            config=FillXLConfig(hide_underscore_sheets=False, delete_sheets=("Report",)),
        )
    )


def test_choosing_the_active_sheet_picks_the_same_tab(tmp_path: Path, parity: Parity) -> None:
    template = sheets_template(tmp_path / "template.xlsx", "Report", "Summary")
    parity.assert_same(*parity.fill_both(template, config=FillXLConfig(active_sheet="Summary")))


def test_the_active_sheet_after_a_chart_sheet_is_the_same(tmp_path: Path, parity: Parity) -> None:
    template = chart_sheet_first_template(tmp_path / "template.xlsx")
    parity.assert_same(*parity.fill_both(template, config=FillXLConfig(active_sheet="Report")))
