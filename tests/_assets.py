"""Paths to the workbook assets used by the test suite."""

from __future__ import annotations

from pathlib import Path

ASSETS = Path(__file__).parent / "assets"
TEMPLATE = ASSETS / "test_template_1.xlsx"
CHARTS_TEMPLATE_2 = ASSETS / "test_template_2_charts.xlsx"
PIVOT_TEMPLATE = ASSETS / "test_template_3_pivot.xlsx"
CHARTS_TEMPLATE_4 = ASSETS / "test_template_4_charts_edge_cases.xlsx"
