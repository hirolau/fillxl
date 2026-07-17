"""Fill all 3 test-template .xlsx files with demo data and save to a folder.

Usage:
    uv run python scripts/fill_demo.py                        # both backends → filled_templates/
    uv run python scripts/fill_demo.py --engine openpyxl
    uv run python scripts/fill_demo.py --engine win32com
    uv run python scripts/fill_demo.py --output-dir tests/assets/golden  # golden copies

Output is written to ``filled_templates/`` (or ``--output-dir``) at the repo
root, in subfolders named after the backend.
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Any, cast

from fillxl import FillXLConfig, fill
from fillxl._types import Engine
from fillxl._validation import ColumnValidationStrategy

ASSETS = Path(__file__).resolve().parent.parent / "tests" / "assets"
OUTPUT = Path(__file__).resolve().parent.parent / "filled_templates"


# ── Template 1: test_template_1.xlsx ─────────────────────────────────────────


def fill_template_1(outdir: Path, engine: Engine = "auto") -> Path:
    """Fill the basic template (table + all fillable named cells).

    Notes:
        ``named_range_multiple_cells`` (a 4-cell range ``E2:F3``) is
        intentionally skipped — :func:`fillxl.fill` only accepts single
        cells for the *values* parameter.
    """
    # ColFormula is a calculated column — omit it so the formula propagates.
    data = [
        {"Col1": "Alpha", "Col2": 10, "ColNotFilled": "x", "TestFormattedPct": 0.12},
        {"Col1": "Beta", "Col2": 20, "ColNotFilled": "y", "TestFormattedPct": 0.34},
        {"Col1": "Gamma", "Col2": 7, "ColNotFilled": "z", "TestFormattedPct": 0.56},
        {"Col1": "Delta", "Col2": 40, "ColNotFilled": "w", "TestFormattedPct": 0.78},
        {"Col1": "Epsilon", "Col2": 50, "ColNotFilled": "v", "TestFormattedPct": 0.90},
    ]
    # win32com can't detect calculated columns before resize, so use "drop"
    # strategy to skip missing columns instead of raising.
    strategy: ColumnValidationStrategy = (
        "drop" if engine == "win32com" else "all_but_formulas"
    )
    return fill(
        template=ASSETS / "test_template_1.xlsx",
        output=outdir / "test_template_1_filled.xlsx",
        tables={"TableTest": data},
        values={
            "single_cell": "Filled by fillxl",
            "single_cell_merged_cells": "Filled merged cell",
            "named_range_on_another_sheet": "Filled cross-sheet",
        },
        engine=engine,
        config=FillXLConfig(column_validation_strategy=strategy),
    )


# ── Template 2: test_template_2_charts.xlsx ──────────────────────────────────


def fill_template_2(outdir: Path, engine: Engine = "auto") -> Path:
    """Fill all 10 chart-bound tables."""
    tables: dict[str, list[dict[str, Any]]] = {
        "ColumnChartTable": [
            {"Region": "Nordic", "Q1": 10000, "Q2": 15000, "Q3": 14000, "Q4": 18000},
            {"Region": "Baltic", "Q1": 8000, "Q2": 9000, "Q3": 9500, "Q4": 11000},
            {"Region": "Balkans", "Q1": 6000, "Q2": 7000, "Q3": 7500, "Q4": 8000},
        ],
        "LineChartTable": [
            {"Date": datetime(2026, 1, 1), "2024": 200000, "2025": 220000},
            {"Date": datetime(2026, 2, 1), "2024": 210000, "2025": 230000},
            {"Date": datetime(2026, 3, 1), "2024": 205000, "2025": 240000},
            {"Date": datetime(2026, 4, 1), "2024": 215000, "2025": 250000},
        ],
        "PieChartTable": [
            {"Company": "Alpha", "Market Share": 0.35},
            {"Company": "Beta", "Market Share": 0.25},
            {"Company": "Gamma", "Market Share": 0.20},
            {"Company": "Delta", "Market Share": 0.20},
        ],
        "BarChartTable": [
            {"Department": "R&D", "Employees": 50},
            {"Department": "Support", "Employees": 30},
            {"Department": "Admin", "Employees": 15},
        ],
        "AreaChartTable": [
            {"Week": "W1", "Desktop": 30000, "Mobile": 50000, "Tablet": 8000},
            {"Week": "W2", "Desktop": 32000, "Mobile": 55000, "Tablet": 9000},
            {"Week": "W3", "Desktop": 35000, "Mobile": 60000, "Tablet": 10000},
            {"Week": "W4", "Desktop": 38000, "Mobile": 65000, "Tablet": 11000},
        ],
        "ScatterChartTable": [
            {"Product": "Widget Pro", "Price ($)": 39, "Units Sold": 2000},
            {"Product": "Widget Lite", "Price ($)": 19, "Units Sold": 4500},
            {"Product": "Gadget X", "Price ($)": 59, "Units Sold": 1200},
            {"Product": "Gadget Mini", "Price ($)": 29, "Units Sold": 3100},
        ],
        "DoughnutChartTable": [
            {"Category": "Salaries", "Budget": 500000},
            {"Category": "Marketing", "Budget": 100000},
            {"Category": "R&D", "Budget": 200000},
            {"Category": "Infra", "Budget": 75000},
        ],
        "StackedColumnChartTable": [
            {
                "Quarter": "Q1 2026",
                "Electronics": 90000,
                "Clothing": 45000,
                "Food": 40000,
                "Home": 30000,
            },
            {
                "Quarter": "Q2 2026",
                "Electronics": 95000,
                "Clothing": 50000,
                "Food": 42000,
                "Home": 32000,
            },
        ],
        "ComboChartTable": [
            {"Date": datetime(2027, 1, 1), "Revenue": 300000, "Growth Rate": 0.04},
            {"Date": datetime(2027, 2, 1), "Revenue": 315000, "Growth Rate": 0.05},
            {"Date": datetime(2027, 3, 1), "Revenue": 330000, "Growth Rate": 0.05},
        ],
        "WaterfallChartTable": [
            {"Category": "Revenue", "Amount": 600000},
            {"Category": "COGS", "Amount": -250000},
            {"Category": "OpEx", "Amount": -180000},
            {"Category": "Net Income", "Amount": 170000},
        ],
    }
    return fill(
        template=ASSETS / "test_template_2_charts.xlsx",
        output=outdir / "test_template_2_charts_filled.xlsx",
        tables=tables,
        engine=engine,
    )


# ── Template 3: test_template_3_pivot.xlsx ───────────────────────────────────


def fill_template_3(outdir: Path, engine: Engine = "auto") -> Path:
    """Fill the pivot source table."""
    data = [
        {
            "Region": "North",
            "Product": "Widgets",
            "Sales Rep": "Alice",
            "Quarter": "Q3",
            "Amount": 13000,
        },
        {
            "Region": "South",
            "Product": "Gadgets",
            "Sales Rep": "Bob",
            "Quarter": "Q3",
            "Amount": 8500,
        },
        {
            "Region": "East",
            "Product": "Widgets",
            "Sales Rep": "Carol",
            "Quarter": "Q3",
            "Amount": 11000,
        },
        {
            "Region": "West",
            "Product": "Gadgets",
            "Sales Rep": "Dave",
            "Quarter": "Q3",
            "Amount": 9500,
        },
        {
            "Region": "North",
            "Product": "Widgets",
            "Sales Rep": "Eve",
            "Quarter": "Q4",
            "Amount": 14500,
        },
        {
            "Region": "South",
            "Product": "Gadgets",
            "Sales Rep": "Frank",
            "Quarter": "Q4",
            "Amount": 9200,
        },
    ]
    return fill(
        template=ASSETS / "test_template_3_pivot.xlsx",
        output=outdir / "test_template_3_pivot_filled.xlsx",
        tables={"SalesData": data},
        engine=engine,
        config=FillXLConfig(refresh_pivot_tables=True),
    )


# ── Main ─────────────────────────────────────────────────────────────────────


def main() -> None:
    """Fill all templates with both backends and print results."""
    parser = argparse.ArgumentParser(description="Fill demo templates with sample data.")
    parser.add_argument(
        "--engine",
        choices=["openpyxl", "win32com", "both"],
        default="both",
        help="Backend engine to use (default: both).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT,
        help=(
            "Output directory (default: filled_templates/ at repo root). "
            "Each backend writes to a subfolder named after it."
        ),
    )
    args = parser.parse_args()

    engines: list[Engine] = (
        ["openpyxl", "win32com"] if args.engine == "both" else [cast(Engine, args.engine)]
    )
    output_root = args.output_dir

    all_paths: list[Path] = []
    for engine in engines:
        engine_dir = output_root / engine
        engine_dir.mkdir(parents=True, exist_ok=True)
        print(f"\n{'─' * 60}")
        print(f"  Backend: {engine}")
        print(f"{'─' * 60}")
        paths = [
            fill_template_1(engine_dir, engine=engine),
            fill_template_2(engine_dir, engine=engine),
            fill_template_3(engine_dir, engine=engine),
        ]
        all_paths.extend(paths)
        for p in paths:
            size = p.stat().st_size
            print(f"  {p.relative_to(output_root.parent)}  ({size:,} bytes)")

    print(f"\n{len(all_paths)} files written to {output_root}")


if __name__ == "__main__":
    main()
