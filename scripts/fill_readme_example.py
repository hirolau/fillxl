"""Fill ``test_template_readme_example.xlsx`` with stock-price data.

Usage:
    uv run python scripts/fill_readme_example.py                          → filled_templates/
    uv run python scripts/fill_readme_example.py --engine win32com
    uv run python scripts/fill_readme_example.py --output-dir /some/path

Output is written to ``filled_templates/`` (or ``--output-dir``) in a
subfolder named after the backend.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import cast

import numpy as np
import pandas as pd

from fillxl import fill
from fillxl._types import Engine

ASSETS = Path(__file__).resolve().parent.parent / "tests" / "assets"
OUTPUT = Path(__file__).resolve().parent.parent / "filled_templates"


def get_fake_stock_data(_ticker: str) -> pd.DataFrame:
    dates = pd.date_range("2020-01-01", "2025-12-31", freq="B")  # business days only
    rng = np.random.default_rng(42)
    changes = rng.normal(0.05, 2.0, size=len(dates))
    prices = np.maximum(5.0, 50.0 + np.cumsum(changes)).round(2)
    return pd.DataFrame({"Date": dates, "Stock Price": prices})

def fill_readme_example(outdir: Path, engine: Engine = "auto") -> Path:

    # Has columns "Date" and "Stock Price"
    stock_data = get_fake_stock_data("ACME")

    return fill(
        template=ASSETS / "test_template_readme_example.xlsx",
        output=outdir / "readme_example_filled.xlsx",
        tables={"StockPriceHistory": stock_data},
        values={
            "title": "ACME Corp Stock Analysis",
            "currency": "USD",
            "target_price": 45,
        },
        engine=engine,
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fill the README example template with stock-price data."
    )
    parser.add_argument(
        "--engine",
        default="openpyxl",
        choices=["auto", "openpyxl", "win32com"],
        help="Backend engine (default: auto)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT,
        help=f"Output directory (default: {OUTPUT})",
    )
    args = parser.parse_args()

    engine = cast(Engine, args.engine)
    outdir = args.output_dir / engine
    outdir.mkdir(parents=True, exist_ok=True)

    path = fill_readme_example(outdir, engine=engine)
    print(f"Filled -> {path}")


if __name__ == "__main__":
    main()
