"""Pivot cache refresh behaviour, verified through the saved file itself.

The ``refreshOnLoad`` attribute lives in ``xl/pivotCache/pivotCacheDefinition*.xml``,
so these tests read the package instead of trusting either backend's object model.
Both engines must leave the same flag: a workbook that recalculates its pivots when
Excel opens it.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl
import pytest

from fillxl import Engine, FillXLConfig, fill

SALES_ROWS: list[dict[str, object]] = [
    {"Region": region, "Product": product, "Sales Rep": rep, "Quarter": quarter, "Amount": amount}
    for region, product, rep, quarter, amount in [
        ("North", "Widgets", "Alice", "Q3", 13000),
        ("South", "Gadgets", "Bob", "Q3", 8500),
        ("East", "Widgets", "Carol", "Q3", 11000),
        ("West", "Gadgets", "Dave", "Q3", 9500),
        ("North", "Widgets", "Eve", "Q4", 14500),
        ("South", "Gadgets", "Frank", "Q4", 9200),
    ]
]

ENGINES = [
    pytest.param("openpyxl", id="openpyxl"),
    pytest.param("win32com", id="win32com", marks=pytest.mark.win32com),
]


def _refresh_on_load_flags(path: Path) -> list[bool]:
    """Return the ``refreshOnLoad`` flag of every pivot cache in *path*."""
    with zipfile.ZipFile(path) as archive:
        definitions = sorted(
            name
            for name in archive.namelist()
            if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d+\.xml", name)
        )
        assert definitions, f"no pivot cache definitions in {path.name}"
        return [
            bool(re.search(r'\brefreshOnLoad="(1|true)"', archive.read(name).decode("utf-8")))
            for name in definitions
        ]


def _fill_pivot(output: Path, template: Path, engine: Engine, config: FillXLConfig) -> Path:
    """Fill the pivot template's source table with *engine*."""
    try:
        fill(
            template=template,
            output=output,
            tables={"SalesData": SALES_ROWS},
            engine=engine,
            config=config,
        )
    except Exception as exc:
        message = str(exc).lower()
        if engine == "win32com" and ("win32com" in message or "excel" in message):
            pytest.skip(f"win32com backend not available: {exc}")
        raise
    return output


@pytest.mark.parametrize("engine", ENGINES)
def test_both_engines_mark_pivot_caches_for_refresh_on_load(
    tmp_path: Path,
    pivot_template: Path,
    engine: Engine,
) -> None:
    """Every pivot cache is flagged so Excel recalculates it when opened."""
    output = _fill_pivot(
        tmp_path / f"pivot_{engine}.xlsx",
        pivot_template,
        engine,
        FillXLConfig(refresh_pivot_tables=True),
    )

    flags = _refresh_on_load_flags(output)
    assert all(flags), f"{engine} left a pivot cache without refreshOnLoad: {flags}"


@pytest.mark.parametrize("engine", ENGINES)
def test_disabling_the_flag_leaves_the_caches_untouched(
    tmp_path: Path,
    pivot_template: Path,
    engine: Engine,
) -> None:
    """refresh_pivot_tables=False must not add refreshOnLoad to the caches."""
    output = _fill_pivot(
        tmp_path / f"pivot_off_{engine}.xlsx",
        pivot_template,
        engine,
        FillXLConfig(refresh_pivot_tables=False),
    )

    flags = _refresh_on_load_flags(output)
    assert not any(flags), f"{engine} set refreshOnLoad despite the flag being off: {flags}"


@pytest.mark.parametrize("engine", ENGINES)
def test_the_pivot_still_reports_the_filled_data(
    tmp_path: Path,
    pivot_template: Path,
    engine: Engine,
) -> None:
    """The fill lands in the pivot's source table, so the cache is not stale."""
    output = _fill_pivot(
        tmp_path / f"pivot_data_{engine}.xlsx",
        pivot_template,
        engine,
        FillXLConfig(refresh_pivot_tables=True),
    )

    wb = openpyxl.load_workbook(output)
    try:
        ws = wb["Sales Data"]
        assert ws.cell(row=2, column=1).value == "North"
        assert ws.cell(row=7, column=5).value == 9200
    finally:
        wb.close()


FIXTURES = Path(__file__).parent.parent / "fixtures"


@pytest.mark.excel
def test_a_pivot_on_a_calculated_column_shows_the_new_data_not_stale_results(
    tmp_path: Path,
) -> None:
    """Refreshing a pivot before recalculation leaves it showing the old numbers."""
    template = FIXTURES / "pivot_on_calculated_column.xlsx"
    rows = [{"fund": "A", "amount": 10}, {"fund": "B", "amount": 5}]

    output = fill(template, tmp_path / "out.xlsx", engine="win32com", tables={"Sales": rows})

    pivot = openpyxl.load_workbook(output, data_only=True)["Pivot"]
    grand_total = next(c for row in pivot.iter_rows() for c in row if c.value == "Grand Total")
    assert pivot.cell(row=grand_total.row, column=grand_total.column + 1).value == 30
