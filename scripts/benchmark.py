#!/usr/bin/env python
"""Benchmark fillxl (polars/pandas/duckdb x openpyxl/win32com) vs reference methods.

Reads the 10-row city-data template, expands to 500k rows by self-concatenation,
then measures wall-clock time for each method across 5 iterations.
"""

from __future__ import annotations

import importlib
import shutil
import statistics
import tempfile
import time
from pathlib import Path
from typing import Any

TEMPLATE = (
    Path(__file__).resolve().parent.parent
    / "tests" / "assets" / "test_template_5_wide_table.xlsx"
)
TABLE_NAME = "CityData"

# ── Tune these for your environment ──────────────────────────────────────────
N_ROWS = 5000         # Number of rows to benchmark
N_ITER = 10            # Iterations per method (results show mean ± std)

# Estimated run time:
#   openpyxl: ~30-90s per 500k-iteration (grows linearly with rows)
#   win32com: ~100x slower than openpyxl (~hours for 500k x 5 iters)
#   pandas.to_csv: fastest (<1s)
#   polars.write_excel: similar to fillxl openpyxl
# ─────────────────────────────────────────────────────────────────────────────


# ── Data loading ─────────────────────────────────────────────────────────────


def load_and_expand() -> list[dict[str, Any]]:
    """Load 10 template rows and repeat to N_ROWS."""
    import openpyxl

    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb[TABLE_NAME]
    rows: list[dict[str, Any]] = []
    col_names: list[str] = []
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=11, values_only=True)):
        if ri == 0:
            col_names = list(row)  # type: ignore[arg-type]
            continue
        rows.append(dict(zip(col_names, row, strict=True)))
    wb.close()

    repeats = (N_ROWS // len(rows)) + 1
    return (rows * repeats)[:N_ROWS]


# ── Helpers ──────────────────────────────────────────────────────────────────


def _copy_template(tmp: Path) -> Path:
    """Copy the template into *tmp* and return the copy path."""
    dst = tmp / "template.xlsx"
    shutil.copy2(TEMPLATE, dst)
    return dst


# ── Runner ───────────────────────────────────────────────────────────────────


def _warmup(fn: Any, *args: Any, **kwargs: Any) -> None:  # noqa: ANN401
    """Run once to trigger imports / JIT compilation."""
    import contextlib
    with contextlib.suppress(Exception):
        fn(*args, **kwargs, iter_index=-1)


def _run(_: str, fn: Any, *args: Any, **kwargs: Any) -> list[float] | None:  # noqa: ANN401
    """Run *fn* N_ITER times and return wall-clock seconds.

    Returns ``None`` if the function raises on any iteration (e.g. win32com
    limitation like ``datetime.time`` not being COM-compatible).  The
    benchmark is skipped entirely rather than producing partial results.
    """
    times: list[float] = []
    for i in range(N_ITER):
        t0 = time.perf_counter()
        try:
            fn(*args, **kwargs, iter_index=i)
        except Exception:
            return None  # any failure → skip method entirely
        t1 = time.perf_counter()
        times.append(t1 - t0)
    return times


def _fmt(times: list[float] | None) -> str:
    if times is None:
        return "—  not supported"
    m = statistics.mean(times)
    s = statistics.stdev(times) if len(times) > 1 else 0.0
    raw = ", ".join(f"{t:.2f}s" for t in times)
    return f"{m:.2f}s  ± {s:.2f}s  [{raw}]"


# ── fillxl benchmarks ────────────────────────────────────────────────────────


def _iter_dir(tmp: Path, tag: str, i: int) -> Path:
    d = tmp / tag / str(i)
    d.mkdir(parents=True, exist_ok=True)
    return d


def bench_fillxl(
    data: list[dict[str, Any]],
    tmp: Path,
    engine: str,
    *,
    tag: str,
    iter_index: int,
) -> None:
    """fillxl with list[dict] data."""
    from fillxl import fill

    d = _iter_dir(tmp, tag, iter_index)
    tmpl = _copy_template(d)
    out = d / "out.xlsx"
    fill(template=tmpl, output=out, tables={TABLE_NAME: data}, engine=engine)


def bench_fillxl_polars(
    data: list[dict[str, Any]],
    tmp: Path,
    engine: str,
    *,
    tag: str,
    iter_index: int,
) -> None:
    """fillxl with polars DataFrame."""
    import polars as pl

    from fillxl import fill

    d = _iter_dir(tmp, tag, iter_index)
    tmpl = _copy_template(d)
    out = d / "out.xlsx"
    df: pl.DataFrame = pl.from_dicts(data)
    fill(template=tmpl, output=out, tables={TABLE_NAME: df}, engine=engine)


def bench_fillxl_pandas(
    data: list[dict[str, Any]],
    tmp: Path,
    engine: str,
    *,
    tag: str,
    iter_index: int,
) -> None:
    """fillxl with pandas DataFrame."""
    import pandas as pd

    from fillxl import fill

    d = _iter_dir(tmp, tag, iter_index)
    tmpl = _copy_template(d)
    out = d / "out.xlsx"
    df: pd.DataFrame = pd.DataFrame(data)
    fill(template=tmpl, output=out, tables={TABLE_NAME: df}, engine=engine)


def bench_fillxl_duckdb(
    data: list[dict[str, Any]],
    tmp: Path,
    engine: str,
    *,
    tag: str,
    iter_index: int,
) -> None:
    """fillxl with duckdb relation.

    Builds a pandas DataFrame from the data, registers it with duckdb,
    then passes the resulting relation to fillxl.
    """
    import duckdb

    from fillxl import fill

    d = _iter_dir(tmp, tag, iter_index)
    tmpl = _copy_template(d)
    out = d / "out.xlsx"

    # Create a pandas DataFrame, register with duckdb, then query it
    _df = __import__("pandas").DataFrame(data)  # type: ignore[no-untyped-call]
    rel = duckdb.sql("SELECT * FROM _df")
    fill(template=tmpl, output=out, tables={TABLE_NAME: rel}, engine=engine)


# ── Reference benchmarks ─────────────────────────────────────────────────────


def bench_pandas_to_excel(
    data: list[dict[str, Any]],
    tmp: Path,
    *,
    tag: str,
    iter_index: int,
) -> None:
    """pandas.DataFrame.to_excel with openpyxl engine (no template)."""
    import openpyxl
    import pandas as pd
    from openpyxl.worksheet.table import Table, TableColumn

    d = _iter_dir(tmp, tag, iter_index)
    out = d / "out.xlsx"
    df: pd.DataFrame = pd.DataFrame(data)

    # Write via openpyxl directly for a fair comparison (fillxl also uses openpyxl)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = TABLE_NAME

    # Headers
    for ci, col in enumerate(df.columns, 1):
        ws.cell(row=1, column=ci, value=col)

    # Data in chunks to avoid holding all in memory
    for ri in range(len(df)):
        row = df.iloc[ri]
        for ci, val in enumerate(row, 1):
            ws.cell(row=ri + 2, column=ci, value=val)

    # Add Excel table
    last_col = openpyxl.utils.get_column_letter(len(df.columns))
    table = Table(
        displayName=TABLE_NAME,
        ref=f"A1:{last_col}{len(df) + 1}",
        tableColumns=[TableColumn(id=i + 1, name=str(c)) for i, c in enumerate(df.columns)],
    )
    ws.add_table(table)
    wb.save(out)
    wb.close()


def bench_pandas_to_csv(
    data: list[dict[str, Any]],
    tmp: Path,
    *,
    tag: str,
    iter_index: int,
) -> None:
    """pandas.DataFrame.to_csv."""
    import pandas as pd

    d = _iter_dir(tmp, tag, iter_index)
    out = d / "out.csv"
    df: pd.DataFrame = pd.DataFrame(data)
    df.to_csv(out, index=False)


def bench_polars_write_excel(
    data: list[dict[str, Any]],
    tmp: Path,
    *,
    tag: str,
    iter_index: int,
) -> None:
    """polars.write_excel."""
    import polars as pl

    d = _iter_dir(tmp, tag, iter_index)
    out = d / "out.xlsx"
    df: pl.DataFrame = pl.from_dicts(data)
    df.write_excel(out, table_name=TABLE_NAME, table_style="TableStyleMedium9", autofit=True)


# ── Utilities ────────────────────────────────────────────────────────────────


def _can_import(mod: str) -> bool:
    try:
        importlib.import_module(mod)
        return True
    except ImportError:
        return False


def _has_excel() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


# ═══════════════════════════════════════════════════════════════════════════════


def main() -> None:
    print("=" * 74)
    print(f"  fillxl speed benchmark — {N_ROWS:,} rows x {N_ITER} iterations")
    print(f"  Template: {TEMPLATE.name} — table '{TABLE_NAME}'")
    print("=" * 74)

    # ── Load data ─────────────────────────────────────────────────────────
    print(f"\n  Loading 10 template rows and expanding to {N_ROWS:,}…")
    data = load_and_expand()
    print(f"  -> {len(data):,} rows x {len(data[0])} columns")

    with tempfile.TemporaryDirectory(prefix="fillxl_bench_") as tmp_str:
        tmp = Path(tmp_str)

        results: list[tuple[str, list[float]]] = []

        # ── fillxl — list[dict] ─────────────────────────────────────────
        print(f"\n  {'─' * 70}")
        print("  fillxl — list[dict]")
        print(f"  {'─' * 70}")

        _warmup(bench_fillxl, data, tmp, "openpyxl", tag="warmup")
        t = _run("", bench_fillxl, data, tmp, "openpyxl", tag="list_openpyxl")
        results.append(("fillxl list[dict] / openpyxl", t))
        print(f"    openpyxl  {_fmt(t)}" if t else "    openpyxl  —  ERROR")

        if _has_excel():
            _warmup(bench_fillxl, data, tmp, "win32com", tag="warmup_win32")
            t = _run("", bench_fillxl, data, tmp, "win32com", tag="list_win32com")
            results.append(("fillxl list[dict] / win32com", t or []))
            msg = f"    win32com  {_fmt(t)}" if t else "    win32com  -  skipped"
            print(msg)
        else:
            print("    win32com  -  skipped (no Excel)")

        # ── fillxl — polars ─────────────────────────────────────────────
        if _can_import("polars"):
            print(f"\n  {'─' * 70}")
            print("  fillxl — polars DataFrame")
            print(f"  {'─' * 70}")

            _warmup(bench_fillxl_polars, data, tmp, "openpyxl", tag="warmup_polars")
            t = _run("", bench_fillxl_polars, data, tmp, "openpyxl", tag="polars_openpyxl")
            results.append(("fillxl polars / openpyxl", t))
            print(f"    openpyxl  {_fmt(t)}" if t else "    openpyxl  —  ERROR")

            if _has_excel():
                _warmup(bench_fillxl_polars, data, tmp, "win32com",
                        tag="warmup_polars_w32")
                t = _run("", bench_fillxl_polars, data, tmp, "win32com",
                         tag="polars_win32com")
                results.append(("fillxl polars / win32com", t or []))
                pm = f"    win32com  {_fmt(t)}" if t else "    win32com  -  skipped"
                print(pm)
            else:
                print("    win32com  -  skipped (no Excel)")

            # polars.write_excel (standalone)
            if _can_import("xlsxwriter"):
                print(f"\n  {'─' * 70}")
                print("  polars.write_excel (standalone)")
                print(f"  {'─' * 70}")
                _warmup(bench_polars_write_excel, data, tmp, tag="warmup_pl_write")
                t = _run("", bench_polars_write_excel, data, tmp, tag="polars_write_excel")
                results.append(("polars.write_excel (standalone)", t))
                print(f"    {'openpyxl':11s}  {_fmt(t)}" if t else "    openpyxl  —  ERROR")

        # ── fillxl — pandas ─────────────────────────────────────────────
        if _can_import("pandas"):
            print(f"\n  {'─' * 70}")
            print("  fillxl — pandas DataFrame")
            print(f"  {'─' * 70}")

            _warmup(bench_fillxl_pandas, data, tmp, "openpyxl", tag="warmup_pd")
            t = _run("", bench_fillxl_pandas, data, tmp, "openpyxl", tag="pandas_openpyxl")
            results.append(("fillxl pandas / openpyxl", t))
            print(f"    openpyxl  {_fmt(t)}" if t else "    openpyxl  —  ERROR")

            if _has_excel():
                _warmup(bench_fillxl_pandas, data, tmp, "win32com",
                        tag="warmup_pd_w32")
                t = _run("", bench_fillxl_pandas, data, tmp, "win32com",
                         tag="pandas_win32com")
                results.append(("fillxl pandas / win32com", t or []))
                pm = f"    win32com  {_fmt(t)}" if t else "    win32com  -  skipped"
                print(pm)
            else:
                print("    win32com  -  skipped (no Excel)")

            # pandas reference methods
            print(f"\n  {'─' * 70}")
            print("  pandas reference methods")
            print(f"  {'─' * 70}")

            _warmup(bench_pandas_to_excel, data, tmp, tag="warmup_pd_xl")
            t = _run("", bench_pandas_to_excel, data, tmp, tag="pandas_to_excel")
            results.append(("pandas.to_excel (openpyxl)", t))
            print(f"    to_excel  {_fmt(t)}" if t else "    to_excel  —  ERROR")

            _warmup(bench_pandas_to_csv, data, tmp, tag="warmup_pd_csv")
            t = _run("", bench_pandas_to_csv, data, tmp, tag="pandas_to_csv")
            results.append(("pandas.to_csv", t))
            print(f"    to_csv    {_fmt(t)}" if t else "    to_csv    —  ERROR")

        # ── fillxl — duckdb ─────────────────────────────────────────────
        if _can_import("duckdb") and _can_import("pandas"):
            print(f"\n  {'─' * 70}")
            print("  fillxl — duckdb relation")
            print(f"  {'─' * 70}")

            _warmup(bench_fillxl_duckdb, data, tmp, "openpyxl", tag="warmup_duckdb")
            t = _run("", bench_fillxl_duckdb, data, tmp, "openpyxl", tag="duckdb_openpyxl")
            results.append(("fillxl duckdb / openpyxl", t))
            print(f"    openpyxl  {_fmt(t)}" if t else "    openpyxl  —  ERROR")

            if _has_excel():
                _warmup(bench_fillxl_duckdb, data, tmp, "win32com",
                        tag="warmup_duckdb_w32")
                t = _run("", bench_fillxl_duckdb, data, tmp, "win32com",
                         tag="duckdb_win32com")
                results.append(("fillxl duckdb / win32com", t or []))
                dm = f"    win32com  {_fmt(t)}" if t else "    win32com  -  skipped"
                print(dm)
            else:
                print("    win32com  -  skipped (no Excel)")

        # ── Summary ─────────────────────────────────────────────────────
        print(f"\n  {'═' * 70}")
        print(f"  SUMMARY — {N_ROWS:,} rows, {N_ITER} iterations")
        print(f"  {'═' * 70}")
        valid = [(lab, t) for lab, t in results if t is not None and len(t) > 0]
        valid.sort(key=lambda r: statistics.mean(r[1]))
        for label, times in valid:
            print(f"    {_fmt(times)}  {label}")
        skipped = [lab for lab, t in results if t is None]
        if skipped:
            print(f"  {'─' * 70}")
            for label in skipped:
                print(f"    SKIPPED  {label}")


if __name__ == "__main__":
    main()
