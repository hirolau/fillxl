# fillxl

[![PyPI version](https://img.shields.io/pypi/v/fillxl.svg)](https://pypi.org/project/fillxl/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](pyproject.toml)
[![CI](https://github.com/hirolau/fillxl/actions/workflows/ci.yml/badge.svg)](https://github.com/hirolau/fillxl/actions/workflows/ci.yml)
[![Docs](https://img.shields.io/badge/docs-mkdocs--material-blue)](https://hirolau.github.io/fillxl/)

**Design in Excel. Populate with Python.**

Create polished templates with named cells and Excel Tables — then fill them from pandas, Polars, DuckDB or plain Python. Charts, formatting and formulas stay intact.

fillxl focuses on a small, reliable API with strong correctness: tables resize properly, formatting is preserved, and edge cases are handled carefully.

### Quick example

Pass your Excel template, then map Excel **Tables** and **named cells** to Python data:

```python
from fillxl import fill
import pandas as pd

fill(
    template="report_template.xlsx",
    output="q4_report.xlsx",
    tables={
        "Products": [
            {"Name": "Widget", "Price": 4.99, "Sold": 150},
            {"Name": "Gadget", "Price": 9.99, "Sold": 85},
        ],
        "Sales": pd.read_csv('sales_data.csv'),
    },
    values={
        "ReportTitle": "Q4 2024 Sales",
        "PreparedBy": "Alice",
    },
)
```

---

## Installation

```bash
pip install fillxl
```

Optional extras:

```bash
pip install fillxl[pandas]   # pandas support
pip install fillxl[all]      # pandas + polars + duckdb
```

---

## Quick Start

1. In Excel, create **named cells** (Name Manager) for single values and **Excel Tables** (`Ctrl + T`) for tabular data.  
2. Fill the template with the code above.

Tables are automatically resized, charts updated, and all formatting preserved.

**Engine**

```python
fill(..., engine="auto")       # default: openpyxl unless Excel is needed
fill(..., engine="openpyxl")   # cross-platform, no Excel needed
fill(..., engine="win32com")   # Windows + Excel installed (maximum fidelity)
```

---

## Visual Example

**Before filling:**

![Before fill](docs/screenshots/before_fill.png)

**After filling — data, chart, and values populated:**

![After fill](docs/screenshots/after_fill.png)

:book: [See the code that produced this](https://hirolau.github.io/fillxl/#visual-example)

---

## Main Features

- Populate named cells and Excel tables
- Automatic table resizing + chart range updates
- Conditional formatting and pivot table refresh
- Two backends: `openpyxl` (default) and `win32com` (Windows + Excel)

---

## API

### `fill()`

```python
fill(
    template: str | Path,
    output: str | Path,
    *,
    values: dict[str, Any] | None = None,
    tables: dict[str, TabularData] | None = None,
    engine: Engine = "auto",   # "auto", "openpyxl", "win32com"
    config: FillXLConfig | None = None
) -> Path
```

See the full documentation for `FillXLConfig` options.

---

## Choosing an engine

`engine="auto"` (the default) picks the engine that can do the job:

| Template → output | Engine used |
|-------------------|-------------|
| `.xlsx` → `.xlsx`, `.xlsm` → `.xlsm` | `openpyxl` (no Excel needed) |
| `.xlsb`/`.xls` template, or a format conversion | `win32com` when Excel is available |
| anything else, without Excel | `openpyxl`, whose error says what is missing |

Pass `engine="openpyxl"` or `engine="win32com"` to choose explicitly. Both
engines produce the same workbook for the same template and data, and a parity
test suite checks it — see [Backends](docs/backends.md).

---

## Dates

- A `date` is written with the number format `yyyy-mm-dd`; a `datetime` with
  `yyyy-mm-dd h:mm:ss`. Both engines use these defaults, so a date never shows
  up as a midnight timestamp.
- A date format already in the template wins: format the cell in Excel (for
  example `dd.mm.yyyy`) and fillxl leaves it alone.
- Dates before **1900-03-01** raise `CellValueError`. Excel counts a 29
  February 1900 that never existed, so the two engines cannot agree before
  that day.
- Timezone-aware datetimes lose their offset; the wall-clock time is kept.

---

## Sheets

After filling, fillxl can tidy the workbook up:

```python
from fillxl import FillXLConfig, fill

fill(
    template="template.xlsx",
    output="report.xlsx",
    config=FillXLConfig(
        hide_underscore_sheets=True,      # hide "_data", "_lookups", …
        delete_sheets=["RawData"],        # remove sheets that must not ship
        active_sheet="Report",            # what the reader sees first
    ),
)
```

- Only **visible** sheets are hidden, and the sheet that was active always
  moves to a visible one.
- Hiding or deleting every visible sheet raises `ValueError`: Excel requires at
  least one.
- A sheet that should have been deleted but could not be raises
  `RuntimeError`, because a leftover sheet must not leak into a report.
- A hidden sheet cannot be made active (`ValueError`), and exactly one tab is
  left selected so Excel does not open the workbook with sheets grouped.

---

## Warnings

Problems fillxl can work around are warnings, and every one of them is a
`fillxl.FillXLWarning` subclass: `warnings.simplefilter("error", FillXLWarning)`
turns them into errors, and `logging.captureWarnings(True)` puts them in your
log. Every category and every exception is listed in [docs/warnings.md](docs/warnings.md).

---

## Supported Data Sources

- `list[dict]`
- `dict[str, sequence]` (lists, tuples, NumPy arrays, pandas/polars Series)
- `[]` or `{}` (empty data — clears the table)
- `pandas.DataFrame`
- `polars.DataFrame`
- `duckdb.DuckDBPyRelation`

---

## Values

Non-empty strings are always stored as text — `"00123"` stays `"00123"`
rather than becoming the number `123`, in both engines. Strings starting with
`=` are kept as literal text by default, so untrusted data cannot inject
formulas; pass `FillXLConfig(escape_formulas=False)` when the input is trusted
and is meant to contain formulas.

| Python value | Written to Excel as |
|--------------|---------------------|
| `str` | text (error-like strings such as `"#N/A"` included) |
| `""` | empty cell (openpyxl cannot round-trip an empty text cell) |
| `int`, `float`, `Decimal`, NumPy scalars | number |
| `bool` | boolean |
| `date` | date serial, shown as `yyyy-mm-dd` |
| `datetime` | date/time serial, shown as `yyyy-mm-dd h:mm:ss` |
| `time` | fraction of a day |
| `timedelta` | fraction of a day |
| `None`, `NaN`, `pd.NA`, `pd.NaT` | empty cell |
| `"=SUM(A1:A2)"` | literal text (formula with `escape_formulas=False`) |

A date format already in the template wins over the defaults, and dates before
1900-03-01 are refused — see [Dates](#dates).

---

## Backends

| Engine     | Cross-platform | Requires Excel | Best for                     |
|------------|----------------|----------------|------------------------------|
| `openpyxl` | Yes            | No             | Most use cases, Linux/macOS  |
| `win32com` | Windows only   | Yes            | `.xlsb`/`.xls`, modern charts |

Both engines write the same values for the same input; the test suite enforces
this with a cross-backend equivalence suite and a parity suite that fills the
same template with both engines and compares the workbooks. See
[Backends](docs/backends.md) for the engine-specific details.

---

## Performance

Benchmark: **5,000 rows × 20 columns** (list[dict] data).

| Method                    | Engine     | Time (mean ± std) |
|---------------------------|------------|-------------------|
| `fillxl`                  | openpyxl   | 2.26s ± 0.03s     |
| `fillxl`                  | win32com   | 2.80s ± 0.03s     |
| `pandas.to_excel`         | openpyxl   | 3.01s ± 0.09s     |

---

## Known Limitations

- `openpyxl` does not support dynamic arrays, `.xlsb`, or `.xls` files and drops some modern chart types (reported by `on_lossy`).
- `win32com` requires Windows + Excel.
- Timezone-aware datetimes are converted to naive (wall-clock preserved).
- Dates before 1900-03-01 are refused: Excel's calendar is wrong before that day.
- Tables with a totals row are supported: the totals row stays directly below the data body, keeping its labels, functions and styling. Totals formulas written as explicit ranges (e.g. `SUM(B2:B5)`) are not rewritten.
- Tables without a header row are supported: Excel names their columns `Column1`, `Column2`, …, so pass those names as the data keys.
- Strings starting with `=` are written as literal text by default; pass `FillXLConfig(escape_formulas=False)` only when the data is trusted and meant to contain formulas.

---

## Development

The project will initially stay at 0.X.X to allow early adopters to try out the package. Once it has been out for while version will be bumped to 1 to signal stability.

No new features are planned, and new features will purly be driven by the community.

```bash
git clone https://github.com/hirolau/fillxl.git
cd fillxl
uv sync --all-extras
uv run pytest
```

**Supported:** Python 3.11–3.15, Windows/Linux (macOS best-effort).  
**Data backends:** pandas 2–3, polars 0.20–2, duckdb 1.x (duckdb has no
Python 3.15 wheels yet, so its extra is limited to Python ≤3.14).  
**License:** [MIT](LICENSE)
