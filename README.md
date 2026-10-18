# fillxl

[![PyPI version](https://img.shields.io/pypi/v/fillxl.svg)](https://pypi.org/project/fillxl/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](pyproject.toml)
[![CI](https://github.com/hirolau/fillxl/actions/workflows/ci.yml/badge.svg)](https://github.com/hirolau/fillxl/actions/workflows/ci.yml)
[![Docs](https://img.shields.io/badge/docs-mkdocs--material-blue)](https://hirolau.github.io/fillxl/)

**Fill Excel templates with data from Python.**

Designed around named cells and Excel tables — no more fragile cell references like A1 or B3. Charts, formulas, formatting, and conditional formatting are preserved and updated automatically.

Fill tables using list, pandas, polars or duckdb!

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

**Backend (optional)**

```python
fill(..., engine="openpyxl")   # default, cross-platform
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
    values: dict[str, Any] | None = None,
    tables: dict[str, TabularData] | None = None,
    *,
    engine: Engine = "auto",   # "auto", "openpyxl", "win32com"
    config: FillXLConfig | None = None
) -> Path
```

See the full documentation for `FillXLConfig` options.

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
| `date`, `time`, `datetime` | date/time serial (timezone stripped) |
| `timedelta` | fraction of a day |
| `None`, `NaN`, `pd.NA`, `pd.NaT` | empty cell |
| `"=SUM(A1:A2)"` | literal text (formula with `escape_formulas=False`) |

---

## Backends

| Engine     | Cross-platform | Requires Excel | Best for                     |
|------------|----------------|----------------|------------------------------|
| `openpyxl` | Yes            | No             | Most use cases, Linux/macOS  |
| `win32com` | Windows only   | Yes            | `.xlsm`/`.xlsb`, modern charts |

Both engines write the same values for the same input; the test suite enforces
this with a cross-backend equivalence suite. See [Backends](backends.md) for the
engine-specific details.

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

- `openpyxl` does not support dynamic arrays, `.xlsm`, or `.xlsb` files and drops some modern chart types.
- `win32com` requires Windows + Excel.
- Timezone-aware datetimes are converted to naive (wall-clock preserved).
- Tables with a totals row are supported: the totals row stays directly below the data body, keeping its labels, functions and styling. Totals formulas written as explicit ranges (e.g. `SUM(B2:B5)`) are not rewritten.
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
