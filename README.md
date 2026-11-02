# fillxl

[![PyPI version](https://img.shields.io/pypi/v/fillxl.svg)](https://pypi.org/project/fillxl/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/hirolau/fillxl/blob/main/LICENSE)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://github.com/hirolau/fillxl/blob/main/pyproject.toml)
[![CI](https://github.com/hirolau/fillxl/actions/workflows/ci.yml/badge.svg)](https://github.com/hirolau/fillxl/actions/workflows/ci.yml)
[![Docs](https://img.shields.io/badge/docs-mkdocs--material-blue)](https://hirolau.github.io/fillxl/)

**Design in Excel. Populate with Python.**

Create polished templates with named cells and Excel Tables — then fill them
from pandas, Polars, DuckDB or plain Python. Charts, formatting and formulas
stay intact.

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
        "Sales": pd.read_csv("sales_data.csv"),
    },
    values={
        "ReportTitle": "Q4 2024 Sales",
        "PreparedBy": "Alice",
    },
)
```

## Installation

```bash
pip install fillxl
```

That is all you need. Filling works with plain Python (`list[dict]`,
`dict[str, list]`) and with the pandas, Polars or DuckDB objects your own code
already creates — there is nothing else to choose:

```python
import pandas as pd
from fillxl import fill

fill(..., tables={"Sales": pd.read_csv("sales.csv")})  # works as-is
```

fillxl never imports a data library itself; it recognises the DataFrame or
relation that you pass in, so any version you already use works.

## Data you can fill from

| Input | Example |
|-------|---------|
| `list[dict]` | `[{"Name": "Widget", "Price": 4.99}]` |
| `dict[str, list]` | `{"Name": ["Widget"], "Price": [4.99]}` |
| `pandas.DataFrame` | `pd.read_csv("sales.csv")` |
| `polars.DataFrame` | `pl.read_csv("sales.csv")` |
| `duckdb.DuckDBPyRelation` | `duckdb.sql("FROM 'sales.csv'")` |

Tables are matched to your data by column name, so the DataFrame order does not
matter, and every format produces the same workbook.

## Visual example

**Before filling:**

![Before fill](https://hirolau.github.io/fillxl/screenshots/before_fill.png)

**After filling — data, chart and values populated:**

![After fill](https://hirolau.github.io/fillxl/screenshots/after_fill.png)

[See the code that produced this](https://hirolau.github.io/fillxl/#visual-example).

## Choosing an engine

`engine="auto"` (the default) picks the engine that can do the job:

| Template → output | Engine used |
|-------------------|-------------|
| `.xlsx` → `.xlsx`, `.xlsm` → `.xlsm` | `openpyxl` (no Excel needed) |
| `.xlsb`/`.xls` template, or a format conversion | `win32com` when Excel is available |
| anything else, without Excel | `openpyxl`, whose error says what is missing |

Three engines are available, and `engine=` selects one explicitly:

| Engine | Needs Excel | Keeps spilled arrays and everything else |
|--------|:---:|---|
| `openpyxl` (default) | no | drops parts the object model cannot hold |
| `fillxl` (experimental) | no | **yes** — edits the OOXML package, never rewrites what it does not understand |
| `win32com` | yes | yes (Windows, Excel required) |

`openpyxl` runs anywhere; `win32com` drives Microsoft Excel on Windows (pywin32
installs with fillxl there) and keeps what openpyxl drops (modern charts,
slicers, macro-heavy workbooks). `engine="fillxl"` is the third way: it needs
no Excel, works on any platform, and keeps spilled dynamic arrays
(`SEQUENCE`, `FILTER`, `SORT`, …), cell metadata, slicers and VBA by writing
only the parts a fill has to change. All engines write the same values for the
same template and data; the
[Backends](https://hirolau.github.io/fillxl/backends/) page lists what still
differs.

> **`engine="fillxl"` is experimental and opt-in.** `engine="auto"` never
> selects it, so a fill only uses it when you ask for it. It keeps spilled
> dynamic arrays, but **it does not calculate them**: the spill is written as
> the template had it, and Excel grows and computes it when it opens the file.
> Until then, anything that reads cached values — `pandas`, `openpyxl` with
> `data_only=True`, previews, search indexers — sees the old numbers or nothing.

**Its limits:**

- `.xlsx` and `.xlsm` only, with no format conversion (use `win32com` for
  `.xlsb`, `.xls` and converting between formats).
- Encrypted workbooks cannot be opened.
- Pivot tables are only flagged for refresh on open; fillxl cannot pivot.
- `chartEx` ranges are not updated.
- Pivot sources and data validations do not follow a table resize — no engine
  does this.

### Cached values

The fillxl engine does not calculate. Formulas it writes have no cached result,
and every other cached result in the workbook — spills included — stays stale
until Excel opens and recalculates the file; the workbook is flagged
`fullCalcOnLoad` so that happens on open. A spill whose size depends on the
table therefore grows only when Excel recalculates.

Tools that read cached values (`pandas`, `openpyxl` with `data_only=True`, file
previews, search indexers) see blanks or old numbers in a fillxl output.
Dynamic arrays survive on Linux, but they are not computed there either. Use
`engine="win32com"` when the saved file must contain computed values.

## Know the edge cases

- One cell holds one value. A `list`, `dict`, `bytes` or a whole DataFrame in a
  named cell or a table cell raises `CellValueError` naming the position,
  instead of writing its text dump; put rows and columns in `tables=`.
- `None` and `NaN` are missing values and become empty cells. ±Infinity cannot
  be stored either, but it warns rather than disappearing quietly.
- Every warning is a `fillxl.FillXLWarning` subclass (`TimezoneWarning`,
  `DanglingReferenceWarning`, `LossyTemplateWarning`), so
  `warnings.simplefilter("error", fillxl.FillXLWarning)` makes data problems
  fail fast instead of quietly changing the output.
- `openpyxl` writes no cached formula results; the output asks Excel to
  recalculate on open (`engine="win32com"` saves computed values).
- `openpyxl` drops some modern chart types, slicers and similar parts, and says
  so: `on_lossy="warn"` (default) warns, `on_lossy="raise"` fails the fill.
  Use `engine="fillxl"` when the template has spilled dynamic arrays or other
  parts openpyxl cannot write: it keeps them.
- `engine="fillxl"` (experimental) keeps spilled arrays and everything else
  unchanged, but it computes nothing: Excel fills the spill and recalculates
  every formula when it opens the file, so cached values stay stale until then.
- `delete_sheets` does not rewrite formulas, charts, names or pivots that point
  at the deleted sheet; fillxl warns with `DanglingReferenceWarning` because
  Excel shows `#REF!` there.
- `hide_underscore_sheets` is **on** by default, so a sheet named `_data` is
  hidden unless you opt out. Hiding or deleting every sheet raises, because
  Excel requires one visible sheet — set `hide_underscore_sheets=False` for a
  workbook whose only sheet is `_report`.

Dates, value types, sheet handling, warning categories, errors and all
`FillXLConfig` options are documented at
[hirolau.github.io/fillxl](https://hirolau.github.io/fillxl/).

## Development

```bash
git clone https://github.com/hirolau/fillxl.git
cd fillxl
uv sync --all-extras
uv run pytest
```

The 0.x line leaves room for breaking changes; 1.0 signals stability. Supported:
Python 3.11–3.15 on Windows and Linux (macOS best-effort), with the data
libraries above on their minimum and newest major release (pandas 2 and 3,
polars 0.20 and 2, duckdb 1 and 2). See
[Development](https://hirolau.github.io/fillxl/development/) for the test
matrix, type checks and supported data-library versions.

**License:** [MIT](https://github.com/hirolau/fillxl/blob/main/LICENSE)