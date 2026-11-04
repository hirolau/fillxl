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

Three engines are available. `engine="auto"` (the default) picks one; pass
`engine=` to choose explicitly.

| | `openpyxl` | `fillxl` | `win32com` |
|---|---|---|---|
| Runs on | Any OS | Any OS | Windows with Excel |
| Needs Excel | No | No | Yes |
| Keeps modern features (dynamic arrays, slicers, chartEx, rich data…) | No; losses are reported via `on_lossy` | Yes | Yes |
| Formula results in the saved file | None; Excel recalculates on open | None; cached results are cleared and Excel recalculates on open | Yes, calculated by Excel |
| Pivot tables | Marked to refresh on open | Marked to refresh on open | Refreshed now |
| Formats | `.xlsx` → `.xlsx`, `.xlsm` → `.xlsm` | Same as openpyxl | `.xlsx`, `.xlsm`, `.xlsb`, with conversion between them |
| Status | Stable | Experimental, opt-in | Stable |

- On Linux or in CI with a simple template: use `"auto"` (it picks `openpyxl`).
- On Linux with modern features in the template: use `engine="fillxl"`. The
  file opens correctly in Excel. Tools that cannot calculate (pandas reading
  the output, previewers) see blank formula results.
- When the saved file must contain calculated values, or for `.xlsb` or format
  conversion: use `engine="win32com"`.
- What `"auto"` does, exactly: `openpyxl`, unless Excel is required and
  available. It never picks `fillxl`.

`openpyxl` runs anywhere; `win32com` drives Microsoft Excel on Windows (pywin32
installs with fillxl there) and keeps what openpyxl drops (modern charts,
slicers, macro-heavy workbooks). `engine="fillxl"` is the third way: it needs
no Excel, works on any platform, and keeps spilled dynamic arrays
(`SEQUENCE`, `FILTER`, `SORT`, …), cell metadata, slicers and VBA by writing
only the parts a fill has to change. All engines write the same values for the
same template and data; the
[Backends](https://hirolau.github.io/fillxl/backends/) page lists what still
differs.

> **`.xls` (Excel 97-2003) is no longer supported.** No engine reads or writes
> it: BIFF8 cannot hold the features fillxl exists to preserve, and reading it
> needs Excel anyway. A fill whose template or output ends in `.xls` is refused
> before anything is opened. Open the file in Excel and save it as `.xlsx`,
> `.xlsm` or `.xlsb`.

> **`engine="fillxl"` is experimental and opt-in.** `engine="auto"` never
> selects it, so a fill only uses it when you ask for it. It keeps spilled
> dynamic arrays, but **it does not calculate them**: the spill is written as
> the template had it, and Excel grows and computes it when it opens the file.
> Until then, anything that reads cached values — `pandas`, `openpyxl` with
> `data_only=True`, previews, search indexers — sees the old numbers or nothing.

**Its limits:**

- `.xlsx` and `.xlsm` only, with no format conversion (use `win32com` for
  `.xlsb` and converting between formats).
- Encrypted workbooks cannot be opened.
- Pivot tables are only flagged for refresh on open; fillxl cannot pivot.
- `chartEx` ranges are not updated.
- Pivot sources, defined names and data validations that cover a table body do
  not follow a resize — no engine can rewrite them. fillxl and `openpyxl` warn
  with `UntrackedRangeWarning`; `win32com` inserts rows, so Excel adjusts them
  itself.

### Speed

Speed is not why it exists, but it is the fastest of the three: on the same
2,000-row fill it takes about **0.7 s**, against 1.5 s for `openpyxl` and 2.6 s
for `win32com` — it writes the same values without building or re-serialising a
workbook model. [Benchmarks](https://hirolau.github.io/fillxl/benchmarks/)
has the tables, the method and the script to reproduce them.

### Cached values

No engine that writes without Excel can calculate. The fillxl engine writes
formulas with no cached result at all, and a fill removes every cached result
already in the workbook: the numbers the template stored belong to the
template's data, not to the filled report, and without a dependency graph there
is no telling which of them are still valid. Spills are kept, but their stored
results are cleared the same way. Sheets without formulas are still copied byte
for byte. Excel recalculates the whole file when it opens it — the workbook is
flagged `fullCalcOnLoad` — and a spill whose size depends on the table grows
then.

Tools that read cached values (`pandas`, `openpyxl` with `data_only=True`, file
previews, search indexers) see blanks in a fillxl output — a blank is
deliberate, because a stale number reads like the report's own. Chart *data*
caches (`numCache`/`strCache` points) are the one exception: they keep the
template's points until Excel redraws the chart on open. Dynamic arrays survive
on Linux, but they are not computed there either. Use `engine="win32com"` when
the saved file must contain computed values.

## Know the edge cases

- One cell holds one value. A `list`, `dict`, `bytes` or a whole DataFrame in a
  named cell or a table cell raises `CellValueError` naming the position,
  instead of writing its text dump; put rows and columns in `tables=`.
- `None` and `NaN` are missing values and become empty cells. ±Infinity cannot
  be stored either, but it warns rather than disappearing quietly.
- Every warning is a `fillxl.FillXLWarning` subclass (`TimezoneWarning`,
  `DanglingReferenceWarning`, `UntrackedRangeWarning`, `LossyTemplateWarning`),
  so `warnings.simplefilter("error", fillxl.FillXLWarning)` makes data problems
  fail fast instead of quietly changing the output.
- `openpyxl` writes no cached formula results; the output asks Excel to
  recalculate on open (`engine="win32com"` saves computed values).
- `openpyxl` drops some modern chart types, slicers and similar parts, and says
  so: `on_lossy="warn"` (default) warns, `on_lossy="raise"` fails the fill.
  Use `engine="fillxl"` when the template has spilled dynamic arrays or other
  parts openpyxl cannot write: it keeps them.
- `engine="fillxl"` (experimental) keeps spilled arrays and everything else
  unchanged, but it computes nothing: every cached formula result is removed
  (a reader without a calculation engine sees blanks instead of the template's
  numbers), and Excel recalculates the file when it opens it.
- `delete_sheets` matches sheet names the way Excel does, ignoring case:
  `delete_sheets=["rawdata"]` deletes `RawData`. A name that matches nothing
  warns rather than being ignored.
- A table fill shows the rows it writes: a template saved with a filter applied
  or with rows hidden by hand is filtered back to "all rows" and unhidden, so no
  filled row stays invisible and no `SUBTOTAL(109, …)` total leaves it out.
- A totals row that holds an A1 formula (`=SUM(B2:B2)`) is resized with the
  body, like a formula anywhere else that tracked the table.
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