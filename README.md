# fillxl

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](pyproject.toml)
[![CI](https://github.com/hirolau/fillxl/actions/workflows/ci.yml/badge.svg)](https://github.com/hirolau/fillxl/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/codecov/c/github/hirolau/fillxl)](https://codecov.io/gh/hirolau/fillxl)

**Fill Excel templates with data from Python.**

Designed around named cells and Excel tables — no more fragile cell references like A1 or B3. Name your tables and cells in the template, then populate them cleanly. Charts, formulas, formatting, and conditional formatting are preserved and updated automatically.

This approach lets you freely modify the template without breaking the data-filling step.

```python
import fillxl

fillxl.fill(
    "template.xlsx",
    "output.xlsx",
    tables={
        "SalesTable": [
            {"Product": "Widget", "Sales": 100},
            {"Product": "Gadget", "Sales": 200},
        ],
    },
    values={
        "Title": "Q4 Report",
        "Date": "2024-01-01",
    }
)
```

---

## Installation

```bash
pip install fillxl
```

Optional extras for DataFrame support:

```bash
pip install fillxl[pandas]   # pandas support
pip install fillxl[all]      # pandas + polars + duckdb
```

Without extras, `fillxl` works with plain `list[dict]` or `dict[str, list]` data.

---

## Quick Start

1. **Prepare a template**  
   In Excel, create named cells (Name Manager) for single values and Excel Tables (`Ctrl + T`) for tabular data.

2. **Fill the template**

```python
from fillxl import fill

fill(
    template="report_template.xlsx",
    output="q4_report.xlsx",
    tables={
        "Products": [
            {"Name": "Widget", "Price": 4.99, "Sold": 150},
            {"Name": "Gadget", "Price": 9.99, "Sold": 85},
        ]
    },
    values={
        "ReportTitle": "Q4 2024 Sales",
        "PreparedBy": "Alice",
    },
)
```

Tables are automatically resized, charts updated, and formatting preserved.

3. **Choose a backend (optional)**

By default, `fillxl` uses `openpyxl` (cross-platform, no Excel needed).
On Windows with Excel installed, switch to the `win32com` backend for
maximum fidelity:

```python
from fillxl import fill

fill(
    template="report_template.xlsx",
    output="q4_report.xlsx",
    tables={"Products": [...]},
    values={"ReportTitle": "Q4 2024 Sales"},
    engine="win32com",
)
```

---

## Example

This example fills a stock-price history table into a template with a named chart, title, currency, and target price.

**Before filling — the empty template:**

![Before fill](docs/screenshots/before_fill.png)

**After filling — data, chart, and values populated:**

![After fill](docs/screenshots/after_fill.png)

```python
import numpy as np
import pandas as pd
from pathlib import Path
from fillxl import fill
from fillxl._types import Engine


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
```

---

## Main Features

- Populate named cells and Excel tables
- Automatic table extension and chart range updates
- Support for conditional formatting and pivot table refresh
- Two backends: `openpyxl` (default, cross-platform) and `win32com` (Windows + Excel, best fidelity for modern features)

---

## API

### `fill()`

```python
fill(template, output, values=None, tables=None, *, engine="auto", config=None) -> Path
```

**Key arguments:**

- `template`: Path to `.xlsx` template
- `output`: Output file path
- `values`: `dict[str, Any]` for named cells
- `tables`: `dict[str, list[dict] | DataFrame | ...]`
- `engine`: Backend engine — `"auto"` (default), `"openpyxl"`, or `"win32com"`
- `config`: `FillXLConfig` for advanced options

### Configuration Highlights

`FillXLConfig` offers control over:

- `extend_tables`, `extend_conditional_formatting`, `copy_table_styles`
- `refresh_pivot_tables`
- `column_validation_strategy` (default: `"all_but_formulas"`)
- `hide_underscore_sheets`, `delete_sheets`, `active_sheet`

See the full documentation for details on validation strategies and other options.

---

## Supported Data Sources

- `list[dict]`
- `dict[str, list]`
- `pandas.DataFrame`
- `polars.DataFrame`
- `duckdb.DuckDBPyRelation`

---

## Backends

| Engine     | Cross-platform | Requires Excel | Best for                     |
|------------|----------------|----------------|------------------------------|
| `openpyxl` | Yes            | No             | Most use cases, Linux/macOS  |
| `win32com` | Windows only   | Yes            | Modern charts, full fidelity |

`openpyxl` is the default. Use `win32com` on Windows when you need maximum Excel compatibility (dynamic arrays, complex charts, rendered conditional formatting).

---

## Performance

Benchmark: **5,000 rows × 20 columns**, 10 iterations, filling a template
table with `list[dict]` data. Measured on Windows with Excel installed.

win32com engine is marginally slower due to its Excel overhead, but is comparable to pandas.to_excel function

| Method                          | Engine     | Time (mean ± std) |
|---------------------------------|------------|-------------------|
| `pandas.to_csv`                 | —          | 0.07s ± 0.00s     |
| `polars.write_excel`            | standalone | 1.15s ± 0.01s     |
| `fillxl` (`list[dict]`)         | openpyxl   | 2.26s ± 0.03s     |
| `fillxl` (`list[dict]`)         | win32com   | 2.80s ± 0.03s     |
| `pandas.to_excel`               | openpyxl   | 3.01s ± 0.09s     |


---

## Data Type Consistency

Both backends produce **identical cell values** for the same input data.
All values are routed through a single :func:`~fillxl._normalize.normalize_value`
function that guarantees cross-backend consistency:

| Python type         | Excel representation        |
|---------------------|----------------------------|
| ``str``             | string                     |
| ``int``             | number                     |
| ``float``           | number                     |
| ``bool``            | TRUE / FALSE               |
| ``None``            | blank cell                 |
| ``datetime``        | Excel date/time serial     |
| ``date``            | ``datetime`` (at midnight) |
| ``time``            | string ``"HH:MM:SS"``       |
| ``timedelta``       | ``float`` (fraction of day)|
| ``Decimal``         | ``float``                  |
| ``float('nan')``    | blank cell (``None``)      |
| ``float('inf')``    | blank cell (``None``)      |
| ``pd.Timestamp``    | ``datetime``               |
| ``pd.NA`` / ``NaT`` | blank cell (``None``)      |
| custom objects      | ``str(obj)`` via __str__   |

Time zone info is stripped (wall-clock preserved) with a warning on first occurrence.
This is **tested** in ``tests/test_dtypes_cross_backend.py`` — 100+ parametrized
tests run the same data through both backends and compare results.

---

## Feature Compatibility

| Feature                        | `openpyxl` | `win32com`                    |
|--------------------------------|------------|-------------------------------|
| Named cells                    | ✅         | ✅                            |
| Tables & resize                | ✅         | ✅                            |
| Chart range updates            | ✅         | ✅ / Experimental             |
| Conditional formatting         | ✅         | ✅                            |
| Calculated columns             | ✅         | ✅                            |
| Pivot table refresh            | ✅         | ✅                            |
| Sheet management (hide/delete) | ✅         | ✅                            |
| Style copying                  | ✅         | ✅                            |
| Formula evaluation             | ❌ *       | ✅ (via Excel)                |
| Dynamic arrays                 | ❌         | ✅                            |
| Modern chart types (waterfall) | ⚠️ Dropped | ✅                            |
| CI tested                      | ✅         | ⚠️ Requires Windows + Excel   |
| Cross-platform                 | ✅         | ❌ Windows only               |

\* Formulas are preserved but not recalculated — Excel recalculates on open.

---

## Known Limitations

- **openpyxl** cannot evaluate formulas (Excel recalculates on open), drops certain modern chart types (waterfall, treemap, sunburst, etc.) with a warning, and does not support dynamic arrays.
- **win32com** requires Windows with Excel installed. It is tested less thoroughly than `openpyxl`.
- **Timezone-aware datetimes** are converted to naive (wall-clock preserved) with a warning — Excel has no native timezone support.
- **Chart range updating** is experimental in the `win32com` backend.
- The library targets **regular Excel tables and named cells** — pivot table sources, data models, and Power Query are not supported.

Users forgive limitations. They dislike surprises. These are documented so you know what to expect.

---

## Contributing

Contributions welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

## Development

```bash
git clone https://github.com/hirolau/fillxl.git
cd fillxl
uv sync --all-extras
uv run pytest
```

---

## Support & Stability

**Supported environments:**

| Environment | Status |
|---|---|
| Python 3.11–3.14 | ✅ Officially supported |
| Windows (openpyxl) | ✅ Officially supported |
| Windows (win32com) | ✅ Officially supported (requires Excel) |
| Linux (openpyxl) | ✅ Officially supported |
| macOS (openpyxl) | ✅ Best-effort (CI coverage limited) |

**API stability:** fillxl is in beta. Minor API changes may occur before 1.0,
but core concepts (`fill()`, tables, values, backends) are expected to remain
stable. Deprecation warnings will be used when possible.

---

**License:** [MIT](LICENSE)
