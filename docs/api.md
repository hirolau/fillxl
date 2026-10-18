# API Reference

## `fill()`

The main entry point for filling Excel templates.

```python
from fillxl import fill
```

### Signature

```python
fill(
    template: str | Path,
    output: str | Path,
    values: dict[str, Any] | None = None,
    tables: dict[str, TabularData] | None = None,
    *,
    engine: Engine = "auto",
    config: FillXLConfig | None = None,
) -> Path
```

### Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `template` | `str \| Path` | Path to the Excel template file (`.xlsx`). |
| `output` | `str \| Path` | Path where the filled workbook will be saved. |
| `values` | `dict[str, Any] \| None` | Map of named cell → value to fill. |
| `tables` | `dict[str, TabularData] \| None` | Map of table name → data to fill. |
| `engine` | `Engine` | Backend engine: `"auto"`, `"openpyxl"`, or `"win32com"`. Default: `"auto"`. |
| `config` | `FillXLConfig \| None` | Advanced configuration (see below). |

### Returns

`Path` — the path to the saved output file.

---

## `FillXLConfig`

Fine-grained control over fill behaviour.

```python
from fillxl import FillXLConfig

config = FillXLConfig(
    validate_columns="strict",
    active_sheet="Report",
    hide_sheets=["RawData"],
)
```

### Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `validate_columns` | `str` | `"strict"` | Strategy for mismatched columns: `"strict"`, `"drop"`, `"warn"`, `"all_but_formulas"`, `"none"`. |
| `active_sheet` | `str \| None` | `None` | Name of the sheet to make active after filling. |
| `hide_sheets` | `list[str] \| None` | `None` | Sheet names to hide after filling. |
| `delete_sheets` | `list[str] \| None` | `None` | Sheet names to delete after filling. |
| `refresh_pivots` | `bool` | `True` | Refresh pivot tables after filling. |
| `extend_conditional_formatting` | `bool` | `True` | Extend conditional formatting ranges when tables grow. |

---

## Column Validation Strategies

| Strategy | Behaviour |
|----------|-----------|
| `"strict"` | Raise `ValueError` if data columns don't match table columns. |
| `"drop"` | Drop extra data columns, fill missing ones with `None`. |
| `"warn"` | Warn about mismatches but continue filling. |
| `"all_but_formulas"` | Like `"strict"`, but exempt calculated/formula columns. |
| `"none"` | No validation — best-effort fill. |
