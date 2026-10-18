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
| `template` | `str \| Path` | Path to the Excel template file. |
| `output` | `str \| Path` | Path where the filled workbook will be saved. |
| `values` | `dict[str, Any] \| None` | Map of named cell → value to fill. |
| `tables` | `dict[str, TabularData] \| None` | Map of table name → data to fill. |
| `engine` | `Engine` | Backend engine: `"auto"`, `"openpyxl"`, or `"win32com"`. Default: `"auto"` (openpyxl). |
| `config` | `FillXLConfig \| None` | Advanced configuration (see below). |

### Returns

`Path` — the path to the saved output file.

### Which engine does `"auto"` pick?

`"auto"` selects the **openpyxl** backend: it works everywhere and needs no
Excel installation. Pass `engine="win32com"` explicitly to drive Microsoft
Excel through COM — that is the engine to use for `.xlsm`/`.xlsb` templates
and for chart types openpyxl cannot write.

---

## `FillXLConfig`

Fine-grained control over fill behaviour.

```python
from fillxl import FillXLConfig

config = FillXLConfig(
    column_validation_strategy="strict",
    escape_formulas=False,  # only for trusted input that may contain formulas
    active_sheet="Report",
    delete_sheets=["RawData"],
)
```

### Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `column_validation_strategy` | `str` | `"all_but_formulas"` | Strategy for mismatched columns (see below). |
| `escape_formulas` | `bool` | `True` | Write strings starting with `=` as literal text instead of formulas. Disable only for trusted data that is meant to contain formulas. |
| `refresh_pivot_tables` | `bool` | `True` | Refresh pivot caches after filling. |
| `extend_conditional_formatting` | `bool` | `True` | Resize conditional formatting ranges with the table. |
| `copy_table_styles` | `bool` | `True` | Copy the first data row's styles to new rows. |
| `hide_underscore_sheets` | `bool` | `True` | Hide sheets whose names start with `_`. |
| `delete_sheets` | `Sequence[str] \| None` | `None` | Sheet names to delete (stored as a tuple). |
| `active_sheet` | `str \| None` | `None` | Name of the sheet to make active after filling. |

---

## Column Validation Strategies

| Strategy | Behaviour |
|----------|-----------|
| `"strict"` | Raise `ValueError` if data columns don't match table columns. |
| `"all_but_formulas"` | Like `"strict"`, but exempt calculated/formula columns. Extra data columns are allowed. |
| `"all_but_formulas_strict"` | Like `"all_but_formulas"`, but extra data columns raise. |
| `"drop"` | Skip table columns that have no data column and ignore extra data columns. |
| `"warn"` | Log warnings about mismatches but continue filling. |
| `"none"` | No validation — best-effort fill. |

Table columns without a matching data column are cleared (or keep their
calculated-column formula), so sample values from the template do not leak
into the output.

---

## Values and Excel's type coercion

Excel re-parses strings that look like typed input: `"00123"` becomes the
number `123` and `"1/2"` becomes a date. Both backends store such strings as
**text**, so the two engines agree and no cell gets Excel's "number stored as
text" marker. Error-like strings (`"#N/A"`, `"#DIV/0!"`, …) are stored as
text as well, not as Excel errors.

Strings that start with `=` are kept as literal text by default, so untrusted
data cannot inject formulas. Set `escape_formulas=False` when the data is
trusted and is meant to create formulas.

Empty strings (`""`) become empty cells in both engines, because openpyxl
cannot represent an empty text cell distinctly (it reads back as `None`).

Datetimes lose their timezone (Excel has no timezone support), and both
engines store the same wall-clock time. `date`, `time`, `timedelta`, `Decimal`,
pandas/NumPy scalars and NumPy date/time scalars are all normalised, so nothing
is silently written as text.
