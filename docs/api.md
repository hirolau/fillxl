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
    *,
    values: dict[str, Any] | None = None,
    tables: dict[str, TabularData] | None = None,
    engine: Engine = "auto",
    config: FillXLConfig | None = None,
) -> Path
```

`values` and `tables` are keyword-only. The 0.5 call
`fill(template, output, values, tables)` was removed in 0.7.0; the CHANGELOG
lists what else changed.

### Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `template` | `str \| Path` | Path to the Excel template file. |
| `output` | `str \| Path` | Path where the filled workbook will be saved. |
| `values` | `dict[str, Any] \| None` | Map of named cell → value to fill. |
| `tables` | `dict[str, TabularData] \| None` | Map of table name → data to fill. |
| `engine` | `Engine` | Backend engine: `"auto"`, `"openpyxl"`, `"fillxl"`, or `"win32com"`. Default: `"auto"` (see below). |
| `config` | `FillXLConfig \| None` | Advanced configuration (see below). |

### Returns

`Path` — the path to the saved output file.

### Which engine does `"auto"` pick?

| Template → output | Engine |
|-------------------|--------|
| `.xlsx` → `.xlsx`, `.xlsm` → `.xlsm` | `openpyxl` — no Excel needed |
| `.xlsb` template, or a different output format | `win32com`, when Excel automation is available |
| the same, without Excel | `openpyxl`, whose error explains what is missing |

Pass `engine="win32com"` explicitly to drive Microsoft Excel through COM — that
is the engine to use for chart types openpyxl cannot write. Both engines write
the same values for the same template and data; a parity test suite fills one
template with both and compares the results.

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
| `extend_data_validation` | `bool` | `True` | Resize data-validation ranges (dropdowns and other rules) with the table. |
| `copy_table_styles` | `bool` | `True` | Copy the first data row's styles to new rows. |
| `hide_underscore_sheets` | `bool` | `True` | Hide sheets whose names start with `_`. On by default, so opt out explicitly to keep such a sheet visible. |
| `delete_sheets` | `Sequence[str] \| None` | `None` | Sheet names to delete (stored as a tuple). References to the deleted sheets are **not** rewritten; see below. |
| `active_sheet` | `str \| None` | `None` | Name of the sheet to make active after filling. |
| `on_lossy` | `"ignore" \| "warn" \| "raise"` | `"warn"` | What to do when the openpyxl engine drops template features. |

---

## Column Validation Strategies

| Strategy | Behaviour |
|----------|-----------|
| `"strict"` | Raise `ValueError` if data columns don't match table columns. |
| `"all_but_formulas"` | Like `"strict"`, but exempt calculated/formula columns. Extra data columns are allowed. |
| `"all_but_formulas_strict"` | Like `"all_but_formulas"`, but extra data columns raise. |
| `"warn"` | Emit one `FillXLWarning` naming the missing and extra columns, then continue filling. |
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
engines store the same wall-clock time.

`datetime.time` values are written as a **fraction of a day** (`0.385763…` for
09:15:30), which is Excel's own time representation. Arithmetic such as
`=A2*24` therefore works; set the template cell's number format (for example
`hh:mm:ss`) to control how the value is displayed.

`date`, `timedelta`, `Decimal`, pandas/NumPy scalars and NumPy date/time
scalars are all normalised, so nothing is silently written as text.

### Missing numbers and infinities

`None`, pandas `NA`/`NaT` and `NaN` — a Python `float('nan')`, a NumPy scalar
or a `Decimal("NaN")` — all mean "missing" and are written as an empty cell,
without a warning.

±Infinity is different: Excel has no way to hold it, so the cell is written
empty, but fillxl reports it with a `FillXLWarning` naming the table, row and
column (once per column, or once per named cell). A number that disappears
silently is data loss; filter on `FillXLWarning` if you want it to fail
instead:

```python
import warnings
from fillxl import FillXLWarning

warnings.simplefilter("error", FillXLWarning)
```

### Dates

A `date` gets the number format `yyyy-mm-dd` and a `datetime` gets
`yyyy-mm-dd h:mm:ss`, in both engines, matching what openpyxl does on its own.
A cell that already has a date format keeps it — format the cell in Excel and
fillxl leaves it alone.

Dates before **1900-03-01** raise `CellValueError`. Excel counts a 29 February
1900 that never existed, and the two engines cannot agree on what a serial
number before that day means.

### Sheets

`FillXLConfig` can hide, delete and activate sheets after filling:

```python
FillXLConfig(
    hide_underscore_sheets=True,   # the default: hide visible sheets named "_…"
    delete_sheets=["RawData"],
    active_sheet="Report",
)
```

Sheet names are matched exactly, in both engines: `delete_sheets=["report"]`
does not delete `Report`, and `active_sheet="report"` leaves the workbook's
active sheet alone. Both cases warn instead of acting on a near-miss, so a
typo cannot delete the wrong sheet.

- Hiding or deleting every visible sheet raises `ValueError`; Excel requires at
  least one visible sheet. A workbook whose only sheet is `_report` therefore
  fails under the default config — the message says to set
  `hide_underscore_sheets=False`.
- When the active sheet is hidden or deleted, the first remaining visible sheet
  becomes active. Exactly one tab is left selected.
- Making a hidden sheet active raises `ValueError`.
- A sheet whose deletion fails raises `RuntimeError`: it must not silently stay
  in the report.
- `veryHidden` sheets stay `veryHidden`, and chart sheets can be deleted or
  activated.
- `delete_sheets` removes the sheet and nothing else: formulas, chart series,
  defined names and pivot caches that refer to it still name it, and Excel shows
  `#REF!` where they do. Every engine warns with `DanglingReferenceWarning` —
  `win32com` included, where the warning is based on the workbook's state
  before Excel deletes the sheet. Update those references yourself, or keep the
  sheet.

### Defined names

A name in `values=` is matched exactly first, then ignoring case, because
Excel does the same: `values={"reporttitle": "Q4"}` fills the name
`ReportTitle`. An exact match always wins, so a `Title` and a `title` that live
in different scopes keep their own destinations.

A name may not point at a **table header cell**: renaming a header would leave
the table's column name out of step with the cell (openpyxl) or silently break
the data-to-column mapping the same fill is using (Excel), so both engines
raise `ValueError`. Point the name at another cell, or rename the column in
Excel.

### Ragged `list[dict]` data

The columns of a `list[dict]` are the ordered union of the keys of **all**
rows (first-seen order); a key that only appears in a later row is still
written, and rows without it get `None`. A non-dict row raises `TypeError`
naming the row number.

### pandas indexes

The DataFrame index is not written. A named, multi-level or non-integer index
probably carries data, so fillxl warns (suggesting `df.reset_index()`) when it
drops one. A default or filtered integer index is silent.

### Containers

One cell holds one value, so a `list`, `tuple`, `set`, `dict`, `bytes` or a
tabular object (DataFrame, Series, relation) where a value is expected raises
`CellValueError` rather than writing its `repr`. The message names the table,
data row and column, or the named cell — and for a named cell it points at
`tables=`.

---

## Errors

| Exception | Meaning |
|-----------|---------|
| `CellValueError` | A value Excel cannot store: text longer than 32,767 characters, an illegal control character, a lone surrogate, a character XML 1.0 forbids such as U+FFFE or U+FFFF, a date before 1900-03-01, an integer at or beyond 1e308, or a container where one value belongs. It is a `ValueError` and carries the table, data row and column (or the named cell). |
| `LossyTemplateError` | The openpyxl engine dropped template features and `FillXLConfig(on_lossy="raise")` was set. Raise it before swapping in the output, so the previous output file is left untouched. |
| `TypeError` | `values=` or `tables=` is not a mapping, `config=` is not a `FillXLConfig`, or the tabular data is not one of the supported shapes. |
| `IsADirectoryError` | `output=` is an existing directory. |
| `PermissionError` | The finished file could not replace `output=`, usually because it is open in Excel. |

Anything that can be seen from the arguments alone — a mistyped mapping, a
directory used as `output=`, an unknown engine or an output format the engine
cannot write — is reported before the template is opened. Problems that need
the workbook, the data or the finished file (`KeyError`, `CellValueError`, a
column mismatch, `LossyTemplateError`, `PermissionError`) arrive later, but
never after the destination has been replaced.

Integers with more than 15 significant digits are not errors: Excel rounds
them, so fillxl warns once per column (or name). Pass such identifiers as text.

Named cells raise `KeyError` only when the name does not exist. A name whose
first destination is a multi-cell range raises `ValueError`; a write failure
(such as a protected sheet through COM) raises `RuntimeError`.

### The output file

`template` and `output` may be the same path: the template is copied to a
temporary directory before it is opened, so an in-place fill replaces the
template with the filled workbook.

`output` is written only after the whole fill has succeeded — saving, closing
the engine and the lossy check all run first, against a staging file in the
same directory. If anything fails, an `output` that already existed is left
exactly as it was and the staging file is removed. This is a guarantee against
*operation* failures, not against power loss: the swap itself is a rename in
the destination directory, which a crash can interrupt.

The openpyxl engine runs a lossy-feature check against the staging file before
the atomic swap. It compares part and element **counts** (modern charts,
slicers, shapes, pictures, embedded objects, VBA, …), so it catches dropped
parts but cannot prove that a preserved part is unchanged. The check reports a
`LossyTemplateWarning`:

```python
import warnings
from fillxl import LossyTemplateWarning

warnings.filterwarnings("error", category=LossyTemplateWarning)
```

[Warnings](warnings.md) lists every category and every exception fillxl
raises.
