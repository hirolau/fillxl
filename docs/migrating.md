# Migrating from 0.6 to 0.7

0.7.0 groups every breaking change into one release, so 0.8 can be additive.
Most projects need only the first two items.

## 1. `values` and `tables` are keyword-only

The positional call announced as deprecated in 0.6.0 is gone.

```python
# 0.6 (and earlier)
fill("template.xlsx", "out.xlsx", values, tables)

# 0.7
fill("template.xlsx", "out.xlsx", values=values, tables=tables)
```

`fill()` also refuses a `config` that is not a `FillXLConfig`, and a
`values`/`tables` mapping whose keys are not strings, instead of failing later
with a confusing "not found" error.

## 2. Containers are refused instead of written as text

One cell holds one value. These used to be written as their `repr`:

```python
# 0.6: writes "   a  b\n0  1  2" into the cell
fill(..., values={"report": dataframe})

# 0.7: TypeError: Named cell 'report': DataFrame is not a single cell value;
#      pass tabular data through tables= instead
fill(..., values={"report": dataframe})
```

The same rule applies inside tables, where the error names the column and the
data row: `list`, `tuple`, `set`, `dict`, `bytes` and DataFrames/Series/
relations are refused (`CellValueError`). Flatten or serialise the value
yourself if the text form is really what you want.

`Enum` members are written as their `value` (not `"Color.RED"`), and
`date`/`datetime` subclasses are written as the exact built-in type.

## 3. Warnings moved out of the log

Problems a caller can act on are warnings now, and every one of them is a
`fillxl.FillXLWarning`:

```python
import warnings
from fillxl import FillXLWarning

warnings.simplefilter("error", FillXLWarning)   # what used to be a log line
```

| 0.6 | 0.7 |
|---|---|
| `logger.warning` — integer has more than 15 significant digits | `PrecisionLossWarning` |
| `logger.warning` — DataFrame index not written | `IgnoredDataWarning` |
| `logger.warning` — data for a calculated column ignored | `IgnoredDataWarning` |
| `logger.warning` — defined name has several destinations | `IgnoredDataWarning` |
| `logger.warning` — sheet to delete / activate not found | `IgnoredDataWarning` |
| `warnings.warn(UserWarning)` — timezone stripped | `TimezoneWarning` |
| `LossyTemplateWarning(UserWarning)` | `LossyTemplateWarning(FillXLWarning)` |

Diagnostics that are not the caller's to act on stay in the log.

## 4. Values Excel cannot store are refused or blanked

- `Decimal("NaN")`, `Decimal("sNaN")`, `Decimal("Infinity")` and a `Decimal`
  too large for a float now become an **empty cell** instead of `nan`/`inf`
  (which made Excel offer to repair the file).
- Text containing a lone surrogate (for example from
  `errors="surrogateescape"`) raises `CellValueError` with the position,
  instead of `UnicodeEncodeError` at save time.
- An integer at or beyond `1e308` raises `CellValueError`: Excel stores
  numbers as doubles and cannot hold it.

## 5. Fail fast

| Situation | 0.6 | 0.7 |
|---|---|---|
| `output` is the same file as `template` | the template was overwritten | `ValueError` |
| `engine="openpyxl"` with a format it cannot write | `ValueError` at `save()` | `ValueError` before the template is loaded |
| `delete_sheets="Sheet1"` | deleted sheets named `S`, `h`, `e`, … | `TypeError` suggesting `["Sheet1"]` |
| `active_sheet="_Data"` with `hide_underscore_sheets=True` | `ValueError` at the end of the fill | `ValueError` when the config is built |
| The output is open in Excel | bare `PermissionError` | `PermissionError` naming the cause |

## 6. Tables behave more like Excel

- A table that would grow into another table or into merged cells raises
  `ValueError` (same message on both engines) instead of producing an
  overlapping table or a cryptic openpyxl error.
- **Empty data no longer clears calculated columns.** The single remaining
  body row keeps its formula, which is what Excel does.
- Shrinking a table with a totals row also resets the formatting of the old
  totals row, so no bold text or rules are left below the table.
- Bubble sizes, multi-level category ranges and custom error bars now follow a
  resized table on the openpyxl path too.
- The lossy check now counts pictures, chart styles and chart colour schemes.
