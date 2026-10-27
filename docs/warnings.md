# Warnings

fillxl reports problems it can work around as warnings, and problems it cannot
as exceptions. Every warning fillxl raises is a subclass of
`fillxl.FillXLWarning`, so one filter can turn them all into errors:

```python
import warnings
from fillxl import FillXLWarning

warnings.simplefilter("error", FillXLWarning)   # in tests, or at start-up
```

The warning is attributed to *your* code — the caller of `fill()` — not to a
module inside fillxl, so `pytest.warns(...)`, `warnings.filterwarnings(...)`
and `-W error` match it.

## Categories

| Category | When | What to do |
|---|---|---|
| `PrecisionLossWarning` | An integer has more than 15 significant digits, so Excel rounds it. Warned once per column. | Pass identifiers as text. |
| `IgnoredDataWarning` | Data was supplied but not written: a DataFrame index, values for a calculated column, all but the first destination of a defined name, a sheet name that matched nothing. | Fix the data or the configuration, or ignore the warning. |
| `TimezoneWarning` | A tz-aware datetime was provided. Excel has no timezone support, so the offset is stripped and the wall-clock time kept. Warned once per fill. | Convert to naive datetimes yourself, or accept the loss. |
| `LossyTemplateWarning` | The openpyxl engine did not write a template feature back (slicers, modern charts, pictures, threaded comments, …). Controlled by `on_lossy`. | Use `engine="win32com"`, or set `on_lossy="ignore"`. |

`LossyTemplateError` (raised with `on_lossy="raise"`) is not a warning; it is
the same message as an exception.

## Errors

| Exception | When |
|---|---|
| `CellValueError` | A value cannot be stored in a cell: text over 32,767 characters, illegal control characters, a lone surrogate, a date before 1900-03-01, an integer at or beyond 1e308, or a container where one value belongs. Carries the table, data row and column (or the named cell). |
| `TypeError` | A named cell was given a list, dict, bytes or tabular object; `values`/`tables` are not mappings of string keys; `config` is not a `FillXLConfig`; `delete_sheets` is a bare string. |
| `ValueError` | Column validation failed, a table would grow into other content, `output` is `template`, the chosen engine cannot write the requested format, or the sheet settings would leave no visible sheet. |
| `KeyError` | A table or defined name in `values`/`tables` does not exist in the workbook. |

## Logging

fillxl logs the choices it makes (`logger.debug`) and stays quiet at
`WARNING` and above: anything a caller can act on is a warning, not a log
line. Configure the `fillxl` logger if you want the details of a fill.
