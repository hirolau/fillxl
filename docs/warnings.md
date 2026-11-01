# Warnings

fillxl reports problems it can work around as warnings, and problems it cannot
as exceptions. Every warning fillxl raises is a subclass of
`fillxl.FillXLWarning`, so one filter can turn them all into errors:

```python
import warnings
from fillxl import FillXLWarning

warnings.simplefilter("error", FillXLWarning)   # in tests, or at start-up
```

Two properties matter when you rely on that:

- **The warning belongs to your code.** It is attributed to the caller of
  `fill()`, not to a module inside fillxl, so `pytest.warns(...)`,
  `warnings.filterwarnings(...)` and `-W error` match it.
- **Python's default filter shows each warning once per call site.** A loop
  that calls `fill()` from the same line therefore reports the same problem
  once, not once per file. Use `warnings.simplefilter("always")` (or
  `pytest.warns`) when you need every occurrence, and
  `logging.captureWarnings(True)` when a batch job should collect them in its
  log.

## Categories

| Category | When | What to do |
|---|---|---|
| `FillXLWarning` | Data was supplied but not written, or is stored differently than supplied: an integer with more than 15 significant digits (warned once per column), a ±infinity written as an empty cell (warned once per column), a DataFrame index that is not written, values supplied for a calculated column, all but the first destination of a defined name, a sheet name that matched nothing. | Fix the data or the configuration, or ignore the warning. |
| `TimezoneWarning` | A tz-aware datetime was provided. Excel has no timezone support, so the offset is stripped and the wall-clock time kept. Warned once per fill. | Convert to naive datetimes yourself, or accept the loss. |
| `LossyTemplateWarning` | The openpyxl engine did not write a template feature back (slicers, modern charts, pictures, threaded comments, …). Controlled by `on_lossy`. | Use `engine="win32com"`, or set `on_lossy="ignore"`. |
| `DanglingReferenceWarning` | A sheet deleted by `delete_sheets` is still named by a formula, chart series, defined name or pivot cache. The openpyxl engine cannot rewrite those references, so Excel shows `#REF!` there. | Point the reference elsewhere, remove it, or keep the sheet. |

`on_lossy="raise"` turns the lossy case into `LossyTemplateError` instead of a
warning.

## Errors

| Exception | When |
|---|---|
| `CellValueError` | A value cannot be stored in a cell: text over 32,767 characters, illegal control characters, a lone surrogate, U+FFFE or U+FFFF (which XML 1.0 forbids), a date before 1900-03-01, an integer at or beyond 1e308, or a container (list, dict, bytes, DataFrame, …) where one value belongs. Carries the table, data row and column — or the named cell. |
| `TypeError` | The wrong *kind* of input: `values`/`tables` is not a mapping or `config` is not a `FillXLConfig`, `delete_sheets` is a bare string or contains a non-name, a `list[dict]` row is not a dict, the data is not tabular at all, or `values`/`tables` were passed positionally. |
| `ValueError` | Column validation failed, a table would grow into other content, a defined name points at a table header, the chosen engine cannot write the requested format, or the sheet settings would leave no visible sheet. `CellValueError` is a `ValueError` as well. |
| `IsADirectoryError` | `output=` names an existing directory. |
| `KeyError` | A table or defined name in `values`/`tables` does not exist in the workbook. |
| `PermissionError` | The finished file could not replace `output=`, usually because it is open in Excel. |
| `RuntimeError` | An action Excel refused: a sheet that could not be deleted, a cell written on a protected sheet through COM. `LossyTemplateError` is a `RuntimeError` as well. |

## Logging

fillxl logs the choices it makes (`logger.debug`) and stays quiet at `WARNING`
and above: anything a caller can act on is a warning, not a log line. Configure
the `fillxl` logger if you want the details of a fill.
