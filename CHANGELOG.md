# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.8.0] - 2026-11-01

Correctness release, not yet 1.0. The public API is unchanged since 0.7.1 —
`fill()`, `FillXLConfig`, the exception classes and the warning classes keep
their signatures, defaults and meanings — so upgrading needs no code changes.
The 0.x line still leaves room for breaking changes.

### Fixed
- openpyxl: a calculated column declared with A1 references writes each row its
  own formula. `calculatedColumnFormula="B5*C5"` used to be copied into every
  row, so every row computed the first row's total (49.9, 49.9, 49.9 where
  Excel shows 49.9, 29.97, 3.5); relative rows now move per row and
  `$`-anchored rows stay put. Structured references (`=[@Qty]*[@Price]`) are
  row-relative already and are written unchanged.
- openpyxl: a worksheet formula whose range covers the table body row for row
  — `=SUM(B4:B4)` for a one-row body — is resized with the table, on the
  table's sheet and on any sheet that refers to it, as Excel's resize does.
  Single-cell references, ranges that start or end outside the body, string
  literals and references to other sheets are left alone.
- Conditional formatting now follows the table by Excel's rule: a range that
  ends on the table's last data row or starts on its header row is resized,
  and a range that reaches beyond the table is left alone. A rule covering
  `B1:B2` (header plus one body row) used to stay at two rows while Excel
  extended it; a rule reaching past the table could be shrunk to the table.
- Text containing U+FFFE or U+FFFF raises `CellValueError` with the table, row
  and column. openpyxl wrote the character raw, `fill()` returned normally and
  replaced the output with a workbook that neither openpyxl nor Excel could
  open. Excel does not preserve the character either — it rewrites it as the
  literal text `_xFFFE_` — so both engines refuse it. The other noncharacters
  XML allows (U+FDD0, U+FDEF, U+1FFFE, U+10FFFF, U+2028, …) are still written
  and round-trip.
- A named cell that points at a table header cell raises `ValueError` instead
  of renaming the header: openpyxl left the table's column name behind (Excel
  refused to open the file) and Excel renamed the column, which invalidated
  the data-to-column mapping the same fill was using.
- A ±infinity written into a table cell or a named cell now warns with
  `FillXLWarning`, naming the cell, instead of quietly becoming a blank
  (`Decimal`, NumPy floats and 0-d arrays included). `NaN` and pandas
  `NA`/`NaT` still mean "missing" and stay silent.
- openpyxl: a defined name is found regardless of case, matching Excel and the
  COM engine. `values={"reporttitle": …}` used to raise `KeyError` with
  `engine="openpyxl"` and fill `ReportTitle` with `engine="win32com"`. An exact
  match still wins.
- `engine="auto"` no longer starts Excel for an output format Excel cannot
  save (a `.txt` name, a suffix-less name, …). The suffix is now checked
  before the template is loaded, exactly as `engine="win32com"` already did,
  instead of failing at `save()` after the whole fill had run. That late
  failure also named the internal staging file — `Cannot save '.fillxl'
  files` — rather than the path the caller passed, and the error now says
  `engine='auto' chose win32com` instead of blaming an engine the caller never
  asked for.
- An `output=` that is an existing directory raises `IsADirectoryError` before
  the template is opened. It used to fail during the final rename and claim
  "permission denied. Is the file open in Excel?".
- `values=` and `tables=` must be mappings, and `config=` a `FillXLConfig`;
  anything else raises `TypeError` naming the argument. A list, a bare string
  or a plain dict of options used to fail with `AttributeError` from deep
  inside the fill, after the template had been loaded.
- The lossy check counts embedded OLE objects (`xl/embeddings/*`). openpyxl
  writes no embeddings, so an embedded Word document, PDF or equation used to
  disappear from the output without any warning; it is now reported by name
  and `on_lossy="raise"` refuses the fill before the swap.
- The DataFrame-index warning says "The DataFrame index is not written" (or
  `named 'rowid'`) instead of printing `[None]`, which read as if the index
  had a value called `None`.
- The error for a workbook whose only sheet is `_report` now says to set
  `hide_underscore_sheets=False`.

### Documentation
- `fill()` documents every exception it can raise, and that `template` and
  `output` may be the same path.
- `docs/backends.md` describes which formulas a fill rewrites — calculated
  columns and formulas whose range covers the table body — and which it leaves
  alone, and lists embedded objects among the features the lossy report
  counts.
- `docs/api.md` documents the missing-value policy (`NaN` is a blank cell) next
  to the infinity warning, and that defined names match case-insensitively
  while sheet names do not.
- `docs/warnings.md` lists the new infinity warning, the U+FFFE/U+FFFF rule
  and the table-header error.

### Tests
- The failure-recovery chapter now covers a destination workbook that already
  exists: a fault injected at saving, at closing, at the lossy check and at the
  final swap each leave it byte-for-byte intact, leave no staging file behind
  and do not block the retry. In-place fills (same template and output path)
  are covered in both directions.
- `tests/integration/test_integrity_guards.py` fills one template with both
  engines and checks the calculated column rows, the resized formulas, the
  conditional-formatting rule, the U+FFFE refusal, the infinity warning, the
  table-header refusal and case-insensitive names. Where Excel is available a
  test opens the openpyxl output in Excel and asserts the computed totals
  (49.9, 29.97, 3.5) rather than the formula text alone.

## [0.7.1] - 2026-11-01

### Changed
- Documentation no longer presents the data-library extras as a choice users have
  to make. `pip install fillxl` is enough: pandas, Polars and DuckDB support
  works with the library version you already have, because fillxl detects the
  object type you pass and never imports those libraries itself. The extras are
  no longer listed in the README.
- On Windows, pywin32 now ships with fillxl, so `engine="win32com"` works after
  a plain `pip install fillxl`. An environment marker keeps it out of Linux and
  macOS installs.

### Removed
- The `win32com` extra: pywin32 is a base dependency on Windows, so there is
  nothing left to opt into. `fillxl[all]` no longer lists it either.

### Fixed
- The `duckdb` extra installs on Python 3.15. duckdb 2.0 ships `cp312-abi3`
  wheels, so the `python_version < '3.15'` marker on the extra is gone and
  duckdb data sources are available on every supported Python; `uv.lock`
  resolves duckdb 2.0. Only the duckdb 1.x line still has no 3.15 wheels.

## [0.7.0] - 2026-11-01

Breaking, but safer. The breaking changes are grouped into this one release so
later releases can be additive where possible.

### Changed (breaking)
- `fill()` takes `values` and `tables` as keyword-only arguments. The
  positional call announced as deprecated in 0.6.0 is gone:

  ```python
  fill("template.xlsx", "out.xlsx", values, tables)   # 0.6
  # TypeError: fill() takes 2 positional arguments but 4 were given
  fill("template.xlsx", "out.xlsx", values=values, tables=tables)   # 0.7
  ```

- One cell holds one value: a `list`, `tuple`, `set`, `dict`, `bytes` or a
  tabular object (DataFrame, Series, relation) raises `CellValueError` with
  the position, instead of writing its `repr` into the cell. For a named cell
  the message suggests `tables=`:

  ```python
  fill(..., values={"report": df})    # 0.6: writes the DataFrame's text dump
  # 0.7: CellValueError: Named cell 'report': DataFrame is not a single cell
  #      value; put rows and columns in tables= instead
  ```

  `Enum` members are written as their `value`, and `date`/`datetime`
  subclasses are written as the exact built-in type (openpyxl looks number
  formats up by type).

- Data problems a caller can act on are warnings instead of log records, and
  every warning is a `fillxl.FillXLWarning` (with `TimezoneWarning` and the
  re-parented `LossyTemplateWarning` as subclasses). Lossy output, a DataFrame
  index that is not written, data supplied for a calculated column, extra
  destinations of a defined name, a sheet name that matches nothing and an
  integer Excel will round are all catchable now:

  ```python
  warnings.simplefilter("error", fillxl.FillXLWarning)
  ```

- Empty data no longer clears calculated columns: the single remaining body
  row keeps its formula, as Excel does.
- A `Decimal` that is NaN, sNaN, infinite or too large for a float becomes an
  empty cell instead of `nan`/`inf`, which made Excel offer to repair the file.
- A format the chosen engine cannot write is refused before the template is
  loaded, not at `save()` after the fill has run. The message says the engine
  "cannot read" an unsupported template and "cannot write" an unsupported
  conversion.
- `delete_sheets` must be a list of names; a bare string was iterated and
  deleted sheets named `S`, `h`, `e`, â€¦
- `active_sheet` starting with `_` together with `hide_underscore_sheets=True`
  is refused when the config is built, not at the end of the fill.

### Fixed
- openpyxl: a table growing into another table or into merged cells is
  refused with the same `ValueError` win32com raises, instead of producing an
  overlapping table or an openpyxl "read-only" `AttributeError`.
- win32com: an empty payload no longer clears calculated columns, which could
  make Excel drop the calculated-column definition.
- openpyxl: shrinking a table with a totals row resets the formatting of the
  old totals row instead of leaving bold text and rules below the table.
- openpyxl: chart updates now include bubble sizes, multi-level category
  ranges and custom error-bar ranges, matching what win32com rewrites.
- Text containing a lone surrogate raises `CellValueError` with the position,
  instead of `UnicodeEncodeError` at save time.
- Integers at or beyond `1e308` raise `CellValueError`; integers outside
  Â±2Â³Â¹ are handed to COM as floats, which loses nothing Excel can store.
- win32com: `TemporaryDirectory` cleanup no longer fails when Excel still
  holds the temporary copy.
- A `PermissionError` from the final rename explains that the output is
  probably open in Excel.
- The lossy check counts pictures (`xl/media/*`, `<xdr:pic>`), chart styles and
  chart colour schemes, so a template that loses them now warns.
- `on_lossy` no longer reports features that only sheets removed by
  `delete_sheets` carried (pictures, charts). Those parts are gone on purpose,
  and `on_lossy="raise"` no longer fails a fill that lost nothing.
- A 0-d NumPy array in a `dict[str, column]` payload is treated as a scalar
  instead of failing with "iteration over a 0-d array".
- `_normalize` rule numbering in the docstrings matches the code.
- Formulas written into calculated columns through `win32com` strip the whole
  XML prefix chain, so `_xlfn._xlws.SORT(` becomes `SORT(` instead of the
  invalid `_xlws.SORT(`.
- A pivot cache the `win32com` engine could not refresh (or could not mark for
  refresh on open) is now a warning instead of a line in the log.
- A `polars.Series` used as a column in `dict[str, column]` table data is
  accepted like a pandas `Series`, instead of being mistaken for a scalar and
  raising `TypeError`.

### Added
- `FillXLWarning` and `TimezoneWarning`, exported from `fillxl`.
- `DanglingReferenceWarning`: when `delete_sheets` leaves a formula, chart
  series, defined name or pivot cache pointing at a removed sheet, the fill now
  says where, because Excel opens those as `#REF!`.
- `docs/warnings.md`: every warning and error fillxl raises, and how to filter
  them.
- `docs/backends.md` now carries the feature-by-feature engine comparison.
- The version matrix covers two release lines per data backend — the minimum
  (pandas 2, polars 0.20, duckdb 1) and the newest major (pandas 3 in
  `uv.lock`, polars 2 and duckdb 2 pinned) — in nox and in a `backends` CI job
  that pins them explicitly. A test that needs one of those libraries carries
  the `polars`/`duckdb` marker, so the pinned sessions select it however it is
  named instead of relying on the test name.

### Verified, not changed
- A sorted table's `sortState` is left exactly as the template had it. Excel 16
  does the same when it resizes a table (in either direction), and the
  workbook still opens without a repair prompt, so rewriting the reference
  would only make the engines disagree.

## [0.6.0] - 2026-10-21

The two engines now produce the same workbook for the same template and data,
and a parity test suite checks it.

### Fixed
- win32com: pivot tables refreshed before formulas were recalculated, so pivots on
  calculated columns showed stale results. Formulas are now also calculated before saving.
- win32com: tables on a sheet other than the active one could not be resized.
- win32com: charts on chart sheets were not updated.
- Both: charts on sheets with non-ASCII or apostrophe names, non-contiguous series, and
  one-row tables (single-cell references) were not updated correctly.
- Both: dates showed as midnight timestamps (openpyxl) or in a locale-dependent format
  (win32com). Dates now get `yyyy-mm-dd` and datetimes `yyyy-mm-dd h:mm:ss`, unless the
  template cell already has a date format.
- openpyxl: rich text in templates (partly bold cells) was flattened to plain text.
- Both: table columns whose names contain line breaks now match data columns.
- Hiding sheets: a hidden sheet could stay active, very hidden sheets became merely
  hidden, and openpyxl outputs could open with sheets grouped.
- Deleting sheets could leave the wrong sheet active; win32com reported every failure
  as "not found".
- `set_active_sheet` picked the wrong sheet (openpyxl) when chart sheets came first.

### Changed
- `engine="auto"` uses Excel (win32com) on Windows when openpyxl cannot do the job:
  `.xlsb`/`.xls` templates, or saving in another format.
- Dates before 1900-03-01 raise `CellValueError` (Excel's calendar is wrong before then).
- `datetime.date` values stay dates through normalisation.
- Column validation errors and logs name the table.
- A `dict[str, column]` that mixes columns and scalars raises `TypeError` instead of
  splitting strings into characters.
- Hiding or deleting every visible sheet raises `ValueError`; making a hidden sheet
  active raises `ValueError`; a failed sheet deletion raises `RuntimeError`.
- win32com matches table names case-sensitively, like openpyxl.
- Lossy-template warnings use the new `LossyTemplateWarning` category and are reported
  once (they are no longer also logged, and the load-time chartEx warning is gone).
- Requires openpyxl >= 3.1.

### Deprecated
- Passing `values` and `tables` to `fill()` positionally. Use keywords; positional use
  stops working in 0.7.0.

### Removed
- `TabularAdapter.is_tabular` and `TabularAdapter.iter_rows` (internal, unused).

## [0.5.0] - 2026-10-20

### Changed (breaking)

- `datetime.time` values are written as a fraction of a day (Excel's own time
  representation) instead of a datetime on 2000-01-01. Arithmetic such as
  `=A2*24` now works. Format the template cell (e.g. `hh:mm:ss`) to control
  the display.
- The openpyxl engine saves `.xlsx` templates only as `.xlsx` and `.xlsm`
  templates only as `.xlsm`.

### Fixed

- Calculated table columns are never overwritten by data, on either engine;
  both engines detect them with the same rule. Constant formulas such as
  `=100` are kept.
- win32com no longer erases calculated columns of `.xlsb`/`.xls` templates.
- `list[dict]` keys that first appear in a later row are no longer dropped.
- Input columns whose names collide after string conversion (for example `1`
  and `"1"`, or two `"amount"` labels in a DataFrame) raise `ValueError`
  instead of silently dropping one column. The check runs before any cell is
  written and does not depend on the column-validation strategy.
- Named cells behave identically on both engines (multi-cell ranges raise
  `ValueError`; `KeyError` only for missing names).
- The win32com overflow guard refuses to resize when it cannot read the cells
  below a table.
- Mixed-type table columns no longer share one style object (openpyxl).
- `totalsRowShown` without `totalsRowCount` is no longer treated as a totals
  row.
- Chart axis repair no longer moves right/top or secondary axes, and handles
  scatter charts.

### Security

- win32com opens templates with macros force-disabled, events off and link
  updates off.

### Added

- `.xlsm` templates on the openpyxl engine (VBA preserved).
- Atomic output: a failed fill never touches an existing output file.
- `CellValueError` with table/row/column context for values Excel cannot
  store.
- `FillXLConfig.on_lossy` and `LossyTemplateError`.
- openpyxl output requests a full recalculation on open.
- win32com works from worker threads (`CoInitialize`) and no longer uses the
  clipboard.

## [0.4.0] - 2026-10-18

### Added

- Tables **without a header row** (the table's "Header Row" option is off) are
  supported by both backends. Excel names their columns `Column1`,
  `Column2`, â€¦, so that is how the data is addressed, and the data is written
  starting in the table's first row instead of one row below it. win32com
  writes and clears such tables row by row (cell by cell for a single-row
  body), because Excel drops the table when a whole headerless body is
  written in one call.

## [0.3.0] - 2026-10-18

### Added

- `FillXLConfig.escape_formulas` (default `True`): write strings that start
  with `=` as literal text instead of formulas. Formula-like strings from
  untrusted data are a formula-injection vector, so escaping is on by default;
  pass `escape_formulas=False` when the input is trusted and is meant to
  contain formulas.
- Empty tabular payloads (`[]` and `{}`) are accepted and clear the table;
  they no longer raise a misleading `TypeError` about an unsupported type.
- Sheet-scoped defined names are now resolved by the openpyxl backend, matching
  win32com, which finds them through `Workbook.Names`.
- Python 3.15 support: added to the classifiers, the CI test matrix, and the
  nox `PYTHONS` list (the `requires-python` floor stays at 3.11).
- polars 2 support: the dev/test environment and `uv.lock` use polars 2, and
  the nox matrix has an explicit `tests_polars-2` session. The supported
  release lines are now documented as pandas 2â€“3, polars 0.20â€“2, and duckdb
  1.x.
- The `duckdb` extra is limited to Python â‰¤3.14: duckdb currently publishes no
  cp315 wheels. Its tests skip on 3.15 until upstream provides them.

### Changed

- **Both backends** now support tables with a **totals row**: the data body is
  replaced and the totals row is moved to stay directly below it, preserving
  its values, labels, functions and styling. Totals formulas that use
  structured references (for example `SUBTOTAL(109,Table1[Value])`) follow the
  table; formulas written as explicit ranges are left untouched.
- **Both backends** now write every non-empty string as text, so Excel's
  typed-input parser cannot reinterpret the data. win32com does this with
  Excel's text marker (apostrophe prefix, no number-format changes), so
  `"00123"` stays `"00123"` instead of becoming `123` and `"1/2"` no longer
  becomes a date. openpyxl stores strings as text as well, including
  error-like strings such as `"#N/A"` that it would otherwise turn into Excel
  errors.
- **Behaviour change**: empty strings (`""`) become empty cells in both
  engines. openpyxl cannot round-trip an empty text cell (it reads back as
  `None`), so writing one would make the two backends disagree.
- **Both backends**: table cell styles are applied *before* the value, so
  datetime columns get a date format in the openpyxl output as well. Datetime
  table cells now read back as datetimes from both engines instead of a bare
  serial number from openpyxl.
- **win32com**: pivot caches are refreshed *and* marked `refreshOnLoad`, so
  both engines leave a workbook that recalculates its pivots on open.
- The tabular adapter accepts any sequence-like column values for
  `dict[str, ...]` data, including NumPy arrays, tuples and pandas Series.
  Columns are read positionally, so a Series with a non-default index (and
  non-string mapping keys) works, and ragged columns raise a `ValueError` that
  names the offending column instead of being silently truncated.
- `delete_sheets` is stored as a tuple, so a frozen `FillXLConfig` cannot be
  mutated through a caller's list.
- CI runs on every branch; the release workflow now gates publishing on a
  passing test run and on the tag matching `pyproject.toml`.

### Fixed

- NumPy scalars (`np.int64`, `np.int32`, `np.float32`, `np.bool_`, â€¦) are
  converted to native Python values instead of `str(value)`, which used to
  produce the "number stored as text" marker in Excel. `np.datetime64` and
  `np.timedelta64` - including nanosecond precision, where `.item()` returns
  an integer - are converted to `datetime`/`timedelta` as well.
- A failed `load()` no longer leaks a hidden `EXCEL.EXE`: the backend is now
  loaded inside the `try`/`finally` so `close()` always runs.
- `table.autoFilter.ref` is updated together with `table.ref`, avoiding the
  mismatch that makes Excel offer to repair the workbook.
- Empty data leaves a valid table with one cleared body row in both backends
  (openpyxl used to write a header-only range, which Excel rejects).
- Table columns missing from the data no longer keep the template's sample
  values; the cells covered by the fill are cleared (calculated columns keep
  their formulas).
- Conditional formatting ranges are matched by geometry, not by string prefix:
  a rule on `B130:B200` is no longer mistaken for one on `B13` (which produced
  the invalid range `B130:B13`), and multi-area sqrefs such as `B13:B20
  D13:D20` are resized per area instead of being split on `:`.
- Chart data ranges are matched by exact sheet name in both backends, so a
  chart on `Data2` is untouched while `Data` is filled.
- **win32com**: the save format is inferred from the output suffix. Filling an
  `.xlsm` template to `.xlsm` no longer strips its macros by saving `.xlsx`
  (`.xlsb` and `.xls` are supported too; anything else raises `ValueError`).
- **win32com**: datetimes written to named cells are no longer shifted by the
  local timezone offset (they now get the same UTC marking as table cells).
- `fill()` now returns the absolute output path it documents, even when the
  caller passes a relative `output`.
- The timezone warning flag is a `ContextVar`, so two concurrent `fill()` calls
  in different threads no longer share it.
- `__version__` is read from package metadata instead of being hard-coded, so
  it cannot drift from `pyproject.toml`. When the package is not installed
  (for example, when importing from a source tree) it reports
  `"0.0.0+unknown"` instead of pretending to be release `0.0.0`.

### Removed

- `FillXLConfig.extend_tables` - it was never read by any backend, so it
  silently did nothing. Tables are always resized to fit the data.

## [0.2.0] - 2026-07-19

Improved the test suite and added functionality enabled by the multiple-backend design.

### Added

- Validation rejecting non-`.xlsx` files in the openpyxl backend with a clear
  error message, preventing silent corruption of `.xlsm`, `.xlsb`, and other
  macro-enabled or binary formats.
- Local openpyxl type stubs in `typings/` for improved type-checking accuracy
  across mypy, pyrefly, and ty.

### Fixed

- Removed an unused `cast` import to keep the codebase clean.
- **win32com**: Preserved original newlines and indentation in calculated column
  formulas when writing via COM, fixing flattened multi-line `LET`/`LAMBDA`
  formulas.

### Changed

- README tagline now mentions pandas, polars, and duckdb; example code shows
  `pd.read_csv()` usage for a more realistic quick-start.
- README example section reorganized with before/after screenshots and code
  placed between them for better flow.
- Added PyPI badge to README.

## [0.2.0-beta.1] - 2026-07-17

### Added

- Cross-backend datatype tests for openpyxl backend (CI-safe).
- GitHub Actions CI workflow (lint, typecheck, test matrix on 3.11â€“3.14).
- GitHub Actions release workflow (publishes to PyPI on version tags).
- `CHANGELOG.md`, issue templates, PR template.
- Test suite split into logical modules (backends, validation, charts, CF, â€¦).
- All test methods renamed to read as descriptive sentences.
- Cross-backend dtype testing (openpyxl + win32com, 68 tests).
- Golden copy regression tests (openpyxl + win32com).
- Story tests: charts, data types, template preservation, failure recovery, report generation.
- Backend protocol tests and openpyxl unit tests.
- win32com performance optimisation (26Ã— speedup for large tables).
- Coverage: overall 85% (excl. win32com: ~95%).
- `ruff` clean, `mypy` strict - 0 errors.
- Author email address in `pyproject.toml`.

### Changed

- Improved CI: split lint/typecheck/test jobs, added coverage upload.
- `FillXLConfig` engine validator now accepts `"win32com"`.
- `_normalize_value` now handles `pd.NA/NaT`, `time`, and tz-aware datetimes.
- Changelog: switched to Keep a Changelog 1.1.0 format.
- CI test runner now excludes `win32com`-marked tests on Linux (where Excel COM is unavailable).
- `.gitignore` now covers `repo_dump.md`, `make_md.py`, and `scripts/__pycache__/`.
- CI now also excludes `win32com`-marked tests alongside `xlwings` tests.

### Fixed

- `extend_conditional_formatting` flag now properly passed to `populate_table`.
- Removed dead `NotImplementedError` in `_backend.py`.
- Removed failing doctest example from `_core.fill()`.
- Chart data range update now refreshes `ptCount` caches.
- `_fix_chart_axes()` restores axis elements dropped by openpyxl on load.
- **win32com**: Formula column detection now happens *before* column validation,
  so that `all_but_formulas` strategy correctly exempts calculated columns.
  Previously the backend passed an empty `formula_columns` list, causing a
  spurious `ValueError` for tables with calculated columns.
- **win32com**: Calculated column formulas now keep their original newlines and
  indentation.  Previously LF/CR were stripped before writing via COM, which
  flattened carefully formatted multi-line `LET` / `LAMBDA` formulas.
- CI / packaging: `pywin32` in the `win32com` extra is now marked
  `sys_platform == 'win32'`, so `uv sync --all-extras` succeeds on Linux.
- Migrated deprecated `tool.uv.dev-dependencies` to `dependency-groups.dev`.

### Removed

- `repo_dump.md` and `make_md.py` - development artifacts that were
  accidentally committed.
- Unused `xlsxwriter` runtime dependency (the library uses `openpyxl` and
  `win32com` exclusively).
- `dist/` build artifacts removed from the working tree.

## [0.1.0] - 2026-07-01

Initial release, focused on turning an internal tool into a proper package with a test suite.

### Added

- Initial beta release.
- Fill named cells and Excel tables via `fill()`.
- Support for pandas, polars, duckdb, list[dict], dict[str, list] data formats.
- openpyxl backend (cross-platform, no Excel required).
- win32com backend (Windows + Excel required).
- Column validation strategies (strict, drop, warn, all_but_formulas, none).
- Conditional formatting extension on table grow/shrink.
- Chart data range updates on table resize.
- Calculated column formula propagation.
- Hide/delete sheets, set active sheet, pivot table refresh.
- Style copying for new table rows.

[0.4.0]: https://github.com/hirolau/fillxl/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/hirolau/fillxl/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/hirolau/fillxl/compare/v0.2.0-beta.1...v0.2.0
[0.2.0-beta.1]: https://github.com/hirolau/fillxl/releases/tag/v0.2.0-beta.1
[0.1.0]: https://github.com/hirolau/fillxl/releases/tag/v0.1.0
