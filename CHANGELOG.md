# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `FillXLConfig.escape_formulas` (default `False`): write strings that start
  with `=` as literal text instead of formulas. Formula-like strings from
  untrusted data are a formula-injection vector, so this is now opt-in-able.
- Sheet-scoped defined names are now resolved by the openpyxl backend, matching
  win32com, which finds them through `Workbook.Names`.
- Python 3.15 support: added to the classifiers, the CI test matrix, and the
  nox `PYTHONS` list (the `requires-python` floor stays at 3.11).
- polars 2 support: the dev/test environment and `uv.lock` use polars 2, and
  the nox matrix has an explicit `tests_polars-2` session. The supported
  release lines are now documented as pandas 2–3, polars 0.20–2, and duckdb
  1.x.
- The `duckdb` extra is limited to Python ≤3.14: duckdb currently publishes no
  cp315 wheels. Its tests skip on 3.15 until upstream provides them.

### Changed

- **Both backends** now write every non-empty string as text, so Excel's
  typed-input parser cannot reinterpret the data. win32com does this with
  Excel's text marker (apostrophe prefix, no number-format changes), so
  `"00123"` stays `"00123"` instead of becoming `123` and `"1/2"` no longer
  becomes a date. openpyxl stores strings as text as well, including
  error-like strings such as `"#N/A"` that it would otherwise turn into Excel
  errors.
- Empty strings (`""`) become empty cells in both engines. openpyxl cannot
  round-trip an empty text cell (it reads back as `None`), so writing one
  would make the two backends disagree.
- **Both backends**: table cell styles are applied *before* the value, so
  datetime columns get a date format in the openpyxl output as well. Datetime
  table cells now read back as datetimes from both engines instead of a bare
  serial number from openpyxl.
- **win32com**: pivot caches are refreshed *and* marked `refreshOnLoad`, so
  both engines leave a workbook that recalculates its pivots on open.
- The tabular adapter accepts any sequence-like column values for
  `dict[str, ...]` data, including NumPy arrays, tuples and pandas Series.
  Columns are read positionally, so a Series with a non-default index (and
  non-string mapping keys) works, and ragged columns raise instead of being
  silently truncated.
- `delete_sheets` is stored as a tuple, so a frozen `FillXLConfig` cannot be
  mutated through a caller's list.

### Fixed

- NumPy scalars (`np.int64`, `np.int32`, `np.float32`, `np.bool_`, …) are
  converted to native Python values instead of `str(value)`, which used to
  produce the "number stored as text" marker in Excel. `np.datetime64` and
  `np.timedelta64` — including nanosecond precision, where `.item()` returns
  an integer — are converted to `datetime`/`timedelta` as well.
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
- Both backends reject tables that have a totals row with a clear error
  instead of overwriting the totals and miscounting rows.
- The timezone warning flag is a `ContextVar`, so two concurrent `fill()` calls
  in different threads no longer share it.
- `__version__` is read from package metadata instead of being hard-coded, so
  it cannot drift from `pyproject.toml`.

### Removed

- `FillXLConfig.extend_tables` — it was never read by any backend, so it
  silently did nothing. Tables are always resized to fit the data.

## [0.2.0] — 2026-07-19

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

## [0.2.0-beta.1] — 2026-07-17

### Added

- Cross-backend datatype tests for openpyxl backend (CI-safe).
- GitHub Actions CI workflow (lint, typecheck, test matrix on 3.11–3.14).
- GitHub Actions release workflow (publishes to PyPI on version tags).
- `CHANGELOG.md`, issue templates, PR template.
- Test suite split into logical modules (backends, validation, charts, CF, …).
- All test methods renamed to read as descriptive sentences.
- Cross-backend dtype testing (openpyxl + win32com, 68 tests).
- Golden copy regression tests (openpyxl + win32com).
- Story tests: charts, data types, template preservation, failure recovery, report generation.
- Backend protocol tests and openpyxl unit tests.
- win32com performance optimisation (26× speedup for large tables).
- Coverage: overall 85% (excl. win32com: ~95%).
- `ruff` clean, `mypy` strict — 0 errors.
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

- `repo_dump.md` and `make_md.py` — development artifacts that were
  accidentally committed.
- Unused `xlsxwriter` runtime dependency (the library uses `openpyxl` and
  `win32com` exclusively).
- `dist/` build artifacts removed from the working tree.

## [0.1.0] — 2026-07-01

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

[unreleased]: https://github.com/hirolau/fillxl/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/hirolau/fillxl/compare/v0.2.0-beta.1...v0.2.0
[0.2.0-beta.1]: https://github.com/hirolau/fillxl/releases/tag/v0.2.0-beta.1
[0.1.0]: https://github.com/hirolau/fillxl/releases/tag/v0.1.0
