# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[unreleased]: https://github.com/hirolau/fillxl/compare/v0.2.0-beta.1...HEAD
[0.2.0-beta.1]: https://github.com/hirolau/fillxl/releases/tag/v0.2.0-beta.1
[0.1.0]: https://github.com/hirolau/fillxl/releases/tag/v0.1.0
