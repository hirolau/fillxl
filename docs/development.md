# Development

## Setup

```bash
git clone https://github.com/hirolau/fillxl.git
cd fillxl
uv sync --all-extras
```

## Running Tests

```bash
# All tests except Excel-COM-dependent ones
uv run pytest -m "not win32com"

# All tests (Windows with Excel only)
uv run pytest
```

Behaviour tests are parametrised over all three engines (`tests/_engines.py`).
The `win32com` case skips itself when Excel is not installed; `openpyxl` and
`fillxl` run everywhere. Tests that genuinely need Excel — opening a file,
repairing, recalculating — carry the `excel` marker and are skipped elsewhere:

```bash
pytest -m excel        # Windows with Microsoft Excel
pytest -m "not excel"  # anywhere
```

When a documented behaviour and Excel disagree, the test wins and the
documentation is wrong. Add a test that reproduces the difference before
changing either engine.

## Code Quality

```bash
uv run ruff check         # Lint
uv run mypy               # Type check (strict mode)
uv run pyrefly check      # Additional type checker
```

## Building Docs

```bash
uv run mkdocs serve        # Live preview at http://localhost:8000
uv run mkdocs build        # Build static site to site/
```

## Supported Environments

| | 3.11 | 3.12 | 3.13 | 3.14 | 3.15 |
|---|:---:|:---:|:---:|:---:|:---:|
| Windows | :material-check: | :material-check: | :material-check: | :material-check: | :material-check: |
| Linux | :material-check: | :material-check: | :material-check: | :material-check: | :material-check: |
| macOS | :material-check: | :material-check: | :material-check: | :material-check: | :material-check: |

!!! note "macOS"
    macOS is best-effort — CI does not currently test on macOS runners.

## Supported Data Backends

The nox version matrix tests the minimum supported release, the version in
`uv.lock`, and explicit major versions of the optional data libraries:

| Backend | Minimum | Also tested | Latest |
|---------|---------|-------------|--------|
| pandas | 2.0 | 3.x | 3.x |
| polars | 0.20 | 2.x | 2.x |
| duckdb | 1.0 | 2.x | 2.x |

Each session (`nox -s tests_polars-2`, `nox -s tests_duckdb-min`, …) pins one
release line and runs every test that exercises that backend. Selection is by
marker — `-k "polars or validation"` matches marker names as well as test
names — so a test that calls `pytest.importorskip("polars")` must also carry
`@pytest.mark.polars` (likewise `duckdb`). Missing that marker silently drops
the test out of every pinned-version run, so
`tests/unit/test_backend_markers.py` fails when one is forgotten.

The `backends` job in CI runs the same selection for the oldest and newest
supported release line of each backend on Python 3.12.

!!! note "duckdb on Python 3.15"
    duckdb 2.x ships abi3 wheels that install on 3.15, so `uv.lock` resolves
    duckdb 2.x everywhere and the `duckdb` extra works on every supported
    Python. duckdb 1.x publishes no cp315 wheels, so `tests_duckdb-min` skips
    3.15 — that is a limit of duckdb 1.x, not of fillxl.

## Project Structure

```
fillxl/
├── src/fillxl/          # Package source
│   ├── _core.py         # Main fill() entry point
│   ├── _backends/       # openpyxl, fillxl and win32com backends
│   ├── _cells.py        # Named cell handling
│   ├── _tables.py       # Excel table population
│   ├── _charts.py       # Chart range updates
│   └── ...
├── tests/               # Test suite
├── docs/                # Documentation (this site)
└── mkdocs.yml           # MkDocs configuration
```

## License

[MIT](https://github.com/hirolau/fillxl/blob/main/LICENSE)
