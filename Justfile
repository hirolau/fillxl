# fillxl — common development tasks
#
# Use `just <recipe>` to run a task.
# `just` alone lists available recipes.

set shell := ["powershell.exe", "-c"]

# Run CI test subset (no Excel required)
test:
    uv run pytest tests/ -m "not xlwings and not win32com" --cov=fillxl --cov-report=term-missing

# Run win32com-only tests (requires Excel + pywin32)
test-win32com:
    uv run pytest tests/ -m win32com -v

# Run full test suite (Excel required — author's machine only)
test-all:
    uv run pytest tests/ --cov=fillxl --cov-report=term-missing

# Run story tests (narrative chapter progression)
test-story:
    uv run pytest tests/story/ -v

# Lint and format with ruff
lint:
    uv run ruff check --fix src/ tests/
    uv run ruff format src/ tests/

# Type-check with mypy (strict mode)
typecheck:
    uv run mypy src/fillxl/

# Run doctests for pure-logic modules
doctest:
    uv run pytest --doctest-modules src/fillxl/ -v

# Fill the 3 test-template .xlsx files with demo data (both backends)
fill-samples:
    uv run python scripts/fill_demo.py --engine both

# Fill samples using only the openpyxl backend
fill-samples-openpyxl:
    uv run python scripts/fill_demo.py --engine openpyxl

# Fill samples using only the win32com backend (requires Excel)
fill-samples-win32com:
    uv run python scripts/fill_demo.py --engine win32com

# Generate golden copies in tests/assets/golden/ (both backends)
fill-golden: fill-golden-openpyxl fill-golden-win32com

# Generate golden copies using the openpyxl backend (CI-safe)
fill-golden-openpyxl:
    uv run python scripts/fill_demo.py --engine openpyxl --output-dir tests/assets/golden

# Generate golden copies using the win32com backend (requires Excel)
fill-golden-win32com:
    uv run python scripts/fill_demo.py --engine win32com --output-dir tests/assets/golden

# CI gate: lint → typecheck → test (no Excel)
ci: lint typecheck test

# Build wheel + sdist
build:
    uvx hatchling build

# Verify built distributions (run after `just build`)
publish-check:
    uvx twine check --strict dist/*
    uvx check-wheel-contents dist/*.whl

# Release gate: lint → typecheck → test-all (Excel required)
release: lint typecheck test-all
