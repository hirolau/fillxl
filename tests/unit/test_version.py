"""The package version comes from the installed metadata, not from source code."""

from __future__ import annotations

import tomllib
from importlib.metadata import version
from pathlib import Path

import fillxl

REPO_ROOT = Path(__file__).resolve().parents[2]


def test_the_reported_version_matches_pyproject() -> None:
    """__version__ cannot drift from the version in pyproject.toml."""
    pyproject = tomllib.loads((REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert fillxl.__version__ == pyproject["project"]["version"]
    assert fillxl.__version__ == version("fillxl")


def test_the_version_is_not_hard_coded_in_the_package_init() -> None:
    """The version is read from metadata instead of being written out."""
    source = (REPO_ROOT / "src" / "fillxl" / "__init__.py").read_text(encoding="utf-8")
    pyproject = tomllib.loads((REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    assert "importlib.metadata" in source
    assert f'"{pyproject["project"]["version"]}"' not in source
