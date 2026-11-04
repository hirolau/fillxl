"""Shared fixtures for the fillxl-engine hardening tests."""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from tests._spills import assert_well_formed as assert_package_well_formed
from tests.ooxml.fixtures.builders import BUILDERS

if TYPE_CHECKING:
    from collections.abc import Callable
    from pathlib import Path


@pytest.fixture
def assert_well_formed() -> Callable[[Path], None]:
    """An assertion that every ``.xml`` and ``.rels`` part of a workbook parses."""
    return assert_package_well_formed


@pytest.fixture(scope="session")
def hostile_fixture_factory(
    tmp_path_factory: pytest.TempPathFactory,
) -> Callable[[str], Path]:
    """Build (once per session) the workbook called *name* of :data:`BUILDERS`."""
    root = tmp_path_factory.mktemp("hostile")
    built: dict[str, Path] = {}

    def build(name: str) -> Path:
        if name not in built:
            built[name] = BUILDERS[name](root / f"{name}.xlsx")
        return built[name]

    return build
