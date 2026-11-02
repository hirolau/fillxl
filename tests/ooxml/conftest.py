"""Shared fixtures for the fillxl-engine hardening tests."""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from tests._spills import assert_well_formed as assert_package_well_formed

if TYPE_CHECKING:
    from collections.abc import Callable
    from pathlib import Path


@pytest.fixture
def assert_well_formed() -> Callable[[Path], None]:
    """An assertion that every ``.xml`` and ``.rels`` part of a workbook parses."""
    return assert_package_well_formed
