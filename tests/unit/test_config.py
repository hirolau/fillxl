"""Config validation tests for FillXLConfig."""

from __future__ import annotations

import pytest

from fillxl import FillXLConfig


class TestConfigValidation:
    """Config attribute validators."""

    def test_config_rejects_invalid_column_validation_strategy(self) -> None:
        """An invalid strategy name should raise ValueError."""
        with pytest.raises(ValueError, match="column_validation_strategy"):
            FillXLConfig(column_validation_strategy="bogus")  # ty:ignore[invalid-argument-type]
