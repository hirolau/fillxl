"""Config validation tests for FillXLConfig."""

from __future__ import annotations

from typing import Any, cast

import pytest

from fillxl import FillXLConfig


class TestConfigValidation:
    """Config attribute validators."""

    def test_config_rejects_invalid_column_validation_strategy(self) -> None:
        """An invalid strategy name should raise ValueError."""
        with pytest.raises(ValueError, match="column_validation_strategy"):
            FillXLConfig(column_validation_strategy=cast(Any, "bogus"))

    def test_the_accepted_strategy_names_come_from_the_literal_type(self) -> None:
        """The validator and the ColumnValidationStrategy Literal cannot drift."""
        from typing import get_args

        from fillxl._config import _KNOWN_STRATEGIES
        from fillxl._validation import ColumnValidationStrategy

        assert frozenset(get_args(ColumnValidationStrategy)) == _KNOWN_STRATEGIES
        for strategy in get_args(ColumnValidationStrategy):
            FillXLConfig(column_validation_strategy=strategy)
