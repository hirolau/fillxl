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


class TestFrozenConfigAttributes:
    """FillXLConfig is frozen, so its attributes must be immutable too."""

    def test_delete_sheets_is_stored_as_a_tuple(self) -> None:
        """A caller's list cannot be mutated after construction."""
        cfg = FillXLConfig(delete_sheets=["Raw", "Scratch"])
        assert cfg.delete_sheets == ("Raw", "Scratch")
        assert isinstance(cfg.delete_sheets, tuple)

    def test_delete_sheets_accepts_a_tuple_and_none(self) -> None:
        assert FillXLConfig(delete_sheets=("Raw",)).delete_sheets == ("Raw",)
        assert FillXLConfig().delete_sheets is None

    def test_formula_escaping_is_on_by_default(self) -> None:
        """Strings starting with "=" stay literal unless the caller opts out."""
        assert FillXLConfig().escape_formulas is True
