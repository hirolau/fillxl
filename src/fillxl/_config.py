"""Configuration for fillxl."""

from __future__ import annotations

from typing import get_args

from attrs import define, field
from attrs.validators import instance_of

from fillxl._validation import ColumnValidationStrategy

_KNOWN_STRATEGIES: frozenset[str] = frozenset(get_args(ColumnValidationStrategy))


def _validate_column_strategy(
    _instance: object,
    _attribute: object,
    value: str,
) -> None:
    """Validate the column validation strategy name.

    Args:
        _instance: attrs instance (unused, required by attrs validator signature)
        _attribute: attrs attribute (unused)
        value: Strategy name to validate

    Raises:
        ValueError: If *value* is not a known strategy name

    Examples:
        >>> _validate_column_strategy(None, None, "strict")
        >>> _validate_column_strategy(None, None, "all_but_formulas")
        >>> _validate_column_strategy(None, None, "bogus")
        Traceback (most recent call last):
            ...
        ValueError: column_validation_strategy must be one of ...got 'bogus'

    """
    if value not in _KNOWN_STRATEGIES:
        msg = (
            f"column_validation_strategy must be one of {sorted(_KNOWN_STRATEGIES)}, got {value!r}"
        )
        raise ValueError(msg)


def _to_sheet_tuple(value: list[str] | tuple[str, ...] | None) -> tuple[str, ...] | None:
    """Freeze a sheet-name list into a tuple.

    ``FillXLConfig`` is frozen, so a mutable list attribute would let callers
    mutate configuration after construction.

    Examples:
        >>> _to_sheet_tuple(["a", "b"])
        ('a', 'b')
        >>> _to_sheet_tuple(None) is None
        True

    """
    if value is None:
        return None
    return tuple(value)


@define(frozen=True, kw_only=True)
class FillXLConfig:
    """Configuration for the fillxl fill operation.

    Attributes:
        column_validation_strategy: How to handle column mismatches between
            data and Excel tables.
        escape_formulas: Whether to write strings starting with ``=`` as
            literal text instead of formulas.  Defaults to ``True`` so that
            data from untrusted sources cannot inject formulas; set it to
            ``False`` when the input is trusted and is meant to contain
            formulas.
        refresh_pivot_tables: Whether to mark pivot caches for refresh on open.
        extend_conditional_formatting: Whether to extend conditional
            formatting rules when a table is resized.
        copy_table_styles: Whether to copy existing table cell styles to new
            rows.
        hide_underscore_sheets: Whether to hide sheets starting with "_".
        delete_sheets: Sheet names to delete.
        active_sheet: Name of the sheet to set as active after filling.

    """

    column_validation_strategy: ColumnValidationStrategy = field(
        default="all_but_formulas",
        validator=_validate_column_strategy,
    )
    escape_formulas: bool = field(default=True, validator=instance_of(bool))
    refresh_pivot_tables: bool = field(default=True, validator=instance_of(bool))
    extend_conditional_formatting: bool = field(default=True, validator=instance_of(bool))
    copy_table_styles: bool = field(default=True, validator=instance_of(bool))
    hide_underscore_sheets: bool = field(default=True, validator=instance_of(bool))
    delete_sheets: tuple[str, ...] | None = field(default=None, converter=_to_sheet_tuple)
    active_sheet: str | None = field(default=None)
