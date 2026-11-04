"""Column validation logic for fillxl.

Provides strategies for handling column mismatches between data and Excel tables.
"""

from __future__ import annotations

import logging
from typing import Literal

from fillxl._warnings import warn_user

logger = logging.getLogger(__name__)

ColumnValidationStrategy = Literal[
    "all_but_formulas",
    "all_but_formulas_strict",
    "none",
    "strict",
    "drop",
    "warn",
]


def validate_columns(
    table_columns: list[str],
    data_columns: list[str],
    strategy: ColumnValidationStrategy = "all_but_formulas",
    *,
    formula_columns: list[str] | None = None,
    table_name: str | None = None,
    data_rows: int = 0,
) -> None:
    """Validate data columns against Excel table columns.

    Compares the columns present in the incoming data against the columns
    expected by the Excel table, and decides how to handle mismatches.
    Backends match columns by *name*, so the data columns never need to be
    filtered here.

    Args:
        table_columns: Column names from the Excel table definition.
        data_columns: Column names from the incoming data.
        strategy: How to handle mismatches:
            - ``"all_but_formulas"`` (default): All non-formula table columns
              must be present in the data. Extra data columns are allowed.
              Formula columns are ignored.
            - ``"all_but_formulas_strict"``: Like ``"all_but_formulas"``, but
              NO extra data columns are allowed.
            - ``"none"``: Skip validation entirely.
            - ``"strict"``: 1-to-1 match between table and data columns.
            - ``"drop"``: Deprecated alias of ``"none"``; it is accepted for
              one release and ``FillXLConfig`` warns with a
              ``DeprecationWarning``.
            - ``"warn"``: Emit one :class:`~fillxl.FillXLWarning` naming the
              missing and extra columns, then proceed with a best-effort fill.
        formula_columns: Names of table columns that are formulas.  Only
            relevant for ``"all_but_formulas"`` / ``"all_but_formulas_strict"``.
        table_name: Name of the table, used in error and log messages.
        data_rows: Number of rows the data holds.  Data with rows but no
            columns still has to pass validation: the table is resized and
            every data column is cleared, which ``"strict"`` must refuse.
            Empty data (no rows, no columns) keeps the old behaviour of
            clearing the table.

    Raises:
        ValueError: If *strategy* is ``"strict"``, ``"all_but_formulas"``, or
            ``"all_but_formulas_strict"`` and the validation fails.

    Examples:
        >>> validate_columns(["A"], [], "all_but_formulas")  # a cleared table
        >>> try:
        ...     validate_columns(["A"], [], "strict", data_rows=3)
        ... except ValueError as error:
        ...     str(error).startswith("Column mismatch for table: missing columns")
        True

    """
    if strategy in ("none", "drop") or (not data_columns and not data_rows):
        return

    where = f"table {table_name!r}" if table_name else "table"
    formula_set = set(formula_columns or [])
    if strategy in ("all_but_formulas", "all_but_formulas_strict"):
        effective = [c for c in table_columns if c not in formula_set]
    else:
        effective = list(table_columns)

    missing = sorted(set(effective) - set(data_columns))
    extra = sorted(set(data_columns) - set(effective))
    if not missing and not extra:
        return

    details = "; ".join(
        part
        for part in (
            f"missing columns: {missing}" if missing else "",
            f"extra columns: {extra}" if extra else "",
        )
        if part
    )

    if strategy == "strict":
        msg = (
            f"Column mismatch for {where}: {details}. "
            f"Table columns: {table_columns}. Data columns: {data_columns}."
        )
        raise ValueError(msg)
    if strategy == "all_but_formulas_strict":
        msg = (
            f"Column mismatch for {where} (all_but_formulas_strict): {details}. "
            f"Table columns (excluding calculated columns): {effective}. "
            f"Data columns: {data_columns}."
        )
        raise ValueError(msg)
    if strategy == "all_but_formulas":
        if missing:
            msg = (
                f"Column mismatch for {where} (all_but_formulas): these table columns have "
                f"no data column: {missing}. Table columns (excluding calculated columns): "
                f"{effective}. Data columns: {data_columns}."
            )
            raise ValueError(msg)
        return
    if strategy == "warn":
        warn_user(
            f"Column mismatch for {where}: {details}. The fill continues with the "
            "columns that match; rename the columns or choose another "
            "column_validation_strategy to change that.",
        )
