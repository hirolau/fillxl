"""Column validation logic for fillxl.

Provides strategies for handling column mismatches between data and Excel tables.
"""

from __future__ import annotations

import logging
from typing import Literal

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
            - ``"drop"``: Silently drop extra data columns and skip missing
              table columns.
            - ``"warn"``: Log warnings but proceed with best-effort fill.
        formula_columns: Names of table columns that are formulas.  Only
            relevant for ``"all_but_formulas"`` / ``"all_but_formulas_strict"``.
        table_name: Name of the table, used in error and log messages.

    Raises:
        ValueError: If *strategy* is ``"strict"``, ``"all_but_formulas"``, or
            ``"all_but_formulas_strict"`` and the validation fails.

    """
    if strategy == "none" or not data_columns:
        return

    where = f"table {table_name!r}" if table_name else "table"
    label = f"Table {table_name!r}" if table_name else "Table"
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
    if strategy == "drop":
        if extra:
            logger.info("%s: dropping extra data columns %s", label, extra)
        if missing:
            logger.info("%s: leaving table columns without data empty: %s", label, missing)
        return
    if missing:
        logger.warning("%s: data is missing columns %s", label, missing)
    if extra:
        logger.warning("%s: data has extra columns %s", label, extra)
