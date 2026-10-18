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

    Raises:
        ValueError: If *strategy* is ``"strict"``, ``"all_but_formulas"``, or
            ``"all_but_formulas_strict"`` and the validation fails.

    """
    # "none" skips all checks
    if strategy == "none":
        return

    formula_set = set(formula_columns or [])

    # For "all_but_formulas" modes, exclude formula columns from the table
    # columns we compare against.
    if strategy in ("all_but_formulas", "all_but_formulas_strict"):
        effective_table_columns = [c for c in table_columns if c not in formula_set]
    else:
        effective_table_columns = table_columns

    table_set = set(effective_table_columns)
    data_set = set(data_columns)

    missing = table_set - data_set  # columns in table but not in data
    extra = data_set - table_set  # columns in data but not in table

    if not missing and not extra:
        return

    # ── Strict modes ───────────────────────────────────────────────────────

    if strategy == "strict":
        msg_parts = []
        if missing:
            msg_parts.append(f"Missing columns: {sorted(missing)}")
        if extra:
            msg_parts.append(f"Extra columns: {sorted(extra)}")
        msg = (
            f"Column mismatch for table. "
            f"{'; '.join(msg_parts)}. "
            f"Table columns: {table_columns}. "
            f"Data columns: {data_columns}."
        )
        raise ValueError(msg)

    if strategy == "all_but_formulas_strict":
        msg_parts = []
        if missing:
            msg_parts.append(f"Missing columns: {sorted(missing)}")
        if extra:
            msg_parts.append(f"Extra columns: {sorted(extra)}")
        if msg_parts:
            msg = (
                f"Column mismatch for table (all_but_formulas_strict). "
                f"{'; '.join(msg_parts)}. "
                f"Table columns (excl. formulas): {effective_table_columns}. "
                f"Data columns: {data_columns}."
            )
            raise ValueError(msg)
        return

    if strategy == "all_but_formulas":
        if missing:
            msg = (
                f"Missing columns in data (all_but_formulas). "
                f"These non-formula table columns have no matching data column: "
                f"{sorted(missing)}. "
                f"Table columns (excl. formulas): {effective_table_columns}. "
                f"Data columns: {data_columns}."
            )
            raise ValueError(msg)
        return

    # ── Non-strict modes ───────────────────────────────────────────────────

    if strategy == "drop":
        if extra:
            logger.info("Dropping extra data columns: %s", sorted(extra))
        if missing:
            logger.info("Skipping missing table columns: %s", sorted(missing))
        return

    if missing:
        logger.warning("Data missing columns expected by table: %s", sorted(missing))
    if extra:
        logger.warning("Data has extra columns not in table: %s", sorted(extra))
