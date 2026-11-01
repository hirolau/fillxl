"""Calculated-column detection shared by both backends.

A table column is a calculated column when its definition carries a
``calculatedColumnFormula`` and the template body agrees (it holds formulas or
is empty).  Both backends apply the same rule, so they keep the same columns.
"""

from __future__ import annotations

import logging
import math
import sys
from typing import TYPE_CHECKING, Any, Literal, NamedTuple

from openpyxl.worksheet.cell_range import CellRange

from fillxl._utils import table_column_name
from fillxl._warnings import warn_user

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping, Sequence

    from openpyxl.worksheet.table import Table
    from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

BodyState = Literal["formula", "mixed", "empty", "values"]
CellKind = Literal["formula", "empty", "value"]


class ColumnFormulaEvidence(NamedTuple):
    """What a backend observed about one table column before writing to it.

    Attributes:
        name: Column name.
        definition: ``calculatedColumnFormula`` text without ``=``, or ``None``.
        definitions_known: ``False`` when the file format hides definitions
            (``.xlsb``/``.xls`` read through COM).
        body_state: What the template body cells of this column hold.
        body_formula: Formula of the first body cell including ``=`` (COM only;
            openpyxl passes ``None`` and writes the definition instead).

    """

    name: str
    definition: str | None
    definitions_known: bool
    body_state: BodyState
    body_formula: str | None = None


def formula_definitions(table: Table) -> dict[str, str]:
    """Map each column that has a ``calculatedColumnFormula`` to its text (no ``=``).

    Constants such as ``100`` are kept: ``=100`` is a valid calculated column.
    Whether a definition is still live is decided by :func:`resolve_formula`.
    """
    definitions: dict[str, str] = {}
    for column in table.tableColumns:
        formula = column.calculatedColumnFormula
        text = getattr(formula, "attr_text", None) if formula is not None else None
        if text:
            definitions[table_column_name(column)] = str(text)
    return definitions


def classify_body(kinds: Iterable[CellKind]) -> BodyState:
    """Summarise the kinds of the body cells of one column.

    Examples:
        >>> classify_body(["formula", "formula"])
        'formula'
        >>> classify_body(["formula", "empty"])
        'formula'
        >>> classify_body(["formula", "value"])
        'mixed'
        >>> classify_body(["empty", "empty"])
        'empty'
        >>> classify_body(["value", "empty"])
        'values'

    """
    seen = set(kinds)
    if "formula" in seen:
        return "mixed" if "value" in seen else "formula"
    return "values" if "value" in seen else "empty"


def resolve_formula(evidence: ColumnFormulaEvidence) -> str | None:
    """Return the formula a column must keep (with ``=``), or ``None`` for a data column.

    Examples:
        >>> E = ColumnFormulaEvidence
        >>> resolve_formula(E("C", "[@A]*2", True, "empty"))
        '=[@A]*2'
        >>> resolve_formula(E("C", "[@A]*2", True, "formula", "=[@A]*2"))
        '=[@A]*2'
        >>> resolve_formula(E("C", "100", True, "formula"))
        '=100'
        >>> resolve_formula(E("C", "100", True, "values")) is None
        True
        >>> resolve_formula(E("C", None, True, "formula")) is None
        True
        >>> resolve_formula(E("C", None, False, "formula", "=A2*2"))
        '=A2*2'
        >>> resolve_formula(E("C", None, False, "mixed", "=A2*2")) is None
        True

    """
    if not evidence.definitions_known:
        # Without definitions only a column that Excel filled completely is trusted.
        return evidence.body_formula if evidence.body_state == "formula" else None
    if evidence.definition is None:
        return None
    if evidence.body_state == "values":
        logger.debug(
            "Column '%s' has a calculated-column definition but holds values; treating it as data",
            evidence.name,
        )
        return None
    return evidence.body_formula or "=" + evidence.definition


def openpyxl_formula_columns(
    worksheet: Worksheet,
    table: Table,
    first_body_row: int,
    last_body_row: int,
) -> dict[str, str]:
    """Calculated columns of an openpyxl table, mapped to the formula to write.

    Must run before any body cell is cleared or written.
    """
    definitions = formula_definitions(table)
    first_col = CellRange(table.ref).min_col
    formulas: dict[str, str] = {}
    for offset, column in enumerate(table.tableColumns):
        name = table_column_name(column)
        definition = definitions.get(name)
        if definition is None:
            continue
        kinds = [
            _cell_kind(worksheet._cells.get((row, first_col + offset)))
            for row in range(first_body_row, last_body_row + 1)
        ]
        evidence = ColumnFormulaEvidence(name, definition, True, classify_body(kinds))
        formula = resolve_formula(evidence)
        if formula is not None:
            formulas[name] = formula
    return formulas


def _cell_kind(cell: Any) -> CellKind:  # noqa: ANN401
    if cell is None or cell.value is None:
        return "empty"
    return "formula" if cell.data_type == "f" else "value"


def _is_blank_cell_value(value: Any) -> bool:  # noqa: ANN401
    """Whether *value* should be treated as an empty cell without normalising it."""
    if value is None:
        return True
    if isinstance(value, str):
        return value == ""
    pd = sys.modules.get("pandas")
    if pd is not None and (value is pd.NA or value is pd.NaT):
        return True
    return bool(isinstance(value, float) and math.isnan(value))


def warn_about_ignored_formula_data(
    table_name: str,
    data_columns: Sequence[str],
    rows: Sequence[Sequence[Any]],
    formulas: Mapping[str, str],
) -> list[str]:
    """Warn when the data carries values for calculated columns, which are not written.

    Columns supplied only as blanks (``None``, NaN, ``""``) are ignored silently.
    """
    ignored = [
        column
        for index, column in enumerate(data_columns)
        if column in formulas and any(not _is_blank_cell_value(row[index]) for row in rows)
    ]
    if ignored:
        warn_user(
            f"Table {table_name!r}: calculated column(s) {ignored} keep their formulas; "
            "the data supplied for them was ignored.",
        )
    return ignored
