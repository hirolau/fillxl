"""Unit tests for the shared calculated-column detection."""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableColumn, TableFormula

from fillxl import FillXLWarning
from fillxl._formulas import (
    formula_definitions,
    openpyxl_formula_columns,
    warn_about_ignored_formula_data,
)

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet


def _table_with_definition(
    definition: str | None,
    body_value: object = None,
) -> tuple[Worksheet, Table]:
    """A one-row, two-column table whose ``C`` column has *definition*."""
    wb = Workbook()
    ws = wb.active
    ws.append(["A", "C"])
    ws.append([1, body_value])
    table = Table(
        displayName="T",
        ref="A1:B2",
        tableColumns=[
            TableColumn(id=1, name="A"),
            TableColumn(
                id=2,
                name="C",
                calculatedColumnFormula=(
                    TableFormula(attr_text=definition) if definition is not None else None
                ),
            ),
        ],
    )
    ws.add_table(table)
    return ws, table


def test_formula_definitions_keeps_constant_formulas() -> None:
    """``=100`` is a valid calculated column, not a plain number."""
    _, table = _table_with_definition("100")

    assert formula_definitions(table) == {"C": "100"}


def test_formula_definitions_skips_columns_without_a_formula() -> None:
    """Columns without a definition are data columns."""
    _, table = _table_with_definition(None)

    assert formula_definitions(table) == {}


def test_openpyxl_formula_columns_uses_the_definition_when_the_body_is_empty() -> None:
    """An empty body accepts the definition as the column's formula."""
    ws, table = _table_with_definition("[@A]*2")

    assert openpyxl_formula_columns(ws, table, 2, 2) == {"C": "=[@A]*2"}


def test_openpyxl_formula_columns_treats_a_definition_over_values_as_data() -> None:
    """A stale definition over constant data must not hide the data."""
    ws, table = _table_with_definition("100", body_value=5)

    assert openpyxl_formula_columns(ws, table, 2, 2) == {}


def test_openpyxl_formula_columns_keeps_a_column_whose_body_holds_formulas() -> None:
    """A body that already holds formulas proves the column is calculated."""
    ws, table = _table_with_definition("[@A]*2", body_value="=A2*2")

    assert openpyxl_formula_columns(ws, table, 2, 2) == {"C": "=[@A]*2"}


def test_warn_about_ignored_formula_data_is_silent_for_blank_values(
    recwarn: pytest.WarningsRecorder,
) -> None:
    """Blank data for a calculated column is ignored without a warning."""
    ignored = warn_about_ignored_formula_data("T", ["C"], [[None], [float("nan")]], {"C": "=1"})

    assert ignored == []
    assert [w for w in recwarn if issubclass(w.category, FillXLWarning)] == []


def test_warn_about_ignored_formula_data_names_the_columns() -> None:
    """Real data supplied for calculated columns is reported once per table."""
    with pytest.warns(FillXLWarning, match="keep their formulas") as caught:
        ignored = warn_about_ignored_formula_data("T", ["C", "D"], [[1, 2]], {"C": "=1", "D": "=2"})

    assert ignored == ["C", "D"]
    assert len(caught) == 1
    assert "['C', 'D']" in str(caught[0].message)
