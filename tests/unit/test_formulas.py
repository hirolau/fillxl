"""Unit tests for the shared calculated-column detection."""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableColumn, TableFormula

from fillxl import FillXLWarning
from fillxl._formulas import (
    extend_table_ranges,
    formula_definitions,
    openpyxl_formula_columns,
    shift_relative_rows,
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


# ═══════════════════════════════════════════════════════════════════════════════
# Filling a calculated column: Excel copies the definition down
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize(
    ("formula", "offset", "expected"),
    [
        # Relative rows move with the copy, like Excel's fill handle.
        ("=B5*C5", 1, "=B6*C6"),
        ("=B5*C5", 2, "=B7*C7"),
        # Absolute and mixed rows stay put.
        ("=B5*$C$5", 2, "=B7*$C$5"),
        ("=B5*C$5", 2, "=B7*C$5"),
        ("=$B5*$C5", 1, "=$B6*$C6"),
        # Structured references are already row-relative.
        ("=[@Qty]*[@Price]", 3, "=[@Qty]*[@Price]"),
        # A string that looks like a reference is text, not a reference.
        ('=IF(A5>0,"B5",B5)', 1, '=IF(A6>0,"B5",B6)'),
        # A sheet-qualified relative row moves too.
        ("=Data!B5*2", 1, "=Data!B6*2"),
        ("='My Sheet'!B5", 1, "='My Sheet'!B6"),
        # A function name that looks like a cell does not move.
        ("=LOG10(B5)", 1, "=LOG10(B6)"),
        # Constants and definitions without references are untouched.
        ("=100", 3, "=100"),
    ],
)
def test_shift_relative_rows_moves_only_relative_a1_rows(
    formula: str,
    offset: int,
    expected: str,
) -> None:
    """The row arithmetic Excel applies when it fills a calculated column."""
    assert shift_relative_rows(formula, offset) == expected


def test_shift_relative_rows_leaves_a_formula_without_references_alone() -> None:
    """No change means the very same string object, not a rebuilt copy."""
    formula = "=[@Qty]*100"

    assert shift_relative_rows(formula, 4) is formula


# ═══════════════════════════════════════════════════════════════════════════════
# Resizing formulas that track the table body, as Excel does
# ═══════════════════════════════════════════════════════════════════════════════

_RESIZE = {
    "sheet_title": "Data",
    "first_col": 2,
    "last_col": 2,
    "old_body": (4, 4),
    "new_body": (4, 7),
}


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=SUM(B4:B4)", "=SUM(B4:B7)"),
        ("=SUM($B$4:$B$4)", "=SUM($B$4:$B$7)"),
        ("=SUM(B$4:B$4)", "=SUM(B$4:B$7)"),
        ("=SUM(Data!B4:B4)", "=SUM(Data!B4:B7)"),
        ('=IF(A1>0,"B4:B4",SUM(B4:B4))', '=IF(A1>0,"B4:B4",SUM(B4:B7))'),
        ("=AVERAGE(B4:B4)+MAX(C1:C2,B4:B4)", "=AVERAGE(B4:B7)+MAX(C1:C2,B4:B7)"),
    ],
)
def test_extend_table_ranges_resizes_the_ranges_that_track_the_body(
    formula: str,
    expected: str,
) -> None:
    """A range over exactly the body rows follows the table, like Excel's resize."""
    assert extend_table_ranges(formula, **_RESIZE) == expected


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(B4:B5)",  # one row more than the body
        "=SUM(C1:C3)",  # outside the table's columns
        "=SUM(D4:D4)",  # a column the table does not own
        "=SUM(Other!B4:B4)",  # another sheet
        "=SUM(Z1:Z1)",  # unrelated rows
        "=B4*2",  # a single cell is not a range
        "=SUM(B$1:$B$4)",  # starts outside the body
        "=50",  # no references at all
    ],
)
def test_extend_table_ranges_leaves_unrelated_formulas_alone(formula: str) -> None:
    """Only a range that covers the table body is resized; everything else is kept."""
    assert extend_table_ranges(formula, **_RESIZE) is None


def test_extend_table_ranges_shrinks_when_the_table_shrinks() -> None:
    """The rule is symmetric: Excel takes the range back down again."""
    assert (
        extend_table_ranges(
            "=SUM(B4:B7)",
            sheet_title="Data",
            first_col=2,
            last_col=2,
            old_body=(4, 7),
            new_body=(4, 4),
        )
        == "=SUM(B4:B4)"
    )
