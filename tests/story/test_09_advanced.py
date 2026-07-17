"""Chapter 9 — Advanced Features.
===============================

Beyond basic table filling, fillxl handles:

* **Calculated columns** — Excel table columns with formulas like ``=[@Col1]&[@Col2]``
  are automatically propagated to new rows.
* **Conditional formatting** — CF rules are extended or shrunk to match the
  new table size.
* **Pivot tables** — Pivot caches can be marked to refresh on open.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
from openpyxl.worksheet.table import (  # ty:ignore[unresolved-import]
    Table,
    TableColumn,
    TableFormula,
)

from fillxl import FillXLConfig, fill

if TYPE_CHECKING:
    from pathlib import Path

# ═══════════════════════════════════════════════════════════════════════════════
# 9.1 Calculated columns
# ═══════════════════════════════════════════════════════════════════════════════


class TestCalculatedColumns:
    """Calculated-column formulas propagate to new rows automatically."""

    def test_calculated_column_formulas_are_written_to_new_rows(
        self,
        output_path: Path,
    ) -> None:
        """A calculated column formula propagates to all data rows."""
        # Build a template with a calculated column
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "First"
        ws["B1"] = "Last"
        ws["C1"] = "FullName"  # calculated: =[@First]&" "&[@Last]
        ws["A2"] = "John"
        ws["B2"] = "Doe"

        col_formula = TableFormula(attr_text='[@First]&" "&[@Last]')
        table = Table(
            displayName="People",
            ref="A1:C2",
            tableColumns=[
                TableColumn(id=1, name="First"),
                TableColumn(id=2, name="Last"),
                TableColumn(id=3, name="FullName", calculatedColumnFormula=col_formula),
            ],
        )
        ws.add_table(table)
        tmpl = output_path.with_name("calc_col_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        # Fill with data — omit the calculated column from data
        data = [
            {"First": "Alice", "Last": "Smith"},
            {"First": "Bob", "Last": "Jones"},
            {"First": "Charlie", "Last": "Brown"},
        ]

        fill(template=tmpl, output=output_path, tables={"People": data})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Data"]
            # Each data row should have a formula in column C
            for row in range(2, 5):
                val = ws2.cell(row=row, column=3).value
                assert val is not None, f"Row {row} has no formula"
                assert isinstance(val, str) and val.startswith("="), (
                    f"Row {row} value is not a formula: {val!r}"
                )
        finally:
            wb2.close()

    def test_calculated_column_formulas_extend_when_table_grows(
        self,
        output_path: Path,
    ) -> None:
        """Even when growing beyond the original table size, formulas are filled."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "X"
        ws["B1"] = "Doubled"  # calculated: =[@X]*2
        ws["A2"] = 1

        col_formula = TableFormula(attr_text="[@X]*2")
        table = Table(
            displayName="Calc",
            ref="A1:B2",
            tableColumns=[
                TableColumn(id=1, name="X"),
                TableColumn(id=2, name="Doubled", calculatedColumnFormula=col_formula),
            ],
        )
        ws.add_table(table)
        tmpl = output_path.with_name("grow_calc_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        data = [{"X": i} for i in range(1, 11)]

        fill(template=tmpl, output=output_path, tables={"Calc": data})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Data"]
            for row in range(2, 12):
                val = ws2.cell(row=row, column=2).value
                assert val is not None, f"Row {row} formula missing"
                assert str(val).startswith("=")
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 9.2 Conditional formatting extension
# ═══════════════════════════════════════════════════════════════════════════════


class TestConditionalFormatting:
    """CF ranges extend and shrink with the table; can be disabled."""

    def test_conditional_formatting_extends_when_table_grows(
        self,
        output_path: Path,
    ) -> None:
        """CF ranges should cover all new rows after growth."""
        # Build template with a table that has CF
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Value"
        ws["A2"] = 5

        from openpyxl.formatting.rule import CellIsRule  # ty:ignore[unresolved-import]

        ws.conditional_formatting.add(
            "A2:A3",
            CellIsRule(
                operator="greaterThan",
                formula=["10"],
                fill=openpyxl.styles.PatternFill(  # ty:ignore[possibly-missing-submodule]
                    start_color="FF0000",
                    end_color="FF0000",
                    fill_type="solid",
                ),
            ),
        )

        table = Table(
            displayName="CFTable",
            ref="A1:A2",
            tableColumns=[TableColumn(id=1, name="Value")],
        )
        ws.add_table(table)
        tmpl = output_path.with_name("cf_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        data = [{"Value": i} for i in range(1, 8)]

        fill(template=tmpl, output=output_path, tables={"CFTable": data})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Data"]
            cf_ranges = [str(cf.sqref) for cf in ws2.conditional_formatting]
            # CF should now cover rows 2-8 (not just 2-3)
            assert any("A8" in r for r in cf_ranges) or any("8" in r for r in cf_ranges), (
                f"CF ranges did not extend: {cf_ranges}"
            )
        finally:
            wb2.close()

    def test_conditional_formatting_shrinks_when_table_shrinks(
        self,
        output_path: Path,
    ) -> None:
        """CF ranges contract when the table has fewer rows."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Value"
        ws["A2"] = 1
        ws["A3"] = 2
        ws["A4"] = 3

        from openpyxl.formatting.rule import CellIsRule  # ty:ignore[unresolved-import]

        ws.conditional_formatting.add(
            "A2:A4",
            CellIsRule(operator="greaterThan", formula=["5"], fill=openpyxl.styles.PatternFill()),  # ty:ignore[possibly-missing-submodule]
        )

        table = Table(
            displayName="ShrinkCF",
            ref="A1:A4",
            tableColumns=[TableColumn(id=1, name="Value")],
        )
        ws.add_table(table)
        tmpl = output_path.with_name("cf_shrink_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        data = [{"Value": 10}]

        fill(template=tmpl, output=output_path, tables={"ShrinkCF": data})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Data"]
            cf_ranges = [str(cf.sqref) for cf in ws2.conditional_formatting]
            # CF should now only cover row 2
            assert not any("A4" in r for r in cf_ranges), (
                f"CF ranges should not include row 4: {cf_ranges}"
            )
        finally:
            wb2.close()

    def test_conditional_formatting_extension_can_be_disabled(
        self,
        output_path: Path,
    ) -> None:
        """Set extend_conditional_formatting=False to leave CF ranges unchanged."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Value"
        ws["A2"] = 1

        from openpyxl.formatting.rule import CellIsRule  # ty:ignore[unresolved-import]

        ws.conditional_formatting.add(
            "A2:A2",
            CellIsRule(operator="greaterThan", formula=["0"], fill=openpyxl.styles.PatternFill()),  # ty:ignore[possibly-missing-submodule]
        )

        table = Table(
            displayName="NoExtCF",
            ref="A1:A2",
            tableColumns=[TableColumn(id=1, name="Value")],
        )
        ws.add_table(table)
        tmpl = output_path.with_name("cf_noext_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        data = [{"Value": i} for i in range(1, 6)]

        fill(
            template=tmpl,
            output=output_path,
            tables={"NoExtCF": data},
            config=FillXLConfig(extend_conditional_formatting=False),
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Data"]
            cf_ranges = [str(cf.sqref) for cf in ws2.conditional_formatting]
            # Original CF range (A2:A2) should NOT have grown
            assert not any("A6" in r for r in cf_ranges), (
                f"CF should not have extended: {cf_ranges}"
            )
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 9.3 Pivot table refresh
# ═══════════════════════════════════════════════════════════════════════════════


class TestPivotTableRefresh:
    """Pivot caches can be marked for refresh on open; default is True."""

    def test_refresh_pivot_tables_flag_is_true_by_default(self) -> None:
        """The config defaults to refreshing pivots on open."""
        cfg = FillXLConfig()
        assert cfg.refresh_pivot_tables is True

    def test_refresh_pivot_tables_can_be_disabled(self) -> None:
        """Set refresh_pivot_tables=False to skip pivot cache marking."""
        cfg = FillXLConfig(refresh_pivot_tables=False)
        assert cfg.refresh_pivot_tables is False


# ═══════════════════════════════════════════════════════════════════════════════
# 9.4 Formulas outside tables
# ═══════════════════════════════════════════════════════════════════════════════


class TestFormulasOutsideTables:
    """Regular Excel formulas outside tables are preserved after fill."""

    def test_formulas_outside_table_are_preserved(
        self,
        output_path: Path,
    ) -> None:
        """A SUM formula below a table remains intact after the table is filled."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Amount"
        ws["A2"] = 10
        ws["A3"] = 20
        # Leave a gap of empty rows between table end and formula
        ws["A10"] = "=SUM(A2:A3)"

        from openpyxl.worksheet.table import Table, TableColumn  # ty:ignore[unresolved-import]

        # Table has just 2 data rows (A2:A3), so the table ref is A1:A3
        table = Table(
            displayName="Amounts",
            ref="A1:A3",
            tableColumns=[TableColumn(id=1, name="Amount")],
        )
        ws.add_table(table)

        tmpl = output_path.with_name("formulas_tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        # Fill with more rows — table grows to A1:A7, formula at A10 is untouched
        data = [{"Amount": 100}, {"Amount": 200}, {"Amount": 300}, {"Amount": 400}, {"Amount": 500}]
        fill(template=tmpl, output=output_path, tables={"Amounts": data})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Data"]
            # Table data
            assert ws2["A2"].value == 100
            assert ws2["A6"].value == 500
            # Formula below the table should be preserved (empty rows between
            # table end at row 6 and formula at row 10)
            formula_cell = ws2["A10"].value
            assert formula_cell is not None, "Formula cell below table is empty"
            assert isinstance(formula_cell, str) and formula_cell.startswith("="), (
                f"Expected formula, got: {formula_cell!r}"
            )
        finally:
            wb2.close()
