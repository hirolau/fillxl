"""Chapter 12 — Template Preservation.
====================================

The core promise of fillxl is:

    "Update data without breaking the workbook."

When you fill specific named cells or tables, everything else in the
template should remain exactly as it was — charts not related to the
changed data, cell formatting, formulas outside tables, hidden sheets,
and named cells that weren't in the fill request.

This chapter verifies that fillxl does NOT touch what it was not asked
to touch.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
from openpyxl.chart import BarChart, Reference  # ty:ignore[unresolved-import]
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableColumn  # ty:ignore[unresolved-import]

from fillxl import fill

if TYPE_CHECKING:
    from pathlib import Path

# ═══════════════════════════════════════════════════════════════════════════════
# 12.1 Charts unrelated to modified tables
# ═══════════════════════════════════════════════════════════════════════════════


class TestChartsUnrelatedToModifiedTables:
    """Charts that reference a different table are left untouched."""

    def test_chart_bound_to_an_unmodified_table_remains_unchanged(
        self,
        output_path: Path,
    ) -> None:
        """Filling TableOne does not affect a chart bound to TableTwo."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"

        # Two tables side by side
        headers = ["Category", "Value"]
        for ci, h in enumerate(headers, 1):
            ws.cell(row=1, column=ci, value=h)
            ws.cell(row=1, column=ci + 3, value=h)

        for ri in range(3):
            ws.cell(row=ri + 2, column=1, value=f"A{ri}")
            ws.cell(row=ri + 2, column=2, value=(ri + 1) * 10)
            ws.cell(row=ri + 2, column=4, value=f"B{ri}")
            ws.cell(row=ri + 2, column=5, value=(ri + 1) * 100)

        t1 = Table(
            displayName="TableOne",
            ref="A1:B5",
            tableColumns=[TableColumn(id=1, name="Category"), TableColumn(id=2, name="Value")],
        )
        t2 = Table(
            displayName="TableTwo",
            ref="D1:E5",
            tableColumns=[TableColumn(id=1, name="Category"), TableColumn(id=2, name="Value")],
        )
        ws.add_table(t1)
        ws.add_table(t2)

        chart_ws = wb.create_sheet("Charts")

        # Chart bound to TableTwo (which we will NOT modify)
        chart = BarChart()
        chart.title = "TableTwo Chart"
        data_ref = Reference(ws, min_col=5, min_row=1, max_row=5)
        cats_ref = Reference(ws, min_col=4, min_row=2, max_row=5)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart_ws.add_chart(chart, "A1")

        tmpl = output_path.with_name("tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        # Fill only TableOne
        fill(
            template=tmpl,
            output=output_path,
            tables={"TableOne": [{"Category": "X", "Value": 99}]},
        )

        # Verify TableTwo chart references are unchanged
        wb2 = openpyxl.load_workbook(output_path)
        try:
            chart_ws2 = wb2["Charts"]
            for c in chart_ws2._charts:
                for ser in c.ser:
                    for label in ("cat", "val", "xVal", "yVal"):
                        item = getattr(ser, label, None)
                        if item is None:
                            continue
                        for ref_type in ("strRef", "numRef"):
                            ref = getattr(item, ref_type, None)
                            if ref is not None and ref.f:
                                # Should still reference row 5 (original table size)
                                assert "$5" in ref.f, f"Chart ref changed: {ref.f}"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 12.2 Formatting outside filled areas
# ═══════════════════════════════════════════════════════════════════════════════


class TestFormattingOutsideFilledAreas:
    """Cell formatting not in the filled range is preserved."""

    def test_font_formatting_outside_tables_remains_unchanged(
        self,
        output_path: Path,
    ) -> None:
        """A cell with special formatting not in a fill target keeps its style."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"

        from openpyxl.workbook.defined_name import DefinedName

        # A named cell we will fill
        ws["A1"] = "old_title"
        ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FF0000")
        wb.defined_names.add(DefinedName("title", attr_text="Report!$A$1"))

        # A cell we will NOT fill — should retain its style
        ws["C1"] = "Static Header"
        ws["C1"].font = Font(name="Arial", size=16, italic=True, color="00AA00")
        ws["C1"].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # A cell below the named cell
        ws["A5"] = "Footer"
        ws["A5"].font = Font(name="Times New Roman", size=10)

        tmpl = output_path.with_name("tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"title": "New Report Title"},
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Report"]
            # Filled cell changed
            assert ws2["A1"].value == "New Report Title"
            # Unrelated cell C1: value should be None (we wrote nothing),
            # but font/fill should be intact
            c1_font = ws2["C1"].font
            assert c1_font.name == "Arial"
            assert c1_font.size == 16
            assert c1_font.italic is True
            # Fill check — openpyxl prepends alpha ("00"), so "FFFF00" becomes "00FFFF00"
            c1_fill = ws2["C1"].fill
            assert c1_fill.start_color.rgb == "00FFFF00" or c1_fill.start_color.value == "00FFFF00"
        finally:
            wb2.close()

    def test_number_format_outside_tables_is_preserved(
        self,
        output_path: Path,
    ) -> None:
        """A cell with a custom number format not involved in fill keeps it."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["B1"] = 0.5
        ws["B1"].number_format = "0.00%"
        ws["B2"] = 1000
        ws["B2"].number_format = "#,##0"

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("pct_val", attr_text="Sheet1!$B$1"))
        tmpl = output_path.with_name("tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        # Fill only B1
        fill(
            template=tmpl,
            output=output_path,
            values={"pct_val": 0.75},
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            ws2 = wb2["Sheet1"]
            # Filled cell updated
            assert ws2["B1"].value == 0.75
            # B2 (not filled) retains number format
            assert ws2["B2"].number_format == "#,##0"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 12.3 Hidden sheets
# ═══════════════════════════════════════════════════════════════════════════════


class TestHiddenSheetPreservation:
    """Sheets hidden in the template stay hidden after fill."""

    def test_hidden_sheets_remain_hidden_after_fill(
        self,
        output_path: Path,
    ) -> None:
        """A sheet that was hidden in the template is not made visible."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Visible"
        ws["A1"] = "hello"

        hidden = wb.create_sheet("HiddenConfig")
        hidden.sheet_state = "hidden"
        hidden["A1"] = "secret"

        very_hidden = wb.create_sheet("SuperHidden")
        very_hidden.sheet_state = "veryHidden"
        very_hidden["A1"] = "top secret"

        tmpl = output_path.with_name("tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(template=tmpl, output=output_path)

        wb2 = openpyxl.load_workbook(output_path)
        try:
            assert wb2["Visible"].sheet_state == "visible"
            assert wb2["HiddenConfig"].sheet_state == "hidden"
            assert wb2["SuperHidden"].sheet_state == "veryHidden"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 12.4 Named cells not in values are preserved
# ═══════════════════════════════════════════════════════════════════════════════


class TestUnrelatedNamedCells:
    """Named cells not included in the values dict keep their original content."""

    def test_unfilled_named_cells_retain_their_original_values(
        self,
        output_path: Path,
    ) -> None:
        """Only the named cells you specify in values are updated."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "keep this"
        ws["B1"] = "change this"

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("keep_me", attr_text="Data!$A$1"))
        wb.defined_names.add(DefinedName("change_me", attr_text="Data!$B$1"))

        tmpl = output_path.with_name("tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"change_me": "updated"},
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            assert wb2["Data"]["A1"].value == "keep this"
            assert wb2["Data"]["B1"].value == "updated"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 12.5 Multiple sheets — changes stay scoped
# ═══════════════════════════════════════════════════════════════════════════════


class TestMultiSheetPreservation:
    """Changes to one sheet do not affect other sheets."""

    def test_filling_a_table_on_one_sheet_does_not_affect_another_sheet(
        self,
        output_path: Path,
    ) -> None:
        """Data on Sheet2 is untouched when filling a table on Sheet1."""
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Sheet1"
        ws1["A1"] = "Name"
        ws1["A2"] = "old"

        t1 = Table(
            displayName="Data",
            ref="A1:A2",
            tableColumns=[TableColumn(id=1, name="Name")],
        )
        ws1.add_table(t1)

        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = "Do not touch"
        ws2["B1"] = 42

        tmpl = output_path.with_name("tmpl.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            tables={"Data": [{"Name": "Alice"}, {"Name": "Bob"}]},
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            # Sheet1 updated
            assert wb2["Sheet1"]["A2"].value == "Alice"
            assert wb2["Sheet1"]["A3"].value == "Bob"
            # Sheet2 untouched
            assert wb2["Sheet2"]["A1"].value == "Do not touch"
            assert wb2["Sheet2"]["B1"].value == 42
        finally:
            wb2.close()
