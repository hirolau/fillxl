"""Chapter 14 — Real Report Generation.
=====================================

The ultimate test of fillxl is generating a complete report workbook
from a template.  This chapter combines named cells, tables, and charts
in a single :func:`fillxl.fill` call, then verifies that everything is
correctly updated and the resulting workbook can be opened without error.

This is the integration test that catches bugs the isolated stories miss.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
from openpyxl.chart import BarChart, Reference  # ty:ignore[unresolved-import]
from openpyxl.worksheet.table import Table, TableColumn  # ty:ignore[unresolved-import]

from fillxl import fill

if TYPE_CHECKING:
    from pathlib import Path

# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_report_template(path: Path) -> None:
    """Build a realistic report template with title, table, and chart."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"

    from openpyxl.workbook.defined_name import DefinedName

    # ── Named cells for report metadata ──────────────────────────────────
    ws["A1"] = "Report Title"
    ws["A2"] = "2025-01-01"
    ws["A3"] = "Prepared by"
    ws["A4"] = "Draft"

    wb.defined_names.add(DefinedName("title", attr_text="Report!$A$1"))
    wb.defined_names.add(DefinedName("report_date", attr_text="Report!$A$2"))
    wb.defined_names.add(DefinedName("prepared_by", attr_text="Report!$A$3"))
    wb.defined_names.add(DefinedName("status", attr_text="Report!$A$4"))

    # ── Table for sales data ─────────────────────────────────────────────
    ws["A6"] = "Product"
    ws["B6"] = "Region"
    ws["C6"] = "Sales"
    ws["D6"] = "Quantity"

    ws["A7"] = "Widget"
    ws["B7"] = "North"
    ws["C7"] = 100
    ws["D7"] = 10

    table = Table(
        displayName="SalesData",
        ref="A6:D7",
        tableColumns=[
            TableColumn(id=1, name="Product"),
            TableColumn(id=2, name="Region"),
            TableColumn(id=3, name="Sales"),
            TableColumn(id=4, name="Quantity"),
        ],
    )
    ws.add_table(table)

    # ── Summary values (also named cells) ────────────────────────────────
    ws["F1"] = "Total Sales"
    ws["G1"] = 0
    ws["F2"] = "Total Qty"
    ws["G2"] = 0
    wb.defined_names.add(DefinedName("total_sales", attr_text="Report!$G$1"))
    wb.defined_names.add(DefinedName("total_qty", attr_text="Report!$G$2"))

    # ── Chart bound to the sales table ───────────────────────────────────
    chart_ws = wb.create_sheet("Chart")
    chart = BarChart()
    chart.title = "Sales by Product"

    data_ref = Reference(ws, min_col=3, min_row=6, max_row=7)
    cats_ref = Reference(ws, min_col=1, min_row=7, max_row=7)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart_ws.add_chart(chart, "A1")

    wb.save(path)
    wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 14.1 Full report generation
# ═══════════════════════════════════════════════════════════════════════════════


class TestFullReportGeneration:
    """A complete report: named cells, table, and chart all update together."""

    def test_report_title_is_updated(
        self,
        output_path: Path,
    ) -> None:
        """The title named cell is filled."""
        _make_report_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={
                "title": "Q4 2025 Sales Report",
                "report_date": "2025-12-31",
                "prepared_by": "Finance Team",
                "status": "Final",
                "total_sales": 12500,
                "total_qty": 340,
            },
            tables={
                "SalesData": [
                    {"Product": "Widget", "Region": "North", "Sales": 3000, "Quantity": 30},
                    {"Product": "Widget", "Region": "South", "Sales": 4500, "Quantity": 45},
                    {"Product": "Gadget", "Region": "North", "Sales": 2000, "Quantity": 20},
                    {"Product": "Gadget", "Region": "South", "Sales": 3000, "Quantity": 35},
                ],
            },
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Report"]
            assert ws["A1"].value == "Q4 2025 Sales Report"
        finally:
            wb.close()

    def test_table_is_populated_with_all_rows(
        self,
        output_path: Path,
    ) -> None:
        """All data rows are written to the table."""
        _make_report_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={
                "title": "Q4 Report",
                "report_date": "2025-12-31",
                "prepared_by": "Finance",
                "status": "Final",
            },
            tables={
                "SalesData": [
                    {"Product": "Widget", "Region": "North", "Sales": 3000, "Quantity": 30},
                    {"Product": "Widget", "Region": "South", "Sales": 4500, "Quantity": 45},
                    {"Product": "Gadget", "Region": "North", "Sales": 2000, "Quantity": 20},
                    {"Product": "Gadget", "Region": "South", "Sales": 3000, "Quantity": 35},
                ],
            },
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Report"]
            assert ws["A7"].value == "Widget"
            assert ws["C7"].value == 3000
            assert ws["A10"].value == "Gadget"
            assert ws["D10"].value == 35
            # Row 11 should be empty (4 data rows + 1 header = 5 rows, so A6:D10)
            assert ws["A11"].value is None
        finally:
            wb.close()

    def test_chart_references_are_updated_to_match_the_grown_table(
        self,
        output_path: Path,
    ) -> None:
        """Chart data ranges expand to cover the new rows."""
        _make_report_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={
                "title": "Q4 Report",
                "report_date": "2025-12-31",
                "prepared_by": "Finance",
                "status": "Final",
            },
            tables={
                "SalesData": [
                    {"Product": "Widget", "Region": "North", "Sales": 3000, "Quantity": 30},
                    {"Product": "Widget", "Region": "South", "Sales": 4500, "Quantity": 45},
                    {"Product": "Gadget", "Region": "North", "Sales": 2000, "Quantity": 20},
                    {"Product": "Gadget", "Region": "South", "Sales": 3000, "Quantity": 35},
                ],
            },
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            chart_ws = wb["Chart"]
            for chart in chart_ws._charts:
                for ser in chart.ser:
                    for label in ("cat", "val", "xVal", "yVal"):
                        item = getattr(ser, label, None)
                        if item is None:
                            continue
                        for ref_type in ("strRef", "numRef"):
                            ref = getattr(item, ref_type, None)
                            if ref is not None and ref.f:
                                # 4 data rows + 1 header = row 10
                                end_row = int(
                                    "".join(ch for ch in ref.f.split("$")[-1] if ch.isdigit()),
                                )
                                assert end_row == 10, (
                                    f"Chart ref {ref.f} ends at row {end_row}, expected 10"
                                )
        finally:
            wb.close()

    def test_summary_values_are_updated(
        self,
        output_path: Path,
    ) -> None:
        """Named cells for summary values are filled correctly."""
        _make_report_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={
                "title": "Q4 Report",
                "report_date": "2025-12-31",
                "prepared_by": "Finance",
                "status": "Final",
                "total_sales": 12500,
                "total_qty": 340,
            },
            tables={
                "SalesData": [
                    {"Product": "Widget", "Region": "North", "Sales": 3000, "Quantity": 30},
                    {"Product": "Widget", "Region": "South", "Sales": 4500, "Quantity": 45},
                    {"Product": "Gadget", "Region": "North", "Sales": 2000, "Quantity": 20},
                    {"Product": "Gadget", "Region": "South", "Sales": 3000, "Quantity": 35},
                ],
            },
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Report"]
            assert ws["G1"].value == 12500
            assert ws["G2"].value == 340
        finally:
            wb.close()

    def test_resulting_workbook_opens_successfully_with_openpyxl(
        self,
        output_path: Path,
    ) -> None:
        """The output workbook can be loaded by openpyxl without error."""
        _make_report_template(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            values={
                "title": "Q4 Report",
                "report_date": "2025-12-31",
                "prepared_by": "Finance",
                "status": "Final",
            },
            tables={
                "SalesData": [
                    {"Product": "Widget", "Region": "North", "Sales": 3000, "Quantity": 30},
                    {"Product": "Widget", "Region": "South", "Sales": 4500, "Quantity": 45},
                    {"Product": "Gadget", "Region": "North", "Sales": 2000, "Quantity": 20},
                    {"Product": "Gadget", "Region": "South", "Sales": 3000, "Quantity": 35},
                ],
            },
        )

        # Should load without exception
        wb = openpyxl.load_workbook(output_path)
        try:
            # All expected sheets present
            assert "Report" in wb.sheetnames
            assert "Chart" in wb.sheetnames
        finally:
            wb.close()
