"""Chapter 1 — Basic fill() behaviour.
====================================

:func:`fillxl.fill` is the single entry point.  It accepts a template
workbook (``.xlsx``), an output path, and optional ``values`` and
``tables`` dictionaries.  The template is never modified — fillxl works
on a copy and writes the result to the output path.

Templates use **named cells** (Excel Name Manager) for scalar values
and **Excel tables** (Insert → Table) for tabular data.  This means
the Python code never references cell coordinates; it only references
names.  Charts, formulas, conditional formatting, and styles in the
template are preserved.

This chapter covers the function signature, parameter validation,
output handling, and the fill lifecycle.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import FillXLConfig, fill

if TYPE_CHECKING:
    from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════
# 1.1 Core fill() behaviour
# ═══════════════════════════════════════════════════════════════════════════════


class TestBasicFillBehaviour:
    """The fill() function: template, values, tables, and output."""

    def test_fill_copies_a_template_to_the_output_path_and_returns_it(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """Even with no data, fill() produces a valid .xlsx copy of the template."""
        result = fill(template=simple_table_template, output=output_path)
        assert result == output_path.resolve()
        assert output_path.exists()

    def test_named_cells_are_filled_from_the_values_dictionary(
        self,
        output_path: Path,
        named_cell_template: Path,
    ) -> None:
        """Each key in values matches a defined name in the workbook."""
        fill(
            template=named_cell_template,
            output=output_path,
            values={"my_cell": "world"},
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Sheet1"]["A1"].value == "world"
        finally:
            wb.close()

    def test_excel_tables_are_filled_from_the_tables_dictionary(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """Each key in tables matches an Excel table name; the table resizes to fit."""
        fill(
            template=simple_table_template,
            output=output_path,
            tables={"SimpleTable": [{"Name": "Alice", "Value": 42}]},
        )
        wb = openpyxl.load_workbook(output_path)
        try:
            assert wb["Data"]["A2"].value == "Alice"
            assert wb["Data"]["B2"].value == 42
        finally:
            wb.close()

    def test_values_and_tables_can_be_combined_in_a_single_fill_call(
        self,
        output_path: Path,
    ) -> None:
        """A single call populates both named cells and tables."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"
        ws["A1"] = "Name"
        ws["B1"] = "Value"
        ws["C1"] = "Old Title"

        from openpyxl.workbook.defined_name import DefinedName
        from openpyxl.worksheet.table import Table, TableColumn

        wb.defined_names.add(DefinedName("ReportTitle", attr_text="Report!$C$1"))
        ws.add_table(
            Table(
                displayName="Data",
                ref="A1:B2",
                tableColumns=[
                    TableColumn(id=1, name="Name"),
                    TableColumn(id=2, name="Value"),
                ],
            ),
        )
        tmpl = output_path.with_name("combined.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"ReportTitle": "Q4 Sales Report"},
            tables={"Data": [{"Name": "Widget", "Value": 100}]},
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            assert wb2["Report"]["C1"].value == "Q4 Sales Report"
            assert wb2["Report"]["A2"].value == "Widget"
            assert wb2["Report"]["B2"].value == 100
        finally:
            wb2.close()

    def test_the_template_file_is_not_modified_by_fill(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """fill() copies the template internally; the original file is untouched."""
        original_mtime = simple_table_template.stat().st_mtime

        fill(
            template=simple_table_template,
            output=output_path,
            tables={"SimpleTable": [{"Name": "Alice", "Value": 42}]},
        )

        assert simple_table_template.stat().st_mtime == original_mtime

    def test_existing_cell_formatting_is_preserved_when_values_are_written(
        self,
        output_path: Path,
    ) -> None:
        """Fillxl replaces the value but keeps the cell's number format."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = 0.5
        ws["A1"].number_format = "0.00%"

        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names.add(DefinedName("pct", attr_text="Data!$A$1"))
        tmpl = output_path.with_name("fmt.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(template=tmpl, output=output_path, values={"pct": 0.75})

        wb2 = openpyxl.load_workbook(output_path)
        try:
            assert wb2["Data"]["A1"].value == 0.75
            assert wb2["Data"]["A1"].number_format == "0.00%"
        finally:
            wb2.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 1.2 The template parameter
# ═══════════════════════════════════════════════════════════════════════════════


class TestTemplateParameter:
    """The template is the Excel workbook that gets filled with data."""

    def test_template_can_be_a_string_path_to_an_xlsx_file(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """A plain string path works just fine."""
        result = fill(
            template=str(simple_table_template),
            output=output_path,
        )
        assert result == output_path.resolve()
        assert output_path.exists()

    def test_template_can_be_a_pathlib_path(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """A pathlib.Path is also accepted."""
        fill(
            template=simple_table_template,
            output=output_path,
        )
        assert output_path.exists()

    def test_a_nonexistent_template_raises_filenotfounderror_before_any_work(
        self,
        output_path: Path,
    ) -> None:
        """The template path is validated immediately — no output file is created."""
        with pytest.raises(FileNotFoundError, match="not found"):
            fill(template="missing.xlsx", output=output_path)
        assert not output_path.exists()


# ═══════════════════════════════════════════════════════════════════════════════
# 1.3 The output parameter
# ═══════════════════════════════════════════════════════════════════════════════


class TestOutputParameter:
    """The output is where the filled workbook is saved."""

    def test_output_can_be_a_string_path(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """A plain string works for the output path too."""
        fill(
            template=simple_table_template,
            output=str(output_path),
        )
        assert output_path.exists()

    def test_output_can_be_a_pathlib_path(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """And so does pathlib.Path."""
        result = fill(
            template=simple_table_template,
            output=output_path,
        )
        assert isinstance(result, Path)

    def test_missing_directories_in_the_output_path_are_created_automatically(
        self,
        tmp_path: Path,
        simple_table_template: Path,
    ) -> None:
        """You don't need to create directories — fill() does it for you."""
        nested = tmp_path / "deep" / "nested" / "folder" / "out.xlsx"
        fill(template=simple_table_template, output=nested)
        assert nested.exists()

    def test_an_existing_output_file_is_overwritten(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """If output already exists, it's replaced with the new file."""
        # First fill
        fill(template=simple_table_template, output=output_path)

        # Second fill — should overwrite
        fill(template=simple_table_template, output=output_path)

        assert output_path.exists()

    def test_the_return_value_is_the_absolute_path_to_the_output_file(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """fill() returns the absolute Path to the output, for chaining."""
        result = fill(template=simple_table_template, output=output_path)
        assert result == output_path.resolve()
        assert result.is_absolute()


# ═══════════════════════════════════════════════════════════════════════════════
# 1.4 The fill lifecycle
# ═══════════════════════════════════════════════════════════════════════════════


class TestFillLifecycle:
    """The fill() pipeline: values → tables → pivots → sheets → save."""

    def test_an_empty_fill_call_produces_a_valid_xlsx_file(
        self,
        output_path: Path,
        simple_table_template: Path,
    ) -> None:
        """Even with no values or tables, fill() copies the template to output."""
        fill(template=simple_table_template, output=output_path)
        wb = openpyxl.load_workbook(output_path)
        try:
            assert "Data" in wb.sheetnames
        finally:
            wb.close()

    def test_the_fill_pipeline_runs_in_a_fixed_order(
        self,
        output_path: Path,
    ) -> None:
        """fill() executes a fixed pipeline: values → tables → pivots →
        hide sheets → delete sheets → set active sheet → save.

        This test verifies that all steps run in a single call and the
        output reflects every stage.
        """
        # Build a template with a table, a named cell, and a _Config sheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Name"
        ws["B1"] = "Value"
        ws["A2"] = "old"
        ws["B2"] = 0
        ws["C1"] = "Old Title"

        from openpyxl.workbook.defined_name import DefinedName
        from openpyxl.worksheet.table import Table, TableColumn

        wb.defined_names.add(DefinedName("title", attr_text="Data!$C$1"))
        ws.add_table(
            Table(
                displayName="People",
                ref="A1:B2",
                tableColumns=[
                    TableColumn(id=1, name="Name"),
                    TableColumn(id=2, name="Value"),
                ],
            ),
        )
        wb.create_sheet("_Config").cell(row=1, column=1, value="secret")
        tmpl = output_path.with_name("pipeline.xlsx")
        wb.save(tmpl)
        wb.close()

        fill(
            template=tmpl,
            output=output_path,
            values={"title": "Pipeline Test"},
            tables={"People": [{"Name": "Alice", "Value": 100}]},
            config=FillXLConfig(active_sheet="Data"),
        )

        wb2 = openpyxl.load_workbook(output_path)
        try:
            # Named cell filled
            assert wb2["Data"]["C1"].value == "Pipeline Test"
            # Table filled
            assert wb2["Data"]["A2"].value == "Alice"
            # _Config sheet hidden
            assert wb2["_Config"].sheet_state == "hidden"
            # Active sheet set
            assert wb2.active.title == "Data"
        finally:
            wb2.close()

    def test_fill_is_idempotent(
        self,
        output_path: Path,
        frozen_clock: datetime,
    ) -> None:
        """Filling the same template twice produces identical output.

        ``frozen_clock`` pins both the wall clock and the timestamps openpyxl
        writes into the zip entries, so the byte-for-byte comparison below
        cannot flake on a two-second zip timestamp boundary.
        """
        import hashlib
        import zipfile

        # Build a template
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Name"
        ws["B1"] = "Value"
        ws["A2"] = "old"
        ws["B2"] = 0

        from openpyxl.workbook.defined_name import DefinedName
        from openpyxl.worksheet.table import Table, TableColumn

        wb.defined_names.add(DefinedName("title", attr_text="Data!$C$1"))
        ws.add_table(
            Table(
                displayName="Items",
                ref="A1:B2",
                tableColumns=[
                    TableColumn(id=1, name="Name"),
                    TableColumn(id=2, name="Value"),
                ],
            ),
        )
        tmpl = output_path.with_name("idempotent.xlsx")
        wb.save(tmpl)
        wb.close()

        data = [{"Name": "Alice", "Value": 100}]

        # First fill
        out1 = output_path.with_name("out1.xlsx")
        fill(template=tmpl, output=out1, values={"title": "Test"}, tables={"Items": data})

        # Second fill with same inputs
        out2 = output_path.with_name("out2.xlsx")
        fill(template=tmpl, output=out2, values={"title": "Test"}, tables={"Items": data})

        def sha256(path: Path) -> str:
            return hashlib.sha256(path.read_bytes()).hexdigest()

        assert sha256(out1) == sha256(out2), (
            "Two fills with same inputs should produce identical output files"
        )

        # Every zip entry must be stamped from the frozen clock rather than the
        # real one; that is what makes the comparison above wall-clock
        # independent.  (The stamp is the frozen instant rendered in the
        # machine's local time, hence the date-only check.)
        frozen_date = (frozen_clock.year, frozen_clock.month, frozen_clock.day)
        for path in (out1, out2):
            with zipfile.ZipFile(path) as archive:
                stamps = {info.date_time for info in archive.infolist()}
            assert {stamp[:3] for stamp in stamps} == {frozen_date}, (
                f"zip timestamps in {path.name} came from the real clock: {stamps}"
            )
            assert len(stamps) == 1, f"mixed zip timestamps in {path.name}: {stamps}"
