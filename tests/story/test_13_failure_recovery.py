"""Chapter 13 — Failure Recovery.
===============================

When a fill operation fails, it's critical that:

* The original template is never modified.
* No partial or corrupt output file is left on disk.
* Resources (file handles, temp files) are properly cleaned up.

This chapter verifies that fillxl handles failures gracefully, leaving
the filesystem in a clean state no matter where in the pipeline the
failure occurs.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

import openpyxl
import pytest
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import FillXLConfig, fill
from fillxl._types import Engine

if TYPE_CHECKING:
    from pathlib import Path

# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_template(path: Path, table_name: str = "MyTable", rows: int = 1) -> None:
    """Create a workbook with a single table for failure tests."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Col"
    ws["A2"] = "placeholder"
    for r in range(2, 2 + rows):
        ws.cell(row=r, column=1, value=f"row{r}")

    last_row = rows + 1
    table = Table(
        displayName=table_name,
        ref=f"A1:A{last_row}",
        tableColumns=[TableColumn(id=1, name="Col")],
    )
    ws.add_table(table)
    wb.save(path)
    wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 13.1 Failure on table collision
# ═══════════════════════════════════════════════════════════════════════════════


class TestFailureOnTableCollision:
    """When a table would overwrite existing content, fill aborts cleanly."""

    def test_fill_aborts_when_a_table_collides_with_existing_content(
        self,
        output_path: Path,
    ) -> None:
        """Table expansion that would overwrite data raises ValueError and
        produces no output file.
        """
        _make_template(output_path.with_name("tmpl.xlsx"), "Blocked", rows=1)

        # Place content directly below the table
        wb = openpyxl.load_workbook(output_path.with_name("tmpl.xlsx"))
        ws = wb["Data"]
        ws["A5"] = "DO NOT OVERWRITE"
        wb.save(output_path.with_name("tmpl.xlsx"))
        wb.close()

        data = [{"Col": f"row{i}"} for i in range(1, 10)]

        with pytest.raises(ValueError, match="cannot expand"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                tables={"Blocked": data},
            )

        # No partial output file should be left behind
        assert not output_path.exists(), "Output file should not exist after a failed fill"

    def test_template_is_not_modified_when_table_collides(
        self,
        output_path: Path,
    ) -> None:
        """The original template file is untouched after a failed fill."""
        _make_template(output_path.with_name("tmpl.xlsx"), "Blocked", rows=1)

        wb = openpyxl.load_workbook(output_path.with_name("tmpl.xlsx"))
        ws = wb["Data"]
        ws["A5"] = "DO NOT OVERWRITE"
        wb.save(output_path.with_name("tmpl.xlsx"))
        wb.close()

        original_mtime = output_path.with_name("tmpl.xlsx").stat().st_mtime

        data = [{"Col": f"row{i}"} for i in range(1, 10)]

        with pytest.raises(ValueError, match="cannot expand"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                tables={"Blocked": data},
            )

        # Template modification time should be unchanged
        assert output_path.with_name("tmpl.xlsx").stat().st_mtime == original_mtime


# ═══════════════════════════════════════════════════════════════════════════════
# 13.2 Failure on validation
# ═══════════════════════════════════════════════════════════════════════════════


class TestFailureOnValidation:
    """When column validation fails, fill aborts before writing anything."""

    def test_fill_aborts_when_column_validation_fails(
        self,
        output_path: Path,
    ) -> None:
        """Missing a required table column raises ValueError and produces
        no output file.
        """
        _make_template(output_path.with_name("tmpl.xlsx"), "Validated", rows=1)

        # Data missing the 'Col' column
        data = [{"WrongColumn": 123}]

        with pytest.raises(ValueError, match=r"Missing columns|Column mismatch"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                tables={"Validated": data},
                config=FillXLConfig(column_validation_strategy="strict"),
            )

        # No partial output file
        assert not output_path.exists(), "Output file should not exist after a validation failure"

    def test_template_is_not_modified_when_validation_fails(
        self,
        output_path: Path,
    ) -> None:
        """The template file is untouched after a validation error."""
        _make_template(output_path.with_name("tmpl.xlsx"), "Validated", rows=1)

        original_mtime = output_path.with_name("tmpl.xlsx").stat().st_mtime

        data = [{"WrongColumn": 123}]

        with pytest.raises(ValueError, match="missing columns"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                tables={"Validated": data},
                config=FillXLConfig(column_validation_strategy="strict"),
            )

        assert output_path.with_name("tmpl.xlsx").stat().st_mtime == original_mtime


# ═══════════════════════════════════════════════════════════════════════════════
# 13.3 Failure on missing names
# ═══════════════════════════════════════════════════════════════════════════════


class TestFailureOnMissingNames:
    """Referencing non-existent names raises and leaves no output."""

    def test_fill_aborts_when_a_named_cell_does_not_exist(
        self,
        output_path: Path,
    ) -> None:
        """A missing named cell raises KeyError; no output file is created."""
        _make_template(output_path.with_name("tmpl.xlsx"))

        with pytest.raises(KeyError, match="not found"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                values={"NonExistentName": "value"},
            )

        assert not output_path.exists()

    def test_fill_aborts_when_a_table_name_does_not_exist(
        self,
        output_path: Path,
    ) -> None:
        """A missing table name raises KeyError; no output file is created."""
        _make_template(output_path.with_name("tmpl.xlsx"))

        with pytest.raises(KeyError):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                tables={"NonExistentTable": [{"Col": 1}]},
            )

        assert not output_path.exists()


# ═══════════════════════════════════════════════════════════════════════════════
# 13.4 Failure on invalid engine
# ═══════════════════════════════════════════════════════════════════════════════


class TestFailureOnInvalidEngine:
    """An invalid engine name is rejected before any file work begins."""

    def test_invalid_engine_raises_before_any_output_is_created(
        self,
        output_path: Path,
    ) -> None:
        """Passing an unknown engine raises ValueError before any I/O."""
        _make_template(output_path.with_name("tmpl.xlsx"))

        with pytest.raises(ValueError, match="Unknown engine"):
            fill(
                template=output_path.with_name("tmpl.xlsx"),
                output=output_path,
                engine=cast(Engine, "bogus"),
            )

        assert not output_path.exists()
