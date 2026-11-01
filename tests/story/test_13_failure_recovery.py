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


# ═══════════════════════════════════════════════════════════════════════════════
# 13.5 A failure never costs you the last report
# ═══════════════════════════════════════════════════════════════════════════════

# Every stage the output has to survive.  Each one is injected in its own test,
# because a failure that quietly destroys yesterday's report is far worse than
# one that leaves no report at all.
_FAILURE_STAGES = {
    "save": "fillxl._backends._openpyxl.OpenpyxlBackend.save",
    "close": "fillxl._backends._openpyxl.OpenpyxlBackend.close",
    "loss check": "fillxl._backends._openpyxl.check_lossy",
    "replace": "pathlib.Path.replace",
}


def _make_previous_output(path: Path) -> None:
    """Write a complete, valid workbook where the next fill will write."""
    wb = openpyxl.Workbook()
    wb.active.title = "Yesterday"
    wb.active["A1"] = "last week's report"
    wb.save(path)
    wb.close()


def _staging_files(directory: Path) -> list[Path]:
    """Staging files fillxl may leave behind in *directory*."""
    return sorted(directory.glob(".*.fillxl*"))


class TestAFailureKeepsThePreviousOutput:
    """The destination is replaced only after the whole fill has succeeded."""

    @pytest.mark.parametrize("stage", list(_FAILURE_STAGES))
    def test_a_failure_at_any_stage_keeps_the_previous_output(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        stage: str,
    ) -> None:
        """Saving, closing, loss checking and the final swap all run first.

        Whenever one of them fails, the file that was already at ``output=``
        is still there, still readable, and still holds its own content.
        """
        template = tmp_path / "tmpl.xlsx"
        _make_template(template)
        template_before = template.read_bytes()

        output = tmp_path / "report.xlsx"
        _make_previous_output(output)
        output_before = output.read_bytes()

        seen_staging: list[Path] = []

        def fail(*_args: object, **_kwargs: object) -> None:
            seen_staging.extend(_staging_files(tmp_path))
            msg = f"injected {stage} failure"
            raise RuntimeError(msg)

        monkeypatch.setattr(_FAILURE_STAGES[stage], fail)

        with pytest.raises(RuntimeError, match="injected"):
            fill(
                template=template,
                output=output,
                tables={"MyTable": [{"Col": "new value"}]},
            )

        assert output.read_bytes() == output_before, "the previous output was destroyed"
        assert template.read_bytes() == template_before, "the template was modified"
        assert _staging_files(tmp_path) == [], "a staging file was left behind"
        if stage != "save":
            # Everything after saving saw a finished staging file.  These are
            # the cases where a careless implementation has already thrown
            # away (or replaced) the destination by now.
            assert seen_staging, f"no staging file existed at the {stage} stage"

        previous = openpyxl.load_workbook(output)
        assert previous["Yesterday"]["A1"].value == "last week's report"
        previous.close()

    @pytest.mark.parametrize("stage", list(_FAILURE_STAGES))
    def test_the_next_fill_still_works_after_a_failure(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        stage: str,
    ) -> None:
        """The failure leaves nothing behind that blocks the retry."""
        template = tmp_path / "tmpl.xlsx"
        _make_template(template)
        output = tmp_path / "report.xlsx"
        _make_previous_output(output)

        def fail(*_args: object, **_kwargs: object) -> None:
            msg = f"injected {stage} failure"
            raise RuntimeError(msg)

        with monkeypatch.context() as patched:
            patched.setattr(_FAILURE_STAGES[stage], fail)
            with pytest.raises(RuntimeError, match="injected"):
                fill(template=template, output=output, tables={"MyTable": [{"Col": "v"}]})

        fill(template=template, output=output, tables={"MyTable": [{"Col": "v"}]})

        filled = openpyxl.load_workbook(output)
        assert filled["Data"]["A2"].value == "v"
        filled.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 13.6 Filling a template onto itself
# ═══════════════════════════════════════════════════════════════════════════════


class TestFillingOntoTheTemplatePath:
    """``template`` and ``output`` may be the same path."""

    def test_the_filled_workbook_replaces_the_template(self, tmp_path: Path) -> None:
        """The template is copied before it is opened, so in-place runs are safe."""
        path = tmp_path / "in_place.xlsx"
        _make_template(path)

        result = fill(
            template=path,
            output=path,
            tables={"MyTable": [{"Col": "filled"}, {"Col": "again"}]},
        )

        assert result == path.resolve()
        filled = openpyxl.load_workbook(path)
        ws = filled["Data"]
        assert ws["A2"].value == "filled"
        assert ws["A3"].value == "again"
        assert ws.tables["MyTable"].ref == "A1:A3"
        filled.close()
        assert _staging_files(tmp_path) == []

    def test_a_failed_in_place_fill_keeps_the_template_intact(self, tmp_path: Path) -> None:
        """An in-place run that fails must not eat the only copy of the template."""
        path = tmp_path / "in_place.xlsx"
        _make_template(path)
        before = path.read_bytes()

        with pytest.raises(KeyError, match="not found"):
            fill(template=path, output=path, tables={"NoSuchTable": [{"Col": 1}]})

        assert path.read_bytes() == before
        assert _staging_files(tmp_path) == []
