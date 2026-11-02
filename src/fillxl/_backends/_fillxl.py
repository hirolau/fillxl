"""The ``fillxl`` engine: fill an OOXML package without Excel and without loss.

This backend does not use a workbook object model.  It copies the template's
parts byte for byte and rewrites only the ones a fill has to change, so
everything the openpyxl model cannot represent — spilled dynamic arrays and
their ``xl/metadata.xml``, ``cm``/``vm`` cell metadata, slicers, ``chartEx``,
rich data, VBA — survives by construction.

See ``IMPLEMENTATION-LOG.md`` for the design decisions.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from fillxl._lossy import check_lossy
from fillxl._ooxml.charts import update_chart_references
from fillxl._ooxml.dangling import warn_about_deleted_sheets
from fillxl._ooxml.fill import fill_named_cell, fill_table
from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.pivots import set_refresh_on_load
from fillxl._ooxml.workbook import WorkbookPart
from fillxl._sheets import first_remaining, require_a_visible_sheet, warn_sheet_not_found

if TYPE_CHECKING:
    from collections.abc import Sequence
    from pathlib import Path

    from fillxl._charts import TableChangeSpecs
    from fillxl._config import FillXLConfig
    from fillxl._lossy import OnLossy

logger = logging.getLogger(__name__)

# Only zip-based OOXML workbooks: .xlsb is binary OOXML and .xls is BIFF8, and
# an encrypted .xlsx is not a zip at all.  All three stay on engine="win32com".
_FILLXL_SUFFIXES = (".xlsx", ".xlsm")


class FillxlBackend:
    """Backend that edits the workbook package directly (no Excel, no model)."""

    def __init__(self) -> None:
        self._package: OoxmlPackage | None = None
        self._workbook: WorkbookPart | None = None
        self._source_suffix: str | None = None
        self._deleted_sheets: list[str] = []

    def _require_workbook(self) -> tuple[OoxmlPackage, WorkbookPart]:
        """Return the loaded package and workbook, or raise."""
        if self._package is None or self._workbook is None:
            msg = "No workbook loaded — call load() first"
            raise RuntimeError(msg)
        return self._package, self._workbook

    # ── Backend protocol ────────────────────────────────────────────────────

    def load(self, path: Path) -> None:
        """Open the workbook package at *path*.

        Args:
            path: Path to the workbook file.

        Raises:
            ValueError: If the file is not an ``.xlsx``/``.xlsm`` package.

        """
        suffix = path.suffix.lower()
        if suffix not in _FILLXL_SUFFIXES:
            msg = (
                f"File type '{suffix}' is not supported by the fillxl engine "
                "(supported: .xlsx, .xlsm). Use engine='win32com' for .xlsb, .xls "
                "and format conversion."
            )
            raise ValueError(msg)
        logger.debug("Loading workbook from %s", path)
        self._source_suffix = suffix
        self._package = OoxmlPackage(path)
        self._workbook = WorkbookPart(self._package)

    def fill_named_cell(self, name: str, value: Any, config: FillXLConfig) -> None:  # noqa: ANN401
        """Write a scalar *value* to the named cell *name*.

        Args:
            name: Name of the named cell / defined name.
            value: Scalar value to write.
            config: Global fill configuration.

        """
        package, workbook = self._require_workbook()
        fill_named_cell(
            package,
            workbook,
            name,
            value,
            escape_formulas=config.escape_formulas,
        )

    def fill_table(
        self,
        name: str,
        data: Any,  # noqa: ANN401
        config: FillXLConfig,
    ) -> TableChangeSpecs:
        """Fill the Excel table *name* with tabular *data*.

        Args:
            name: Name of the Excel table.
            data: Tabular data (DataFrame, list[dict], …).
            config: Global fill configuration.

        Returns:
            A :class:`TableChangeSpecs` describing the old and new dimensions.

        """
        package, workbook = self._require_workbook()
        return fill_table(package, workbook, name, data, config)

    def update_charts(self, spec: TableChangeSpecs) -> None:
        """Update chart data ranges after a table resize.

        Args:
            spec: The table dimension change to reflect in charts.

        """
        package, _ = self._require_workbook()
        update_chart_references(package, spec)

    def set_pivot_refresh(self) -> None:
        """Mark all pivot caches for refresh on load."""
        package, _ = self._require_workbook()
        set_refresh_on_load(package)

    def hide_sheets(self, prefix: str) -> None:
        """Hide visible sheets whose names start with *prefix*.

        Args:
            prefix: Sheet name prefix (e.g. ``"_"``).

        """
        _, workbook = self._require_workbook()
        visible = [entry.name for entry in workbook.sheets() if entry.visible]
        to_hide = [name for name in visible if name.startswith(prefix)]
        if not to_hide:
            return
        require_a_visible_sheet(
            visible,
            to_hide,
            "Hiding",
            hint="Set hide_underscore_sheets=False to keep it visible.",
        )
        active = workbook.active_sheet_name()
        for name in to_hide:
            workbook.set_sheet_state(name, "hidden")
            logger.debug("Hidden sheet '%s'", name)
        if active is None or active in to_hide:
            workbook.set_active_sheet(first_remaining(visible, to_hide))

    def delete_sheets(self, names: Sequence[str]) -> None:
        """Delete the sheets named in *names*.

        Args:
            names: Sheet names to remove.

        """
        package, workbook = self._require_workbook()
        present = [name for name in dict.fromkeys(names) if workbook.sheet(name) is not None]
        for name in names:
            if workbook.sheet(name) is None:
                warn_sheet_not_found(name, "nothing was deleted")
        if not present:
            return
        visible = [entry.name for entry in workbook.sheets() if entry.visible]
        require_a_visible_sheet(visible, present, "Deleting")
        active = workbook.active_sheet_name()
        workbook.remove_sheets(present)
        for name in present:
            logger.debug("Deleted sheet '%s'", name)
        if active is not None and active not in present:
            workbook.set_active_sheet(active)
        else:
            workbook.set_active_sheet(first_remaining(visible, present))
        self._deleted_sheets.extend(present)
        warn_about_deleted_sheets(package, workbook, present)

    def set_active_sheet(self, name: str) -> None:
        """Set the active (displayed) sheet.

        Args:
            name: Name of the worksheet to activate.

        Raises:
            ValueError: If the sheet is hidden.

        """
        _, workbook = self._require_workbook()
        entry = workbook.sheet(name)
        if entry is None:
            warn_sheet_not_found(name, "the workbook keeps its active sheet")
            return
        if not entry.visible:
            msg = f"Sheet '{name}' is hidden and cannot be the active sheet."
            raise ValueError(msg)
        workbook.set_active_sheet(entry.name)
        logger.debug("Set active sheet to '%s'", name)

    def save(self, path: Path) -> None:
        """Write the workbook to *path*.

        Args:
            path: Destination path.

        Raises:
            ValueError: If *path* is not the template's format.

        """
        package, workbook = self._require_workbook()
        suffix = path.suffix.lower()
        if suffix != self._source_suffix:
            msg = (
                f"The fillxl engine cannot turn a '{self._source_suffix}' template "
                f"into a '{suffix}' file. Save to a '{self._source_suffix}' path, or "
                "use engine='win32com' to convert between formats."
            )
            raise ValueError(msg)
        # Formulas are written without cached results, and a calcChain that
        # disagrees with the workbook makes Excel refuse the file; Excel
        # rebuilds the chain on load anyway.
        workbook.set_full_calc_on_load()
        workbook.remove_calc_chain()
        logger.debug("Saving workbook to %s", path)
        package.save(path)

    def close(self) -> None:
        """Release the in-memory package."""
        self._package = None
        self._workbook = None

    def check_output(self, template: Path, output: Path, policy: OnLossy) -> None:
        """Report template features the fill did not write back.

        The engine rewrites only what it understands, so this is expected to
        stay silent; it is kept as an assertion that a surgical edit really
        was surgical.
        """
        check_lossy(
            template,
            output,
            policy,
            deleted_sheets=self._deleted_sheets,
            engine="fillxl",
        )
