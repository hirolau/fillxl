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
from fillxl._ooxml.styles import StylesPart
from fillxl._ooxml.workbook import WorkbookPart
from fillxl._ooxml.worksheet import SheetStore
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
        self._sheets: SheetStore | None = None
        self._styles: StylesPart | None = None
        self._source_suffix: str | None = None
        self._deleted_sheets: list[str] = []

    def _require_workbook(self) -> tuple[OoxmlPackage, WorkbookPart]:
        """Return the loaded package and workbook, or raise."""
        if self._package is None or self._workbook is None:
            msg = "No workbook loaded — call load() first"
            raise RuntimeError(msg)
        return self._package, self._workbook

    def _require_session(self) -> tuple[OoxmlPackage, WorkbookPart, SheetStore, StylesPart]:
        """Return the package, the workbook and the per-session part editors."""
        package, workbook = self._require_workbook()
        if self._sheets is None or self._styles is None:
            msg = "No workbook loaded — call load() first"
            raise RuntimeError(msg)
        return package, workbook, self._sheets, self._styles

    def _save_parts(self) -> None:
        """Write back the cached sheet editors, so ``package.text()`` is current."""
        _, _, sheets, styles = self._require_session()
        sheets.save_all()
        styles.save()

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
        # One editor per part and one style sheet per session: filling several
        # named cells must not parse the same sheet again for each of them.
        self._sheets = SheetStore(self._package)
        self._styles = StylesPart(self._package)

    def fill_named_cell(self, name: str, value: Any, config: FillXLConfig) -> None:  # noqa: ANN401
        """Write a scalar *value* to the named cell *name*.

        Args:
            name: Name of the named cell / defined name.
            value: Scalar value to write.
            config: Global fill configuration.

        """
        package, workbook, sheets, styles = self._require_session()
        fill_named_cell(
            package,
            workbook,
            name,
            value,
            escape_formulas=config.escape_formulas,
            sheets=sheets,
            styles=styles,
        )
        self._save_parts()

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
        package, workbook, sheets, styles = self._require_session()
        spec = fill_table(package, workbook, name, data, config, sheets=sheets, styles=styles)
        self._save_parts()
        return spec

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
        _, _, sheets, _ = self._require_session()
        for name in to_hide:
            workbook.set_sheet_state(name, "hidden")
            entry = workbook.sheet(name)
            if entry is not None:
                sheets.forget(entry.part)
            logger.debug("Hidden sheet '%s'", name)
        if active is None or active in to_hide:
            workbook.set_active_sheet(first_remaining(visible, to_hide))
            for entry in workbook.sheets():
                sheets.forget(entry.part)

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
        _, _, sheets, _ = self._require_session()
        present_parts = {
            entry.part for entry in workbook.sheets() if entry.name in set(present) and entry.part
        }
        workbook.remove_sheets(present)
        for part in present_parts:
            sheets.forget(part)
        for name in present:
            logger.debug("Deleted sheet '%s'", name)
        if active is not None and active not in present:
            workbook.set_active_sheet(active)
        else:
            workbook.set_active_sheet(first_remaining(visible, present))
        for entry in workbook.sheets():
            sheets.forget(entry.part)
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
        _, _, sheets, _ = self._require_session()
        for sheet_entry in workbook.sheets():
            sheets.forget(sheet_entry.part)
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
        self._sheets = None
        self._styles = None

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
