"""Openpyxl backend implementation for fillxl."""

from __future__ import annotations

import logging
import warnings
from typing import TYPE_CHECKING, Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from fillxl._cells import populate_named_cells
from fillxl._charts import (
    TableChangeSpecs,
    detect_unsupported_chart_sheets,
    update_chart_data_ranges,
)
from fillxl._tables import populate_table
from fillxl._validation import ColumnValidationStrategy, validate_columns

if TYPE_CHECKING:
    from pathlib import Path

    from fillxl._config import FillXLConfig

logger = logging.getLogger(__name__)

_CHARTEX_WARNING = (
    "Sheet(s) {sheets} contain chartEx charts (waterfall, treemap, "
    "sunburst, histogram, boxwhisker, funnel, etc.) that openpyxl "
    "cannot preserve. These charts will be lost when the workbook "
    "is saved. Use the win32com backend (engine='win32com') to "
    "preserve them."
)


class OpenpyxlBackend:
    """Backend implementation using openpyxl.

    This backend works cross-platform and does not require Microsoft Excel
    to be installed.  It uses the openpyxl library to read/write ``.xlsx``
    files directly.
    """

    def __init__(self) -> None:
        self._workbook: openpyxl.Workbook | None = None
        self._unsupported_chart_sheets: list[str] = []

    def _require_workbook(self) -> openpyxl.Workbook:
        """Return the loaded workbook or raise.

        Raises:
            RuntimeError: If no workbook has been loaded yet.

        """
        if self._workbook is None:
            msg = "No workbook loaded — call load() first"
            raise RuntimeError(msg)
        return self._workbook

    # ── Backend protocol ────────────────────────────────────────────────────

    def load(self, path: Path) -> None:
        """Open an Excel workbook from *path*.

        Args:
            path: Path to the workbook file.

        """
        logger.debug("Loading workbook from %s", path)

        ext = path.suffix.lower()
        if ext != ".xlsx":
            msg = (
                f"File type '{ext}' is not supported by the openpyxl backend. "
                f"Only .xlsx files are supported. "
                f"Use engine='win32com' to fill other formats (.xlsm, .xlsb, .xls, etc.) instead."
            )
            raise ValueError(msg)

        # Detect chartEx charts (waterfall, treemap, etc.) that openpyxl
        # cannot preserve — they will be silently dropped on save.
        self._unsupported_chart_sheets = detect_unsupported_chart_sheets(path)
        if self._unsupported_chart_sheets:
            msg = _CHARTEX_WARNING.format(
                sheets=", ".join(repr(s) for s in self._unsupported_chart_sheets),
            )
            logger.warning(msg)
            warnings.warn(msg, stacklevel=2)
        self._workbook = openpyxl.load_workbook(path)
        _fix_chart_axes(self._workbook)

    def fill_named_cell(self, name: str, value: Any, config: FillXLConfig) -> None:  # noqa: ANN401
        """Write a scalar *value* to the named cell *name*.

        Args:
            name: Name of the named cell / defined name.
            value: Scalar value to write.
            config: Global fill configuration.

        """
        workbook = self._require_workbook()
        populate_named_cells(
            workbook,
            {name: value},
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
            A :class:`TableChangeSpecs` describing the old and new table
            dimensions (so chart references can be updated).

        """
        workbook = self._require_workbook()

        # Column validation
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(data)
        _validate_table_columns(
            workbook=workbook,
            table_name=name,
            data_columns=adapter.columns,
            strategy=config.column_validation_strategy,
        )

        return populate_table(
            workbook,
            name,
            data,
            copy_style_in_tables=config.copy_table_styles,
            extend_conditional_formatting=config.extend_conditional_formatting,
            escape_formulas=config.escape_formulas,
        )

    def update_charts(self, spec: TableChangeSpecs) -> None:
        """Update chart data ranges after a table resize.

        Args:
            spec: The table dimension change to reflect in charts.

        """
        update_chart_data_ranges(spec, self._require_workbook())

    def set_pivot_refresh(self) -> None:
        """Mark all pivot caches for refresh on load."""
        for worksheet in self._require_workbook().worksheets:
            for pivot in worksheet._pivots:  # type: ignore[attr-defined, unused-ignore]
                pivot.cache.refreshOnLoad = True
        logger.debug("Set all pivot tables to refresh on load")

    def hide_sheets(self, prefix: str) -> None:
        """Hide worksheets whose names start with *prefix*.

        Args:
            prefix: Sheet name prefix (e.g. ``"_"``).

        """
        workbook = self._require_workbook()
        for sheet_name in workbook.sheetnames:
            if sheet_name.startswith(prefix):
                workbook[sheet_name].sheet_state = "hidden"
                logger.debug("Hidden worksheet '%s'", sheet_name)

    def delete_sheets(self, names: list[str]) -> None:
        """Delete the worksheets named in *names*.

        Args:
            names: Sheet names to remove.

        """
        workbook = self._require_workbook()
        for sheet_name in names:
            if sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
                if isinstance(ws, Worksheet):
                    workbook.remove(ws)
                    logger.debug("Deleted worksheet '%s'", sheet_name)
                else:
                    logger.warning("Sheet '%s' is not a regular worksheet, skipping", sheet_name)
            else:
                logger.warning("Sheet '%s' not found, skipping deletion", sheet_name)

    def set_active_sheet(self, name: str) -> None:
        """Set the active (displayed) worksheet.

        Args:
            name: Name of the worksheet to activate.

        """
        workbook = self._require_workbook()
        if name in workbook.sheetnames:
            ws = workbook[name]
            if isinstance(ws, Worksheet):
                workbook.active = workbook.index(ws)
            logger.debug("Set active worksheet to '%s'", name)
        else:
            logger.warning("Sheet '%s' not found, keeping default active sheet", name)

    def save(self, path: Path) -> None:
        """Save the workbook to *path*.

        Args:
            path: Destination path.

        """
        workbook = self._require_workbook()
        if self._unsupported_chart_sheets:
            msg = _CHARTEX_WARNING.format(
                sheets=", ".join(repr(s) for s in self._unsupported_chart_sheets),
            )
            logger.warning(msg)
            warnings.warn(msg, stacklevel=2)
        _fix_chart_axes(workbook)
        logger.debug("Saving workbook to %s", path)
        workbook.save(path)

    def close(self) -> None:
        """Close the workbook and release resources."""
        if self._workbook is not None:
            self._workbook.close()
            self._workbook = None


def _validate_table_columns(
    workbook: openpyxl.Workbook,
    table_name: str,
    data_columns: list[str],
    strategy: ColumnValidationStrategy,
) -> None:
    """Validate columns of incoming data against the Excel table definition.

    Args:
        workbook: The openpyxl Workbook containing the table.
        table_name: Name of the table to validate against.
        data_columns: Column names from the incoming data.
        strategy: Validation strategy (see :func:`fillxl._validation.validate_columns`).

    Raises:
        ValueError: If the strategy requires exact match and columns don't align.

    """
    from fillxl._utils import find_worksheet_and_table

    _, table = find_worksheet_and_table(workbook, table_name)
    table_columns = [col.name for col in table.tableColumns]
    formula_columns = [col.name for col in table.tableColumns if col.calculatedColumnFormula]

    validate_columns(
        table_columns=table_columns,
        data_columns=data_columns,
        strategy=strategy,
        formula_columns=formula_columns,
    )


def _fix_chart_axes(workbook: openpyxl.Workbook) -> None:
    """Fix chart axis definitions that openpyxl fails to load correctly.

    openpyxl has a bug where it creates default axis objects with wrong
    ``axPos`` values for certain chart types (e.g., :class:`~openpyxl.chart.AreaChart`,
    :class:`~openpyxl.chart.LineChart`) instead of parsing the actual axis
    definitions from the XML.  This results in the category axis being
    placed on the left instead of at the bottom, and missing critical
    elements like ``delete``, ``crosses``, ``auto``, ``lblAlgn``, and
    ``noMultiLvlLbl``, causing x-axis labels to disappear in Excel.

    The fix inspects each chart and ensures:
    * The category axis (``TextAxis``) is positioned at the bottom (``axPos='b'``)
      for vertical chart types, or on the left (``axPos='l'``) for horizontal
      bar charts.
    * The value axis (``NumericAxis``) is positioned on the left (``axPos='l'``)
      for vertical charts, or at the bottom (``axPos='b'``) for horizontal bars.
    * ``delete`` is set to ``False`` (axis visible).
    * ``tickLblPos`` is set to ``'nextTo'`` if not already set.
    * ``crosses`` is set to ``'autoZero'`` if not already set.
    * TextAxis-specific: ``auto``, ``lblAlgn``, ``noMultiLvlLbl`` are set.
    * NumericAxis-specific: ``crossBetween`` is set to ``'midCat'``.

    Args:
        workbook: The openpyxl Workbook to fix.

    """
    from openpyxl.chart import BarChart
    from openpyxl.chart.axis import NumericAxis, TextAxis
    from openpyxl.chart.data_source import NumFmt

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        for chart in getattr(ws, "_charts", []):
            axes = chart._axes
            if not axes:
                continue

            is_horizontal_bar = isinstance(chart, BarChart) and chart.barDir == "bar"

            for ax in axes.values():
                # Common fixes for both axis types
                if ax.delete is None:
                    ax.delete = False
                if ax.tickLblPos is None:
                    ax.tickLblPos = "nextTo"
                if ax.crosses is None:
                    ax.crosses = "autoZero"

                if isinstance(ax, TextAxis):
                    # Category axis
                    ax.axPos = "l" if is_horizontal_bar else "b"
                    if ax.numFmt is None:
                        ax.numFmt = NumFmt(formatCode="General", sourceLinked=True)
                    if ax.auto is None:
                        ax.auto = True
                    if ax.lblAlgn is None:
                        ax.lblAlgn = "ctr"
                    if ax.noMultiLvlLbl is None:
                        ax.noMultiLvlLbl = False
                elif isinstance(ax, NumericAxis):
                    # Value axis
                    ax.axPos = "b" if is_horizontal_bar else "l"
                    if ax.crossBetween is None:
                        ax.crossBetween = "midCat"
