"""Win32com (pywin32) backend implementation for fillxl.

Drives real Microsoft Excel via COM automation (pywin32), providing
capabilities that openpyxl cannot: rendered conditional formatting colors,
real formula evaluation, display format application, and true pivot table
refresh.

Only works on **Windows** with Microsoft Excel installed.
"""

from __future__ import annotations

import logging
import sys
import threading
import time
from collections.abc import Mapping, Sequence
from typing import TYPE_CHECKING, Any

from openpyxl.worksheet.cell_range import CellRange

from fillxl._adapters import TabularAdapter
from fillxl._cells import named_cell_target
from fillxl._charts import TableChangeSpecs, split_formula_arguments, update_reference_text
from fillxl._formats import needs_temporal_format, temporal_format, temporal_format_runs
from fillxl._formulas import (
    BodyState,
    ColumnFormulaEvidence,
    formula_definitions,
    resolve_formula,
    warn_about_ignored_formula_data,
)
from fillxl._normalize import (
    check_cell_value,
    com_value_from_normalized,
    prepare_cell_value,
)
from fillxl._plan import TableGeometry, plan_table_fill
from fillxl._sheets import first_remaining, require_a_visible_sheet
from fillxl._utils import col_to_excel
from fillxl._validation import validate_columns

if TYPE_CHECKING:
    from collections.abc import Iterator
    from pathlib import Path

    from fillxl._config import FillXLConfig
    from fillxl._lossy import OnLossy

logger = logging.getLogger(__name__)

_is_windows = sys.platform == "win32"

# COM constants (fallback if early binding is unavailable)
_XL_SHEET_HIDDEN = 0
_XL_SHEET_VERY_HIDDEN = 2
_XL_FILL_FORMATS = 3

_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
_RPC_E_CHANGED_MODE = -2147417850
_XL_CALCULATION_MANUAL = -4135

_OOXML_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})

# SaveAs file formats, keyed by output suffix.  Saving an .xlsm template with
# the .xlsx format silently strips its macros, so the format is inferred from
# the requested output path instead of being hard-coded.
_FILE_FORMATS: dict[str, int] = {
    ".xlsx": 51,  # xlOpenXMLWorkbook
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12
    ".xls": 56,  # xlExcel8
}


class Win32ComBackend:
    """Backend implementation using pywin32 (win32com) to drive real Excel.

    Only works on Windows with Microsoft Excel installed.  Provides lazy
    Excel application initialisation — Excel is launched on the first call
    to :meth:`load` rather than in ``__init__``.

    Attributes:
        _app: The Excel ``Application`` COM object (``None`` until first use).
        _workbook: The currently open ``Workbook`` COM object (``None``
            until :meth:`load` is called).

    """

    def __init__(self) -> None:
        self._app: Any = None
        self._workbook: Any = None
        self._load_path: Path | None = None
        # Calculated-column definitions per table; None = file format cannot be read.
        self._definitions: dict[str, dict[str, str]] | None = None
        self._definitions_loaded = False
        self._com_initialized = False
        self._com_thread: int | None = None
        # Excel's calculation mode before this fill; restored before SaveAs.
        self._previous_calculation: int | None = None

    # ── Internal helpers ───────────────────────────────────────────────────

    def _ensure_platform(self) -> None:
        """Raise a clear error if not on Windows."""
        if not _is_windows:
            msg = (
                "The win32com backend requires Windows and Microsoft Excel. "
                "Use the 'openpyxl' engine on non-Windows platforms."
            )
            raise RuntimeError(msg)

    def _ensure_app(self) -> None:
        """Start a hidden Excel that cannot run macros, fire events or update links."""
        if self._app is not None:
            return
        self._ensure_platform()
        import pythoncom
        import win32com.client

        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error as exc:
            # The thread already uses another apartment model; COM is usable as is.
            # getattr keeps this readable where pywin32 is not installed (Linux CI).
            if getattr(exc, "hresult", None) != _RPC_E_CHANGED_MODE:
                raise
        else:
            self._com_initialized = True
            self._com_thread = threading.get_ident()

        # DispatchEx always creates a new, isolated Excel instance.
        self._app = win32com.client.DispatchEx("Excel.Application")
        # Security settings must hold before any workbook opens; failures propagate.
        self._app.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        self._app.EnableEvents = False
        self._app.AskToUpdateLinks = False
        for attr, val in [("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False)]:
            try:
                setattr(self._app, attr, val)
            except AttributeError:
                logger.warning("Could not set Excel.Application.%s", attr)
        logger.debug("Excel Application started (hidden, macros disabled)")

    def _ensure_workbook(self) -> None:
        """Raise :class:`RuntimeError` if no workbook has been loaded."""
        if self._workbook is None:
            msg = "No workbook loaded — call load() first"
            raise RuntimeError(msg)

    def _table_formula_definitions(self, table_name: str) -> dict[str, str] | None:
        """Definitions of *table_name*, or ``None`` when the template format hides them."""
        if not self._definitions_loaded:
            self._definitions = _read_table_formula_definitions(self._load_path)
            self._definitions_loaded = True
        if self._definitions is None:
            return None
        return self._definitions.get(table_name, {})

    # ── Backend protocol ───────────────────────────────────────────────────

    def load(self, path: Path) -> None:
        """Open an Excel workbook from *path*.

        Args:
            path: Path to the workbook file.

        """
        self._ensure_app()
        t0 = time.perf_counter()
        abs_path = path.resolve()
        logger.debug("Loading workbook from %s", abs_path)
        self._load_path = abs_path
        self._definitions = None
        self._definitions_loaded = False
        self._workbook = self._app.Workbooks.Open(
            str(abs_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        # Filling recalculates constantly in automatic mode; save() restores this.
        if self._previous_calculation is None:
            self._previous_calculation = self._app.Calculation
        self._app.Calculation = _XL_CALCULATION_MANUAL
        logger.debug("[win32com] load took %.3fs", time.perf_counter() - t0)

    def fill_named_cell(self, name: str, value: Any, config: FillXLConfig) -> None:  # noqa: ANN401
        """Write a scalar *value* to the named cell *name*.

        Args:
            name: Name of the named cell / defined name.
            value: Scalar value to write.
            config: Global fill configuration.

        Raises:
            KeyError: If *name* is not a defined name in the workbook.

        """
        self._ensure_workbook()
        try:
            name_obj = self._workbook.Names(name)
        except Exception as exc:
            msg = f"Named cell '{name}' not found in workbook"
            raise KeyError(msg) from exc

        target = named_cell_target(name, _com_name_destinations(str(name_obj.RefersTo)))
        if target is None:
            return
        sheet_name, address = target
        normalized = prepare_cell_value(value)
        check_cell_value(normalized, name=name)
        prepared = com_value_from_normalized(normalized, escape_formulas=config.escape_formulas)
        try:
            target_range = self._workbook.Worksheets(sheet_name).Range(address)
            # Read the format first: Excel changes a General cell when a date arrives.
            existing_format = self._existing_number_format(target_range)
            target_range.Value = prepared
            fmt = temporal_format(normalized)
            if fmt is not None and needs_temporal_format(existing_format):
                set_number_format(target_range, fmt)
        except Exception as exc:
            msg = f"Could not write named cell '{name}' ({sheet_name}!{address}): {exc}"
            raise RuntimeError(msg) from exc

    def _existing_number_format(self, target: Any) -> Any:  # noqa: ANN401
        """The range's format, with Excel's name for General spelled ``General``.

        Excel names the General format in the language of its object model
        (``Standard`` in Swedish, for instance), and openpyxl cannot tell such
        a name from a date format.
        """
        return _number_format_of(target.NumberFormat, _general_format_name(self._app))

    def _find_list_object(self, name: str) -> tuple[Any, Any]:
        """The ListObject called exactly *name* and its worksheet."""
        for sheet in self._workbook.Worksheets:
            try:
                candidate = sheet.ListObjects(name)
            except Exception:  # not on this sheet
                continue
            # ListObjects() ignores case; openpyxl does not, and both must agree.
            if candidate is not None and str(candidate.Name) == name:
                return candidate, sheet
        msg = f"Table '{name}' not found in workbook"
        raise KeyError(msg)

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

        Raises:
            KeyError: If the table is not found in the workbook.

        """
        self._ensure_workbook()
        list_obj, worksheet = self._find_list_object(name)

        adapter = TabularAdapter(data)
        data_columns = adapter.columns
        rows = adapter.rows

        t_meta = time.perf_counter()
        geometry = _com_table_geometry(list_obj)
        column_count = geometry.last_col - geometry.first_col + 1
        # ListColumn.Name also works for tables shown without a header row.
        table_columns = [str(list_obj.ListColumns(i).Name) for i in range(1, column_count + 1)]
        logger.debug("[win32com] table metadata took %.3fs", time.perf_counter() - t_meta)

        # Detection reads the template body, so it runs before anything changes.
        formulas = _com_formula_columns(
            self._app, list_obj, table_columns, self._table_formula_definitions(name)
        )
        validate_columns(
            table_columns=table_columns,
            data_columns=data_columns,
            strategy=config.column_validation_strategy,
            formula_columns=list(formulas),
            table_name=name,
        )
        warn_about_ignored_formula_data(name, data_columns, rows, formulas)

        plan = plan_table_fill(geometry, len(rows))
        first_letter = col_to_excel(geometry.first_col)
        last_letter = col_to_excel(geometry.last_col)

        if plan.overflow_rows is not None:
            _check_overflow_com(
                worksheet,
                table_name=name,
                start_col=geometry.first_col,
                start_col_letter=first_letter,
                end_col_letter=last_letter,
                first_row=plan.overflow_rows[0],
                last_row=plan.overflow_rows[1],
                num_data_rows=len(rows),
            )

        if not rows:
            _clear_body_com(list_obj.DataBodyRange, headerless=geometry.header_rows == 0)

        # Resize first: Excel fills calculated columns into new rows itself.
        # worksheet.Range, not Application.Range: the latter means the ACTIVE sheet.
        t_resize = time.perf_counter()
        if plan.resized:
            resize_last_row = _resize_target_row(
                plan.new_last_row, geometry.last_row, geometry.totals_rows
            )
            list_obj.Resize(
                worksheet.Range(f"{first_letter}{geometry.top_row}:{last_letter}{resize_last_row}")
            )
        if plan.vacated_rows is not None:
            first_row, last_row = plan.vacated_rows
            worksheet.Range(f"{first_letter}{first_row}:{last_letter}{last_row}").ClearContents()
        logger.debug("[win32com] resize took %.3fs", time.perf_counter() - t_resize)

        if rows:
            _write_rows_com(
                list_obj,
                table_name=name,
                table_columns=table_columns,
                data_columns=data_columns,
                rows=rows,
                formulas=formulas,
                headerless=geometry.header_rows == 0,
                config=config,
                general_format=_general_format_name(self._app),
            )

        if config.extend_conditional_formatting:
            _extend_conditional_formatting_com(
                self._app,
                worksheet,
                start_col=geometry.first_col,
                end_col=geometry.last_col,
                first_data_row=plan.first_data_row,
                last_data_row=plan.new_body_last_row,
            )

        return plan.change_spec(str(worksheet.Name))

    def update_charts(self, spec: TableChangeSpecs) -> None:
        """Point chart series that show the resized table at its new rows.

        Covers charts embedded in worksheets and charts on their own chart sheet.

        Args:
            spec: The table dimension change to reflect in charts.

        """
        self._ensure_workbook()
        for chart in self._iter_charts():
            for series in chart.SeriesCollection():
                try:
                    formula = str(series.Formula)
                except Exception:  # e.g. pivot chart series have no formula
                    continue
                updated = _update_series_formula(formula, spec)
                if updated is None:
                    continue
                try:
                    series.Formula = updated
                except Exception as exc:
                    msg = f"Could not update chart series {formula!r} to {updated!r}: {exc}"
                    raise RuntimeError(msg) from exc
                logger.debug("Updated chart series formula: %s -> %s", formula, updated)

    def _iter_charts(self) -> Iterator[Any]:
        for sheet in self._workbook.Worksheets:
            for chart_object in sheet.ChartObjects():
                yield chart_object.Chart
        # A chart sheet is itself a Chart; it holds no ChartObjects.
        yield from self._workbook.Charts

    def set_pivot_refresh(self) -> None:
        """Mark all pivot caches for refresh, and refresh them now.

        The openpyxl backend can only mark caches for refresh on open; this
        backend does the same *and* recalculates the caches, so the saved
        workbook is up to date whichever engine produced it.
        """
        self._ensure_workbook()
        # Calculation is manual while filling; pivots must read up-to-date formulas.
        self._app.Calculate()
        for pcache in self._workbook.PivotCaches():
            try:
                pcache.Refresh()
            except Exception:
                logger.exception("Failed to refresh pivot cache")
            # The COM property is ``RefreshOnFileOpen``; it is what the file
            # format stores as ``refreshOnLoad``.  (``RefreshOnLoad`` is the
            # XML attribute name, not a COM property.)
            try:
                pcache.RefreshOnFileOpen = True
            except Exception:
                logger.exception("Failed to mark pivot cache for refresh on open")
        logger.debug("Refreshed all pivot caches")

    def _sheet_names(self) -> list[str]:
        return [str(sheet.Name) for sheet in self._workbook.Sheets]

    def _visible_sheet_names(self) -> list[str]:
        return [str(sheet.Name) for sheet in self._workbook.Sheets if _com_is_visible(sheet)]

    def hide_sheets(self, prefix: str) -> None:
        """Hide visible sheets whose names start with *prefix*."""
        self._ensure_workbook()
        visible = self._visible_sheet_names()
        to_hide = [name for name in visible if name.startswith(prefix)]
        if not to_hide:
            return
        require_a_visible_sheet(visible, to_hide, "Hiding")
        if str(self._workbook.ActiveSheet.Name) in to_hide:
            self._workbook.Sheets(first_remaining(visible, to_hide)).Activate()
        for name in to_hide:
            try:
                self._workbook.Sheets(name).Visible = _XL_SHEET_HIDDEN
            except Exception as exc:
                msg = f"Could not hide sheet '{name}': {exc}"
                raise RuntimeError(msg) from exc
            logger.debug("Hidden sheet '%s'", name)

    def delete_sheets(self, names: Sequence[str]) -> None:
        """Delete the sheets named in *names*."""
        self._ensure_workbook()
        existing = set(self._sheet_names())
        present = [name for name in dict.fromkeys(names) if name in existing]
        for name in names:
            if name not in existing:
                logger.warning("Sheet '%s' not found, skipping deletion", name)
        if not present:
            return
        visible = self._visible_sheet_names()
        require_a_visible_sheet(visible, present, "Deleting")
        if str(self._workbook.ActiveSheet.Name) in present:
            self._workbook.Sheets(first_remaining(visible, present)).Activate()
        for name in present:
            try:
                self._workbook.Sheets(name).Delete()
            except Exception as exc:
                msg = f"Could not delete sheet '{name}': {exc}"
                raise RuntimeError(msg) from exc
            logger.debug("Deleted sheet '%s'", name)

    def set_active_sheet(self, name: str) -> None:
        """Set the active (displayed) sheet."""
        self._ensure_workbook()
        if name not in self._sheet_names():
            logger.warning("Sheet '%s' not found, keeping default active sheet", name)
            return
        sheet = self._workbook.Sheets(name)
        if not _com_is_visible(sheet):
            msg = f"Sheet '{name}' is hidden and cannot be the active sheet."
            raise ValueError(msg)
        sheet.Activate()
        logger.debug("Set active sheet to '%s'", name)

    def save(self, path: Path) -> None:
        """Save the workbook to *path*.

        The file format is inferred from the output suffix, so filling an
        ``.xlsm`` template keeps its macros instead of silently saving the
        workbook as macro-free ``.xlsx``.

        Args:
            path: Destination path.

        Raises:
            ValueError: If the output suffix is not a supported Excel format.

        """
        self._ensure_workbook()
        t0 = time.perf_counter()
        suffix = path.suffix.lower()
        file_format = _FILE_FORMATS.get(suffix)
        if file_format is None:
            supported = ", ".join(sorted(_FILE_FORMATS))
            msg = (
                f"Cannot save '{suffix}' files. Supported output formats "
                f"for the win32com backend: {supported}."
            )
            raise ValueError(msg)

        # Resolve to absolute path — COM/Excel resolves relative paths
        # from the Excel process working directory, not the Python CWD.
        abs_path = path.resolve()
        abs_path.parent.mkdir(parents=True, exist_ok=True)
        # Never save stale results, even if the workbook's own mode is manual.
        self._app.Calculate()
        if self._previous_calculation is not None:
            self._app.Calculation = self._previous_calculation
            self._previous_calculation = None
        logger.debug("Saving workbook to %s (format %d)", abs_path, file_format)
        self._workbook.SaveAs(str(abs_path), FileFormat=file_format)
        logger.debug("[win32com] save took %.3fs", time.perf_counter() - t0)

    def close(self) -> None:
        """Close the workbook, quit Excel and release COM on this thread."""
        try:
            if self._workbook is not None:
                try:
                    self._workbook.Close(SaveChanges=False)
                except Exception:
                    logger.exception("Error closing workbook")
                finally:
                    self._workbook = None
        finally:
            try:
                if self._app is not None:
                    try:
                        self._app.Quit()
                    except Exception:
                        logger.exception("Error during Excel quit")
                    finally:
                        self._app = None
                        logger.debug("Excel Application quit")
            finally:
                self._release_com()

    def check_output(self, template: Path, output: Path, policy: OnLossy) -> None:
        """Excel writes back every feature it read, so there is nothing to check."""

    def _release_com(self) -> None:
        if not self._com_initialized:
            return
        self._com_initialized = False
        if threading.get_ident() != self._com_thread:
            logger.warning("close() ran on another thread than load(); COM was not released")
            return
        import pythoncom

        pythoncom.CoUninitialize()


# ── Standalone helpers ──────────────────────────────────────────────────────


def _com_is_visible(sheet: Any) -> bool:  # noqa: ANN401
    return sheet.Visible not in (_XL_SHEET_HIDDEN, _XL_SHEET_VERY_HIDDEN)


_EN_US_LCID = 1033
_XL_GENERAL_FORMAT_NAME = 26  # XlApplicationInternational.xlGeneralFormatName


def _general_format_name(app: Any) -> str:  # noqa: ANN401
    """Excel's own name for the General format (``General``, ``Standard``, …).

    ``Application.International`` is a method on a typed Excel, but pywin32
    exposes it as a tuple ordered by the ``XlApplicationInternational``
    constants.
    """
    try:
        international = app.International
        if callable(international):
            return str(international(_XL_GENERAL_FORMAT_NAME))
        return str(international[_XL_GENERAL_FORMAT_NAME - 1])
    except Exception:  # an unknown name simply never matches
        return "General"


def _number_format_of(existing: Any, general_format: str) -> Any:  # noqa: ANN401
    """*existing* with Excel's own name for General spelled ``General``."""
    if isinstance(existing, str) and existing == general_format:
        return "General"
    return existing


def set_number_format(target: Any, format_code: str) -> None:  # noqa: ANN401
    """Apply an English *format_code* to a COM range, whatever Excel's language.

    Excel parses a format code in the language of the installed product, so a
    Swedish Excel reads ``yyyy`` as four literal ``y`` characters.  The
    property is set under an en-US locale so the code is read as written; a
    plain assignment is the fallback when the COM object has no low-level
    property access.
    """
    try:
        ole = target._oleobj_
        dispid = ole.GetIDsOfNames("NumberFormat")
    except Exception:  # e.g. a wrapper without COM internals
        target.NumberFormat = format_code
        return
    import pythoncom

    ole.Invoke(dispid, _EN_US_LCID, pythoncom.DISPATCH_PROPERTYPUT, 0, format_code)


def _com_has_totals_row(list_obj: Any) -> bool:  # noqa: ANN401
    """Whether a COM ``ListObject`` is configured with a totals row."""
    return bool(getattr(list_obj, "ShowTotals", False))


def _resize_target_row(
    new_table_last_row: int,
    old_table_last_row: int,
    totals_rows: int,
) -> int:
    """Last row to pass to ``ListObject.Resize`` for a totals-row table.

    Excel interprets the resize range as the final table range while a table
    grows, but as the range *above* the totals row while it shrinks — the
    totals row is then appended one row below it.  Adjusting for that makes
    both directions produce *new_table_last_row* as the table's last row.

    Examples:
        >>> _resize_target_row(5, 3, 1)
        5
        >>> _resize_target_row(3, 5, 1)
        2
        >>> _resize_target_row(3, 5, 0)
        3

    """
    if new_table_last_row > old_table_last_row:
        return new_table_last_row
    return new_table_last_row - totals_rows


def _prepare_formula_for_com(formula: str) -> str:
    """Normalize a calculated-column formula for Excel COM assignment.

    Preserves newlines and indentation so carefully formatted multi-line
    formulas (e.g. ``LET`` / ``LAMBDA`` blocks) keep their layout.

    Excel ``Formula2`` accepts LF (``\\n``) line breaks but rejects bare CR
    (``\\r``) / CRLF, so those are normalized to LF.  openpyxl XML prefixes
    ``_xlfn.`` / ``_xlpm.`` are stripped because COM does not understand them.

    Args:
        formula: Formula string, typically with a leading ``=``.

    Returns:
        Formula string ready for ``Range.Formula2`` / ``Range.Formula``.

    Examples:
        >>> _prepare_formula_for_com("=LET(\\n    x, 1,\\n    x+1\\n)")
        '=LET(\\n    x, 1,\\n    x+1\\n)'
        >>> _prepare_formula_for_com("=_xlfn.LET(_xlpm.x,1,_xlpm.x+1)")
        '=LET(x,1,x+1)'
        >>> _prepare_formula_for_com("=A1\\r\\n+B1")
        '=A1\\n+B1'

    """
    return (
        formula.replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("_xlfn.", "")
        .replace("_xlpm.", "")
    )


def _check_overflow_com(
    worksheet: Any,  # noqa: ANN401
    *,
    table_name: str,
    start_col: int,
    start_col_letter: str,
    end_col_letter: str,
    first_row: int,
    last_row: int,
    num_data_rows: int,
) -> None:
    """Refuse to grow a table over non-empty cells, or when the cells cannot be read."""
    range_str = f"{start_col_letter}{first_row}:{end_col_letter}{last_row}"
    try:
        values = worksheet.Range(range_str).Value
    except Exception as exc:
        msg = (
            f"Table '{table_name}' cannot grow: the cells below it ({range_str}) "
            "could not be read, so fillxl cannot confirm that they are empty."
        )
        raise RuntimeError(msg) from exc
    if values is None:
        return
    grid = values if isinstance(values, tuple) else ((values,),)
    for row_offset, row_values in enumerate(grid):
        for col_offset, cell in enumerate(row_values):
            if cell is not None:
                cell_ref = f"{col_to_excel(start_col + col_offset)}{first_row + row_offset}"
                msg = (
                    f"Table '{table_name}' cannot expand to {num_data_rows} rows: "
                    f"cell {cell_ref} already contains data ({cell!r}). "
                    "Move or remove the content below the table first."
                )
                raise ValueError(msg)


def _com_name_destinations(refers_to: str) -> list[tuple[str, str]]:
    """Parse ``Name.RefersTo`` exactly the way the openpyxl backend parses names."""
    from openpyxl.workbook.defined_name import DefinedName

    probe = DefinedName("fillxl_probe", attr_text=refers_to.removeprefix("="))
    return list(probe.destinations)


def _read_table_formula_definitions(
    path: Path | None,
) -> dict[str, dict[str, str]] | None:
    """Every table's calculated-column definition texts, or ``None`` if unreadable."""
    if path is None or path.suffix.lower() not in _OOXML_SUFFIXES:
        return None
    import openpyxl

    try:
        workbook = openpyxl.load_workbook(path)
    except Exception:
        logger.warning(
            "Could not read table definitions from %s; calculated columns are "
            "detected from cell formulas only",
            path,
            exc_info=True,
        )
        return None
    try:
        return {
            str(table.name): formula_definitions(table)
            for worksheet in workbook.worksheets
            for table in worksheet.tables.values()
        }
    finally:
        workbook.close()


def _read_formula(cell: Any) -> str:  # noqa: ANN401
    try:
        return str(cell.Formula2)
    except AttributeError:
        return str(cell.Formula)


def _com_body_state(app: Any, body: Any) -> tuple[BodyState, str | None]:  # noqa: ANN401
    """Classify a column's body range; ``HasFormula`` is ``None`` for a mixed range."""
    if body is None:
        return "empty", None
    has_formula = body.HasFormula
    if has_formula is False:
        is_empty = app.WorksheetFunction.CountA(body) == 0
        return ("empty" if is_empty else "values"), None
    first = body.Cells(1, 1)
    first_formula = _read_formula(first) if first.HasFormula else None
    return ("formula" if has_formula is True else "mixed"), first_formula


def _com_formula_columns(
    app: Any,  # noqa: ANN401
    list_obj: Any,  # noqa: ANN401
    table_columns: Sequence[str],
    definitions: Mapping[str, str] | None,
) -> dict[str, str]:
    """Calculated columns of a ListObject, mapped to the formula Excel must keep."""
    formulas: dict[str, str] = {}
    for index, name in enumerate(table_columns, start=1):
        definition = definitions.get(name) if definitions is not None else None
        if definitions is not None and definition is None:
            continue
        state, first_formula = _com_body_state(app, list_obj.ListColumns(index).DataBodyRange)
        evidence = ColumnFormulaEvidence(
            name, definition, definitions is not None, state, first_formula
        )
        formula = resolve_formula(evidence)
        if formula is not None:
            formulas[name] = formula
    return formulas


def _com_table_geometry(list_obj: Any) -> TableGeometry:  # noqa: ANN401
    """Where a ListObject sits; the same numbers openpyxl reads from the table ref."""
    table_range = list_obj.Range
    top_row = int(table_range.Row)
    first_col = int(table_range.Column)
    return TableGeometry(
        top_row=top_row,
        last_row=top_row + int(table_range.Rows.Count) - 1,
        first_col=first_col,
        last_col=first_col + int(table_range.Columns.Count) - 1,
        header_rows=1 if list_obj.ShowHeaders else 0,
        totals_rows=1 if _com_has_totals_row(list_obj) else 0,
    )


def _clear_body_com(body: Any, *, headerless: bool) -> None:  # noqa: ANN401
    """Clear a table body's values.

    Clearing a headerless table's whole body in one call drops the ListObject.
    Sub-ranges are safe, so such bodies are cleared row by row (cell by cell
    for a single row).
    """
    if body is None:
        return
    if not headerless:
        body.ClearContents()
        return
    row_count = body.Rows.Count
    if row_count == 1:
        for col_index in range(1, body.Columns.Count + 1):
            body.Cells(1, col_index).ClearContents()
    else:
        for row_index in range(1, row_count + 1):
            body.Rows(row_index).ClearContents()


def _write_rows_com(
    list_obj: Any,  # noqa: ANN401
    *,
    table_name: str,
    table_columns: Sequence[str],
    data_columns: Sequence[str],
    rows: Sequence[Sequence[Any]],
    formulas: Mapping[str, str],
    headerless: bool,
    config: FillXLConfig,
    general_format: str,
) -> None:
    """Write *rows* into the (already resized) table body."""
    body = list_obj.DataBodyRange
    # Formats go first, so a template date format is in place before values arrive.
    if config.copy_table_styles and len(rows) > 1:
        body.Rows(1).AutoFill(Destination=body, Type=_XL_FILL_FORMATS)

    data_col_index = {column: index for index, column in enumerate(data_columns)}
    writable = [index for index, column in enumerate(table_columns) if column not in formulas]
    sources = {index: data_col_index.get(table_columns[index]) for index in writable}

    warned: set[tuple[str | None, str | None]] = set()
    matrix: list[list[Any]] = []
    normalized_columns: dict[int, list[Any]] = {index: [] for index in writable}
    for row_number, row in enumerate(rows, start=1):
        row_values: list[Any] = [None] * len(table_columns)
        for col_idx in writable:
            source = sources[col_idx]
            normalized = None
            if source is not None:
                normalized = prepare_cell_value(row[source])
                check_cell_value(
                    normalized,
                    table=table_name,
                    row=row_number,
                    column=table_columns[col_idx],
                    warned=warned,
                )
                row_values[col_idx] = com_value_from_normalized(
                    normalized, escape_formulas=config.escape_formulas
                )
            normalized_columns[col_idx].append(normalized)
        matrix.append(row_values)

    # Read formats BEFORE writing: Excel reformats General cells that receive dates.
    date_formats = _temporal_format_plan(body, normalized_columns, general_format)
    _write_body_blocks(body, matrix, writable, headerless=headerless)
    _apply_temporal_formats(body, date_formats)
    _ensure_formula_columns(list_obj, table_columns, formulas, table_name)


def _temporal_format_plan(
    body: Any,  # noqa: ANN401
    columns: Mapping[int, Sequence[Any]],
    general_format: str,
) -> list[tuple[int, int, int, str]]:
    """``(column, first_row, last_row, format)`` cell runs (0-based) that need a date format.

    This applies openpyxl's rule: a date written to a cell without a date
    format gets the default date format.
    """
    plan: list[tuple[int, int, int, str]] = []
    for col_idx, values in columns.items():
        runs = temporal_format_runs(values)
        if not runs:
            continue
        existing = _number_format_of(body.Columns(col_idx + 1).NumberFormat, general_format)
        if existing is None:
            # The column mixes formats: decide cell by cell, like openpyxl does.
            for first, last, fmt in runs:
                for offset in range(first, last + 1):
                    cell_format = _number_format_of(
                        body.Cells(offset + 1, col_idx + 1).NumberFormat, general_format
                    )
                    if values[offset] is not None and needs_temporal_format(cell_format):
                        plan.append((col_idx, offset, offset, fmt))
            continue
        if needs_temporal_format(existing):
            plan.extend((col_idx, first, last, fmt) for first, last, fmt in runs)
    return plan


def _apply_temporal_formats(
    body: Any,  # noqa: ANN401
    plan: Sequence[tuple[int, int, int, str]],
) -> None:
    sheet = body.Worksheet
    for col_idx, first, last, fmt in plan:
        set_number_format(
            sheet.Range(body.Cells(first + 1, col_idx + 1), body.Cells(last + 1, col_idx + 1)),
            fmt,
        )


def _column_blocks(indices: Sequence[int]) -> list[tuple[int, int]]:
    """Group column indices into inclusive runs of adjacent columns.

    Examples:
        >>> _column_blocks([0, 1, 2, 4, 6, 7])
        [(0, 2), (4, 4), (6, 7)]
        >>> _column_blocks([])
        []

    """
    blocks: list[tuple[int, int]] = []
    for index in sorted(indices):
        if blocks and index == blocks[-1][1] + 1:
            blocks[-1] = (blocks[-1][0], index)
        else:
            blocks.append((index, index))
    return blocks


def _write_body_blocks(
    body: Any,  # noqa: ANN401
    matrix: list[list[Any]],
    writable: Sequence[int],
    *,
    headerless: bool,
) -> None:
    """Write *matrix* into *body*, touching only the *writable* (0-based) columns.

    Excel drops a headerless ListObject when its whole body is assigned at once,
    so headerless tables are written row by row (cell by cell for one row).
    """
    if not matrix:
        return
    sheet = body.Worksheet
    row_count = len(matrix)
    for start, end in _column_blocks(writable):
        block = [row[start : end + 1] for row in matrix]
        if not headerless:
            target = sheet.Range(body.Cells(1, start + 1), body.Cells(row_count, end + 1))
            target.Value = block
        elif row_count == 1:
            for offset, value in enumerate(block[0]):
                body.Cells(1, start + 1 + offset).Value = value
        else:
            for row_number, row_values in enumerate(block, start=1):
                target = sheet.Range(
                    body.Cells(row_number, start + 1), body.Cells(row_number, end + 1)
                )
                target.Value = [row_values]


def _ensure_formula_columns(
    list_obj: Any,  # noqa: ANN401
    table_columns: Sequence[str],
    formulas: Mapping[str, str],
    table_name: str,
) -> None:
    """Write the formula into any calculated column that Excel did not fill itself."""
    for index, name in enumerate(table_columns, start=1):
        formula = formulas.get(name)
        if formula is None:
            continue
        body = list_obj.ListColumns(index).DataBodyRange
        if body is None or body.HasFormula is True:
            continue
        prepared = _prepare_formula_for_com(formula)
        try:
            try:
                body.Formula2 = prepared
            except AttributeError:
                body.Formula = prepared
        except Exception as exc:
            msg = (
                f"Table '{table_name}': could not write the calculated column "
                f"'{name}' formula {prepared[:120]!r}: {exc}"
            )
            raise RuntimeError(msg) from exc


def _extend_conditional_formatting_com(
    app: Any,  # noqa: ANN401
    worksheet: Any,  # noqa: ANN401
    *,
    start_col: int,
    end_col: int,
    first_data_row: int,
    last_data_row: int,
) -> None:
    """Resize CF ranges that start on the first data row to the table's last row."""
    from fillxl._tables import _cf_range_tracks_table

    conditions = worksheet.Cells.FormatConditions
    for index in range(1, conditions.Count + 1):
        condition = conditions.Item(index)
        applies_to = condition.AppliesTo
        new_areas: list[Any] = []
        changed = False
        for area_index in range(1, applies_to.Areas.Count + 1):
            area = applies_to.Areas(area_index)
            bounds = CellRange(
                min_col=area.Column,
                min_row=area.Row,
                max_col=area.Column + area.Columns.Count - 1,
                max_row=area.Row + area.Rows.Count - 1,
            )
            if _cf_range_tracks_table(
                bounds,
                table_start_column=start_col,
                table_end_column=end_col,
                first_data_row=first_data_row,
                table_last_row=last_data_row,
            ):
                area = area.Resize(last_data_row - first_data_row + 1, area.Columns.Count)
                changed = True
            new_areas.append(area)
        if not changed:
            continue
        new_range = new_areas[0]
        for area in new_areas[1:]:
            new_range = app.Union(new_range, area)
        condition.ModifyAppliesToRange(new_range)


_SERIES_PREFIX = "=SERIES("
# =SERIES(name, categories, values, order[, bubble sizes]): the name and the
# plot order are never table data, so only these arguments are rewritten.
_SERIES_DATA_ARGUMENTS = (1, 2, 4)


def _update_series_formula(formula: str, spec: TableChangeSpecs) -> str | None:
    """Update a chart series formula to reflect the table's new last row.

    Returns:
        The updated formula string, or ``None`` if no update was needed.

    Examples:
        >>> spec = TableChangeSpecs("Data", 12, 2, 13, 13, 20)
        >>> _update_series_formula("=SERIES(T,'Data'!$B$13:$B$13,1)", spec)
        "=SERIES(T,'Data'!$B$13:$B$20,1)"
        >>> _update_series_formula("=SERIES(T,'Data2'!$B$13:$B$13,1)", spec) is None
        True
        >>> _update_series_formula("=SERIES(Data!$B$13,,Data!$C$13,1)", spec)
        '=SERIES(Data!$B$13,,Data!$C$13:$C$20,1)'

    """
    if not (formula.upper().startswith(_SERIES_PREFIX) and formula.endswith(")")):
        return None
    head = formula[: len(_SERIES_PREFIX)]
    arguments = split_formula_arguments(formula[len(_SERIES_PREFIX) : -1])
    changed = False
    for index in _SERIES_DATA_ARGUMENTS:
        if index < len(arguments):
            updated = update_reference_text(arguments[index], spec)
            if updated is not None:
                arguments[index] = updated
                changed = True
    return head + ",".join(arguments) + ")" if changed else None
