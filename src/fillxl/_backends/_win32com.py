"""Win32com (pywin32) backend implementation for fillxl.

Drives real Microsoft Excel via COM automation (pywin32), providing
capabilities that openpyxl cannot: rendered conditional formatting colors,
real formula evaluation, display format application, and true pivot table
refresh.

Only works on **Windows** with Microsoft Excel installed.
"""

from __future__ import annotations

import contextlib
import logging
import re
import sys
import time
from typing import TYPE_CHECKING, Any

from openpyxl.worksheet.cell_range import CellRange

from fillxl._charts import TableChangeSpecs, _unquote_sheet_name
from fillxl._normalize import prepare_com_value

if TYPE_CHECKING:
    from collections.abc import Sequence
    from pathlib import Path

    from fillxl._config import FillXLConfig

logger = logging.getLogger(__name__)

_is_windows = sys.platform == "win32"

# COM constants (fallback if early binding is unavailable)
_XL_SHEET_HIDDEN = 0
_XL_PASTE_FORMATS = -4122

# SaveAs file formats, keyed by output suffix.  Saving an .xlsm template with
# the .xlsx format silently strips its macros, so the format is inferred from
# the requested output path instead of being hard-coded.
_FILE_FORMATS: dict[str, int] = {
    ".xlsx": 51,  # xlOpenXMLWorkbook
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12
    ".xls": 56,  # xlExcel8
}

# A chart series formula contains references such as 'Sheet 1'!$B$13:$B$13.
_SERIES_RANGE_RE = re.compile(
    r"(?:'[^']+'|[A-Za-z0-9_.]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?",
)


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
        """Lazy-initialise the Excel Application COM object.

        Creates a hidden Excel instance with alerts and screen updating
        disabled.  Uses ``win32com.client.Dispatch`` directly (rather than
        ``gencache.EnsureDispatch``) to avoid RPC issues when the type
        library cache is being built.
        """
        if self._app is not None:
            return
        self._ensure_platform()
        # Late import so the module is importable on non-Windows (for tests)
        import pythoncom  # noqa: F401
        import win32com.client

        # Use DispatchEx to always create a new, isolated Excel instance.
        # Dispatch would reuse a stale/hung instance which can cause
        # AttributeError on property access.
        self._app = win32com.client.DispatchEx("Excel.Application")
        # These properties may not be settable depending on the Excel
        # process state or permissions — non-fatal, just best-effort.
        for attr, val in [("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False)]:
            try:
                setattr(self._app, attr, val)
            except AttributeError:
                logger.warning("Could not set Excel.Application.%s", attr)
        logger.debug("Excel Application started (hidden)")

    def _ensure_workbook(self) -> None:
        """Raise :class:`RuntimeError` if no workbook has been loaded."""
        if self._workbook is None:
            msg = "No workbook loaded — call load() first"
            raise RuntimeError(msg)

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
        self._workbook = self._app.Workbooks.Open(str(abs_path))
        logger.info("[win32com] load took %.3fs", time.perf_counter() - t0)

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

        # Same normalisation as the table path (pandas objects, date/time
        # promotion, Decimal, tz stripping) plus the apostrophe prefix that
        # keeps COM from reinterpreting strings, and the UTC marking that
        # keeps it from shifting datetimes by the local timezone offset.
        prepared = prepare_com_value(value, escape_formulas=config.escape_formulas)

        try:
            name_obj = self._workbook.Names(name)
            name_obj.RefersToRange.Value = prepared
        except Exception as exc:
            msg = f"Named cell '{name}' not found in workbook"
            raise KeyError(msg) from exc

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

        # ── Locate the ListObject ──────────────────────────────────────────
        list_obj: Any = None
        worksheet: Any = None
        for sheet in self._workbook.Sheets:
            try:
                candidate = sheet.ListObjects(name)
                if candidate:
                    list_obj = candidate
                    worksheet = sheet
                    break
            except Exception:
                continue

        if list_obj is None or worksheet is None:
            msg = f"Table '{name}' not found in workbook"
            raise KeyError(msg)

        if _com_has_totals_row(list_obj):
            msg = (
                f"Table '{name}' has a totals row, which fillxl does not "
                f"support: writing the data would overwrite the totals and "
                f"miscount the rows. Remove the totals row from the template."
            )
            raise ValueError(msg)

        # ── Normalise data ────────────────────────────────────────────────
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(data)
        data_columns = adapter.columns
        rows = adapter.rows
        num_data_rows = len(rows)

        # ── Get table metadata ────────────────────────────────────────────
        t_meta = time.perf_counter()
        header_range = list_obj.HeaderRowRange
        header_row = header_range.Row
        start_col = header_range.Column
        num_cols = header_range.Columns.Count

        from fillxl._utils import col_to_excel

        start_col_letter = col_to_excel(start_col)
        end_col_letter = col_to_excel(start_col + num_cols - 1)

        # Column names from the table — read all at once (single COM call)
        header_values = header_range.Value
        if isinstance(header_values, tuple) and len(header_values) > 0:
            row_tuple = header_values[0]
            table_columns = [str(row_tuple[i]) for i in range(num_cols)]
        else:
            table_columns = [str(header_values)]
        logger.info("[win32com] table metadata took %.3fs", time.perf_counter() - t_meta)

        # ── Column validation (must happen before resize so we know what to write) ──
        from fillxl._validation import validate_columns

        formula_columns, formula_map = _read_calculated_column_formulas(self._load_path, name)

        validate_columns(
            table_columns=table_columns,
            data_columns=data_columns,
            strategy=config.column_validation_strategy,
            formula_columns=formula_columns,
        )

        # ── Measure the existing table ────────────────────────────────────
        data_body_range = list_obj.DataBodyRange
        if data_body_range is not None:
            old_data_row = data_body_range.Row
            old_last_row = old_data_row + data_body_range.Rows.Count - 1
        else:
            old_data_row = header_row + 1
            old_last_row = header_row

        worksheet_title = str(worksheet.Name)

        # ── Empty data: keep a single, cleared body row ───────────────────
        # (openpyxl does the same; an Excel table needs at least one body row.)
        if num_data_rows == 0:
            if data_body_range is not None:
                data_body_range.ClearContents()
            new_last_row = header_row + 1
            if old_last_row != new_last_row:
                new_range = self._app.Range(
                    f"{start_col_letter}{header_row}:{end_col_letter}{new_last_row}",
                )
                list_obj.Resize(new_range)
                worksheet.Range(
                    f"{start_col_letter}{new_last_row}:{end_col_letter}{new_last_row}",
                ).ClearContents()
            return TableChangeSpecs(
                worksheet_title=worksheet_title,
                header_row=header_row,
                first_col=start_col,
                last_col=start_col + num_cols - 1,
                old_last_row=old_last_row,
                new_last_row=new_last_row,
            )

        new_end_row = header_row + num_data_rows
        new_table_range = f"{start_col_letter}{header_row}:{end_col_letter}{new_end_row}"

        # Check for overflow — cells below the table should be empty.
        # Read the entire overflow range in a single COM call (batch read)
        # instead of cell-by-cell (which is ~100x slower for large expansions).
        t_overflow = time.perf_counter()
        if num_data_rows > (old_last_row - header_row):
            overflow_start = old_last_row + 1
            overflow_range_str = f"{start_col_letter}{overflow_start}:{end_col_letter}{new_end_row}"
            try:
                overflow_values = worksheet.Range(overflow_range_str).Value
                if overflow_values is not None:
                    if isinstance(overflow_values, tuple):
                        for row_idx, row_tuple in enumerate(overflow_values):
                            for col_idx, val in enumerate(row_tuple):
                                if val is not None:
                                    check_row = overflow_start + row_idx
                                    check_col = start_col + col_idx
                                    cell_ref = f"{col_to_excel(check_col)}{check_row}"
                                    msg = (
                                        f"Table '{name}' cannot expand to "
                                        f"{num_data_rows} rows: cell {cell_ref} "
                                        f"already contains data ({val!r}). "
                                        "Move or remove the content below the "
                                        "table first."
                                    )
                                    raise ValueError(msg)
                    elif overflow_values is not None:
                        # Single cell overflow range
                        check_row = overflow_start
                        check_col = start_col
                        cell_ref = f"{col_to_excel(check_col)}{check_row}"
                        msg = (
                            f"Table '{name}' cannot expand to {num_data_rows} "
                            f"rows: cell {cell_ref} already contains data "
                            f"({overflow_values!r}). Move or remove the content "
                            "below the table first."
                        )
                        raise ValueError(msg)
            except ValueError:
                raise
            except Exception:
                # COM errors during overflow check → fall through to resize
                pass
        logger.info("[win32com] overflow check took %.3fs", time.perf_counter() - t_overflow)

        # Resize the table first — Excel auto-fills calculated column formulas
        # in new rows when the table is resized.
        t_resize = time.perf_counter()
        excel_range = self._app.Range(new_table_range)
        list_obj.Resize(excel_range)
        logger.info("[win32com] resize took %.3fs", time.perf_counter() - t_resize)

        # Clear old data rows if the table has shrunk — batch clear
        t_clear = time.perf_counter()
        if num_data_rows < (old_last_row - header_row):
            clear_range_str = f"{start_col_letter}{new_end_row + 1}:{end_col_letter}{old_last_row}"
            with contextlib.suppress(Exception):
                worksheet.Range(clear_range_str).ClearContents()
        logger.info("[win32com] clear took %.3fs", time.perf_counter() - t_clear)

        # ── Write data ────────────────────────────────────────────────────
        new_data_body = list_obj.DataBodyRange
        data_col_index = {col: idx for idx, col in enumerate(data_columns)}

        # Build a 2D array (list of lists) matching table column order.
        # Formula columns get None (Excel already filled them).
        t_write = time.perf_counter()
        values_matrix: list[list[Any]] = []
        for row in rows:
            row_values: list[Any] = []
            for col_name in table_columns:
                if col_name in data_col_index and col_name not in formula_columns:
                    # Central normalisation — guarantees identical results
                    # with the openpyxl backend (apostrophe-prefixed strings
                    # and the UTC marking for datetimes).
                    val = prepare_com_value(
                        row[data_col_index[col_name]],
                        escape_formulas=config.escape_formulas,
                    )
                    row_values.append(val)
                else:
                    row_values.append(None)
            values_matrix.append(row_values)

        # Write the entire matrix at once (much faster than cell-by-cell)
        if values_matrix:
            new_data_body.Value = values_matrix
        logger.info("[win32com] write data took %.3fs", time.perf_counter() - t_write)

        # Re-apply calculated column formulas.  Writing ``None`` above clears
        # the auto-filled formulas, so we must restore them here.  Preserve
        # newlines/indentation so user formatting is not destroyed.
        t_formula = time.perf_counter()
        for col_idx, col_name in enumerate(table_columns):
            if col_name in formula_columns:
                formula = formula_map.get(col_name)
                if formula:
                    col_data_range = new_data_body.Columns(col_idx + 1)
                    formula_str = _prepare_formula_for_com(str(formula))
                    try:
                        col_data_range.Formula2 = formula_str
                    except AttributeError:
                        # Excel < 2016 — fall back to Formula (255-char limit)
                        col_data_range.Formula = formula_str
                    except Exception:
                        # The formula may contain functions (LET, LAMBDA, SCAN,
                        # etc.) not available in this Excel version.  Log and
                        # skip — Excel already auto-filled the column and the
                        # data rows have been written.
                        logger.warning(
                            "[win32com] Could not set formula for column '%s': %s",
                            col_name,
                            formula_str[:120],
                        )
        logger.info("[win32com] formula columns took %.3fs", time.perf_counter() - t_formula)

        # ── Copy styles from first data row ───────────────────────────────
        t_style = time.perf_counter()
        if config.copy_table_styles and num_data_rows > 1:
            # Copy only the table data body columns (not the full worksheet
            # row) to avoid a column shift when pasting.
            first_data_row_range = new_data_body.Rows(1)
            first_data_row_range.Copy()
            dest_range = worksheet.Range(
                f"{start_col_letter}{header_row + 2}:{end_col_letter}{new_end_row}",
            )
            dest_range.PasteSpecial(Paste=_XL_PASTE_FORMATS)
            # Clear the clipboard (copy mode)
            self._app.CutCopyMode = False
        logger.info("[win32com] copy styles took %.3fs", time.perf_counter() - t_style)

        # ── Extend conditional formatting ─────────────────────────────────
        t_cf = time.perf_counter()
        if config.extend_conditional_formatting:
            _extend_conditional_formatting_com(
                worksheet=worksheet,
                start_col_letter=start_col_letter,
                end_col_letter=end_col_letter,
                first_data_row=header_row + 1,
                last_data_row=new_end_row,
            )
        logger.info("[win32com] conditional formatting took %.3fs", time.perf_counter() - t_cf)

        return TableChangeSpecs(
            worksheet_title=worksheet_title,
            header_row=header_row,
            first_col=start_col,
            last_col=start_col + num_cols - 1,
            old_last_row=old_last_row,
            new_last_row=new_end_row,
        )

    def update_charts(self, spec: TableChangeSpecs) -> None:
        """Update chart data ranges after a table resize.

        Iterates all ``ChartObject`` items across all worksheets and updates
        series formulas that reference the resized table.

        Args:
            spec: The table dimension change to reflect in charts.

        """
        self._ensure_workbook()

        for sheet in self._workbook.Sheets:
            try:
                chart_objects = sheet.ChartObjects()
            except Exception:
                continue

            for chart_obj in chart_objects:
                chart = chart_obj.Chart
                for series in chart.SeriesCollection():
                    try:
                        formula = series.Formula
                        if not formula:
                            continue
                    except Exception:
                        continue

                    # Parse the formula to find end row references
                    # Formula format: =SERIES(name, x_values, y_values, plot_order)
                    # Each range is like: 'Sheet'!$A$2:$A$N
                    updated = _update_series_formula(formula, spec)
                    if updated:
                        logger.debug("Updated chart series formula: %s -> %s", formula, updated)
                        series.Formula = updated

    def set_pivot_refresh(self) -> None:
        """Mark all pivot caches for refresh on load."""
        self._ensure_workbook()
        for pcache in self._workbook.PivotCaches():
            try:
                pcache.Refresh()
            except Exception:
                logger.exception("Failed to refresh pivot cache")
        logger.debug("Refreshed all pivot caches")

    def hide_sheets(self, prefix: str) -> None:
        """Hide worksheets whose names start with *prefix*.

        Args:
            prefix: Sheet name prefix (e.g. ``"_"``).

        """
        self._ensure_workbook()
        for sheet in self._workbook.Sheets:
            try:
                name = str(sheet.Name)
                if name.startswith(prefix):
                    sheet.Visible = _XL_SHEET_HIDDEN
                    logger.debug("Hidden worksheet '%s'", name)
            except Exception:
                logger.exception("Failed to hide sheet")

    def delete_sheets(self, names: Sequence[str]) -> None:
        """Delete the worksheets named in *names*.

        Args:
            names: Sheet names to remove.

        """
        self._ensure_workbook()
        for sheet_name in names:
            try:
                sheet = self._workbook.Sheets(sheet_name)
                sheet.Delete()
                logger.debug("Deleted worksheet '%s'", sheet_name)
            except Exception:
                logger.warning("Sheet '%s' not found, skipping deletion", sheet_name)

    def set_active_sheet(self, name: str) -> None:
        """Set the active (displayed) worksheet.

        Args:
            name: Name of the worksheet to activate.

        """
        self._ensure_workbook()
        try:
            sheet = self._workbook.Sheets(name)
            sheet.Activate()
            logger.debug("Set active worksheet to '%s'", name)
        except Exception:
            logger.warning("Sheet '%s' not found, keeping default active sheet", name)

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
        logger.debug("Saving workbook to %s (format %d)", abs_path, file_format)
        self._workbook.SaveAs(str(abs_path), FileFormat=file_format)
        logger.info("[win32com] save took %.3fs", time.perf_counter() - t0)

    def close(self) -> None:
        """Close the workbook and release the Excel Application."""
        try:
            if self._workbook is not None:
                try:
                    self._workbook.Close(SaveChanges=False)
                except Exception:
                    logger.exception("Error closing workbook")
                finally:
                    self._workbook = None
        finally:
            if self._app is not None:
                try:
                    self._app.Quit()
                except Exception:
                    logger.exception("Error during Excel quit")
                finally:
                    self._app = None
                    logger.debug("Excel Application quit")


# ── Standalone helpers ──────────────────────────────────────────────────────


def _com_has_totals_row(list_obj: Any) -> bool:  # noqa: ANN401
    """Whether a COM ``ListObject`` is configured with a totals row."""
    return bool(getattr(list_obj, "ShowTotals", False))


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


def _read_calculated_column_formulas(
    path: Path | None,
    table_name: str,
) -> tuple[list[str], dict[str, str]]:
    """Read calculated column formulas from the template using openpyxl.

    COM does not reliably expose ``calculatedColumnFormula`` for tables
    created by openpyxl.  This helper reads the table definition directly
    from the ``.xlsx`` file to detect formula columns.

    Args:
        path: Path to the workbook file.
        table_name: Name of the table to inspect.

    Returns:
        A tuple of (formula_column_names, formula_map) where *formula_map*
        maps column names to their formula strings (with ``=`` prefix).

    """
    if path is None or not path.exists():
        return [], {}

    try:
        import openpyxl as _op

        wb = _op.load_workbook(str(path))
        try:
            from fillxl._utils import find_worksheet_and_table

            _, table = find_worksheet_and_table(wb, table_name)

            formula_columns: list[str] = []
            formula_map: dict[str, str] = {}

            for col in table.tableColumns:
                if col.calculatedColumnFormula:
                    name = str(col.name)
                    attr_text = col.calculatedColumnFormula.attr_text
                    if attr_text:
                        # openpyxl sets calculatedColumnFormula for any
                        # non-empty default value, not just real formulas.
                        # Skip plain numbers, booleans, and simple strings
                        # — they are default values, not formulas.
                        text = str(attr_text)
                        if text.isdigit() or text in ("TRUE", "FALSE", "true", "false"):
                            continue
                        try:
                            float(text)
                            continue
                        except ValueError:
                            pass
                        formula_str = "=" + text
                        formula_columns.append(name)
                        formula_map[name] = formula_str
                        logger.debug(
                            "Detected calculated column '%s' formula=%s",
                            name,
                            formula_str,
                        )

            return formula_columns, formula_map
        finally:
            wb.close()
    except Exception:
        logger.exception("Failed to read calculated column formulas from %s", path)
        return [], {}


def _extend_conditional_formatting_com(
    worksheet: Any,  # noqa: ANN401
    start_col_letter: str,
    end_col_letter: str,
    first_data_row: int,
    last_data_row: int,
) -> None:
    """Extend conditional formatting rules to cover all data rows.

    Works by copying formatting from the first data row down (which carries
    CF rules along when pasting formats).

    Args:
        worksheet: The COM Worksheet object.
        start_col_letter: Starting column letter (e.g., ``'B'``).
        end_col_letter: Ending column letter (e.g., ``'M'``).
        first_data_row: Row number of the first data row.
        last_data_row: Row number of the last data row.

    """
    try:
        # Copy from the entire first data row (not just the first cell) so
        # that each column's individual NumberFormat is preserved.
        source_range = worksheet.Range(
            f"{start_col_letter}{first_data_row}:{end_col_letter}{first_data_row}",
        )
        source_range.Copy()
        dest_range = worksheet.Range(
            f"{start_col_letter}{first_data_row}:{end_col_letter}{last_data_row}",
        )
        dest_range.PasteSpecial(Paste=_XL_PASTE_FORMATS)
        # Clear clipboard — use the worksheet's own Application reference
        with contextlib.suppress(Exception):
            worksheet.Application.CutCopyMode = False
    except Exception:
        logger.exception("Failed to extend conditional formatting")


def _update_series_formula(formula: str, spec: TableChangeSpecs) -> str | None:
    """Update a chart series formula to reflect the table's new last row.

    Only references that point at the table's data body — starting one row
    below the header, ending at the table's last row, and within the table's
    columns — are updated.  The sheet name is compared exactly, so a chart on
    ``Data2`` is left alone while ``Data`` is filled.

    Args:
        formula: The original series formula (e.g., ``=SERIES(...)``).
        spec: The table change to apply.

    Returns:
        The updated formula string, or ``None`` if no update was needed.

    Examples:
        >>> spec = TableChangeSpecs("Data", 12, 2, 13, 13, 20)
        >>> _update_series_formula("=SERIES(T,'Data'!$B$13:$B$13,1)", spec)
        "=SERIES(T,'Data'!$B$13:$B$20,1)"
        >>> _update_series_formula("=SERIES(T,'Data2'!$B$13:$B$13,1)", spec) is None
        True

    """
    updated_any = False

    def _replace_reference(match: re.Match[str]) -> str:
        nonlocal updated_any
        reference = match.group(0)
        sheet_part, separator, cell_part = reference.rpartition("!")
        if not separator or _unquote_sheet_name(sheet_part) != spec.worksheet_title:
            return reference
        try:
            cell_range = CellRange(cell_part)
        except ValueError:
            return reference
        if cell_range.min_row != spec.first_data_row or cell_range.max_row != spec.old_last_row:
            return reference
        if cell_range.min_col < spec.first_col or cell_range.max_col > spec.last_col:
            return reference
        updated_any = True
        # Replace only the trailing row number, preserving "$" prefixes.
        new_cell_part = re.sub(r"\d+$", str(spec.new_last_row), cell_part)
        return f"{sheet_part}!{new_cell_part}"

    updated = str(_SERIES_RANGE_RE.sub(_replace_reference, formula))
    return updated if updated_any else None
