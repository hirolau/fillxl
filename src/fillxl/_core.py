"""Core fill orchestrator for fillxl."""

from __future__ import annotations

import logging
import shutil
import tempfile
from pathlib import Path
from typing import TYPE_CHECKING, Any

from fillxl._backend import get_backend
from fillxl._config import FillXLConfig
from fillxl._normalize import reset_tz_warning_state

if TYPE_CHECKING:
    from fillxl._types import Engine, PathLike, TabularData

logger = logging.getLogger(__name__)


def fill(
    template: PathLike,
    output: PathLike,
    values: dict[str, Any] | None = None,
    tables: dict[str, TabularData] | None = None,
    *,
    engine: Engine = "auto",
    config: FillXLConfig | None = None,
) -> Path:
    """Fill an Excel template with data.

    Accepts separate mappings for table data (``tables``) and named-cell
    values (``values``).  Tabular data can be any of the supported formats
    (pandas/polars/duckdb DataFrames, ``list[dict]``, or
    ``dict[str, sequence]``).

    Args:
        template: Path to the Excel template file.
        output: Path where the filled output will be saved.
        tables: Mapping of Excel table names to tabular data.
        values: Mapping of named cell / defined name to scalar values.
        engine: Backend engine to use (``"auto"``, ``"openpyxl"``, or
            ``"win32com"``).  Defaults to ``"auto"``, which selects the
            best available backend for the current platform.
        config: Configuration options.  Uses :class:`FillXLConfig` defaults
            if not provided.

    Returns:
        Absolute :class:`~pathlib.Path` to the output file.

    Raises:
        FileNotFoundError: If *template* does not exist.
        KeyError: If a name in *values* or *tables* is not found in the
            workbook.
        ValueError: If *engine* is unknown or if column validation fails
            (when ``column_validation_strategy`` is ``"strict"``).

    Note:
        This function requires a real Excel file — it is **not** suitable for
        doctest execution.  See the test suite for runnable examples.

    """
    cfg = config or FillXLConfig()
    tables = tables or {}
    values = values or {}

    # Reset the per-fill timezone warning state so each fill session
    # gets its own "first occurrence" warning.
    reset_tz_warning_state()

    template_path = Path(template)
    output_path = Path(output)

    # Validate template exists
    if not template_path.exists():
        msg = f"Template file not found: {template_path}"
        raise FileNotFoundError(msg)

    # Work on a copy to preserve the original template
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_template_path = Path(temp_dir) / template_path.name
        shutil.copy2(template_path, temp_template_path)

        backend = get_backend(engine)
        backend.load(temp_template_path)

        try:
            # 1. Fill named cells with scalar values
            if values:
                for name, value in values.items():
                    backend.fill_named_cell(name, value, cfg)
                logger.debug("Filled %d named cell(s)", len(values))

            # 2. Fill tables and update chart references
            if tables:
                for table_name, table_data in tables.items():
                    spec = backend.fill_table(table_name, table_data, cfg)
                    backend.update_charts(spec)
                logger.debug("Filled %d table(s)", len(tables))

            # 3. Pivot cache refresh
            if cfg.refresh_pivot_tables:
                backend.set_pivot_refresh()

            # 4. Hide sheets starting with underscore
            if cfg.hide_underscore_sheets:
                backend.hide_sheets("_")

            # 5. Delete specified sheets
            if cfg.delete_sheets:
                backend.delete_sheets(cfg.delete_sheets)

            # 6. Set active sheet
            if cfg.active_sheet:
                backend.set_active_sheet(cfg.active_sheet)

            # 7. Ensure output directory exists
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # 8. Save and close
            backend.save(output_path)

        finally:
            backend.close()

    return output_path
