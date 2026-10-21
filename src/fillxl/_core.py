"""Core fill orchestrator for fillxl."""

from __future__ import annotations

import logging
import shutil
import tempfile
import uuid
import warnings
from pathlib import Path
from typing import TYPE_CHECKING, Any

from fillxl._backend import get_backend, resolve_engine
from fillxl._config import FillXLConfig
from fillxl._normalize import reset_tz_warning_state

if TYPE_CHECKING:
    from fillxl._types import Engine, PathLike, TabularData

logger = logging.getLogger(__name__)


def fill(
    template: PathLike,
    output: PathLike,
    *legacy_args: Any,  # noqa: ANN401 - deprecated positional values/tables, removed in 0.7.0
    values: dict[str, Any] | None = None,
    tables: dict[str, TabularData] | None = None,
    engine: Engine = "auto",
    config: FillXLConfig | None = None,
) -> Path:
    """Fill an Excel template with data.

    Args:
        template: Path to the Excel template file.
        output: Path where the filled output will be saved.
        values: Mapping of named cell / defined name to scalar values.
        tables: Mapping of Excel table names to tabular data.
        engine: ``"openpyxl"``, ``"win32com"`` or ``"auto"`` (default).
            ``"auto"`` uses openpyxl for ``.xlsx``/``.xlsm`` templates saved in
            the same format, and Excel (win32com) when it is needed and
            available: ``.xlsb``/``.xls`` templates, or a different output format.
        config: Configuration options.  Uses :class:`FillXLConfig` defaults
            if not provided.

    Returns:
        Absolute :class:`~pathlib.Path` to the output file.

    Raises:
        FileNotFoundError: If *template* does not exist.
        KeyError: If a name in *values* or *tables* is not found in the workbook.
        ValueError: If *engine* is unknown, column validation fails, or the
            sheet settings would leave no visible sheet.

    Note:
        Passing *values* and *tables* positionally is deprecated (0.6.0) and
        will stop working in 0.7.0.

    """
    if legacy_args:
        values, tables = _legacy_positional_arguments(legacy_args, values, tables)
    cfg = config or FillXLConfig()
    tables = tables or {}
    values = values or {}

    # Each fill session gets its own "first occurrence" timezone warning.
    reset_tz_warning_state()

    template_path = Path(template)
    output_path = Path(output)
    if not template_path.exists():
        msg = f"Template file not found: {template_path}"
        raise FileNotFoundError(msg)
    resolved_engine = resolve_engine(engine, template_path, output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    # A sibling path keeps the final rename on one volume, so the swap is atomic.
    staging_path = output_path.with_name(
        f".{output_path.stem}.{uuid.uuid4().hex[:12]}.fillxl{output_path.suffix}",
    )
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_template_path = Path(temp_dir) / template_path.name
            shutil.copy2(template_path, temp_template_path)
            backend = get_backend(resolved_engine)
            try:
                # load() runs inside try so a failed open still closes Excel.
                backend.load(temp_template_path)
                for name, value in values.items():
                    backend.fill_named_cell(name, value, cfg)
                for table_name, table_data in tables.items():
                    spec = backend.fill_table(table_name, table_data, cfg)
                    backend.update_charts(spec)
                if cfg.refresh_pivot_tables:
                    backend.set_pivot_refresh()
                if cfg.hide_underscore_sheets:
                    backend.hide_sheets("_")
                if cfg.delete_sheets:
                    backend.delete_sheets(cfg.delete_sheets)
                if cfg.active_sheet:
                    backend.set_active_sheet(cfg.active_sheet)
                backend.save(staging_path)
            finally:
                backend.close()
        backend.check_output(template_path, staging_path, cfg.on_lossy)
        staging_path.replace(output_path)
    except BaseException:
        staging_path.unlink(missing_ok=True)
        raise

    return output_path.resolve()


def _legacy_positional_arguments(
    args: tuple[Any, ...],
    values: dict[str, Any] | None,
    tables: dict[str, TabularData] | None,
) -> tuple[dict[str, Any] | None, dict[str, TabularData] | None]:
    """Accept the 0.5 call ``fill(template, output, values, tables)`` with a warning."""
    if len(args) > 2:
        msg = f"fill() takes at most 4 positional arguments ({2 + len(args)} given)"
        raise TypeError(msg)
    if values is not None:
        msg = "fill() got multiple values for argument 'values'"
        raise TypeError(msg)
    if len(args) == 2 and tables is not None:
        msg = "fill() got multiple values for argument 'tables'"
        raise TypeError(msg)
    warnings.warn(
        "Passing 'values' and 'tables' to fill() positionally is deprecated and stops "
        "working in fillxl 0.7.0; use fill(template, output, values=..., tables=...).",
        DeprecationWarning,
        stacklevel=3,
    )
    return args[0], (args[1] if len(args) == 2 else tables)
