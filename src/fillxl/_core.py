"""Core fill orchestrator for fillxl."""

from __future__ import annotations

import logging
import shutil
import tempfile
import uuid
from collections.abc import Mapping
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
    *,
    values: dict[str, Any] | None = None,
    tables: dict[str, TabularData] | None = None,
    engine: Engine = "auto",
    config: FillXLConfig | None = None,
) -> Path:
    """Fill an Excel template with data.

    Args:
        template: Path to the Excel template file.
        output: Path where the filled output will be saved.  Must differ from
            *template*: the template is never overwritten.
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
        TypeError: If *config*, *values* or *tables* is of the wrong type, or a
            named cell is given something other than a single value.
        ValueError: If *engine* is unknown, *output* is *template*, the engine
            cannot write the requested format, column validation fails, or the
            sheet settings would leave no visible sheet.

    """
    cfg = _checked_config(config)
    named_values = _checked_mapping("values", values)
    table_data = _checked_mapping("tables", tables)

    # Each fill session gets its own "first occurrence" timezone warning.
    reset_tz_warning_state()

    template_path = Path(template)
    output_path = Path(output)
    if not template_path.exists():
        msg = f"Template file not found: {template_path}"
        raise FileNotFoundError(msg)
    if output_path.resolve() == template_path.resolve():
        msg = (
            f"output and template are the same file ({template_path}); fillxl never "
            "overwrites its template. Write to another path and move the result "
            "yourself if that is really what you want."
        )
        raise ValueError(msg)
    resolved_engine = resolve_engine(engine, template_path, output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    # A sibling path keeps the final rename on one volume, so the swap is atomic.
    staging_path = output_path.with_name(
        f".{output_path.stem}.{uuid.uuid4().hex[:12]}.fillxl{output_path.suffix}",
    )
    try:
        # ignore_cleanup_errors: Excel can still hold the temp copy for a
        # moment after Quit, and a failed cleanup must not fail the fill.
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp_dir:
            temp_template_path = Path(temp_dir) / template_path.name
            shutil.copy2(template_path, temp_template_path)
            backend = get_backend(resolved_engine)
            try:
                # load() runs inside try so a failed open still closes Excel.
                backend.load(temp_template_path)
                for name, value in named_values.items():
                    backend.fill_named_cell(name, value, cfg)
                for table_name, data in table_data.items():
                    spec = backend.fill_table(table_name, data, cfg)
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
        _replace_output(staging_path, output_path)
    except BaseException:
        staging_path.unlink(missing_ok=True)
        raise

    return output_path.resolve()


def _checked_config(config: FillXLConfig | None) -> FillXLConfig:
    """The configuration to use, refusing anything that is not a ``FillXLConfig``."""
    if config is None:
        return FillXLConfig()
    if not isinstance(config, FillXLConfig):
        msg = (
            f"config must be a FillXLConfig, got {type(config).__name__}. "
            "Pass FillXLConfig(...) rather than a dict of settings."
        )
        raise TypeError(msg)
    return config


def _checked_mapping(name: str, mapping: Mapping[str, Any] | None) -> dict[str, Any]:
    """Validate a ``values``/``tables`` argument and copy it into a ``dict``.

    Keys are checked because every lookup matches Excel names as text: a
    ``1`` key would never find the table named ``"1"``, and the failure would
    look like a missing table.
    """
    if mapping is None:
        return {}
    if not isinstance(mapping, Mapping):
        msg = (
            f"{name} must be a mapping of name to data, got {type(mapping).__name__}. "
            f"Pass a dict: {name}={{...}}"
        )
        raise TypeError(msg)
    for key in mapping:
        if not isinstance(key, str):
            msg = (
                f"{name} keys must be strings, got {type(key).__name__} {key!r}. "
                "Excel names are matched as text."
            )
            raise TypeError(msg)
    return dict(mapping)


def _replace_output(staging_path: Path, output_path: Path) -> None:
    """Move the finished file into place, explaining the common failure.

    On Windows the rename fails while anything holds the destination open —
    most often the output being open in Excel.
    """
    try:
        staging_path.replace(output_path)
    except PermissionError as exc:
        msg = (
            f"Could not write {output_path}: permission denied. Is the file open in "
            "Excel? Close it and run the fill again."
        )
        raise PermissionError(msg) from exc
