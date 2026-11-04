"""Core fill orchestrator for fillxl."""

from __future__ import annotations

import logging
import shutil
import tempfile
import uuid
from collections.abc import Mapping
from pathlib import Path
from typing import TYPE_CHECKING, Any

from fillxl._backend import get_backend, reject_xls_paths, resolve_engine
from fillxl._config import FillXLConfig
from fillxl._normalize import reset_tz_warning_state

if TYPE_CHECKING:
    from fillxl._types import Engine, PathLike, TabularData

logger = logging.getLogger(__name__)


def _mapping_argument(value: object, name: str) -> Mapping[str, Any]:
    """*value* as a mapping, or ``TypeError`` naming the argument.

    A list of ``(name, value)`` pairs, a bare string or a DataFrame used to
    reach ``values.items()`` and fail with ``AttributeError`` after the
    template had been opened.
    """
    if value is None:
        return {}
    if not isinstance(value, Mapping):
        msg = (
            f"fill() argument {name}= must be a mapping of names to values, "
            f"not {type(value).__name__}."
        )
        raise TypeError(msg)
    return value


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

    The template is never modified: it is copied to a temporary directory, the
    copy is filled, and the result is written to a staging file next to
    *output*.  The staging file replaces *output* only after saving, closing
    the backend and the ``on_lossy`` check have all succeeded, so a failed fill
    leaves an existing *output* exactly as it was.  *template* and *output* may
    be the same path, in which case the filled workbook replaces the template.

    Args:
        template: Path to the Excel template file.
        output: Path where the filled output will be saved.
        values: Mapping of named cell / defined name to scalar values.
        tables: Mapping of Excel table names to tabular data.
        engine: ``"openpyxl"``, ``"fillxl"``, ``"win32com"`` or ``"auto"``
            (default).  ``"auto"`` uses openpyxl for ``.xlsx``/``.xlsm``
            templates saved in the same format, and Excel (win32com) when it is
            needed and available: ``.xlsb`` templates, or a different
            output format.  ``"fillxl"`` edits the OOXML package itself: no
            Excel, and no template feature lost.  It is experimental and
            opt-in — ``"auto"`` never selects it — and it only handles
            ``.xlsx``/``.xlsm``, writes no calculated values, and leaves
            ``chartEx``, pivot sources and data validations alone; see the
            README for the full list.
        config: Configuration options.  Uses :class:`FillXLConfig` defaults
            if not provided.

    Returns:
        Absolute :class:`~pathlib.Path` to the output file.

    Raises:
        FileNotFoundError: If *template* does not exist.
        IsADirectoryError: If *output* is an existing directory.
        TypeError: If *values* or *tables* is not a mapping, or *config* is not
            a :class:`FillXLConfig`.
        KeyError: If a name in *values* or *tables* is not found in the workbook.
        CellValueError: If a value cannot be stored in an Excel cell.
        LossyTemplateError: If ``on_lossy="raise"`` and the openpyxl engine
            dropped template features.
        PermissionError: If *output* cannot be replaced, usually because it is
            open in Excel.
        ValueError: If *engine* is unknown, the engine cannot write the
            requested format, column validation fails, or the sheet settings
            would leave no visible sheet.

    """
    cfg = config if config is not None else FillXLConfig()
    if not isinstance(cfg, FillXLConfig):
        msg = f"fill() argument config= must be a FillXLConfig, not {type(cfg).__name__}."
        raise TypeError(msg)
    named_values = _mapping_argument(values, "values")
    table_data = _mapping_argument(tables, "tables")

    # Each fill session gets its own "first occurrence" timezone warning.
    reset_tz_warning_state()

    template_path = Path(template)
    output_path = Path(output)
    # An unsupported format is refused before the template is looked at, so the
    # message is about the format and not a missing file.
    reject_xls_paths(template_path, output_path)
    if not template_path.exists():
        msg = f"Template file not found: {template_path}"
        raise FileNotFoundError(msg)
    if output_path.is_dir():
        msg = (
            f"Output path {output_path} is a directory; pass the file name to write "
            "the filled workbook to."
        )
        raise IsADirectoryError(msg)
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
                if cfg.active_sheet is not None:
                    backend.set_active_sheet(cfg.active_sheet)
                # Must stay inside the TemporaryDirectory block: the fillxl
                # engine streams untouched parts from the temporary copy.
                backend.save(staging_path)
            finally:
                backend.close()
        backend.check_output(template_path, staging_path, cfg.on_lossy)
        _replace_output(staging_path, output_path)
    except BaseException:
        staging_path.unlink(missing_ok=True)
        raise

    return output_path.resolve()


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
