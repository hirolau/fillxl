"""Engine helpers shared by the cross-backend tests."""

from __future__ import annotations

import contextlib
from functools import cache
from typing import TYPE_CHECKING, Any

import pytest

from fillxl import fill

if TYPE_CHECKING:
    from pathlib import Path

    from fillxl import Engine

ENGINES = [pytest.param("openpyxl", id="openpyxl"), pytest.param("win32com", id="win32com")]


@cache
def _win32com_available() -> bool:
    """Whether this machine can actually start Excel via COM.

    pywin32 can be installed on Windows CI runners that have no Excel, in
    which case ``DispatchEx`` raises a COM error that says nothing about
    Excel or win32com.  Checking up front turns those runs into skips
    instead of failures.
    """
    try:
        import pythoncom
        import win32com.client

        # A backend fill releases COM on its thread, so direct COM use must init it.
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Excel.Application")
    except Exception:
        return False

    with contextlib.suppress(Exception):
        app.Quit()
    return True


def _fill_output(
    template: Path,
    output_path: Path,
    engine: Engine,
    **kwargs: Any,  # noqa: ANN401
) -> Path:
    """Fill *template* with *engine*; skip when Excel is unavailable."""
    if engine == "win32com" and not _win32com_available():
        pytest.skip("win32com backend requires Windows and Excel")
    output = output_path.with_name(f"{output_path.stem}_{engine}{output_path.suffix}")
    fill(template=template, output=output, engine=engine, **kwargs)
    return output
