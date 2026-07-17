"""Win32com backend unit tests — CI-safe guards (no Excel required).

Tests the no-workbook guard, factory, and platform checks for the
win32com backend.  These do not require Excel and can run in CI.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from fillxl import FillXLConfig


class TestWin32ComBackendNoWorkbook:
    """Backend method calls without loading a workbook (CI-safe)."""

    @pytest.mark.parametrize(
        "method_call",
        [
            pytest.param(lambda b: b.fill_named_cell("test", "value"), id="fill_named_cell"),
            pytest.param(lambda b: b.fill_table("test", [], FillXLConfig()), id="fill_table"),
            pytest.param(lambda b: b.save(Path("output.xlsx")), id="save"),
        ],
    )
    def test_win32com_backend_raises_runtime_error_without_workbook(
        self,
        method_call: object,
    ) -> None:
        """Any backend operation without loading raises RuntimeError."""
        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            method_call(backend)  # ty:ignore[call-non-callable]

    def test_win32com_backend_close_without_load_does_not_raise(self) -> None:
        """Calling close() without loading should not raise."""
        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        backend.close()


class TestWin32ComBackendFactory:
    """Factory tests for the win32com engine (CI-safe)."""

    def test_get_backend_returns_win32com_for_win32com_engine(self) -> None:
        """get_backend('win32com') should return a Win32ComBackend instance."""
        from fillxl._backend import get_backend
        from fillxl._backends._win32com import Win32ComBackend

        backend = get_backend("win32com")
        assert isinstance(backend, Win32ComBackend)

    def test_win32com_backend_does_not_launch_excel_on_init(self) -> None:
        """Win32ComBackend() should not launch Excel until load() is called."""
        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        assert backend._app is None  # type: ignore[attr-defined]
        assert backend._workbook is None  # type: ignore[attr-defined]


class TestWin32ComPlatformGuard:
    """Platform guard for non-Windows systems (CI-safe)."""

    def test_win32com_backend_raises_runtime_error_on_non_windows(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """On non-Windows, load() should raise RuntimeError with a clear message."""
        from fillxl._backends import _win32com

        monkeypatch.setattr(_win32com, "_is_windows", False)

        backend = _win32com.Win32ComBackend()
        with pytest.raises(RuntimeError, match="win32com"):
            backend.load(tmp_path / "fake.xlsx")
