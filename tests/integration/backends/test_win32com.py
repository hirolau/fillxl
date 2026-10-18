"""Win32com backend unit tests — CI-safe guards (no Excel required).

Tests the no-workbook guard, factory, and platform checks for the
win32com backend.  These do not require Excel and can run in CI.
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pytest

from fillxl import FillXLConfig


class TestWin32ComBackendNoWorkbook:
    """Backend method calls without loading a workbook (CI-safe)."""

    @pytest.mark.parametrize(
        "method_call",
        [
            pytest.param(
                lambda b: b.fill_named_cell("test", "value", FillXLConfig()),  # pyrefly: ignore[implicit-any-lambda]
                id="fill_named_cell",
            ),
            pytest.param(
                lambda b: b.fill_table("test", [], FillXLConfig()),  # pyrefly: ignore[implicit-any-lambda]
                id="fill_table",
            ),
            pytest.param(
                lambda b: b.save(Path("output.xlsx")),  # pyrefly: ignore[implicit-any-lambda]
                id="save",
            ),
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
            method_call(backend)  # type: ignore[operator]  # ty:ignore[call-non-callable]

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
        assert backend._app is None
        assert backend._workbook is None


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


def _skip_without_excel() -> None:
    """Skip the calling test when this machine cannot drive Excel."""
    if sys.platform != "win32":
        pytest.skip("win32com backend requires Windows")

    try:
        import win32com.client

        app = win32com.client.DispatchEx("Excel.Application")
    except Exception as exc:
        pytest.skip(f"Excel COM is not available: {exc}")
    else:
        app.Quit()


class TestWin32ComTotalsRows:
    """Tables with totals rows are rejected with a clear message."""

    def test_a_table_with_a_totals_row_is_rejected(self, tmp_path: Path) -> None:
        """Totals rows are not supported — the error says so instead of corrupting them."""
        from openpyxl.worksheet.table import Table, TableColumn

        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "totals_template.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Name"
        ws["B1"] = "Value"
        ws["A2"] = "a"
        ws["B2"] = 1
        ws["A3"] = "Total"
        ws["B3"] = "=SUBTOTAL(109,B2:B2)"
        table = Table(
            displayName="Totals",
            ref="A1:B3",
            tableColumns=[
                TableColumn(id=1, name="Name", totalsRowLabel="Total"),
                TableColumn(id=2, name="Value", totalsRowFunction="sum"),
            ],
            totalsRowCount=1,
            totalsRowShown=True,
        )
        ws.add_table(table)
        wb.save(template)
        wb.close()

        # Excel only exposes ShowTotals for real list objects, so this case is
        # exercised through the win32com backend as well as openpyxl.
        with pytest.raises(ValueError, match="totals row"):
            fill(
                template=template,
                output=tmp_path / "out.xlsx",
                tables={"Totals": [{"Name": "b", "Value": 2}]},
                engine="win32com",
            )
