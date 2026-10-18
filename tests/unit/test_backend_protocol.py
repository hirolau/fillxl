"""Backend protocol conformance tests.

All backends conform to the same protocol (load, fill, save, close, …).
These tests verify the Backend protocol interface contract — they belong
in unit tests, not story tests, because they describe architecture, not
user behavior.
"""

from __future__ import annotations

import pytest

from fillxl._backend import Backend
from fillxl._backends._openpyxl import OpenpyxlBackend
from fillxl._charts import TableChangeSpecs


class TestBackendProtocol:
    """All backends conform to the same protocol (load, fill, save, close, …)."""

    def test_backend_protocol_defines_the_full_interface(self) -> None:
        """The Backend protocol has load, fill_named_cell, fill_table,
        update_charts, set_pivot_refresh, hide_sheets, delete_sheets,
        set_active_sheet, save, and close.
        """
        from fillxl._backend import Backend

        expected_methods = {
            "load",
            "fill_named_cell",
            "fill_table",
            "update_charts",
            "set_pivot_refresh",
            "hide_sheets",
            "delete_sheets",
            "set_active_sheet",
            "save",
            "close",
        }
        # Get abstract methods from the protocol
        protocol_methods = {
            name
            for name, _ in Backend.__dict__.items()
            if not name.startswith("_") and callable(getattr(Backend, name, None))
        }
        assert expected_methods <= protocol_methods, (
            f"Missing methods: {expected_methods - protocol_methods}"
        )

    def test_openpyxl_backend_is_runtime_checkable_as_backend(self) -> None:
        """Isinstance check works against the Backend protocol."""
        backend = OpenpyxlBackend()
        assert isinstance(backend, Backend)


class TestOpenpyxlBackendRuntimeErrorGuards:
    """Calling backend methods before load() raises RuntimeError."""

    def test_fill_named_cell_raises_when_not_loaded(self) -> None:
        """Calling fill_named_cell before load raises RuntimeError."""
        from fillxl._config import FillXLConfig

        backend = OpenpyxlBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            backend.fill_named_cell("test", 42, FillXLConfig())

    def test_fill_table_raises_when_not_loaded(self) -> None:
        """Calling fill_table before load raises RuntimeError."""
        from fillxl._config import FillXLConfig

        backend = OpenpyxlBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            backend.fill_table("test", [], FillXLConfig())

    def test_update_charts_raises_when_not_loaded(self) -> None:
        """Calling update_charts before load raises RuntimeError."""
        backend = OpenpyxlBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            backend.update_charts(TableChangeSpecs("Sheet1", 1, 1, 2, 2, 3))

    def test_set_pivot_refresh_raises_when_not_loaded(self) -> None:
        """Calling set_pivot_refresh before load raises RuntimeError."""
        backend = OpenpyxlBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            backend.set_pivot_refresh()

    def test_hide_sheets_raises_when_not_loaded(self) -> None:
        """Calling hide_sheets before load raises RuntimeError."""
        backend = OpenpyxlBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            backend.hide_sheets("_")

    def test_delete_sheets_raises_when_not_loaded(self) -> None:
        """Calling delete_sheets before load raises RuntimeError."""
        backend = OpenpyxlBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            backend.delete_sheets(["Sheet1"])

    def test_set_active_sheet_raises_when_not_loaded(self) -> None:
        """Calling set_active_sheet before load raises RuntimeError."""
        backend = OpenpyxlBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            backend.set_active_sheet("Sheet1")
