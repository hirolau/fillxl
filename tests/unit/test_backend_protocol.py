"""Backend protocol conformance tests.

All backends conform to the same protocol (load, fill, save, close, …).
These tests verify the Backend protocol interface contract — they belong
in unit tests, not story tests, because they describe architecture, not
user behavior.
"""

from __future__ import annotations


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
        from fillxl._backend import Backend
        from fillxl._backends._openpyxl import OpenpyxlBackend

        backend = OpenpyxlBackend()
        assert isinstance(backend, Backend)
