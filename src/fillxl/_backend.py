"""Backend abstraction layer for fillxl.

Defines the :class:`Backend` protocol that all Excel backends must implement,
and provides the :func:`get_backend` factory for creating backend instances.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Protocol, get_args, runtime_checkable

from fillxl._types import Engine

if TYPE_CHECKING:
    from collections.abc import Sequence
    from pathlib import Path

    from fillxl._charts import TableChangeSpecs
    from fillxl._config import FillXLConfig


@runtime_checkable
class Backend(Protocol):
    """Protocol that all Excel backends must implement.

    This is the seam where future backends (com32, xlwings, etc.) will plug in.
    """

    def load(self, path: Path) -> None:
        """Open an Excel workbook from *path*.

        Args:
            path: Path to the workbook file.

        """
        ...

    def fill_named_cell(self, name: str, value: Any, config: FillXLConfig) -> None:  # noqa: ANN401
        """Write a scalar *value* to the named cell *name*.

        Args:
            name: Name of the named cell / defined name.
            value: Scalar value to write.
            config: Global fill configuration.

        """
        ...

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

        """
        ...

    def update_charts(self, spec: TableChangeSpecs) -> None:
        """Update chart data ranges after a table resize.

        Args:
            spec: The table dimension change to reflect in charts.

        """
        ...

    def set_pivot_refresh(self) -> None:
        """Mark all pivot caches for refresh on load."""
        ...

    def hide_sheets(self, prefix: str) -> None:
        """Hide worksheets whose names start with *prefix*.

        Args:
            prefix: Sheet name prefix (e.g. ``"_"``).

        """
        ...

    def delete_sheets(self, names: Sequence[str]) -> None:
        """Delete the worksheets named in *names*.

        Args:
            names: Sheet names to remove.

        """
        ...

    def set_active_sheet(self, name: str) -> None:
        """Set the active (displayed) worksheet.

        Args:
            name: Name of the worksheet to activate.

        """
        ...

    def save(self, path: Path) -> None:
        """Save the workbook to *path*.

        Args:
            path: Destination path.

        """
        ...

    def close(self) -> None:
        """Close the workbook and release resources."""
        ...


_KNOWN_ENGINES: frozenset[str] = frozenset(get_args(Engine))


def get_backend(engine: Engine) -> Backend:
    """Return a :class:`Backend` implementation for *engine*.

    Args:
        engine: Engine name (``"auto"``, ``"openpyxl"``, or ``"win32com"``).

    Returns:
        A backend instance ready to load workbooks.

    Raises:
        NotImplementedError: If the engine is not implemented yet.
        ValueError: If the engine name is unknown.

    """
    if engine not in _KNOWN_ENGINES:
        known = ", ".join(sorted(_KNOWN_ENGINES))
        msg = f"Unknown engine {engine!r}. Known engines: {known}."
        raise ValueError(msg)

    if engine in ("auto", "openpyxl"):
        from fillxl._backends._openpyxl import OpenpyxlBackend

        return OpenpyxlBackend()

    if engine == "win32com":
        from fillxl._backends._win32com import Win32ComBackend

        return Win32ComBackend()

    # All known engine names are handled above — this line is unreachable.
    msg = f"Unhandled engine {engine!r}"
    raise AssertionError(msg)  # pragma: no cover
