"""Backend abstraction layer for fillxl.

Defines the :class:`Backend` protocol that all Excel backends must implement,
and provides the :func:`get_backend` factory for creating backend instances.
"""

from __future__ import annotations

import importlib.util
import logging
import sys
from typing import TYPE_CHECKING, Any, Protocol, get_args, runtime_checkable

from fillxl._types import ConcreteEngine, Engine

if TYPE_CHECKING:
    from collections.abc import Sequence
    from pathlib import Path

    from fillxl._charts import TableChangeSpecs
    from fillxl._config import FillXLConfig
    from fillxl._lossy import OnLossy

logger = logging.getLogger(__name__)


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
            A :class:`TableChangeSpecs` describing the table's old and new
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

    def check_output(self, template: Path, output: Path, policy: OnLossy) -> None:
        """Apply the ``on_lossy`` policy to the saved *output* of *template*."""
        ...


_KNOWN_ENGINES: frozenset[str] = frozenset(get_args(Engine))
_OPENPYXL_SUFFIXES = frozenset({".xlsx", ".xlsm"})
_WIN32COM_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xlsb", ".xls"})


def _win32com_available() -> bool:
    return sys.platform == "win32" and importlib.util.find_spec("win32com") is not None


def resolve_engine(engine: Engine, template: Path, output: Path) -> ConcreteEngine:
    """The engine that will fill *template* into *output*.

    ``"auto"`` picks openpyxl when it can do the job: an ``.xlsx``/``.xlsm``
    template saved in the same format.  Otherwise it picks win32com when Excel
    automation is available (``.xlsb``/``.xls`` templates, format conversion).

    A format openpyxl cannot write raises here, before the template is loaded,
    instead of at ``save()`` after all the filling work is done.

    Examples:
        >>> from pathlib import Path
        >>> resolve_engine("auto", Path("t.xlsx"), Path("o.xlsx"))
        'openpyxl'
        >>> resolve_engine("win32com", Path("t.xlsx"), Path("o.xlsx"))
        'win32com'
        >>> resolve_engine("openpyxl", Path("t.xlsx"), Path("o.xlsb"))  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: The openpyxl engine cannot write a '.xlsb' file ...

    """
    if engine not in _KNOWN_ENGINES:
        known = ", ".join(sorted(_KNOWN_ENGINES))
        msg = f"Unknown engine {engine!r}. Known engines: {known}."
        raise ValueError(msg)
    if engine == "win32com":
        if output.suffix.lower() not in _WIN32COM_SUFFIXES:
            msg = (
                f"The win32com engine cannot write a '{output.suffix.lower()}' file. "
                "Use one of: .xlsx, .xlsm, .xlsb or .xls."
            )
            raise ValueError(msg)
        return engine
    if engine == "auto" and _needs_excel(template, output) and _win32com_available():
        logger.debug("engine='auto' chose win32com for %s -> %s", template, output)
        return "win32com"
    _reject_impossible_output(template, output)
    return "openpyxl"


def _needs_excel(template: Path, output: Path) -> bool:
    """Whether openpyxl cannot turn *template* into *output*."""
    suffix = template.suffix.lower()
    return suffix not in _OPENPYXL_SUFFIXES or output.suffix.lower() != suffix


def _reject_impossible_output(template: Path, output: Path) -> None:
    """Raise when openpyxl is about to be asked for a file it cannot write."""
    if not _needs_excel(template, output):
        return
    suffix = template.suffix.lower()
    if suffix not in _OPENPYXL_SUFFIXES:
        msg = (
            f"The openpyxl engine cannot read a '{suffix}' template "
            "(it reads .xlsx and .xlsm). Use engine='win32com', which needs Windows "
            "with Microsoft Excel installed, or save the template as .xlsx."
        )
        raise ValueError(msg)
    msg = (
        f"The openpyxl engine cannot write a '{output.suffix.lower()}' file from a "
        f"'{suffix}' template. Save to a '{suffix}' path, or use engine='win32com', "
        "which needs Windows with Microsoft Excel installed, to convert between formats."
    )
    raise ValueError(msg)


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
