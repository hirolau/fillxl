"""Backend abstraction layer for fillxl.

Defines the :class:`Backend` protocol that all Excel backends must implement,
and provides the :func:`get_backend` factory for creating backend instances.
"""

from __future__ import annotations

import importlib.util
import logging
import sys
from typing import TYPE_CHECKING, Any, Protocol, get_args, runtime_checkable

from fillxl._types import (
    FILLXL_SUFFIXES,
    OPENPYXL_SUFFIXES,
    WIN32COM_SUFFIXES,
    ConcreteEngine,
    Engine,
)

if TYPE_CHECKING:
    from collections.abc import Sequence
    from pathlib import Path

    from fillxl._charts import TableChangeSpecs
    from fillxl._config import FillXLConfig
    from fillxl._lossy import OnLossy

logger = logging.getLogger(__name__)


@runtime_checkable
class Backend(Protocol):
    """What a fill needs from an engine.

    Three engines implement it: ``openpyxl`` (an object model, lossy),
    ``fillxl`` (surgical edits of the OOXML package) and ``win32com`` (drives
    Microsoft Excel).  Each one turns :func:`fillxl.fill`'s steps into that
    engine's own operations; the fill itself decides what to write.
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
            dimensions.  The charts that follow the table have already been
            updated by the time this returns.

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
_CONCRETE_ENGINES: frozenset[str] = frozenset(get_args(ConcreteEngine))


def _win32com_available() -> bool:
    """Whether this machine can actually start Excel through COM.

    pywin32 is installed on every Windows install of fillxl; Excel is not.
    A build server or CI runner has the package and no ``Excel.Application``
    ProgID, so checking only for the package makes ``engine="auto"`` pick
    win32com for a fill that cannot open a workbook.  The ProgID lookup is
    cheap enough to run on every ``"auto"`` decision that needs Excel, and it
    does not start Excel just to ask.
    """
    if sys.platform != "win32" or importlib.util.find_spec("win32com") is None:
        return False
    import winreg

    # typeshed lists winreg's members for a win32 target only, so checkers that
    # run on Linux reject the direct lookup; the alias keeps this file clean on
    # both platforms.
    registry: Any = winreg
    try:
        with registry.OpenKey(registry.HKEY_CLASSES_ROOT, "Excel.Application"):
            return True
    except OSError:
        return False


_XLS_MESSAGE = (
    "fillxl does not support .xls (Excel 97-2003) files. Open the file in Excel and "
    "save it as .xlsx, .xlsm or .xlsb."
)


def _require_known_engine(engine: str) -> None:
    """Raise unless *engine* is one of the names :data:`Engine` allows.

    Examples:
        >>> _require_known_engine("openpyxl")
        >>> _require_known_engine("bogus")
        Traceback (most recent call last):
            ...
        ValueError: Unknown engine 'bogus'. Known engines: auto, fillxl, openpyxl, win32com.

    """
    if engine in _KNOWN_ENGINES:
        return
    known = ", ".join(sorted(_KNOWN_ENGINES))
    msg = f"Unknown engine {engine!r}. Known engines: {known}."
    raise ValueError(msg)


def reject_xls_paths(*paths: Path) -> None:
    """Raise when any of *paths* is a ``.xls`` file.

    BIFF8 cannot hold the features fillxl exists to preserve, and Excel is
    needed to read it, so the request is refused before anything is opened.

    Examples:
        >>> from pathlib import Path
        >>> reject_xls_paths(Path("t.xlsx"), Path("o.xlsx"))
        >>> reject_xls_paths(Path("t.xls"))  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: fillxl does not support .xls (Excel 97-2003) files. ... t.xls.

    """
    for path in paths:
        if path.suffix.lower() == ".xls":
            raise ValueError(f"{_XLS_MESSAGE} Offending path: {path}.")


def resolve_engine(engine: Engine, template: Path, output: Path) -> ConcreteEngine:
    """The engine that will fill *template* into *output*.

    ``"auto"`` picks openpyxl when it can do the job: an ``.xlsx``/``.xlsm``
    template saved in the same format.  Otherwise it picks win32com when Excel
    automation is available (``.xlsb`` templates, format conversion).

    ``"fillxl"`` edits the OOXML package itself: no Excel, and no feature loss.
    It reads and writes ``.xlsx``/``.xlsm`` only, in the same format.

    A format an engine cannot write raises here, before the template is loaded,
    instead of at ``save()`` after all the filling work is done.

    Examples:
        >>> from pathlib import Path
        >>> resolve_engine("auto", Path("t.xlsx"), Path("o.xlsx"))
        'openpyxl'
        >>> resolve_engine("win32com", Path("t.xlsx"), Path("o.xlsx"))
        'win32com'
        >>> resolve_engine("fillxl", Path("t.xlsx"), Path("o.xlsx"))
        'fillxl'
        >>> resolve_engine("fillxl", Path("t.xlsx"), Path("o.xlsb"))  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: The fillxl engine cannot write a '.xlsb' file ...
        >>> resolve_engine("openpyxl", Path("t.xlsx"), Path("o.xlsb"))  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: The openpyxl engine cannot write a '.xlsb' file ...

    """
    _require_known_engine(engine)
    reject_xls_paths(template, output)
    if engine == "win32com":
        _require_excel_output(output)
        return engine
    if engine == "fillxl":
        _require_fillxl_output(template, output)
        return engine
    if engine == "auto" and _needs_excel(template, output) and _win32com_available():
        # The suffix is checked here too: "auto" must not start Excel for a
        # fill that save() is bound to reject.
        _require_excel_output(output, engine="auto")
        logger.debug("engine='auto' chose win32com for %s -> %s", template, output)
        return "win32com"
    _reject_impossible_output(template, output)
    return "openpyxl"


def _require_excel_output(output: Path, *, engine: str = "win32com") -> None:
    """Raise unless Excel can save *output*.

    *engine* only shapes the message: for ``"auto"`` the caller did not ask for
    win32com, so the error says that the choice was made for them.

    Examples:
        >>> from pathlib import Path
        >>> _require_excel_output(Path("out.xlsb"))
        >>> _require_excel_output(Path("out.csv"))  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: The win32com engine cannot write a '.csv' file. ...
        >>> _require_excel_output(Path("out"), engine="auto")  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: engine='auto' chose win32com, which cannot write a file without ...

    """
    suffix = output.suffix.lower()
    if suffix in WIN32COM_SUFFIXES:
        return
    what = f"a '{suffix}' file" if suffix else "a file without an extension"
    if engine == "auto":
        msg = (
            f"engine='auto' chose win32com, which cannot write {what}. Use one of: "
            ".xlsx, .xlsm or .xlsb, or name an .xlsx output to stay on openpyxl."
        )
    else:
        msg = f"The win32com engine cannot write {what}. Use one of: .xlsx, .xlsm or .xlsb."
    raise ValueError(msg)


def _require_fillxl_output(template: Path, output: Path) -> None:
    """Raise unless the fillxl engine can turn *template* into *output*.

    Examples:
        >>> from pathlib import Path
        >>> _require_fillxl_output(Path("t.xlsx"), Path("o.xlsx"))
        >>> _require_fillxl_output(Path("t.xlsx"), Path("o.xlsb"))  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: The fillxl engine cannot write a '.xlsb' file...
        >>> _require_fillxl_output(Path("t.xlsb"), Path("o.xlsb"))  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: The fillxl engine cannot read a '.xlsb' template...

    """
    template_suffix = template.suffix.lower()
    output_suffix = output.suffix.lower()
    if template_suffix not in FILLXL_SUFFIXES:
        msg = (
            f"The fillxl engine cannot read a '{template_suffix}' template "
            "(it reads .xlsx and .xlsm). Use engine='win32com' for .xlsb, "
            "or save the template as .xlsx."
        )
        raise ValueError(msg)
    if output_suffix != template_suffix:
        what = f"a '{output_suffix}' file" if output_suffix else "a file without an extension"
        msg = (
            f"The fillxl engine cannot write {what} from a '{template_suffix}' "
            f"template. Save to a '{template_suffix}' path, or use engine='win32com' "
            "to convert between formats."
        )
        raise ValueError(msg)


def _needs_excel(template: Path, output: Path) -> bool:
    """Whether openpyxl cannot turn *template* into *output*."""
    suffix = template.suffix.lower()
    return suffix not in OPENPYXL_SUFFIXES or output.suffix.lower() != suffix


def _reject_impossible_output(template: Path, output: Path) -> None:
    """Raise when openpyxl is about to be asked for a file it cannot write."""
    if not _needs_excel(template, output):
        return
    suffix = template.suffix.lower()
    if suffix not in OPENPYXL_SUFFIXES:
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


def get_backend(engine: ConcreteEngine) -> Backend:
    """Return the :class:`Backend` implementation for *engine*.

    Args:
        engine: A concrete engine name: ``"openpyxl"``, ``"fillxl"`` or
            ``"win32com"``.  ``"auto"`` is turned into one of these by
            :func:`resolve_engine` before a fill builds a backend.

    Returns:
        A backend instance ready to load workbooks.

    Raises:
        ValueError: If the engine name is unknown, or is ``"auto"``.

    """
    _require_known_engine(engine)
    if engine not in _CONCRETE_ENGINES:
        msg = (
            "get_backend() needs a concrete engine; 'auto' is resolved by "
            "resolve_engine() before a backend is created."
        )
        raise ValueError(msg)

    if engine == "openpyxl":
        from fillxl._backends._openpyxl import OpenpyxlBackend

        return OpenpyxlBackend()

    if engine == "win32com":
        from fillxl._backends._win32com import Win32ComBackend

        return Win32ComBackend()

    if engine == "fillxl":
        from fillxl._backends._fillxl import FillxlBackend

        return FillxlBackend()

    # All concrete engine names are handled above — this line is unreachable.
    msg = f"Unhandled engine {engine!r}"
    raise AssertionError(msg)  # pragma: no cover
