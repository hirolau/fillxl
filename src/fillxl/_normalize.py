"""Value normalisation for Excel compatibility.

Both the ``openpyxl`` and ``win32com`` backends route ALL cell values
through :func:`normalize_value` so that they produce **identical results**
for the same input data.

Normalisation rules (applied in order):
    1. ``pd.NA`` / ``pd.NaT`` → ``None`` (pandas sentinels — cannot be
       represented in Excel).
    2. Empty strings → ``None``.  openpyxl cannot round-trip an empty text
       cell (it reads back as ``None``), so both backends write an empty
       cell instead of storing text that only one engine can represent.
    3. NumPy scalars (``np.int64``, ``np.float32``, ``np.bool_``,
       ``np.datetime64``, …) → the equivalent native Python value.
       Excel stores anything it does not recognise as text, which is how
       numpy scalars used to end up with the "number stored as text"
       marker.
    4. ``float('nan')``, ``float('inf'), float('-inf')`` → ``None``
       (Excel cannot represent IEEE-754 special values).
    5. ``datetime.date`` → ``datetime.datetime`` (Excel has no native
       ``date`` type — everything is a serial number).
    6. ``datetime.time`` → ``float`` (fraction of a 24-hour day), which is
       Excel's own time representation, so arithmetic such as ``=A2*24``
       works.  The template's number format controls the display.
    7. ``datetime.timedelta`` → ``float`` (fraction of a 24-hour day).
       Excel has no native ``timedelta`` type.
    8. ``Decimal`` → ``float``.  Excel has no native ``Decimal`` type.
    9. Timezone-aware ``datetime`` → naive ``datetime`` (tzinfo stripped,
       wall-clock preserved).  ``pd.Timestamp`` (subclass of ``datetime``)
       is handled identically.
   10. Unrecognised objects → ``str(obj)`` via ``__str__``.
"""

from __future__ import annotations

import datetime as dt
import math
import sys
import warnings
from contextvars import ContextVar
from decimal import Decimal
from typing import Any, TypeVar

T = TypeVar("T")

# ── Per-fill-session state ──────────────────────────────────────────────────
# A ContextVar (rather than a module global) so that two concurrent fills in
# different threads each get their own "first occurrence" warning.
_tz_warning_issued: ContextVar[bool] = ContextVar("fillxl_tz_warning_issued", default=False)


def reset_tz_warning_state() -> None:
    """Reset the timezone warning state for a new fill session.

    Call this at the start of :func:`fillxl.fill` so that each fill
    session gets its own "first occurrence" warning.

    Examples:
        >>> reset_tz_warning_state()
        >>> _tz_warning_issued.get()
        False

    """
    _tz_warning_issued.set(False)


def normalize_datetime(value: T) -> T:
    """Strip timezone info from tz-aware datetimes, warning on first occurrence.

    Excel has no native timezone support.  Datetimes are stored as
    floating-point serial numbers with no timezone metadata.  Any
    tz-aware value must be converted to naive at write time.

    Args:
        value: Any Python value.  Non-datetime values are returned
            unchanged.

    Returns:
        The value with timezone info stripped, or the original value
        if it is not a tz-aware datetime.

    Examples:
        >>> import datetime as dt
        >>> reset_tz_warning_state()
        >>> naive = dt.datetime(2025, 6, 15, 14, 30)
        >>> normalize_datetime(naive) is naive
        True
        >>> aware = dt.datetime(2025, 6, 15, 14, 30, tzinfo=dt.timezone.utc)
        >>> with warnings.catch_warnings():
        ...     warnings.simplefilter("ignore")
        ...     result = normalize_datetime(aware)
        >>> result.tzinfo is None
        True
        >>> result == dt.datetime(2025, 6, 15, 14, 30)
        True

    """
    if isinstance(value, dt.datetime) and value.tzinfo is not None:
        if not _tz_warning_issued.get():
            warnings.warn(
                "Timezone-aware datetimes were provided. Excel does not "
                "support timezones — the tzinfo will be stripped.",
                UserWarning,
                stacklevel=2,
            )
            _tz_warning_issued.set(True)
        return value.replace(tzinfo=None)

    return value


# ==============================================================================
# General value normalisation — single entry point for ALL backends
# ==============================================================================


def normalize_value(value: Any) -> Any:  # noqa: ANN401
    """Normalise *any* Python value for Excel compatibility.

    This is the **single entry point** for all cell values in both
    backends.  It guarantees identical output regardless of backend.

    Rules (applied in order):

    1. ``pd.NA`` / ``pd.NaT`` → ``None`` (pandas sentinels).
    2. Empty strings → ``None`` (openpyxl cannot round-trip them).
    3. NumPy scalars → native Python values.
    4. IEEE-754 special floats (NaN, Inf, -Inf) → ``None``.
    5. ``date`` → ``datetime`` (Excel has no native ``date`` type).
    6. ``time`` → ``float`` (fraction of a day).
    7. ``timedelta`` → ``float`` (fraction of a 24-hour day).
    8. ``Decimal`` → ``float``.
    9. Timezone-aware ``datetime`` → naive ``datetime``.
    10. Unrecognised objects → ``str(obj)``.

    Args:
        value: Any Python value to write to an Excel cell.

    Returns:
        A value suitable for writing to an Excel cell.

    Examples:
        >>> normalize_value("hello")
        'hello'
        >>> normalize_value(42)
        42
        >>> normalize_value(3.14)
        3.14
        >>> normalize_value(True)
        True
        >>> normalize_value(None) is None
        True
        >>> normalize_value("") is None
        True
        >>> import math
        >>> normalize_value(float('nan')) is None
        True
        >>> normalize_value(float('inf')) is None
        True
        >>> from datetime import date, datetime, time, timedelta
        >>> normalize_value(date(2025, 12, 25))
        datetime.datetime(2025, 12, 25, 0, 0)
        >>> round(normalize_value(time(9, 15, 30)) * 24, 6)
        9.258333
        >>> normalize_value(timedelta(hours=6))
        0.25
        >>> from decimal import Decimal
        >>> normalize_value(Decimal('12.34'))
        12.34
        >>> class Greeting:
        ...     def __str__(self) -> str:
        ...         return "Hello!"
        >>> normalize_value(Greeting())
        'Hello!'

    """
    # ── 1. pd.NA / pd.NaT → None (before any other type checks) ───────────
    pd = sys.modules.get("pandas")
    if pd is not None and (value is pd.NA or value is pd.NaT):
        return None

    # ── 2. Empty string → None ────────────────────────────────────────────
    # openpyxl reads an empty text cell back as None, so writing "" would
    # make the two backends disagree.  Both write a blank cell instead.
    if isinstance(value, str) and not value:
        return None

    # ── 3. NumPy scalars → native Python ──────────────────────────────────
    # np.int64, np.float32, np.bool_, np.datetime64 … do not subclass the
    # corresponding Python types, so without this they would be written as
    # text (Excel then flags the cell with "number stored as text").
    if _is_numpy_scalar(value):
        value = _numpy_scalar_to_python(value)

    # ── 4. IEEE-754 special floats → None ─────────────────────────────────
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None

    # ── 5. date → datetime ────────────────────────────────────────────────
    if isinstance(value, dt.date) and not isinstance(value, dt.datetime):
        value = dt.datetime(value.year, value.month, value.day)

    # ── 6. time → fraction of a day (Excel's own representation) ──────────
    if isinstance(value, dt.time):
        return (
            value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
        ) / 86_400

    # ── 7. timedelta → float (fraction of a day) ──────────────────────────
    if isinstance(value, dt.timedelta):
        value = value.total_seconds() / 86400.0

    # ── 8. Decimal → float ────────────────────────────────────────────────
    if isinstance(value, Decimal):
        value = float(value)

    # ── 9. Strip timezone from datetime (must come after date→datetime) ───
    value = normalize_datetime(value)

    # ── 10. Recognised Excel-compatible primitives pass through unchanged ─
    if value is None or isinstance(
        value,
        (str, int, float, bool, dt.datetime),
    ):
        return value

    # ── 11. Unrecognised → string ────────────────────────────────────────
    return str(value)


def _is_numpy_scalar(value: Any) -> bool:  # noqa: ANN401
    """Whether *value* is a NumPy scalar such as ``np.int64`` or ``np.float32``.

    NumPy is looked up in :data:`sys.modules` rather than imported: a NumPy
    scalar can only exist if NumPy has already been imported, and importing
    it per cell is measurably slow.

    Examples:
        >>> _is_numpy_scalar(42)
        False
        >>> _is_numpy_scalar("42")
        False

    """
    numpy = sys.modules.get("numpy")
    if numpy is None:
        return False
    if isinstance(value, numpy.generic):
        return True
    # 0-d arrays (``np.array(5)``) behave like scalars.
    return isinstance(value, numpy.ndarray) and value.ndim == 0


def _numpy_scalar_to_python(value: Any) -> Any:  # noqa: ANN401
    """Convert a NumPy scalar to the closest native Python object."""
    numpy = sys.modules.get("numpy")
    if numpy is None:  # pragma: no cover - only called for NumPy values
        return value.item()
    item = value.item()

    if isinstance(item, int) and not isinstance(item, bool):
        # ``.item()`` returns a raw integer for datetime64/timedelta64 at
        # ns/ps/fs/as resolution.  Let NumPy convert the unit to
        # microseconds instead of assuming the count is in nanoseconds.
        if isinstance(value, numpy.datetime64):
            return value.astype("datetime64[us]").item()
        if isinstance(value, numpy.timedelta64):
            return value.astype("timedelta64[us]").item()

    if (
        isinstance(item, numpy.generic)
        and hasattr(item, "dtype")
        # ``np.longdouble.item()`` hands back another NumPy scalar, but Excel
        # wants a plain float for real numbers.
        and numpy.issubdtype(item.dtype, numpy.floating)
    ):
        return float(item)

    return item


# ==============================================================================
# Backend-facing helpers
#
# ``prepare_cell_value`` and ``prepare_com_value`` are the two entry points the
# backends use.  Keeping them here guarantees that both backends see exactly
# the same value for the same input — including the win32com quirk that COM
# applies the local timezone offset to naive datetimes.
# ==============================================================================


def _to_python_datetime(value: Any) -> Any:  # noqa: ANN401
    """Convert ``pd.Timestamp`` to ``datetime`` (COM cannot marshal it)."""
    pd = sys.modules.get("pandas")
    if pd is not None and isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def is_formula_string(value: Any) -> bool:  # noqa: ANN401
    """Whether *value* is a string that Excel would treat as a formula.

    Examples:
        >>> is_formula_string("=SUM(A1:A2)")
        True
        >>> is_formula_string("00123")
        False
        >>> is_formula_string(123)
        False

    """
    return isinstance(value, str) and value.startswith("=")


def is_text_cell_value(value: Any, *, escape_formulas: bool = False) -> bool:  # noqa: ANN401
    """Whether *value* must be forced to the text cell type in openpyxl.

    openpyxl treats a string starting with ``=`` as a formula and a string
    such as ``"#N/A"`` as an error value.  Forcing the remaining strings to
    text keeps the openpyxl output identical to win32com, which stores every
    string as text.

    Args:
        value: The normalised value about to be written.
        escape_formulas: If True, formula-like strings are text as well.

    Examples:
        >>> is_text_cell_value("hello")
        True
        >>> is_text_cell_value("#N/A")
        True
        >>> is_text_cell_value("=1+1")
        False
        >>> is_text_cell_value("=1+1", escape_formulas=True)
        True
        >>> is_text_cell_value(42)
        False

    """
    if not isinstance(value, str):
        return False
    return escape_formulas or not is_formula_string(value)


def prepare_cell_value(value: Any) -> Any:  # noqa: ANN401
    """Prepare *value* for writing through the openpyxl backend.

    Args:
        value: Raw value from user data.

    Returns:
        A value that openpyxl can store without loss.

    Examples:
        >>> import datetime as dt
        >>> prepare_cell_value(dt.date(2025, 1, 2))
        datetime.datetime(2025, 1, 2, 0, 0)

    """
    return normalize_value(_to_python_datetime(value))


def prepare_com_value(value: Any, *, escape_formulas: bool = False) -> Any:  # noqa: ANN401
    """Prepare *value* for writing through the win32com backend.

    Applies the same normalisation as :func:`prepare_cell_value`, then takes
    care of two COM-specific behaviours:

    * Datetimes are marked as UTC.  COM converts naive datetimes using the
      local timezone, which shifts the serial number Excel stores; marking
      the value as UTC makes that serial match the wall-clock time the
      openpyxl backend writes.
    * Strings are prefixed with an apostrophe, Excel's marker for "store this
      as text".  Without it COM parses strings the way typed input is parsed:
      ``"00123"`` becomes the number 123 and ``"1/2"`` becomes a date, while
      openpyxl keeps them as text.  Strings starting with ``=`` are left alone
      so they stay formulas, unless *escape_formulas* is set.

    Args:
        value: Raw value from user data.
        escape_formulas: If True, strings starting with ``=`` are stored as
            literal text instead of becoming formulas.

    Returns:
        A value suitable for assigning to a COM ``Range``.

    Examples:
        >>> import datetime as dt
        >>> prepare_com_value(3.14)
        3.14
        >>> prepare_com_value("00123")
        "'00123"
        >>> prepare_com_value("=1+1")
        '=1+1'
        >>> prepare_com_value("=1+1", escape_formulas=True)
        "'=1+1"
        >>> prepare_com_value(dt.datetime(2025, 6, 15, 14, 30)).isoformat()
        '2025-06-15T14:30:00+00:00'

    """
    value = normalize_value(_to_python_datetime(value))

    if isinstance(value, str):
        if escape_formulas or not value.startswith("="):
            return "'" + value
        return value

    if isinstance(value, dt.datetime):
        return value.replace(tzinfo=dt.UTC)

    return value
