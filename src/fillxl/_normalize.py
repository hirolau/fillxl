"""Value normalisation for Excel compatibility.

Both the ``openpyxl`` and ``win32com`` backends route ALL cell values
through :func:`normalize_value` so that they produce **identical results**
for the same input data.

Normalisation rules (applied in order):
    1. ``pd.NA`` / ``pd.NaT`` → ``None`` (pandas sentinels — cannot be
       represented in Excel).
    2. ``float('nan')``, ``float('inf'), float('-inf')`` → ``None``
       (Excel cannot represent IEEE-754 special values).
    3. ``datetime.date`` → ``datetime.datetime`` (Excel has no native
       ``date`` type — everything is a serial number).
    4. ``datetime.time`` → ``datetime.datetime`` on ``2000-01-01``
       (COM cannot marshal raw ``time`` objects, so they are promoted to
       ``datetime`` for cross-backend consistency).  ``2000-01-01`` is a
       universally COM-compatible anchor date.
    5. ``datetime.timedelta`` → ``float`` (fraction of a 24-hour day).
       Excel has no native ``timedelta`` type.
    6. ``Decimal`` → ``float``.  Excel has no native ``Decimal`` type.
    7. Timezone-aware ``datetime`` → naive ``datetime`` (tzinfo stripped,
       wall-clock preserved).  ``pd.Timestamp`` (subclass of ``datetime``)
       is handled identically.
    8. Unrecognised objects → ``str(obj)`` via ``__str__``.
"""

from __future__ import annotations

import datetime as dt
import math
import warnings
from decimal import Decimal
from typing import Any, TypeVar

T = TypeVar("T")

# ── Per-fill-session state ──────────────────────────────────────────────────
_tz_warning_issued: bool = False


def reset_tz_warning_state() -> None:
    """Reset the timezone warning state for a new fill session.

    Call this at the start of :func:`fillxl.fill` so that each fill
    session gets its own "first occurrence" warning.

    Examples:
        >>> reset_tz_warning_state()
        >>> _tz_warning_issued
        False

    """
    global _tz_warning_issued
    _tz_warning_issued = False


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
    global _tz_warning_issued

    if isinstance(value, dt.datetime) and value.tzinfo is not None:
        if not _tz_warning_issued:
            warnings.warn(
                "Timezone-aware datetimes were provided. Excel does not "
                "support timezones — the tzinfo will be stripped.",
                UserWarning,
                stacklevel=2,
            )
            _tz_warning_issued = True
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
    2. IEEE-754 special floats (NaN, Inf, -Inf) → ``None``.
    3. ``date`` → ``datetime`` (Excel has no native ``date`` type).
    4. ``time`` → ``datetime`` on the Excel epoch (``1899-12-30``).
    5. ``timedelta`` → ``float`` (fraction of a 24-hour day).
    6. ``Decimal`` → ``float``.
    7. Timezone-aware ``datetime`` → naive ``datetime``.
    8. Unrecognised objects → ``str(obj)``.

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
        >>> import math
        >>> normalize_value(float('nan')) is None
        True
        >>> normalize_value(float('inf')) is None
        True
        >>> from datetime import date, datetime, time, timedelta
        >>> normalize_value(date(2025, 12, 25))
        datetime.datetime(2025, 12, 25, 0, 0)
        >>> normalize_value(time(9, 15, 30))
        datetime.datetime(2000, 1, 1, 9, 15, 30)
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
    try:
        import pandas as pd

        if value is pd.NA or value is pd.NaT:
            return None
    except ImportError:
        pass

    # ── 2. IEEE-754 special floats → None ──────────────────────────────────
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None

    # ── 3. date → datetime ─────────────────────────────────────────────────
    if isinstance(value, dt.date) and not isinstance(value, dt.datetime):
        value = dt.datetime(value.year, value.month, value.day)

    # ── 4. time → datetime on anchor date (2000-01-01) ────────────────────
    #     COM cannot marshal raw time objects, so they are promoted to
    #     datetime for cross-backend consistency.  2000-01-01 is used as
    #     the anchor date because it is universally COM-compatible (some
    #     Windows/pywin32 versions cannot marshal dates before ~1969).
    #     The time fraction (hours:minutes:seconds) is preserved.
    if isinstance(value, dt.time) and not isinstance(value, dt.datetime):
        value = dt.datetime(2000, 1, 1, value.hour, value.minute, value.second, value.microsecond)

    # ── 5. timedelta → float (fraction of a day) ───────────────────────────
    if isinstance(value, dt.timedelta):
        value = value.total_seconds() / 86400.0

    # ── 6. Decimal → float ─────────────────────────────────────────────────
    if isinstance(value, Decimal):
        value = float(value)

    # ── 7. Strip timezone from datetime (must come after date→datetime) ────
    value = normalize_datetime(value)

    # ── 8. Recognised Excel-compatible primitives pass through unchanged ────
    if value is None or isinstance(
        value,
        (str, int, float, bool, dt.datetime),
    ):
        return value

    # ── 9. Unrecognised → string ───────────────────────────────────────────
    return str(value)
