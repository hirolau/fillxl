"""Value normalisation for Excel compatibility.

Both the ``openpyxl`` and ``win32com`` backends route ALL cell values
through :func:`normalize_value` so that they produce **identical results**
for the same input data.

Normalisation rules, in the order :func:`normalize_value` applies them:
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
    4. ``Decimal`` → ``float`` (Excel has no native ``Decimal`` type).
        ``Decimal`` NaN, sNaN and infinities → ``None``, and a ``Decimal``
        too large for a float becomes an empty cell too, because the IEEE
        check below then sees the ``inf``.
    5. ``float('nan')``, ``float('inf'), float('-inf')`` → ``None``
        (Excel cannot represent IEEE-754 special values).
    6. ``datetime.time`` → ``float`` (fraction of a 24-hour day), which is
        Excel's own time representation, so arithmetic such as ``=A2*24``
        works.  The template's number format controls the display.
    7. ``datetime.timedelta`` → ``float`` (fraction of a 24-hour day).
        Excel has no native ``timedelta`` type.
    8. Timezone-aware ``datetime`` → naive ``datetime`` (tzinfo stripped,
        wall-clock preserved), with a :class:`TimezoneWarning`.
    9. ``datetime`` and ``date`` subclasses (``pendulum``, ``arrow``, …) →
        the exact built-in type, so openpyxl recognises them and the date
        stays a ``date`` rather than becoming a datetime.
   10. ``enum.Enum`` members → their ``value`` (``str(Color.RED)`` is
        ``"Color.RED"``, which is never what the caller meant).
   11. Recognised primitives pass through unchanged.
   12. Unrecognised objects → ``str(obj)`` via ``__str__``.

Values Excel cannot store faithfully (text longer than 32,767 characters,
illegal control characters, lone surrogates, dates before 1900-03-01,
integers with more than 15 significant digits or outside the double range)
are reported by :func:`check_cell_value`.
"""

from __future__ import annotations

import datetime as dt
import enum
import logging
import math
import sys
from contextvars import ContextVar
from decimal import Decimal
from typing import Any, Literal, TypeVar

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from fillxl._warnings import TimezoneWarning, warn_user

T = TypeVar("T")

logger = logging.getLogger(__name__)

_MAX_TEXT_LENGTH = 32_767
_MAX_SIGNIFICANT_DIGITS = 15
_SIGNIFICANT_DIGIT_FLOOR = 10**_MAX_SIGNIFICANT_DIGITS
# Beyond this Excel stores nothing: a number that large is not a double.
_MAX_DOUBLE = 1e308
# pywin32 marshals a Python int as VT_I4 when it fits, and Excel holds
# numbers as doubles anyway, so wider integers are handed over as floats.
_MAX_COM_INT = 2**31
_FIRST_EXCEL_DATE = dt.date(1900, 3, 1)

Severity = Literal["error", "warning"]


class CellValueError(ValueError):
    """A value Excel cannot store, with the cell it was meant for."""

    def __init__(
        self,
        problem: str,
        *,
        table: str | None = None,
        row: int | None = None,
        column: str | None = None,
        name: str | None = None,
    ) -> None:
        self.problem = problem
        self.table = table
        self.row = row
        self.column = column
        self.name = name
        location = describe_cell(table=table, row=row, column=column, name=name)
        super().__init__(f"{location}: {problem}")


def describe_cell(
    *,
    table: str | None = None,
    row: int | None = None,
    column: str | None = None,
    name: str | None = None,
) -> str:
    """Human-readable location of a cell; *row* is the 1-based row in the data.

    Examples:
        >>> describe_cell(table="Sales", row=3, column="Notes")
        "Table 'Sales', data row 3, column 'Notes'"
        >>> describe_cell(name="report_date")
        "Named cell 'report_date'"

    """
    if name is not None:
        return f"Named cell {name!r}"
    return f"Table {table!r}, data row {row}, column {column!r}"


def cell_value_problem(value: Any) -> tuple[Severity, str] | None:  # noqa: ANN401
    """Why Excel cannot store a normalised *value* faithfully, or ``None``.

    Examples:
        >>> cell_value_problem("ok") is None
        True
        >>> cell_value_problem("x" * 32_768)[0]
        'error'
        >>> cell_value_problem("bell\\x07")[0]
        'error'
        >>> cell_value_problem(1234567890123456)[0]
        'warning'
        >>> cell_value_problem(10**15) is None
        True
        >>> cell_value_problem(10**400)[0]
        'error'
        >>> cell_value_problem("\\ud800")[0]
        'error'
        >>> cell_value_problem(dt.datetime(1899, 12, 31))[0]
        'error'

    """
    if isinstance(value, str):
        if len(value) > _MAX_TEXT_LENGTH:
            return "error", (
                f"text is {len(value):,} characters long; an Excel cell holds at most 32,767"
            )
        # isprintable() is a fast C scan; only a non-printable string needs
        # the regex, which tells tabs/newlines (allowed) from illegal chars.
        if not value.isprintable():
            match = ILLEGAL_CHARACTERS_RE.search(value)
            if match:
                return "error", (
                    f"text contains the control character {match.group()!r} at position "
                    f"{match.start()}, which Excel cannot store"
                )
            # A lone surrogate is non-printable but is not a control
            # character, so the regex above misses it.  It cannot be encoded
            # as UTF-8, which is what openpyxl does while saving.
            try:
                value.encode("utf-8")
            except UnicodeEncodeError as exc:
                return "error", (
                    f"text contains the lone surrogate {value[exc.start]!r} at position "
                    f"{exc.start}, which cannot be stored as UTF-8"
                )
        return None
    if isinstance(value, int) and not isinstance(value, bool):
        if abs(value) >= _MAX_DOUBLE:
            return "error", (
                f"integer {value} is too large for an Excel cell; Excel stores numbers as "
                "doubles and cannot hold values at or beyond 1e308"
            )
        if abs(value) >= _SIGNIFICANT_DIGIT_FLOOR and (
            len(str(abs(value)).rstrip("0")) > _MAX_SIGNIFICANT_DIGITS
        ):
            return "warning", (
                f"integer {value} has more than 15 significant digits and Excel will "
                "round it; pass it as text if it is an identifier"
            )
        return None
    if isinstance(value, dt.date):
        day = value.date() if isinstance(value, dt.datetime) else value
        if day < _FIRST_EXCEL_DATE:
            return "error", (
                f"date {value.isoformat()} is before 1900-03-01; Excel's calendar is wrong "
                "before that day (it counts a 29 February 1900), so it cannot be written "
                "reliably"
            )
    return None


def check_cell_value(
    value: Any,  # noqa: ANN401
    *,
    table: str | None = None,
    row: int | None = None,
    column: str | None = None,
    name: str | None = None,
    warned: set[tuple[str | None, str | None]] | None = None,
) -> None:
    """Raise :class:`CellValueError`, or warn once per column, for values Excel cannot store."""
    problem = cell_value_problem(value)
    if problem is None:
        return
    severity, text = problem
    if severity == "error":
        raise CellValueError(text, table=table, row=row, column=column, name=name)
    key = (table or name, column)
    if warned is not None:
        if key in warned:
            return
        warned.add(key)
    warn_user(f"{describe_cell(table=table, row=row, column=column, name=name)}: {text}")


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
        >>> import warnings
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
            warn_user(
                "Timezone-aware datetimes were provided. Excel does not "
                "support timezones — the tzinfo will be stripped.",
                TimezoneWarning,
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
    4. ``Decimal`` → ``float``; NaN, sNaN, infinity and values too large for
       a float → ``None``.
    5. IEEE-754 special floats (NaN, Inf, -Inf) → ``None``.
    6. ``time`` → ``float`` (fraction of a day).
    7. ``timedelta`` → ``float`` (fraction of a 24-hour day).
    8. Timezone-aware ``datetime`` → naive ``datetime`` (with a warning).
    9. ``datetime``/``date`` subclasses → the exact built-in type.
    10. ``enum.Enum`` members → their ``value``.
    11. Recognised primitives pass through unchanged.
    12. Unrecognised objects → ``str(obj)``.

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
        datetime.date(2025, 12, 25)
        >>> round(normalize_value(time(9, 15, 30)) * 24, 6)
        9.258333
        >>> normalize_value(timedelta(hours=6))
        0.25
        >>> from decimal import Decimal
        >>> normalize_value(Decimal('12.34'))
        12.34
        >>> normalize_value(Decimal('NaN')) is None
        True
        >>> normalize_value(Decimal('Infinity')) is None
        True
        >>> normalize_value(Decimal('1e400')) is None
        True
        >>> class Greeting:
        ...     def __str__(self) -> str:
        ...         return "Hello!"
        >>> normalize_value(Greeting())
        'Hello!'
        >>> import enum
        >>> class Color(enum.Enum):
        ...     RED = "red"
        >>> normalize_value(Color.RED)
        'red'

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

    # ── 4. Decimal → float ────────────────────────────────────────────────
    # Before the IEEE check, so that a Decimal that is NaN, infinite or too
    # large for a float is caught instead of reaching the file as nan/inf.
    if isinstance(value, Decimal):
        if not value.is_finite():  # NaN, sNaN, ±Infinity
            return None
        value = float(value)

    # ── 5. IEEE-754 special floats → None ─────────────────────────────────
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None

    # ── 6. time → fraction of a day (Excel's own representation) ──────────
    if isinstance(value, dt.time):
        return (
            value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
        ) / 86_400

    # ── 7. timedelta → float (fraction of a day) ──────────────────────────
    if isinstance(value, dt.timedelta):
        value = value.total_seconds() / 86400.0

    # ── 8. Strip timezone from datetime ───────────────────────────────────
    value = normalize_datetime(value)

    # ── 9. datetime/date subclasses → the exact built-in type ─────────────
    # A ``date`` stays a ``date`` so it can get a date-only format: Excel has
    # no date type, but openpyxl and the COM edge need the distinction to
    # choose between ``yyyy-mm-dd`` and ``yyyy-mm-dd h:mm:ss``.  Subclasses
    # are flattened because openpyxl looks number formats up by exact type.
    if isinstance(value, dt.datetime) and type(value) is not dt.datetime:
        value = dt.datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            value.tzinfo,
            fold=value.fold,
        )
    elif (
        isinstance(value, dt.date)
        and not isinstance(value, dt.datetime)
        and type(value) is not dt.date
    ):
        value = dt.date(value.year, value.month, value.day)

    # ── 10. Enum members → their value ────────────────────────────────────
    if isinstance(value, enum.Enum):
        return normalize_value(value.value)

    # ── 11. Recognised Excel-compatible primitives pass through unchanged ─
    if value is None or isinstance(value, (str, int, float, bool, dt.date)):
        return value

    # ── 12. Unrecognised → string ────────────────────────────────────────
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


def scalar_value_problem(value: Any) -> str | None:  # noqa: ANN401
    """Why *value* is not something one Excel cell can hold, or ``None``.

    Containers, binary data and tabular objects used to reach the file as
    their ``repr``: a DataFrame in a named cell became its text dump and
    ``b"abc"`` became ``"b'abc'"``.  They are refused instead, with the same
    exception for a named cell and for a table cell.

    Examples:
        >>> scalar_value_problem("text") is None
        True
        >>> scalar_value_problem(42) is None
        True
        >>> scalar_value_problem([1, 2])
        'list is not a single cell value'
        >>> scalar_value_problem(b"abc")
        'bytes is not a single cell value; decode it to text first'

    """
    if isinstance(value, (bytes, bytearray, memoryview)):
        return f"{type(value).__name__} is not a single cell value; decode it to text first"
    if isinstance(value, (list, tuple, set, frozenset, dict)):
        return f"{type(value).__name__} is not a single cell value"
    if _is_tabular(value):
        return f"{type(value).__name__} is not a single cell value"
    return None


def check_named_cell_scalar(name: str, value: Any) -> None:  # noqa: ANN401
    """Refuse a named-cell value that is not a single Excel value.

    Raises:
        CellValueError: If *value* is a container, binary data or tabular.

    Examples:
        >>> check_named_cell_scalar("report_date", 42)
        >>> check_named_cell_scalar("report", [1, 2])  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        fillxl._normalize.CellValueError: Named cell 'report': list is not a single cell value; ...
        >>> check_named_cell_scalar("blob", b"abc")  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        fillxl._normalize.CellValueError: Named cell 'blob': bytes is not a single cell value; ...

    """
    problem = scalar_value_problem(value)
    if problem is not None:
        # The tables= hint only fits rows-and-columns data, not binary blobs.
        hint = "" if isinstance(value, (bytes, bytearray, memoryview)) else _TABLES_HINT
        raise CellValueError(f"{problem}{hint}", name=name)


def check_table_cell_scalar(
    value: Any,  # noqa: ANN401
    *,
    table: str,
    row: int,
    column: str,
) -> None:
    """Refuse a table cell value that is not a single Excel value.

    Raises:
        CellValueError: If *value* is a container, binary data or tabular.

    """
    problem = scalar_value_problem(value)
    if problem is not None:
        raise CellValueError(problem, table=table, row=row, column=column)


_TABLES_HINT = "; put rows and columns in tables= instead"

_TABULAR_TYPES: tuple[tuple[str, str], ...] = (
    ("pandas", "DataFrame"),
    ("pandas", "Series"),
    ("polars", "DataFrame"),
    ("polars", "Series"),
    ("polars", "LazyFrame"),
    ("duckdb", "DuckDBPyRelation"),
)


def _is_tabular(value: Any) -> bool:  # noqa: ANN401
    """Whether *value* is a DataFrame, Series or query relation.

    The modules are looked up in :data:`sys.modules`: the object can only
    exist if its library was already imported.

    Examples:
        >>> _is_tabular([[1, 2]])
        False

    """
    if isinstance(value, (str, bytes, bytearray)):
        return False
    for module_name, type_name in _TABULAR_TYPES:
        module = sys.modules.get(module_name)
        candidate = getattr(module, type_name, None) if module is not None else None
        if isinstance(candidate, type) and isinstance(value, candidate):
            return True
    numpy = sys.modules.get("numpy")
    return numpy is not None and isinstance(value, numpy.ndarray) and value.ndim > 0


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
        datetime.date(2025, 1, 2)

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
    return com_value_from_normalized(prepare_cell_value(value), escape_formulas=escape_formulas)


def com_value_from_normalized(value: Any, *, escape_formulas: bool = False) -> Any:  # noqa: ANN401
    """The COM-specific part of :func:`prepare_com_value` for an already normalised value.

    Integers outside the 32-bit range are handed over as floats: pywin32
    marshals a wider Python int as VT_I8 or fails outright, while Excel holds
    every number as a double anyway.

    Examples:
        >>> com_value_from_normalized(dt.date(2025, 6, 15)).isoformat()
        '2025-06-15T00:00:00+00:00'
        >>> com_value_from_normalized(2**31)
        2147483648.0
        >>> com_value_from_normalized(2**31 - 1)
        2147483647

    """
    if isinstance(value, str):
        if escape_formulas or not value.startswith("="):
            return "'" + value
        return value
    if isinstance(value, dt.datetime):
        return value.replace(tzinfo=dt.UTC)
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day, tzinfo=dt.UTC)
    if isinstance(value, int) and not isinstance(value, bool) and abs(value) >= _MAX_COM_INT:
        return float(value)
    return value
