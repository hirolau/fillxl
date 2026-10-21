"""Number formats for dates, shared so both backends display dates the same way.

openpyxl gives a date or datetime the formats below when the cell has no date
format yet.  The win32com backend applies the same rule explicitly, because
Excel would otherwise pick a locale-dependent format.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Any

from openpyxl.styles.numbers import FORMAT_DATE_DATETIME, FORMAT_DATE_YYYYMMDD2, is_date_format

if TYPE_CHECKING:
    from collections.abc import Sequence

DATE_FORMAT = FORMAT_DATE_YYYYMMDD2
DATETIME_FORMAT = FORMAT_DATE_DATETIME


def temporal_format(value: Any) -> str | None:  # noqa: ANN401
    """The default number format for a normalised *value*, or ``None`` if it is not a date.

    Examples:
        >>> temporal_format(dt.date(2025, 1, 2))
        'yyyy-mm-dd'
        >>> temporal_format(dt.datetime(2025, 1, 2, 9, 30))
        'yyyy-mm-dd h:mm:ss'
        >>> temporal_format(45000.0) is None
        True

    """
    if isinstance(value, dt.datetime):
        return DATETIME_FORMAT
    if isinstance(value, dt.date):
        return DATE_FORMAT
    return None


def needs_temporal_format(existing: object) -> bool:
    """Whether a cell whose number format is *existing* needs a date format.

    ``None`` is what COM returns for a range that mixes formats; it is
    unknown, so it never qualifies.

    Examples:
        >>> needs_temporal_format("General")
        True
        >>> needs_temporal_format("0.00")
        True
        >>> needs_temporal_format("dd.mm.yyyy")
        False
        >>> needs_temporal_format(None)
        False

    """
    return isinstance(existing, str) and not is_date_format(existing)


def temporal_format_runs(values: Sequence[Any]) -> list[tuple[int, int, str]]:
    """Inclusive 0-based runs of rows that need the same date format.

    Blank values neither start nor break a run; other non-dates break it.

    Examples:
        >>> d, t = dt.date(2025, 1, 2), dt.datetime(2025, 1, 2, 9, 30)
        >>> temporal_format_runs([d, None, d, 5, t])
        [(0, 2, 'yyyy-mm-dd'), (4, 4, 'yyyy-mm-dd h:mm:ss')]
        >>> temporal_format_runs([None, "x"])
        []

    """
    runs: list[tuple[int, int, str]] = []
    current: str | None = None
    start = last = 0
    for index, value in enumerate(values):
        if value is None:
            continue
        fmt = temporal_format(value)
        if fmt is not None and fmt == current:
            last = index
            continue
        if current is not None:
            runs.append((start, last, current))
        current, start, last = fmt, index, index
    if current is not None:
        runs.append((start, last, current))
    return runs
