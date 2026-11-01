"""Sheet visibility rules shared by both backends."""

from __future__ import annotations

from typing import TYPE_CHECKING

from fillxl._warnings import warn_user

if TYPE_CHECKING:
    from collections.abc import Collection, Sequence


def warn_sheet_not_found(name: str, consequence: str) -> None:
    """Warn that a configured sheet name matched nothing.

    Both engines emit the same text, because a name that silently does
    nothing is a config bug the caller needs to see.

    Examples:
        >>> import warnings
        >>> with warnings.catch_warnings(record=True) as caught:
        ...     warnings.simplefilter("always")
        ...     warn_sheet_not_found("Data", "nothing was deleted")
        >>> caught[0].category.__name__
        'FillXLWarning'

    """
    warn_user(f"Sheet {name!r} not found; {consequence}")


def require_a_visible_sheet(visible: Sequence[str], removed: Collection[str], action: str) -> None:
    """Raise if removing *removed* would leave no visible sheet.

    Examples:
        >>> require_a_visible_sheet(["Report", "_data"], ["_data"], "Hiding")
        >>> require_a_visible_sheet(["_data"], ["_data"], "Hiding")  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: Hiding ['_data'] would leave the workbook without a visible sheet...

    """
    if any(name not in removed for name in visible):
        return
    msg = (
        f"{action} {sorted(removed)} would leave the workbook without a visible sheet; "
        "Excel requires at least one."
    )
    raise ValueError(msg)


def first_remaining(visible: Sequence[str], removed: Collection[str]) -> str:
    """The first visible sheet that is not being removed.

    Examples:
        >>> first_remaining(["_data", "Report", "Notes"], ["_data"])
        'Report'

    """
    return next(name for name in visible if name not in removed)
