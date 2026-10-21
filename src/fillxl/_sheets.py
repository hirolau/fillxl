"""Sheet visibility rules shared by both backends."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Collection, Sequence


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
