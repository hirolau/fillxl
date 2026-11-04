"""Sheet visibility rules shared by both backends.

Hiding, deleting and activating a sheet are three operations with one rule
book.  The *planners* here decide what should happen — which sheets go, which
one ends up active, which name matched nothing — and each backend only applies
the result, so the engines cannot disagree about the rules.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from attrs import define

from fillxl._warnings import warn_user

if TYPE_CHECKING:
    from collections.abc import Callable, Collection, Sequence


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


def resolve_sheet_name(name: str, titles: Sequence[str]) -> str | None:
    """The sheet that *name* refers to, matching exactly first and then ignoring case.

    Excel resolves sheet names without regard to case: ``Sheets("rawdata")``
    finds ``RawData``.  Every engine has to do the same, or a fill that asks
    to delete an internal sheet leaves it in the distributed report.

    Examples:
        >>> resolve_sheet_name("rawdata", ["RawData", "Report"])
        'RawData'
        >>> resolve_sheet_name("Report", ["_data", "Report"])
        'Report'
        >>> resolve_sheet_name("missing", ["RawData"]) is None
        True
        >>> resolve_sheet_name("report", ["Report", "REPORT"]) is None
        True

    """
    if name in titles:
        return name
    folded = name.casefold()
    matches = [title for title in titles if title.casefold() == folded]
    return matches[0] if len(matches) == 1 else None


def require_a_visible_sheet(
    visible: Sequence[str],
    removed: Collection[str],
    action: str,
    *,
    hint: str | None = None,
) -> None:
    """Raise if removing *removed* would leave no visible sheet.

    Examples:
        >>> require_a_visible_sheet(["Report", "_data"], ["_data"], "Hiding")
        >>> require_a_visible_sheet(["_data"], ["_data"], "Hiding")  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: Hiding ['_data'] would leave the workbook without a visible sheet...
        >>> try:
        ...     require_a_visible_sheet(["_data"], ["_data"], "Hiding", hint="Set X.")
        ... except ValueError as error:
        ...     str(error).endswith("Excel requires at least one. Set X.")
        True

    """
    if any(name not in removed for name in visible):
        return
    msg = (
        f"{action} {sorted(removed)} would leave the workbook without a visible sheet; "
        "Excel requires at least one."
    )
    if hint:
        msg = f"{msg} {hint}"
    raise ValueError(msg)


def first_remaining(visible: Sequence[str], removed: Collection[str]) -> str:
    """The first visible sheet that is not being removed.

    Examples:
        >>> first_remaining(["_data", "Report", "Notes"], ["_data"])
        'Report'

    """
    return next(name for name in visible if name not in removed)


@define(frozen=True, kw_only=True)
class HidePlan:
    """What hiding the sheets with a given prefix does.

    Attributes:
        to_hide: Visible sheets that will be hidden, in tab order.
        activate: The sheet that has to become active afterwards, or ``None``
            when the active sheet stays the active one.

    """

    to_hide: tuple[str, ...]
    activate: str | None


@define(frozen=True, kw_only=True)
class DeletePlan:
    """What deleting a list of sheets does.

    Attributes:
        present: Sheets that will be deleted, in the order they were asked for.
        not_found: Requested names that matched no sheet.
        activate: The sheet that has to become active afterwards, or ``None``
            when there is nothing to delete.

    """

    present: tuple[str, ...]
    not_found: tuple[str, ...]
    activate: str | None


def plan_hide(prefix: str, visible: Sequence[str], active: str | None) -> HidePlan:
    """Plan hiding the visible sheets whose names start with *prefix*.

    Args:
        prefix: The name prefix, e.g. ``"_"``.
        visible: Titles of the visible sheets, in tab order.
        active: Title of the sheet Excel shows, if it knows one.

    Returns:
        The sheets to hide and the sheet to activate afterwards.

    Raises:
        ValueError: If every visible sheet would be hidden.

    Examples:
        >>> plan_hide("_", ["_data", "Report"], "Report")
        HidePlan(to_hide=('_data',), activate=None)
        >>> plan_hide("_", ["_data", "Report"], "_data")
        HidePlan(to_hide=('_data',), activate='Report')
        >>> plan_hide("_", ["Report"], "Report")
        HidePlan(to_hide=(), activate=None)
        >>> plan_hide("_", ["_data"], "_data")  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: Hiding ['_data'] would leave the workbook without a visible sheet...

    """
    to_hide = tuple(name for name in visible if name.startswith(prefix))
    if not to_hide:
        return HidePlan(to_hide=(), activate=None)
    require_a_visible_sheet(
        visible,
        to_hide,
        "Hiding",
        hint="Set hide_underscore_sheets=False to keep it visible.",
    )
    kept = active is not None and active not in to_hide
    return HidePlan(
        to_hide=to_hide,
        activate=None if kept else first_remaining(visible, to_hide),
    )


def plan_delete(
    requested: Sequence[str],
    titles: Sequence[str],
    visible: Sequence[str],
    active: str | None,
) -> DeletePlan:
    """Plan deleting the sheets named in *requested*.

    Args:
        requested: Sheet names as the caller wrote them.
        titles: Titles of every sheet, in tab order.
        visible: Titles of the visible sheets, in tab order.
        active: Title of the sheet Excel shows, if it knows one.

    Returns:
        The sheets to delete, the names that matched nothing, and the sheet to
        activate afterwards.

    Raises:
        ValueError: If deleting would leave no visible sheet.

    Examples:
        >>> plan_delete(["Report"], ["Report", "Notes"], ["Report", "Notes"], "Report")
        DeletePlan(present=('Report',), not_found=(), activate='Notes')
        >>> plan_delete(["report", "Missing"], ["Report", "Notes"], ["Report", "Notes"], "Notes")
        DeletePlan(present=('Report',), not_found=('Missing',), activate='Notes')
        >>> plan_delete(["Missing"], ["Report"], ["Report"], "Report")
        DeletePlan(present=(), not_found=('Missing',), activate=None)
        >>> plan_delete(["Report"], ["Report"], ["Report"], "Report")  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: Deleting ['Report'] would leave the workbook without a visible sheet...

    """
    present: list[str] = []
    not_found: list[str] = []
    for name in requested:
        resolved = resolve_sheet_name(name, titles)
        if resolved is None:
            not_found.append(name)
        elif resolved not in present:
            present.append(resolved)
    if not present:
        return DeletePlan(present=(), not_found=tuple(not_found), activate=None)
    require_a_visible_sheet(visible, present, "Deleting")
    kept = active is not None and active not in present
    return DeletePlan(
        present=tuple(present),
        not_found=tuple(not_found),
        activate=active if kept else first_remaining(visible, present),
    )


def plan_activate(
    name: str,
    titles: Sequence[str],
    is_visible: Callable[[str], bool],
) -> str | None:
    """Plan making *name* the active sheet.

    Args:
        name: Sheet name as the caller wrote it.
        titles: Titles of every sheet, in tab order.
        is_visible: Whether a sheet is visible, by its own title.

    Returns:
        The sheet's own spelling, or ``None`` when no sheet matched.

    Raises:
        ValueError: If the sheet is hidden.

    Examples:
        >>> plan_activate("report", ["Report", "Notes"], lambda title: True)
        'Report'
        >>> plan_activate("missing", ["Report"], lambda title: True) is None
        True
        >>> plan_activate("Notes", ["Report", "Notes"], lambda title: title != "Notes")
        ... # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: Sheet 'Notes' is hidden and cannot be the active sheet.

    """
    resolved = resolve_sheet_name(name, titles)
    if resolved is None:
        return None
    if not is_visible(resolved):
        msg = f"Sheet '{resolved}' is hidden and cannot be the active sheet."
        raise ValueError(msg)
    return resolved
