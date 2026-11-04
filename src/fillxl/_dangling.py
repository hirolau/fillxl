"""Find workbook content that still points at a sheet that was deleted.

Excel rewrites a formula to ``#REF!`` when the sheet it names disappears.  The
openpyxl engine removes sheets from the in-memory workbook and leaves every
reference to them untouched, so the output would ship formulas, charts, names
and pivots that Excel can only show as ``#REF!``.  A warning is the only honest
option: fillxl cannot know where a reference should point instead.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from fillxl._charts import chart_references
from fillxl._utils import unquote_sheet_name
from fillxl._warnings import DanglingReferenceWarning, warn_user

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence
    from typing import Any

    from openpyxl.cell.cell import Cell
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.workbook.workbook import Workbook

# A sheet token followed by "!", quoted ('Bob''s Data'!) or bare (RawData!).
# The bare part matches whole names only, so a reference to MyData! is not
# mistaken for a reference to Data!.
_SHEET_REFERENCE = re.compile(r"(?:'(?:[^']|'')+'|[\w.\\]+)!")
_STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')
_MAX_REPORTED = 5


def names_deleted_sheet(text: object, deleted: set[str]) -> bool:
    """Whether *text* references one of *deleted*, ignoring string literals.

    *deleted* holds lower-case sheet names, so the lookup stays free of the
    case folding Excel applies to sheet references.  Anything that is not a
    non-empty string (an empty cell, a number, a date) names nothing.

    Examples:
        >>> names_deleted_sheet("=SUM(RawData!B2:B2)", {"rawdata"})
        True
        >>> names_deleted_sheet("=SUM('Raw Data'!B2)", {"raw data"})
        True
        >>> names_deleted_sheet("=SUM(rawdata!B2)", {"rawdata"})
        True
        >>> names_deleted_sheet("=SUM(MyData!B2)", {"data"})
        False
        >>> names_deleted_sheet('="RawData!A1 is gone"', {"rawdata"})
        False
        >>> names_deleted_sheet(None, {"rawdata"})
        False

    """
    if not isinstance(text, str) or not text:
        return False
    # A sheet name inside a string literal is data, not a reference.
    formula = _STRING_LITERAL.sub('""', text)
    return any(
        unquote_sheet_name(match.group(0)[:-1]).lower() in deleted
        for match in _SHEET_REFERENCE.finditer(formula)
    )


def _formula_cells(sheet: Any) -> Iterator[Cell]:  # noqa: ANN401
    """The cells of *sheet* that hold a formula (a chart sheet has none)."""
    # ``_cells`` is openpyxl's cell store; iter_rows() would visit every empty
    # cell of a filled table as well.
    return (cell for cell in getattr(sheet, "_cells", {}).values() if cell.data_type == "f")


def _defined_names(workbook: Workbook) -> Iterable[tuple[str, DefinedName]]:
    """Workbook-global and sheet-scoped defined names."""
    yield from workbook.defined_names.items()
    for worksheet in workbook.worksheets:
        yield from worksheet.defined_names.items()


def _chart_mentions(chart: Any, deleted: set[str]) -> bool:  # noqa: ANN401
    """Whether any series of *chart* reads from a deleted sheet."""
    return any(names_deleted_sheet(reference.f, deleted) for reference in chart_references(chart))


def _pivot_mentions(pivot: Any, deleted: set[str]) -> bool:  # noqa: ANN401
    """Whether a pivot table's cache is built from a deleted sheet."""
    source = getattr(getattr(pivot.cache, "cacheSource", None), "worksheetSource", None)
    sheet = getattr(source, "sheet", None)
    return isinstance(sheet, str) and sheet.lower() in deleted


def find_dangling_references(workbook: Workbook, deleted: Sequence[str]) -> list[str]:
    """Locations in *workbook* that still reference one of the *deleted* sheets.

    Args:
        workbook: The workbook after the sheets were removed.
        deleted: Names of the sheets that were deleted.

    Returns:
        Human-readable locations such as ``"Report!A1"`` or
        ``"a chart on 'Report'"``, in sheet order.

    """
    names = {name.lower() for name in deleted}
    if not names:
        return []
    found: list[str] = []
    for title in workbook.sheetnames:
        sheet = workbook[title]
        found.extend(
            f"{title}!{cell.coordinate}"
            for cell in _formula_cells(sheet)
            if names_deleted_sheet(cell.value, names)
        )
        charts = getattr(sheet, "_charts", [])
        if any(_chart_mentions(chart, names) for chart in charts):
            found.append(f"a chart on {title!r}")
        pivots = getattr(sheet, "_pivots", [])
        if any(_pivot_mentions(pivot, names) for pivot in pivots):
            found.append(f"a pivot table on {title!r}")
    found.extend(
        f"the defined name {name!r}"
        for name, defined in _defined_names(workbook)
        if names_deleted_sheet(defined.attr_text, names)
    )
    return found


def warn_dangling_references(deleted: Sequence[str], locations: Sequence[str]) -> None:
    """Warn that *locations* still reference the sheets named in *deleted*.

    Examples:
        >>> import warnings
        >>> with warnings.catch_warnings(record=True) as caught:
        ...     warnings.simplefilter("always")
        ...     warn_dangling_references(["RawData"], ["Report!A1"])
        >>> caught[0].category.__name__
        'DanglingReferenceWarning'
        >>> warn_dangling_references(["RawData"], [])

    """
    if not locations:
        return
    sheets = ", ".join(repr(name) for name in deleted)
    noun, verb = ("sheet", "is") if len(deleted) == 1 else ("sheets", "are")
    shown = locations[:_MAX_REPORTED]
    remaining = len(locations) - len(shown)
    listed = ", ".join(shown) + (f" and {remaining} more" if remaining else "")
    warn_user(
        f"Deleted {noun} {sheets}, which {verb} still referenced by: {listed}. "
        "openpyxl does not rewrite references, so Excel shows #REF! there. "
        "Update or remove those formulas, charts, names and pivots, or keep the sheet.",
        DanglingReferenceWarning,
    )
