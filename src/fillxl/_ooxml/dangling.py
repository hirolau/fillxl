"""Find workbook content that still points at a sheet that was deleted.

Deleting a sheet leaves every formula, chart, name and pivot source that
referred to it untouched, and Excel can only show ``#REF!`` there.  A warning
is the only honest option: fillxl cannot know where a reference should point
instead.

This is the XML counterpart of :mod:`fillxl._dangling`; the sheet-name matching
and the warning text are shared with the openpyxl engine so the engines report
the same thing.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

from openpyxl.utils import get_column_letter

from fillxl._dangling import names_deleted_sheet, warn_dangling_references
from fillxl._ooxml import xmltext
from fillxl._ooxml.charts import CHART_PART
from fillxl._ooxml.worksheet import SheetPart

if TYPE_CHECKING:
    from collections.abc import Sequence

    from fillxl._ooxml.package import OoxmlPackage
    from fillxl._ooxml.workbook import WorkbookPart

logger = logging.getLogger(__name__)


def find_dangling_references(
    package: OoxmlPackage,
    workbook: WorkbookPart,
    deleted: Sequence[str],
) -> list[str]:
    """Locations in the workbook that still reference one of the *deleted* sheets.

    Args:
        package: The open package, after the sheets were removed.
        workbook: The workbook part, after the sheets were removed.
        deleted: Names of the sheets that were deleted.

    Returns:
        Human-readable locations such as ``"Report!A1"`` or
        ``"a chart on 'Report'"``, in sheet order.

    """
    names = {name.lower() for name in deleted}
    if not names:
        return []
    found: list[str] = []
    chart_parts = _chart_parts(package)
    for entry in workbook.worksheet_parts():
        sheet = SheetPart(package, entry.part)
        found.extend(
            f"{entry.name}!{get_column_letter(column)}{row}"
            for row, column, formula in sheet.formula_cells()
            if names_deleted_sheet(formula, names)
        )
        reachable = package.reachable_from([entry.part]) & chart_parts
        if any(_chart_mentions(package, part, names) for part in sorted(reachable)):
            found.append(f"a chart on {entry.name!r}")
        if any(
            _pivot_mentions(package, part, names)
            for part in package.related_parts(entry.part, "/pivotTable")
        ):
            found.append(f"a pivot table on {entry.name!r}")
    found.extend(
        f"the defined name {entry.name!r}"
        for entry in workbook.defined_names()
        if names_deleted_sheet(entry.text, names)
    )
    return found


def _chart_parts(package: OoxmlPackage) -> set[str]:
    """Every chart part of the package."""
    return {part for part in package.part_names if CHART_PART.match(part)}


def _chart_mentions(package: OoxmlPackage, part: str, deleted: set[str]) -> bool:
    """Whether any ``<c:f>`` formula of the chart *part* reads a deleted sheet."""
    text = package.text(part)
    for match in xmltext.element_re("f").finditer(text):
        formula = xmltext.element_text(match.group(0), "f")
        if formula is not None and names_deleted_sheet("=" + formula, deleted):
            return True
    return False


def _pivot_mentions(package: OoxmlPackage, part: str, deleted: set[str]) -> bool:
    """Whether the pivot table *part* is built from a deleted sheet."""
    for cache in package.related_parts(part, "/pivotCacheDefinition"):
        text = package.text(cache)
        for match in xmltext.element_re("worksheetSource").finditer(text):
            sheet = xmltext.attribute(match.group(0), "sheet")
            if sheet is not None and sheet.lower() in deleted:
                return True
    return False


def warn_about_deleted_sheets(
    package: OoxmlPackage,
    workbook: WorkbookPart,
    deleted: Sequence[str],
) -> None:
    """Warn about everything that still refers to the sheets in *deleted*."""
    warn_dangling_references(deleted, find_dangling_references(package, workbook, deleted))
