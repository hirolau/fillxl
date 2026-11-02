"""Pivot cache definitions.

A fill cannot refresh a pivot table without Excel, but it can tell Excel to
refresh the cache when the workbook opens — which is what the other engines do
and all that is needed for a resized source range to be picked up.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING

from fillxl._ooxml import xmltext

if TYPE_CHECKING:
    from fillxl._ooxml.package import OoxmlPackage

logger = logging.getLogger(__name__)

PIVOT_CACHE = re.compile(r"^xl/pivotCache/pivotCacheDefinition\d*\.xml$")
_ROOT = xmltext.start_tag_re("pivotCacheDefinition")
_WORKSHEET_SOURCE = xmltext.element_re("worksheetSource")


def pivot_cache_parts(package: OoxmlPackage) -> list[str]:
    """Every pivot cache definition part of the package."""
    return [part for part in package.part_names if PIVOT_CACHE.match(part)]


def pivot_worksheet_sources(package: OoxmlPackage) -> list[tuple[str, str]]:
    """``(sheet, ref)`` of every pivot cache source that names an A1 range.

    A cache source can also name an Excel table (``<worksheetSource name="Sales"/>``);
    Excel keeps that one in step with the table, so it is not returned.
    """
    sources: list[tuple[str, str]] = []
    for part in pivot_cache_parts(package):
        for match in _WORKSHEET_SOURCE.finditer(package.text(part)):
            element = match.group(0)
            sheet = xmltext.attribute(element, "sheet")
            reference = xmltext.attribute(element, "ref")
            if sheet is None or reference is None:
                continue
            sources.append((xmltext.unescape(sheet), xmltext.unescape(reference)))
    return sources


def set_refresh_on_load(package: OoxmlPackage) -> int:
    """Mark every pivot cache for refresh; returns how many caches changed."""
    changed = 0
    for part in pivot_cache_parts(package):
        text = package.text(part)
        root = _ROOT.search(text)
        if root is None:  # pragma: no cover - the part is named after its root
            continue
        updated = xmltext.set_tag_attribute(root.group(0), "refreshOnLoad", "1")
        if updated == root.group(0):
            continue
        package.set_text(part, text[: root.start()] + updated + text[root.end() :])
        logger.debug("Marked %s for refresh on load", part)
        changed += 1
    return changed
