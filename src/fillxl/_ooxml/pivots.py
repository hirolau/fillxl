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


def pivot_cache_parts(package: OoxmlPackage) -> list[str]:
    """Every pivot cache definition part of the package."""
    return [part for part in package.part_names if PIVOT_CACHE.match(part)]


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
