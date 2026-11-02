"""Chart data references.

A classic chart stores each series source as a formula string
(``SERIES(Data!$C$2,Data!$B$3:$B$4,Data!$C$3:$C$4,1)``) plus a cached copy of
the points.  When a fill resizes a table, the formulas that pointed at the old
data body are rewritten with the tested logic in
:func:`fillxl._charts.update_reference_text`, and the cached point count is
brought in step — a stale count makes Excel drop axis labels.

Modern charts (``chartEx``) are left alone: their sources are not these
``<c:f>`` references, and rewriting what the engine does not model is worse
than leaving it untouched.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING

from fillxl._charts import point_count, point_count_rows, update_reference_text
from fillxl._ooxml import xmltext

if TYPE_CHECKING:
    from fillxl._charts import TableChangeSpecs
    from fillxl._ooxml.package import OoxmlPackage

logger = logging.getLogger(__name__)

CHART_PART = re.compile(r"^xl/charts/chart\d+\.xml$")
_TAG = r"(?:[A-Za-z_][\w.\-]*:)?"
_REFERENCE_BLOCK = re.compile(
    rf"<{_TAG}(?P<kind>strRef|numRef|multiLvlStrRef)\b[^>]*?(?:/>|>.*?</{_TAG}(?P=kind)>)",
    re.S,
)
_DATA_SOURCE = re.compile(
    rf"<{_TAG}(?P<source>cat|val|xVal|yVal|bubbleSize|plus|minus)\b[^>]*>"
    rf".*?</{_TAG}(?P=source)>",
    re.S,
)
_FORMULA = xmltext.element_re("f")
_POINT_COUNT = re.compile(rf"<{_TAG}ptCount\b[^>]*?/>")


def chart_parts(package: OoxmlPackage) -> list[str]:
    """Every classic chart part of the package."""
    return [part for part in package.part_names if CHART_PART.match(part)]


def update_chart_references(package: OoxmlPackage, spec: TableChangeSpecs) -> int:
    """Point every chart series that showed the table body at its new rows.

    Args:
        package: The open package.
        spec: The table's old and new dimensions.

    Returns:
        How many references were rewritten.

    """
    changed = 0
    for part in chart_parts(package):
        text = package.text(part)
        updated, count = _rewrite_part(text, spec)
        if count:
            package.set_text(part, updated)
            logger.debug("Rewrote %d chart reference(s) in %s", count, part)
        changed += count
    return changed


def _rewrite_part(text: str, spec: TableChangeSpecs) -> tuple[str, int]:
    """*text* with the references of *spec* rewritten, and how many changed.

    Only the data-source blocks are rewritten.  A series name lives in
    ``<c:tx>`` and points at a header cell that often sits next to — or, for a
    one-row table, on — the data rows; rewriting it would turn a label into a
    range.
    """
    total = 0

    def replace_block(match: re.Match[str]) -> str:
        nonlocal total
        block = match.group(0)
        formula = _FORMULA.search(block)
        if formula is None:
            return block
        current = xmltext.element_text(formula.group(0), "f")
        if current is None:
            return block
        updated = update_reference_text(current, spec)
        if updated is None:
            return block
        total += 1
        block = block[: formula.start()] + xmltext.replace_element_text(
            formula.group(0),
            "f",
            updated,
        ) + block[formula.end() :]
        multi_level = match.group("kind") == "multiLvlStrRef"
        count = point_count_rows(updated) if multi_level else point_count(updated)
        return _POINT_COUNT.sub(
            lambda m: xmltext.set_tag_attribute(m.group(0), "val", str(count)),
            block,
        )

    def replace_source(match: re.Match[str]) -> str:
        return _REFERENCE_BLOCK.sub(replace_block, match.group(0))

    return _DATA_SOURCE.sub(replace_source, text), total
