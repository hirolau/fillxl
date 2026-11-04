"""Number formats come from the file, not from a built-in table.

Excel lets a workbook redefine a built-in format id: a ``<numFmt numFmtId="14">``
entry replaces what the application would otherwise show.  Reading the built-in
table first hides that override.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.styles import StylesPart
from tests.ooxml._build import table_template
from tests.ooxml._patch import patch_part

if TYPE_CHECKING:
    from pathlib import Path

STYLES_PART = "xl/styles.xml"


def test_a_custom_format_that_reuses_a_builtin_id_wins(tmp_path: Path) -> None:
    """Built-in id 14 is normally ``mm-dd-yy``; the file says otherwise."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    patch_part(template, STYLES_PART, _override_builtin_format)

    styles = StylesPart(OoxmlPackage(template))

    assert styles.number_format("0") == "dd.mm.yyyy"


def _override_builtin_format(styles_xml: str) -> str:
    """Declare a custom format under id 14 and give the first xf that id."""
    declared = re.sub(
        r"<numFmts\b[^>]*/>",
        '<numFmts count="1"><numFmt numFmtId="14" formatCode="dd.mm.yyyy"/></numFmts>',
        styles_xml,
        count=1,
    )
    return re.sub(
        r'(<cellXfs\b[^>]*>\s*<xf\b[^>]*?numFmtId=")0(")',
        r"\g<1>14\g<2>",
        declared,
        count=1,
    )
