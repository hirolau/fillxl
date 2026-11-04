"""Error messages describe what a cell holds, not how the file stores it.

A shared string cell stores an index into ``xl/sharedStrings.xml``; reporting
the index tells the user nothing about what is in their way.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from fillxl import fill
from fillxl._ooxml.package import OoxmlPackage
from tests.ooxml._build import body_rows, table_template
from tests.ooxml._patch import patch_part, put_cell

if TYPE_CHECKING:
    from pathlib import Path

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SHARED_STRINGS = "xl/sharedStrings.xml"
RELATIONSHIPS = "xl/_rels/workbook.xml.rels"
SHEET_PART = "xl/worksheets/sheet1.xml"

SHARED_STRINGS_XML = (
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'count="1" uniqueCount="1"><si><t>blocker text</t></si><si><r><t>rich</t></r>'
    "<r><t> text</t></r></si></sst>"
)


def test_the_overflow_error_names_the_blocking_text_not_its_index(tmp_path: Path) -> None:
    """``A3 already contains data ('blocker text')``, not ``('0')``."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    _add_shared_strings(template)
    patch_part(
        template,
        SHEET_PART,
        lambda xml: put_cell(xml, '<c r="A3" t="s"><v>0</v></c>'),
    )

    with pytest.raises(ValueError, match=r"A3 already contains data \('blocker text'\)"):
        fill(
            template,
            tmp_path / "out.xlsx",
            engine="fillxl",
            tables={"Data": body_rows(3)},
        )


def _add_shared_strings(path: Path) -> None:
    """Give *path* a shared-strings part, declared the way Excel declares it."""
    package = OoxmlPackage(path)
    package.set_text(SHARED_STRINGS, SHARED_STRINGS_XML)
    package.add_override(
        SHARED_STRINGS,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
    )
    relationships = package.text(RELATIONSHIPS)
    package.set_text(
        RELATIONSHIPS,
        relationships.replace(
            "</Relationships>",
            f'<Relationship Id="rIdShared" Type="{REL_NS}/sharedStrings" '
            'Target="sharedStrings.xml"/></Relationships>',
        ),
    )
    package.save(path)
