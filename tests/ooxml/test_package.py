"""The package container: writing a workbook back and refusing a broken one.

A surgical text edit that produces malformed XML only shows up when a reader
opens the file, so the writer checks every part it changed before it writes
anything at all.
"""

from __future__ import annotations

import re
import zipfile
from typing import TYPE_CHECKING

import pytest

from fillxl._ooxml.package import OoxmlPackage
from tests.ooxml._build import table_template
from tests.ooxml._patch import part_text

if TYPE_CHECKING:
    from pathlib import Path

WORKBOOK = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"


def test_a_corrupted_part_is_refused_before_the_workbook_is_written(tmp_path: Path) -> None:
    """No output file may exist when the XML is broken: a partial file is worse."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    package = OoxmlPackage(template)
    package.set_text("xl/workbook.xml", "<workbook>")

    with pytest.raises(RuntimeError, match=r"xl/workbook\.xml"):
        package.save(tmp_path / "out.xlsx")

    assert not (tmp_path / "out.xlsx").exists()


def test_removing_a_relationship_from_a_package_that_has_not_read_it_yet_removes_it(
    tmp_path: Path,
) -> None:
    """The relationship reader is a lazy cache, not a list of the parts that exist."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    with zipfile.ZipFile(template) as archive:
        rels_xml = archive.read(WORKBOOK_RELS).decode("utf-8")
    rel_id = re.search(r'Id="([^"]+)"', rels_xml)
    assert rel_id is not None
    removed = rel_id.group(1)

    package = OoxmlPackage(template)
    package.remove_relationship(WORKBOOK, removed)

    assert all(rel.rel_id != removed for rel in package.relationships(WORKBOOK))
    output = tmp_path / "out.xlsx"
    package.save(output)
    assert f'Id="{removed}"' not in part_text(output, WORKBOOK_RELS)
