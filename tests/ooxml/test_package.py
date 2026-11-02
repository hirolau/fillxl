"""The package container: writing a workbook back and refusing a broken one.

A surgical text edit that produces malformed XML only shows up when a reader
opens the file, so the writer checks every part it changed before it writes
anything at all.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from fillxl._ooxml.package import OoxmlPackage
from tests.ooxml._build import table_template

if TYPE_CHECKING:
    from pathlib import Path


def test_a_corrupted_part_is_refused_before_the_workbook_is_written(tmp_path: Path) -> None:
    """No output file may exist when the XML is broken: a partial file is worse."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    package = OoxmlPackage(template)
    package.set_text("xl/workbook.xml", "<workbook>")

    with pytest.raises(RuntimeError, match=r"xl/workbook\.xml"):
        package.save(tmp_path / "out.xlsx")

    assert not (tmp_path / "out.xlsx").exists()
