"""Parts that are not UTF-8 have to be refused, not corrupted.

Excel always writes UTF-8, but other tools can save a part as UTF-16.  Editing
such a part as text would decode it wrongly and write the result back as
garbage; the honest answer is an error that names the part and points at the
engine that can read it.
"""

from __future__ import annotations

import os
import zipfile
from typing import TYPE_CHECKING

import pytest

from fillxl import fill
from tests.ooxml._build import body_rows, table_template

if TYPE_CHECKING:
    from pathlib import Path

SHEET_PART = "xl/worksheets/sheet1.xml"


def test_a_utf16_part_is_refused_with_a_pointer_to_win32com(tmp_path: Path) -> None:
    """The error names the part and the engine that can edit it."""
    template = table_template(tmp_path / "template.xlsx", rows=1)
    _encode_as_utf16(template, SHEET_PART)

    with pytest.raises(ValueError, match=f"{SHEET_PART} is not UTF-8 encoded") as caught:
        fill(
            template,
            tmp_path / "out.xlsx",
            engine="fillxl",
            tables={"Data": body_rows(2)},
        )

    assert "engine='win32com'" in str(caught.value)


def _encode_as_utf16(path: Path, part: str) -> None:
    """Rewrite *part* of the package at *path* as UTF-16 with a byte-order mark."""
    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in infos}
    parts[part] = parts[part].decode("utf-8").encode("utf-16")
    temporary = path.with_name(f"{path.name}.utf16")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(info, parts[info.filename])
    os.replace(temporary, path)
