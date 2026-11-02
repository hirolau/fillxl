"""Patch one part of a workbook package, keeping every other entry byte for byte.

openpyxl cannot write the shapes that break a surgical editor: shared formulas,
cells without an ``r`` attribute, the 1904 date system, ``extLst`` lookalikes.
The tests therefore start from a workbook openpyxl *can* write and patch the
one part they care about as text.
"""

from __future__ import annotations

import contextlib
import os
import re
import zipfile
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Callable
    from pathlib import Path

_TAG = r"(?:[A-Za-z_][\w.\-]*:)?"


def part_text(path: Path, part_name: str) -> str:
    """The decoded text of *part_name* in the package at *path*."""
    with zipfile.ZipFile(path) as archive:
        return archive.read(part_name).decode("utf-8")


def put_cell(sheet_xml: str, element: str) -> str:
    """*sheet_xml* with the ``<c>`` *element* in place of the cell it replaces.

    The row is found by the reference in *element*; a cell that is not there
    yet is appended to its row, which is how a test adds a spill next to a
    table.

    Examples:
        >>> put_cell('<row r="2"><c r="A2"/></row>', '<c r="A2"><v>1</v></c>')
        '<row r="2"><c r="A2"><v>1</v></c></row>'
        >>> put_cell('<row r="2"><c r="A2"/></row>', '<c r="B2"/>')
        '<row r="2"><c r="A2"/><c r="B2"/></row>'

    """
    reference = re.search(r'\br="([A-Z]+\d+)"', element)
    if reference is None:
        msg = f"the element has no cell reference: {element!r}"
        raise AssertionError(msg)
    address = reference.group(1)
    row_number = int(address[len(address.rstrip("0123456789")) :])
    row = re.compile(
        rf"<(?P<prefix>{_TAG})row\b[^>]*\br=\"{row_number}\"[^>]*?(?:/>|>.*?</(?P=prefix)row>)",
        re.S,
    ).search(sheet_xml)
    if row is None:
        msg = f"no <row> for {address}"
        raise AssertionError(msg)
    element_xml = row.group(0)
    if element_xml.endswith("/>"):
        opening = element_xml[:-2].rstrip()
        return (
            f"{sheet_xml[: row.start()]}{opening}>{element}</{row.group('prefix')}row>"
            f"{sheet_xml[row.end() :]}"
        )
    cell = re.compile(
        rf"<(?P<prefix>{_TAG})c\b[^>]*\br=\"{address}\"[^>]*?(?:/>|>.*?</(?P=prefix)c>)",
        re.S,
    ).search(element_xml)
    if cell is None:
        closing = element_xml.rfind("</")
        updated = f"{element_xml[:closing]}{element}{element_xml[closing:]}"
    else:
        updated = f"{element_xml[: cell.start()]}{element}{element_xml[cell.end() :]}"
    return f"{sheet_xml[: row.start()]}{updated}{sheet_xml[row.end() :]}"



def patch_part(path: Path, part_name: str, transform: Callable[[str], str]) -> None:
    """Rewrite one part of the zip at *path* in place, keeping every other entry byte for byte.

    Args:
        path: The workbook to rewrite.
        part_name: The part to transform, e.g. ``xl/worksheets/sheet1.xml``.
        transform: Receives the part's text and returns its replacement.

    Raises:
        KeyError: If the package has no such part.
        AssertionError: If *transform* returned the text unchanged.  A patch
            that silently does nothing makes a test pass for the wrong reason.

    """
    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in infos}
    if part_name not in parts:
        raise KeyError(part_name)
    original = parts[part_name].decode("utf-8")
    updated = transform(original)
    if updated == original:
        msg = f"transform did not change {part_name}"
        raise AssertionError(msg)
    parts[part_name] = updated.encode("utf-8")

    temporary = path.with_name(f"{path.name}.patching")
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            for info in infos:
                archive.writestr(info, parts[info.filename])
        os.replace(temporary, path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            temporary.unlink()
