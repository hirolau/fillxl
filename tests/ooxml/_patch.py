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

from fillxl._ooxml import xmltext

_TAG = r"(?:[A-Za-z_][\w.\-]*:)?"


def part_text(path: Path, part_name: str) -> str:
    """The decoded text of *part_name* in the package at *path*."""
    with zipfile.ZipFile(path) as archive:
        return archive.read(part_name).decode("utf-8")


def _cell_regex(reference: str) -> re.Pattern[str]:
    return re.compile(
        rf"<(?P<prefix>{_TAG})c\b[^>]*\br=\"{reference}\"[^>]*?(?:/>|>.*?</(?P=prefix)c>)",
        re.S,
    )


_FORMULA_ELEMENT = re.compile(
    rf"<(?P<prefix>{_TAG})f\b[^>]*?(?:/>|>.*?</(?P=prefix)f>)",
    re.S,
)


def share_formulas(
    sheet_xml: str,
    first: str,
    last: str,
    *,
    si: str = "0",
    master_text: str | None = None,
) -> str:
    """Make the formulas from *first* to *last* one shared group.

    The master keeps its text and gains ``t="shared" ref="…" si="…"``; every
    other cell of the range gets an empty ``<f t="shared" si="…"/>``, which is
    what Excel writes when a formula is filled down or right.

    Examples:
        >>> xml = '<row r="2"><c r="A2"><f>B2*2</f></c><c r="B2"><f>C2*2</f></c></row>'
        >>> share_formulas(xml, "A2", "B2")
        '<row r="2"><c r="A2"><f t="shared" ref="A2:B2" si="0">B2*2</f></c><c r="B2"><f t="shared" si="0"/></c></row>'

    """
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.cell_range import CellRange

    bounds = CellRange(f"{first}:{last}")
    addresses = [
        f"{get_column_letter(column)}{row}"
        for cells in bounds.rows
        for row, column in cells
    ]
    if master_text is None:
        cell = _cell_regex(addresses[0]).search(sheet_xml)
        if cell is None:
            msg = f"no cell {addresses[0]} in the sheet"
            raise AssertionError(msg)
        formula = _FORMULA_ELEMENT.search(cell.group(0))
        if formula is None:
            msg = f"cell {addresses[0]} has no formula to share"
            raise AssertionError(msg)
        master_text = xmltext.element_text(formula.group(0), "f")
        if master_text is None:
            msg = f"cell {addresses[0]} has an empty formula to share"
            raise AssertionError(msg)
    for index, address in enumerate(addresses):
        element = (
            f'<f t="shared" ref="{bounds.coord}" si="{si}">{master_text}</f>'
            if index == 0
            else f'<f t="shared" si="{si}"/>'
        )
        cell = _cell_regex(address).search(sheet_xml)
        if cell is None:
            msg = f"no cell {address} in the sheet"
            raise AssertionError(msg)
        cell_xml = cell.group(0)
        formula = _FORMULA_ELEMENT.search(cell_xml)
        if formula is None:
            msg = f"cell {address} has no formula to share"
            raise AssertionError(msg)
        updated = f"{cell_xml[: formula.start()]}{element}{cell_xml[formula.end() :]}"
        sheet_xml = f"{sheet_xml[: cell.start()]}{updated}{sheet_xml[cell.end() :]}"
    return sheet_xml


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
    row_pattern = re.compile(
        rf"<(?P<prefix>{_TAG})row\b[^>]*\br=\"(?P<row>\d+)\"[^>]*?"
        rf"(?:/>|>.*?</(?P=prefix)row>)",
        re.S,
    )
    row = next(
        (match for match in row_pattern.finditer(sheet_xml) if int(match.group("row")) == row_number),
        None,
    )
    if row is None:
        return _insert_row(sheet_xml, row_pattern, row_number, element)
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


def _insert_row(
    sheet_xml: str,
    row_pattern: re.Pattern[str],
    row_number: int,
    element: str,
) -> str:
    """*sheet_xml* with a new row *row_number* holding *element*, in row order."""
    new_row = f'<row r="{row_number}">{element}</row>'
    for match in row_pattern.finditer(sheet_xml):
        if int(match.group("row")) > row_number:
            return f"{sheet_xml[: match.start()]}{new_row}{sheet_xml[match.start() :]}"
    closing = re.search(rf"</(?P<prefix>{_TAG})sheetData>", sheet_xml)
    if closing is None:
        msg = "the part has no <sheetData>"
        raise AssertionError(msg)
    return f"{sheet_xml[: closing.start()]}{new_row}{sheet_xml[closing.start() :]}"



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
