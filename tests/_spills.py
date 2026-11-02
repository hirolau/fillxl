"""Templates that carry Excel's dynamic-array (spilled formula) markup.

Excel marks a spilled formula with ``<f t="array" ref="…">`` on the master
cell, ``cm="1"`` to point at ``xl/metadata.xml``, and cached values on the
children — nothing else.  openpyxl cannot write that markup, so these fixtures
are authored with openpyxl and then patched to the exact shape Excel 16 writes
(verified end to end in ``BACKEND-RESEARCH.md`` §3.5).
"""

from __future__ import annotations

import re
import xml.parsers.expat as expat
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn, TableFormula

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
METADATA_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
)
METADATA_PART = "xl/metadata.xml"

METADATA_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    f'<metadata xmlns="{MAIN_NS}" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" '
    'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
    'clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
    "</metadataTypes>"
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext '
    'uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
    "</futureMetadata>"
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    "</metadata>"
)


def _cell_element(xml: str, reference: str) -> tuple[int, int] | None:
    """The span of the ``<c>`` element for *reference* in *xml*."""
    pattern = re.compile(
        rf"<c\b[^>]*\br=\"{reference}\"[^>]*?(?:/>|>.*?</c>)",
        re.S,
    )
    match = pattern.search(xml)
    return None if match is None else (match.start(), match.end())


def replace_cell(xml: str, reference: str, element: str) -> str:
    """*xml* with the ``<c>`` element of *reference* replaced by *element*."""
    span = _cell_element(xml, reference)
    if span is None:
        msg = f"no cell {reference} in the sheet"
        raise AssertionError(msg)
    return f"{xml[: span[0]]}{element}{xml[span[1] :]}"


def _sheet_part(parts: dict[str, bytes]) -> str:
    """The part name of the first worksheet."""
    names = sorted(name for name in parts if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name))
    return names[0]


def patch_package(path: Path, parts: dict[str, bytes]) -> Path:
    """Rewrite the package at *path* with *parts*."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


def read_package(path: Path) -> dict[str, bytes]:
    """Every part of the package at *path*, by name."""
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def add_metadata_part(parts: dict[str, bytes]) -> None:
    """Declare and add ``xl/metadata.xml``, as Excel does for dynamic arrays."""
    parts[METADATA_PART] = METADATA_XML.encode()
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>",
        (
            f'<Override PartName="/{METADATA_PART}" ContentType="{METADATA_CONTENT_TYPE}"/></Types>'
        ).encode(),
    )
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    assert rels.count("</Relationships>") == 1
    parts["xl/_rels/workbook.xml.rels"] = rels.replace(
        "</Relationships>",
        f'<Relationship Id="rIdSheetMeta" Type="{REL_NS}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>",
    ).encode()


def _build_table_sheet(workbook: openpyxl.Workbook) -> openpyxl.worksheet.worksheet.Worksheet:
    """``Data`` with table ``Tbl`` (B2:D5): header, two body rows, totals row."""
    sheet = workbook.active
    sheet.title = "Data"
    for column, name in enumerate(("Region", "Amount", "Double"), start=2):
        sheet.cell(row=2, column=column, value=name)
    sheet["B3"], sheet["C3"], sheet["D3"] = "North", 10, "=C3*2"
    sheet["B4"], sheet["C4"], sheet["D4"] = "South", 20, "=C4*2"
    sheet["B5"] = "Total"
    sheet["C5"] = "=SUBTOTAL(109,Tbl[Amount])"
    sheet["D5"] = "=SUBTOTAL(109,Tbl[Double])"
    columns = [
        TableColumn(id=1, name="Region", totalsRowLabel="Total"),
        TableColumn(id=2, name="Amount", totalsRowFunction="sum"),
        TableColumn(id=3, name="Double", totalsRowFunction="sum"),
    ]
    columns[2].calculatedColumnFormula = TableFormula(attr_text="[@Amount]*2")
    table = Table(
        displayName="Tbl",
        ref="B2:D5",
        totalsRowCount=1,
        tableColumns=columns,
    )
    sheet.add_table(table)
    workbook.defined_names.add(DefinedName("report_title", attr_text="Data!$B$1"))
    sheet["B1"] = "Report"
    return sheet


def build_spill_template(path: Path) -> Path:
    """A template with a totals-row table and two spilled dynamic arrays.

    ``F8`` spills ``SEQUENCE(5)`` over ``F8:F12``, ``H8`` sums it with the
    ``F8#`` operator, and ``J3`` spills ``SORT(UNIQUE(Tbl[Region]))`` over
    ``J3:J4`` — the same shape the research prototype was verified against.
    """
    workbook = openpyxl.Workbook()
    sheet = _build_table_sheet(workbook)
    for offset, value in enumerate(range(1, 6)):
        sheet.cell(row=8 + offset, column=6, value=value)
    sheet["H8"] = "=SUM(F8)"
    sheet["J3"], sheet["J4"] = "North", "South"
    workbook.save(path)
    workbook.close()

    parts = read_package(path)
    part = _sheet_part(parts)
    xml = parts[part].decode()
    xml = replace_cell(
        xml,
        "F8",
        '<c r="F8" cm="1"><f t="array" ref="F8:F12">_xlfn.SEQUENCE(5)</f><v>1</v></c>',
    )
    xml = replace_cell(
        xml,
        "H8",
        '<c r="H8"><f>SUM(_xlfn.ANCHORARRAY(F8))</f><v>15</v></c>',
    )
    xml = replace_cell(
        xml,
        "J3",
        '<c r="J3" t="str" cm="1"><f t="array" ref="J3:J4">'
        "_xlfn.SORT(_xlfn.UNIQUE(Tbl[Region]))</f><v>North</v></c>",
    )
    parts[part] = xml.encode()
    add_metadata_part(parts)
    return patch_package(path, parts)


def build_spill_in_growth_path(path: Path) -> Path:
    """A template whose spilled array sits where the table has to grow.

    The spill's children are *blank* (``<c r="B7" t="str"><v/></c>``), which is
    what Excel writes when a spilled result is empty: a value-based occupancy
    check sees an empty cell, and only treating the array range as occupied
    stops the fill from writing over part of somebody's formula.
    """
    workbook = openpyxl.Workbook()
    sheet = _build_table_sheet(workbook)
    sheet["B6"] = "north"
    sheet["B7"] = None
    sheet["B8"] = None
    sheet["B9"] = None
    workbook.save(path)
    workbook.close()

    parts = read_package(path)
    part = _sheet_part(parts)
    xml = parts[part].decode()
    xml = replace_cell(
        xml,
        "B6",
        '<c r="B6" cm="1"><f t="array" ref="B6:B9">_xlfn.UNIQUE(Tbl[Region])</f><v>north</v></c>',
    )
    for row in (7, 8, 9):
        span = _cell_element(xml, f"B{row}")
        if span is None:
            continue
        xml = f'{xml[: span[0]]}<c r="B{row}" t="str"><v/></c>{xml[span[1] :]}'
    parts[part] = xml.encode()
    add_metadata_part(parts)
    return patch_package(path, parts)


def assert_well_formed(path: Path) -> None:
    """Every XML part of the package at *path* parses.

    A surgical text edit that appends an attribute to a closing tag produces a
    file no reader can open; openpyxl and Excel both parse, so ``fillxl`` must
    too ? and a silent corruption would only show up when a user opens it.
    """
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            parser = expat.ParserCreate()
            try:
                parser.Parse(archive.read(name), True)
            except expat.ExpatError as exc:
                msg = f"{name} is not well-formed XML: {exc}"
                raise AssertionError(msg) from exc
