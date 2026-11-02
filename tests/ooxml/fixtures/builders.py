"""Workbooks built to break a surgical XML editor.

Each builder returns a valid workbook whose markup a different writer would
produce: another namespace prefix, CDATA, an ``extLst`` whose element names
collide with cells and rows, no cell references, no sheet views, LibreOffice's
namespace noise, or xlsxwriter's table output.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

import openpyxl

from tests.ooxml._build import HEADERS, table_template
from tests.ooxml._patch import part_text, patch_part, put_cell

if TYPE_CHECKING:
    from pathlib import Path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SHEET_PART = "xl/worksheets/sheet1.xml"

_TAG = re.compile(r"<(?P<slash>/?)(?P<name>[A-Za-z_][\w.\-]*)(?=[\s/>])")


def prefixed(path: Path) -> Path:
    """Every element of the sheet uses an ``x:`` prefix."""
    table_template(path, rows=2)

    def transform(xml: str) -> str:
        renamed = _TAG.sub(lambda match: f"<{match.group('slash')}x:{match.group('name')}", xml)
        root = re.search(r"<x:worksheet\b[^>]*>", renamed)
        assert root is not None
        declared = root.group(0).replace(
            "<x:worksheet",
            f'<x:worksheet xmlns:x="{MAIN_NS}"',
            1,
        )
        return f"{renamed[: root.start()]}{declared}{renamed[root.end() :]}"

    patch_part(path, SHEET_PART, transform)
    return path


def cdata(path: Path) -> Path:
    """An out-of-table inline string is stored as CDATA."""
    table_template(path, rows=2)
    patch_part(
        path,
        SHEET_PART,
        lambda xml: put_cell(
            xml,
            '<c r="H1" t="inlineStr"><is><t><![CDATA[a<b]]></t></is></c>',
        ),
    )
    return path


def extlst_lookalikes(path: Path) -> Path:
    """``extLst`` holds elements whose local names are ``c`` and ``row``."""
    table_template(path, rows=2)
    extension = (
        '<extLst><ext uri="{fillxl-test}">'
        '<foo:c xmlns:foo="urn:fillxl:test" r="Z99"/>'
        '<foo:row xmlns:foo="urn:fillxl:test" r="77">'
        '<foo:c xmlns:foo="urn:fillxl:test" r="Z98"/></foo:row>'
        "</ext></extLst>"
    )
    patch_part(
        path,
        SHEET_PART,
        lambda xml: xml.replace("</worksheet>", f"{extension}</worksheet>"),
    )
    return path


def no_cell_refs(path: Path) -> Path:
    """Cells in the body rows carry no ``r`` attribute."""
    table_template(path, rows=2)
    return _without_cell_references(path, rows=(2, 3))


def _without_cell_references(path: Path, *, rows: tuple[int, int]) -> Path:
    def transform(xml: str) -> str:
        for row in range(rows[0], rows[1] + 1):
            start = xml.find(f'<row r="{row}"')
            end = xml.find("</row>", start)
            body = re.sub(r' r="[A-Z]+[0-9]+"', "", xml[start:end])
            xml = f"{xml[:start]}{body}{xml[end:]}"
        return xml

    patch_part(path, SHEET_PART, transform)
    return path


def no_sheet_views(path: Path) -> Path:
    """Neither sheet has a ``<sheetViews>`` block, and one has no row defaults."""
    workbook = openpyxl.load_workbook(table_template(path, rows=2))
    workbook.create_chartsheet("Chart")
    workbook.save(path)
    workbook.close()

    def transform(xml: str) -> str:
        without_views = re.sub(r"<sheetViews>.*?</sheetViews>", "", xml, flags=re.S)
        return re.sub(r"<sheetFormatPr\b[^>]*/>", "", without_views)

    patch_part(path, SHEET_PART, transform)
    chartsheets = [part_name for part_name in _sheet_parts(path) if "/chartsheets/" in part_name]
    assert chartsheets, "the workbook has no chartsheet"
    patch_part(
        path,
        chartsheets[0],
        lambda xml: re.sub(r"<sheetViews>.*?</sheetViews>", "", xml, flags=re.S),
    )
    return path


def libreoffice(path: Path) -> Path:
    """The namespace noise and element order a LibreOffice Calc save produces."""
    table_template(path, rows=2)

    def transform(xml: str) -> str:
        xml = xml.replace(
            "<worksheet ",
            '<worksheet xmlns:loext="http://schemas.libreoffice.org/" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'mc:Ignorable="loext" ',
            1,
        )
        xml = xml.replace(
            "<sheetData",
            '<cols><col min="1" max="5" width="12.5"/></cols><sheetData',
            1,
        )
        return xml.replace(
            "</worksheet>",
            '<extLst><ext uri="{7626C862-2A13-11E5-B345-FEFF819CDC9F}">'
            '<loext:extCalcPr stringRefSyntax="CalcA1"/></ext></extLst></worksheet>',
        )

    patch_part(path, SHEET_PART, transform)
    patch_part(
        path,
        _table_part(path),
        lambda xml: xml.replace("<table ", '<table totalsRowShown="0" ', 1),
    )
    return path


def streaming_writer(path: Path) -> Path:
    """A workbook xlsxwriter wrote with ``add_table``."""
    pytest = __import__("pytest")
    xlsxwriter = pytest.importorskip("xlsxwriter")
    workbook = xlsxwriter.Workbook(str(path))
    worksheet = workbook.add_worksheet("Data")
    for column, name in enumerate(HEADERS):
        worksheet.write(0, column, name)
    for index in range(2):
        worksheet.write(1 + index, 0, "North")
        worksheet.write(1 + index, 1, index + 1)
    worksheet.add_table(
        0,
        0,
        1 + 2,
        len(HEADERS) - 1,
        {
            "name": "Data",
            "columns": [{"header": name} for name in HEADERS],
        },
    )
    workbook.close()
    return path


def _sheet_parts(path: Path) -> list[str]:
    """Every worksheet part name of the package at *path*."""
    import zipfile

    with zipfile.ZipFile(path) as archive:
        return sorted(
            name
            for name in archive.namelist()
            if re.fullmatch(r"xl/(worksheets|chartsheets)/[^/]+\.xml", name)
        )


def _table_part(path: Path) -> str:
    """The table part name of the package at *path*."""
    import zipfile

    with zipfile.ZipFile(path) as archive:
        names = sorted(
            name for name in archive.namelist() if re.fullmatch(r"xl/tables/[^/]+\.xml", name)
        )
    assert names, "the template has no table"
    return names[0]


def part(path: Path, part_name: str) -> str:
    """The text of *part_name*; a small helper for assertions."""
    return part_text(path, part_name)


BUILDERS = {
    "prefixed": prefixed,
    "cdata": cdata,
    "extlst_lookalikes": extlst_lookalikes,
    "no_cell_refs": no_cell_refs,
    "no_sheet_views": no_sheet_views,
    "libreoffice": libreoffice,
    "streaming_writer": streaming_writer,
}
