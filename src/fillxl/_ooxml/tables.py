"""Excel table parts (``xl/tables/tableN.xml``).

A table part holds the range Excel treats as a table, the autofilter range and
per-column metadata: names, totals-row functions and calculated-column
formulas.  Only ``ref`` and the autofilter range are rewritten when a fill
resizes the table; the column metadata is what makes a resized table still a
table, so it is never touched.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from attrs import define
from openpyxl.utils.escape import unescape
from openpyxl.worksheet.cell_range import CellRange

from fillxl._ooxml import xmltext
from fillxl._plan import TableGeometry

if TYPE_CHECKING:
    from fillxl._ooxml.package import OoxmlPackage
    from fillxl._ooxml.workbook import SheetEntry

TABLE_SUFFIX = "/table"
_COLUMNS = xmltext.element_re("tableColumns")
_COLUMN = xmltext.element_re("tableColumn")
_AUTOFILTER = xmltext.element_re("autoFilter")
_ROOT = xmltext.start_tag_re("table")


@define(frozen=True)
class TableColumnInfo:
    """One ``<tableColumn>`` of a table.

    Attributes:
        name: Column name as Excel shows it (escaped characters decoded).
        calculated_formula: ``<calculatedColumnFormula>`` text without ``=``,
            or ``None`` for a plain data column.
        totals_function: ``totalsRowFunction`` such as ``sum``, or ``None``.
        totals_label: ``totalsRowLabel`` such as ``Total``, or ``None``.

    """

    name: str
    calculated_formula: str | None = None
    totals_function: str | None = None
    totals_label: str | None = None


class TablePart:
    """One table definition, tied to the sheet that shows it."""

    def __init__(self, package: OoxmlPackage, part: str, sheet: SheetEntry) -> None:
        """Parse the table part *part*.

        Args:
            package: The open package.
            part: Part name, e.g. ``xl/tables/table1.xml``.
            sheet: The sheet the table sits on.

        """
        self.package = package
        self.part = part
        self.sheet = sheet

    @property
    def text(self) -> str:
        """The current XML of the table part."""
        return self.package.text(self.part)

    @property
    def display_name(self) -> str:
        """The name Excel uses for the table (``displayName``)."""
        root = _ROOT.search(self.text)
        if root is None:  # pragma: no cover - a table part always has a root
            return ""
        return xmltext.attribute(root.group(0), "displayName") or xmltext.attribute(
            root.group(0),
            "name",
        ) or ""

    @property
    def ref(self) -> str:
        """The range the table covers, e.g. ``B2:D8``."""
        root = _ROOT.search(self.text)
        if root is None:  # pragma: no cover
            return ""
        return xmltext.attribute(root.group(0), "ref") or ""

    def geometry(self) -> TableGeometry:
        """Where the table sits and whether it has a header and a totals row."""
        bounds = CellRange(self.ref)
        root = _ROOT.search(self.text)
        root_tag = root.group(0) if root is not None else ""
        header = xmltext.attribute(root_tag, "headerRowCount")
        totals = xmltext.attribute(root_tag, "totalsRowCount")
        return TableGeometry(
            top_row=bounds.min_row,
            last_row=bounds.max_row,
            first_col=bounds.min_col,
            last_col=bounds.max_col,
            header_rows=0 if header == "0" else 1,
            totals_rows=int(totals) if totals and totals.isdigit() else 0,
        )

    def columns(self) -> list[TableColumnInfo]:
        """The table's columns, in table order."""
        block = _COLUMNS.search(self.text)
        if block is None:
            return []
        columns: list[TableColumnInfo] = []
        for match in _COLUMN.finditer(block.group(0)):
            element = match.group(0)
            columns.append(
                TableColumnInfo(
                    name=unescape(xmltext.attribute(element, "name") or ""),
                    calculated_formula=xmltext.element_text(
                        element,
                        "calculatedColumnFormula",
                    ),
                    totals_function=xmltext.attribute(element, "totalsRowFunction"),
                    totals_label=xmltext.attribute(element, "totalsRowLabel"),
                ),
            )
        return columns

    def set_refs(self, ref: str, autofilter_ref: str | None) -> None:
        """Point the table (and its autofilter) at the new range."""
        text = self.text
        root = _ROOT.search(text)
        if root is None:  # pragma: no cover
            return
        updated_root = xmltext.set_tag_attribute(root.group(0), "ref", ref)
        text = text[: root.start()] + updated_root + text[root.end() :]
        if autofilter_ref is not None:
            match = _AUTOFILTER.search(text)
            if match is not None:
                updated = xmltext.set_tag_attribute(match.group(0), "ref", autofilter_ref)
                text = text[: match.start()] + updated + text[match.end() :]
        if text != self.package.text(self.part):
            self.package.set_text(self.part, text)


def tables_on_sheet(package: OoxmlPackage, sheet: SheetEntry) -> list[TablePart]:
    """Every table defined on *sheet*."""
    return [
        TablePart(package, part, sheet)
        for part in package.related_parts(sheet.part, TABLE_SUFFIX)
    ]


def find_table(package: OoxmlPackage, sheets: list[SheetEntry], name: str) -> TablePart:
    """The table called *name*, or :class:`KeyError`.

    Table names are matched exactly, as the other engines do: Excel forbids two
    tables that differ only by case, so an inexact match can only be a caller
    mistake.
    """
    for sheet in sheets:
        for table in tables_on_sheet(package, sheet):
            if table.display_name == name:
                return table
    msg = f"Table '{name}' not found in workbook"
    raise KeyError(msg)


def table_part_names(package: OoxmlPackage) -> list[str]:
    """Every table part of the package, in part order."""
    pattern = re.compile(r"^xl/tables/[^/]+\.xml$")
    return [part for part in package.part_names if pattern.match(part)]
