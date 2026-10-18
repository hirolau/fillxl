"""Chart data range update logic."""

from __future__ import annotations

import logging
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path, PurePosixPath
from typing import TYPE_CHECKING, Any, NamedTuple

if TYPE_CHECKING:
    from collections.abc import Generator

    from openpyxl.workbook.workbook import Workbook

from openpyxl.worksheet.cell_range import CellRange

logger = logging.getLogger(__name__)

# XML namespaces used in the OOXML package relationship files.
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class TableChangeSpecs(NamedTuple):
    """Specification for a table range change.

    All row and column numbers are absolute worksheet coordinates and
    ``header_row`` is the row that holds the column headers, so the first
    data row is always ``header_row + 1``.  ``old_last_row`` and
    ``new_last_row`` are the last rows of the table's *data body*, which is
    what chart data ranges are compared against.  A totals row, if the table
    has one, sits below ``new_last_row`` and is not included.  Empty data
    still leaves one blank body row (Excel requires it), so a table with no
    data rows has ``new_last_row == header_row + 1``.

    Attributes:
        worksheet_title: Title of the worksheet containing the table.
        header_row: Row number of the table header.
        first_col: 1-based index of the table's first column.
        last_col: 1-based index of the table's last column.
        old_last_row: Last data row of the table before the fill.
        new_last_row: Last data row of the table after the fill.

    Examples:
        >>> spec = TableChangeSpecs("Sheet1", 12, 2, 13, 13, 20)
        >>> spec.first_data_row
        13
        >>> spec.old_ref
        'B12:M13'
        >>> spec.new_ref
        'B12:M20'

    """

    worksheet_title: str
    header_row: int
    first_col: int
    last_col: int
    old_last_row: int
    new_last_row: int

    @property
    def first_data_row(self) -> int:
        """Row number of the first data row (the row below the header)."""
        return self.header_row + 1

    @property
    def old_ref(self) -> str:
        """The table's reference before the fill (e.g. ``B12:M13``)."""
        return _format_ref(self.first_col, self.header_row, self.last_col, self.old_last_row)

    @property
    def new_ref(self) -> str:
        """The table's reference after the fill (e.g. ``B12:M20``)."""
        return _format_ref(self.first_col, self.header_row, self.last_col, self.new_last_row)


def _format_ref(first_col: int, first_row: int, last_col: int, last_row: int) -> str:
    """Format a 1-based rectangle as an A1-style reference."""
    from openpyxl.utils import get_column_letter

    return f"{get_column_letter(first_col)}{first_row}:{get_column_letter(last_col)}{last_row}"


def _unquote_sheet_name(sheet_name: str) -> str:
    """Strip the quoting Excel adds around sheet names.

    Examples:
        >>> _unquote_sheet_name("'My Sheet'")
        'My Sheet'
        >>> _unquote_sheet_name("Data")
        'Data'

    """
    if len(sheet_name) >= 2 and sheet_name.startswith("'") and sheet_name.endswith("'"):
        return sheet_name[1:-1].replace("''", "'")
    return sheet_name


def all_chart_series(workbook: Workbook) -> Generator[Any, None, None]:
    """Yield all chart series from all worksheets in the workbook.

    Args:
        workbook: openpyxl Workbook instance

    Yields:
        Chart series objects

    """
    for sheet in workbook.sheetnames:
        for chart in workbook[sheet]._charts:
            yield from chart.ser


def all_chart_series_references(workbook: Workbook) -> Generator[Any, None, None]:
    """Yield all chart series data references (strRef/numRef) from all series.

    Args:
        workbook: openpyxl Workbook instance

    Yields:
        Reference objects (strRef or numRef) that contain chart data ranges

    """
    labels = ("cat", "val", "xVal", "yVal")
    ref_types = ("strRef", "numRef")

    for series in all_chart_series(workbook=workbook):
        for label in labels:
            if not hasattr(series, label):
                continue
            chart_label = getattr(series, label)
            for ref_type in ref_types:
                if not hasattr(chart_label, ref_type):
                    continue
                reference = getattr(chart_label, ref_type)
                if reference is None:
                    continue
                yield reference


def update_chart_data_ranges(spec: TableChangeSpecs, workbook: Workbook) -> None:
    """Update chart data ranges after a table has been resized.

    When a table's row count changes, charts that reference that table need their
    data ranges updated to match the new table dimensions.

    Only references that point at the table's *data body* — that is, ranges
    starting one row below the header, ending at the table's last data row,
    and within the table's columns — are updated.

    Args:
        spec: TableChangeSpecs describing the old and new table dimensions
        workbook: openpyxl Workbook instance

    """
    for reference in all_chart_series_references(workbook):
        try:
            ref = reference.f
            if not ref:
                continue

            new_ref = _update_series_reference(ref, spec)
            if new_ref is None or new_ref == ref:
                continue

            reference.f = new_ref

            # Update the cached data point count to match the new range.
            # openpyxl does not update these automatically, and leaving stale
            # ptCount values can cause Excel to drop axis labels (especially
            # for Area charts).
            new_pt_count = spec.new_last_row - spec.first_data_row + 1
            if hasattr(reference, "strCache") and reference.strCache is not None:
                reference.strCache.ptCount = new_pt_count
            if hasattr(reference, "numCache") and reference.numCache is not None:
                reference.numCache.ptCount = new_pt_count

            logger.debug("Updated chart reference: %s -> %s", ref, new_ref)

        except Exception:
            logger.exception(
                "Unable to update chart series reference",
                extra={
                    "ref": getattr(reference, "f", None),
                    "worksheet": spec.worksheet_title,
                },
            )


def _update_series_reference(ref: str, spec: TableChangeSpecs) -> str | None:
    """Return *ref* with its end row updated, or ``None`` if it does not match.

    Args:
        ref: A chart series reference, e.g. ``'My Sheet'!$B$13:$B$13``.
        spec: The table change to apply.

    Returns:
        The updated reference, or ``None`` when the reference does not point
        at the table's data body.

    Examples:
        >>> spec = TableChangeSpecs("Data", 12, 2, 13, 13, 20)
        >>> _update_series_reference("Data!$B$13:$B$13", spec)
        'Data!$B$13:$B$20'
        >>> _update_series_reference("Data2!$B$13:$B$13", spec) is None
        True
        >>> _update_series_reference("Data!$B$12:$B$13", spec) is None
        True
        >>> _update_series_reference("Data!$B$13:$B$40", spec) is None
        True

    """
    sheet_name, separator, cell_ref = ref.rpartition("!")
    if not separator or _unquote_sheet_name(sheet_name) != spec.worksheet_title:
        return None

    # A reference may cover several areas: "$B$13:$B$13 $D$13:$D$13".
    areas = cell_ref.split()
    updated = False
    for index, area in enumerate(areas):
        cell_range = CellRange(area)
        if cell_range.min_row != spec.first_data_row or cell_range.max_row != spec.old_last_row:
            continue
        if cell_range.min_col < spec.first_col or cell_range.max_col > spec.last_col:
            continue
        # Replace only the trailing row number, preserving "$" prefixes.
        areas[index] = re.sub(r"\d+$", str(spec.new_last_row), area)
        updated = True

    if not updated:
        return None
    return f"{sheet_name}!{' '.join(areas)}"


def detect_unsupported_chart_sheets(path: Path) -> list[str]:
    """Return sheet names that contain chartEx (extension) charts.

    Excel stores "modern" chart types — waterfall, treemap, sunburst,
    histogram, boxwhisker, funnel, etc. — as *chartEx* parts
    (``xl/charts/chartEx*.xml``) using an extension namespace that
    openpyxl cannot read or write.  When the workbook is saved through
    openpyxl, these charts are **silently dropped**.

    This function scans the raw ``.xlsx`` zip *before* openpyxl loads it
    and traces the relationship chain (chartEx → drawing → worksheet)
    to identify which sheets are affected.

    Args:
        path: Path to the ``.xlsx`` file.

    Returns:
        A list of worksheet names that contain unsupported chartEx charts.
        Empty if the workbook has none.

    """
    try:
        with zipfile.ZipFile(path) as z:
            # 1. Collect chartEx basenames (e.g. "chartEx1.xml")
            chartex_basenames: set[str] = set()
            for name in z.namelist():
                if "/charts/chartEx" in name and name.endswith(".xml"):
                    chartex_basenames.add(PurePosixPath(name).name)

            if not chartex_basenames:
                return []

            # 2. Find drawings that reference any chartEx file
            affected_drawings: set[str] = set()
            for name in z.namelist():
                if "/drawings/_rels/" not in name or not name.endswith(".rels"):
                    continue
                root = ET.fromstring(z.read(name))
                for rel in root:
                    rel_type = rel.get("Type", "")
                    target_basename = PurePosixPath(rel.get("Target", "")).name
                    if "chartEx" in rel_type or target_basename in chartex_basenames:
                        # rels file: xl/drawings/_rels/drawingN.xml.rels
                        # → drawing file: drawingN.xml
                        drawing_file = name.split("_rels/")[1].removesuffix(".rels")
                        affected_drawings.add(drawing_file)

            if not affected_drawings:
                return []

            # 3. Find worksheets that reference those drawings
            affected_sheet_files: set[str] = set()
            for name in z.namelist():
                if "/worksheets/_rels/" not in name or not name.endswith(".rels"):
                    continue
                root = ET.fromstring(z.read(name))
                for rel in root:
                    if "drawing" not in rel.get("Type", ""):
                        continue
                    target_basename = PurePosixPath(rel.get("Target", "")).name
                    if target_basename in affected_drawings:
                        sheet_file = name.split("_rels/")[1].removesuffix(".rels")
                        affected_sheet_files.add(sheet_file)

            if not affected_sheet_files:
                return []

            # 4. Map worksheet files → sheet names via workbook.xml
            wb_rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rid_to_sheet_file: dict[str, str] = {}
            for rel in wb_rels_root:
                if "worksheet" in rel.get("Type", ""):
                    rid_to_sheet_file[rel.get("Id", "")] = PurePosixPath(rel.get("Target", "")).name

            wb_root = ET.fromstring(z.read("xl/workbook.xml"))
            result: list[str] = []
            for sheet_elem in wb_root.iter(f"{{{_MAIN_NS}}}sheet"):
                sheet_name = sheet_elem.get("name", "")
                rid = sheet_elem.get(f"{{{_R_NS}}}id", "")
                if rid_to_sheet_file.get(rid, "") in affected_sheet_files:
                    result.append(sheet_name)

            return result

    except Exception:
        logger.debug("Failed to scan for chartEx charts", exc_info=True)
        return []
