"""Chart data range update logic."""

from __future__ import annotations

import logging
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path, PurePosixPath
from typing import TYPE_CHECKING, Any, NamedTuple

if TYPE_CHECKING:
    from collections.abc import Generator

    from openpyxl.workbook.workbook import Workbook

from openpyxl.utils import column_index_from_string

logger = logging.getLogger(__name__)

# XML namespaces used in the OOXML package relationship files.
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class TableChangeSpecs(NamedTuple):
    """Specification for a table range change.

    Attributes:
        old_ref: Original table reference (e.g., "B12:M13")
        new_ref: New table reference after resizing (e.g., "B12:M20")
        worksheet_title: Title of the worksheet containing the table

    """

    old_ref: str
    new_ref: str
    worksheet_title: str


def all_chart_series(workbook: Workbook) -> Generator[Any, None, None]:
    """Yield all chart series from all worksheets in the workbook.

    Args:
        workbook: openpyxl Workbook instance

    Yields:
        Chart series objects

    """
    for sheet in workbook.sheetnames:
        for chart in workbook[sheet]._charts:  # type: ignore[attr-defined, unused-ignore]
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

    Args:
        spec: TableChangeSpecs describing the old and new table references
        workbook: openpyxl Workbook instance

    """
    for reference in all_chart_series_references(workbook):
        try:
            ref = reference.f
            if not ref:
                continue

            # Check if this reference is from the correct worksheet
            if not ref.replace("'", "").startswith(spec.worksheet_title.replace("'", "")):
                continue

            # Parse the reference (e.g., "Sheet1!$B$13:$B$20")
            ref_start = ref.split("!")[-1].split(":")[0]
            ref_end_row = int(ref.split("$")[-1])
            _, ref_start_column, ref_start_row = ref_start.split("$")
            ref_start_row = int(ref_start_row)

            # Parse the old table reference
            table_start_row = int("".join([x for x in spec.old_ref.split(":")[0] if x.isdigit()]))
            old_ref_end_row = int("".join([x for x in spec.old_ref.split(":")[1] if x.isdigit()]))

            # Check if the reference starts one row after the table header
            # (i.e., it's referencing the data rows, not the header)
            if ref_start_row - table_start_row != 1:
                continue

            # Check if the column is within the table column range
            table_start_col_digit, table_end_col_digit = [
                column_index_from_string("".join(letter for letter in x if not letter.isdigit()))
                for x in spec.old_ref.split(":")
            ]
            ref_start_column_digit = column_index_from_string(ref_start_column)

            if not (table_start_col_digit <= ref_start_column_digit <= table_end_col_digit):
                continue

            # Check if the reference end row matches the old table end row
            if old_ref_end_row != ref_end_row:
                continue

            # Calculate the new end row from the new table reference
            new_end_row = int("".join([x for x in spec.new_ref.split(":")[1] if x.isdigit()]))

            # Update the reference (preserve the trailing "$" before the row number)
            new_data_ref = "$".join(ref.split("$")[:-1]) + "$" + str(new_end_row)
            reference.f = new_data_ref

            # Update the cached data point count to match the new range.
            # openpyxl does not update these automatically, and leaving stale
            # ptCount values can cause Excel to drop axis labels (especially
            # for Area charts).
            new_pt_count = new_end_row - ref_start_row + 1
            if hasattr(reference, "strCache") and reference.strCache is not None:
                reference.strCache.ptCount = new_pt_count
            if hasattr(reference, "numCache") and reference.numCache is not None:
                reference.numCache.ptCount = new_pt_count

            logger.debug("Updated chart reference: %s -> %s", ref, new_data_ref)

        except Exception:
            logger.exception(
                "Unable to update chart series reference",
                extra={
                    "ref": getattr(reference, "f", None),
                    "worksheet": spec.worksheet_title,
                },
            )


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
