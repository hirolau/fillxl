"""Detect template features that the openpyxl engine did not write back."""

from __future__ import annotations

import re
import warnings
import zipfile
from collections import Counter
from typing import TYPE_CHECKING, Literal

if TYPE_CHECKING:
    from pathlib import Path

OnLossy = Literal["ignore", "warn", "raise"]


class LossyTemplateError(RuntimeError):
    """The output lost template features and ``on_lossy="raise"`` was set."""


class LossyTemplateWarning(UserWarning):
    """The openpyxl engine dropped template features."""


_PART_FEATURES: tuple[tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"^xl/charts/chartEx\d+\.xml$"), "modern charts (chartEx)"),
    (re.compile(r"^xl/slicers/[^/]+\.xml$"), "slicers"),
    (re.compile(r"^xl/slicerCaches/[^/]+\.xml$"), "slicer caches"),
    (re.compile(r"^xl/timelines/[^/]+\.xml$"), "timelines"),
    (re.compile(r"^xl/timelineCaches/[^/]+\.xml$"), "timeline caches"),
    (re.compile(r"^xl/threadedComments/[^/]+\.xml$"), "threaded comments"),
    (re.compile(r"^xl/persons/[^/]+\.xml$"), "threaded-comment authors"),
    (re.compile(r"^xl/ctrlProps/[^/]+\.xml$"), "form controls"),
    (re.compile(r"^xl/activeX/[^/]+$"), "ActiveX controls"),
    (re.compile(r"^xl/model/[^/]+$"), "data model (Power Pivot)"),
    (re.compile(r"^customXml/item\d+\.xml$"), "custom XML parts (e.g. Power Query)"),
    (re.compile(r"^xl/connections\.xml$"), "data connections"),
    (re.compile(r"^xl/queryTables/[^/]+\.xml$"), "query tables"),
    (re.compile(r"^xl/richData/[^/]+$"), "rich data (pictures in cells, data types)"),
    (re.compile(r"^xl/externalLinks/externalLink\d+\.xml$"), "external links"),
    (re.compile(r"^xl/webextensions/[^/]+$"), "Office add-ins"),
    (re.compile(r"^xl/vbaProject\.bin$"), "VBA project"),
)
_DRAWING_PART = re.compile(r"^xl/drawings/drawing\d+\.xml$")
_WORKSHEET_PART = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
_SHAPE_TAG = re.compile(rb"<(?:\w+:)?(sp|cxnSp|grpSp)[\s>/]")
_EXT_URI = re.compile(rb"<(?:\w+:)?ext\b[^>]*?\buri=\"([^\"]+)\"")
_SHAPE_LABELS = {
    b"sp": "shapes and text boxes",
    b"cxnSp": "connector shapes",
    b"grpSp": "grouped shapes",
}
# Verify against [MS-XLSX] before adding more; unknown URIs are reported raw.
_EXTENSION_LABELS = {
    "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}": "sparklines",
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}": "extended conditional formatting (x14)",
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}": "extended data validation (x14)",
}


def feature_counts(path: Path) -> Counter[str]:
    """Count lossy-sensitive features in an OOXML package (empty if it is not a zip)."""
    counts: Counter[str] = Counter()
    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return counts
    with archive:
        for part in archive.namelist():
            for pattern, label in _PART_FEATURES:
                if pattern.match(part):
                    counts[label] += 1
            if _DRAWING_PART.match(part):
                for match in _SHAPE_TAG.finditer(archive.read(part)):
                    counts[_SHAPE_LABELS[match.group(1)]] += 1
            elif _WORKSHEET_PART.match(part):
                for match in _EXT_URI.finditer(archive.read(part)):
                    uri = match.group(1).decode("ascii", "replace")
                    counts[_EXTENSION_LABELS.get(uri, f"worksheet extension {uri}")] += 1
    return counts


def find_lost_features(template: Path, output: Path) -> list[str]:
    """Features present in *template* that are missing (or fewer) in *output*."""
    lost = feature_counts(template) - feature_counts(output)
    return [f"{label} ({count} lost)" for label, count in sorted(lost.items())]


def check_lossy(template: Path, output: Path, policy: OnLossy, *, stacklevel: int = 2) -> None:
    """Apply the ``on_lossy`` policy to a filled *output*."""
    if policy == "ignore":
        return
    lost = find_lost_features(template, output)
    if not lost:
        return
    msg = (
        "The openpyxl engine did not preserve these template features: "
        + "; ".join(lost)
        + ". Use engine='win32com' to keep them, or set on_lossy='ignore'."
    )
    if policy == "raise":
        raise LossyTemplateError(msg)
    warnings.warn(msg, LossyTemplateWarning, stacklevel=stacklevel)
