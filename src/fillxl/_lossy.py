"""Detect template features that the openpyxl engine did not write back."""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ElementTree
import zipfile
from collections import Counter
from typing import TYPE_CHECKING, Literal

from fillxl._warnings import FillXLWarning, warn_user

if TYPE_CHECKING:
    from collections.abc import Collection
    from pathlib import Path

OnLossy = Literal["ignore", "warn", "raise"]

# The formats whose parts this module can compare part by part.
_XML_SUFFIXES = (".xlsx", ".xlsm")
_VBA_LABEL = "VBA project"


class LossyTemplateError(RuntimeError):
    """The output lost template features and ``on_lossy="raise"`` was set."""


class LossyTemplateWarning(FillXLWarning):
    """An engine dropped template features the fill could not write back."""


_MEDIA_PREFIX = "xl/media/"
_PART_FEATURES: tuple[tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"^xl/charts/chartEx\d+\.xml$"), "modern charts (chartEx)"),
    (re.compile(r"^xl/charts/style\d+\.xml$"), "chart styles"),
    (re.compile(r"^xl/charts/colors\d+\.xml$"), "chart colour schemes"),
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
    (re.compile(r"^xl/metadata\.xml$"), "cell metadata (dynamic arrays, data types)"),
    (re.compile(r"^xl/externalLinks/externalLink\d+\.xml$"), "external links"),
    (re.compile(r"^xl/webextensions/[^/]+$"), "Office add-ins"),
    # openpyxl writes no embeddings at all, so an embedded Word document, PDF
    # or equation disappears without a trace unless it is counted here.
    (re.compile(r"^xl/embeddings/[^/]+$"), "embedded objects (OLE)"),
    (re.compile(r"^xl/vbaProject\.bin$"), "VBA project"),
)
_DRAWING_PART = re.compile(r"^xl/drawings/drawing\d+\.xml$")
_WORKSHEET_PART = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
_SHAPE_TAG = re.compile(rb"<(?:\w+:)?(sp|cxnSp|grpSp|pic)[\s>/]")
_EXT_URI = re.compile(rb"<(?:\w+:)?ext\b[^>]*?\buri=\"([^\"]+)\"")
_SHAPE_LABELS = {
    b"sp": "shapes and text boxes",
    b"cxnSp": "connector shapes",
    b"grpSp": "grouped shapes",
    b"pic": "pictures",
}
# Verify against [MS-XLSX] before adding more; unknown URIs are reported raw.
_EXTENSION_LABELS = {
    "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}": "sparklines",
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}": "extended conditional formatting (x14)",
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}": "extended data validation (x14)",
}
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_OFFICE_DOCUMENT_REL = "/officeDocument"


def _rels_part(part: str) -> str:
    """The ``.rels`` part that lists *part*'s relationships.

    Examples:
        >>> _rels_part("xl/worksheets/sheet1.xml")
        'xl/worksheets/_rels/sheet1.xml.rels'
        >>> _rels_part("workbook.xml")
        '_rels/workbook.xml.rels'

    """
    folder, _, name = part.rpartition("/")
    return posixpath.join(folder, "_rels", f"{name}.rels")


def _read_xml(archive: zipfile.ZipFile, part: str) -> ElementTree.Element | None:
    """*part* parsed as XML, or ``None`` when it is missing or unparsable."""
    try:
        raw = archive.read(part)
    except KeyError:
        return None
    try:
        return ElementTree.fromstring(raw)
    except ElementTree.ParseError:
        return None


def _relationship_map(archive: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Map relationship ids of *part* to the package parts they point at."""
    root = _read_xml(archive, _rels_part(part))
    if root is None:
        return {}
    folder = posixpath.dirname(part)
    targets: dict[str, str] = {}
    for rel in root:
        target = rel.get("Target")
        rel_id = rel.get("Id")
        if target is None or rel_id is None or rel.get("TargetMode") == "External":
            continue
        targets[rel_id] = posixpath.normpath(posixpath.join(folder, target)).lstrip("/")
    return targets


def _workbook_part(archive: zipfile.ZipFile) -> str | None:
    """The package part that holds ``workbook.xml``."""
    root = _read_xml(archive, "_rels/.rels")
    if root is None:
        return None
    for rel in root:
        target = rel.get("Target")
        if target is not None and (rel.get("Type") or "").endswith(_OFFICE_DOCUMENT_REL):
            return posixpath.normpath(target.lstrip("/"))
    return None


def _sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    """Map each sheet name in the package to the part that stores it."""
    workbook_part = _workbook_part(archive)
    if workbook_part is None:
        return {}
    root = _read_xml(archive, workbook_part)
    if root is None:
        return {}
    by_id = _relationship_map(archive, workbook_part)
    sheets: dict[str, str] = {}
    for element in root.iter(f"{{{_MAIN_NS}}}sheet"):
        name = element.get("name")
        part = by_id.get(element.get(f"{{{_REL_NS}}}id", ""))
        if name is not None and part is not None:
            sheets[name] = part
    return sheets


def _part_owners(archive: zipfile.ZipFile) -> dict[str, set[str]]:
    """Map each part reachable from a sheet to the sheets that reach it."""
    owners: dict[str, set[str]] = {}
    for sheet_name, part in _sheet_parts(archive).items():
        pending = [part]
        seen: set[str] = set()
        while pending:
            current = pending.pop()
            if current in seen:
                continue
            seen.add(current)
            owners.setdefault(current, set()).add(sheet_name)
            pending.extend(_relationship_map(archive, current).values())
    return owners


def _parts_owned_only_by(archive: zipfile.ZipFile, sheets: Collection[str]) -> frozenset[str]:
    """Parts that are gone only because *sheets* were deleted.

    A part that a surviving sheet also references (a drawing shared by two
    sheets, say) is kept, so its loss still counts.
    """
    deleted = set(sheets)
    if not deleted:
        return frozenset()
    return frozenset(part for part, owners in _part_owners(archive).items() if owners <= deleted)


def feature_counts(path: Path, *, exclude_sheets: Collection[str] = ()) -> Counter[str]:
    """Count lossy-sensitive features in an OOXML package (empty if it is not a zip).

    Parts that only *exclude_sheets* reference are skipped: deleting a sheet
    removes its pictures and charts on purpose, so their absence from the
    output is not the openpyxl engine dropping anything.

    Args:
        path: The workbook package to inspect.
        exclude_sheets: Names of sheets whose parts are not counted.

    """
    counts: Counter[str] = Counter()
    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return counts
    with archive:
        excluded = _parts_owned_only_by(archive, exclude_sheets)
        for part in archive.namelist():
            if part in excluded:
                continue
            if part.startswith(_MEDIA_PREFIX):
                counts["images (xl/media)"] += 1
                continue
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


def find_lost_features(
    template: Path,
    output: Path,
    *,
    deleted_sheets: Collection[str] = (),
) -> list[str]:
    """Features present in *template* that are missing (or fewer) in *output*.

    Args:
        template: The template package.
        output: The package the openpyxl engine wrote.
        deleted_sheets: Sheets ``delete_sheets`` removed from *output*, whose
            own parts are expected to be gone.

    """
    lost = feature_counts(template, exclude_sheets=deleted_sheets) - feature_counts(output)
    return [f"{label} ({count} lost)" for label, count in sorted(lost.items())]


def check_lossy(
    template: Path,
    output: Path,
    policy: OnLossy,
    *,
    deleted_sheets: Collection[str] = (),
    engine: str = "openpyxl",
) -> None:
    """Apply the ``on_lossy`` policy to a filled *output*."""
    if policy == "ignore":
        return
    lost = find_lost_features(template, output, deleted_sheets=deleted_sheets)
    if not lost:
        return
    features = "; ".join(lost)
    if engine == "fillxl":
        msg = (
            f"The fillxl engine did not preserve these template features: {features}. "
            "This is a fillxl engine bug; please report it with the template. "
            "engine='win32com' preserves them meanwhile, or set on_lossy='ignore'."
        )
    else:
        msg = (
            f"The {engine} engine did not preserve these template features: {features}. "
            "Use engine='win32com' to keep them, or set on_lossy='ignore'."
        )
    _apply_lossy_policy(msg, policy)


def check_conversion_losses(
    template: Path,
    output: Path,
    policy: OnLossy,
    *,
    engine: str = "win32com",
    has_vba: bool | None = None,
) -> None:
    """Report template features that writing *output* in another format dropped.

    Excel writes back everything it read, but the format it wrote to may have
    no room for it: an ``.xlsm`` template saved as ``.xlsx`` loses its VBA
    project.  Nothing warns about that on its own -- Excel's compatibility
    checker is silenced along with the other alerts during a fill.

    Args:
        template: The template the workbook was filled from.
        output: The file the backend wrote.
        policy: The ``on_lossy`` policy.
        engine: Engine name for the message.
        has_vba: Whether the filled workbook has a VBA project, when the
            backend can say so.

    """
    if policy == "ignore" or template.suffix.lower() == output.suffix.lower():
        return
    output_suffix = output.suffix.lower()
    lost: list[str] = []
    if output_suffix in _XML_SUFFIXES:
        # Both sides are XML packages, so their parts can be compared.  Only
        # things the output cannot hold are reported: a part that is gone from
        # the output but was never there either is not a loss.
        lost = find_lost_features(template, output)
    # A binary OOXML output (``.xlsb``) holds the same feature set as the XML
    # formats; comparing parts by name would report every XML part as lost.
    vba = _vba_loss(template, output_suffix, has_vba=has_vba)
    if vba is not None and not any(item.startswith(_VBA_LABEL) for item in lost):
        lost.append(vba)
    if not lost:
        return
    msg = (
        f"The {engine} engine saved a '{template.suffix}' template as "
        f"'{output.suffix}', which cannot hold everything it had: {'; '.join(lost)}. "
        "Save to a format that keeps them, or set on_lossy='ignore'."
    )
    _apply_lossy_policy(msg, policy)


def _vba_loss(template: Path, output_suffix: str, *, has_vba: bool | None) -> str | None:
    """The VBA loss for *output_suffix*, or ``None`` when the format keeps macros.

    Only ``.xlsx`` cannot hold a VBA project: ``.xlsm`` and ``.xlsb`` both can.
    """
    if output_suffix != ".xlsx":
        return None
    count = feature_counts(template).get(_VBA_LABEL, 0)
    if not count and has_vba:
        count = 1
    return f"{_VBA_LABEL} ({count} lost)" if count else None


def _apply_lossy_policy(message: str, policy: OnLossy) -> None:
    """Raise or warn about *message*, depending on *policy*."""
    if policy == "raise":
        raise LossyTemplateError(message)
    warn_user(message, LossyTemplateWarning)
