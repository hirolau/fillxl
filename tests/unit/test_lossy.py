"""Unit tests for the lossy-feature detector."""

from __future__ import annotations

import warnings
import zipfile
from collections import Counter
from pathlib import Path
from typing import TYPE_CHECKING

import pytest

from fillxl._lossy import LossyTemplateError, check_lossy, feature_counts, find_lost_features

if TYPE_CHECKING:
    from collections.abc import Sequence

_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_PICTURE = (
    b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
    b"<xdr:pic/></xdr:wsDr>"
)


def _package(path: Path, parts: dict[str, bytes]) -> Path:
    """Write a synthetic OOXML package with *parts*."""
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


def _relationships(*entries: tuple[str, str, str]) -> bytes:
    """A ``.rels`` part from ``(id, type-suffix, target)`` tuples."""
    rels = "".join(
        f'<Relationship Id="{rel_id}" Type="{_REL_NS}/{kind}" Target="{target}"/>'
        for rel_id, kind, target in entries
    )
    return f'<Relationships xmlns="{_PACKAGE_NS}">{rels}</Relationships>'.encode()


def _picture_package(
    path: Path,
    sheets: Sequence[str],
    *,
    owners: Sequence[str] = (),
    pictures: int = 1,
) -> Path:
    """Package whose *owners* link to one shared drawing holding *pictures* pictures."""
    parts: dict[str, bytes] = {
        "_rels/.rels": _relationships(("rId1", "officeDocument", "xl/workbook.xml")),
        "xl/workbook.xml": (
            f'<workbook xmlns="{_MAIN_NS}" xmlns:r="{_REL_NS}"><sheets>'
            + "".join(
                f'<sheet name="{name}" sheetId="{index}" r:id="rId{index}"/>'
                for index, name in enumerate(sheets, start=1)
            )
            + "</sheets></workbook>"
        ).encode(),
        "xl/_rels/workbook.xml.rels": _relationships(
            *[
                (f"rId{index}", "worksheet", f"worksheets/sheet{index}.xml")
                for index in range(1, len(sheets) + 1)
            ]
        ),
        **{
            f"xl/worksheets/sheet{index}.xml": f'<worksheet xmlns="{_MAIN_NS}"/>'.encode()
            for index in range(1, len(sheets) + 1)
        },
    }
    for name in owners:
        index = list(sheets).index(name) + 1
        parts[f"xl/worksheets/_rels/sheet{index}.xml.rels"] = _relationships(
            ("rId1", "drawing", "../drawings/drawing1.xml"),
        )
    if owners:
        parts["xl/drawings/drawing1.xml"] = _PICTURE * pictures
        parts["xl/drawings/_rels/drawing1.xml.rels"] = _relationships(
            *[
                (f"rId{index}", "image", f"../media/image{index}.png")
                for index in range(1, pictures + 1)
            ]
        )
        parts.update(
            {f"xl/media/image{index}.png": b"png" for index in range(1, pictures + 1)},
        )
    return _package(path, parts)


def test_identical_packages_lose_nothing(tmp_path: Path) -> None:
    parts = {"xl/workbook.xml": b"<workbook/>", "xl/slicers/slicer1.xml": b"<slicer/>"}
    template = _package(tmp_path / "template.xlsx", parts)
    output = _package(tmp_path / "output.xlsx", parts)

    assert find_lost_features(template, output) == []


def test_a_missing_slicer_part_is_reported(tmp_path: Path) -> None:
    template = _package(
        tmp_path / "template.xlsx",
        {"xl/workbook.xml": b"<workbook/>", "xl/slicers/slicer1.xml": b"<slicer/>"},
    )
    output = _package(tmp_path / "output.xlsx", {"xl/workbook.xml": b"<workbook/>"})

    assert find_lost_features(template, output) == ["slicers (1 lost)"]


def test_fewer_shapes_in_a_drawing_are_reported(tmp_path: Path) -> None:
    drawing = "xl/drawings/drawing1.xml"
    template = _package(
        tmp_path / "template.xlsx",
        {drawing: b"<xdr:wsDr><xdr:sp/><xdr:cxnSp/></xdr:wsDr>"},
    )
    output = _package(
        tmp_path / "output.xlsx",
        {drawing: b"<xdr:wsDr><xdr:sp/></xdr:wsDr>"},
    )

    assert find_lost_features(template, output) == ["connector shapes (1 lost)"]


def test_a_missing_sparkline_extension_is_reported_by_name(tmp_path: Path) -> None:
    sparkline = (
        b'<worksheet><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">'
        b"</ext></extLst></worksheet>"
    )
    template = _package(
        tmp_path / "template.xlsx",
        {"xl/worksheets/sheet1.xml": sparkline},
    )
    output = _package(
        tmp_path / "output.xlsx",
        {"xl/worksheets/sheet1.xml": b"<worksheet/>"},
    )

    assert find_lost_features(template, output) == ["sparklines (1 lost)"]


def test_an_unknown_extension_is_reported_by_uri(tmp_path: Path) -> None:
    template = _package(
        tmp_path / "template.xlsx",
        {"xl/worksheets/sheet1.xml": b'<worksheet><ext uri="{UNKNOWN}"/></worksheet>'},
    )
    output = _package(
        tmp_path / "output.xlsx",
        {"xl/worksheets/sheet1.xml": b"<worksheet/>"},
    )

    assert find_lost_features(template, output) == ["worksheet extension {UNKNOWN} (1 lost)"]


def test_a_non_zip_file_counts_nothing(tmp_path: Path) -> None:
    path = tmp_path / "not_a_zip.xlsx"
    path.write_bytes(b"not a zip")

    assert feature_counts(path) == Counter()


def test_ignore_policy_is_silent(tmp_path: Path) -> None:
    template = _package(tmp_path / "template.xlsx", {"xl/slicers/slicer1.xml": b"<slicer/>"})
    output = _package(tmp_path / "output.xlsx", {"xl/workbook.xml": b"<workbook/>"})

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        check_lossy(template, output, "ignore")

    assert caught == []


def test_warn_policy_warns(tmp_path: Path) -> None:
    template = _package(tmp_path / "template.xlsx", {"xl/slicers/slicer1.xml": b"<slicer/>"})
    output = _package(tmp_path / "output.xlsx", {"xl/workbook.xml": b"<workbook/>"})

    with pytest.warns(UserWarning, match="did not preserve"):
        check_lossy(template, output, "warn")


def test_raise_policy_raises(tmp_path: Path) -> None:
    template = _package(tmp_path / "template.xlsx", {"xl/slicers/slicer1.xml": b"<slicer/>"})
    output = _package(tmp_path / "output.xlsx", {"xl/workbook.xml": b"<workbook/>"})

    with pytest.raises(LossyTemplateError, match="slicers"):
        check_lossy(template, output, "raise")


def test_pictures_are_counted_when_the_drawing_loses_them(tmp_path: Path) -> None:
    """openpyxl drops images when Pillow is missing, so they must be counted."""
    drawing = b"<xdr:wsDr><xdr:pic><xdr:blipFill/></xdr:pic></xdr:wsDr>"
    template = _package(
        tmp_path / "template.xlsx",
        {
            "xl/drawings/drawing1.xml": drawing,
            "xl/media/image1.png": b"png",
        },
    )
    output = _package(tmp_path / "output.xlsx", {"xl/workbook.xml": b"<workbook/>"})

    lost = find_lost_features(template, output)

    assert any("pictures" in item for item in lost)
    assert any("images (xl/media)" in item for item in lost)


def test_chart_style_and_colour_parts_are_counted(tmp_path: Path) -> None:
    """openpyxl writes neither part back, and both carry visible formatting."""
    template = _package(
        tmp_path / "template.xlsx",
        {
            "xl/charts/style1.xml": b"<cs:chartStyle/>",
            "xl/charts/colors1.xml": b"<cs:colorStyle/>",
        },
    )
    output = _package(tmp_path / "output.xlsx", {"xl/workbook.xml": b"<workbook/>"})

    lost = find_lost_features(template, output)

    assert any("chart styles" in item for item in lost)
    assert any("chart colour schemes" in item for item in lost)


def test_the_lossy_warning_belongs_to_the_fillxl_warning_hierarchy(tmp_path: Path) -> None:
    """One ``warnings.simplefilter`` can turn every fillxl warning into an error."""
    from fillxl import FillXLWarning

    template = _package(tmp_path / "template.xlsx", {"xl/slicers/slicer1.xml": b"<slicer/>"})
    output = _package(tmp_path / "output.xlsx", {"xl/workbook.xml": b"<workbook/>"})

    with pytest.warns(FillXLWarning, match="slicers"):
        check_lossy(template, output, "warn")


def test_pictures_of_a_deleted_sheet_are_not_reported_as_lost(tmp_path: Path) -> None:
    """``delete_sheets`` removes a sheet's pictures on purpose, not by accident."""
    template = _picture_package(
        tmp_path / "template.xlsx", ["Report", "Scratch"], owners=["Scratch"]
    )
    output = _picture_package(tmp_path / "output.xlsx", ["Report"])

    assert find_lost_features(template, output) == [
        "images (xl/media) (1 lost)",
        "pictures (1 lost)",
    ]
    assert find_lost_features(template, output, deleted_sheets=["Scratch"]) == []


def test_a_sheet_name_that_did_not_match_excludes_nothing(tmp_path: Path) -> None:
    """Only the sheets that really went are excused."""
    template = _picture_package(
        tmp_path / "template.xlsx", ["Report", "Scratch"], owners=["Scratch"]
    )
    output = _picture_package(tmp_path / "output.xlsx", ["Report"])

    lost = find_lost_features(template, output, deleted_sheets=["Scratch2"])

    assert "pictures (1 lost)" in lost


def test_a_drawing_a_surviving_sheet_also_uses_still_counts(tmp_path: Path) -> None:
    """A part shared with a remaining sheet is not excused by the deletion."""
    template = _picture_package(
        tmp_path / "template.xlsx",
        ["Report", "Scratch"],
        owners=["Report", "Scratch"],
        pictures=2,
    )
    output = _picture_package(tmp_path / "output.xlsx", ["Report"], owners=["Report"])

    assert find_lost_features(template, output, deleted_sheets=["Scratch"]) == [
        "images (xl/media) (1 lost)",
        "pictures (1 lost)",
    ]


def test_excluding_sheets_from_a_non_package_is_harmless(tmp_path: Path) -> None:
    """A file that is not a zip counts nothing, with or without exclusions."""
    path = tmp_path / "not_a_zip.xlsx"
    path.write_bytes(b"not a zip")

    assert feature_counts(path, exclude_sheets=["Scratch"]) == Counter()
