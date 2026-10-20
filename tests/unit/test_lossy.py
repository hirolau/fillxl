"""Unit tests for the lossy-feature detector."""

from __future__ import annotations

import warnings
import zipfile
from collections import Counter
from pathlib import Path

import pytest

from fillxl._lossy import LossyTemplateError, check_lossy, feature_counts, find_lost_features


def _package(path: Path, parts: dict[str, bytes]) -> Path:
    """Write a synthetic OOXML package with *parts*."""
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


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
