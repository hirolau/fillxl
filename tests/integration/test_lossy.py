"""Integration tests for the ``on_lossy`` policy."""

from __future__ import annotations

import warnings
import zipfile
from pathlib import Path

import openpyxl
import pytest

from fillxl import FillXLConfig, LossyTemplateError, fill
from tests.conftest import TEMPLATE


def _inject_part(path: Path, part: str, data: bytes) -> None:
    """Add *part* to the OOXML package at *path*."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
        target.writestr(part, data)
    tmp.replace(path)


def _template_with_slicer(path: Path) -> Path:
    wb = openpyxl.Workbook()
    wb.active["A1"] = "x"
    wb.save(path)
    wb.close()
    _inject_part(path, "xl/slicers/slicer1.xml", b"<slicers/>")
    return path


def test_on_lossy_raise_keeps_the_previous_output(tmp_path: Path) -> None:
    """The lossy check runs before the swap, so "raise" keeps the old output."""
    template = _template_with_slicer(tmp_path / "slicer.xlsx")
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"old")

    with pytest.raises(LossyTemplateError, match="slicers"):
        fill(template=template, output=output, config=FillXLConfig(on_lossy="raise"))

    assert output.read_bytes() == b"old"


def test_a_plain_template_does_not_warn(tmp_path: Path) -> None:
    """An ordinary template has no features to lose."""
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        fill(
            template=TEMPLATE,
            output=tmp_path / "out.xlsx",
            values={"single_cell": "x"},
        )

    assert not [item for item in caught if "did not preserve" in str(item.message)]


def test_on_lossy_rejects_unknown_values() -> None:
    """The policy is validated when the config is built."""
    with pytest.raises(ValueError):
        FillXLConfig(
            on_lossy="oops",  # ty: ignore[invalid-argument-type]  # pyrefly: ignore[bad-argument-type]
        )
