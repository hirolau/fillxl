"""Integration tests for the ``on_lossy`` policy."""

from __future__ import annotations

import warnings
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import FillXLConfig, LossyTemplateError, LossyTemplateWarning, fill
from tests._assets import TEMPLATE

if TYPE_CHECKING:
    from collections.abc import Sequence

    from fillxl._lossy import OnLossy


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


def test_an_embedded_object_is_reported_and_keeps_the_previous_output(
    tmp_path: Path,
) -> None:
    """An embedded Word document or PDF would vanish without any warning.

    openpyxl cannot write ``xl/embeddings``, and once its drawing shape is
    gone nothing else in the package mentions it.
    """
    wb = openpyxl.Workbook()
    wb.active["A1"] = "x"
    template = tmp_path / "embedded.xlsx"
    wb.save(template)
    wb.close()
    _inject_part(template, "xl/embeddings/oleObject1.bin", b"\xd0\xcf\x11\xe0 fake package")

    output = tmp_path / "out.xlsx"
    output.write_bytes(b"old")

    with pytest.raises(LossyTemplateError, match="embedded objects"):
        fill(template=template, output=output, config=FillXLConfig(on_lossy="raise"))

    assert output.read_bytes() == b"old"

    with pytest.warns(LossyTemplateWarning, match="embedded objects"):
        fill(template=template, output=output)

    with zipfile.ZipFile(output) as archive:
        assert [n for n in archive.namelist() if n.startswith("xl/embeddings/")] == []


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


def test_the_openpyxl_backend_excuses_the_sheets_it_deleted(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A deleted sheet's pictures are its own, so ``on_lossy`` must not count them."""
    from fillxl._backends import _openpyxl

    recorded: dict[str, object] = {}
    real_check = _openpyxl.check_lossy

    def spy(
        template: Path,
        output: Path,
        policy: OnLossy,
        *,
        deleted_sheets: Sequence[str] = (),
    ) -> None:
        recorded["deleted"] = list(deleted_sheets)
        real_check(template, output, policy, deleted_sheets=deleted_sheets)

    monkeypatch.setattr(_openpyxl, "check_lossy", spy)
    workbook = openpyxl.Workbook()
    workbook.active.title = "Report"
    workbook.create_sheet("Scratch")["A1"] = "x"
    template = tmp_path / "template.xlsx"
    workbook.save(template)
    workbook.close()

    fill(
        template=template,
        output=tmp_path / "out.xlsx",
        engine="openpyxl",
        config=FillXLConfig(hide_underscore_sheets=False, delete_sheets=["Scratch"]),
    )

    assert recorded["deleted"] == ["Scratch"]
