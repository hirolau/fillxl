"""Integration tests for the ``on_lossy`` policy."""

from __future__ import annotations

import shutil
import warnings
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING, Any

import openpyxl
import pytest

from fillxl import FillXLConfig, LossyTemplateError, LossyTemplateWarning, fill
from tests._assets import ASSETS

MACRO_TEMPLATE = ASSETS / "test_template_1_slsm.xlsm"

if TYPE_CHECKING:
    from collections.abc import Iterator, Sequence

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


def _plain_template(path: Path) -> Path:
    """A template with nothing the openpyxl engine can lose."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "Name"
    worksheet["A2"] = "sample"
    workbook.save(path)
    workbook.close()
    return path


def test_a_plain_template_does_not_warn(tmp_path: Path) -> None:
    """An ordinary template has no features to lose."""
    template = _plain_template(tmp_path / "plain.xlsx")
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        fill(template=template, output=tmp_path / "out.xlsx")

    assert not [item for item in caught if "did not preserve" in str(item.message)]


def test_losing_cell_metadata_is_reported(tmp_path: Path) -> None:
    """``xl/metadata.xml`` carries spilling and data-type information.

    It is what makes an edited formula start spilling again, so a fill that
    drops it silently changes what the user's formulas do.
    """
    template = _plain_template(tmp_path / "metadata.xlsx")
    _inject_part(
        template,
        "xl/metadata.xml",
        b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    )

    with pytest.raises(LossyTemplateError, match="cell metadata"):
        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            config=FillXLConfig(on_lossy="raise"),
        )


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


def _vba_template(tmp_path: Path) -> Path:
    """A real macro-enabled template: the VBA project is what a format change drops."""
    target = tmp_path / "macro.xlsm"
    shutil.copyfile(MACRO_TEMPLATE, target)
    return target


@pytest.mark.win32com
def test_converting_a_macro_template_to_xlsx_reports_the_vba_loss(tmp_path: Path) -> None:
    """Excel writes back what it read, but ``.xlsx`` cannot hold a macro project."""
    from tests._engines import _fill_output

    template = _vba_template(tmp_path)
    output = tmp_path / "out.xlsx"

    with pytest.warns(LossyTemplateWarning, match="VBA project"):
        result = _fill_output(template, output, "win32com")

    assert result.suffix == ".xlsx"


@pytest.mark.win32com
def test_converting_a_macro_template_to_xlsx_can_fail_the_fill(tmp_path: Path) -> None:
    """``on_lossy="raise"`` has to apply to a conversion, not only to a lossy engine."""
    from tests._engines import _fill_output

    template = _vba_template(tmp_path)

    with pytest.raises(LossyTemplateError, match="VBA project"):
        _fill_output(
            template,
            tmp_path / "out.xlsx",
            "win32com",
            config=FillXLConfig(on_lossy="raise"),
        )


@pytest.fixture(scope="module")
def excel() -> Iterator[Any]:
    """A hidden Excel instance, for the conversions only Excel can perform."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    try:
        yield app
    finally:
        app.Quit()
        pythoncom.CoUninitialize()


def _converted(template: Path, *, suffix: str, file_format: int, excel: Any) -> Path:  # noqa: ANN401
    """*template* saved again in another format by Excel, macros included."""
    target = template.with_name(f"{template.stem}{suffix}")
    workbook = excel.Workbooks.Open(
        str(template),
        UpdateLinks=0,
        ReadOnly=False,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False,
    )
    try:
        workbook.SaveAs(str(target), FileFormat=file_format)
    finally:
        workbook.Close(False)
    return target


def _conversion_warnings(caught: list[warnings.WarningMessage]) -> list[str]:
    return [
        str(record.message) for record in caught if isinstance(record.message, LossyTemplateWarning)
    ]


@pytest.mark.win32com
def test_a_macro_template_saved_as_xls_reports_what_biff8_drops(
    tmp_path: Path,
) -> None:
    """.xls keeps macros, so VBA is not a loss -- but dynamic arrays are.

    The message must name the features the format really cannot hold, and must
    not claim a VBA loss the file does not have.
    """
    from tests._engines import _fill_output

    template = _vba_template(tmp_path)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        _fill_output(template, tmp_path / "out.xls", "win32com")

    messages = _conversion_warnings(caught)
    assert messages, "an .xls output should report what it cannot hold"
    assert any("cell metadata" in message for message in messages)
    assert not any("VBA" in message for message in messages)


@pytest.mark.win32com
def test_an_xlsb_template_saved_as_xlsm_loses_nothing(
    tmp_path: Path,
    excel: Any,  # noqa: ANN401
) -> None:
    """Binary OOXML holds the same feature set; nothing may be reported."""
    from tests._engines import _fill_output

    template = _converted(
        _vba_template(tmp_path),
        suffix=".xlsb",
        file_format=50,
        excel=excel,
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        _fill_output(template, tmp_path / "out.xlsm", "win32com")

    assert _conversion_warnings(caught) == []


@pytest.mark.win32com
def test_an_xls_template_can_still_lose_its_macros_to_xlsx(
    tmp_path: Path,
    excel: Any,  # noqa: ANN401
) -> None:
    """A BIFF8 template cannot be read part by part, so COM has to say it has VBA."""
    from tests._engines import _fill_output

    template = _converted(
        _vba_template(tmp_path),
        suffix=".xls",
        file_format=56,
        excel=excel,
    )

    with pytest.warns(LossyTemplateWarning, match="VBA project"):
        _fill_output(template, tmp_path / "out.xlsx", "win32com")
