"""Chapter 15 — Custom Excel Formats (.xlsm, .xlsb, .xls, etc.).
=========================================================

fillxl supports ``.xlsx`` and macro-enabled ``.xlsm`` with the openpyxl
backend.  Binary ``.xlsb`` and legacy ``.xls`` require win32com.

* openpyxl raises :class:`ValueError` for any other template format, and it
  never converts between formats.
* win32com (Windows + Excel) opens and fills ``.xlsm`` and ``.xlsb`` successfully.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from fillxl import fill

ASSETS = Path(__file__).resolve().parent.parent / "assets"


# ═══════════════════════════════════════════════════════════════════════════════
# 15.1 openpyxl rejects unsupported formats
# ═══════════════════════════════════════════════════════════════════════════════


class TestOpenpyxlRejectsUnsupportedFormats:
    """openpyxl fails fast with a clear error for unsupported files."""

    @pytest.mark.parametrize(
        ("template", "ext"),
        [
            pytest.param(ASSETS / "test_template_1_xlsb.xlsb", ".xlsb", id=".xlsb"),
        ],
    )
    def test_openpyxl_raises_valueerror_for_unsupported_format(
        self, tmp_path: Path, template: Path, ext: str
    ) -> None:
        """openpyxl raises ValueError with a message suggesting win32com."""
        output = tmp_path / "out.xlsx"
        with pytest.raises(ValueError, match="cannot read a"):
            fill(template=template, output=output, engine="openpyxl")

    @pytest.mark.parametrize(
        "ext",
        [".xls", ".xltx", ".xltm", ".csv", ".txt", ".zip", ".ods", ".pdf"],
    )
    def test_openpyxl_rejects_any_non_xlsx_extension(self, tmp_path: Path, ext: str) -> None:
        """Even made-up files with non-.xlsx extensions raise ValueError."""
        template = tmp_path / f"dummy{ext}"
        template.write_text("not a real excel file")
        output = tmp_path / "out.xlsx"
        with pytest.raises(ValueError, match="cannot read a"):
            fill(template=template, output=output, engine="openpyxl")

    def test_auto_engine_hands_a_non_xlsx_template_to_excel(self, tmp_path: Path) -> None:
        """'auto' sends .xlsb to Excel; without Excel, openpyxl explains why."""
        template = ASSETS / "test_template_1_xlsb.xlsb"
        output = tmp_path / "out.xlsb"
        try:
            fill(template=template, output=output, engine="auto")
        except ValueError as exc:
            assert "cannot read a" in str(exc)
            pytest.skip("needs Excel to fill a .xlsb template")
        assert output.exists()

    def test_openpyxl_accepts_xlsx(self, tmp_path: Path) -> None:
        """Plain .xlsx works fine with openpyxl."""
        template = ASSETS / "test_template_1.xlsx"
        output = tmp_path / "out.xlsx"
        result = fill(
            template=template,
            output=output,
            values={"single_cell": "works"},
            engine="openpyxl",
        )
        assert result == output.resolve()
        assert output.exists()


# ═══════════════════════════════════════════════════════════════════════════════
# 15.2 openpyxl fills .xlsm and never converts formats
# ═══════════════════════════════════════════════════════════════════════════════


class TestOpenpyxlXlsm:
    """The openpyxl backend keeps a macro-enabled template macro-enabled."""

    def test_openpyxl_fills_an_xlsm_template_and_keeps_its_vba(self, tmp_path: Path) -> None:
        template = ASSETS / "test_template_1_slsm.xlsm"
        output = tmp_path / "filled.xlsm"

        fill(
            template=template,
            output=output,
            values={"single_cell": "works"},
            engine="openpyxl",
        )

        with zipfile.ZipFile(output) as archive:
            assert "xl/vbaProject.bin" in archive.namelist()
            content_types = archive.read("[Content_Types].xml").decode()
        assert "macroEnabled" in content_types

    def test_openpyxl_refuses_to_save_an_xlsx_template_as_xlsm(self, tmp_path: Path) -> None:
        output = tmp_path / "out.xlsm"
        with pytest.raises(ValueError, match="cannot write"):
            fill(
                template=ASSETS / "test_template_1.xlsx",
                output=output,
                values={"single_cell": "x"},
                engine="openpyxl",
            )
        assert not output.exists()

    def test_openpyxl_refuses_to_save_an_xlsm_template_as_xlsx(self, tmp_path: Path) -> None:
        output = tmp_path / "out.xlsx"
        with pytest.raises(ValueError, match="cannot write"):
            fill(
                template=ASSETS / "test_template_1_slsm.xlsm",
                output=output,
                values={"single_cell": "x"},
                engine="openpyxl",
            )
        assert not output.exists()

    def test_openpyxl_output_asks_excel_to_recalculate_on_open(self, tmp_path: Path) -> None:
        output = tmp_path / "out.xlsx"
        fill(
            template=ASSETS / "test_template_1.xlsx",
            output=output,
            values={"single_cell": "x"},
            engine="openpyxl",
        )

        with zipfile.ZipFile(output) as archive:
            assert b'fullCalcOnLoad="1"' in archive.read("xl/workbook.xml")


# ═══════════════════════════════════════════════════════════════════════════════
# 15.2 win32com handles .xlsm and .xlsb
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.win32com
class TestWin32comHandlesCustomFormats:
    """win32com fills both .xlsm and .xlsb (Windows + Excel required)."""

    @pytest.mark.parametrize(
        "template",
        [
            pytest.param(ASSETS / "test_template_1_xlsb.xlsb", id=".xlsb"),
            pytest.param(ASSETS / "test_template_1_slsm.xlsm", id=".xlsm"),
        ],
    )
    def test_win32com_fills_xlsb_and_xlsm(self, tmp_path: Path, template: Path) -> None:
        """win32com succeeds where openpyxl fails."""
        output = tmp_path / "filled.xlsx"
        result = fill(
            template=template,
            output=output,
            values={"single_cell": "Filled by win32com"},
            engine="win32com",
        )
        assert result == output.resolve()
        assert output.exists()
        assert output.stat().st_size > 0
