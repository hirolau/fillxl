"""Chapter 15 — Custom Excel Formats (.xlsm, .xlsb, .xls, etc.).
=========================================================

fillxl supports standard ``.xlsx`` files with both backends, but other
formats (macro-enabled ``.xlsm``, binary ``.xlsb``, legacy ``.xls``, etc.)
require the win32com backend because openpyxl cannot handle them.

* openpyxl raises :class:`ValueError` for any non-``.xlsx`` file.
* win32com (Windows + Excel) opens and fills ``.xlsm`` and ``.xlsb`` successfully.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from fillxl import fill

ASSETS = Path(__file__).resolve().parent.parent / "assets"


# ═══════════════════════════════════════════════════════════════════════════════
# 15.1 openpyxl rejects non-.xlsx formats
# ═══════════════════════════════════════════════════════════════════════════════


class TestOpenpyxlRejectsUnsupportedFormats:
    """openpyxl fails fast with a clear error for any non-.xlsx file."""

    @pytest.mark.parametrize(
        ("template", "ext"),
        [
            pytest.param(ASSETS / "test_template_1_xlsb.xlsb", ".xlsb", id=".xlsb"),
            pytest.param(ASSETS / "test_template_1_slsm.xlsm", ".xlsm", id=".xlsm"),
        ],
    )
    def test_openpyxl_raises_valueerror_for_unsupported_format(
        self, tmp_path: Path, template: Path, ext: str
    ) -> None:
        """openpyxl raises ValueError with a message suggesting win32com."""
        output = tmp_path / "out.xlsx"
        with pytest.raises(ValueError, match="not supported by the openpyxl backend"):
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
        with pytest.raises(ValueError, match="not supported by the openpyxl backend"):
            fill(template=template, output=output, engine="openpyxl")

    def test_auto_engine_also_rejects_non_xlsx(self, tmp_path: Path) -> None:
        """'auto' maps to openpyxl, so .xlsb fails the same way."""
        template = ASSETS / "test_template_1_xlsb.xlsb"
        output = tmp_path / "out.xlsx"
        with pytest.raises(ValueError, match="not supported by the openpyxl backend"):
            fill(template=template, output=output, engine="auto")

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
