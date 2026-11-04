"""Text that looks like an Excel escape sequence is stored as text.

Excel decodes ``_xHHHH_`` in a cell string when it reads the file, so writing
``file_x0020_name`` as-is shows ``file name``.  A user's text has to survive a
round trip through Excel, which means the offline engines escape the leading
underscore the way Excel's own writer does.
"""

from __future__ import annotations

import contextlib
import importlib
import time
from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn

from fillxl import FillXLConfig
from tests._engines import ENGINES, _fill_output, _win32com_available
from tests.ooxml._patch import part_text

if TYPE_CHECKING:
    from pathlib import Path
    from typing import Any

    from fillxl import Engine

SHEET_PART = "xl/worksheets/sheet1.xml"

TYPED = "a_x0041_b"
"""What the user types, and what Excel has to show."""

ESCAPED = "a_x005F_x0041_b"
"""What the file has to hold for Excel to show *TYPED*."""


def _template(path: Path) -> Path:
    """A one-row table with a text column."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"], worksheet["B1"] = "Name", "Note"
    worksheet["A2"], worksheet["B2"] = "row2", "sample"
    worksheet.add_table(
        Table(
            displayName="Data",
            ref="A1:B2",
            tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Note")],
        ),
    )
    workbook.save(path)
    workbook.close()
    return path


def _named_template(path: Path) -> Path:
    """A workbook with one named cell in ``Data``."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "sample"
    workbook.defined_names.add(DefinedName("note", attr_text="Data!$A$1"))
    workbook.save(path)
    workbook.close()
    return path


def _value_as_excel_shows_it(path: Path, cell: str) -> object:
    """The value of ``Data!*cell*`` as Excel reads it back.

    A new Excel is started after the fill.  COM can hand back the engine's own
    instance while it is still shutting down, and a dying instance cannot open
    a workbook, so the attempt is repeated once the old one is gone.
    """
    if not _win32com_available():
        pytest.skip("win32com backend requires Windows and Excel")
    pythoncom = importlib.import_module("pythoncom")
    win32com_client = importlib.import_module("win32com.client")
    pythoncom.CoInitialize()
    failure: Exception | None = None
    for attempt in range(3):
        app = win32com_client.DispatchEx("Excel.Application")
        try:
            return _read_value(app, path, cell)
        except Exception as exc:  # a dying instance cannot open a workbook
            failure = exc
            time.sleep(1 + attempt)
        finally:
            with contextlib.suppress(Exception):
                app.Quit()
    raise AssertionError(f"Could not read {cell} back through Excel: {failure}")


def _read_value(app: Any, path: Path, cell: str) -> object:  # noqa: ANN401
    """Open *path* read-only and return the value of ``Data!*cell*``."""
    app.Visible = False
    app.DisplayAlerts = False
    workbook = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
    try:
        return workbook.Worksheets("Data").Range(cell).Value
    finally:
        workbook.Close(SaveChanges=False)


@pytest.mark.parametrize("engine", ENGINES)
def test_text_that_looks_like_an_excel_escape_is_written_as_typed(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """``a_x0041_b`` must not reach Excel as ``aAb``."""
    template = _template(tmp_path / "template.xlsx")

    output = _fill_output(
        template,
        tmp_path / "out.xlsx",
        engine,
        tables={"Data": [{"Name": "n", "Note": TYPED}]},
    )

    if engine == "win32com":
        assert _value_as_excel_shows_it(output, "B2") == TYPED
    else:
        sheet = part_text(output, SHEET_PART)
        assert ESCAPED in sheet
        assert f">{TYPED}<" not in sheet


@pytest.mark.parametrize("engine", ENGINES)
def test_named_text_that_looks_like_an_excel_escape_is_written_as_typed(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """The same rule for a named cell."""
    template = _named_template(tmp_path / "template.xlsx")

    output = _fill_output(template, tmp_path / "out.xlsx", engine, values={"note": TYPED})

    if engine == "win32com":
        assert _value_as_excel_shows_it(output, "A1") == TYPED
    else:
        assert ESCAPED in part_text(output, SHEET_PART)


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_formula_is_never_escaped(engine: Engine, tmp_path: Path) -> None:
    """A formula is markup Excel reads, not text it decodes."""
    template = _template(tmp_path / "template.xlsx")

    output = _fill_output(
        template,
        tmp_path / "out.xlsx",
        engine,
        tables={"Data": [{"Name": "n", "Note": '=B1&"_x0041_"'}]},
        config=FillXLConfig(escape_formulas=False),
    )

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook["Data"]["B2"].value == '=B1&"_x0041_"'
    finally:
        workbook.close()


def test_the_escaping_helper_leaves_plain_text_alone() -> None:
    """Only the leading underscore of a real escape is escaped."""
    from fillxl._ooxml.xmltext import escape_ooxml_escapes

    assert escape_ooxml_escapes("file name") == "file name"
    assert escape_ooxml_escapes("_x005F_x0041_") == "_x005F_x005F_x0041_"
    assert escape_ooxml_escapes("a_x0041_b") == "a_x005F_x0041_b"
