"""Every fillxl output has to open in real Excel without a repair prompt.

The XML can be well-formed and still be a workbook Excel refuses; only Excel can
say.  This runs on the self-hosted Windows runner that has Excel installed (the
nightly parity job), and skips everywhere else.

``CorruptLoad=0`` means "never repair": Excel reports a broken file instead of
quietly dropping the parts it dislikes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

import pytest

from fillxl import fill
from tests._engines import _win32com_available
from tests.ooxml._build import body_rows
from tests.ooxml.test_hostile import _HOSTILE
from tests.ooxml.test_parity import SCENARIOS

if TYPE_CHECKING:
    from collections.abc import Callable, Iterator, Sequence
    from pathlib import Path

pytestmark = [pytest.mark.excel, pytest.mark.win32com]

if not _win32com_available():  # pragma: no cover - depends on the machine
    pytest.skip("needs Windows with Microsoft Excel", allow_module_level=True)


@pytest.fixture(scope="module")
def excel() -> Iterator[Any]:
    """A hidden Excel instance that never asks the user anything."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        yield app
    finally:
        app.Quit()
        pythoncom.CoUninitialize()


def _opens_without_repair(excel: Any, path: Path, expected: Sequence[str]) -> None:  # noqa: ANN401
    """Open *path* with corruption-repair disabled; Excel must show every sheet."""
    workbook = excel.Workbooks.Open(
        Filename=str(path),
        UpdateLinks=0,
        ReadOnly=True,
        CorruptLoad=0,
    )
    try:
        names = [
            str(workbook.Sheets(index + 1).Name) for index in range(int(workbook.Sheets.Count))
        ]
    finally:
        workbook.Close(False)
    assert names == list(expected)


def _expected_sheets(template: Path, kwargs: dict[str, Any]) -> list[str]:
    """The sheets the fill leaves behind, as the template itself lists them."""
    from fillxl._ooxml.package import OoxmlPackage
    from fillxl._ooxml.workbook import WorkbookPart

    names = [entry.name for entry in WorkbookPart(OoxmlPackage(template)).sheets()]
    config = kwargs.get("config")
    deleted = set(config.delete_sheets or ()) if config is not None else set()
    return [name for name in names if name not in deleted]


_EXCEL_SCENARIOS = [
    # Same scenarios, without their marks: a scenario that cannot be compared
    # between the two engines can still open cleanly in Excel.
    pytest.param(param.values[0], id=param.id)
    for param in SCENARIOS
]


@pytest.mark.parametrize("build", _EXCEL_SCENARIOS)
def test_a_parity_output_opens_in_excel_without_repair(
    build: Callable[[Path], tuple[Path, dict[str, Any]]],
    tmp_path: Path,
    excel: Any,  # noqa: ANN401
) -> None:
    """Each parity scenario, filled by fillxl, opens cleanly."""
    template, kwargs = build(tmp_path)
    expected = _expected_sheets(template, kwargs)

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", **kwargs)

    _opens_without_repair(excel, output, expected)


_EXCEL_FIXTURES = [
    param
    if param.id != "no-sheet-views"
    else pytest.param(
        "no_sheet_views",
        id=param.id,
        marks=pytest.mark.skip(
            reason="Excel refuses the template itself: a chartsheet needs "
            "<sheetViews>, which is exactly what this fixture removes",
        ),
    )
    for param in _HOSTILE
]


@pytest.mark.parametrize("name", _EXCEL_FIXTURES)
def test_a_hostile_fixture_output_opens_in_excel_without_repair(
    name: str,
    hostile_fixture_factory: Callable[[str], Path],
    tmp_path: Path,
    excel: Any,  # noqa: ANN401
) -> None:
    """A workbook another writer produced still opens after a fillxl fill."""
    template = hostile_fixture_factory(name)
    expected = _expected_sheets(template, {})

    output = fill(
        template,
        tmp_path / template.name,
        engine="fillxl",
        tables={"Data": body_rows(3)},
    )

    _opens_without_repair(excel, output, expected)
