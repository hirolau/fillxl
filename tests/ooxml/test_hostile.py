"""Workbooks written by other programs must fill cleanly and read the same.

Each fixture in :mod:`tests.ooxml.fixtures.builders` is a valid workbook whose
markup the fillxl engine has never seen: a different namespace prefix, CDATA,
an ``extLst`` full of lookalike element names, missing cell references, missing
sheet views, LibreOffice's namespaces, xlsxwriter's tables.
"""

from __future__ import annotations

import importlib.util
from typing import TYPE_CHECKING

import pytest

from fillxl import fill
from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.worksheet import SheetPart
from tests.ooxml._build import body_rows
from tests.ooxml.fixtures.builders import BUILDERS, SHEET_PART
from tests.ooxml.test_parity import _differences, fill_both

if TYPE_CHECKING:
    from collections.abc import Callable
    from pathlib import Path

_NEEDS_XLSXWRITER = pytest.mark.skipif(
    importlib.util.find_spec("xlsxwriter") is None,
    reason="needs xlsxwriter",
)

_HOSTILE = [
    pytest.param("prefixed", id="prefixed"),
    pytest.param("cdata", id="cdata"),
    pytest.param("extlst_lookalikes", id="extlst-lookalikes"),
    pytest.param(
        "no_cell_refs",
        id="no-cell-refs",
        marks=pytest.mark.xfail(reason="Task 9: cells without an r attribute"),
    ),
    pytest.param("no_sheet_views", id="no-sheet-views"),
    pytest.param("libreoffice", id="libreoffice"),
    pytest.param("streaming_writer", id="streaming-writer", marks=_NEEDS_XLSXWRITER),
]

_PARITY = [
    param for param in _HOSTILE if param.id != "no-sheet-views"
] + [
    pytest.param(
        "no_sheet_views",
        id="no-sheet-views",
        marks=pytest.mark.xfail(
            reason="openpyxl cannot read a chartsheet without <sheetViews>, "
            "so there is no second engine to compare with",
        ),
    ),
]


@pytest.mark.parametrize("name", _HOSTILE)
def test_every_hostile_fixture_fills_cleanly(
    name: str,
    hostile_fixture_factory: Callable[[str], Path],
    tmp_path: Path,
    assert_well_formed: Callable[[Path], None],
) -> None:
    """A valid workbook must survive a fill whatever writer produced it."""
    template = hostile_fixture_factory(name)

    output = fill(
        template,
        tmp_path / template.name,
        engine="fillxl",
        tables={"Data": body_rows(3)},
    )

    assert_well_formed(output)


@pytest.mark.parametrize("name", _PARITY)
def test_every_hostile_fixture_matches_openpyxl(
    name: str,
    hostile_fixture_factory: Callable[[str], Path],
    tmp_path: Path,
) -> None:
    """The fill of a foreign workbook must read the way openpyxl's fill does."""
    template = hostile_fixture_factory(name)

    expected, actual = fill_both(template, tmp_path, tables={"Data": body_rows(3)})

    differences = _differences(expected, actual)
    assert not differences, "The engines disagree:\n" + "\n".join(differences[:50])


def test_an_extlst_row_is_not_treated_as_a_sheet_row(
    hostile_fixture_factory: Callable[[str], Path],
) -> None:
    """``_ROW`` and ``_CELL`` must only ever see the ``<sheetData>`` body.

    An extension can hold ``<foo:row r="77"><foo:c r="Z98"/></foo:row>``; if the
    parser looked at the whole part, the fill would believe row 77 exists.
    """
    assert "extlst_lookalikes" in BUILDERS
    template = hostile_fixture_factory("extlst_lookalikes")

    sheet = SheetPart(OoxmlPackage(template), SHEET_PART)

    assert sheet.row_numbers() == [1, 2, 3]
    assert sheet.cell(77, 26) is None
