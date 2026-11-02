"""One sheet editor and one style sheet per fill session.

Parsing a sheet is not free, and a cached editor that is rebuilt per write
throws away the earlier work — or, worse, writes stale text over a part another
operation already changed.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from fillxl import FillXLConfig, fill
from fillxl._ooxml.worksheet import SheetPart
from tests.ooxml._build import body_rows, named_cell_template, table_template

if TYPE_CHECKING:
    from pathlib import Path


def test_many_named_cells_on_one_sheet_parse_the_sheet_once(
    tmp_path: Path,
    monkeypatch: object,
) -> None:
    """Three writes to one sheet must not mean three parses."""
    template = named_cell_template(
        tmp_path / "template.xlsx",
        cells={"a_number": ("B2", 1), "another": ("B3", 2), "a_third": ("B4", 3)},
    )
    parses: list[str] = []
    original = SheetPart.__init__

    def counting(self: SheetPart, package: object, part: str) -> None:  # noqa: ANN401
        parses.append(part)
        original(self, package, part)  # type: ignore[arg-type]

    monkeypatch.setattr(SheetPart, "__init__", counting)  # type: ignore[attr-defined]

    fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        values={"a_number": 10, "another": 20, "a_third": 30},
    )

    assert parses == ["xl/worksheets/sheet1.xml"]


def test_values_survive_activating_a_sheet_after_the_fill(tmp_path: Path) -> None:
    """The activation changes parts directly; the session cache must not undo it."""
    template = table_template(tmp_path / "template.xlsx", rows=1, sum_sheet="Summary")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(3)},
        config=FillXLConfig(active_sheet="Summary"),
    )

    import openpyxl

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.active.title == "Summary"
        assert workbook["Data"]["A2"].value == "R0"
        assert workbook["Data"]["A4"].value == "R2"
    finally:
        workbook.close()
