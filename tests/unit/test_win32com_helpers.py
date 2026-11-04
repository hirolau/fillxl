"""Pure decisions the win32com engine makes about a table's shape.

They are separated from COM so they can be tested without Excel, which is
where the interesting cases are: a headerless table has no header row at all.
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from fillxl._backends._win32com import _is_header_cell, _merged_regions_com


def test_a_cell_on_the_header_row_of_a_table_with_headers_is_a_header_cell() -> None:
    table = SimpleNamespace(ShowHeaders=True, HeaderRowRange=SimpleNamespace(Row=1))

    assert _is_header_cell(SimpleNamespace(Row=1), table) is True


def test_a_cell_in_the_body_of_a_table_with_headers_is_not_a_header_cell() -> None:
    table = SimpleNamespace(ShowHeaders=True, HeaderRowRange=SimpleNamespace(Row=1))

    assert _is_header_cell(SimpleNamespace(Row=2), table) is False


def test_a_cell_in_a_headerless_table_is_never_a_header_cell() -> None:
    table = SimpleNamespace(ShowHeaders=False, HeaderRowRange=None)

    assert _is_header_cell(SimpleNamespace(Row=1), table) is False


class _FakeRange:
    def __init__(self, *, address: str = "A2:A2", merged: bool = False) -> None:
        self.Address = address
        self.MergeCells = None if merged else False
        self.Row = 2
        self.Column = 1
        self.Rows = SimpleNamespace(Count=1)
        self.Columns = SimpleNamespace(Count=1)
        self.Cells = SimpleNamespace(Count=2 if merged else 1)


class _FakeCell:
    def __init__(self, merge_area: _FakeRange) -> None:
        self.MergeArea = merge_area


class _FakeWorksheet:
    """``MergeCells`` is ``None`` (a mix of merged and plain cells)."""

    def __init__(self, merged: tuple[int, int]) -> None:
        self.merged = merged
        self.cells_read: list[tuple[int, int]] = []
        self._area = _FakeRange(merged=True)

    def Range(  # noqa: N802 - the name has to match Excel's object model
        self,
        top_left: _FakeCell,
        bottom_right: _FakeCell,
    ) -> _FakeRange:
        del top_left, bottom_right
        area = _FakeRange()
        area.MergeCells = None
        return area

    def Cells(self, row: int, column: int) -> _FakeCell:  # noqa: N802 - Excel's method
        self.cells_read.append((row, column))
        if (row, column) == self.merged:
            return _FakeCell(self._area)
        return _FakeCell(_FakeRange())


def test_the_first_merged_region_found_is_the_one_reported() -> None:
    worksheet = _FakeWorksheet(merged=(2, 1))

    regions = _merged_regions_com(worksheet, first_row=1, last_row=4, first_col=1, last_col=3)

    assert len(regions) == 1
    assert regions[0].first_row == 2


def test_reading_stops_at_the_first_merged_region() -> None:
    """One merged range is enough to refuse the fill, so the rest is not visited."""
    worksheet = _FakeWorksheet(merged=(1, 2))

    _merged_regions_com(worksheet, first_row=1, last_row=10, first_col=1, last_col=5)

    # The two corners build the range; the merged cell is the first one read.
    assert worksheet.cells_read == [(1, 1), (10, 5), (1, 1), (1, 2)]


# -------------------------------------------------------------------------------
# Calculated-column definitions are read from the package, not from openpyxl
# -------------------------------------------------------------------------------

_TESTS = Path(__file__).resolve().parents[1]
_TEMPLATES = sorted(
    path
    for folder in (_TESTS / "assets", _TESTS / "fixtures")
    for path in folder.glob("*")
    if path.suffix.lower() in {".xlsx", ".xlsm"}
)


def _definitions_from_openpyxl(path: Path) -> dict[str, dict[str, str]] | None:
    """The reading the function used to do, kept as the reference to compare with."""
    import openpyxl

    from fillxl._formulas import formula_definitions

    try:
        workbook = openpyxl.load_workbook(path)
    except Exception:
        return None
    try:
        return {
            str(table.displayName or table.name): formula_definitions(table)
            for worksheet in workbook.worksheets
            for table in worksheet.tables.values()
        }
    finally:
        workbook.close()


@pytest.mark.parametrize("template", _TEMPLATES, ids=lambda path: path.name)
def test_the_definitions_read_from_the_package_match_the_openpyxl_reading(
    template: Path,
) -> None:
    """A whole workbook is no longer loaded just to read a table's columns."""
    from fillxl._backends._win32com import _read_table_formula_definitions

    assert _read_table_formula_definitions(template) == _definitions_from_openpyxl(template)


def test_a_file_that_cannot_be_read_returns_none() -> None:
    """A damaged package falls back to detecting calculated columns from cells."""
    from fillxl._backends._win32com import _read_table_formula_definitions

    assert _read_table_formula_definitions(Path("missing-file.xlsx")) is None


def test_a_binary_template_is_not_read_for_table_definitions(tmp_path: Path) -> None:
    """``.xlsb`` is not an OOXML package, so there is nothing to read."""
    from fillxl._backends._win32com import _read_table_formula_definitions

    binary = tmp_path / "template.xlsb"
    binary.write_bytes(b"not a zip")

    assert _read_table_formula_definitions(binary) is None
