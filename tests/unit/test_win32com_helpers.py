"""Pure decisions the win32com engine makes about a table's shape.

They are separated from COM so they can be tested without Excel, which is
where the interesting cases are: a headerless table has no header row at all.
"""

from __future__ import annotations

from types import SimpleNamespace

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
