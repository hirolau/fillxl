"""Pure decisions the win32com engine makes about a table's shape.

They are separated from COM so they can be tested without Excel, which is
where the interesting cases are: a headerless table has no header row at all.
"""

from __future__ import annotations

from types import SimpleNamespace

from fillxl._backends._win32com import _is_header_cell


def test_a_cell_on_the_header_row_of_a_table_with_headers_is_a_header_cell() -> None:
    table = SimpleNamespace(ShowHeaders=True, HeaderRowRange=SimpleNamespace(Row=1))

    assert _is_header_cell(SimpleNamespace(Row=1), table) is True


def test_a_cell_in_the_body_of_a_table_with_headers_is_not_a_header_cell() -> None:
    table = SimpleNamespace(ShowHeaders=True, HeaderRowRange=SimpleNamespace(Row=1))

    assert _is_header_cell(SimpleNamespace(Row=2), table) is False


def test_a_cell_in_a_headerless_table_is_never_a_header_cell() -> None:
    table = SimpleNamespace(ShowHeaders=False, HeaderRowRange=None)

    assert _is_header_cell(SimpleNamespace(Row=1), table) is False
