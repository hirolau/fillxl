"""A table's sort state has to stay inside the table body.

``<sortState ref="…">`` records the range that was last sorted.  When a table
shrinks, a stale range points at rows the table no longer has, and Excel offers
to sort cells that are no longer part of the table.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
from openpyxl.worksheet.filters import SortState

from fillxl import fill
from tests.ooxml._build import body_rows, table_template
from tests.ooxml._patch import part_text

if TYPE_CHECKING:
    from pathlib import Path

TABLE_PART = "xl/tables/table1.xml"


def test_a_sorted_table_keeps_a_sort_range_inside_its_body_after_shrinking(
    tmp_path: Path,
) -> None:
    """The sort range follows the body down to its new last row."""
    template = table_template(tmp_path / "template.xlsx", rows=4)
    workbook = openpyxl.load_workbook(template)
    try:
        table = workbook["Data"].tables["Data"]
        table.sortState = SortState(ref="A1:E5")
        workbook.save(template)
    finally:
        workbook.close()

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(2)},
    )

    table_xml = part_text(output, TABLE_PART)
    assert '<sortState ref="A2:E3"' in table_xml
