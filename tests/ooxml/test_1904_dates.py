"""The 1904 date system: a workbook can count days from 1904 instead of 1900.

Excel for Mac used to save workbooks with ``date1904="1"``, which moves every
serial number by 1462 days.  Writing 1900-based serials into such a workbook
makes every date jump four years forward the moment Excel opens it.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

import openpyxl

from fillxl import fill
from tests.ooxml._build import body_rows, date1904_template

if TYPE_CHECKING:
    from pathlib import Path


def test_dates_written_into_a_1904_workbook_read_back_as_the_same_dates(
    tmp_path: Path,
) -> None:
    """A date column and a dated named cell must survive the epoch change."""
    template = date1904_template(tmp_path / "template.xlsx")

    output = fill(
        template,
        tmp_path / "out.xlsx",
        engine="fillxl",
        tables={"Data": body_rows(2)},
        values={"a_date": dt.date(2025, 5, 5)},
    )

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["Data"]
        assert [sheet.cell(row=row, column=3).value for row in (2, 3)] == [
            dt.datetime(2024, 1, 1),
            dt.datetime(2024, 1, 2),
        ]
        assert sheet["G1"].value == dt.datetime(2025, 5, 5)
    finally:
        workbook.close()
