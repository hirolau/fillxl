"""The 1904 date system: a workbook can count days from 1904 instead of 1900.

Excel for Mac used to save workbooks with ``date1904="1"``, which moves every
serial number by 1462 days.  Writing 1900-based serials into such a workbook
makes every date jump four years forward the moment Excel opens it.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import CellValueError, fill
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


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_named_date_before_the_1904_epoch_is_refused_by_both_engines(
    engine: str,
    tmp_path: Path,
) -> None:
    """Excel shows a date before day zero as ####, so neither engine may write one."""
    template = date1904_template(tmp_path / "template.xlsx")

    with pytest.raises(CellValueError, match=r"1904-01-01.*1904 date system") as caught:
        fill(
            template,
            tmp_path / f"{engine}.xlsx",
            engine=engine,
            values={"a_date": dt.date(1903, 12, 31)},
        )

    assert "1903-12-31" in str(caught.value)
    assert not (tmp_path / f"{engine}.xlsx").exists()


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_table_date_before_the_1904_epoch_is_refused_by_both_engines(
    engine: str,
    tmp_path: Path,
) -> None:
    """A table cell is validated by the same rule as a named cell."""
    template = date1904_template(tmp_path / "template.xlsx")

    with pytest.raises(CellValueError, match=r"1904-01-01.*1904 date system"):
        fill(
            template,
            tmp_path / f"{engine}.xlsx",
            engine=engine,
            tables={"Data": body_rows(2, first_day=dt.date(1903, 12, 31))},
        )

    assert not (tmp_path / f"{engine}.xlsx").exists()


@pytest.mark.parametrize("engine", ["openpyxl", "fillxl"])
def test_a_date_on_the_1904_epoch_is_written(engine: str, tmp_path: Path) -> None:
    """Day zero of the 1904 system is a serial Excel can show."""
    template = date1904_template(tmp_path / "template.xlsx")

    output = fill(
        template,
        tmp_path / f"{engine}.xlsx",
        engine=engine,
        values={"a_date": dt.date(1904, 1, 1)},
    )

    workbook = openpyxl.load_workbook(output)
    try:
        # 1904-01-01 is serial 0, which openpyxl reads back as a time of day.
        assert workbook["Data"]["G1"].value == dt.time(0, 0)
    finally:
        workbook.close()
