"""Dates display as dates, with a format the template can override."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from fillxl import fill
from tests.parity.builders import named_cells_template, table_template


def _fill_table(
    tmp_path: Path,
    rows: list[dict[str, object]],
    *,
    number_formats: dict[int, str] | None = None,
) -> Path:
    template = table_template(
        tmp_path / "template.xlsx",
        header=("fund", "nav"),
        number_formats=number_formats,
    )
    output = tmp_path / "filled.xlsx"
    fill(template=template, output=output, engine="openpyxl", tables={"Sales": rows})
    return output


def _number_format(path: Path, coordinate: str) -> str:
    workbook = openpyxl.load_workbook(path)
    try:
        return str(workbook["Data"][coordinate].number_format)
    finally:
        workbook.close()


def test_a_date_in_a_general_cell_is_shown_without_a_time(tmp_path: Path) -> None:
    """A date must not arrive as a midnight timestamp."""
    output = _fill_table(tmp_path, [{"fund": "A", "nav": dt.date(2025, 1, 2)}])

    assert _number_format(output, "B2") == "yyyy-mm-dd"


def test_a_datetime_keeps_its_time_of_day(tmp_path: Path) -> None:
    """A datetime needs a format that shows the time it was given."""
    output = _fill_table(tmp_path, [{"fund": "A", "nav": dt.datetime(2025, 1, 2, 9, 30)}])

    assert _number_format(output, "B2") == "yyyy-mm-dd h:mm:ss"


def test_a_date_format_in_the_template_wins_over_the_default(tmp_path: Path) -> None:
    """The person who made the template already chose how the date looks."""
    output = _fill_table(
        tmp_path,
        [{"fund": "A", "nav": dt.date(2025, 1, 2)}],
        number_formats={2: "dd.mm.yyyy"},
    )

    assert _number_format(output, "B2") == "dd.mm.yyyy"


def _fill_named_cell(tmp_path: Path, value: object, *, template_format: str | None = None) -> Path:
    template = named_cells_template(tmp_path / "template.xlsx", cells={"report_date": ("A1", None)})
    if template_format is not None:
        workbook = openpyxl.load_workbook(template)
        workbook["Data"]["A1"].number_format = template_format
        workbook.save(template)
    output = tmp_path / "filled.xlsx"
    fill(
        template=template,
        output=output,
        engine="openpyxl",
        values={"report_date": value},
    )
    return output


def test_a_date_in_a_named_cell_is_shown_without_a_time(tmp_path: Path) -> None:
    output = _fill_named_cell(tmp_path, dt.date(2025, 1, 2))

    assert _number_format(output, "A1") == "yyyy-mm-dd"


def test_a_datetime_in_a_named_cell_keeps_its_time_of_day(tmp_path: Path) -> None:
    output = _fill_named_cell(tmp_path, dt.datetime(2025, 1, 2, 9, 30))

    assert _number_format(output, "A1") == "yyyy-mm-dd h:mm:ss"


def test_a_named_cell_keeps_a_date_format_from_the_template(tmp_path: Path) -> None:
    output = _fill_named_cell(tmp_path, dt.date(2025, 1, 2), template_format="dd.mm.yyyy")

    assert _number_format(output, "A1") == "dd.mm.yyyy"


@pytest.mark.parametrize("value", ["2025-01-02", "text"])
def test_a_named_cell_that_is_not_a_date_keeps_its_general_format(
    tmp_path: Path, value: str
) -> None:
    """Only dates get a date format."""
    output = _fill_named_cell(tmp_path, value)

    assert _number_format(output, "A1") == "General"
