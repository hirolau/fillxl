"""Chart series references follow their table in the openpyxl engine."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
import pytest

from fillxl import fill
from tests.parity.builders import chart_template

_ROWS = [{"fund": f"F{i}", "nav": i + 1, "alt": i + 2} for i in range(5)]


def _fill(template: Path, tmp_path: Path) -> Path:
    output = tmp_path / "filled.xlsx"
    fill(
        template=template,
        output=output,
        engine="openpyxl",
        tables={"Sales": _ROWS},
    )
    return output


def _chart(path: Path) -> Any:  # noqa: ANN401 - openpyxl chart classes vary by type
    workbook = openpyxl.load_workbook(path)
    try:
        for name in workbook.sheetnames:
            charts = getattr(workbook[name], "_charts", [])
            if charts:
                return charts[0]
    finally:
        workbook.close()
    msg = "no chart in the workbook"
    raise AssertionError(msg)


def _value_reference(path: Path) -> str:
    chart = _chart(path)
    return str(chart.series[0].val.numRef.f)


def test_a_chart_on_a_one_row_table_grows_into_a_range_not_another_cell(tmp_path: Path) -> None:
    """Excel writes a one-row table as a single cell; the fill must widen it to a range."""
    template = chart_template(tmp_path / "template.xlsx", reference="Data!$B$2")

    assert _value_reference(_fill(template, tmp_path)) == "Data!$B$2:$B$6"


def test_a_non_contiguous_series_grows_with_its_table(tmp_path: Path) -> None:
    """A parenthesised list of areas is one series, and every area follows the table."""
    template = chart_template(
        tmp_path / "template.xlsx",
        header=("fund", "nav", "alt"),
        body=(("A", 1, 2),),
        reference="(Data!$B$2,Data!$C$2)",
    )

    assert _value_reference(_fill(template, tmp_path)) == "(Data!$B$2:$B$6,Data!$C$2:$C$6)"


def test_a_chart_on_a_swedish_sheet_follows_its_table(tmp_path: Path) -> None:
    """An unquoted non-ASCII sheet name is a sheet name like any other."""
    template = chart_template(
        tmp_path / "template.xlsx",
        sheet="Försäljning",
        reference="Försäljning!$B$2",
    )

    assert _value_reference(_fill(template, tmp_path)) == "Försäljning!$B$2:$B$6"


def test_a_chart_on_a_sheet_with_an_apostrophe_follows_its_table(tmp_path: Path) -> None:
    """The doubled quote inside a quoted sheet name is part of the name."""
    template = chart_template(
        tmp_path / "template.xlsx",
        sheet="Bob's Data",
        reference="'Bob''s Data'!$B$2",
    )

    assert _value_reference(_fill(template, tmp_path)) == "'Bob''s Data'!$B$2:$B$6"


def test_a_chart_on_its_own_chart_sheet_follows_its_table(tmp_path: Path) -> None:
    """A chart that lives on a chart sheet is updated too."""
    template = chart_template(tmp_path / "template.xlsx", chart_sheet=True)

    assert _value_reference(_fill(template, tmp_path)) == "Data!$B$2:$B$6"


def test_the_cached_point_count_matches_the_new_range(tmp_path: Path) -> None:
    """A stale ptCount makes Excel draw the wrong number of points."""
    template = chart_template(tmp_path / "template.xlsx", cache_points=1)

    chart = _chart(_fill(template, tmp_path))

    assert chart.series[0].val.numRef.numCache.ptCount == 5


@pytest.mark.parametrize("target", ["Data2!$B$2", "Data!$B$1", "Other!$B$2"])
def test_a_reference_outside_the_table_body_is_left_alone(tmp_path: Path, target: str) -> None:
    """Only references to the table's data body move with it."""
    template = chart_template(tmp_path / "template.xlsx", reference=target)

    assert _value_reference(_fill(template, tmp_path)) == target
