"""Unit tests for the dangling-reference detector."""

from __future__ import annotations

import warnings
from typing import TYPE_CHECKING

import openpyxl
import pytest
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.defined_name import DefinedName

from fillxl._dangling import (
    find_dangling_references,
    names_deleted_sheet,
    warn_dangling_references,
)
from tests._assets import PIVOT_TEMPLATE

if TYPE_CHECKING:
    from openpyxl.workbook.workbook import Workbook


def test_a_formula_that_names_the_sheet_is_a_dangling_reference() -> None:
    """The plain case: a formula reads a cell on the sheet that went away."""
    assert names_deleted_sheet("=SUM(RawData!B2:B2)", {"rawdata"})


def test_a_quoted_sheet_name_is_matched_and_unquoted() -> None:
    """Excel quotes names with spaces and doubles quotes inside them."""
    assert names_deleted_sheet("=SUM('Raw Data'!B2)", {"raw data"})
    assert names_deleted_sheet("=SUM('Bob''s Data'!B2)", {"bob's data"})


def test_a_sheet_name_inside_a_string_literal_is_not_a_reference() -> None:
    """Text that merely spells a reference must not trigger a warning."""
    assert not names_deleted_sheet('="RawData!B2 is gone"', {"rawdata"})


def test_a_longer_sheet_name_is_not_mistaken_for_a_shorter_one() -> None:
    """``MyData!`` is not a reference to ``Data!``."""
    assert not names_deleted_sheet("=SUM(MyData!B2)", {"data"})


def _report(workbook: Workbook) -> None:
    """Add every kind of reference to a ``RawData`` sheet to a report sheet."""
    report = workbook.active
    assert report is not None
    report.title = "Report"
    report["A1"] = "=SUM(RawData!B2:B2)"
    raw = workbook.create_sheet("RawData")
    raw.append(["k", "v"])
    raw.append(["a", 1])
    raw.append(["b", 2])
    chart = BarChart()
    chart.add_data(Reference(raw, min_col=2, min_row=1, max_row=3), titles_from_data=True)
    report.add_chart(chart, "D2")
    workbook.defined_names["dangling_name"] = DefinedName("dangling_name", attr_text="RawData!$B$2")


def test_every_kind_of_reference_is_located() -> None:
    """Formulas, charts and defined names all keep the old sheet name."""
    workbook = openpyxl.Workbook()
    try:
        _report(workbook)
        workbook.remove(workbook["RawData"])

        found = find_dangling_references(workbook, ["RawData"])
    finally:
        workbook.close()

    assert found == ["Report!A1", "a chart on 'Report'", "the defined name 'dangling_name'"]


def test_a_name_the_workbook_does_not_use_reports_nothing() -> None:
    """A configured name that matched no reference produces no warning."""
    workbook = openpyxl.Workbook()
    try:
        _report(workbook)
        workbook.remove(workbook["RawData"])

        found = find_dangling_references(workbook, ["Nowhere"])
    finally:
        workbook.close()

    assert found == []


def test_a_pivot_over_a_deleted_sheet_is_located() -> None:
    """A pivot cache whose source sheet is gone cannot refresh."""
    workbook = openpyxl.load_workbook(PIVOT_TEMPLATE)
    try:
        found = find_dangling_references(workbook, ["Sales Data"])
    finally:
        workbook.close()

    assert "a pivot table on 'Pivot Analysis'" in found


def test_a_chart_on_a_chart_sheet_is_located() -> None:
    """A combo or chart-sheet chart keeps its own reference list."""
    workbook = openpyxl.Workbook()
    try:
        report = workbook.active
        assert report is not None
        report.title = "Report"
        raw = workbook.create_sheet("RawData")
        raw.append(["k", "v"])
        raw.append(["a", 1])
        chart = BarChart()
        chart.add_data(Reference(raw, min_col=2, min_row=1, max_row=2), titles_from_data=True)
        workbook.create_chartsheet("Chart").add_chart(chart)
        workbook.remove(raw)

        found = find_dangling_references(workbook, ["RawData"])
    finally:
        workbook.close()

    assert "a chart on 'Chart'" in found


def test_the_warning_names_the_locations_and_the_sheet() -> None:
    """The message says what still points where, so the caller can fix it."""
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        warn_dangling_references(["RawData", "Old"], ["Report!A1"])

    message = str(caught[0].message)
    assert "Deleted sheets 'RawData', 'Old'" in message
    assert "Report!A1" in message
    assert "#REF!" in message


def test_only_the_first_locations_are_listed() -> None:
    """A big workbook must not produce a warning with a thousand locations."""
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        warn_dangling_references(["RawData"], [f"Report!A{row}" for row in range(1, 9)])

    message = str(caught[0].message)
    assert "Report!A5" in message
    assert "Report!A6" not in message
    assert "and 3 more" in message


def test_a_single_location_reads_as_one_sheet() -> None:
    """The message stays grammatical for one sheet and one location."""
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        warn_dangling_references(["RawData"], ["Report!A1"])

    assert "Deleted sheet 'RawData', which is still referenced by: Report!A1" in str(
        caught[0].message,
    )


def test_nothing_is_warned_without_locations() -> None:
    """A clean deletion says nothing."""
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        warn_dangling_references(["RawData"], [])

    assert caught == []


@pytest.mark.parametrize("text", [None, ""])
def test_empty_text_never_names_a_sheet(text: str | None) -> None:
    """Chart references without a formula are not references."""
    assert not names_deleted_sheet(text, {"rawdata"})
