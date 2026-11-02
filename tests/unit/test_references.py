"""Ranges that track a table body without being a cell formula.

A pivot cache source, a defined name and a data-validation list all hold a
worksheet rectangle.  No engine rewrites them when a table is resized, so a
fill has to say so: a pivot that refreshes on load would otherwise show a
total that quietly leaves the filled rows out.
"""

from __future__ import annotations

import warnings
from typing import TYPE_CHECKING

import pytest

from fillxl import UntrackedRangeWarning
from fillxl._ooxml.package import OoxmlPackage
from fillxl._ooxml.pivots import pivot_worksheet_sources
from fillxl._plan import TableGeometry
from fillxl._references import RangeSource, warn_about_untracked_ranges
from tests.unit.test_ooxml import build_package

if TYPE_CHECKING:
    from pathlib import Path

PIVOT_CACHE = (
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<cacheSource type="worksheet"><worksheetSource ref="A2:B2" sheet="Data"/></cacheSource>'
    "</pivotCacheDefinition>"
)
PIVOT_CACHE_AS_TABLE = (
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource>'
    "</pivotCacheDefinition>"
)

GEOMETRY = TableGeometry(
    top_row=1,
    last_row=2,
    first_col=1,
    last_col=2,
    header_rows=1,
    totals_rows=0,
)


def _warn(sources: list[RangeSource], *, new_body_last_row: int = 4) -> list[str]:
    """The messages :func:`warn_about_untracked_ranges` emits for *sources*."""
    with pytest.warns(UntrackedRangeWarning) as caught:
        warn_about_untracked_ranges(
            sources,
            table_name="Sales",
            sheet_title="Data",
            geometry=GEOMETRY,
            new_body_last_row=new_body_last_row,
        )
    return [str(record.message) for record in caught]


def _assert_quiet(sources: list[RangeSource], *, new_body_last_row: int = 4) -> None:
    """Fail if :func:`warn_about_untracked_ranges` says anything about *sources*."""
    with warnings.catch_warnings():
        warnings.simplefilter("error", UntrackedRangeWarning)
        warn_about_untracked_ranges(
            sources,
            table_name="Sales",
            sheet_title="Data",
            geometry=GEOMETRY,
            new_body_last_row=new_body_last_row,
        )


def test_a_defined_name_over_the_old_body_is_reported() -> None:
    messages = _warn([RangeSource("the defined name 'Body'", "Data", "$A$2:$B$2")])

    assert len(messages) == 1
    assert "the defined name 'Body'" in messages[0]
    assert "Data!$A$2:$B$2" in messages[0]


def test_a_source_over_the_whole_old_table_is_reported() -> None:
    """A pivot source usually includes the header row."""
    assert _warn([RangeSource("a pivot cache source", "Data", "A1:B2")])


def test_a_source_that_already_covers_every_row_is_quiet() -> None:
    _assert_quiet([RangeSource("a pivot cache source", "Data", "A1:B10")])


def test_a_source_on_another_sheet_is_quiet() -> None:
    _assert_quiet([RangeSource("a pivot cache source", "Other", "A1:B2")])


def test_a_table_that_did_not_change_size_says_nothing() -> None:
    _assert_quiet(
        [RangeSource("the defined name 'Body'", "Data", "$A$2:$B$2")],
        new_body_last_row=GEOMETRY.body_last_row,
    )


def test_a_pivot_cache_source_that_names_a_range_is_found(tmp_path: Path) -> None:
    path = build_package(
        tmp_path / "pivot.xlsx",
        extra={"xl/pivotCache/pivotCacheDefinition1.xml": PIVOT_CACHE},
    )

    assert pivot_worksheet_sources(OoxmlPackage(path)) == [("Data", "A2:B2")]


def test_a_pivot_cache_source_that_names_a_table_is_ignored(tmp_path: Path) -> None:
    """Excel keeps a table-named source in step with the table itself."""
    path = build_package(
        tmp_path / "pivot.xlsx",
        extra={"xl/pivotCache/pivotCacheDefinition1.xml": PIVOT_CACHE_AS_TABLE},
    )

    assert pivot_worksheet_sources(OoxmlPackage(path)) == []
