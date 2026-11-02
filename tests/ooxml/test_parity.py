"""The fillxl engine must write what the openpyxl engine writes.

Parity is the cheapest guard against a surgical edit that looks right in the
XML but reads differently in a workbook: both outputs are reloaded with
openpyxl and compared cell by cell, plus everything a reader can observe
around the cells — sheet order and state, the active tab, defined names with
their scope, and table ranges.
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from typing import TYPE_CHECKING, Any

import openpyxl
import pytest
from openpyxl.worksheet.worksheet import Worksheet

from fillxl import FillXLConfig, fill
from tests.ooxml._build import (
    HEADERS,
    body_rows,
    date1904_template,
    named_cell_template,
    table_template,
    table_with_calculated_column_template,
    table_with_sum_below_template,
    table_with_totals_template,
    underscore_sheet_template,
)

if TYPE_CHECKING:
    from pathlib import Path

_MISSING = object()


def snapshot(path: Path) -> dict[str, Any]:
    """Everything a reader of the filled workbook notices, as plain data."""
    workbook = openpyxl.load_workbook(path, data_only=False)
    try:
        sheets: dict[str, Any] = {}
        for name in workbook.sheetnames:
            sheet = workbook[name]
            entry: dict[str, Any] = {"state": sheet.sheet_state}
            if isinstance(sheet, Worksheet):
                entry["cells"] = {
                    cell.coordinate: (cell.value, cell.number_format)
                    for row in sheet.iter_rows()
                    for cell in row
                    if cell.value is not None
                }
                entry["tables"] = {table.name: table.ref for table in sheet.tables.values()}
            sheets[name] = entry
        active = workbook.active
        return {
            "order": list(workbook.sheetnames),
            "active": active.title if active is not None else None,
            "sheets": sheets,
            "names": _defined_names(workbook),
        }
    finally:
        workbook.close()


def _defined_names(workbook: Any) -> dict[Any, str]:  # noqa: ANN401
    """Global and sheet-scoped defined names, keyed by name and scope."""
    names: dict[Any, str] = {}
    for name, defined in workbook.defined_names.items():
        names[(name, defined.localSheetId)] = str(defined.attr_text)
    for sheet_name in workbook.sheetnames:
        for name, defined in workbook[sheet_name].defined_names.items():
            names[(name, defined.localSheetId)] = str(defined.attr_text)
    return names


def fill_both(
    template: Path,
    tmp_path: Path,
    **kwargs: Any,  # noqa: ANN401
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Fill *template* with both engines and return the two snapshots.

    The openpyxl snapshot comes first.  *kwargs* are passed to :func:`fillxl.fill`
    unchanged, so a scenario configures both engines identically.
    """
    outputs = {
        engine: fill(template, tmp_path / engine / template.name, engine=engine, **kwargs)
        for engine in ("openpyxl", "fillxl")
    }
    return snapshot(outputs["openpyxl"]), snapshot(outputs["fillxl"])


def _differences(left: Any, right: Any, path: str = "") -> list[str]:  # noqa: ANN401
    """Every place the two snapshots disagree, as readable strings."""
    if isinstance(left, dict) and isinstance(right, dict):
        found: list[str] = []
        for key in sorted(left.keys() | right.keys(), key=str):
            found.extend(
                _differences(left.get(key, _MISSING), right.get(key, _MISSING), f"{path}/{key}"),
            )
        return found
    if left == right:
        return []
    return [f"{path}: openpyxl={left!r} fillxl={right!r}"]


# ── scenarios ──────────────────────────────────────────────────────────────


def _grows(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_template(tmp_path / "template.xlsx", rows=1),
        {"tables": {"Data": body_rows(5)}},
    )


def _shrinks(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_template(tmp_path / "template.xlsx", rows=4),
        {"tables": {"Data": body_rows(2)}},
    )


def _emptied(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    pandas = pytest.importorskip("pandas")
    empty = pandas.DataFrame(columns=list(HEADERS))
    return (
        table_template(tmp_path / "template.xlsx", rows=2),
        {"tables": {"Data": empty}},
    )


def _totals_grow(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_with_totals_template(tmp_path / "template.xlsx", rows=2),
        {"tables": {"Data": body_rows(5)}},
    )


def _totals_shrink(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_with_totals_template(tmp_path / "template.xlsx", rows=5),
        {"tables": {"Data": body_rows(2)}},
    )


def _headerless(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_template(tmp_path / "template.xlsx", rows=2, headerless=True),
        {"tables": {"Data": body_rows(4)}},
    )


def _calculated_column(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    data = [{key: value for key, value in row.items() if key != "Total"} for row in body_rows(4)]
    return (
        table_with_calculated_column_template(tmp_path / "template.xlsx", rows=2),
        {"tables": {"Data": data}},
    )


def _sum_below_same_sheet(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_with_sum_below_template(tmp_path / "template.xlsx", rows=1),
        {"tables": {"Data": body_rows(4)}},
    )


def _sum_below_other_sheet(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_with_sum_below_template(
            tmp_path / "template.xlsx",
            rows=1,
            other_sheet="Summary",
        ),
        {"tables": {"Data": body_rows(4)}},
    )


def _general_dates(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    data = body_rows(3)
    data[1]["Tax"] = dt.datetime(2024, 2, 2, 13, 30)
    return (
        table_template(tmp_path / "template.xlsx", rows=1),
        {"tables": {"Data": data}},
    )


def _formatted_dates(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        table_template(
            tmp_path / "template.xlsx",
            rows=1,
            number_formats={"Tax": "dd.mm.yyyy"},
        ),
        {"tables": {"Data": body_rows(3)}},
    )


def _text_that_looks_like_an_error(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    data = body_rows(3)
    data[0]["Note"] = "#N/A"
    return (
        table_template(tmp_path / "template.xlsx", rows=1),
        {"tables": {"Data": data}},
    )


def _formula_text_escaped(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    data = body_rows(2)
    data[0]["Note"] = "=1+1"
    return (
        table_template(tmp_path / "template.xlsx", rows=1),
        {"tables": {"Data": data}, "config": FillXLConfig(escape_formulas=True)},
    )


def _formula_text_unescaped(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    data = body_rows(2)
    data[0]["Note"] = "=1+1"
    return (
        table_template(tmp_path / "template.xlsx", rows=1),
        {"tables": {"Data": data}, "config": FillXLConfig(escape_formulas=False)},
    )


def _date_1904(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        date1904_template(tmp_path / "template.xlsx"),
        {
            "tables": {"Data": body_rows(2)},
            "values": {"a_date": dt.date(2025, 5, 5)},
        },
    )


def _named_cell_scalars(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    template = named_cell_template(
        tmp_path / "template.xlsx",
        cells={
            "a_number": ("B2", 1),
            "a_string": ("B3", "text"),
            "a_date": ("B4", dt.date(2024, 1, 1)),
        },
    )
    return (
        template,
        {"values": {"a_number": 5, "a_string": "new", "a_date": dt.date(2025, 5, 5)}},
    )


def _hide_underscore_sheets(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    return (
        underscore_sheet_template(tmp_path / "template.xlsx"),
        {"tables": {"Data": body_rows(3)}},
    )


def _delete_sheets(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    template = named_cell_template(tmp_path / "template.xlsx", other_sheets=("Old",))
    return (
        template,
        {
            "values": {"report_title": 7},
            "config": FillXLConfig(delete_sheets=["Old"]),
        },
    )


def _active_sheet(tmp_path: Path) -> tuple[Path, dict[str, Any]]:
    template = table_template(tmp_path / "template.xlsx", rows=1, sum_sheet="Summary")
    return (
        template,
        {
            "tables": {"Data": body_rows(2)},
            "config": FillXLConfig(active_sheet="Summary"),
        },
    )


SCENARIOS = [
    pytest.param(_grows, id="table-grows"),
    pytest.param(_shrinks, id="table-shrinks"),
    pytest.param(_emptied, id="table-becomes-empty"),
    pytest.param(_totals_grow, id="totals-grows"),
    pytest.param(_totals_shrink, id="totals-shrinks"),
    pytest.param(_headerless, id="headerless"),
    pytest.param(_calculated_column, id="calculated-column"),
    pytest.param(_sum_below_same_sheet, id="sum-below-same-sheet"),
    pytest.param(_sum_below_other_sheet, id="sum-below-other-sheet"),
    pytest.param(_general_dates, id="general-dates"),
    pytest.param(_formatted_dates, id="formatted-dates"),
    pytest.param(_date_1904, id="date-1904"),
    pytest.param(_text_that_looks_like_an_error, id="hash-na-text"),
    pytest.param(_formula_text_escaped, id="formula-text-escaped"),
    pytest.param(_formula_text_unescaped, id="formula-text-unescaped"),
    pytest.param(_named_cell_scalars, id="named-cells"),
    pytest.param(_hide_underscore_sheets, id="hide-underscore-sheets"),
    pytest.param(_delete_sheets, id="delete-sheets"),
    pytest.param(_active_sheet, id="active-sheet"),
]


@pytest.mark.parametrize("build", SCENARIOS)
def test_the_engines_write_the_same_workbook(build: Any, tmp_path: Path) -> None:  # noqa: ANN401
    """A fill that reads differently in the two engines is a fillxl bug."""
    template, kwargs = build(tmp_path)

    expected, actual = fill_both(template, tmp_path, **kwargs)

    differences = _differences(expected, actual)
    assert not differences, "The engines disagree:\n" + "\n".join(differences[:50])


def test_untouched_parts_are_byte_identical_to_the_template(tmp_path: Path) -> None:
    """Filling one named cell must not rewrite anything but that cell's sheet."""
    template = named_cell_template(tmp_path / "template.xlsx")
    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"report_title": 42})

    before = _entries(template)
    after = _entries(output)

    assert set(before) - set(after) <= {"xl/calcChain.xml"}
    for name in sorted(after):
        if name in ("xl/worksheets/sheet1.xml", "xl/workbook.xml"):
            continue
        if name in ("xl/_rels/workbook.xml.rels", "[Content_Types].xml"):
            assert after[name] == _without_calc_chain(before[name], name)
            continue
        assert after[name] == before[name], name


def _entries(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _without_calc_chain(data: bytes, name: str) -> bytes:
    """*data* without the parts that point at a dropped ``calcChain.xml``."""
    element = "Relationship" if name.endswith(".rels") else "Override"
    pattern = rf'<{element}\b[^>]*calcChain\.xml[^>]*/>'
    return re.sub(pattern, "", data.decode("utf-8")).encode("utf-8")
