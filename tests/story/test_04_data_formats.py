"""Chapter 4 — Data Formats for Tables.
=====================================

fillxl accepts tabular data in **five** formats:

1. ``list[dict]`` — a list of rows, each row a dict of column→value.
2. ``dict[str, list]`` — a dict of columns, each value a list of row values.
3. ``pandas.DataFrame`` — if pandas is installed.
4. ``polars.DataFrame`` — if polars is installed.
5. ``duckdb.DuckDBPyRelation`` — if duckdb is installed.

All formats are normalised through the :class:`~fillxl._adapters.TabularAdapter`
before being written to the Excel table.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import openpyxl
import pytest

from fillxl import fill
from fillxl._adapters import TabularAdapter

if TYPE_CHECKING:
    from pathlib import Path

# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_simple_table(path: Path, table_name: str = "Data") -> None:
    """Single-column table for format tests."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    from openpyxl.worksheet.table import Table, TableColumn

    ws["A1"] = "Name"
    ws["B1"] = "Score"
    ws["A2"] = "x"
    ws["B2"] = 0
    table = Table(
        displayName=table_name,
        ref="A1:B2",
        tableColumns=[
            TableColumn(id=1, name="Name"),
            TableColumn(id=2, name="Score"),
        ],
    )
    ws.add_table(table)
    wb.save(path)
    wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 4.1 Python built-in formats
# ═══════════════════════════════════════════════════════════════════════════════


class TestPythonBuiltInFormats:
    """list[dict] and dict[str, list] — no extra dependencies needed."""

    def test_list_of_dicts_is_the_most_natural_python_format(
        self,
        output_path: Path,
    ) -> None:
        """Each dict in the list is one row; keys must match table column names."""
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        data = [
            {"Name": "Alice", "Score": 95},
            {"Name": "Bob", "Score": 87},
            {"Name": "Charlie", "Score": 92},
        ]

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": data},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A2"].value == "Alice"
            assert ws["B3"].value == 87
            assert ws["A4"].value == "Charlie"
        finally:
            wb.close()

    def test_dict_of_lists_is_column_oriented(
        self,
        output_path: Path,
    ) -> None:
        """Each key is a column name, each value is a list of row values."""
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        data: dict[str, list[object]] = {
            "Name": ["Alice", "Bob", "Charlie"],
            "Score": [95, 87, 92],
        }

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": data},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A2"].value == "Alice"
            assert ws["B4"].value == 92
        finally:
            wb.close()

    def test_empty_dict_of_lists_clears_the_table(
        self,
        output_path: Path,
    ) -> None:
        """Passing dict[str, list] with empty lists clears all data rows."""
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": {"Name": [], "Score": []}},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A1"].value == "Name"  # header
            assert ws["A2"].value is None  # cleared
        finally:
            wb.close()

    def test_empty_list_clears_the_table(
        self,
        output_path: Path,
    ) -> None:
        """A plain empty list is valid tabular data with no rows or columns."""
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": []},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A1"].value == "Name"  # header
            assert ws["A2"].value is None  # cleared
            assert wb["Sheet1"].tables["Data"].ref == "A1:B2"
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 4.2 pandas DataFrame
# ═══════════════════════════════════════════════════════════════════════════════


class TestPandasDataFrame:
    """pandas DataFrames are accepted directly; Timestamps and NA sentinels handled."""

    def test_a_pandas_dataframe_fills_the_table_directly(
        self,
        output_path: Path,
    ) -> None:
        """No conversion needed — pass the DataFrame as-is."""
        pd = pytest.importorskip("pandas")
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [95, 87]})

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": df},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A2"].value == "Alice"
            assert ws["B3"].value == 87
        finally:
            wb.close()

    def test_pandas_timestamp_values_are_converted_to_python_datetime(
        self,
        output_path: Path,
    ) -> None:
        """pandas.Timestamp is automatically converted for Excel compatibility."""
        pd = pytest.importorskip("pandas")
        from datetime import datetime

        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        df = pd.DataFrame({"Name": [pd.Timestamp("2025-06-15 10:30:00")], "Score": [100]})

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": df},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            val = ws["A2"].value
            assert isinstance(val, datetime), f"Expected datetime, got {type(val)}"
            assert val == datetime(2025, 6, 15, 10, 30, 0)
        finally:
            wb.close()

    def test_pandas_na_and_nat_sentinels_are_written_as_empty_cells(
        self,
        output_path: Path,
    ) -> None:
        """pd.NA and pd.NaT become None (empty) in Excel."""
        pd = pytest.importorskip("pandas")
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        df = pd.DataFrame({"Name": [pd.NA], "Score": [pd.NaT]})

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": df},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A2"].value is None
            assert ws["B2"].value is None
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 4.3 polars DataFrame
# ═══════════════════════════════════════════════════════════════════════════════


class TestPolarsDataFrame:
    """polars DataFrames are detected and converted automatically."""

    def test_a_polars_dataframe_fills_the_table(
        self,
        output_path: Path,
    ) -> None:
        """Polars DataFrames are detected and converted automatically."""
        pl = pytest.importorskip("polars")
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        df = pl.DataFrame({"Name": ["Alice", "Bob"], "Score": [95, 87]})

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": df},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A2"].value == "Alice"
            assert ws["B3"].value == 87
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 4.4 duckdb Relation
# ═══════════════════════════════════════════════════════════════════════════════


class TestDuckDBRelation:
    """duckdb query results can be piped directly into Excel tables."""

    def test_a_duckdb_relation_fills_the_table(
        self,
        output_path: Path,
    ) -> None:
        """Run a SQL query with duckdb and fill the result into Excel."""
        duckdb = pytest.importorskip("duckdb")
        _make_simple_table(output_path.with_name("tmpl.xlsx"))

        rel = duckdb.sql("SELECT 'Alice' AS Name, 95 AS Score")

        fill(
            template=output_path.with_name("tmpl.xlsx"),
            output=output_path,
            tables={"Data": rel},
        )

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["Sheet1"]
            assert ws["A2"].value == "Alice"
            assert ws["B2"].value == 95
        finally:
            wb.close()


# ═══════════════════════════════════════════════════════════════════════════════
# 4.5 TabularAdapter internals
# ═══════════════════════════════════════════════════════════════════════════════


class TestTabularAdapterInternals:
    """Column/row extraction and TypeError for unsupported types."""

    def test_adapter_extracts_columns_and_rows_from_list_of_dict(self) -> None:
        """Columns come from dict keys; rows from dict values."""
        adapter = TabularAdapter([{"x": 1, "y": 2}, {"x": 3, "y": 4}])
        assert adapter.columns == ["x", "y"]
        assert list(adapter.rows) == [[1, 2], [3, 4]]

    def test_adapter_extracts_columns_and_rows_from_dict_of_list(self) -> None:
        """Columns are the dict keys; rows are transposed."""
        adapter = TabularAdapter({"x": [1, 3], "y": [2, 4]})
        assert adapter.columns == ["x", "y"]
        assert list(adapter.rows) == [[1, 2], [3, 4]]

    def test_adapter_raises_typeerror_for_unsupported_types(self) -> None:
        """A wrong-shaped dict raises TypeError; empty payloads are valid."""
        assert TabularAdapter([]).rows == []
        assert TabularAdapter({}).rows == []
        with pytest.raises(TypeError):
            TabularAdapter({"a": 1})
