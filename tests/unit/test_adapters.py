"""Unit tests for TabularAdapter detection and conversion logic."""

from __future__ import annotations

import pytest


class TestTabularAdapter:
    """Unit tests for TabularAdapter detection logic."""

    def test_is_tabular_returns_true_for_list_of_dict(self) -> None:
        """is_tabular returns True for list[dict]."""
        from fillxl._adapters import TabularAdapter

        assert TabularAdapter.is_tabular([{"a": 1, "b": 2}])
        assert TabularAdapter.is_tabular([{"x": "hello"}])

    def test_is_tabular_returns_true_for_dict_of_list(self) -> None:
        """is_tabular returns True for dict[str, list]."""
        from fillxl._adapters import TabularAdapter

        assert TabularAdapter.is_tabular({"a": [1, 2], "b": [3, 4]})
        assert TabularAdapter.is_tabular({"x": ["hello"]})

    def test_is_tabular_returns_true_for_pandas_dataframe(self) -> None:
        """is_tabular returns True for a pandas DataFrame (if pandas installed)."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        assert TabularAdapter.is_tabular(pd.DataFrame({"a": [1, 2]}))
        assert TabularAdapter.is_tabular(pd.DataFrame({"x": ["hello"]}))

    def test_is_tabular_returns_false_for_unsupported_types(self) -> None:
        """is_tabular returns False for unsupported types."""
        from fillxl._adapters import TabularAdapter

        assert not TabularAdapter.is_tabular(42)
        assert not TabularAdapter.is_tabular("string")
        assert not TabularAdapter.is_tabular(None)
        assert not TabularAdapter.is_tabular([1, 2, 3])
        assert not TabularAdapter.is_tabular({1, 2, 3})

    def test_adapter_raises_typeerror_for_unsupported_data(self) -> None:
        """TabularAdapter raises TypeError for unsupported data."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError, match="Unsupported data type"):
            TabularAdapter(42)
        with pytest.raises(TypeError, match="Unsupported data type"):
            TabularAdapter("not tabular")
        with pytest.raises(TypeError, match="Unsupported data type"):
            TabularAdapter([1, 2, 3])

    def test_adapter_extracts_columns_and_rows_from_list_of_dict(self) -> None:
        """Adapter extracts correct columns and rows from list[dict]."""
        from fillxl._adapters import TabularAdapter

        data = [
            {"name": "Alice", "age": 30},
            {"name": "Bob", "age": 25},
        ]
        adapter = TabularAdapter(data)
        assert adapter.columns == ["name", "age"]
        rows = list(adapter.rows)
        assert rows == [["Alice", 30], ["Bob", 25]]

    def test_adapter_extracts_columns_and_rows_from_dict_of_list(self) -> None:
        """Adapter extracts correct columns and rows from dict[str, list]."""
        from fillxl._adapters import TabularAdapter

        data = {"name": ["Alice", "Bob"], "age": [30, 25]}
        adapter = TabularAdapter(data)
        assert adapter.columns == ["name", "age"]
        rows = list(adapter.rows)
        assert rows == [["Alice", 30], ["Bob", 25]]

    def test_adapter_detects_polars_dataframe(self) -> None:
        """TabularAdapter should detect and extract columns/rows from a polars DataFrame."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        df = pl.DataFrame({"a": [1, 2], "b": [3, 4]})
        adapter = TabularAdapter(df)
        assert adapter.columns == ["a", "b"]
        rows = list(adapter.rows)
        assert rows == [(1, 3), (2, 4)]

    def test_adapter_detects_duckdb_relation(self) -> None:
        """TabularAdapter should detect and extract columns/rows from a duckdb relation."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        rel = duckdb.sql("SELECT 1 AS a, 2 AS b UNION SELECT 3, 4")
        adapter = TabularAdapter(rel)
        assert adapter.columns == ["a", "b"]

    def test_adapter_raises_typeerror_for_empty_list(self) -> None:
        """An empty list cannot be detected as tabular — TypeError should be raised."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError, match="Unsupported data type"):
            TabularAdapter([])

    def test_adapter_raises_typeerror_for_dict_with_non_list_values(self) -> None:
        """A dict with non-list values should raise TypeError."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError, match="Unsupported data type"):
            TabularAdapter({"a": 1, "b": 2})

    def test_is_tabular_returns_true_for_polars(self) -> None:
        """is_tabular should return True for polars DataFrame."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        assert TabularAdapter.is_tabular(pl.DataFrame({"a": [1]}))

    def test_is_tabular_returns_true_for_duckdb(self) -> None:
        """is_tabular should return True for duckdb relation."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        assert TabularAdapter.is_tabular(duckdb.sql("SELECT 1"))

    def test_adapter_columns_raises_runtime_error_when_not_initialized(self) -> None:
        """Verify .columns on an uninitialized adapter raises RuntimeError."""
        from fillxl._adapters import TabularAdapter

        # Create an adapter that bypasses successful type detection
        # by starting with _columns=None
        adapter = object.__new__(TabularAdapter)
        adapter._columns = None
        adapter._rows = None
        with pytest.raises(RuntimeError, match="Columns not initialized"):
            _ = adapter.columns

    def test_adapter_rows_raises_runtime_error_when_not_initialized(self) -> None:
        """Verify .rows on an uninitialized adapter raises RuntimeError."""
        from fillxl._adapters import TabularAdapter

        adapter = object.__new__(TabularAdapter)
        adapter._columns = None
        adapter._rows = None
        with pytest.raises(RuntimeError, match="Rows not initialized"):
            _ = adapter.rows

    def test_rows_can_be_read_more_than_once(self) -> None:
        """Rows are materialised, so a second access is not empty."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter([{"a": 1}, {"a": 2}])
        assert list(adapter.rows) == [[1], [2]]
        assert list(adapter.rows) == [[1], [2]]
        assert list(adapter.iter_rows) == [[1], [2]]
