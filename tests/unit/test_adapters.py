"""Unit tests for TabularAdapter detection and conversion logic."""

from __future__ import annotations

import datetime as dt
from decimal import Decimal

import pytest

from fillxl import FillXLWarning


class TestTabularAdapter:
    """Unit tests for TabularAdapter detection logic."""

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

    @pytest.mark.polars
    def test_adapter_detects_polars_dataframe(self) -> None:
        """TabularAdapter should detect and extract columns/rows from a polars DataFrame."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        df = pl.DataFrame({"a": [1, 2], "b": [3, 4]})
        adapter = TabularAdapter(df)
        assert adapter.columns == ["a", "b"]
        rows = list(adapter.rows)
        assert rows == [(1, 3), (2, 4)]

    @pytest.mark.duckdb
    def test_adapter_detects_duckdb_relation(self) -> None:
        """TabularAdapter should detect and extract columns/rows from a duckdb relation."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        rel = duckdb.sql("SELECT * FROM (VALUES (1, 2), (3, 4)) AS t(a, b) ORDER BY a")
        adapter = TabularAdapter(rel)
        assert adapter.columns == ["a", "b"]
        assert adapter.rows == [(1, 2), (3, 4)]

    def test_adapter_accepts_empty_list_as_zero_rows(self) -> None:
        """An empty list is valid, schema-less data that clears a table."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter([])
        assert adapter.columns == []
        assert adapter.rows == []

    def test_adapter_accepts_empty_dict_as_zero_rows(self) -> None:
        """An empty dict is valid, schema-less data that clears a table."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({})
        assert adapter.columns == []
        assert adapter.rows == []

    def test_adapter_raises_typeerror_for_dict_with_non_list_values(self) -> None:
        """A dict of scalars points at named cells instead of failing obscurely."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError, match="dict of scalars"):
            TabularAdapter({"a": 1, "b": 2})

    def test_the_constructor_is_the_only_way_to_set_columns_and_rows(self) -> None:
        """Detection always assigns both, so the properties are plain reads."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter([{"a": 1}])

        assert adapter.columns == ["a"]
        assert adapter.rows == [[1]]
        assert adapter._columns is adapter.columns

    def test_adapter_accepts_a_dict_of_numpy_arrays(self) -> None:
        """dict[str, np.ndarray] is column-oriented data like dict[str, list]."""
        numpy = pytest.importorskip("numpy")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({"a": numpy.arange(3), "b": numpy.array([4.5, 5.5, 6.5])})
        assert adapter.columns == ["a", "b"]
        assert len(adapter.rows) == 3
        assert adapter.rows[1][0] == 1

    def test_adapter_accepts_a_dict_of_tuples_and_series(self) -> None:
        """Tuples and pandas Series are column data too."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({"a": (1, 2), "b": pd.Series([3, 4])})
        assert adapter.columns == ["a", "b"]
        assert adapter.rows == [[1, 3], [2, 4]]

    def test_list_of_dicts_uses_the_ordered_union_of_keys(self) -> None:
        """Column order follows first appearance across all rows."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter([{"a": 1}, {"b": 2, "a": 3}])

        assert adapter.columns == ["a", "b"]
        assert adapter.rows == [[1, None], [3, 2]]

    def test_a_key_first_seen_in_a_later_row_is_kept(self) -> None:
        """A key that appears late is still written, with None above it."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(
            [{"Name": "a", "Value": 1}, {"Name": "b", "Value": 2, "Notes": "late"}]
        )

        assert adapter.columns == ["Name", "Value", "Notes"]
        assert adapter.rows[0] == ["a", 1, None]
        assert adapter.rows[1] == ["b", 2, "late"]

    def test_a_non_dict_row_raises_with_its_row_number(self) -> None:
        """A ragged list of dicts fails loudly instead of dropping the row."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError, match="row 2 is int"):
            TabularAdapter([{"a": 1}, 5])

    def test_a_named_pandas_index_warns_that_it_is_dropped(self) -> None:
        """A named index is probably data the user meant to write."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        df = pd.DataFrame({"x": [1, 2]}, index=pd.Index(["a", "b"], name="Date"))

        with pytest.warns(FillXLWarning, match="is not written to Excel") as caught:
            adapter = TabularAdapter(df)

        assert adapter.columns == ["x"]
        assert "Date" in str(caught[0].message)

    def test_a_datetime_index_warns_that_it_is_dropped(self) -> None:
        """A datetime index is data too, not just row numbers."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        df = pd.DataFrame(
            {"x": [1, 2, 3]},
            index=pd.date_range("2024-01-01", periods=3, freq="D"),
        )
        resampled = df.resample("2D").sum()

        with pytest.warns(FillXLWarning, match="is not written to Excel"):
            TabularAdapter(resampled)

    def test_an_unnamed_index_is_not_described_as_one_named_none(self) -> None:
        """The message used to print the index *name*, so it read ``index [None]``."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        df = pd.DataFrame({"x": [1, 2]}, index=pd.Index(["a", "b"]))

        with pytest.warns(FillXLWarning, match="is not written to Excel") as caught:
            TabularAdapter(df)

        message = str(caught[0].message)
        assert message.startswith("The DataFrame index is not written")
        assert "None" not in message

    def test_a_default_index_is_silent(self, recwarn: pytest.WarningsRecorder) -> None:
        """A RangeIndex is row numbering, not data."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(pd.DataFrame({"x": [1, 2]}))

        assert [w for w in recwarn if not issubclass(w.category, DeprecationWarning)] == []

        assert adapter.columns == ["x"]

    def test_a_filtered_integer_index_is_silent(self, recwarn: pytest.WarningsRecorder) -> None:
        """Filtering keeps an integer index; it is still row numbering."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        df = pd.DataFrame({"x": [1, -2, 3]})
        filtered = df[df.x > 0]

        adapter = TabularAdapter(filtered)

        assert adapter.columns == ["x"]
        assert [w for w in recwarn if issubclass(w.category, FillXLWarning)] == []

    def test_a_series_with_a_non_default_index_is_read_positionally(self) -> None:
        """Series values are positional; their index labels must not be used."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({"a": pd.Series([10, 20], index=["x", "y"])})
        assert adapter.rows == [[10], [20]]

    def test_non_string_mapping_keys_are_supported(self) -> None:
        """Column keys may be any type; only the output names are stringified."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({1: [10, 20], 2: [30, 40]})
        assert adapter.columns == ["1", "2"]
        assert adapter.rows == [[10, 30], [20, 40]]

    def test_ragged_columns_raise_instead_of_truncating(self) -> None:
        """Columns of different lengths are an error, not a silent truncation."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(ValueError, match=r"Ragged column data.*'b'.*expected 2"):
            TabularAdapter({"a": [1, 2], "b": [3]})

    def test_a_scalar_beside_columns_is_refused_rather_than_spelled_out(self) -> None:
        """A string beside real columns must not be split into characters."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError) as excinfo:
            TabularAdapter({"amount": [1, 2, 3], "fund": "NAE"})

        assert "'fund'" in str(excinfo.value)

    def test_a_dict_of_scalars_points_the_user_to_named_cells(self) -> None:
        """A dict of scalars is named-cell data, and the error says so."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError) as excinfo:
            TabularAdapter({"a": 1})

        assert "values=" in str(excinfo.value)

    def test_columns_of_any_sequence_type_still_work_side_by_side(self) -> None:
        """A list, a tuple and an array are all column values."""
        numpy = pytest.importorskip("numpy")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({"a": [1, 2], "b": (3, 4), "c": numpy.array([5.5, 6.5])})

        assert adapter.columns == ["a", "b", "c"]
        assert adapter.rows == [[1, 3, 5.5], [2, 4, 6.5]]

    def test_rows_can_be_read_more_than_once(self) -> None:
        """Rows are materialised, so a second access is not empty."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter([{"a": 1}, {"a": 2}])
        assert list(adapter.rows) == [[1], [2]]
        assert list(adapter.rows) == [[1], [2]]


@pytest.mark.polars
class TestPolarsFrames:
    """What a polars frame hands over, and what is refused from it."""

    def test_values_arrive_as_plain_python_objects(self) -> None:
        """polars yields ints and strings, not numpy scalars, so Excel sees them as such."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(pl.DataFrame({"a": [1, 2], "b": [1.5, 2.5], "c": ["x", "y"]}))

        assert adapter.rows == [(1, 1.5, "x"), (2, 2.5, "y")]
        assert [type(value) for value in adapter.rows[0]] == [int, float, str]

    def test_a_null_is_none_not_nan(self) -> None:
        """A polars null arrives as None, which is what an empty cell needs."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(pl.DataFrame({"a": [1, None], "b": ["x", None], "c": [None, 2]}))

        assert adapter.rows == [(1, "x", None), (None, None, 2)]

    def test_nan_stays_a_value(self) -> None:
        """polars keeps NaN and null apart; only the null becomes None here."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(pl.DataFrame({"a": [float("nan"), None]}))

        rows = adapter.rows
        assert rows[0][0] != rows[0][0]  # NaN is never equal to itself
        assert rows[1] == (None,)

    def test_a_column_of_only_nulls_reads_as_none(self) -> None:
        """An all-null column has no dtype to widen to, but still one value per row."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(pl.DataFrame({"a": [None, None]}))

        assert adapter.columns == ["a"]
        assert adapter.rows == [(None,), (None,)]

    def test_a_frame_with_no_rows_keeps_its_columns(self) -> None:
        """A typed empty frame still names its columns, so a table can be cleared."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(pl.DataFrame(schema={"a": pl.Int64, "b": pl.Utf8}))

        assert adapter.columns == ["a", "b"]
        assert adapter.rows == []

    def test_dates_times_and_durations_stay_python_objects(self) -> None:
        """Temporal and decimal columns are converted by polars, not by fillxl."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        frame = pl.DataFrame(
            {
                "day": [dt.date(2025, 1, 2)],
                "clock": [dt.time(9, 30)],
                "stamp": [dt.datetime(2025, 1, 2, 3, 4, 5)],
                "wait": [dt.timedelta(minutes=5)],
                "money": [Decimal("1.23")],
            },
        )

        (row,) = TabularAdapter(frame).rows

        assert row == (
            dt.date(2025, 1, 2),
            dt.time(9, 30),
            dt.datetime(2025, 1, 2, 3, 4, 5),
            dt.timedelta(minutes=5),
            Decimal("1.23"),
        )

    def test_a_lazyframe_is_refused_by_name(self) -> None:
        """A LazyFrame is a query; the error says which type to collect."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        with pytest.raises(TypeError, match=r"Unsupported data type: LazyFrame"):
            TabularAdapter(pl.DataFrame({"a": [1]}).lazy())

    def test_a_series_can_be_a_column(self) -> None:
        """polars Series carry no names, so they slot into dict[str, column] data."""
        pl = pytest.importorskip("polars")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({"a": pl.Series([1, 2]), "b": pl.Series(["x", "y"])})

        assert adapter.columns == ["a", "b"]
        assert adapter.rows == [[1, "x"], [2, "y"]]


@pytest.mark.duckdb
class TestDuckDBRelations:
    """What a duckdb relation hands over, and what is refused from it."""

    def test_rows_are_fetched_from_the_relation(self) -> None:
        """The relation is materialised through ``fetchall``, column by column."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        rel = duckdb.sql("SELECT * FROM (VALUES (1, 'x'), (2, 'y')) AS t(a, b) ORDER BY a")

        assert TabularAdapter(rel).rows == [(1, "x"), (2, "y")]

    def test_a_null_becomes_none(self) -> None:
        """SQL NULL is an empty cell, not the text "NULL"."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        rel = duckdb.sql("SELECT * FROM (VALUES (1, 'x'), (NULL, NULL)) AS t(a, b) ORDER BY a")

        assert TabularAdapter(rel).rows == [(1, "x"), (None, None)]

    def test_an_empty_relation_keeps_its_columns(self) -> None:
        """A query that matches nothing still describes the table to clear."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter(duckdb.sql("SELECT 1 AS a, 'x' AS b WHERE 1 = 0"))

        assert adapter.columns == ["a", "b"]
        assert adapter.rows == []

    def test_sql_types_arrive_as_python_values(self) -> None:
        """duckdb converts DATE, TIME, DECIMAL and INTERVAL before fillxl sees them."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        rel = duckdb.sql(
            "SELECT DATE '2025-01-02' AS day, TIME '09:30:00' AS clock,"
            " TIMESTAMP '2025-01-02 03:04:05' AS stamp,"
            " CAST(1.23 AS DECIMAL(10, 2)) AS money, INTERVAL 5 DAY AS wait",
        )

        (row,) = TabularAdapter(rel).rows

        assert row == (
            dt.date(2025, 1, 2),
            dt.time(9, 30),
            dt.datetime(2025, 1, 2, 3, 4, 5),
            Decimal("1.23"),
            dt.timedelta(days=5),
        )

    def test_duplicate_relation_column_names_are_refused(self) -> None:
        """SQL allows ``SELECT 1 AS a, 2 AS a``; one of them would be lost."""
        duckdb = pytest.importorskip("duckdb")
        from fillxl._adapters import TabularAdapter

        rel = duckdb.sql("SELECT 1 AS a, 2 AS a")

        with pytest.raises(ValueError, match=r"Duplicate column names.*'a'"):
            TabularAdapter(rel)


class TestDuplicateColumnNames:
    """Column labels that collide after string conversion are rejected."""

    def test_duplicate_literal_names_raise(self) -> None:
        """Two columns named "amount" cannot be told apart when writing."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        df = pd.DataFrame([[1, 2]], columns=["amount", "amount"])

        with pytest.raises(ValueError, match=r"Duplicate column names.*'amount'"):
            TabularAdapter(df)

    def test_a_normalization_collision_raises(self) -> None:
        """``1`` and ``"1"`` become the same Excel column name."""
        pd = pytest.importorskip("pandas")
        from fillxl._adapters import TabularAdapter

        df = pd.DataFrame([[1, 2]], columns=[1, "1"])

        with pytest.raises(ValueError, match=r"Duplicate column names.*'1'"):
            TabularAdapter(df)

    def test_a_dict_key_collision_raises(self) -> None:
        """Mappings may hold ``1`` and ``"1"`` as distinct keys."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(ValueError, match=r"Duplicate column names.*'1'"):
            TabularAdapter({1: [10], "1": [20]})

    def test_a_list_of_dicts_key_collision_raises(self) -> None:
        """Keys from different rows are unioned before normalisation."""
        from fillxl._adapters import TabularAdapter

        with pytest.raises(ValueError, match=r"Duplicate column names.*'1'"):
            TabularAdapter([{1: 10}, {"1": 20}])

    def test_unique_names_are_accepted(self) -> None:
        """Distinct labels keep working, including non-string ones."""
        from fillxl._adapters import TabularAdapter

        adapter = TabularAdapter({"amount": [1], "Amount": [2], 3: [4]})

        assert adapter.columns == ["amount", "Amount", "3"]
        assert adapter.rows == [[1, 2, 4]]

    @pytest.mark.polars
    def test_polars_already_rejects_duplicate_names(self) -> None:
        """polars refuses duplicate column names before fillxl sees them.

        Its own error means the adapter can never receive such a frame.
        """
        pl = pytest.importorskip("polars")

        with pytest.raises(pl.exceptions.DuplicateError):
            pl.DataFrame([[1, 2]], schema=["amount", "amount"], orient="row")
