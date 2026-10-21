"""Data adapter layer for different tabular data types."""

from __future__ import annotations

import logging
import sys
from collections.abc import Sequence
from typing import TYPE_CHECKING, Any

# Optional dependencies are looked up in ``sys.modules`` instead of imported:
# an object of one of these types can only exist if the module has already
# been imported, and a per-call ``import`` is both slow and repeatedly fails
# to find packages that are not installed.
if TYPE_CHECKING:
    from collections.abc import Iterable

logger = logging.getLogger(__name__)


def _is_column_values(values: Any) -> bool:  # noqa: ANN401
    """Whether *values* is a column of data (list, tuple, array, Series, …).

    Strings and mappings are excluded: ``{"a": "hello"}`` is a single value,
    not a column.

    Examples:
        >>> _is_column_values([1, 2, 3])
        True
        >>> _is_column_values((1, 2))
        True
        >>> _is_column_values("abc")
        False
        >>> _is_column_values({"a": 1})
        False

    """
    if isinstance(values, (str, bytes, bytearray, dict)):
        return False
    if isinstance(values, Sequence):
        return True
    # NumPy arrays and pandas/polars Series are not ``Sequence`` instances.
    return hasattr(values, "tolist") and hasattr(values, "__len__")


def _check_column_lengths(keys: list[Any], column_values: list[list[Any]]) -> None:
    """Raise a column-aware error when the columns have different lengths.

    ``zip(..., strict=True)`` reports ``zip() argument 2 is shorter...``,
    which does not say which column is wrong.

    Examples:
        >>> _check_column_lengths(["a", "b"], [[1, 2], [3, 4]])
        >>> _check_column_lengths(["a", "b"], [[1, 2], [3]])
        Traceback (most recent call last):
            ...
        ValueError: Ragged column data: column 'b' has 1 row(s), expected 2.

    """
    expected = len(column_values[0])
    ragged = [
        f"column {str(key)!r} has {len(values)} row(s), expected {expected}"
        for key, values in zip(keys, column_values, strict=True)
        if len(values) != expected
    ]
    if ragged:
        msg = f"Ragged column data: {'; '.join(ragged)}."
        raise ValueError(msg)


def _normalize_column_names(names: Iterable[Any]) -> list[str]:
    """Stringify column names, rejecting collisions that would lose a column.

    Every backend maps column names to indices, so two labels that become the
    same string would make one column overwrite the other.

    Examples:
        >>> _normalize_column_names(["amount", 1])
        ['amount', '1']
        >>> _normalize_column_names(["amount", "amount"])  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: Duplicate column names after normalization: ['amount']...
        >>> _normalize_column_names([1, "1"])  # doctest: +ELLIPSIS
        Traceback (most recent call last):
            ...
        ValueError: Duplicate column names after normalization: ['1']...

    """
    normalized = [str(name) for name in names]
    seen: set[str] = set()
    duplicates: list[str] = []
    for name in normalized:
        if name in seen and name not in duplicates:
            duplicates.append(name)
        seen.add(name)
    if duplicates:
        msg = (
            f"Duplicate column names after normalization: {duplicates}. "
            "Input column names must be unique after normalization."
        )
        raise ValueError(msg)
    return normalized


def _ordered_key_union(rows: list[Any]) -> list[Any]:
    """Keys of all row dicts in first-seen order (missing keys become ``None``).

    Examples:
        >>> _ordered_key_union([{"a": 1}, {"b": 2, "a": 3}])
        ['a', 'b']
        >>> _ordered_key_union([{"a": 1}, 5])
        Traceback (most recent call last):
            ...
        TypeError: list[dict] row 2 is int, not dict.

    """
    keys: dict[Any, None] = {}
    for number, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            msg = f"list[dict] row {number} is {type(row).__name__}, not dict."
            raise TypeError(msg)
        keys.update(dict.fromkeys(row))
    return list(keys)


def _pandas_index_carries_data(index: Any) -> bool:  # noqa: ANN401
    """Whether a pandas index probably holds data (named, multi-level or non-integer)."""
    if index.nlevels > 1 or any(name is not None for name in index.names):
        return True
    pd = sys.modules["pandas"]
    return len(index) > 0 and not pd.api.types.is_integer_dtype(index.dtype)


class TabularAdapter:
    """Adapter that normalizes different tabular data types to a common interface."""

    def __init__(self, data: Any) -> None:  # noqa: ANN401
        """Initialize adapter with data.

        Args:
            data: Tabular data (DataFrame, list[dict], dict[str, list], etc.)

        """
        self.data = data
        self._columns: list[str] | None = None
        self._rows: list[Sequence[Any]] | None = None
        self._detect_type()

    def _detect_type(self) -> None:
        """Detect the type of data and prepare column names and rows."""
        data = self.data

        # Try pandas DataFrame
        pd = sys.modules.get("pandas")
        if pd is not None and isinstance(data, pd.DataFrame):
            if _pandas_index_carries_data(data.index):
                logger.warning(
                    "The DataFrame index %s is not written to Excel; "
                    "call df.reset_index() to include it.",
                    list(data.index.names),
                )
            self._columns = _normalize_column_names(data.columns)
            self._rows = list(data.itertuples(index=False, name=None))
            logger.debug("Detected pandas DataFrame with %d columns", len(self._columns))
            return

        # Try polars DataFrame
        pl = sys.modules.get("polars")
        if pl is not None and isinstance(data, pl.DataFrame):
            self._columns = _normalize_column_names(data.columns)
            self._rows = list(data.iter_rows())
            logger.debug("Detected polars DataFrame with %d columns", len(self._columns))
            return

        # Try duckdb Relation
        duckdb = sys.modules.get("duckdb")
        if duckdb is not None and isinstance(data, duckdb.DuckDBPyRelation):
            self._columns = _normalize_column_names(data.columns)
            self._rows = [tuple(row) for row in data.fetchall()]
            logger.debug("Detected duckdb Relation with %d columns", len(self._columns))
            return

        # Empty payloads are valid: they mean "no rows" (and, for an empty
        # list, no schema).  The table is cleared instead of raising a
        # misleading TypeError.
        if isinstance(data, list) and not data:
            self._columns = []
            self._rows = []
            logger.debug("Detected empty list (no rows, no columns)")
            return

        # Try list[dict]
        if isinstance(data, list) and isinstance(data[0], dict):
            keys = _ordered_key_union(data)
            self._columns = _normalize_column_names(keys)
            self._rows = [[row.get(key) for key in keys] for row in data]
            logger.debug("Detected list[dict] with %d columns", len(keys))
            return

        # Try dict[str, column] — column values may be lists, tuples, numpy
        # arrays, pandas/polars Series, or anything else sequence-like.
        if isinstance(data, dict) and not data:
            self._columns = []
            self._rows = []
            logger.debug("Detected empty dict (no rows, no columns)")
            return

        if isinstance(data, dict):
            scalars = [str(key) for key, value in data.items() if not _is_column_values(value)]
            if scalars and len(scalars) < len(data):
                msg = (
                    f"dict[str, column] data mixes columns and scalars: {scalars} are not "
                    "columns. Wrap a single value in a list, or pass it through values= if "
                    "it belongs in a named cell."
                )
                raise TypeError(msg)
            if not scalars:
                keys = list(data)
                self._columns = _normalize_column_names(keys)
                # Materialise each column positionally: list(Series) ignores
                # a non-default index, and this also handles arrays/tuples.
                column_values = [list(data[key]) for key in keys]
                _check_column_lengths(keys, column_values)
                self._rows = [list(row) for row in zip(*column_values, strict=True)]
                logger.debug("Detected dict[str, column] with %d columns", len(keys))
                return
            msg = (
                "A dict of scalars is not tabular data. Pass it as values= to fill named "
                "cells, or wrap every value in a list to make a one-row table."
            )
            raise TypeError(msg)

        # If we get here, it's not tabular data
        msg = (
            f"Unsupported data type: {type(data).__name__}. "
            "Expected pandas.DataFrame, polars.DataFrame, duckdb.DuckDBPyRelation, "
            "list[dict], or dict[str, sequence]. "
            "Install the appropriate optional dependency (pandas, polars, or duckdb) "
            "if you're trying to use one of those backends."
        )
        raise TypeError(
            msg,
        )

    @property
    def columns(self) -> list[str]:
        """Get column names."""
        if self._columns is None:
            msg = "Columns not initialized"
            raise RuntimeError(msg)
        return self._columns

    @property
    def rows(self) -> list[Sequence[Any]]:
        """Get the materialised rows.

        Rows are materialised in the constructor so that the property can be
        read repeatedly — a one-shot iterator would silently yield nothing on
        the second access.
        """
        if self._rows is None:
            msg = "Rows not initialized"
            raise RuntimeError(msg)
        return self._rows
