"""Data adapter layer for different tabular data types."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from collections.abc import Iterable, Sequence

logger = logging.getLogger(__name__)


class TabularAdapter:
    """Adapter that normalizes different tabular data types to a common interface."""

    def __init__(self, data: Any) -> None:  # noqa: ANN401
        """Initialize adapter with data.

        Args:
            data: Tabular data (DataFrame, list[dict], dict[str, list], etc.)

        """
        self.data = data
        self._columns: list[str] | None = None
        self._rows: Iterable[Sequence[Any]] | None = None
        self._detect_type()

    def _detect_type(self) -> None:
        """Detect the type of data and prepare column names and row iterator."""
        data = self.data

        # Try pandas DataFrame
        try:
            import pandas as pd

            if isinstance(data, pd.DataFrame):
                self._columns = data.columns.tolist()
                self._rows = data.itertuples(index=False, name=None)
                logger.debug("Detected pandas DataFrame with %d columns", len(self._columns))
                return
        except ImportError:
            pass

        # Try polars DataFrame
        try:
            import polars as pl

            if isinstance(data, pl.DataFrame):
                self._columns = data.columns
                self._rows = data.iter_rows()
                logger.debug("Detected polars DataFrame with %d columns", len(self._columns))
                return
        except ImportError:
            pass

        # Try duckdb Relation
        try:
            import duckdb

            if isinstance(data, duckdb.DuckDBPyRelation):
                self._columns = data.columns
                self._rows = iter(data.fetchall())
                logger.debug("Detected duckdb Relation with %d columns", len(self._columns))
                return
        except ImportError:
            pass

        # Try list[dict]
        if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
            self._columns = list(data[0].keys())  # ty:ignore[invalid-assignment]
            self._rows = ([row.get(col) for col in self._columns] for row in data)  # ty:ignore[not-iterable]
            logger.debug("Detected list[dict] with %d columns", len(self._columns))  # ty:ignore[invalid-argument-type]
            return

        # Try dict[str, list]
        if isinstance(data, dict) and len(data) > 0:
            first_value = next(iter(data.values()))
            if isinstance(first_value, list):
                self._columns = list(data.keys())  # ty:ignore[invalid-assignment]
                # Transpose: dict of columns -> list of rows
                num_rows = len(first_value)
                self._rows = ([data[col][i] for col in self._columns] for i in range(num_rows))  # ty:ignore[not-iterable]
                logger.debug("Detected dict[str, list] with %d columns", len(self._columns))  # ty:ignore[invalid-argument-type]
                return

        # If we get here, it's not tabular data
        msg = (
            f"Unsupported data type: {type(data).__name__}. "
            "Expected pandas.DataFrame, polars.DataFrame, duckdb.DuckDBPyRelation, "
            "list[dict], or dict[str, list]. "
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
    def rows(self) -> Iterable[Sequence[Any]]:
        """Get row iterator."""
        if self._rows is None:
            msg = "Rows not initialized"
            raise RuntimeError(msg)
        return self._rows

    @staticmethod
    def is_tabular(data: Any) -> bool:  # noqa: ANN401
        """Check if data appears to be tabular.

        Args:
            data: Data to check

        Returns:
            True if data is likely tabular, False otherwise

        """
        # Quick check without importing heavy dependencies
        if isinstance(data, (list, dict)):
            if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                return True
            if isinstance(data, dict) and len(data) > 0:
                first_value = next(iter(data.values()))
                if isinstance(first_value, list):
                    return True

        # Check for known DataFrame types by class name to avoid imports
        return type(data).__name__ in ("DataFrame", "DuckDBPyRelation")
