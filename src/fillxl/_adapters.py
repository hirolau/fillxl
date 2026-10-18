"""Data adapter layer for different tabular data types."""

from __future__ import annotations

import logging
import sys
from typing import TYPE_CHECKING, Any

# Optional dependencies are looked up in ``sys.modules`` instead of imported:
# an object of one of these types can only exist if the module has already
# been imported, and a per-call ``import`` is both slow and repeatedly fails
# to find packages that are not installed.
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
        self._rows: list[Sequence[Any]] | None = None
        self._detect_type()

    def _detect_type(self) -> None:
        """Detect the type of data and prepare column names and rows."""
        data = self.data

        # Try pandas DataFrame
        pd = sys.modules.get("pandas")
        if pd is not None and isinstance(data, pd.DataFrame):
            self._columns = [str(col) for col in data.columns]
            self._rows = list(data.itertuples(index=False, name=None))
            logger.debug("Detected pandas DataFrame with %d columns", len(self._columns))
            return

        # Try polars DataFrame
        pl = sys.modules.get("polars")
        if pl is not None and isinstance(data, pl.DataFrame):
            self._columns = [str(col) for col in data.columns]
            self._rows = list(data.iter_rows())
            logger.debug("Detected polars DataFrame with %d columns", len(self._columns))
            return

        # Try duckdb Relation
        duckdb = sys.modules.get("duckdb")
        if duckdb is not None and isinstance(data, duckdb.DuckDBPyRelation):
            self._columns = [str(col) for col in data.columns]
            self._rows = [tuple(row) for row in data.fetchall()]
            logger.debug("Detected duckdb Relation with %d columns", len(self._columns))
            return

        # Try list[dict]
        if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
            columns = [str(col) for col in data[0]]
            self._columns = columns
            self._rows = [[row.get(col) for col in columns] for row in data]
            logger.debug("Detected list[dict] with %d columns", len(columns))
            return

        # Try dict[str, list]
        if isinstance(data, dict) and len(data) > 0:
            first_value = next(iter(data.values()))
            if isinstance(first_value, list):
                self._columns = list(data.keys())
                # Transpose: dict of columns -> list of rows
                num_rows = len(first_value)
                self._rows = [[data[col][i] for col in self._columns] for i in range(num_rows)]
                logger.debug("Detected dict[str, list] with %d columns", len(self._columns))
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

    @property
    def iter_rows(self) -> Iterable[Sequence[Any]]:
        """Get a fresh iterator over the rows."""
        return iter(self.rows)

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
