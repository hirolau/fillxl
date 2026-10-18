"""Type aliases for fillxl."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal, TypeAlias, Union

if TYPE_CHECKING:
    from pathlib import Path

    from duckdb import DuckDBPyRelation
    from pandas import DataFrame
    from polars import DataFrame as PolarsDataFrame

    # Tabular data types that can be used to fill Excel tables.  The
    # column-oriented mapping accepts any sequence-like column values (list,
    # tuple, numpy array, pandas/polars Series, …).
    TabularData: TypeAlias = (
        DataFrame | PolarsDataFrame | DuckDBPyRelation | list[dict[str, Any]] | dict[str, Any]
    )
else:
    TabularData: TypeAlias = Any

# Path-like types
PathLike: TypeAlias = Union[str, "Path"]

# Backend engine selection
Engine: TypeAlias = Literal["auto", "openpyxl", "win32com"]
