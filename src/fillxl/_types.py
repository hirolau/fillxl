"""Type aliases and protocols for fillxl."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal, Protocol, TypeAlias, Union

if TYPE_CHECKING:
    from pathlib import Path

    from duckdb import DuckDBPyRelation
    from pandas import DataFrame
    from polars import DataFrame as PolarsDataFrame

# Tabular data types that can be used to fill Excel tables
# Use string annotations at runtime to avoid import overhead
if TYPE_CHECKING:
    TabularData: TypeAlias = (
        DataFrame | PolarsDataFrame | DuckDBPyRelation | list[dict[str, Any]] | dict[str, list[Any]]
    )
else:
    TabularData: TypeAlias = Any

# Values dict contains scalar values for named cells
Values: TypeAlias = dict[str, Any]

# Tables dict maps table names to tabular data
Tables: TypeAlias = dict[str, Any]

# Path-like types
PathLike: TypeAlias = Union[str, "Path"]

# Backend engine selection
Engine: TypeAlias = Literal["auto", "openpyxl", "win32com"]


class SupportsIterColumns(Protocol):
    """Protocol for tabular data that supports column iteration."""

    def columns(self) -> list[str]: ...
    def iter_rows(self) -> Any: ...  # noqa: ANN401
