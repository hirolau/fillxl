from collections.abc import Iterator
from typing import Any

class AutoFilter:
    ref: str | None
    filterColumn: Any
    sortState: Any
    def __init__(
        self,
        ref: str | None = ...,
        filterColumn: Any = ...,
        sortState: Any = ...,
        extLst: Any = ...,
    ) -> None: ...
    def add_filter_column(
        self,
        col_id: int,
        vals: Any = ...,
        blank: bool = ...,
    ) -> None: ...
    def __iter__(self) -> Iterator[Any]: ...

class SortCondition:
    ref: str
    def __init__(
        self,
        ref: str = ...,
        descending: bool | None = ...,
        sortBy: str | None = ...,
        customList: str | None = ...,
        dxfId: int | None = ...,
        iconSet: Any = ...,
        iconId: Any = ...,
    ) -> None: ...

class SortState:
    ref: str | None
    caseSensitive: bool | None
    columnSort: bool | None
    sortMethod: str | None
    sortCondition: Any
    def __init__(
        self,
        ref: str | None = ...,
        caseSensitive: bool | None = ...,
        columnSort: bool | None = ...,
        sortMethod: str | None = ...,
        sortCondition: Any = ...,
    ) -> None: ...
