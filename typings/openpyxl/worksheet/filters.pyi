from collections.abc import Iterator, Sequence
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

class Filters:
    blank: bool | None
    calendarType: str | None
    filter: Sequence[str]
    dateGroupItem: Sequence[Any]
    def __init__(
        self,
        blank: bool | None = ...,
        calendarType: str | None = ...,
        filter: Sequence[str] = ...,
        dateGroupItem: Sequence[Any] = ...,
    ) -> None: ...
    def __iter__(self) -> Iterator[Any]: ...

class FilterColumn:
    colId: int | None
    hiddenButton: bool | None
    showButton: bool | None
    filters: Filters | None
    top10: Any
    customFilters: Any
    dynamicFilter: Any
    colorFilter: Any
    iconFilter: Any
    def __init__(
        self,
        colId: int | None = ...,
        hiddenButton: bool | None = ...,
        showButton: bool | None = ...,
        filters: Filters | None = ...,
        top10: Any = ...,
        customFilters: Any = ...,
        dynamicFilter: Any = ...,
        colorFilter: Any = ...,
        iconFilter: Any = ...,
        extLst: Any = ...,
        blank: bool | None = ...,
        vals: Any = ...,
    ) -> None: ...
    def __iter__(self) -> Iterator[Any]: ...
