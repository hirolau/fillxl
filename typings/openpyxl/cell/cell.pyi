import re
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from typing import Any

ILLEGAL_CHARACTERS_RE: re.Pattern[str]

_CellValue = (
    bool
    | int
    | float
    | Decimal
    | str
    | bytes
    | datetime
    | date
    | time
    | timedelta
    | None
)

class Cell:
    row: int
    column: int
    coordinate: str
    data_type: str
    number_format: str
    font: Any
    fill: Any
    border: Any
    alignment: Any
    protection: Any
    comment: Any
    hyperlink: Any
    value: _CellValue
    internal_value: Any
    parent: Any
    _style: Any
    def __init__(
        self,
        worksheet: Any,
        row: int | None = ...,
        column: int | None = ...,
        value: Any | None = ...,
        style_array: Any = ...,
    ) -> None: ...
    def offset(self, row: int = ..., column: int = ...) -> Cell: ...
    @property
    def col_idx(self) -> int: ...
    @property
    def column_letter(self) -> str: ...
    @property
    def encoding(self) -> str: ...
    @property
    def base_date(self) -> Any: ...
    @property
    def is_date(self) -> bool: ...

class MergedCell:
    row: int
    column: int
    coordinate: str
    value: _CellValue
    def __init__(self, worksheet: Any, row: int | None = ..., column: int | None = ...) -> None: ...
    def offset(self, row: int = ..., column: int = ...) -> Cell | MergedCell: ...
