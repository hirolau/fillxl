from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

class Reference:
    def __init__(
        self,
        worksheet: Worksheet | None = ...,
        min_col: int | None = ...,
        min_row: int | None = ...,
        max_col: int | None = ...,
        max_row: int | None = ...,
        range_string: str | None = ...,
    ) -> None: ...
    def __len__(self) -> int: ...
    @property
    def rows(self) -> Any: ...
    @property
    def cols(self) -> Any: ...
