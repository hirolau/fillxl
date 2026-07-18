from collections.abc import Iterator
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from typing import Any, Iterable, Literal

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.chart._chart_base import ChartBase
from openpyxl.workbook.child import _WorkbookChild
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.table import Table, TableList

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

class Worksheet(_WorkbookChild):
    sheet_state: Literal["visible", "hidden", "veryHidden"] | str
    sheet_view: Any
    sheet_format: Any
    sheet_properties: Any
    page_setup: Any
    page_margins: Any
    print_options: Any
    protection: Any
    views: Any
    data_validations: Any
    conditional_formatting: Any
    legacy_drawing: Any
    auto_filter: Any
    merged_cells: Any
    row_dimensions: Any
    column_dimensions: Any
    _charts: list[Any]
    _images: list[Any]
    _tables: TableList
    _pivots: list[Any]
    _cells: dict[Any, Any]
    _current_row: int

    def __init__(self, parent: Workbook | Any, title: str | int | None = ...) -> None: ...
    def __delitem__(self, key: str) -> None: ...
    def __getitem__(self, key: str | slice | int) -> Any: ...
    def __iter__(self) -> Iterator[Any]: ...
    def __setitem__(self, key: str, value: _CellValue) -> None: ...
    def add_chart(self, chart: ChartBase | Any, anchor: Any | None = ...) -> None: ...
    def add_table(self, table: Table) -> None: ...
    def add_image(self, img: Any, anchor: Any | None = ...) -> None: ...
    def add_pivot(self, pivot: Any) -> None: ...
    def append(self, iterable: Iterable[Any]) -> None: ...
    def cell(
        self,
        row: int,
        column: int,
        value: Any | None = ...,
    ) -> Cell: ...
    def delete_cols(self, idx: int, amount: int = ...) -> None: ...
    def delete_rows(self, idx: int, amount: int = ...) -> None: ...
    def insert_cols(self, idx: int, amount: int = ...) -> None: ...
    def insert_rows(self, idx: int, amount: int = ...) -> None: ...
    def iter_rows(
        self,
        min_row: int | None = ...,
        max_row: int | None = ...,
        min_col: int | None = ...,
        max_col: int | None = ...,
        values_only: bool = ...,
    ) -> Iterator[Any]: ...
    def iter_cols(
        self,
        min_row: int | None = ...,
        max_row: int | None = ...,
        min_col: int | None = ...,
        max_col: int | None = ...,
        values_only: bool = ...,
    ) -> Iterator[Any]: ...
    def merge_cells(
        self,
        range_string: str | None = ...,
        start_row: int | None = ...,
        start_column: int | None = ...,
        end_row: int | None = ...,
        end_column: int | None = ...,
    ) -> None: ...
    def unmerge_cells(
        self,
        range_string: str | None = ...,
        start_row: int | None = ...,
        start_column: int | None = ...,
        end_row: int | None = ...,
        end_column: int | None = ...,
    ) -> None: ...
    @property
    def dimensions(self) -> str: ...
    @property
    def max_column(self) -> int: ...
    @property
    def max_row(self) -> int: ...
    @property
    def min_column(self) -> int: ...
    @property
    def min_row(self) -> int: ...
    @property
    def tables(self) -> TableList: ...
    @property
    def freeze_panes(self) -> str | None: ...
    @freeze_panes.setter
    def freeze_panes(self, topLeftCell: str | Cell | None = ...) -> None: ...
    @property
    def print_title_cols(self) -> str | None: ...
    @property
    def print_title_rows(self) -> str | None: ...
    @property
    def print_area(self) -> str: ...
    @print_area.setter
    def print_area(self, value: str | Iterable[str] | None) -> None: ...
    def move_range(
        self,
        cell_range: Any,
        rows: int = ...,
        cols: int = ...,
        translate: bool = ...,
    ) -> None: ...
