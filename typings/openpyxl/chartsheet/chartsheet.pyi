from typing import Any

from openpyxl.workbook.child import _WorkbookChild

class Chartsheet(_WorkbookChild):
    def __init__(self, parent: Any | None = ..., title: str | None = ..., **kw: Any) -> None: ...
    def add_chart(self, chart: Any) -> None: ...
