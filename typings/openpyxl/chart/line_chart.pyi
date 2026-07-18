from typing import Any

from openpyxl.chart._chart_base import ChartBase

class LineChart(ChartBase):
    tagname: str
    grouping: str | None
    def __init__(self, **kw: Any) -> None: ...

class LineChart3D(LineChart): ...
