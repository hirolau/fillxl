from typing import Any

from openpyxl.chart._chart_base import ChartBase

class AreaChart(ChartBase):
    tagname: str
    grouping: str | None
    def __init__(self, **kw: Any) -> None: ...

class AreaChart3D(AreaChart): ...
