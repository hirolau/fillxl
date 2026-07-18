from typing import Any

from openpyxl.chart._chart_base import ChartBase

class ScatterChart(ChartBase):
    tagname: str
    scatterStyle: str | None
    def __init__(self, **kw: Any) -> None: ...
