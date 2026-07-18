from typing import Any

from openpyxl.chart._chart_base import ChartBase

class SurfaceChart(ChartBase):
    def __init__(self, **kw: Any) -> None: ...

class SurfaceChart3D(SurfaceChart): ...
