from typing import Any, Literal

from openpyxl.chart._chart_base import ChartBase

class BarChart(ChartBase):
    tagname: str
    barDir: Literal["bar", "col"] | str
    grouping: str | None
    gapWidth: int | str | None
    overlap: int | str | None
    shape: str | None
    type: str | None
    def __init__(
        self,
        gapWidth: int | str = ...,
        overlap: int | str | None = ...,
        serLines: Any = ...,
        extLst: Any = ...,
        **kw: Any,
    ) -> None: ...

class BarChart3D(BarChart): ...
