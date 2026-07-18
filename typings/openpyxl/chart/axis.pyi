from typing import Any

from openpyxl.chart.data_source import NumFmt

class ChartLines:
    def __init__(self, spPr: Any = ...) -> None: ...

class Scaling:
    logBase: Any
    orientation: str | None
    max: Any
    min: Any
    def __init__(
        self,
        logBase: Any = ...,
        orientation: str = ...,
        max: Any = ...,
        min: Any = ...,
        extLst: Any = ...,
    ) -> None: ...

class _BaseAxis:
    axId: Any
    scaling: Scaling | None
    delete: bool | str | None
    axPos: str | None
    majorGridlines: ChartLines | None
    minorGridlines: ChartLines | None
    title: Any
    numFmt: NumFmt | None
    majorTickMark: str | None
    minorTickMark: str | None
    tickLblPos: str | None
    spPr: Any
    txPr: Any
    crossAx: Any
    crosses: str | None
    crossesAt: Any
    def __init__(
        self,
        axId: Any = ...,
        scaling: Scaling | None = ...,
        delete: bool | str | None = ...,
        axPos: str = ...,
        majorGridlines: ChartLines | None = ...,
        minorGridlines: ChartLines | None = ...,
        title: Any = ...,
        numFmt: NumFmt | None = ...,
        majorTickMark: str | None = ...,
        minorTickMark: str | None = ...,
        tickLblPos: str | None = ...,
        spPr: Any = ...,
        txPr: Any = ...,
        crossAx: Any = ...,
        crosses: str | None = ...,
        crossesAt: Any = ...,
    ) -> None: ...

class TextAxis(_BaseAxis):
    auto: bool | str | None
    lblAlgn: str | None
    lblOffset: int | str | None
    tickLblSkip: Any
    tickMarkSkip: Any
    noMultiLvlLbl: bool | str | None
    def __init__(
        self,
        auto: bool | str | None = ...,
        lblAlgn: str | None = ...,
        lblOffset: int | str | None = ...,
        tickLblSkip: Any = ...,
        tickMarkSkip: Any = ...,
        noMultiLvlLbl: bool | str | None = ...,
        extLst: Any = ...,
        **kw: Any,
    ) -> None: ...

class NumericAxis(_BaseAxis):
    crossBetween: str | None
    majorUnit: Any
    minorUnit: Any
    dispUnits: Any
    def __init__(
        self,
        crossBetween: str | None = ...,
        majorUnit: Any = ...,
        minorUnit: Any = ...,
        dispUnits: Any = ...,
        extLst: Any = ...,
        **kw: Any,
    ) -> None: ...
    @classmethod
    def from_tree(cls, node: Any) -> NumericAxis: ...

class DateAxis(_BaseAxis):
    def __init__(self, **kw: Any) -> None: ...

class SeriesAxis(_BaseAxis):
    def __init__(self, **kw: Any) -> None: ...

class DisplayUnitsLabel:
    def __init__(self, **kw: Any) -> None: ...

class DisplayUnitsLabelList:
    def __init__(self, **kw: Any) -> None: ...
