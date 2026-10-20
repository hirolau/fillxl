class CalcProperties:
    calcId: int
    calcMode: str | None
    fullCalcOnLoad: bool | None
    refMode: str | None
    iterate: bool | None
    iterateCount: int | None
    iterateDelta: float | None
    fullPrecision: bool | None
    calcCompleted: bool | None
    calcOnSave: bool | None
    concurrentCalc: bool | None
    concurrentManualCount: int | None
    forceFullCalc: bool | None
    def __init__(
        self,
        calcId: int = ...,
        calcMode: str | None = ...,
        fullCalcOnLoad: bool | None = ...,
        refMode: str | None = ...,
        iterate: bool | None = ...,
        iterateCount: int | None = ...,
        iterateDelta: float | None = ...,
        fullPrecision: bool | None = ...,
        calcCompleted: bool | None = ...,
        calcOnSave: bool | None = ...,
        concurrentCalc: bool | None = ...,
        concurrentManualCount: int | None = ...,
        forceFullCalc: bool | None = ...,
    ) -> None: ...
