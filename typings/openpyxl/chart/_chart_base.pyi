from typing import Any

class ChartBase:
    title: Any
    legend: Any
    anchor: Any
    width: float
    height: float
    style: int | None
    y_axis: Any
    x_axis: Any
    series: list[Any]
    ser: list[Any]
    _charts: list[Any]
    _axes: dict[Any, Any]
    def __init__(self, **kw: Any) -> None: ...
    def add_data(
        self,
        data: Any,
        from_rows: bool = ...,
        titles_from_data: bool = ...,
    ) -> None: ...
    def set_categories(self, labels: Any) -> None: ...
    def append(self, value: Any) -> None: ...
