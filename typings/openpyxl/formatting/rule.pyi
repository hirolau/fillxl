from typing import Any


class CellIsRule:
    operator: str | None
    formula: list[str] | None
    font: Any
    fill: Any
    def __init__(
        self,
        operator: str | None = ...,
        formula: list[str] | None = ...,
        font: Any = ...,
        fill: Any = ...,
    ) -> None: ...