from typing import Any

class NamedStyle:
    name: str
    font: Any
    fill: Any
    border: Any
    alignment: Any
    number_format: str | None
    protection: Any
    def __init__(
        self,
        name: str = ...,
        font: Any = ...,
        fill: Any = ...,
        border: Any = ...,
        alignment: Any = ...,
        number_format: str | None = ...,
        protection: Any = ...,
        builtinId: Any = ...,
        hidden: bool | None = ...,
        xfId: Any = ...,
    ) -> None: ...
