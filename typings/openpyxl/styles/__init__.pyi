from typing import Any

class Font:
    name: str | None
    size: float | None
    bold: bool | None
    italic: bool | None
    color: Any
    def __init__(
        self,
        name: str | None = ...,
        size: float | None = ...,
        bold: bool | None = ...,
        italic: bool | None = ...,
        color: Any = ...,
        # ... other params omitted
    ) -> None: ...


class PatternFill:
    fill_type: str | None
    start_color: Any
    end_color: Any
    def __init__(
        self,
        fill_type: str | None = ...,
        start_color: str | None = ...,
        end_color: str | None = ...,
    ) -> None: ...


class Fill:
    fill_type: str | None
    start_color: Any
    end_color: Any