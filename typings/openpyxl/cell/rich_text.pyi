from openpyxl.cell.cell import Cell

class InlineFont:
    b: bool | None
    i: bool | None
    def __init__(
        self,
        b: bool | None = ...,
        i: bool | None = ...,
        **kwargs: object,
    ) -> None: ...

class TextBlock:
    font: InlineFont | None
    text: str
    def __init__(self, font: InlineFont | None = None, text: str = "") -> None: ...

class CellRichText(list[TextBlock | str]):
    def __init__(self, *args: object) -> None: ...
    def __str__(self) -> str: ...
