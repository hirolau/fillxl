from openpyxl.cell.text import InlineFont

class TextBlock:
    font: InlineFont | None
    text: str
    def __init__(self, font: InlineFont | None = None, text: str = "") -> None: ...

class CellRichText(list[TextBlock | str]):
    def __init__(self, *args: object) -> None: ...
    def __str__(self) -> str: ...
