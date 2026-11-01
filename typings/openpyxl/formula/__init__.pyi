"""Minimal stubs for ``openpyxl.formula`` (only what fillxl uses)."""

from typing import Literal

class Token:
    value: str
    type: Literal[
        "OPERAND",
        "FUNC",
        "OP-IN",
        "OP-PRE",
        "OP-POST",
        "WSPACE",
        "SEP",
        "ARRAY",
        "PAREN",
    ]
    subtype: str | None
    def __init__(self, value: str, type: str, subtype: str | None = ...) -> None: ...

class Tokenizer:
    formula: str
    items: list[Token]
    def __init__(self, formula: str) -> None: ...
    def render(self) -> str: ...
