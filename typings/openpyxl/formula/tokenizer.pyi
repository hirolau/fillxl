"""Minimal stubs for ``openpyxl.formula.tokenizer`` (only what fillxl uses)."""

from . import Token

class TokenizerError(Exception): ...

class Tokenizer:
    formula: str
    items: list[Token]
    def __init__(self, formula: str) -> None: ...
    def render(self) -> str: ...
