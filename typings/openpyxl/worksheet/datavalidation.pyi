from collections.abc import Iterator
from typing import Any

class DataValidation:
    type: str | None
    formula1: Any
    formula2: Any
    sqref: Any
    allow_blank: bool | None
    showDropDown: bool | None
    showErrorMessage: bool | None
    showInputMessage: bool | None
    prompt: str | None
    promptTitle: str | None
    error: str | None
    errorTitle: str | None
    errorStyle: str | None
    imeMode: str | None
    operator: str | None
    def __init__(
        self,
        type: str | None = ...,
        formula1: Any = ...,
        formula2: Any = ...,
        showErrorMessage: bool | None = ...,
        showInputMessage: bool | None = ...,
        showDropDown: bool | None = ...,
        allow_blank: bool | None = ...,
        promptTitle: str | None = ...,
        errorStyle: str | None = ...,
        error: str | None = ...,
        prompt: str | None = ...,
        errorTitle: str | None = ...,
        imeMode: str | None = ...,
        operator: str | None = ...,
        sqref: Any = ...,
    ) -> None: ...
    def add(self, cell: Any) -> None: ...
    def __iter__(self) -> Iterator[Any]: ...

class DataValidationList:
    dataValidation: Any
    def __iter__(self) -> Iterator[Any]: ...
    def __len__(self) -> int: ...
