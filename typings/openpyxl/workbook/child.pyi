from typing import Any, Literal

class _WorkbookChild:
    title: str
    sheet_state: Literal["visible", "hidden", "veryHidden"] | str
    HeaderFooter: Any
    def __init__(
        self,
        parent: Any | None = ...,
        title: str | bytes | int | None = ...,
    ) -> None: ...
    def __repr__(self) -> str: ...
    @property
    def encoding(self) -> str: ...
    @property
    def parent(self) -> Any: ...
    @property
    def path(self) -> str: ...
