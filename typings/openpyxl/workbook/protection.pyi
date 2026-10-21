from typing import Any

class WorkbookProtection:
    workbookPassword: str | None
    lockStructure: bool | None
    lockWindows: bool | None
    def __init__(
        self,
        workbookPassword: str | None = ...,
        lockStructure: bool | None = ...,
        lockWindows: bool | None = ...,
        **kwargs: Any,
    ) -> None: ...
