from typing import Any

class NumFmt:
    formatCode: str | None
    sourceLinked: bool | None
    def __init__(
        self,
        formatCode: str | None = ...,
        sourceLinked: bool | None = ...,
    ) -> None: ...

class NumDataSource:
    numRef: Any
    numLit: Any
    def __init__(self, numRef: Any = ..., numLit: Any = ...) -> None: ...

class StrDataSource:
    strRef: Any
    strLit: Any
    def __init__(self, strRef: Any = ..., strLit: Any = ...) -> None: ...

class NumRef:
    f: str | None
    numCache: Any
    def __init__(self, f: str | None = ..., numCache: Any = ..., **kw: Any) -> None: ...

class StrRef:
    f: str | None
    strCache: Any
    def __init__(self, f: str | None = ..., strCache: Any = ..., **kw: Any) -> None: ...

class NumData:
    ptCount: int | None
    pt: Any
    def __init__(self, **kw: Any) -> None: ...

class StrData:
    ptCount: int | None
    pt: Any
    def __init__(self, **kw: Any) -> None: ...

class AxDataSource:
    strRef: Any
    numRef: Any
    multiLvlStrRef: Any
    numLit: Any
    strLit: Any
    def __init__(self, **kw: Any) -> None: ...

class MultiLevelStrRef:
    f: str | None
    multiLvlStrCache: Any
    def __init__(
        self, f: str | None = ..., multiLvlStrCache: Any = ..., extLst: Any = ...
    ) -> None: ...

class MultiLevelStrData:
    ptCount: int | None
    lvl: Any
    def __init__(self, ptCount: int | None = ..., lvl: Any = ..., extLst: Any = ...) -> None: ...
