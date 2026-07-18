from typing import Any, Iterable, Iterator

class TableStyleInfo:
    name: str | None
    showFirstColumn: bool | None
    showLastColumn: bool | None
    showRowStripes: bool | None
    showColumnStripes: bool | None
    def __init__(
        self,
        name: str | None = ...,
        showFirstColumn: bool | None = ...,
        showLastColumn: bool | None = ...,
        showRowStripes: bool | None = ...,
        showColumnStripes: bool | None = ...,
    ) -> None: ...

class TableFormula:
    array: bool | None
    attr_text: str | None
    text: str | None
    def __init__(self, array: bool | None = ..., attr_text: str | None = ...) -> None: ...

class TableColumn:
    id: int | None
    uniqueName: str | None
    name: str
    totalsRowFunction: str | None
    totalsRowLabel: str | None
    calculatedColumnFormula: TableFormula | None
    totalsRowFormula: TableFormula | None
    def __init__(
        self,
        id: int | None = ...,
        uniqueName: str | None = ...,
        name: str | None = ...,
        totalsRowFunction: str | None = ...,
        totalsRowLabel: str | None = ...,
        queryTableFieldId: Any = ...,
        headerRowDxfId: Any = ...,
        dataDxfId: Any = ...,
        totalsRowDxfId: Any = ...,
        headerRowCellStyle: Any = ...,
        dataCellStyle: Any = ...,
        totalsRowCellStyle: Any = ...,
        calculatedColumnFormula: TableFormula | str | None = ...,
        totalsRowFormula: TableFormula | str | None = ...,
        xmlColumnPr: Any = ...,
        extLst: Any = ...,
    ) -> None: ...

class TableList(dict[str, "Table"]):
    def add(self, table: Table) -> None: ...
    def get(self, name: str | None = ..., table_range: str | None = ...) -> Table | None: ...  # type: ignore[override]
    def items(self) -> Any: ...
    def values(self) -> Any: ...
    def __iter__(self) -> Iterator[str]: ...

class Table:
    id: int | None
    displayName: str | None
    ref: str
    name: str | None
    comment: str | None
    headerRowCount: int | None
    totalsRowCount: int | None
    totalsRowShown: bool | None
    autoFilter: Any
    tableColumns: list[TableColumn]
    tableStyleInfo: TableStyleInfo | None
    def __init__(
        self,
        id: int = ...,
        displayName: str | None = ...,
        ref: str | None = ...,
        name: str | None = ...,
        comment: str | None = ...,
        tableType: str | None = ...,
        headerRowCount: int | None = ...,
        insertRow: Any = ...,
        insertRowShift: Any = ...,
        totalsRowCount: int | None = ...,
        totalsRowShown: bool | None = ...,
        published: Any = ...,
        headerRowDxfId: Any = ...,
        dataDxfId: Any = ...,
        totalsRowDxfId: Any = ...,
        headerRowBorderDxfId: Any = ...,
        tableBorderDxfId: Any = ...,
        totalsRowBorderDxfId: Any = ...,
        headerRowCellStyle: Any = ...,
        dataCellStyle: Any = ...,
        totalsRowCellStyle: Any = ...,
        connectionId: Any = ...,
        autoFilter: Any = ...,
        sortState: Any = ...,
        tableColumns: Iterable[TableColumn] | tuple[TableColumn, ...] = ...,
        tableStyleInfo: TableStyleInfo | None = ...,
        extLst: Any = ...,
    ) -> None: ...
    def __contains__(self, key: object) -> bool: ...
