FORMAT_GENERAL: str
FORMAT_TEXT: str
FORMAT_DATE_YYYYMMDD2: str
FORMAT_DATE_DATETIME: str
BUILTIN_FORMATS: dict[int, str]
BUILTIN_FORMATS_MAX_SIZE: int
def is_date_format(fmt: str | None) -> bool: ...
def is_timedelta_format(fmt: str | None) -> bool: ...
