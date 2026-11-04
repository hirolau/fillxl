from datetime import date, datetime, time, timedelta

WINDOWS_EPOCH: datetime
MAC_EPOCH: datetime

def from_excel(value: float, epoch: object = ...) -> datetime: ...
def to_excel(value: date | datetime | time | timedelta, epoch: object = ...) -> float: ...
