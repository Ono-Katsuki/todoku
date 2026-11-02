"""Excelのシリアル値と日付の相互変換（1900年うるう年バグも再現する）。"""

from __future__ import annotations

import datetime as dt

_BASE = dt.date(1899, 12, 30)


def date_to_serial(d: dt.date) -> int:
    n = (d - _BASE).days
    # Excelは1900年2月29日（実在しない）を60として数えるので、それより前は1ずれる
    if n < 61:
        n -= 1
    return n


def serial_to_date(n: float) -> dt.date:
    i = int(n)
    if i == 60:
        raise ValueError("1900-02-29")
    if i < 60:
        i += 1
    return _BASE + dt.timedelta(days=i)


def serial_parts(n: float) -> tuple[int, int, int]:
    """(年, 月, 日)。シリアル60は1900年2月29日として返す。"""
    if int(n) == 60:
        return (1900, 2, 29)
    if int(n) == 0:
        return (1900, 1, 0)
    d = serial_to_date(n)
    return d.year, d.month, d.day


# 和暦（元号の開始日）
ERAS = [
    ("令和", "R", dt.date(2019, 5, 1)),
    ("平成", "H", dt.date(1989, 1, 8)),
    ("昭和", "S", dt.date(1926, 12, 25)),
    ("大正", "T", dt.date(1912, 7, 30)),
    ("明治", "M", dt.date(1868, 1, 1)),
]


def to_wareki(d: dt.date) -> tuple[str, str, int]:
    for name, abbr, start in ERAS:
        if d >= start:
            return name, abbr, d.year - start.year + 1
    return "", "", d.year
