"""Excelのシリアル値と日付の相互変換（1900年うるう年バグも再現する）。

ブックの日付の数え方は2つある（ECMA-376 第1部 18.2.28 workbookPr の date1904、BIFF8 の DATEMODE レコード）：
- 1900 年の数え方（既定）：1 = 1900/1/1。実在しない 1900/2/29 を 60 として数える
- 1904 年の数え方：0 = 1904/1/1。同じ日付のシリアル値は、1900 年の数え方より 1462 小さい
どちらで数えるかは、引数 date1904 か、計算中のブック（current_1904）で決める。
"""

from __future__ import annotations

import contextvars
import datetime as dt

_BASE = dt.date(1899, 12, 30)
_BASE_1904 = dt.date(1904, 1, 1)
OFFSET_1904 = 1462  # 1904/1/1 の、1900 年の数え方でのシリアル値

# 計算中のブックの日付の数え方（数式の計算・TEXT 関数などで、引数を渡さずに使う）
_CURRENT_1904: contextvars.ContextVar[bool] = contextvars.ContextVar("todoku_date1904", default=False)


def current_1904() -> bool:
    return _CURRENT_1904.get()


def use_1904(flag: bool):
    """with で使う：その間は flag の数え方で日付を数える。"""
    return _Use(bool(flag))


class _Use:
    def __init__(self, flag: bool):
        self.flag = flag
        self.token = None

    def __enter__(self):
        self.token = _CURRENT_1904.set(self.flag)
        return self

    def __exit__(self, *exc):
        _CURRENT_1904.reset(self.token)
        return False


def _flag(date1904: bool | None) -> bool:
    return _CURRENT_1904.get() if date1904 is None else bool(date1904)


def date_to_serial(d: dt.date, date1904: bool | None = None) -> int:
    if _flag(date1904):
        return (d - _BASE_1904).days
    n = (d - _BASE).days
    # Excelは1900年2月29日（実在しない）を60として数えるので、それより前は1ずれる
    if n < 61:
        n -= 1
    return n


def serial_to_date(n: float, date1904: bool | None = None) -> dt.date:
    i = int(n)
    if _flag(date1904):
        return _BASE_1904 + dt.timedelta(days=i)
    if i == 60:
        raise ValueError("1900-02-29")
    if i < 60:
        i += 1
    return _BASE + dt.timedelta(days=i)


def serial_parts(n: float, date1904: bool | None = None) -> tuple[int, int, int]:
    """(年, 月, 日)。1900 年の数え方のシリアル60は1900年2月29日、0 は 1900年1月0日として返す（Excel の表示と同じ）。"""
    if _flag(date1904):
        d = serial_to_date(n, True)
        return d.year, d.month, d.day
    if int(n) == 60:
        return (1900, 2, 29)
    if int(n) == 0:
        return (1900, 1, 0)
    d = serial_to_date(n, False)
    return d.year, d.month, d.day


def weekday_sun0(n: float, date1904: bool | None = None) -> int:
    """曜日（日曜 = 0）。1900 年の数え方では、実在しない 1900/2/29 も1日として通しで数える（Excel と同じ）。"""
    if _flag(date1904):
        return (serial_to_date(n, True).weekday() + 1) % 7
    v = int(n)
    return (serial_to_date(v, False).weekday() + 1) % 7 if v >= 61 else (v - 1) % 7


# 和暦（元号の開始日）。Windows の和暦の表（レジストリの Calendars\Japanese\Eras。明治は「1868 01 01」）と同じ日：
# 明治はその年の初めから数える（改元の詔は 1868 年 10 月 23 日だが、その年の初めにさかのぼる）。
# Excel の表示形式と Word の日付の欄（docx/fill.py）の両方で、この1つの表を使う
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
