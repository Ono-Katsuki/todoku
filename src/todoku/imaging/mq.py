"""MQ 算術符号の復号（ITU-T T.88 附属書 E・ISO/IEC 15444-1 附属書 C。二つの仕様は同じ符号）。

文脈の状態は 1 つの整数：2 × 表の番号 + MPS。表 QS（Qe を 16 ビット上げた値）・QE・NM（MPS のときの次）・NL（LPS のときの次、
SWITCH も反映）を引く。速さが要る所（JBIG2 の汎用領域・JPEG 2000 の符号ブロック）は、decode と同じ手順をその場に書き写す。
データの終わりの先は、仕様どおり 0xFF が続くものとして読む（番兵の 2 バイトを足すので、範囲の確かめは要らない）。
"""

from __future__ import annotations

_T = [(0x5601, 1, 1, 1), (0x3401, 2, 6, 0), (0x1801, 3, 9, 0), (0x0AC1, 4, 12, 0), (0x0521, 5, 29, 0), (0x0221, 38, 33, 0),
      (0x5601, 7, 6, 1), (0x5401, 8, 14, 0), (0x4801, 9, 14, 0), (0x3801, 10, 14, 0), (0x3001, 11, 17, 0), (0x2401, 12, 18, 0),
      (0x1C01, 13, 20, 0), (0x1601, 29, 21, 0), (0x5601, 15, 14, 1), (0x5401, 16, 14, 0), (0x5101, 17, 15, 0), (0x4801, 18, 16, 0),
      (0x3801, 19, 17, 0), (0x3401, 20, 18, 0), (0x3001, 21, 19, 0), (0x2801, 22, 19, 0), (0x2401, 23, 20, 0), (0x2201, 24, 21, 0),
      (0x1C01, 25, 22, 0), (0x1801, 26, 23, 0), (0x1601, 27, 24, 0), (0x1401, 28, 25, 0), (0x1201, 29, 26, 0), (0x1101, 30, 27, 0),
      (0x0AC1, 31, 28, 0), (0x09C1, 32, 29, 0), (0x08A1, 33, 30, 0), (0x0521, 34, 31, 0), (0x0441, 35, 32, 0), (0x02A1, 36, 33, 0),
      (0x0221, 37, 34, 0), (0x0141, 38, 35, 0), (0x0111, 39, 36, 0), (0x0085, 40, 37, 0), (0x0049, 41, 38, 0), (0x0025, 42, 39, 0),
      (0x0015, 43, 40, 0), (0x0009, 44, 41, 0), (0x0005, 45, 42, 0), (0x0001, 45, 43, 0), (0x5601, 46, 46, 0)]

QE = [q for q, _m, _l, _s in _T for _ in (0, 1)]
QS = [q << 16 for q in QE]
NM = [2 * m + k for _q, m, _l, _s in _T for k in (0, 1)]
NL = [2 * l + (k ^ s) for _q, _m, l, s in _T for k in (0, 1)]


class MQ:
    """1 つの符号語の復号器。文脈の表（bytearray か list）は呼ぶ側が持つ（同じ表を別の符号語で使い続ける仕様があるため）。"""

    __slots__ = ("data", "bp", "c", "a", "ct")

    def __init__(self, data: bytes | bytearray | memoryview, start: int = 0, end: int | None = None):
        self.data = bytes(data[start:end if end is not None else len(data)]) + b"\xff\xff"
        self.bp = 0
        self.c = self.data[0] << 16
        self.ct = 0
        self._bytein()
        self.c <<= 7
        self.ct -= 7
        self.a = 0x8000

    def _bytein(self) -> None:
        d = self.data
        bp = self.bp
        if d[bp] == 0xFF:
            b1 = d[bp + 1]
            if b1 > 0x8F:
                self.c += 0xFF00
                self.ct = 8
            else:
                self.bp = bp + 1
                self.c += b1 << 9
                self.ct = 7
        else:
            self.bp = bp + 1
            self.c += d[bp + 1] << 8
            self.ct = 8

    def decode(self, cx, i: int) -> int:
        s = cx[i]
        qs = QS[s]
        q = QE[s]
        a = self.a - q
        c = self.c
        if c < qs:
            if a < q:
                d = s & 1
                cx[i] = NM[s]
            else:
                d = (s & 1) ^ 1
                cx[i] = NL[s]
            a = q
        else:
            c -= qs
            if a >= 0x8000:
                self.a = a
                self.c = c
                return s & 1
            if a < q:
                d = (s & 1) ^ 1
                cx[i] = NL[s]
            else:
                d = s & 1
                cx[i] = NM[s]
        ct = self.ct
        while True:
            if ct == 0:
                self.c = c
                self._bytein()
                c = self.c
                ct = self.ct
            a <<= 1
            c <<= 1
            ct -= 1
            if a >= 0x8000:
                break
        self.a = a
        self.c = c
        self.ct = ct
        return d
