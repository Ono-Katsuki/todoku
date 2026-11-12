"""ISO-2022-JP の Windows の拡張（CP50220・CP50221・CP50222）の読み書き：Python の codec として登録する（名前 iso2022_jp_ms）。

Python の iso2022_jp の仲間は、2 バイトの字を JIS X 0208 の表で読むので、Windows が書く NEC 特殊文字（13 区：①・㈱・㍻ など）と
NEC 選定 IBM 拡張文字（89〜92 区：髙・﨑 など）のあるファイルを読めない。ここでは、2 バイトの字を、Windows と同じに
CP932（Windows-31J）の表で読む：JIS の区点 → Shift_JIS のバイト → CP932 の字。

読む印（どれも ISO/IEC 2022 の印。ほかの ESC の並びは、厳密に読めないものとして断る）：
- ESC ( B：ASCII。ESC ( J：JIS X 0201 のローマ字（0x5C は ¥、0x7E は ‾）。ESC ( I：半角カナ（0x21〜0x5F。CP50221 の書き方）。
- ESC $ @・ESC $ B：2 バイトの字（どちらも CP932 の表）。
- SO・SI：半角カナへの切り替えと戻し（CP50222 の書き方）。SO の間は、ESC の印で選んだ集合にかかわらず半角カナ。
- 制御文字（改行・タブなど）と空白は、どの状態でもそのままの字（ISO/IEC 2022：C0 と SP は集合の外）。
書くとき：ASCII の字（制御文字・空白も）は ESC ( B の状態で（RFC 1468：行の終わりは ASCII の状態）、¥・‾ は ESC ( J、半角カナは ESC ( I、ほかの字は ESC $ B で、終わりは ESC ( B に戻す。同じ字が
2 つの区にあるとき（∵ は 2 区と 13 区、￢ は 2 区と 92 区）は、区の番号の小さい方（JIS X 0208 → 13 区 → 89〜92 区）。

確かめ方：94 × 94 の区点の全部（字のある 7,336 点・7,326 字）の読みと、その字の書きが、macOS の iconv の CP50221 と同じ（tests/
test_iso2022_ms.py に、表の字のハッシュを置く）。ESC $ @ で 13 区を読むことと、Windows そのものの動きは確かめていない
（iconv は、ESC $ @ を 1978 年の表として読み、13 区を断る）。
"""

from __future__ import annotations

import codecs
import re

NAME = "iso2022_jp_ms"
_ASCII, _ROMAN, _KANA, _WIDE = 0, 1, 2, 3
_SO = 4  # 状態の数の 3 ビット目：SO の間
_ESC = {b"\x1b(B": _ASCII, b"\x1b(J": _ROMAN, b"\x1b(I": _KANA, b"\x1b$@": _WIDE, b"\x1b$B": _WIDE}
_PARTIAL = (b"\x1b", b"\x1b(", b"\x1b$")
# 印・SO/SI・字になるバイトのひと続き・制御文字と空白のひと続き
_TOK = re.compile(rb"(\x1b(?:\([BJI]|\$[@B]))|([\x0e\x0f])|([\x21-\x7e]+)|([\x00-\x0d\x10-\x1a\x1c-\x20\x7f]+)")
_ROMAN_MAP = {0x5C: "¥", 0x7E: "‾"}
_KANA_MAP = {b: chr(0xFF61 + b - 0x21) if b <= 0x5F else None for b in range(0x21, 0x7F)}
_BAD = "￿"
_tables: tuple[dict, dict] | None = None


def _sjis(row: int, cell: int) -> bytes:
    """JIS の区点（1〜94）→ Shift_JIS の 2 バイト。"""
    s1 = (row - 1) // 2 + (0x81 if row <= 62 else 0xC1)
    s2 = cell + 0x9E if row % 2 == 0 else cell + (0x3F if cell <= 63 else 0x40)
    return bytes((s1, s2))


def tables() -> tuple[dict, dict]:
    """(読む表, 書く表)。読む表：2 バイトを UTF-16BE の 1 字と見た符号位置 → 字（字のない区点は U+FFFF）。書く表：字 → 2 バイト
    （同じ字が 2 つの区にあれば、区の番号の小さい方）。初めて使うときに 1 回だけ作る。"""
    global _tables
    if _tables is None:
        dec: dict[int, str] = {}
        enc: dict[str, bytes] = {}
        for row in range(1, 95):
            for cell in range(1, 95):
                code = ((row + 0x20) << 8) | (cell + 0x20)
                try:
                    ch = _sjis(row, cell).decode("cp932")
                except UnicodeDecodeError:
                    dec[code] = _BAD
                    continue
                dec[code] = ch
                enc.setdefault(ch, bytes((row + 0x20, cell + 0x20)))
        _tables = (dec, enc)
    return _tables


class IncrementalDecoder(codecs.IncrementalDecoder):
    def __init__(self, errors: str = "strict"):
        super().__init__(errors)
        self.pending = b""
        self.state = 0

    def reset(self) -> None:
        self.pending = b""
        self.state = 0

    def getstate(self) -> tuple[bytes, int]:
        return self.pending, self.state

    def setstate(self, state) -> None:
        self.pending, self.state = bytes(state[0]), int(state[1])

    def _fail(self, data: bytes, a: int, b: int, why: str, out: list) -> int:
        """厳密に読めない所：errors が strict なら断る。replace・ignore なら、その 1 バイトを飛ばして続ける。"""
        if self.errors == "strict":
            raise UnicodeDecodeError(NAME, data, a, b, why)
        if self.errors == "replace":
            out.append("�")
        return a + 1

    def decode(self, input, final: bool = False) -> str:
        data = self.pending + bytes(input)
        self.pending = b""
        out: list[str] = []
        pos, n, st = 0, len(data), self.state
        while pos < n:
            m = _TOK.match(data, pos)
            if m is None:
                rest = data[pos:]
                if rest in _PARTIAL and not final:
                    self.pending = rest
                    break
                pos = self._fail(data, pos, min(n, pos + 3), "ISO-2022-JP の切り替えの印として読めないバイト", out)
                continue
            esc, shift, run, ctl = m.group(1), m.group(2), m.group(3), m.group(4)
            if esc is not None:
                st = (st & _SO) | _ESC[esc]
            elif shift is not None:
                st = (st | _SO) if shift == b"\x0e" else (st & ~_SO)
            elif ctl is not None:
                out.append(ctl.decode("ascii"))
            else:
                kind = _KANA if st & _SO else st & 3
                if kind == _ASCII:
                    out.append(run.decode("ascii"))
                elif kind == _ROMAN:
                    out.append(run.decode("ascii").translate(_ROMAN_MAP))
                elif kind == _KANA:
                    t = run.decode("ascii").translate(_KANA_MAP)
                    if len(t) != len(run):
                        bad = next(i for i, b in enumerate(run) if b > 0x5F)
                        out.append(run[:bad].decode("ascii").translate(_KANA_MAP))
                        pos = self._fail(data, pos + bad, pos + bad + 1, "半角カナにないバイト", out)
                        continue
                    out.append(t)
                else:
                    if len(run) % 2:
                        if m.end() == n and not final:
                            self.pending = run[-1:]  # 2 バイトの字の途中
                            run = run[:-1]
                        else:
                            t = run[:-1].decode("utf-16-be").translate(tables()[0])
                            if _BAD not in t:
                                out.append(t)
                                pos = self._fail(data, m.end() - 1, m.end(), "2 バイトの字の途中で終わっている", out)
                                continue
                    t = run[:len(run) & ~1].decode("utf-16-be").translate(tables()[0])
                    if _BAD in t:
                        bad = t.index(_BAD)
                        out.append(t[:bad])
                        pos = self._fail(data, pos + 2 * bad, pos + 2 * bad + 2, "CP932 の表に字のない区点", out)
                        if self.errors != "strict":
                            pos += 1
                            self.pending = b""
                        continue
                    out.append(t)
            pos = m.end()
        self.state = st
        if final and self.pending:
            self.pending = b""
        return "".join(out)


class IncrementalEncoder(codecs.IncrementalEncoder):
    """切り替えのない状態から書き、final で切り替えのない状態（ESC ( B）に戻す。"""

    def __init__(self, errors: str = "strict"):
        super().__init__(errors)
        self.mode = _ASCII

    def reset(self) -> None:
        self.mode = _ASCII

    def getstate(self) -> int:
        return self.mode

    def setstate(self, state) -> None:
        self.mode = int(state)

    def encode(self, input: str, final: bool = False) -> bytes:
        out = bytearray()
        mode, table = self.mode, tables()[1]
        marks = {_ASCII: b"\x1b(B", _ROMAN: b"\x1b(J", _KANA: b"\x1b(I", _WIDE: b"\x1b$B"}
        for i, ch in enumerate(input):
            o = ord(ch)
            if o < 0x80:
                want, b = _ASCII, bytes((o,))
                if o in (0x1B, 0x0E, 0x0F):
                    raise UnicodeEncodeError(NAME, input, i, i + 1, "切り替えの印になる制御文字")
            elif ch == "¥" or ch == "‾":
                want, b = _ROMAN, (b"\\" if ch == "¥" else b"~")
            elif 0xFF61 <= o <= 0xFF9F:
                want, b = _KANA, bytes((o - 0xFF61 + 0x21,))
            else:
                want, b = _WIDE, table.get(ch)
                if b is None:
                    raise UnicodeEncodeError(NAME, input, i, i + 1, "CP932 の表にない字")
            if want != mode:
                out += marks[want]
                mode = want
            out += b
        if final and mode != _ASCII:
            out += marks[_ASCII]
            mode = _ASCII
        self.mode = mode
        return bytes(out)


def _decode(input, errors: str = "strict") -> tuple[str, int]:
    return IncrementalDecoder(errors).decode(bytes(input), True), len(input)


def _encode(input: str, errors: str = "strict") -> tuple[bytes, int]:
    return IncrementalEncoder(errors).encode(input, True), len(input)


class _Writer(codecs.StreamWriter):
    encode = staticmethod(_encode)


class _Reader(codecs.StreamReader):
    decode = staticmethod(_decode)


_INFO = codecs.CodecInfo(name=NAME, encode=_encode, decode=_decode, incrementalencoder=IncrementalEncoder,
                         incrementaldecoder=IncrementalDecoder, streamwriter=_Writer, streamreader=_Reader)


def _search(name: str):
    return _INFO if name.replace("-", "_") == NAME else None


codecs.register(_search)
