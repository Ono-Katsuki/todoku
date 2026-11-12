"""ISO-2022-JP の Windows の拡張（CP50220・CP50221・CP50222）の読み書き：Python の codec として登録する（名前 iso2022_jp_ms）。

Python の iso2022_jp の仲間は、Windows やメールの道具が書く NEC 特殊文字（13 区：①・㈱・㍻ など）と NEC 選定 IBM 拡張文字
（89〜92 区：髙・﨑 など）、Windows が ESC $ ( D で書く IBM 拡張文字（補助漢字の 83・84 区に置いた 﨑・髙 など）のある
ファイルを読めない。ここでは、規格の読み方を変えずに、その先を足す：規格の ISO-2022-JP として読めるバイトの並びは、
Python の iso2022_jp・iso2022_jp_1・iso2022_jp_ext と同じ字に読む。

読む印（どれも ISO/IEC 2022 の印。ほかの ESC の並びは、厳密に読めないものとして断る）：
- ESC ( B：ASCII。ESC ( J：JIS X 0201 のローマ字（0x5C は ¥、0x7E は ‾）。ESC ( I：半角カナ（0x21〜0x5F。CP50221 の書き方）。
- ESC $ @・ESC $ B：2 バイトの字。JIS X 0208 にある区点は、規格の表（Python の iso2022_jp と同じ字：1 区 33 点は 〜 U+301C）。
  JIS X 0208 にない区点（13 区・89〜92 区）は、CP932（Windows-31J）の表：JIS の区点 → Shift_JIS のバイト → CP932 の字。
- ESC $ ( D：JIS X 0212（補助漢字。Python の iso2022_jp_1 と同じ字）と、Windows が 83・84 区に置く IBM 拡張文字 92 字。
- SO・SI：半角カナへの切り替えと戻し（CP50222 の書き方）。SO の間は、ESC の印で選んだ集合にかかわらず半角カナ。ESC の印は
  SO を終える（Windows は、SI を書かずに ESC ( B・ESC $ B で半角カナを終える。iconv は SI を必ず書くので、どちらが書いた
  ファイルも、書いたとおりの字に読める）。
- 制御文字（改行・タブなど）と空白は、どの状態でもそのままの字（ISO/IEC 2022：C0 と SP は集合の外）。
書くとき：ASCII の字（制御文字・空白も）は ESC ( B の状態で（RFC 1468：行の終わりは ASCII の状態）、¥・‾ は ESC ( J、半角カナは
ESC ( I（Windows の 50220・50221・50222 と iconv の CP50221・CP50222 が、どれも読む書き方）、ほかの字は ESC $ B で、終わりは
ESC ( B に戻す。同じ字が 2 つの区にあるときは、区の番号の小さい方（JIS X 0208 → 13 区 → 89〜92 区）。ESC $ ( D のあるファイル
（x0212=True）では、89〜92 区の字と、2 バイトの表にない字を、ESC $ ( D で書く（Windows の書き方：Windows は 89〜92 区を
外字として読むので、そのファイルを Windows が読める）。

確かめ方：tests/test_iso2022_ms.py（表の字のハッシュ・iconv の CP50221 と同じ区点）と、tests/test_windows_codepages.py
（Windows でだけ動く：Windows そのものの 50220〜50222 の読み書きと、94 × 94 の区点の全部で照らす）。Windows（10.0.26100）
そのものと違う所：89〜92 区（Windows は外字 U+E000〜 として読む。ここは、メールの道具・iconv・ブラウザと同じに、NEC 選定 IBM
拡張文字として読む）、13 区の 64 点（Windows は 〞 U+301E、ここは CP932 と同じ 〝 U+301D）、13 区のうち JIS X 0208 と同じ字の
9 点（≒ ≡ ∫ √ ⊥ ∠ ∵ ∩ ∪：Windows は読まない）、ESC ( J（Windows は 0x5C・0x7E を \・~ と読む）、補助漢字の 2 区 23 点・
81 点（~・№：Windows は読まない）、85〜94 区の外字（ここは読まない）。
"""

from __future__ import annotations

import codecs
import re

NAME = "iso2022_jp_ms"
_ASCII, _ROMAN, _KANA, _WIDE, _X212 = 0, 1, 2, 3, 4
_SO = 8  # 状態の数の 4 ビット目：SO の間
_ESC = {b"\x1b(B": _ASCII, b"\x1b(J": _ROMAN, b"\x1b(I": _KANA, b"\x1b$@": _WIDE, b"\x1b$B": _WIDE, b"\x1b$(D": _X212}
_PARTIAL = (b"\x1b", b"\x1b(", b"\x1b$", b"\x1b$(")
# 印・SO/SI・字になるバイトのひと続き・制御文字と空白のひと続き
_TOK = re.compile(rb"(\x1b(?:\([BJI]|\$[@B]|\$\(D))|([\x0e\x0f])|([\x21-\x7e]+)|([\x00-\x0d\x10-\x1a\x1c-\x20\x7f]+)")
_ROMAN_MAP = {0x5C: "¥", 0x7E: "‾"}
_KANA_MAP = {b: chr(0xFF61 + b - 0x21) if b <= 0x5F else None for b in range(0x21, 0x7F)}
_BAD = "￿"
_tables: tuple[dict, dict] | None = None
_tables212: tuple[dict, dict] | None = None
# Windows（10.0.26100）の 50220〜50222 が、ESC $ ( D の 83 区・84 区に置く IBM 拡張文字（JIS X 0212 にない 92 字）：点の順の
# UTF-16BE の 16 進（ffff は字のない点）。互換漢字（﨑 U+FA11 など）を含むので、字でなく符号位置で置く。
# tests/test_windows_codepages.py が、Windows の上で、Windows そのものの表と照らす
_IBM212 = {
    83: "2170217121722173217421752176217721782179ffffffffffffffffffffffffffffffffffffffffff07ff02ffffffffffff70bb4efc50f451ec5307"
        "5324fa0e548a5759fa0ffa10589e5bec5cf55d53fa115fb760856120654effff6665fa12f9296801fa13fa146a6b6ae26df86df27028fa15fa167501"
        "7682769efa177930fa18fa19fa1afa1b7ae7fa1cfa1d7da07dd6fa1e8362fa1f85b0fa20fa218807fa228b7f8cf48d76fa23fa24fa2590defa269115"
        "fa27fa289592f9dc",
    84: "fa29973bffff9751fa2afa2bfa2c999e9ad99b72fa2d9ed1" + "ffff" * 54 + "974d" + "ffff" * 27,
}


def _sjis(row: int, cell: int) -> bytes:
    """JIS の区点（1〜94）→ Shift_JIS の 2 バイト。"""
    s1 = (row - 1) // 2 + (0x81 if row <= 62 else 0xC1)
    s2 = cell + 0x9E if row % 2 == 0 else cell + (0x3F if cell <= 63 else 0x40)
    return bytes((s1, s2))


def _std(mark: bytes, kk: bytes, codec: str) -> str | None:
    """規格の表の字（Python の codec で、その 2 バイトだけを読む）。字のない区点は None。"""
    try:
        return (mark + kk + b"\x1b(B").decode(codec)
    except UnicodeDecodeError:
        return None


def tables() -> tuple[dict, dict]:
    """ESC $ B の (読む表, 書く表)。読む表：2 バイトを UTF-16BE の 1 字と見た符号位置 → 字（字のない区点は U+FFFF）。
    JIS X 0208 にある区点は規格の表（Python の iso2022_jp）、ない区点は CP932。書く表：字 → 2 バイト（同じ字が 2 つの区に
    あれば、区の番号の小さい方）。初めて使うときに 1 回だけ作る。"""
    global _tables
    if _tables is None:
        dec: dict[int, str] = {}
        enc: dict[str, bytes] = {}
        for row in range(1, 95):
            for cell in range(1, 95):
                kk = bytes((row + 0x20, cell + 0x20))
                ch = _std(b"\x1b$B", kk, "iso2022_jp")
                if ch is None:
                    try:
                        ch = _sjis(row, cell).decode("cp932")
                    except UnicodeDecodeError:
                        dec[(kk[0] << 8) | kk[1]] = _BAD
                        continue
                dec[(kk[0] << 8) | kk[1]] = ch
                enc.setdefault(ch, kk)
        _tables = (dec, enc)
    return _tables


def tables212() -> tuple[dict, dict]:
    """ESC $ ( D の (読む表, 書く表)：JIS X 0212（Python の iso2022_jp_1 と同じ字）と、Windows が 83・84 区に置く IBM 拡張文字。
    書く表は、ESC $ B の表にない字と、89〜92 区の字（Windows は、その区を外字として読む）だけに使う。"""
    global _tables212
    if _tables212 is None:
        dec: dict[int, str] = {}
        enc: dict[str, bytes] = {}
        for row in range(1, 95):
            ibm = bytes.fromhex(_IBM212[row]).decode("utf-16-be") if row in _IBM212 else None
            for cell in range(1, 95):
                kk = bytes((row + 0x20, cell + 0x20))
                ch = _std(b"\x1b$(D", kk, "iso2022_jp_1")
                if ch is None and ibm is not None and ibm[cell - 1] != _BAD:
                    ch = ibm[cell - 1]
                dec[(kk[0] << 8) | kk[1]] = _BAD if ch is None else ch
                if ch is not None and ord(ch) >= 0x80:
                    enc.setdefault(ch, kk)
        _tables212 = (dec, enc)
    return _tables212


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
                st = _ESC[esc]  # ESC の印は SO も終える（Windows は SI を書かない）
            elif shift is not None:
                st = (st | _SO) if shift == b"\x0e" else (st & ~_SO)
            elif ctl is not None:
                out.append(ctl.decode("ascii"))
            else:
                kind = _KANA if st & _SO else st & 7
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
                    table = tables()[0] if kind == _WIDE else tables212()[0]
                    if len(run) % 2:
                        if m.end() == n and not final:
                            self.pending = run[-1:]  # 2 バイトの字の途中
                            run = run[:-1]
                        else:
                            t = run[:-1].decode("utf-16-be").translate(table)
                            if _BAD not in t:
                                out.append(t)
                                pos = self._fail(data, m.end() - 1, m.end(), "2 バイトの字の途中で終わっている", out)
                                continue
                    t = run[:len(run) & ~1].decode("utf-16-be").translate(table)
                    if _BAD in t:
                        bad = t.index(_BAD)
                        out.append(t[:bad])
                        pos = self._fail(data, pos + 2 * bad, pos + 2 * bad + 2, "表に字のない区点", out)
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

    x0212 = False  # 89〜92 区の字と、ESC $ B の表にない字を、ESC $ ( D で書くか（ESC $ ( D のあるファイルの書き方）

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
        marks = {_ASCII: b"\x1b(B", _ROMAN: b"\x1b(J", _KANA: b"\x1b(I", _WIDE: b"\x1b$B", _X212: b"\x1b$(D"}
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
                if self.x0212 and (b is None or b[0] >= 0x79):
                    b2 = tables212()[1].get(ch)
                    if b2 is not None:
                        want, b = _X212, b2
                if b is None:
                    raise UnicodeEncodeError(NAME, input, i, i + 1, "JIS X 0208 と CP932 の表にない字")
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


def encode(s: str, x0212: bool = False) -> bytes:
    """s を、切り替えのない状態から書いて、切り替えのない状態で終わるバイト列。x0212：ESC $ ( D のあるファイルの書き方。"""
    e = IncrementalEncoder("strict")
    e.x0212 = x0212
    return e.encode(s, True)


class _Writer(codecs.StreamWriter):
    encode = staticmethod(_encode)


class _Reader(codecs.StreamReader):
    decode = staticmethod(_decode)


_INFO = codecs.CodecInfo(name=NAME, encode=_encode, decode=_decode, incrementalencoder=IncrementalEncoder,
                         incrementaldecoder=IncrementalDecoder, streamwriter=_Writer, streamreader=_Reader)


def _search(name: str):
    return _INFO if name.replace("-", "_") == NAME else None


codecs.register(_search)
