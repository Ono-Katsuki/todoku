"""PDFのオブジェクト（数・文字列・名前・配列・辞書・参照・ストリーム）を読み書きする。外部のライブラリは使わない。"""

from __future__ import annotations

import re
from dataclasses import dataclass


class PdfError(ValueError):
    pass


@dataclass(frozen=True)
class Name:
    name: str

    def __repr__(self) -> str:
        return "/" + self.name


@dataclass(frozen=True)
class Ref:
    num: int
    gen: int = 0


@dataclass
class Stream:
    dict: dict
    raw: bytes  # 圧縮されたまま（暗号化も解いていない）
    num: int = 0
    gen: int = 0


class Real(float):
    """小数で、書き直すと桁が変わるもの（595.303937007874 など）：元の字のまま書き戻すために、元の字を覚える。
    増分更新で書き換えたオブジェクト（ページなど）の、書いていない数が1桁も変わらないように。"""

    __slots__ = ("raw",)

    def __new__(cls, raw: bytes):
        o = super().__new__(cls, float(raw))
        o.raw = bytes(raw)
        return o

    def __reduce__(self):
        return (Real, (self.raw,))


class PdfString(bytes):
    """文字列（バイト列）。hex で書かれていたかを覚える。"""

    hexform: bool = False


WS = b" \t\r\n\x0c\x00"
DELIM = b"()<>[]{}/%"
_NUM = re.compile(rb"[+-]?(?:\d+\.?\d*|\.\d+)")


class Lexer:
    def __init__(self, data: bytes, pos: int = 0):
        self.data = data
        self.pos = pos

    def skip(self) -> None:
        d = self.data
        n = len(d)
        while self.pos < n:
            c = d[self.pos]
            if c in WS:
                self.pos += 1
            elif c == 0x25:  # % コメント
                while self.pos < n and d[self.pos] not in b"\r\n":
                    self.pos += 1
            else:
                break

    def token(self):
        """次の字句。(種類, 値)。種類は num / str / name / kw / '[' ']' '<<' '>>'。"""
        self.skip()
        d = self.data
        if self.pos >= len(d):
            return ("eof", None)
        c = d[self.pos]
        if c == 0x2F:  # /名前
            j = self.pos + 1
            while j < len(d) and d[j] not in WS and d[j] not in DELIM:
                j += 1
            raw = d[self.pos + 1:j]
            self.pos = j
            name = re.sub(rb"#([0-9A-Fa-f]{2})", lambda m: bytes([int(m[1], 16)]), raw)
            return ("name", Name(name.decode("latin-1")))
        if c == 0x28:  # (文字列)
            return ("str", self._literal())
        if c == 0x3C:
            if d[self.pos:self.pos + 2] == b"<<":
                self.pos += 2
                return ("<<", None)
            j = d.index(b">", self.pos)
            hexs = re.sub(rb"\s", b"", d[self.pos + 1:j])
            if len(hexs) % 2:
                hexs += b"0"
            self.pos = j + 1
            s = PdfString(bytes.fromhex(hexs.decode("latin-1")))
            s.hexform = True
            return ("str", s)
        if c == 0x3E and d[self.pos:self.pos + 2] == b">>":
            self.pos += 2
            return (">>", None)
        if c in b"[]{}":
            self.pos += 1
            return (chr(c), None)
        m = _NUM.match(d, self.pos)
        if m and (m.end() >= len(d) or d[m.end()] in WS or d[m.end()] in DELIM):
            self.pos = m.end()
            t = m[0]
            if b"." not in t:
                return ("num", int(t))
            if len(t) - t.index(b".") > 7:
                return ("num", Real(t))  # 7 桁以上の小数：書き戻すと丸まるので、元の字を覚える
            return ("num", float(t))
        j = self.pos
        while j < len(d) and d[j] not in WS and d[j] not in DELIM:
            j += 1
        if j == self.pos:
            self.pos += 1
            return ("kw", d[j:j + 1].decode("latin-1"))
        kw = d[self.pos:j].decode("latin-1")
        self.pos = j
        return ("kw", kw)

    def _literal(self) -> PdfString:
        d = self.data
        i = self.pos + 1
        depth = 1
        out = bytearray()
        while i < len(d):
            c = d[i]
            if c == 0x5C:  # \
                i += 1
                e = d[i]
                esc = {0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12, 0x28: 0x28, 0x29: 0x29, 0x5C: 0x5C}
                if e in esc:
                    out.append(esc[e])
                    i += 1
                elif 0x30 <= e <= 0x37:
                    j = i
                    while j < i + 3 and j < len(d) and 0x30 <= d[j] <= 0x37:
                        j += 1
                    out.append(int(d[i:j], 8) & 0xFF)
                    i = j
                elif e == 0x0D:
                    i += 2 if d[i + 1:i + 2] == b"\n" else 1
                elif e == 0x0A:
                    i += 1
                else:
                    out.append(e)
                    i += 1
                continue
            if c == 0x28:
                depth += 1
            elif c == 0x29:
                depth -= 1
                if depth == 0:
                    self.pos = i + 1
                    return PdfString(bytes(out))
            out.append(c)
            i += 1
        raise PdfError("文字列が閉じていません")


def parse_object(lx: Lexer):
    """オブジェクトを1つ読む（参照 n g R も読む）。"""
    kind, val = lx.token()
    return _obj(lx, kind, val)


def _obj(lx: Lexer, kind, val):
    if kind == "num":
        # 参照（n g R）かもしれない
        save = lx.pos
        k2, v2 = lx.token()
        if k2 == "num" and isinstance(val, int) and isinstance(v2, int):
            k3, v3 = lx.token()
            if k3 == "kw" and v3 == "R":
                return Ref(val, v2)
        lx.pos = save
        return val
    if kind in ("str", "name"):
        return val
    if kind == "[":
        arr = []
        while True:
            k, v = lx.token()
            if k == "]":
                return arr
            if k == "eof":
                raise PdfError("配列が閉じていません")
            arr.append(_obj(lx, k, v))
    if kind == "<<":
        d = {}
        while True:
            k, v = lx.token()
            if k == ">>":
                return d
            if k == "eof":
                raise PdfError("辞書が閉じていません")
            if k != "name":
                continue
            kk, vv = lx.token()
            if kk == ">>":
                d[v.name] = None
                return d
            d[v.name] = _obj(lx, kk, vv)
    if kind == "kw":
        if val == "true":
            return True
        if val == "false":
            return False
        if val == "null":
            return None
        return ("kw", val)
    if kind == "eof":
        raise PdfError("思いがけずファイルの終わりに来ました")
    return None


# ====================================================================== 書く
def _esc_name(n: str) -> str:
    out = []
    for ch in n.encode("latin-1", errors="replace"):
        if ch < 0x21 or ch > 0x7E or ch in b"#()<>[]{}/%":
            out.append(f"#{ch:02X}")
        else:
            out.append(chr(ch))
    return "".join(out)


def dumps(o) -> bytes:
    if o is None:
        return b"null"
    if o is True:
        return b"true"
    if o is False:
        return b"false"
    if isinstance(o, Name):
        return ("/" + _esc_name(o.name)).encode("latin-1")
    if isinstance(o, Ref):
        return f"{o.num} {o.gen} R".encode()
    if isinstance(o, bool):
        return b"true" if o else b"false"
    if isinstance(o, int):
        return str(o).encode()
    if isinstance(o, Real):
        return o.raw
    if isinstance(o, float):
        s = f"{o:.6f}".rstrip("0").rstrip(".")
        return (s or "0").encode()
    if isinstance(o, (bytes, PdfString)):
        if getattr(o, "hexform", False):
            return b"<" + o.hex().upper().encode() + b">"
        return b"(" + o.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)").replace(b"\r", b"\\r") + b")"
    if isinstance(o, str):
        return dumps(text_string(o))
    if isinstance(o, list):
        return b"[" + b" ".join(dumps(x) for x in o) + b"]"
    if isinstance(o, dict):
        return b"<<" + b"".join(dumps(Name(k)) + b" " + dumps(v) for k, v in o.items()) + b">>"
    if isinstance(o, tuple) and o and o[0] == "kw":
        return o[1].encode()
    raise PdfError(f"書けないオブジェクト: {type(o)}")


def text_string(s: str) -> PdfString:
    """PDFの文字列（UTF-16BE の印付き）。ASCII だけなら そのまま。"""
    try:
        raw = s.encode("ascii")
        return PdfString(raw)
    except UnicodeEncodeError:
        out = PdfString(b"\xfe\xff" + s.encode("utf-16-be"))
        out.hexform = True
        return out


PDFDOC = {0x18: "˘", 0x19: "ˇ", 0x1A: "ˆ", 0x1B: "˙", 0x1C: "˝", 0x1D: "˛", 0x1E: "˚", 0x1F: "˜", 0x80: "•", 0x81: "†",
          0x82: "‡", 0x83: "…", 0x84: "—", 0x85: "–", 0x86: "ƒ", 0x87: "⁄", 0x88: "‹", 0x89: "›", 0x8A: "−", 0x8B: "‰",
          0x8C: "„", 0x8D: "“", 0x8E: "”", 0x8F: "‘", 0x90: "’", 0x91: "‚", 0x92: "™", 0x93: "ﬁ", 0x94: "ﬂ", 0x95: "Ł",
          0x96: "Œ", 0x97: "Š", 0x98: "Ÿ", 0x99: "Ž", 0x9A: "ı", 0x9B: "ł", 0x9C: "œ", 0x9D: "š", 0x9E: "ž", 0xA0: "€"}


def decode_text(b) -> str:
    """PDFの「テキスト文字列」（UTF-16BE の印付き、または PDFDocEncoding）を文字列にする。"""
    if b is None:
        return ""
    if isinstance(b, str):
        return b
    if b[:2] == b"\xfe\xff":
        return b[2:].decode("utf-16-be", errors="replace")
    if b[:2] == b"\xff\xfe":
        return b[2:].decode("utf-16-le", errors="replace")
    if b[:3] == b"\xef\xbb\xbf":
        return b[3:].decode("utf-8", errors="replace")
    return "".join(PDFDOC.get(c, chr(c)) for c in b)
