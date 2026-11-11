"""表の絵（拡張メタファイル EMF、MS-EMF）を読み、字の記録だけを消す・描き足す。外部のライブラリは使わない。

Word は、文書に埋め込んだ Excel の表を、表そのものではなく「表の絵」（Excel が描いた EMF）で表示・印刷する。
埋め込みの表のセルに書いたとき、この絵を作り直さないと、Word には元の値のまま出る。
ここでは、絵を描き直すのではなく、元の絵のレコードを1バイトも変えずに残し、

- 変わったセルの字のレコード（EMR_EXTTEXTOUTW）を取り除き
- 新しい字のレコードを、GDI の描く部分の終わり（EMF+ の終わりのコメントの前）に足す

だけにする。足すレコードは、状態を保存して（SAVEDC）、セルの範囲で切り抜き（INTERSECTCLIPRECT）、字体を作って選び、
字の色・背景を透明に・寄せ方（TA_LEFT／TA_RIGHT／TA_CENTER）を決めて字を描き、状態を戻し（RESTOREDC）、
作った字体を消す（DELETEOBJECT）一続きの塊。寄せ方を GDI に任せるので、右寄せ・中央寄せの位置は字の幅が分からなくても正しい。

読むときは、再生の状態（SAVEDC・RESTOREDC の段、選んでいる字体、字の色、寄せ方、切り抜き、座標変換）を
順にたどって、字のレコードごとに「どの字体で、どこを基準に、どの寄せ方で描いたか」と、罫線（細い矩形の塗りと線）を取り出す。
これを、埋め込みのブックの列の幅・行の高さと照らし合わせるのは todoku.sheetpic。

EMF+ のコメントがある絵（EMF+ Dual）は、EMF+ に対応した再生でも GDI のレコードが描かれる区間（EmfPlusGetDC の後）に
足すときだけ受け付ける。そうでない絵は、足した字が EMF+ の再生で消えるので、書き直さない。

ほかに、OfficeArt の絵の識別子（MD4、RFC 1320）もここで計算する（Python の hashlib は MD4 を持たないことがあるため）。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

EMR_HEADER, EMR_EOF, EMR_COMMENT = 0x01, 0x0E, 0x46
EMR_SETWINDOWEXTEX, EMR_SETWINDOWORGEX, EMR_SETVIEWPORTEXTEX, EMR_SETVIEWPORTORGEX = 0x09, 0x0A, 0x0B, 0x0C
EMR_SETMAPMODE, EMR_SETBKMODE, EMR_SETTEXTALIGN, EMR_SETTEXTCOLOR = 0x11, 0x12, 0x16, 0x18
EMR_OFFSETCLIPRGN, EMR_MOVETOEX, EMR_SETMETARGN, EMR_EXCLUDECLIPRECT, EMR_INTERSECTCLIPRECT = 0x1A, 0x1B, 0x1C, 0x1D, 0x1E
EMR_SCALEVIEWPORTEXTEX, EMR_SCALEWINDOWEXTEX, EMR_SAVEDC, EMR_RESTOREDC = 0x1F, 0x20, 0x21, 0x22
EMR_SETWORLDTRANSFORM, EMR_MODIFYWORLDTRANSFORM, EMR_SELECTOBJECT, EMR_DELETEOBJECT = 0x23, 0x24, 0x25, 0x28
EMR_LINETO, EMR_SELECTCLIPPATH, EMR_EXTSELECTCLIPRGN, EMR_BITBLT = 0x36, 0x43, 0x4B, 0x4C
EMR_EXTCREATEFONTINDIRECTW, EMR_EXTTEXTOUTA, EMR_EXTTEXTOUTW = 0x52, 0x53, 0x54
EMR_POLYLINE16, EMR_POLYTEXTOUTA, EMR_POLYTEXTOUTW, EMR_SMALLTEXTOUT = 0x57, 0x60, 0x61, 0x6C
# 新しい GDI のオブジェクトを作るレコード（オブジェクトの表の番号を使う）
CREATES = {0x26, 0x27, 0x31, 0x52, 0x5D, 0x5E, 0x5F, 0x63, 0x7A}
TRANSFORMS = {EMR_SETWINDOWEXTEX, EMR_SETWINDOWORGEX, EMR_SETVIEWPORTEXTEX, EMR_SETVIEWPORTORGEX, EMR_SETMAPMODE,
              EMR_SCALEVIEWPORTEXTEX, EMR_SCALEWINDOWEXTEX, EMR_SETWORLDTRANSFORM, EMR_MODIFYWORLDTRANSFORM}
TA_LEFT, TA_RIGHT, TA_CENTER, TA_UPDATECP = 0, 2, 6, 1
EMFPLUS_GETDC = 0x4004
EMF_SIGNATURE = 0x464D4520  # " EMF"


class MetafileError(ValueError):
    pass


@dataclass
class LogFont:
    """字体（LogFont の 92 バイト）と、それを含むレコードの中身（作り直すときに写す）。"""

    height: int
    weight: int
    italic: bool
    underline: bool
    strike: bool
    charset: int
    face: str
    elw: bytes  # レコードの ihFonts の後ろ全部（LogFontExDv など）

    @property
    def key(self) -> tuple:
        """字の幅が同じになる字体の印（大きさは含めない）。"""
        return (self.face.lower(), self.weight >= 600, self.italic)


@dataclass
class TextRun:
    """字のレコード1つ：何を、どの字体で、どこを基準に、どの寄せ方で描いたか。"""

    index: int  # レコードの番号
    text: str
    x: int
    y: int
    align: int  # そのときの SETTEXTALIGN
    font: LogFont | None
    color: int  # COLORREF（0x00BBGGRR）
    dx: list[int]
    bounds: tuple[int, int, int, int]  # GDI が書いた字の外接矩形（含む・含む）。再生では使わない
    clip: object  # None（切り抜きなし）・(左, 上, 右, 下)・"complex"（矩形でない切り抜き）
    xform: tuple  # 座標変換の状態（同じかどうかを比べる）
    gmode: int
    scales: tuple[float, float]
    kind: int = EMR_EXTTEXTOUTW

    @property
    def width(self) -> int | None:
        """字の幅（送り幅の合計）。送り幅の配列がなければ分からない。"""
        return sum(self.dx) if self.dx and len(self.dx) == len(self.text) else None

    @property
    def box_height(self) -> int | None:
        """字の高さ（GDI が書いた外接矩形から）。上寄せの基準（TA_TOP）で描いたものだけ。"""
        l_, t, r, b = self.bounds
        if r < l_ or b < t or self.align & 0x18 or t != self.y:
            return None
        return b - t + 1


@dataclass
class Line:
    """罫線の候補：縦（v）か横（h）の線の中心の位置と、範囲。"""

    orient: str
    at: float
    lo: int
    hi: int
    thick: int


@dataclass
class Emf:
    data: bytes
    records: list[tuple[int, int, int]] = field(default_factory=list)  # (位置, 種類, 大きさ)

    def __post_init__(self):
        d = self.data
        if len(d) < 88 or struct.unpack_from("<I", d, 0)[0] != EMR_HEADER:
            raise MetafileError("拡張メタファイル（EMF）ではありません")
        if struct.unpack_from("<I", d, 40)[0] != EMF_SIGNATURE:
            raise MetafileError("拡張メタファイル（EMF）の印がありません")
        self.bounds = struct.unpack_from("<4i", d, 8)
        self.n_bytes, self.n_records, self.n_handles = struct.unpack_from("<IIH", d, 48)
        pos = 0
        while pos + 8 <= len(d):
            t, size = struct.unpack_from("<II", d, pos)
            if size < 8 or size % 4 or pos + size > len(d):
                raise MetafileError(f"EMF のレコードの大きさが壊れています（位置 {pos}）")
            self.records.append((pos, t, size))
            pos += size
            if t == EMR_EOF:
                break
        if not self.records or self.records[-1][1] != EMR_EOF:
            raise MetafileError("EMF の終わり（EMR_EOF）がありません")
        if self.n_records != len(self.records):
            raise MetafileError("EMF の見出しのレコード数が、実際の数と合いません")
        self._play()

    def rec(self, i: int) -> bytes:
        p, _t, s = self.records[i]
        return self.data[p:p + s]

    # ------------------------------------------------------------------ 再生の状態をたどる
    def _play(self) -> None:
        st = {"align": 0, "color": 0, "bk": 2, "font": None, "clip": None, "xform": (), "cur": (0, 0)}
        stack: list[dict] = []
        objects: dict[int, object] = {}
        self.runs: list[TextRun] = []
        self.lines: list[Line] = []
        self.fonts: list[LogFont] = []
        self.unknown_text = 0  # 字を描くが、ここでは読まないレコード（POLYTEXTOUT・SMALLTEXTOUT）
        self.plus_ops: list[tuple[int, int]] = []  # (レコードの番号, EMF+ のレコードの種類)
        self.dual = False
        self.state_at: list[dict] = []
        for i, (pos, t, size) in enumerate(self.records):
            r = self.data[pos:pos + size]
            self.state_at.append({k: v for k, v in st.items()})
            if t == EMR_SAVEDC:
                stack.append(dict(st))
            elif t == EMR_RESTOREDC:
                n = struct.unpack_from("<i", r, 8)[0]
                if n >= 0 or -n > len(stack):
                    raise MetafileError("EMF の状態の戻し（RESTOREDC）が、保存（SAVEDC）と合いません")
                for _ in range(-n):
                    st = stack.pop()
            elif t == EMR_SETTEXTALIGN:
                st["align"] = struct.unpack_from("<I", r, 8)[0]
            elif t == EMR_SETTEXTCOLOR:
                st["color"] = struct.unpack_from("<I", r, 8)[0]
            elif t == EMR_SETBKMODE:
                st["bk"] = struct.unpack_from("<I", r, 8)[0]
            elif t in CREATES:
                ih = struct.unpack_from("<I", r, 8)[0]
                if t == EMR_EXTCREATEFONTINDIRECTW:
                    f = _logfont(r)
                    objects[ih] = f
                    self.fonts.append(f)
                else:
                    objects[ih] = "other"
            elif t == EMR_SELECTOBJECT:
                ih = struct.unpack_from("<I", r, 8)[0]
                if ih & 0x80000000:
                    if 0x8000000A <= ih <= 0x80000011:
                        st["font"] = None  # 標準の字体
                else:
                    o = objects.get(ih)
                    if isinstance(o, LogFont):
                        st["font"] = o
            elif t == EMR_DELETEOBJECT:
                objects.pop(struct.unpack_from("<I", r, 8)[0], None)
            elif t == EMR_INTERSECTCLIPRECT:
                rc = struct.unpack_from("<4i", r, 8)
                c = st["clip"]
                if c is None:
                    st["clip"] = rc
                elif c != "complex":
                    st["clip"] = (max(c[0], rc[0]), max(c[1], rc[1]), min(c[2], rc[2]), min(c[3], rc[3]))
            elif t == EMR_EXTSELECTCLIPRGN:
                cb, mode = struct.unpack_from("<II", r, 8)
                st["clip"] = None if (cb == 0 and mode == 5) else "complex"
            elif t in (EMR_EXCLUDECLIPRECT, EMR_OFFSETCLIPRGN, EMR_SELECTCLIPPATH, EMR_SETMETARGN):
                st["clip"] = "complex"
            elif t in TRANSFORMS:
                if st["clip"] is not None:
                    st["clip"] = "complex"  # 切り抜きは、決めたときの座標で残る
                st["xform"] = st["xform"] + ((t, r[8:]),)
            elif t == EMR_MOVETOEX:
                st["cur"] = struct.unpack_from("<2i", r, 8)
            elif t == EMR_LINETO:
                x2, y2 = struct.unpack_from("<2i", r, 8)
                self._seg(st, st["cur"], (x2, y2))
                st["cur"] = (x2, y2)
            elif t == EMR_POLYLINE16:
                n = struct.unpack_from("<I", r, 24)[0]
                pts = [struct.unpack_from("<2h", r, 28 + 4 * k) for k in range(n)]
                for a, b in zip(pts, pts[1:]):
                    self._seg(st, a, b)
            elif t == EMR_BITBLT:
                x, y, cx, cy = struct.unpack_from("<4i", r, 24)
                cb_src = struct.unpack_from("<I", r, 88)[0] if size >= 100 else 1  # cbBmiSrc
                if cb_src == 0 and cx > 0 and cy > 0:  # 元の絵のない塗り（PatBlt）：細ければ罫線
                    if cx <= cy:
                        self.lines.append(Line("v", x + cx // 2, y, y + cy, cx))
                    else:
                        self.lines.append(Line("h", y + cy // 2, x, x + cx, cy))
            elif t in (EMR_EXTTEXTOUTW, EMR_EXTTEXTOUTA):
                self.runs.append(_text_run(i, t, r, st))
            elif t in (EMR_POLYTEXTOUTA, EMR_POLYTEXTOUTW, EMR_SMALLTEXTOUT):
                self.unknown_text += 1
            elif t == EMR_COMMENT and size >= 16 and r[12:16] == b"EMF+":
                ds = struct.unpack_from("<I", r, 8)[0]
                q, end = 16, min(12 + ds, size)
                while q + 12 <= end:
                    ptype, pflags, psize = struct.unpack_from("<HHI", r, q)
                    self.plus_ops.append((i, ptype))
                    if ptype == 0x4001:
                        self.dual = bool(pflags & 1)
                    if psize < 12:
                        break
                    q += psize
        self.final_state = st
        self.stack_left = len(stack)

    def _seg(self, st: dict, a, b) -> None:
        (x1, y1), (x2, y2) = a, b
        if x1 == x2 and y1 != y2:
            self.lines.append(Line("v", x1, min(y1, y2), max(y1, y2), 0))
        elif y1 == y2 and x1 != x2:
            self.lines.append(Line("h", y1, min(x1, x2), max(x1, x2), 0))

    # ------------------------------------------------------------------ 足す場所
    def insertion_point(self) -> int:
        """新しいレコードを足す位置（この番号のレコードの前）：GDI の描く部分の終わり。

        終わりの EMR_EOF と、その前に続く EMF+ のコメントの前。EMF+ のある絵では、そこが EmfPlusGetDC の後
        （EMF+ に対応した再生でも GDI のレコードを描く区間）でなければ、足した字が描かれないので断る。"""
        at = len(self.records) - 1
        while at > 0:
            _p, t, _s = self.records[at - 1]
            if t == EMR_COMMENT and self.rec(at - 1)[12:16] == b"EMF+":
                at -= 1
                continue
            break
        before = [op for i, op in self.plus_ops if i < at]
        if before and before[-1] != EMFPLUS_GETDC:
            raise MetafileError("この表の絵は EMF+ で描かれていて、足した字が EMF+ に対応した表示で描かれません")
        if self.state_at[at]["clip"] == "complex":
            raise MetafileError("表の絵の終わりで、切り抜きが矩形でない状態のままです")
        return at

    # ------------------------------------------------------------------ 組み立て直す
    def rebuild(self, remove: set[int], insert: list[bytes], handles: int) -> bytes:
        """remove の番号のレコード（字）を除き、insert のレコードを足す所（insertion_point）に入れる。
        ほかのレコードはバイト列のまま。見出しのバイト数・レコード数・オブジェクトの表の大きさを直す。"""
        at = self.insertion_point()
        for i in remove:
            if self.records[i][1] not in (EMR_EXTTEXTOUTW, EMR_EXTTEXTOUTA):
                raise MetafileError("字のレコードでないものは取り除けません")
            if self.state_at[i]["align"] & TA_UPDATECP:
                raise MetafileError("描いた後に位置を進める字（TA_UPDATECP）は取り除けません")
        parts = []
        n = 0
        for i in range(len(self.records)):
            if i == at:
                parts += insert
                n += len(insert)
            if i in remove:
                continue
            parts.append(self.rec(i))
            n += 1
        out = bytearray(b"".join(parts))
        struct.pack_into("<IIH", out, 48, len(out), n, max(self.n_handles, handles))
        return bytes(out)


def _logfont(r: bytes) -> LogFont:
    elw = r[12:]
    if len(elw) < 92:
        raise MetafileError("字体のレコードが短すぎます")
    h, _w, _e, _o, weight, it, ul, so, cs = struct.unpack_from("<iiiiiBBBB", elw, 0)
    face = elw[28:92].decode("utf-16-le", errors="replace").split("\x00")[0]
    return LogFont(h, weight, bool(it), bool(ul), bool(so), cs, face, bytes(elw))


def _text_run(i: int, t: int, r: bytes, st: dict) -> TextRun:
    bounds = struct.unpack_from("<4i", r, 8)
    gmode, exs, eys = struct.unpack_from("<Iff", r, 24)
    x, y, nch, off_s, opts = struct.unpack_from("<iiIII", r, 36)
    off_dx = struct.unpack_from("<I", r, 72 if not opts & 0x100 else 56)[0]
    if t == EMR_EXTTEXTOUTW:
        text = r[off_s:off_s + 2 * nch].decode("utf-16-le", errors="replace")
    else:
        from .namecode import read_ansi

        text = read_ansi(r[off_s:off_s + nch], st["font"].charset if st["font"] is not None else None)[0]
    dx = []
    if off_dx and off_dx + 4 * nch <= len(r) and not opts & 0x2000:  # ETO_PDY のときは縦の送りも並ぶので読まない
        dx = list(struct.unpack_from(f"<{nch}i", r, off_dx))
    return TextRun(i, text, x, y, st["align"], st["font"], st["color"], dx, bounds, st["clip"], st["xform"], gmode,
                   (exs, eys), t)


# ====================================================================== レコードを作る
def _rec(t: int, body: bytes) -> bytes:
    if len(body) % 4:
        body += b"\x00" * (4 - len(body) % 4)
    return struct.pack("<II", t, 8 + len(body)) + body


def savedc() -> bytes:
    return _rec(EMR_SAVEDC, b"")


def restoredc() -> bytes:
    return _rec(EMR_RESTOREDC, struct.pack("<i", -1))


def intersect_clip(rect: tuple[int, int, int, int]) -> bytes:
    return _rec(EMR_INTERSECTCLIPRECT, struct.pack("<4i", *rect))


def select(ih: int) -> bytes:
    return _rec(EMR_SELECTOBJECT, struct.pack("<I", ih))


def delete(ih: int) -> bytes:
    return _rec(EMR_DELETEOBJECT, struct.pack("<I", ih))


def text_align(a: int) -> bytes:
    return _rec(EMR_SETTEXTALIGN, struct.pack("<I", a))


def text_color(colorref: int) -> bytes:
    return _rec(EMR_SETTEXTCOLOR, struct.pack("<I", colorref))


def bk_transparent() -> bytes:
    return _rec(EMR_SETBKMODE, struct.pack("<I", 1))


def create_font(ih: int, base: LogFont, height: int, weight: int, italic: bool, underline: bool, strike: bool,
                charset: int, face: str) -> bytes:
    """字体を作るレコード。絵の中にある字体のレコード（base）を写し、大きさ・太さ・斜体・下線・取り消し線・字の集合・書体名だけを変える。"""
    elw = bytearray(base.elw)
    struct.pack_into("<iiiii", elw, 0, height, 0, 0, 0, weight)
    elw[20], elw[21], elw[22], elw[23] = int(italic), int(underline), int(strike), charset
    name = face.encode("utf-16-le")[:62]
    elw[28:92] = name + b"\x00" * (64 - len(name))
    return _rec(EMR_EXTCREATEFONTINDIRECTW, struct.pack("<I", ih) + bytes(elw))


def ext_text_out(x: int, y: int, text: str, dx: list[int] | None, bounds: tuple[int, int, int, int], gmode: int,
                 scales: tuple[float, float]) -> bytes:
    """EMR_EXTTEXTOUTW：(x, y) を基準に text を描く。送り幅（dx）がなければ、字の幅は再生する側の字体に任せる。"""
    n = len(text)
    s = text.encode("utf-16-le")
    s += b"\x00" * ((4 - len(s) % 4) % 4)
    off_s = 76
    off_dx = off_s + len(s) if dx else 0
    body = struct.pack("<4iIff", *bounds, gmode, *scales)
    body += struct.pack("<iiIII4iI", x, y, n, off_s, 0, 0, 0, -1, -1, off_dx)
    body += s
    if dx:
        body += struct.pack(f"<{n}i", *dx)
    return _rec(EMR_EXTTEXTOUTW, body)


def colorref(rgb: str | None) -> int:
    """"RRGGBB" → COLORREF（0x00BBGGRR）。None は黒（Excel の自動の色）。"""
    if not rgb:
        return 0
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return r | (g << 8) | (b << 16)


# ====================================================================== MD4（RFC 1320）
def md4(msg: bytes) -> bytes:
    """MD4 の要約（RFC 1320 のとおり）。OfficeArt の絵（BLIP）の識別子に使う。"""

    def rol(v: int, s: int) -> int:
        v &= 0xFFFFFFFF
        return ((v << s) | (v >> (32 - s))) & 0xFFFFFFFF

    n = len(msg)
    msg = msg + b"\x80" + b"\x00" * ((55 - n) % 64) + struct.pack("<Q", (n * 8) & 0xFFFFFFFFFFFFFFFF)
    a, b, c, d = 0x67452301, 0xEFCDAB89, 0x98BADCFE, 0x10325476
    order2 = (0, 4, 8, 12, 1, 5, 9, 13, 2, 6, 10, 14, 3, 7, 11, 15)
    order3 = (0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15)
    for off in range(0, len(msg), 64):
        x = struct.unpack_from("<16I", msg, off)
        aa, bb, cc, dd = a, b, c, d
        for rnd, (order, shifts, k_add) in enumerate(((range(16), (3, 7, 11, 19), 0),
                                                      (order2, (3, 5, 9, 13), 0x5A827999),
                                                      (order3, (3, 9, 11, 15), 0x6ED9EBA1))):
            for j, k in enumerate(order):
                if rnd == 0:
                    f = lambda p, q, r: (p & q) | (~p & r)  # noqa: E731
                elif rnd == 1:
                    f = lambda p, q, r: (p & q) | (p & r) | (q & r)  # noqa: E731
                else:
                    f = lambda p, q, r: p ^ q ^ r  # noqa: E731
                s = shifts[j % 4]
                if j % 4 == 0:
                    a = rol(a + f(b, c, d) + x[k] + k_add, s)
                elif j % 4 == 1:
                    d = rol(d + f(a, b, c) + x[k] + k_add, s)
                elif j % 4 == 2:
                    c = rol(c + f(d, a, b) + x[k] + k_add, s)
                else:
                    b = rol(b + f(c, d, a) + x[k] + k_add, s)
        a, b, c, d = (a + aa) & 0xFFFFFFFF, (b + bb) & 0xFFFFFFFF, (c + cc) & 0xFFFFFFFF, (d + dd) & 0xFFFFFFFF
    return struct.pack("<4I", a, b, c, d)
