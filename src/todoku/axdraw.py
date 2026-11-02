"""ActiveX の部品の「控えの絵」を描き直す（EMF：MS-EMF・WMF：MS-WMF）。

Excel・Word は、ActiveX の部品を開いて動かすまで（ActiveX を止めている・保護ビュー・Mac 版・ブラウザーや
プレビューで見る・部品を読み込まないで印刷する）は、ファイルの中の控えの絵を表示する。値だけを書き換えると、
控えの絵は古い値のまま残り、画面や印刷が書いた値と食い違う。そこで、書いた部品の控えの絵を、書いた後の状態
（文字・チェック・選んだ項目・位置）で描き直す。

描き方は Forms 2.0 の部品の見た目に合わせる（くぼんだ枠・白い地・チェックの印・丸・押し下げたボタンなど）。
色はファイルの中の性質（BackColor・ForeColor：システムの色は Windows の標準の色）、字は TextProps の
字体名・大きさ・太さ・文字セット。絵の大きさと解像度は、元の控えの絵のヘッダー（EMF の枠と基準の装置、
WMF の置き場所の枠）と同じにする。字体の寸法は持たないので、字の送り幅は全角を1字、半角を半字と見積もる
（ＭＳ ゴシック・ＭＳ 明朝のような等幅の字体なら、そのとおりの幅）。
"""

from __future__ import annotations

import struct
import unicodedata

# OLE_COLOR のシステムの色（0x800000nn）→ Windows の標準の色（RGB）
SYSCOLOR = {0x00: 0xC8C8C8, 0x01: 0x000000, 0x02: 0x99B4D1, 0x03: 0xBFCDDB, 0x04: 0xF0F0F0, 0x05: 0xFFFFFF,
            0x06: 0x646464, 0x07: 0x000000, 0x08: 0x000000, 0x09: 0x000000, 0x0A: 0xB4B4B4, 0x0B: 0xF4F7FC,
            0x0C: 0xABABAB, 0x0D: 0x0078D7, 0x0E: 0xFFFFFF, 0x0F: 0xF0F0F0, 0x10: 0xA0A0A0, 0x11: 0x6D6D6D,
            0x12: 0x000000, 0x13: 0x000000, 0x14: 0xFFFFFF, 0x15: 0x696969, 0x16: 0xE3E3E3, 0x17: 0x000000,
            0x18: 0xFFFFE1}
FACE, LIGHT, HILITE, SHADOW, DKSHADOW = 0xF0F0F0, 0xE3E3E3, 0xFFFFFF, 0xA0A0A0, 0x696969
GRAYTEXT, SELECTED = 0x6D6D6D, 0x0078D7


def rgb(ole: int | None, fallback: int = 0x000000) -> int:
    """OLE_COLOR（2.4.9）→ RGB（0xRRGGBB）。"""
    if ole is None:
        return fallback
    typ = ole >> 24 & 0xFF
    if typ == 0x80:
        return SYSCOLOR.get(ole & 0xFF, fallback)
    r, g, b = ole & 0xFF, ole >> 8 & 0xFF, ole >> 16 & 0xFF
    return r << 16 | g << 8 | b


def _colorref(c: int) -> int:
    """RGB（0xRRGGBB）→ COLORREF（0x00BBGGRR）。"""
    return (c & 0xFF) << 16 | (c & 0xFF00) | (c >> 16 & 0xFF)


def char_width(ch: str, em: float) -> float:
    return em if unicodedata.east_asian_width(ch) in ("W", "F") else em / 2


def text_width(s: str, em: float) -> float:
    return sum(char_width(ch, em) for ch in s)


def advances(s: str, px: float) -> list[int]:
    """字ごとの送り幅（ピクセル）。足し合わせた位置が見積もりの幅とずれないように丸める。"""
    out, at, acc = [], 0.0, 0
    for ch in s:
        at += char_width(ch, px)
        w = round(at) - acc
        out.append(w)
        acc += w
    return out


def wrap(text: str, width: float, em: float) -> list[str]:
    """折り返し：行ごとに、幅に収まる字まで（英字の単語の途中では、なるべく切らない）。"""
    out = []
    for para in text.split("\n"):
        line, w = "", 0.0
        for ch in para:
            cw = char_width(ch, em)
            if line and w + cw > width:
                cut = line.rfind(" ")
                if cut > 0 and ch != " " and unicodedata.east_asian_width(ch) not in ("W", "F"):
                    out.append(line[:cut])
                    line = line[cut + 1:]
                    w = text_width(line, em)
                else:
                    out.append(line)
                    line, w = "", 0.0
            line += ch
            w += cw
        out.append(line)
    return out


# ====================================================================== 描く中身（どちらの形式にも書ける命令の並び）
class Canvas:
    """部品の絵の命令の並び（ピクセル単位）。"""

    def __init__(self, w: int, h: int):
        self.w, self.h = max(int(w), 1), max(int(h), 1)
        self.ops: list[tuple] = []

    def fill(self, x0, y0, x1, y1, color):
        """左上 (x0, y0) から右下 (x1, y1) の手前まで塗る。"""
        if x1 > x0 and y1 > y0:
            self.ops.append(("fill", int(x0), int(y0), int(x1), int(y1), color))

    def line(self, pts, color, width=1):
        self.ops.append(("line", [(int(x), int(y)) for x, y in pts], color, width))

    def ellipse(self, x0, y0, x1, y1, fill, edge):
        self.ops.append(("ellipse", int(x0), int(y0), int(x1), int(y1), fill, edge))

    def polygon(self, pts, color):
        self.ops.append(("polygon", [(int(x), int(y)) for x, y in pts], color))

    def text(self, x, y, s, font: dict, color, clip):
        if s:
            self.ops.append(("text", int(round(x)), int(round(y)), s, dict(font), color, tuple(int(v) for v in clip)))

    def edge(self, x0, y0, x1, y1, outer_tl, outer_br, inner_tl=None, inner_br=None):
        """立体の枠（外側と内側の2本の線）。右下の手前 (x1, y1) までの箱。"""
        self.fill(x0, y0, x1, y0 + 1, outer_tl)
        self.fill(x0, y0, x0 + 1, y1, outer_tl)
        self.fill(x0, y1 - 1, x1, y1, outer_br)
        self.fill(x1 - 1, y0, x1, y1, outer_br)
        if inner_tl is not None:
            self.fill(x0 + 1, y0 + 1, x1 - 1, y0 + 2, inner_tl)
            self.fill(x0 + 1, y0 + 1, x0 + 2, y1 - 1, inner_tl)
            self.fill(x0 + 1, y1 - 2, x1 - 1, y1 - 1, inner_br)
            self.fill(x1 - 2, y0 + 1, x1 - 1, y1 - 1, inner_br)

    def texts(self) -> list[str]:
        return [op[3] for op in self.ops if op[0] == "text"]


def _frame(cv: Canvas, x0, y0, x1, y1, effect: int, border: bool, border_color: int) -> int:
    """SpecialEffect（2.5.80）に応じた枠。2 くぼみ・1 盛り上がり・3 彫り・6 浮き彫り・0 平ら（BorderStyle が 1 なら1本の線）。
    枠の太さ（内側の始まり）を返す。"""
    if effect == 2:
        cv.edge(x0, y0, x1, y1, SHADOW, HILITE, DKSHADOW, LIGHT)
        return 2
    if effect == 1:
        cv.edge(x0, y0, x1, y1, LIGHT, DKSHADOW, HILITE, SHADOW)
        return 2
    if effect == 3:
        cv.edge(x0, y0, x1, y1, SHADOW, HILITE, HILITE, SHADOW)
        return 2
    if effect == 6:
        cv.edge(x0, y0, x1, y1, HILITE, SHADOW, SHADOW, HILITE)
        return 2
    if border:
        cv.edge(x0, y0, x1, y1, border_color, border_color)
        return 1
    return 0


def render(ctl, w: int, h: int, dpi: float, items: list[str] | None = None) -> Canvas:
    """部品 ctl（activex.Control）の、いまの状態の絵。w・h はピクセル、dpi は字の大きさを点からピクセルにするため。"""
    cv = Canvas(w, h)
    k = ctl.kind
    get = ctl.get
    back = rgb(get("BackColor"), 0xFFFFFF)
    fore = rgb(get("ForeColor"))
    enabled = ctl.bit("Enabled")
    if not enabled:
        fore = GRAYTEXT
    f = ctl.font
    em = max(1.0, (f["height"] or 160) / 20 * dpi / 72)
    font = {"name": f["name"] or "MS Sans Serif", "px": em, "weight": f["weight"] or 400, "charset": f["charset"],
            "effects": f["effects"] or 0}
    align = f["align"] or 1  # 1 左・2 右・3 中央（PARAFORMAT の揃え）
    opaque = ctl.bit("BackStyle")
    border = (get("BorderStyle") or 0) == 1
    effect = get("SpecialEffect") or 0
    bcol = rgb(get("BorderColor"), 0x646464)
    W, H = cv.w, cv.h
    lh = em * 1.2

    def place(s: str, x0, x1, y_top, color=fore, al=align):
        tw = text_width(s, em)
        x = x0 if al == 1 else (x1 - tw if al == 2 else x0 + (x1 - x0 - tw) / 2)
        cv.text(max(x, x0), y_top, s, font, color, (x0, 0, x1, H))

    if k in ("TextBox", "ComboBox", "ListBox"):
        if opaque or k == "ListBox":
            cv.fill(0, 0, W, H, back)
        inset = _frame(cv, 0, 0, W, H, effect, border, bcol)
        right = W - inset
        if k in ("ComboBox", "TextBox") and get("ShowDropButtonWhen") == 2:  # いつもドロップボタンを出す
            bw = max(min(int(H * 0.75), 17, W // 2), 1)
            right = W - inset - bw
            cv.fill(right, inset, W - inset, H - inset, FACE)
            cv.edge(right, inset, W - inset, H - inset, LIGHT, DKSHADOW, HILITE, SHADOW)
            cx, cy = (right + W - inset) // 2, H // 2
            cv.polygon([(cx - 3, cy - 1), (cx + 4, cy - 1), (cx, cy + 3)], 0x000000 if enabled else GRAYTEXT)
        text = ctl.value
        if k == "TextBox" and get("PasswordChar"):
            text = get("PasswordChar") * len(text)
        x0 = inset + 2
        if k == "ListBox":
            rows = list(items) if items else ([text] if text else [])
            for i, item in enumerate(rows):
                y = inset + 1 + i * lh
                if y > H:
                    break
                if text and item == text:
                    cv.fill(inset, y, right, min(y + lh, H - inset), SELECTED)
                    place(item, x0, right - 1, y + (lh - em) / 2, color=0xFFFFFF, al=1)
                else:
                    place(item, x0, right - 1, y + (lh - em) / 2, al=1)
            return cv
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        if k == "TextBox" and ctl.bit("MultiLine"):
            lines = wrap(text, right - x0 - 2, em) if ctl.bit("WordWrap") else text.split("\n")
            for i, line in enumerate(lines):
                place(line, x0, right - 1, inset + 1 + i * lh)
        else:
            place(text.replace("\n", " "), x0, right - 1, inset + 1)
        return cv
    if k in ("CheckBox", "OptionButton"):
        if opaque:
            cv.fill(0, 0, W, H, back)
        box = max(min(13, H, W), 1)
        left_caption = ctl.bit("Alignment")  # 1：見出しが左、印が右（2.5.96.1 と注 15）
        bx = W - box - 1 if left_caption else 1
        by = (H - box) // 2
        st = ctl.checked
        if k == "CheckBox":
            cv.fill(bx, by, bx + box, by + box, 0xFFFFFF if st is not None else LIGHT)
            if effect == 0:
                cv.edge(bx, by, bx + box, by + box, 0x000000, 0x000000)
            else:
                cv.edge(bx, by, bx + box, by + box, SHADOW, HILITE, DKSHADOW, LIGHT)
            if st is not False:
                mark = 0x000000 if st else SHADOW
                x, y = bx + 3, by + box // 2
                for d in range(3):
                    cv.line([(x, y + d - 1), (x + 2, y + d + 1), (x + 7, y + d - 4)], mark)
        else:
            cv.ellipse(bx, by, bx + box, by + box, 0xFFFFFF if st is not None else LIGHT, SHADOW)
            cv.ellipse(bx + 1, by + 1, bx + box - 1, by + box - 1, None, DKSHADOW)
            if st:
                c = box // 2
                cv.ellipse(bx + c - 2, by + c - 2, bx + c + 3, by + c + 3, 0x000000, 0x000000)
        x0, x1 = (1, bx - 3) if left_caption else (bx + box + 4, W - 1)
        cap = ctl.caption
        if ctl.bit("WordWrap") and text_width(cap, em) > x1 - x0:
            lines = wrap(cap, x1 - x0, em)
            top = (H - len(lines) * lh) / 2
            for i, line in enumerate(lines):
                place(line, x0, x1, top + i * lh, al=1)
        else:
            place(cap, x0, x1, (H - em) / 2, al=1)
        return cv
    if k == "ToggleButton":
        st = ctl.checked
        cv.fill(0, 0, W, H, back if not st else 0xF8F8F8)
        d = 1 if st else 0
        if st:
            cv.edge(0, 0, W, H, DKSHADOW, HILITE, SHADOW, LIGHT)
        else:
            cv.edge(0, 0, W, H, HILITE, DKSHADOW, LIGHT, SHADOW)
        place(ctl.caption, 3 + d, W - 3 + d, (H - em) / 2 + d, al=3 if align == 1 else align)
        return cv
    if k in ("SpinButton", "ScrollBar"):
        o = get("Orientation")
        vertical = o == 0 or (o not in (0, 1) and H >= W)  # -1：縦横の長い方（fmOrientationAuto）
        cv.fill(0, 0, W, H, FACE if k == "SpinButton" else 0xF0F0F0)
        lo, hi = min(get("Min"), get("Max")), max(get("Min"), get("Max"))
        pos = get("Position")
        arrow = 0x000000 if enabled else GRAYTEXT
        a = min(W, H // 2) if vertical else min(H, W // 2)
        if k == "SpinButton":
            a = H // 2 if vertical else W // 2
        boxes = [(0, 0, W, a), (0, H - a, W, H)] if vertical else [(0, 0, a, H), (W - a, 0, W, H)]
        for i, (x0, y0, x1, y1) in enumerate(boxes):
            cv.fill(x0, y0, x1, y1, FACE)
            cv.edge(x0, y0, x1, y1, LIGHT, DKSHADOW, HILITE, SHADOW)
            cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
            if vertical:
                pts = [(cx - 3, cy + 1), (cx + 4, cy + 1), (cx, cy - 3)] if i == 0 else \
                    [(cx - 3, cy - 1), (cx + 4, cy - 1), (cx, cy + 3)]
            else:
                pts = [(cx + 1, cy - 4), (cx + 1, cy + 4), (cx - 3, cy)] if i == 0 else \
                    [(cx - 1, cy - 4), (cx - 1, cy + 4), (cx + 3, cy)]
            cv.polygon(pts, arrow)
        if k == "ScrollBar" and hi > lo:
            span = (H if vertical else W) - 2 * a
            thumb = max(8, min(span, a))
            t = a + (span - thumb) * (pos - lo) / (hi - lo)
            box_ = (0, t, W, t + thumb) if vertical else (t, 0, t + thumb, H)
            cv.fill(*box_, FACE)
            cv.edge(*box_, LIGHT, DKSHADOW, HILITE, SHADOW)
        return cv
    # ラベル・コマンドボタンなど（値を書かない部品）：見出しだけ
    cv.fill(0, 0, W, H, back)
    place(ctl.caption, 2, W - 2, (H - em) / 2)
    return cv


# ====================================================================== 物の番号
class _Handles:
    """絵の中の物（ペン・ブラシ・字体）の番号：消した番号は次に使い回す（いちばん小さい空き番号）。"""

    def __init__(self, first: int):
        self.free: list[int] = []
        self.next = first
        self.top = first - 1

    def new(self) -> int:
        if self.free:
            return self.free.pop(0)
        self.next += 1
        self.top = max(self.top, self.next - 1)
        return self.next - 1

    def drop(self, i: int) -> None:
        self.free.append(i)
        self.free.sort()


# ====================================================================== EMF
NULL_PEN, NULL_BRUSH, SYSTEM_FONT = 0x80000008, 0x80000005, 0x8000000D  # ストックオブジェクト（MS-EMF 2.1.31）


def _rec(typ: int, body: bytes) -> bytes:
    return struct.pack("<II", typ, 8 + len(body)) + body


def _logfont_w(font: dict) -> bytes:
    """LogFontExDv（2.2.15）：LogFont（92）・FullName・Style・Script・DesignVector（印と軸の数 0）。"""
    face = (font["name"] or "")[:31].encode("utf-16-le")
    eff = font.get("effects") or 0
    cs = font.get("charset")
    lf = struct.pack("<iiiii8B", -max(1, round(font["px"])), 0, 0, 0, int(font.get("weight") or 400),
                     1 if eff & 2 else 0, 1 if eff & 4 else 0, 1 if eff & 8 else 0, (1 if cs is None else int(cs)) & 0xFF,
                     0, 0, 0, 0) + face + bytes(64 - len(face))
    return lf + bytes(128 + 64 + 64) + struct.pack("<II", 0x08007664, 0)


def to_emf(cv: Canvas, frame: tuple[int, int, int, int] | None = None, device=(1920, 1080), mm=(508, 285)) -> bytes:
    """Canvas を EMF にする。frame は絵の大きさ（0.01 mm の RectL）。device・mm は基準の装置（元の絵のヘッダーと同じにする）。
    物は使ったらすぐ消す（消す前に、決まった物（ストックオブジェクト）を選び直す）。"""
    W, H = cv.w, cv.h
    recs: list[bytes] = []
    hs = _Handles(1)

    def emr(typ, body):
        recs.append(_rec(typ, body))

    def sel(h):
        emr(0x25, struct.pack("<I", h))  # SELECTOBJECT

    def brush(color) -> int:
        h = hs.new()
        emr(0x27, struct.pack("<IIII", h, 0, _colorref(color), 0))  # CREATEBRUSHINDIRECT（BS_SOLID）
        return h

    def pen(color, width=1) -> int:
        h = hs.new()
        emr(0x26, struct.pack("<IIiiI", h, 0, width, 0, _colorref(color)))  # CREATEPEN（PS_SOLID）
        return h

    def done(*objs):
        sel(NULL_PEN)
        sel(NULL_BRUSH)
        for h in objs:
            if h is not None:
                emr(0x28, struct.pack("<I", h))  # DELETEOBJECT
                hs.drop(h)

    emr(0x11, struct.pack("<I", 8))  # SETMAPMODE：MM_ANISOTROPIC
    emr(0x09, struct.pack("<ii", W, H))  # SETWINDOWEXTEX
    emr(0x0B, struct.pack("<ii", W, H))  # SETVIEWPORTEXTEX
    emr(0x0A, struct.pack("<ii", 0, 0))  # SETWINDOWORGEX
    emr(0x12, struct.pack("<I", 1))  # SETBKMODE：TRANSPARENT
    emr(0x16, struct.pack("<I", 0))  # SETTEXTALIGN：TA_LEFT | TA_TOP
    for op in cv.ops:
        kind = op[0]
        if kind == "fill":
            _, x0, y0, x1, y1, color = op
            b = brush(color)
            sel(b)
            sel(NULL_PEN)
            emr(0x2B, struct.pack("<iiii", x0, y0, x1 + 1, y1 + 1))  # RECTANGLE（線なしは右下の1つ手前まで塗る）
            done(b)
        elif kind in ("line", "polygon"):
            pts, color = op[1], op[2]
            p = pen(color, op[3] if kind == "line" else 1)
            sel(p)
            b = brush(color) if kind == "polygon" else None
            sel(b if b is not None else NULL_BRUSH)
            xs, ys = [q[0] for q in pts], [q[1] for q in pts]
            body = struct.pack("<iiiiI", min(xs), min(ys), max(xs), max(ys), len(pts))
            body += b"".join(struct.pack("<hh", x, y) for x, y in pts)
            emr(0x56 if kind == "polygon" else 0x57, body)  # POLYGON16 / POLYLINE16
            done(p, b)
        elif kind == "ellipse":
            _, x0, y0, x1, y1, fill, edge = op
            p = pen(edge)
            sel(p)
            b = brush(fill) if fill is not None else None
            sel(b if b is not None else NULL_BRUSH)
            emr(0x2A, struct.pack("<iiii", x0, y0, x1, y1))  # ELLIPSE
            done(p, b)
        elif kind == "text":
            _, x, y, s, font, color, clip = op
            f = hs.new()
            emr(0x52, struct.pack("<I", f) + _logfont_w(font))  # EXTCREATEFONTINDIRECTW
            sel(f)
            emr(0x18, struct.pack("<I", _colorref(color)))  # SETTEXTCOLOR
            raw = s.encode("utf-16-le")
            nchar = len(raw) // 2
            dx = []
            for ch, wd in zip(s, advances(s, font["px"])):
                dx += [wd] if ord(ch) < 0x10000 else [wd, 0]  # サロゲートペア：UTF-16 の単位ごとに送り幅
            off_str = 76  # レコードの頭から文字まで（8 + 16 + 12 + EmrText の 40）
            off_dx = off_str + len(raw) + (-len(raw)) % 4
            bounds = (x, y, x + sum(dx), y + int(font["px"] * 1.2))
            body = struct.pack("<iiii", *bounds) + struct.pack("<Iff", 1, 0.0, 0.0)  # GM_COMPATIBLE
            body += struct.pack("<iiIII", x, y, nchar, off_str, 0x04)  # ETO_CLIPPED
            body += struct.pack("<iiii", *clip) + struct.pack("<I", off_dx)
            body += raw + bytes((-len(raw)) % 4) + struct.pack(f"<{nchar}i", *dx)
            emr(0x54, body)  # EXTTEXTOUTW
            sel(SYSTEM_FONT)
            emr(0x28, struct.pack("<I", f))
            hs.drop(f)
    emr(0x0E, struct.pack("<III", 0, 16, 20))  # EOF（パレットなし。最後の欄はこのレコードの大きさ）
    if frame is None:
        frame = (0, 0, round(W * mm[0] * 100 / device[0]), round(H * mm[1] * 100 / device[1]))
    nbytes = 108 + sum(len(r) for r in recs)
    head = struct.pack("<II", 1, 108) + struct.pack("<iiii", 0, 0, W - 1, H - 1) + struct.pack("<iiii", *frame)
    head += struct.pack("<IIIIHHIII", 0x464D4520, 0x00010000, nbytes, len(recs) + 1, hs.top + 1, 0, 0, 0, 0)
    head += struct.pack("<ii", *device) + struct.pack("<ii", *mm)
    head += struct.pack("<III", 0, 0, 0) + struct.pack("<ii", mm[0] * 1000, mm[1] * 1000)
    return head + b"".join(recs)


def emf_header(data: bytes) -> dict | None:
    """EMF のヘッダー（2.2.9）：枠（0.01 mm）・装置の大きさ（ピクセル・mm）・絵の大きさ（ピクセル）。EMF でなければ None。"""
    if len(data) < 88 or struct.unpack_from("<I", data, 0)[0] != 1 or struct.unpack_from("<I", data, 40)[0] != 0x464D4520:
        return None
    bounds = struct.unpack_from("<iiii", data, 8)
    frame = struct.unpack_from("<iiii", data, 24)
    device = struct.unpack_from("<ii", data, 72)
    mm = struct.unpack_from("<ii", data, 80)
    w = round((frame[2] - frame[0]) / 100 * device[0] / mm[0]) if mm[0] else bounds[2] - bounds[0] + 1
    h = round((frame[3] - frame[1]) / 100 * device[1] / mm[1]) if mm[1] else bounds[3] - bounds[1] + 1
    return {"bounds": bounds, "frame": frame, "device": device, "mm": mm, "w": w, "h": h,
            "dpi": device[0] / mm[0] * 25.4 if mm[0] else 96.0}


def emf_records(data: bytes) -> list[tuple[int, bytes]]:
    """EMF のレコードを順に（種類, 中身）。大きさが合わなければ ValueError。ヘッダーの数と大きさも確かめる。"""
    out = []
    p = 0
    while p + 8 <= len(data):
        typ, size = struct.unpack_from("<II", data, p)
        if size < 8 or size % 4 or p + size > len(data):
            raise ValueError(f"EMF のレコード（{p} バイト目）の大きさが不正です")
        out.append((typ, data[p + 8:p + size]))
        p += size
        if typ == 0x0E:
            break
    if p != len(data):
        raise ValueError("EMF の終わり（EOF）の後ろにバイトがあります")
    if not out or out[0][0] != 1 or out[-1][0] != 0x0E:
        raise ValueError("EMF の頭（HEADER）か終わり（EOF）がありません")
    nbytes, nrec = struct.unpack_from("<II", out[0][1], 40)
    if nbytes != len(data) or nrec != len(out):
        raise ValueError("EMF のヘッダーの大きさ・レコードの数が、中身と合いません")
    return out


def emf_texts(data: bytes) -> list[str]:
    """EMF の中の、字を書く命令（EXTTEXTOUTW）の字。"""
    out = []
    for typ, body in emf_records(data):
        if typ == 0x54:
            n, off = struct.unpack_from("<II", body, 36)
            out.append(body[off - 8:off - 8 + 2 * n].decode("utf-16-le", errors="replace"))
    return out


# ====================================================================== WMF
def _wrec(func: int, body: bytes) -> bytes:
    if len(body) % 2:
        body += b"\x00"
    return struct.pack("<IH", 3 + len(body) // 2, func) + body


_CODEPAGE = {128: "cp932", 0: "cp1252", 1: "cp1252", 129: "cp949", 134: "gbk", 136: "cp950"}


def to_wmf(cv: Canvas, box: tuple[int, int, int, int] | None = None, inch: int = 96, placeable: bool = True) -> bytes:
    """Canvas を WMF（MS-WMF）にする。box・inch は置き場所の頭（Placeable）の枠と1インチの単位（元の絵と同じにする）。
    字は字体の文字セットの符号（日本語は Shift_JIS）で書く。符号にできない字があれば ValueError（字を落とした絵は作らない）。"""
    W, H = cv.w, cv.h
    recs: list[bytes] = []
    hs = _Handles(0)

    def wr(func, body):
        recs.append(_wrec(func, body))

    def cref(c):
        return struct.pack("<I", _colorref(c))

    def brush(color, style=0) -> int:
        h = hs.new()
        wr(0x02FC, struct.pack("<H", style) + cref(color) + struct.pack("<H", 0))  # CREATEBRUSHINDIRECT
        return h

    def pen(color, width=1, style=0) -> int:
        h = hs.new()
        wr(0x02FA, struct.pack("<Hhh", style, width, 0) + cref(color))  # CREATEPENINDIRECT
        return h

    def use(*objs):
        for h in objs:
            wr(0x012D, struct.pack("<H", h))  # SELECTOBJECT

    def drop(*objs):
        for h in objs:
            wr(0x01F0, struct.pack("<H", h))  # DELETEOBJECT
            hs.drop(h)

    wr(0x0103, struct.pack("<h", 8))  # SETMAPMODE：MM_ANISOTROPIC
    wr(0x020B, struct.pack("<hh", 0, 0))  # SETWINDOWORG（y, x）
    wr(0x020C, struct.pack("<hh", H, W))  # SETWINDOWEXT（y, x）
    wr(0x0102, struct.pack("<h", 1))  # SETBKMODE：TRANSPARENT
    wr(0x012E, struct.pack("<H", 0))  # SETTEXTALIGN
    for op in cv.ops:
        kind = op[0]
        if kind == "fill":
            _, x0, y0, x1, y1, color = op
            b, p = brush(color), pen(0, 0, 5)  # PS_NULL
            use(b, p)
            wr(0x041B, struct.pack("<hhhh", y1 + 1, x1 + 1, y0, x0))  # RECTANGLE（bottom, right, top, left）
            drop(b, p)
        elif kind in ("line", "polygon"):
            pts, color = op[1], op[2]
            p = pen(color, op[3] if kind == "line" else 1)
            b = brush(color) if kind == "polygon" else brush(0, 1)  # 線だけのときは中を塗らない（BS_NULL）
            use(p, b)
            body = struct.pack("<h", len(pts)) + b"".join(struct.pack("<hh", x, y) for x, y in pts)
            wr(0x0324 if kind == "polygon" else 0x0325, body)  # POLYGON / POLYLINE
            drop(p, b)
        elif kind == "ellipse":
            _, x0, y0, x1, y1, fill, edge = op
            p = pen(edge)
            b = brush(0, 1) if fill is None else brush(fill)
            use(p, b)
            wr(0x0418, struct.pack("<hhhh", y1, x1, y0, x0))  # ELLIPSE
            drop(p, b)
        elif kind == "text":
            _, x, y, s, font, color, clip = op
            cs = font.get("charset")
            cs = (1 if cs is None else int(cs)) & 0xFF
            enc = _CODEPAGE.get(cs, "cp1252")
            try:
                raw = s.encode(enc)
            except UnicodeEncodeError:
                # 字体の文字セットで書けない字（英語の字体の部品に日本語を書いた）：Windows が字を代わりの字体で出すのと同じく、
                # 日本語の文字セット（SHIFTJIS_CHARSET）で書く
                try:
                    cs, enc = 128, "cp932"
                    raw = s.encode(enc)
                except UnicodeEncodeError:
                    raise ValueError(f"「{s}」には、WMF の字（Shift_JIS）で書けない字があります") from None
            f = hs.new()
            face = (font["name"] or "").encode(enc, errors="replace")[:31]
            eff = font.get("effects") or 0
            lf = struct.pack("<hhhhh8B", -max(1, round(font["px"])), 0, 0, 0, int(font.get("weight") or 400),
                             1 if eff & 2 else 0, 1 if eff & 4 else 0, 1 if eff & 8 else 0, cs, 0, 0, 0, 0)
            wr(0x02FB, lf + face + bytes(32 - len(face)))  # CREATEFONTINDIRECT（字体名は 32 バイトの欄）
            use(f)
            wr(0x0209, cref(color))  # SETTEXTCOLOR
            dx: list[int] = []
            for ch, wd in zip(s, advances(s, font["px"])):
                dx += [wd] + [0] * (len(ch.encode(enc)) - 1)  # 2バイトの字は、初めのバイトに送り幅
            body = struct.pack("<hhhH", y, x, len(raw), 0x04) + struct.pack("<hhhh", *clip)  # ETO_CLIPPED
            body += raw + (b"\x00" if len(raw) % 2 else b"") + struct.pack(f"<{len(dx)}h", *dx)
            wr(0x0A32, body)  # EXTTEXTOUT
            drop(f)
    wr(0x0000, b"")  # EOF
    body = b"".join(recs)
    max_rec = max(struct.unpack_from("<I", r, 0)[0] for r in recs)
    head = struct.pack("<HHHIHIH", 1, 9, 0x0300, (18 + len(body)) // 2, hs.top + 1, max_rec, 0)
    out = head + body
    if not placeable:
        return out
    l, t, r, b = box if box is not None else (0, 0, W, H)
    ph = struct.pack("<IHhhhhHI", 0x9AC6CDD7, 0, l, t, r, b, inch, 0)
    check = 0
    for wd in struct.unpack("<10H", ph):
        check ^= wd
    return ph + struct.pack("<H", check) + out


def wmf_header(data: bytes) -> dict | None:
    """WMF の大きさ：置き場所の頭（Placeable、22バイト）があれば、その枠と1インチの単位。"""
    if len(data) >= 22 and struct.unpack_from("<I", data, 0)[0] == 0x9AC6CDD7:
        _k, _h, l, t, r, b, inch = struct.unpack_from("<IHhhhhH", data, 0)
        return {"placeable": True, "box": (l, t, r, b), "inch": inch, "offset": 22}
    if len(data) >= 18 and struct.unpack_from("<H", data, 2)[0] == 9 and struct.unpack_from("<H", data, 0)[0] in (1, 2):
        return {"placeable": False, "offset": 0}
    return None


def wmf_records(data: bytes) -> list[tuple[int, bytes]]:
    """WMF のレコードを順に（関数の番号, 中身）。大きさがヘッダーと合わなければ ValueError。"""
    hd = wmf_header(data)
    if hd is None:
        raise ValueError("WMF ではありません")
    p = hd["offset"]
    size_words = struct.unpack_from("<I", data, p + 6)[0]
    if size_words * 2 != len(data) - p:
        raise ValueError("WMF のヘッダーの大きさが、中身と合いません")
    p += 18
    out = []
    while p + 6 <= len(data):
        size, func = struct.unpack_from("<IH", data, p)
        if size < 3 or p + size * 2 > len(data):
            raise ValueError(f"WMF のレコード（{p} バイト目）の大きさが不正です")
        out.append((func, data[p + 6:p + size * 2]))
        p += size * 2
        if func == 0:
            break
    if p != len(data) or not out or out[-1][0] != 0:
        raise ValueError("WMF の終わり（EOF）が合いません")
    return out


def wmf_texts(data: bytes, encoding: str = "cp932") -> list[str]:
    out = []
    for func, body in wmf_records(data):
        if func == 0x0A32:
            n, opts = struct.unpack_from("<hH", body, 4)
            at = 8 + (8 if opts & 0x06 else 0)
            out.append(body[at:at + n].decode(encoding, errors="replace"))
    return out


def picture(ctl, original: bytes, items: list[str] | None = None) -> bytes:
    """部品 ctl の、いまの状態の控えの絵を、元の控えの絵（original）と同じ形式・大きさで作る。"""
    hd = emf_header(original)
    if hd is not None:
        cv = render(ctl, hd["w"], hd["h"], hd["dpi"], items)
        return to_emf(cv, hd["frame"], hd["device"], hd["mm"])
    wh = wmf_header(original)
    if wh is not None:
        if wh["placeable"]:
            l, t, r, b = wh["box"]
            inch = wh["inch"] or 1440
            w, h = round(abs(r - l) * 96 / inch), round(abs(b - t) * 96 / inch)
            return to_wmf(render(ctl, w, h, 96.0, items), (l, t, r, b), inch)
        sw, sh = ctl.size
        cv = render(ctl, round(sw / 2540 * 96), round(sh / 2540 * 96), 96.0, items)
        return to_wmf(cv, placeable=False)
    raise ValueError("控えの絵が EMF でも WMF でもないので、描き直せません")


def picture_texts(data: bytes) -> list[str]:
    """控えの絵（EMF か WMF）の中の字。"""
    if emf_header(data) is not None:
        return emf_texts(data)
    return wmf_texts(data)


def check_picture(data: bytes) -> list[str]:
    """控えの絵の形の検査（レコードの大きさ・ヘッダーの数・終わり）。問題の一覧。"""
    try:
        if emf_header(data) is not None:
            emf_records(data)
        else:
            wmf_records(data)
    except ValueError as e:
        return [str(e)]
    return []
