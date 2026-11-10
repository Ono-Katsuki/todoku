"""数式エディター 3.0（Equation.3）の数式を、字の並びにする（Equation Native のストリームの MTEF 版 3）。

Equation Native は、先頭に長さ cbHdr（最初の 2 バイト。ふつう 28）の見出しがあり、その後ろが MTEF。
MTEF 版 3（Design Science の公開の仕様 "MTEF v.3 (Equation Editor 3.x)"）：
- 見出し 5 バイト：版（3）・環境（0 Mac・1 Windows）・製品（0 MathType・1 数式エディター）・製品の版・副版
- レコード：1 バイトの印（下の 4 ビットが種類、上の 4 ビットが選択）。0 END・1 LINE・2 CHAR・3 TMPL・4 PILE・5 MATRIX・
  6 EMBELL・7 RULER・8 FONT・9 SIZE・10〜14 FULL/SUB/SUB2/SYM/SUBSYM。選択 0x8（xfLMOVE）があれば、位置のずれ（nudge）が続く
- LINE：［ずれ］［行の間隔（0x4）］［RULER（0x2）］中身の並び（0x1 xfNULL なら中身なし）
- CHAR：［ずれ］字体（値＋128）・字（2 バイト。文字の字体では Unicode、記号の字体では Symbol フォントの番号。MT Extra の字は、対応表を確かめていないので、番号のまま示す）・［飾り（0x2）］
- TMPL：［ずれ］型・変種・選択・中身の並び（枠の LINE と、かっこなどの字）。PILE・MATRIX は行（LINE）の並び

ここでは、決まった規則だけで字を取り出す：分数（型 14）は (分子)/(分母)、根号（13）は √(…)、添え字（15）は _(…)・^(…)、
それ以外の型は中身をそのまま順に並べる。読めない所に来たら、そこまでの字を返す（理由を添える）。
"""

from __future__ import annotations

import struct

from .. import symfont

_SYMBOL_FACES = {4, 5, 6}  # 小文字のギリシャ文字・大文字のギリシャ文字・記号：字は Symbol フォントの番号（対応表は symfont）
_MT_EXTRA = 11  # MT Extra：対応表を確かめていない書体。字を決めず、書体の名前と番号を示す


class MtefError(ValueError):
    pass


def equation_text(native: bytes) -> tuple[str, str]:
    """Equation Native の中身 → (数式の字, 読めなかった理由。読めたら空)。"""
    if len(native) < 2:
        return "", "Equation Native が短すぎます"
    cb = struct.unpack_from("<H", native, 0)[0]
    data = native[cb:] if 0 < cb < len(native) else native
    if len(data) < 5 or data[0] not in (2, 3):
        return "", f"MTEF の版（{data[0] if data else '?'}）に対応していません"
    r = _Reader(data, 5)
    out: list[str] = []
    try:
        while r.pos < len(data):
            obj = r.obj()
            if obj is None:
                break
            out.append(obj)
    except (MtefError, IndexError, struct.error) as e:
        return "".join(out), f"数式のデータの途中で読めなくなりました（{e}）"
    return "".join(out).strip(), ""


class _Reader:
    def __init__(self, data: bytes, pos: int):
        self.d = data
        self.pos = pos
        self.depth = 0

    def u8(self) -> int:
        v = self.d[self.pos]
        self.pos += 1
        return v

    def nudge(self, opts: int) -> None:
        if opts & 0x8:
            dx, dy = self.u8(), self.u8()
            if dx == 128 and dy == 128:
                self.pos += 4

    def objects(self) -> list[str]:
        """END までの中身の並び。"""
        out = []
        while True:
            o = self.obj()
            if o is None:
                return out
            out.append(o)

    def obj(self) -> str | None:
        """1つのレコード。END なら None、字を持たないレコードは空の字。"""
        if self.pos >= len(self.d):
            return None
        tag = self.u8()
        kind, opts = tag & 0x0F, tag >> 4
        self.depth += 1
        if self.depth > 200:
            raise MtefError("入れ子が深すぎます")
        try:
            return self._obj(kind, opts)
        finally:
            self.depth -= 1

    def _obj(self, kind: int, opts: int) -> str | None:
        if kind == 0:
            return None
        if kind == 1:  # LINE
            self.nudge(opts)
            if opts & 0x4:
                self.pos += 1
            if opts & 0x2:
                self.ruler()
            if opts & 0x1:
                return ""
            return "".join(self.objects())
        if kind == 2:  # CHAR
            self.nudge(opts)
            face = self.u8() - 128
            code = struct.unpack_from("<H", self.d, self.pos)[0]
            self.pos += 2
            if opts & 0x2:
                self.objects()  # 飾り（EMBELL）の並び
            return _char(face, code)
        if kind == 3:  # TMPL
            self.nudge(opts)
            sel, var, _opt = self.u8(), self.u8(), self.u8()
            parts = self.objects()
            return _template(sel, var, parts)
        if kind == 4:  # PILE
            self.nudge(opts)
            self.pos += 2
            if opts & 0x2:
                self.ruler()
            return " ／ ".join(p for p in self.objects() if p)
        if kind == 5:  # MATRIX
            self.nudge(opts)
            _va, _hj, _vj, rows, cols = (self.u8() for _ in range(5))
            self.pos += (2 * (rows + 1) + 7) // 8 + (2 * (cols + 1) + 7) // 8
            cells = self.objects()
            lines = [" ".join(cells[i:i + cols]) for i in range(0, len(cells), max(cols, 1))]
            return "[" + "; ".join(lines) + "]"
        if kind == 6:  # EMBELL
            self.nudge(opts)
            self.pos += 1
            return ""
        if kind == 7:  # RULER
            self.ruler(read_tag=False)
            return ""
        if kind == 8:  # FONT
            self.pos += 2
            while self.d[self.pos] != 0:
                self.pos += 1
            self.pos += 1
            return ""
        if kind == 9:  # SIZE
            b = self.u8()
            if b == 101:
                self.pos += 2
            elif b == 100:
                self.pos += 3
            else:
                self.pos += 1
            return ""
        if 10 <= kind <= 14:
            return ""
        raise MtefError(f"知らない種類のレコード（{kind}）")

    def ruler(self, read_tag: bool = True) -> None:
        if read_tag:
            self.pos += 1  # RULER の印
        n = self.u8()
        self.pos += 3 * n


def _char(face: int, code: int) -> str:
    if face in _SYMBOL_FACES and code < 0x100:
        # 空白は空白のまま。表にない番号は、私用の領域の字（U+F0xx）
        return symfont.decode("Symbol", code) if code > 0x20 else " " if code == 0x20 else ""
    if face == _MT_EXTRA and 0x20 < code < 0x100:
        return f"〔MT Extra の字 0x{code:02X}〕"
    if code in (0, 0xFFFF):
        return ""
    if 0xF020 <= code <= 0xF0FF:  # 記号のフォントの私用の領域
        return symfont.decode("Symbol", code)
    return chr(code)


def _template(sel: int, var: int, parts: list[str]) -> str:
    slots = [p for p in parts]
    if sel == 14 and len(slots) >= 2:  # 分数
        return f"({slots[0]})/({slots[1]})"
    if sel == 13 and slots:  # 根号（2つ目の枠は、何乗根か）
        idx = slots[1] if len(slots) > 1 and slots[1] else ""
        return (f"[{idx}]" if idx else "") + f"√({slots[0]})"
    if sel == 15:  # 添え字：変種 0 下・1 上・2 両方（枠は下、上の順）
        if var == 0 and slots:
            return f"_({slots[0]})"
        if var == 1 and slots:
            return f"^({slots[0]})"
        if len(slots) >= 2:
            return f"_({slots[0]})^({slots[1]})"
    return "".join(slots)
