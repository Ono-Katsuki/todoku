"""埋め込んだ Excel の表の絵（EMF）の中で、各セルがどこにあるかを求め、書いたセルの字を Excel と同じ規則で描き直す。

Excel が描いた表の絵には、セルの位置そのものは書かれていない。そこで、
1. ブックから、絵に映る範囲（OleObjectSize・oleSize、映るのは選ばれているシート）の列の幅（文字数）・行の高さ（ポイント）を取り、
   セルの境の位置を「既知の幅の和 S」と「既定の幅の列の数 n」で表す（既定の幅は、絵の上での幅を未知数にする）。
2. 絵の中の証拠を集める：罫線（細い塗りの矩形と線）、結合セルの切り抜きの矩形、すでに描かれている字（どのセルの字かは、
   字が一致するセルが1つだけのものから決め、位置の合うものを足していく）。
3. 証拠ごとに一次式を立て、最小二乗で解く。
   横：x = ax + bx·S + v·n（境）、左寄せの字の左端 = 境 + m、右寄せの字の右端 = 境 − m − p·h（p は書式の「_)」などの空白の幅を
   字の高さ h あたりで表したもの）、中央寄せの字の中心 = セルの中心。
   縦：y = ay + by·（行の高さの和）、中央寄せの字の上端 + 字の高さ/2 = 行の中心 + dc、上寄せ・下寄せも同様。
   外れた証拠（字の飾りの線など）は、残差の大きいものから外して解き直す。
4. 新しい字を描くのに要る値（右寄せなら「境 − m − p·h」など）が、証拠から決まるかを、連立式の行の空間に入るかで確かめる。
   決まらなければ、何が足りないかを日本語で示して書かない（推測で描かない）。境の位置そのものは決まらなくてよい：
   目盛り線のない印刷用の絵で、字が右寄せの数字だけなら「右の境 − 余白」は決まるので、右寄せの数字は描ける。
   そのときセルに入るかは、使える幅の下限（セルの余白は左右で同じ・目盛り線の幅は 0 以上・Excel が字を描く切り抜きの
   左端は範囲の最初の列の中の左端、から求めたもの）で確かめ、切り抜きは Excel がセルの字を描いたときのものを使う。

字の大きさは、すでにある字の字体の高さと、そのセルの字の大きさ（ポイント）の比で決める。「縮小して全体を表示する」セルは、
Excel と同じく h' = ⌊h × 使える幅 ÷（字の幅＋空白の幅）⌋ にする（使える幅 = セルの幅 − 余白×2。函館市の記入例の絵で、
Excel が描いた 76・76・75 と一致することを確かめた）。字の幅は、絵の中の同じ字体の字の送り幅から求め、分からない字があって
収まるかが決められないときは書かない。数字は、同じ字体で2つ以上の数字の幅が分かり、どれも同じときに限り、ほかの数字も同じ幅とみなす。

縦の位置は「行の中心 − 字の高さ/2 + dc」を四捨五入（Excel の絵の字の上端と、どれも1単位以内で合う）。横は、右寄せ・中央寄せを
GDI の寄せ方（TA_RIGHT・TA_CENTER）で描くので、字の幅が分からなくても右端・中心は正しい。
"""

from __future__ import annotations

import bisect
import math
import re
import struct
from dataclasses import dataclass, field

from .formula.numfmt import format_value, split_sections
from .formula.parser import addr
from .formula.values import ExcelError
from .metafile import (TA_CENTER, TA_LEFT, TA_RIGHT, Emf, MetafileError, TextRun, bk_transparent,
                       colorref, create_font, delete, ext_text_out, intersect_clip, restoredc, savedc, select,
                       text_align, text_color)


class PictureError(ValueError):
    pass


# ====================================================================== ブックの書式（配置・字体）
@dataclass
class CellFont:
    face: str
    size: float  # ポイント
    bold: bool = False
    italic: bool = False
    underline: int = 0  # 0 なし・1 一重・ほか（二重・会計）
    strike: bool = False
    script: int = 0  # 0 ふつう・1 上付き・2 下付き
    charset: int | None = None

    @property
    def key(self) -> tuple:
        return (self.face.lower(), self.bold, self.italic)


@dataclass
class CellFormat:
    halign: str = "general"  # general / left / center / right / fill / justify / centerAcross / distributed
    valign: str = "bottom"  # top / center / bottom / justify / distributed
    wrap: bool = False
    shrink: bool = False
    indent: int = 0
    rotation: int = 0
    font: CellFont | None = None


_HALIGN_JA = {"left": "左寄せ", "right": "右寄せ", "center": "中央寄せ"}
_HALIGN = {0: "general", 1: "left", 2: "center", 3: "right", 4: "fill", 5: "justify", 6: "centerAcross", 7: "distributed"}
_VALIGN = {0: "top", 1: "center", 2: "bottom", 3: "justify", 4: "distributed"}


def _xls_formats(wb) -> tuple[list[CellFormat], dict]:
    """古い Excel（BIFF8）：FONT と XF のレコードから、配置と字体を読む（todoku.xls.book が読まない所）。"""
    from .xls.book import records, xlstring

    fonts: list[CellFont] = []
    out: list[CellFormat] = []
    info: dict = {}
    for _pos, t, b in records(wb.data):
        if t == 0x000A:
            break
        if t == 0x0031 and len(b) >= 16:  # FONT
            h, fl, _icv, bls, sss, uls, _fam, cs = struct.unpack_from("<HHHHHBBB", b, 0)
            name = xlstring(b, 14, cch_size=1)[0]
            fonts.append(CellFont(name, h / 20, bls >= 600, bool(fl & 0x02), uls, bool(fl & 0x08), sss, cs))
        elif t == 0x00E0 and len(b) >= 20:  # XF
            ifnt = struct.unpack_from("<H", b, 0)[0]
            b6, trot, b8 = b[6], b[7], b[8]
            fi = ifnt - 1 if ifnt >= 4 else ifnt  # FONT の4番は欠番
            out.append(CellFormat(_HALIGN.get(b6 & 7, "general"), _VALIGN.get((b6 >> 4) & 7, "bottom"), bool(b6 & 8),
                                  bool(b8 & 0x10), b8 & 0x0F, trot, fonts[fi] if 0 <= fi < len(fonts) else None))
        elif t == 0x00DE and len(b) >= 8:  # OleObjectSize：絵に映る範囲
            r1, r2, c1, c2 = struct.unpack_from("<HHBB", b, 2)
            info["range"] = (r1 + 1, c1 + 1, r2 + 1, c2 + 1)
        elif t == 0x003D and len(b) >= 12:  # Window1：選ばれているシート
            info["active"] = struct.unpack_from("<H", b, 10)[0]
    return out, info


def _xlsx_formats(wb) -> tuple[list[CellFormat], dict]:
    from .xlsx.workbook import M

    info: dict = {}
    root = wb.xml(wb.wb_path)
    ole = root.find(M + "oleSize")
    if ole is not None and ole.get("ref"):
        info["range"] = _ref_range(ole.get("ref"))
    bv = root.find(M + "bookViews")
    view = bv.find(M + "workbookView") if bv is not None else None
    if view is not None:
        info["active"] = int(view.get("activeTab", "0") or 0)
    path = next((t for typ, t in wb.rels(wb.wb_path).values() if typ == "styles"), None)
    if path is None or not wb.zip.has(path):
        return [], info
    st = wb.xml(path)

    def on(el, name) -> bool:
        x = el.find(M + name)
        return x is not None and x.get("val", "1") not in ("0", "false")

    fonts = []
    for f in st.find(M + "fonts") or []:
        sz = f.find(M + "sz")
        nm = f.find(M + "name")
        u = f.find(M + "u")
        va = f.find(M + "vertAlign")
        cs = f.find(M + "charset")
        fonts.append(CellFont(nm.get("val", "") if nm is not None else "", float(sz.get("val", "11")) if sz is not None else 11.0,
                              on(f, "b"), on(f, "i"), 0 if u is None else (1 if u.get("val", "single") == "single" else 2),
                              on(f, "strike"), {"superscript": 1, "subscript": 2}.get(va.get("val", "")) if va is not None else 0,
                              int(cs.get("val")) if cs is not None and cs.get("val", "").isdigit() else None))
    out = []
    for xf in st.find(M + "cellXfs") or []:
        al = xf.find(M + "alignment")
        fo = int(xf.get("fontId", "0") or 0)
        cf = CellFormat(font=fonts[fo] if fo < len(fonts) else None)
        if al is not None:
            cf.halign = al.get("horizontal", "general")
            if cf.halign == "centerContinuous":
                cf.halign = "centerAcross"
            cf.valign = al.get("vertical", "bottom")
            cf.wrap = al.get("wrapText") in ("1", "true")
            cf.shrink = al.get("shrinkToFit") in ("1", "true")
            cf.indent = int(al.get("indent", "0") or 0)
            cf.rotation = int(al.get("textRotation", "0") or 0)
        out.append(cf)
    return out, info


def _ref_range(ref: str) -> tuple[int, int, int, int]:
    from .formula.parser import parse_cell

    a, _, b = ref.replace("$", "").partition(":")
    r1, c1 = parse_cell(a)
    r2, c2 = parse_cell(b or a)
    return (r1, c1, r2, c2)


# ====================================================================== セルに見える字
@dataclass
class CellText:
    """Excel がセルに描く字と、描き方。text が空なら何も描かない。problem があれば、この字は描けない。"""

    text: str = ""
    halign: str = "left"  # left / center / right（General を値の種類で決めたもの）
    valign: str = "bottom"
    pad_left: str = ""  # 表示形式の「_x」の空白（x の幅）
    pad_right: str = ""
    color: str | None = None  # "RRGGBB"（None は自動＝黒）
    font: CellFont | None = None
    shrink: bool = False
    wrap: bool = False
    parts: tuple[str, str] | None = None  # 「* 」で埋める表示形式：左に寄せる字と、右に寄せる字
    problem: str = ""

    @property
    def lines(self) -> list[str]:
        if self.parts is not None:
            return [p for p in self.parts if p]
        return [ln for ln in self.text.split("\n") if ln]


_COLORS = {"red": "FF0000", "black": "000000", "blue": "0000FF", "green": "00FF00", "white": "FFFFFF", "yellow": "FFFF00",
           "magenta": "FF00FF", "cyan": "00FFFF"}


def _mark(fmt: str) -> tuple[str, dict[str, str], dict[str, str]]:
    """表示形式の「_x」（x の幅の空白）と「*x」（x で埋める）を、私用の字の印に置き換える（印 → x）。"""
    out, pads, fills = [], {}, {}
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            j = fmt.find('"', i + 1)
            j = len(fmt) - 1 if j < 0 else j
            out.append(fmt[i:j + 1])
            i = j + 1
        elif ch == "\\" and i + 1 < len(fmt):
            out.append(fmt[i:i + 2])
            i += 2
        elif ch == "[":
            j = fmt.find("]", i)
            j = len(fmt) - 1 if j < 0 else j
            out.append(fmt[i:j + 1])
            i = j + 1
        elif ch in "_*" and i + 1 < len(fmt):
            table = pads if ch == "_" else fills
            base = 0xE000 if ch == "_" else 0xE800
            mk = next((m for m, x in table.items() if x == fmt[i + 1]), None) or chr(base + len(table))
            table[mk] = fmt[i + 1]
            out.append('"' + mk + '"')
            i += 2
        else:
            out.append(ch)
            i += 1
    return "".join(out), pads, fills


def _section(fmt: str, v) -> str:
    """値に使う表示形式の節（format_value と同じ選び方）。"""
    secs = split_sections(fmt)
    if isinstance(v, str):
        return secs[3] if len(secs) >= 4 else (secs[0] if len(secs) == 1 and "@" in secs[0] else "")
    if len(secs) == 1:
        return secs[0]
    x = float(v)
    if x > 0 or (x == 0 and len(secs) < 3):
        return secs[0]
    if x < 0:
        return secs[1]
    return secs[2] if len(secs) > 2 else secs[0]


def cell_text(wb, sh, fmts: list[CellFormat], r: int, c: int) -> CellText:
    """セル (r, c) に Excel が描く字（結合セルの左上でないセルは、何も描かない）。"""
    cell = sh.cells.get((r, c))
    v = cell.value if cell is not None else None
    xf = sh.style_of(r, c)
    cf = fmts[xf] if 0 <= xf < len(fmts) else CellFormat()
    st = wb.style(xf)
    out = CellText(valign=cf.valign, color=st.font_color, font=cf.font, shrink=cf.shrink, wrap=cf.wrap)
    if v is None or v == "":
        return out
    kind = "text" if isinstance(v, str) else ("center" if isinstance(v, (bool, ExcelError)) else "number")
    if isinstance(v, ExcelError):
        out.text = v.code
    elif isinstance(v, bool):
        out.text = "TRUE" if v else "FALSE"
    else:
        fmt = st.num_fmt or "General"
        marked, pads, fills = _mark(fmt)
        try:
            s = format_value(v, marked, getattr(wb, "date1904", False))
        except Exception as e:  # noqa: BLE001
            out.problem = f"表示形式「{fmt}」で、Excel と同じ字にできません（{e}）"
            return out
        sec = _section(fmt, v)
        m = re.search(r"\[(red|black|blue|green|white|yellow|magenta|cyan|color\s*\d+)\]", sec, re.I)
        if m:
            name = m[1].lower()
            if name not in _COLORS:
                out.problem = f"表示形式の色の番号（{m[0]}）は、まだ描けません"
                return out
            out.color = _COLORS[name]
        fill_at = [i for i, ch in enumerate(s) if ch in fills]
        if fill_at:
            if any(fills[s[i]] != " " for i in fill_at):
                out.problem = "空白以外の字でセルを埋める表示形式（*x）は、まだ描けません"
                return out
            left, right = s[:fill_at[0]], s[fill_at[0] + 1:]
            right = "".join(ch for ch in right if ch not in fills)
        else:
            left, right = "", s
        lead = re.match(f"[{''.join(pads)}]*", left or right)[0] if pads else ""
        trail = re.search(f"[{''.join(pads)}]*$", right)[0] if pads else ""
        out.pad_left = "".join(pads[ch] for ch in lead)
        out.pad_right = "".join(pads[ch] for ch in trail)
        core_l = (left[len(lead):] if left else "")
        core_r = right[len(lead) if not left else 0:len(right) - len(trail)]
        if any(ch in pads for ch in core_l + core_r):
            out.problem = "表示形式の字の間にある空白（_x）は、まだ描けません"
            return out
        if fill_at and core_l.strip():
            out.parts = (core_l.strip(), core_r.strip())
            out.text = (core_l + " " + core_r).strip()
        elif fill_at:
            # 左が空白だけの「* 」（会計の表示形式など）：字はセルの右に寄る
            out.text = core_r
            out.pad_left = ""
            kind = "fill-right"
        else:
            out.text = core_r
    h = cf.halign
    if kind == "fill-right":
        h = "right"
    elif h == "general":
        h = {"text": "left", "number": "right", "center": "center"}[kind]
    elif h == "justify":
        h = "left"  # 1行に収まる両端揃えは左寄せと同じ
    out.halign = h
    if h not in ("left", "center", "right"):
        out.problem = f"配置「{cf.halign}」のセルの字は、まだ描けません"
    elif cf.rotation:
        out.problem = "字の向きを変えたセルの字は、まだ描けません"
    elif cf.indent:
        out.problem = "字下げ（インデント）のあるセルの字は、まだ描けません"
    elif cf.valign not in ("top", "center", "bottom"):
        out.problem = f"縦の配置「{cf.valign}」のセルの字は、まだ描けません"
    elif cf.font is None:
        out.problem = "セルの字体が分かりません"
    elif cf.font.underline not in (0, 1) or cf.font.script:
        out.problem = "二重下線・会計用の下線・上付き・下付きの字は、まだ描けません"
    elif "\n" in out.text:
        out.problem = "改行のある字は、まだ描けません"
    return out


# ====================================================================== 絵に映るシート
def shown_sheet(wb):
    """表の絵に映るシート（ブックで選ばれているシート：BIFF の Window1.itabCur、.xlsx の workbookView の activeTab）。
    範囲の外を指していれば None。"""
    kind = getattr(wb, "kind", "xlsx")
    _f, info = _xls_formats(wb) if kind == "xls" else _xlsx_formats(wb)
    sheets = list(wb.sheets)
    ai = info.get("active", 0)
    return sheets[ai] if 0 <= ai < len(sheets) else None


class _RowTops:
    """行の境：範囲の上から行 r の上の境までの高さ（ポイント）。高さの指定のある行と非表示の行だけの控えで求める
    （遠くの行のセルがあっても、あいだの行を 1 つずつ足さない）。"""

    def __init__(self, sh, r1: int, r2: int):
        self.r1 = r1
        self.base = float(sh.default_row_height)
        self.rows = sorted(r for r in {*sh.row_heights, *sh.hidden_rows} if r1 <= r <= r2)
        self.total = [0.0]
        for r in self.rows:
            h = 0.0 if r in sh.hidden_rows else float(sh.row_heights.get(r, sh.default_row_height))
            self.total.append(self.total[-1] + h - self.base)

    def __getitem__(self, r: int) -> float:
        return (r - self.r1) * self.base + self.total[bisect.bisect_left(self.rows, r)]


class SheetModel:
    """絵に映るシートと範囲、列・行の境（既知の幅の和と、既定の幅の列の数）、セルの字。"""

    def __init__(self, wb):
        self.wb = wb
        kind = getattr(wb, "kind", "xlsx")
        self.formats, info = _xls_formats(wb) if kind == "xls" else _xlsx_formats(wb)
        sheets = list(wb.sheets)
        ai = info.get("active", 0)
        sh = sheets[ai] if 0 <= ai < len(sheets) else None
        if sh is None or sh.state.startswith("non-"):
            raise PictureError("表の絵に映るシート（選ばれているシート）が、セルのあるシートではありません")
        self.sheet = sh
        self.range_recorded = "range" in info
        if "range" in info:
            self.range = info["range"]
        else:
            # 範囲の記録がない：使われているセルの範囲（A1 から）
            self.range = (1, 1, max(sh.max_row, max((r for r, _c in sh.cells), default=1)),
                          max(sh.max_col, max((c for _r, c in sh.cells), default=1)))
        r1, c1, r2, c2 = self.range
        self.cols = list(range(c1, c2 + 1))
        # 列の境：k 番目の境 = 範囲の左から k 列の右（既知の幅の和, 既定の幅の列の数）
        self.xb: dict[int, tuple[float, int]] = {}
        s, n = 0.0, 0
        for c in self.cols:
            self.xb[c] = (s, n)
            if c in sh.hidden_cols:
                continue
            w = sh.col_widths.get(c)
            if w is None:
                n += 1
            else:
                s += w
        self.xb[c2 + 1] = (s, n)
        self.yb = _RowTops(sh, r1, r2)

    def visible(self, r: int, c: int) -> bool:
        r1, c1, r2, c2 = self.range
        return r1 <= r <= r2 and c1 <= c <= c2 and r not in self.sheet.hidden_rows and c not in self.sheet.hidden_cols

    def area(self, r: int, c: int) -> tuple[int, int, int, int]:
        """セルが描かれる範囲（結合セルなら結合の範囲。範囲の外にはみ出す所は切る）。"""
        m = self.sheet.merged_at(r, c)
        r1, c1, r2, c2 = m if m else (r, c, r, c)
        R1, C1, R2, C2 = self.range
        return max(r1, R1), max(c1, C1), min(r2, R2), min(c2, C2)

    def display_cells(self):
        """字を描けるセル（結合セルの左上だけ）。ファイルに要素のあるマスだけを見る（要素のないマスに字はない。範囲の広さによらない）。"""
        for r, c in self.sheet.existing_in(*self.range):
            m = self.sheet.merged_at(r, c)
            if m and (r, c) != (m[0], m[1]):
                continue
            if self.visible(r, c):
                yield r, c

    def text(self, r: int, c: int) -> CellText:
        return cell_text(self.wb, self.sheet, self.formats, r, c)


# ====================================================================== 最小二乗（小さな連立式）
def _minus(a: dict[str, float], b: dict[str, float]) -> dict[str, float]:
    return {k: a.get(k, 0.0) - b.get(k, 0.0) for k in set(a) | set(b)}


class _System:
    """一次式の証拠を集めて解く。解けない（決まらない）組み合わせは estimable で分かる。"""

    def __init__(self, names: list[str]):
        self.names = list(names)
        self.rows: list[tuple[list[float], float, object]] = []

    def var(self, name: str) -> int:
        if name not in self.names:
            self.names.append(name)
            for coef, _o, _t in self.rows:
                coef.append(0.0)
        return self.names.index(name)

    def add(self, coef: dict[str, float], obs: float, tag=None) -> None:
        for k in coef:
            self.var(k)
        row = [0.0] * len(self.names)
        for k, v in coef.items():
            row[self.names.index(k)] = float(v)
        self.rows.append((row, float(obs), tag))

    def vec(self, coef: dict[str, float]) -> list[float] | None:
        if any(k not in self.names for k in coef if coef[k]):
            return None
        v = [0.0] * len(self.names)
        for k, x in coef.items():
            if x:
                v[self.names.index(k)] = float(x)
        return v

    def solve(self, tol_of, max_drop: float = 0.25) -> None:
        """最小二乗で解き、残差が tol_of(tag) を超える証拠を大きいものから外して解き直す。"""
        active = list(range(len(self.rows)))
        dropped = 0
        while True:
            self._fit(active)
            worst, wi = 0.0, None
            for i in active:
                row, obs, tag = self.rows[i]
                res = abs(sum(a * x for a, x in zip(row, self.x)) - obs) / max(tol_of(tag), 1e-9)
                if res > worst:
                    worst, wi = res, i
            if worst <= 1.0 or wi is None:
                break
            if dropped + 1 > max_drop * len(self.rows):
                raise PictureError("表の絵の中の罫線と字の位置が、ブックの列の幅・行の高さと合いません")
            active.remove(wi)
            dropped += 1
        self.active = active

    def _fit(self, active: list[int]) -> None:
        n = len(self.names)
        scale = [max(math.sqrt(sum(self.rows[i][0][j] ** 2 for i in active)), 1e-12) for j in range(n)]
        A = [[self.rows[i][0][j] / scale[j] for j in range(n)] for i in active]
        y = [self.rows[i][1] for i in active]
        N = [[sum(A[k][i] * A[k][j] for k in range(len(A))) for j in range(n)] for i in range(n)]
        b = [sum(A[k][i] * y[k] for k in range(len(A))) for i in range(n)]
        # 消去（主元を選ぶ）。小さすぎる主元の変数は 0 にする（決まらない変数）
        M = [N[i][:] + [b[i]] for i in range(n)]
        pivots: list[tuple[int, int]] = []  # (行, 列)：Gauss-Jordan で消去し終えた主元
        used_r, used_c = set(), set()
        for _ in range(n):
            best, bi, bj = 1e-9, None, None
            for i in range(n):
                if i in used_r:
                    continue
                for j in range(n):
                    if j not in used_c and abs(M[i][j]) > best:
                        best, bi, bj = abs(M[i][j]), i, j
            if bi is None:
                break
            used_r.add(bi)
            used_c.add(bj)
            row = M[bi]
            for i in range(n):
                if i != bi and M[i][bj]:
                    f = M[i][bj] / row[bj]
                    M[i] = [a - f * c for a, c in zip(M[i], row)]
            pivots.append((bi, bj))
        x = [0.0] * n
        for bi, bj in pivots:
            x[bj] = M[bi][n] / M[bi][bj]
        self.x = [x[j] / scale[j] for j in range(n)]
        self._basis = []  # 行の空間の正規直交基底（決まる組み合わせの判定）
        for i in active:
            v = self.rows[i][0][:]
            for q in self._basis:
                d = sum(a * c for a, c in zip(v, q))
                v = [a - d * c for a, c in zip(v, q)]
            nv = math.sqrt(sum(a * a for a in v))
            if nv > 1e-7 * max(1.0, math.sqrt(sum(a * a for a in self.rows[i][0]))):
                self._basis.append([a / nv for a in v])

    def estimable(self, coef: dict[str, float]) -> bool:
        v = self.vec(coef)
        if v is None:
            return False
        norm = math.sqrt(sum(a * a for a in v)) or 1.0
        for q in self._basis:
            d = sum(a * c for a, c in zip(v, q))
            v = [a - d * c for a, c in zip(v, q)]
        return math.sqrt(sum(a * a for a in v)) <= 1e-6 * norm

    def value(self, coef: dict[str, float]) -> float:
        v = self.vec(coef)
        return sum(a * x for a, x in zip(v, self.x))


# ====================================================================== 絵の中のセルの位置
@dataclass
class Placed:
    """描く字1つ：レコードの塊と、描いた所（証明で照らす）。"""

    cell: tuple[int, int]
    text: str
    records: list[bytes]
    ref: tuple[int, int]
    band: tuple[float, float] = (0.0, 0.0)  # 字の上端と下端


@dataclass
class Layout:
    emf: Emf
    model: SheetModel
    anchors: dict[int, tuple[int, int]] = field(default_factory=dict)  # 字のレコードの番号 → セル

    def __post_init__(self):
        emf, model = self.emf, self.model
        runs = [r for r in emf.runs if r.font is not None and r.text.strip()]
        if not runs:
            raise PictureError("表の絵に、セルの字が1つも描かれていません（位置を確かめる手がかりがありません）")
        forms: dict[tuple, int] = {}
        for r in runs:
            forms[r.xform] = forms.get(r.xform, 0) + 1
        self.xform = max(forms, key=lambda k: forms[k])
        self.runs = [r for r in runs if r.xform == self.xform]
        self.texts = {rc: model.text(*rc) for rc in model.display_cells()}
        self._widths()
        self._solve()

    # ------------------------------------------------------------------ 字の幅
    def _widths(self) -> None:
        """字体ごとの、字の送り幅（字の高さ h のときの幅そのもの、と h あたりの幅）。"""
        self.exact: dict[tuple, dict[str, int]] = {}  # (字体の印, h) → 字 → 幅
        em: dict[tuple, dict[str, list[float]]] = {}
        self.heights: dict[tuple, list[float]] = {}  # 字体の印 → 字の高さ / |h|
        for r in self.runs:
            f = r.font
            h = abs(f.height)
            if not h:
                continue
            if r.dx and len(r.dx) == len(r.text):
                ex = self.exact.setdefault((f.key, h), {})
                for ch, w in zip(r.text, r.dx):
                    ex.setdefault(ch, w)
                    em.setdefault(f.key, {}).setdefault(ch, []).append(w / h)
            bh = r.box_height
            if bh:
                self.heights.setdefault(f.key, []).append(bh / h)
        self.em = {k: {ch: sorted(v)[len(v) // 2] for ch, v in d.items()} for k, d in em.items()}
        for d in list(self.exact.values()) + list(self.em.values()):
            digits = {d[ch] for ch in "0123456789" if ch in d}
            if len([ch for ch in "0123456789" if ch in d]) >= 2 and len(digits) == 1:
                w = digits.pop()
                for ch in "0123456789":
                    d.setdefault(ch, w)

    def text_width(self, key: tuple, h: int, text: str) -> tuple[float, float, bool]:
        """(少なくとも, 多くとも, 正確か)。分からない字は 0.25〜1.05 字分とみなす。"""
        ex = self.exact.get((key, h), {})
        if all(ch in ex for ch in text):
            w = float(sum(ex[ch] for ch in text))
            return w, w, True
        emk = self.em.get(key, {})
        lo = hi = 0.0
        for ch in text:
            if ch in ex:
                lo += ex[ch]
                hi += ex[ch]
            elif ch in emk:
                lo += emk[ch] * h * 0.95
                hi += emk[ch] * h * 1.05
            else:
                lo += 0.25 * h
                hi += 1.05 * h
        return lo, hi, False

    def known_width(self, key: tuple, h: int, text: str) -> float | None:
        ex = self.exact.get((key, h), {})
        emk = self.em.get(key, {})
        if all(ch in ex for ch in text):
            return float(sum(ex[ch] for ch in text))
        if all(ch in ex or ch in emk for ch in text):
            return sum(ex[ch] if ch in ex else emk[ch] * h for ch in text)
        return None

    # ------------------------------------------------------------------ 証拠を集めて解く
    def _xb(self, c: int) -> dict[str, float]:
        s, n = self.model.xb[c]
        return {"ax": 1.0, "bx": s, "v": n}

    def _yb(self, r: int) -> dict[str, float]:
        return {"ay": 1.0, "by": self.model.yb[r]}

    def _pad(self, pad: str, key: tuple) -> str:
        return f"p:{pad}:{key}"

    def _solve(self) -> None:
        model, emf = self.model, self.emf
        r1, c1, r2, c2 = model.range
        bl, bt, br, bb = emf.bounds
        # 罫線との照合の初めの見当：表の範囲が絵の外接矩形いっぱいに描かれているとする
        s_tot, n_tot = model.xb[c2 + 1]
        w_def = 9.0
        bx0 = (br - bl) / max(s_tot + n_tot * w_def, 1e-6)
        by0 = (bb - bt) / max(model.yb[r2 + 1], 1e-6)
        guess = {"ax": bl, "bx": bx0, "v": bx0 * w_def, "ay": bt, "by": by0}
        anchors = self._unique_anchors()
        for _pass in range(5):
            xs = _System(["ax", "bx", "v", "m"])
            ys = _System(["ay", "by", "dc", "dt", "db"])
            tolx = max(2.0, 0.2 * guess["bx"])
            toly = max(2.0, 0.2 * guess["by"] * 12)
            self._k = self._scale_of(anchors)
            self._lines(xs, ys, guess, tolx, toly)
            for idx, rc in anchors.items():
                self._anchor_eqs(xs, ys, self._run(idx), rc)
            xs.solve(lambda tag: tolx)
            ys.solve(lambda tag: toly)
            self.xs, self.ys = xs, ys
            # 境の位置そのもの（表の左上の位置・列の幅・行の高さ）は、すべて決まらなくてもよい：右寄せの字の右端（境 − 余白）や、
            # 下寄せの字の下端（境 − ずれ）のように、描くのに要る組み合わせが証拠から決まるかを、描くときに一つずつ確かめる
            # （既定の幅の列しかなければ、幅の分かる列の倍率 bx は要らない）
            self.x_fixed = all(xs.estimable(self._xb(c)) for c in self.cols_b)
            # 行の境の式は (1, 高さ) の形：いちばん上と下で決まれば、あいだの行の境も、その 2 つの組み合わせなので決まる
            self.y_fixed = all(ys.estimable(self._yb(r)) for r in (self.rows_b[0], self.rows_b[-1]))
            guess = {k: (xs.value({k: 1}) if k in ("ax", "bx", "v") and xs.estimable({k: 1}) else
                         ys.value({k: 1}) if k in ("ay", "by") and ys.estimable({k: 1}) else guess[k]) for k in guess}
            self._guess = guess
            more = self._position_anchors(anchors)
            if not more:
                break
            anchors.update(more)
        self.anchors = anchors
        # 罫線のある境は、罫線の位置そのものを使う（列ごとの丸めの差を残さない）
        self.snap_x = self._snap("v", self.cols_b, self._xb, xs, tolx)
        self.snap_y = self._snap("h", self.rows_b, self._yb, ys, toly)
        self.tolx, self.toly = tolx, toly
        if not model.range_recorded and not all(self._inside(self._run(i), rc) for i, rc in anchors.items()):
            # 映る範囲の記録がなく、使われているセルの範囲と見なした：字が自分のセルの中に来ないなら、絵に映る範囲はもっと狭い
            # （見なした範囲に合わせて解くと、行の高さを取り違える。遠くのセルは絵に映らないので、ここには書けない）
            raise PictureError("表の絵の中の罫線と字の位置が、ブックの列の幅・行の高さと合いません"
                               "（映る範囲の記録がなく、使われているセルの範囲が絵に収まりません）")
        self._offsets()

    def _offsets(self) -> None:
        """境から字までのずれ（余白 m・書式の空白の幅・中央寄せのずれ・縦のずれ）を、境（罫線の位置）からじかに測る。
        どれも、すでに描かれている字の中央値。縮小して描いた字は使わない（余白も縮むので）。"""
        def med(v):
            v = sorted(v)
            return v[len(v) // 2] if v else None

        lefts, rights, centers, pads, gaps = [], [], [], {}, {}
        dcs, dts, dbs = [], [], []
        for idx, rc in self.anchors.items():
            run, t = self._run(idx), self.texts[rc]
            if t.problem or t.parts is not None:
                continue
            r1, c1, r2, c2 = self.model.area(*rc)
            L, R, T, B = self.bx(c1), self.bx(c2 + 1), self.by(r1), self.by(r2 + 1)
            h = abs(run.font.height)
            x0, x1, xc = self.edges(run)
            if not self._shrunk(run, t):
                if t.halign == "left" and not t.pad_left and x0 is not None:
                    lefts.append(x0 - L)
                elif t.halign == "right" and not t.pad_right and x1 is not None:
                    rights.append(R - x1)
                elif t.halign == "right" and x1 is not None:
                    gaps.setdefault((self._pad(t.pad_right, run.font.key), h), []).append(R - x1)
                elif t.halign == "center" and not (t.pad_left or t.pad_right) and xc is not None:
                    centers.append(xc - (L + R) / 2)
            bh = run.box_height
            if bh and not run.align & 0x18:
                if t.valign == "center":
                    dcs.append(run.y + bh / 2 - (T + B) / 2)
                elif t.valign == "top":
                    dts.append(run.y - T)
                elif t.valign == "bottom":
                    dbs.append(B - run.y - bh)
        self.m = med(lefts + rights)
        # 左の余白（左の境から左寄せの字まで）と右の余白（右寄せの字から右の境まで）は、目盛り線の幅の分だけ違うことがある
        # （境は線の左端：画面の絵では左 3・右 2）。測れた側はその値を、測れない側は両方の真ん中の値を使う
        self.ml = med(lefts) if lefts else self.m
        self.mr = med(rights) if rights else self.m
        self.gaps = {k: med(v) for k, v in gaps.items()}
        self.pad_em = {}
        if self.m is not None:
            for (name, h), g in self.gaps.items():
                self.pad_em.setdefault(name, []).append((g - self.mr) / h)
            self.pad_em = {k: med(v) for k, v in self.pad_em.items()}
        self.dh = med(centers)
        self.dc, self.dt, self.db = med(dcs), med(dts), med(dbs)
        # 表示形式の「_x」の空白は、ちょうど字 x の幅：測った空白の幅を、その字の幅として使う（「_)」なら「)」の幅）
        if self.m is not None:
            for (name, h), g in self.gaps.items():
                pad, key = name[2:].split(":", 1)
                if len(pad) == 1:
                    k_ = next((r.font.key for r in self.runs if str(r.font.key) == key), None)
                    if k_ is not None:
                        self.exact.setdefault((k_, h), {}).setdefault(pad, int(round(g - self.mr)))
                        self.em.setdefault(k_, {}).setdefault(pad, (g - self.mr) / h)
        # 対の括弧は同じ幅（字体の作りの決まり）：片方しか分からないときは、もう片方にも使う
        for d in list(self.exact.values()) + list(self.em.values()):
            for a, b in ("()", "[]", "{}", "（）", "「」", "［］"):
                if a in d and b not in d:
                    d[b] = d[a]
                elif b in d and a not in d:
                    d[a] = d[b]

    def _shrunk(self, run: TextRun, t: CellText) -> bool:
        return bool(t.shrink and (self._k is None or t.font is None or abs(run.font.height) < round(t.font.size * self._k)))

    @property
    def cols_b(self) -> list[int]:
        return sorted(self.model.xb)

    @property
    def rows_b(self) -> range:
        r1, _c1, r2, _c2 = self.model.range
        return range(r1, r2 + 2)

    def _run(self, idx: int) -> TextRun:
        return next(r for r in self.runs if r.index == idx)

    def _near(self, orient: str, want: float, tol: float) -> float | None:
        """want の近くの罫線の中心。塗りの矩形の線があれば、それだけを見る（同じ所に描く細い線は、矩形の端に描かれる）。
        同じ太さの線が、太さの2倍の間隔で並ぶものは二重線として、その真ん中を境とする。"""
        cand = [ln for ln in self.emf.lines if ln.orient == orient and abs(ln.at - want) <= 2.5 * tol]
        thick = [ln for ln in cand if ln.thick > 0] or cand
        spots = [float(ln.at) for ln in thick]
        for a in thick:
            for b in thick:
                if a.at < b.at and a.thick == b.thick and a.thick > 0 and \
                        abs((b.at - a.at) - 2 * a.thick) <= (1 if a.thick >= 4 else 0):
                    spots.append((a.at + b.at) / 2)
        spots = [x for x in spots if abs(x - want) <= tol]
        if not spots:
            return None
        return min(spots, key=lambda x: abs(x - want))

    def _tied(self, orient: str, keys, at_of) -> list:
        """絵の罫線ごとの、いちばん近いセルの境（行・列の番号の順）。罫線 1 本は、いちばん近い境の証拠にしかならない：
        見るのは絵の線の数だけで、範囲の広さ（遠くのセルまでの行の数）によらない。at_of(境)：いまの見当の位置。"""
        keys = list(keys) if not isinstance(keys, (list, range)) else keys
        rising = at_of(keys[0]) <= at_of(keys[-1])
        key = at_of if rising else (lambda k: -at_of(k))
        out = set()
        for at in {ln.at for ln in self.emf.lines if ln.orient == orient}:
            i = bisect.bisect_left(keys, at if rising else -at, key=key)
            near = [keys[j] for j in (i - 1, i) if 0 <= j < len(keys)]
            out.add(min(near, key=lambda k: abs(at_of(k) - at)))
        return sorted(out)

    def _lines(self, xs: _System, ys: _System, guess: dict, tolx: float, toly: float) -> None:
        def x_at(c: int) -> float:
            co = self._xb(c)
            return guess["ax"] + guess["bx"] * co["bx"] + guess["v"] * co["v"]

        def y_at(r: int) -> float:
            return guess["ay"] + guess["by"] * self._yb(r)["by"]

        for c in self._tied("v", self.cols_b, x_at):
            at = self._near("v", x_at(c), 2 * tolx)
            if at is not None:
                xs.add(self._xb(c), at, ("line", c))
        for r in self._tied("h", self.rows_b, y_at):
            at = self._near("h", y_at(r), 2 * toly)
            if at is not None:
                ys.add(self._yb(r), at, ("line", r))

    def _snap(self, orient: str, keys, coef_of, sysm: _System, tol: float) -> dict:
        """罫線のある境は、その罫線の位置そのものに：{境の、幅・高さの和: 罫線の位置}（幅のない行・列は、同じ境）。"""
        out = {}
        # 線に結びつけた境は、位置が決まらない向きでも、そこだけは決まることがあるので必ず見る
        tagged = {t[1] for _c, _o, t in sysm.rows if isinstance(t, tuple) and t[0] == "line"}
        for k in sorted({*self._tied(orient, keys, lambda k_: sysm.value(coef_of(k_))), *tagged}):
            co = coef_of(k)
            if not sysm.estimable(co):
                continue  # 位置の決まらない境は、近くの線に寄せない
            at = self._near(orient, sysm.value(co), tol)
            if at is not None:
                out[self.model.xb[k] if orient == "v" else self.model.yb[k]] = at
        return out

    def bx(self, c: int) -> float:
        """列 c の左の境の、絵の上の x。"""
        got = self.snap_x.get(self.model.xb[c])
        return self.xs.value(self._xb(c)) if got is None else got

    def by(self, r: int) -> float:
        got = self.snap_y.get(self.model.yb[r])
        return self.ys.value(self._yb(r)) if got is None else got

    def rect(self, r: int, c: int) -> tuple[float, float, float, float]:
        """セルの矩形（求めた解）。境の位置が決まらない向きでは、本当の位置からずれていることがある（x_range・y_fixed を見る）。"""
        r1, c1, r2, c2 = self.model.area(r, c)
        return self.bx(c1), self.by(r1), self.bx(c2 + 1), self.by(r2 + 1)

    def text_clip(self) -> tuple[int, int, int, int] | None:
        """Excel がセルの字を描いたときの切り抜き（表の範囲の、セルの中の部分）。字によって違えば None。"""
        clips = {self._run(i).clip for i in self.anchors}
        if len(clips) != 1:
            return None
        c = next(iter(clips))
        return c if isinstance(c, tuple) else None

    def x_shift(self) -> tuple[float, float]:
        """列の境が決まらないとき（位置の分かる字が右寄せだけで、罫線もないとき）、本当の境が求めた解から
        どれだけずれうるか（t の範囲。列の幅は決まっているので、どの境も同じだけずれる）。決まっていれば (0, 0)。
        - 上限：Excel が字を描くときの切り抜きの左端は、範囲の最初の列の中（左の境 ＋ 目盛り線の幅、幅は 0 以上）の左端
        - 下限：右寄せの字の右端は、列の右の境より右に出ない（右の余白は 0 以上）"""
        if self.x_fixed:
            return 0.0, 0.0
        c0 = self.cols_b[0]
        if not all(self.xs.estimable(_minus(self._xb(c), self._xb(c0))) for c in self.cols_b):
            raise PictureError("表の絵の中で、列の境の位置が決まりません（罫線も、位置の分かる字も足りません）")
        clip = self.text_clip()
        if clip is None:
            raise PictureError("表の絵の中で、列の境の位置が決まりません（罫線がなく、字を描いた切り抜きも字ごとに違います）")
        hi = clip[0] - self.bx(self.cols_b[0])
        lows = []
        for idx, rc in self.anchors.items():
            t = self.texts[rc]
            run = self._run(idx)
            x1 = self.edges(run)[1]
            if t.halign == "right" and not t.pad_right and x1 is not None and not self._shrunk(run, t):
                lows.append(x1 - self.bx(self.model.area(*rc)[3] + 1))
        if not lows:
            raise PictureError("表の絵の中で、列の境の位置が決まりません（罫線も、右寄せの字もありません）")
        lo = max(lows)
        if lo > hi + 1:
            raise PictureError("表の絵の中の字の位置と、字を描いた切り抜きが合いません")
        return lo, max(lo, hi)

    def x_range(self, r: int, c: int) -> tuple[float, float, float, float]:
        """セルの左右の境のありうる範囲：(左の最小, 左の最大, 右の最小, 右の最大)。境が決まっていれば幅のない範囲。"""
        L, _T, R, _B = self.rect(r, c)
        lo, hi = self.x_shift()
        return L + lo, L + hi, R + lo, R + hi

    # ------------------------------------------------------------------ 字とセルを結ぶ
    def _candidates(self) -> dict[str, list[tuple[int, int]]]:
        out: dict[str, list[tuple[int, int]]] = {}
        for rc, t in self.texts.items():
            for ln in set(t.lines):
                out.setdefault(ln, []).append(rc)
        return out

    def _unique_anchors(self) -> dict[int, tuple[int, int]]:
        cand = self._candidates()
        count: dict[str, int] = {}
        for r in self.runs:
            count[r.text] = count.get(r.text, 0) + 1
        out = {}
        for r in self.runs:
            cells = cand.get(r.text, [])
            if len(cells) == 1 and count[r.text] == 1 and len(self.texts[cells[0]].lines) == 1:
                out[r.index] = cells[0]
        return out

    def _position_anchors(self, have: dict[int, tuple[int, int]]) -> dict[int, tuple[int, int]]:
        """字が一致するセルのうち、字の位置がそのセルの中にあるものが1つだけなら結ぶ。
        境の位置が決まらない向きでは、そのセルの寄せ方で字が描かれるはずの所（右寄せなら「右の境 − 余白」など、
        証拠から決まる組み合わせ）に字があるかで確かめる。決まらなければ結ばない。"""
        cand = self._candidates()
        taken = set(have.values())
        out = {}
        for r in self.runs:
            if r.index in have:
                continue
            hits = [rc for rc in cand.get(r.text, []) if rc not in taken and self._inside(r, rc)]
            if len(hits) == 1 and len(self.texts[hits[0]].lines) == 1:
                out[r.index] = hits[0]
                taken.add(hits[0])
        return out

    def _inside(self, run: TextRun, rc: tuple[int, int]) -> bool:
        """字 run が、セル rc の字として描かれた所にあるか。"""
        L, T, R, B = self._rect_fit(*rc)
        r1, c1, r2, c2 = self.model.area(*rc)
        x0, x1, xc = self.edges(run)
        bh = run.box_height
        t = self.texts[rc]
        if self.x_fixed:
            mx = xc if xc is not None else (x0 if x0 is not None else x1)  # 字の中の1点（寄せ方で決まる端もセルの中）
            if not L - self._tolx() <= mx <= R + self._tolx():
                return False
        else:
            if t.pad_left or t.pad_right or self._shrunk(run, t):
                return False
            if t.halign == "right" and x1 is not None:
                co, at = {**self._xb(c2 + 1), "m": -1.0}, x1
            elif t.halign == "left" and x0 is not None:
                co, at = {**self._xb(c1), "m": 1.0}, x0
            else:
                return False
            if not self.xs.estimable(co) or abs(self.xs.value(co) - at) > self._tolx():
                return False
        if self.y_fixed:
            h = bh or abs(run.font.height)
            return T - self._toly() <= run.y + h / 2 <= B + self._toly()
        if not bh or run.align & 0x18:
            return False
        Tc, Bc = self._yb(r1), self._yb(r2 + 1)
        if t.valign == "center":
            co, at = {**{k: (Tc.get(k, 0.0) + Bc.get(k, 0.0)) / 2 for k in set(Tc) | set(Bc)}, "dc": 1.0}, run.y + bh / 2
        elif t.valign == "top":
            co, at = {**Tc, "dt": 1.0}, run.y
        else:
            co, at = {**Bc, "db": -1.0}, run.y + bh
        return self.ys.estimable(co) and abs(self.ys.value(co) - at) <= self._toly()

    def _tolx(self) -> float:
        # 文字1つ分の幅の 0.2 倍（幅の分かる列がなく、倍率が決まらないときは、見当の値）
        return max(2.0, 0.2 * self._guess["bx"])

    def _toly(self) -> float:
        return max(2.0, 0.2 * self._guess["by"] * 12)

    def _scale_of(self, anchors: dict) -> float | None:
        """絵の上の字の高さ ÷ ポイント：縮小していないセルの字の、真ん中の値。"""
        ks = sorted(abs(self._run(i).font.height) / self.texts[rc].font.size for i, rc in anchors.items()
                    if self.texts[rc].font is not None and not self.texts[rc].shrink and self.texts[rc].font.size
                    and self.texts[rc].font.face.lower() == self._run(i).font.face.lower())
        return ks[len(ks) // 2] if ks else None

    def _rect_fit(self, r: int, c: int):
        r1, c1, r2, c2 = self.model.area(r, c)
        return (self.xs.value(self._xb(c1)), self.ys.value(self._yb(r1)), self.xs.value(self._xb(c2 + 1)),
                self.ys.value(self._yb(r2 + 1)))

    @staticmethod
    def edges(r: TextRun) -> tuple[float | None, float | None, float | None]:
        """(左端, 右端, 中心)。送り幅のない字（ここで足した字など）は、寄せ方で決まる端だけが分かる。"""
        w = r.width
        a = r.align & 6
        if a == TA_RIGHT:
            return (r.x - w if w is not None else None), r.x, (r.x - w / 2 if w is not None else None)
        if a == TA_CENTER:
            return (r.x - w / 2 if w is not None else None), (r.x + w / 2 if w is not None else None), r.x
        return r.x, (r.x + w if w is not None else None), (r.x + w / 2 if w is not None else None)

    def span(self, r: TextRun) -> tuple[float, float]:
        """字の横の範囲。送り幅がなければ、字の切り抜きの矩形（ここで足した字はセルで切り抜く）か、1字 1.05 字幅の見積もり。"""
        x0, x1, _c = self.edges(r)
        if x0 is not None and x1 is not None:
            return x0, x1
        if isinstance(r.clip, tuple):
            # セルの大きさの切り抜き（ここで足した字は、そのセルで切り抜いて描く）。表全体の切り抜きは手がかりにしない
            w_clip = r.clip[2] - r.clip[0]
            if 0 < w_clip < 0.5 * (self.emf.bounds[2] - self.emf.bounds[0]) and r.clip[0] <= r.x <= r.clip[2]:
                return float(r.clip[0]), float(r.clip[2])
        w = 1.05 * abs(r.font.height) * len(r.text) if r.font else 0.0
        if x0 is not None:
            return x0, x0 + w
        if x1 is not None:
            return x1 - w, x1
        c = self.edges(r)[2]
        return c - w / 2, c + w / 2

    def _anchor_eqs(self, xs: _System, ys: _System, run: TextRun, rc: tuple[int, int]) -> None:
        t = self.texts[rc]
        if t.problem or t.parts is not None:
            return
        r1, c1, r2, c2 = self.model.area(*rc)
        h = abs(run.font.height)
        key = run.font.key
        x0, x1, xc = self.edges(run)
        if not self._shrunk(run, t):  # 縮小して描いた字：余白も縮むので、横の位置の証拠にしない（縦の中心は使う）
            L, R = self._xb(c1), self._xb(c2 + 1)
            if t.halign == "left" and x0 is not None:
                co = dict(L)
                co["m"] = 1.0
                if t.pad_left:
                    co[self._pad(t.pad_left, key)] = h
                xs.add(co, x0, ("left", rc))
            elif t.halign == "right" and x1 is not None:
                co = dict(R)
                co["m"] = -1.0
                if t.pad_right:
                    co[self._pad(t.pad_right, key)] = -h
                xs.add(co, x1, ("right", rc))
            elif t.halign == "center" and xc is not None:
                co = {k: (L.get(k, 0) + R.get(k, 0)) / 2 for k in set(L) | set(R)}
                if t.pad_left:
                    co[self._pad(t.pad_left, key)] = h / 2
                if t.pad_right:
                    co[self._pad(t.pad_right, key)] = -h / 2
                xs.add(co, xc, ("center", rc))
        bh = run.box_height
        if bh and not run.align & 0x18:
            T, B = self._yb(r1), self._yb(r2 + 1)
            if t.valign == "center":
                co = {k: (T.get(k, 0) + B.get(k, 0)) / 2 for k in set(T) | set(B)}
                co["dc"] = 1.0
                ys.add(co, run.y + bh / 2, ("center", rc))
            elif t.valign == "top":
                co = dict(T)
                co["dt"] = 1.0
                ys.add(co, run.y, ("top", rc))
            elif t.valign == "bottom":
                co = dict(B)
                co["db"] = -1.0
                ys.add(co, run.y + bh, ("bottom", rc))

    # ------------------------------------------------------------------ 字の大きさ
    def font_scale(self) -> float:
        """絵の上の字の高さ ÷ ポイント（縮小していないセルの字から）。"""
        ks = []
        for idx, rc in self.anchors.items():
            t = self.texts[rc]
            run = self._run(idx)
            if t.font is None or t.shrink or not t.font.size:
                continue
            if t.font.face.lower() != run.font.face.lower():
                continue
            ks.append(abs(run.font.height) / t.font.size)
        if not ks:
            raise PictureError("表の絵の字の大きさと、セルの字の大きさ（ポイント）の比が分かりません")
        ks.sort()
        k = ks[len(ks) // 2]
        for idx, rc in self.anchors.items():
            t = self.texts[rc]
            run = self._run(idx)
            if t.font is None or t.shrink or t.font.face.lower() != run.font.face.lower():
                continue
            if abs(round(t.font.size * k) - abs(run.font.height)) > 1:
                raise PictureError("表の絵の字の大きさが、セルの字の大きさと比例していません")
        return k

    def runs_of(self, rc: tuple[int, int]) -> list[TextRun]:
        return [self._run(i) for i, cell in self.anchors.items() if cell == rc]

    # ------------------------------------------------------------------ 描く
    def draw(self, rc: tuple[int, int], t: CellText, ih: int, new_texts: dict) -> Placed:
        """セル rc に、t の字を描くレコードの塊を作る（描けないときは PictureError）。"""
        where = addr(*rc)
        if t.problem:
            raise PictureError(f"{where}: {t.problem}")
        if t.parts is not None:
            raise PictureError(f"{where}: セルを空白で埋める表示形式（* ）の字は、まだ描けません")
        f = t.font
        k = self.font_scale()
        h_nom = int(round(f.size * k))
        base = next((x for x in self.emf.fonts if x.face.lower() == f.face.lower()), None) or self.emf.fonts[0]
        key = (f.face.lower(), f.bold, f.italic)
        L, T, R, B = self.rect(*rc)
        r1, c1, r2, c2 = self.model.area(*rc)
        m = self.m
        if m is None:
            raise PictureError(f"{where}: 表の絵の中で、セルの字の余白が分かりません（左寄せ・右寄せの字がありません）")
        # 表示形式の空白（_x）の幅：同じ空白・同じ大きさで描いた右寄せの字があれば、その右端からの間をそのまま使う
        pads: dict[str, float] = {}
        for side, pad in (("l", t.pad_left), ("r", t.pad_right)):
            if not pad:
                continue
            name = self._pad(pad, key)
            if name in self.pad_em:
                pads[side] = self.pad_em[name]
            else:
                w = self.known_width(key, h_nom, pad)
                if w is None:
                    raise PictureError(f"{where}: 表示形式の空白（_{pad}）の幅が、表の絵から分かりません")
                pads[side] = w / h_nom
        # 横の位置に要る組み合わせ（左寄せ：左の境＋余白、右寄せ：右の境−余白、中央寄せ：境の中点）が、証拠から決まること
        xs, ys = self.xs, self.ys
        Lc, Rc = self._xb(c1), self._xb(c2 + 1)
        need = {"left": {**Lc, "m": 1.0}, "right": {**Rc, "m": -1.0}}.get(
            t.halign, {k: (Lc.get(k, 0.0) + Rc.get(k, 0.0)) / 2 for k in set(Lc) | set(Rc)})
        if not xs.estimable(need):
            raise PictureError(f"{where}: 表の絵の中で、このセルに{_HALIGN_JA.get(t.halign, '')}の字を描く横の位置が決まりません"
                               "（罫線がなく、同じ寄せ方の字もないので、列の境の位置が決まりません）")
        lo, hi, _exact = self.text_width(key, h_nom, t.text)
        padw = (pads.get("l", 0) + pads.get("r", 0)) * h_nom
        exact_avail = xs.estimable(_minus(_minus(Rc, Lc), {"m": 2.0}))
        if exact_avail:
            avail = (R - L) - self.ml - self.mr
        else:
            # 列の境が決まらない（右寄せの字だけ）：使える幅の下限で、入ることを確かめる。セルの余白は左右で同じ
            # （MS-OI29500 2.1.598・ECMA-376 第1部 18.3.1.13：列の幅は、字の幅に、両側の余白と目盛り線の分を足したもの）、
            # 目盛り線の幅は 0 以上、
            # Excel が字を描く切り抜きの左端は範囲の最初の列の中の左端、とすると
            # 使える幅 = 2×（右寄せの字の右端）− 2×（切り抜きの左端）−（最初の列の左の境からこのセルの左右の境まで）＋ 線の幅 ≧ 下の値
            clip = self.text_clip()
            self.x_shift()  # 字の位置と切り抜きが食い違っていないこと
            first = self.bx(self.cols_b[0])
            avail = 2 * (R - self.mr) - 2 * clip[0] - (L - first) - (R - first)
        h = h_nom
        overflow_to = None
        if hi + padw > avail and not exact_avail:
            raise PictureError(f"{where}: 「{t.text}」が欄の幅に入るか確かめられません（表の絵に列の境の線がなく、"
                               "使える幅が分からないため。入りきらなければ、Excel では"
                               + ("「####」になります" if t.halign == "right" else "隣のセルに掛かるか、切れます") + "）")
        if hi + padw > avail:
            if t.shrink:
                w = self.known_width(key, h_nom, t.text)
                if w is None:
                    raise PictureError(f"{where}: 「{t.text}」の字の幅が表の絵から分からないので、縮小する大きさを決められません")
                h = min(h_nom, int(math.floor(h_nom * avail / (w + padw))))
                if h < 1:
                    raise PictureError(f"{where}: 「{t.text}」は欄に入りきりません")
            elif lo + padw > avail and t.halign == "left" and not t.wrap and (r1, c1) == (r2, c2):
                overflow_to = self._overflow_end(rc, t, lo, hi, L, new_texts)
            else:
                raise PictureError(f"{where}: 「{t.text}」が欄の幅に入るか確かめられません（入りきらなければ、Excel では"
                                   + ("「####」になります" if t.halign == "right" else "隣のセルに掛かるか、切れます") + "）")
        ratios = self.heights.get(key)
        same = None if ratios else self._same_height_run(rc, t, key, h)
        if not ratios and same is None:
            raise PictureError(f"{where}: 表の絵の中に、この字体（{f.face}）で描いた字の高さの手がかりがありません")
        hb = sorted(ratios)[len(ratios) // 2] * h if ratios else same[1]
        # 縦の位置：Excel と同じく、行の中心 − 字の高さ/2 + ずれ を切り捨てる
        if same is not None:
            # 字の高さが分からない（GDI が書いた字の外接矩形が切り抜きで欠けている）：同じ字体・同じ大きさ・同じ縦の寄せ方で
            # 描いた字の上端から、行の境の差（証拠から決まるもの）だけずらす。同じ行なら、その字と同じ高さ
            y = same[0]
        elif t.valign == "center":
            if self.dc is None:
                raise PictureError(f"{where}: 表の絵の中に、縦の中央に寄せた字がないので、描く高さが分かりません")
            y = (T + B) / 2 + self.dc - hb / 2
        elif t.valign == "top":
            if self.dt is None:
                raise PictureError(f"{where}: 表の絵の中に、上に寄せた字がないので、描く高さが分かりません")
            y = T + self.dt
        else:
            if self.db is None:
                raise PictureError(f"{where}: 表の絵の中に、下に寄せた字がないので、描く高さが分かりません")
            y = B - self.db - hb
        Tc, Bc = self._yb(r1), self._yb(r2 + 1)
        yneed = {"center": {**{k: (Tc.get(k, 0.0) + Bc.get(k, 0.0)) / 2 for k in set(Tc) | set(Bc)}, "dc": 1.0},
                 "top": {**Tc, "dt": 1.0}}.get(t.valign, {**Bc, "db": -1.0})
        if same is None and not ys.estimable(yneed):
            raise PictureError(f"{where}: 表の絵の中で、このセルの字を描く高さが決まりません（罫線がなく、同じ縦の寄せ方の字もありません）")
        # 横の位置（右寄せ・中央寄せは GDI の寄せ方で描くので、字の幅が分からなくても右端・中心は正しい）
        s = h / h_nom  # 縮小して描くときは、余白と空白も同じ割合で縮む（記入例の絵の字の位置と1単位以内で合う）
        gap = self.gaps.get((self._pad(t.pad_right, key), h)) if t.pad_right else None
        if t.halign == "left":
            x, align = L + self.ml * s + pads.get("l", 0) * h, TA_LEFT
        elif t.halign == "right":
            x = R - gap if gap is not None else R - self.mr * s - pads.get("r", 0) * h
            align = TA_RIGHT
        else:
            x, align = (L + R) / 2 + (self.dh or 0.0) + (pads.get("l", 0) - pads.get("r", 0)) * h / 2, TA_CENTER
        xi, yi = int(math.floor(x + 1e-6)), int(math.floor(y + 1e-6))  # 解の計算の小さな誤差（847.9999…）で1つ下がらない
        dx = None
        ex = self.exact.get((key, h), {})
        if all(ch in ex for ch in t.text):
            dx = [ex[ch] for ch in t.text]
        w_est = sum(dx) if dx else (self.known_width(key, h, t.text) or hi * h / h_nom)
        x0 = xi if align == TA_LEFT else (xi - w_est if align == TA_RIGHT else xi - w_est / 2)
        bounds = (int(x0), yi, int(x0 + w_est) - 1, int(yi + hb) - 1)
        ref = self.runs[0]
        clip = (int(math.floor(L)), int(math.floor(T)), int(math.ceil(overflow_to if overflow_to else R)), int(math.ceil(B)))
        if not (self.x_fixed and self.y_fixed):
            # 境の位置が決まらない向きは、Excel がセルの字を描いたときと同じ切り抜き（表の範囲）で描く。
            # 字はセルに入ることを確かめてあるので、セルで切り抜いたときと同じに描かれる
            tc = self.text_clip()
            if tc is None:
                raise PictureError(f"{where}: 表の絵の中で、セルの境が決まらず、Excel が字を描いた切り抜きも字ごとに違います")
            clip = (clip[0] if self.x_fixed else tc[0], clip[1] if self.y_fixed else tc[1],
                    clip[2] if self.x_fixed else tc[2], clip[3] if self.y_fixed else tc[3])
        recs = [savedc(), intersect_clip(clip),
                create_font(ih, base, -h, 700 if f.bold else 400, f.italic, f.underline == 1, f.strike,
                            f.charset if f.charset is not None else base.charset, f.face),
                select(ih), text_color(colorref(t.color)), bk_transparent(), text_align(align),
                ext_text_out(xi, yi, t.text, dx, bounds, ref.gmode, ref.scales), restoredc(), delete(ih)]
        return Placed(rc, t.text, recs, (xi, yi), (yi, yi + hb))

    def _same_height_run(self, rc: tuple[int, int], t: CellText, key: tuple, h: int) -> tuple[float, float] | None:
        """同じ字体・同じ大きさ・同じ縦の寄せ方で描いた字から、セル rc に描く字の上端と、字の高さ（GDI が書いた外接矩形の
        下端まで）を求める。行の境の差が証拠から決まるものだけを使う。なければ None。"""
        r1, _c1, r2, _c2 = self.model.area(*rc)
        Tn, Bn = self._yb(r1), self._yb(r2 + 1)
        for idx, cell in sorted(self.anchors.items()):
            run, ct = self._run(idx), self.texts[cell]
            if run.font is None or run.font.key != key or abs(run.font.height) != h or run.align & 0x18:
                continue
            if ct.valign != t.valign or ct.problem or self._shrunk(run, ct):
                continue
            q1, _d1, q2, _d2 = self.model.area(*cell)
            To, Bo = self._yb(q1), self._yb(q2 + 1)
            if t.valign == "top":
                diff = _minus(Tn, To)
            elif t.valign == "center":
                diff = _minus({k: (Tn.get(k, 0.0) + Bn.get(k, 0.0)) / 2 for k in set(Tn) | set(Bn)},
                              {k: (To.get(k, 0.0) + Bo.get(k, 0.0)) / 2 for k in set(To) | set(Bo)})
            else:
                diff = _minus(Bn, Bo)
            if not any(abs(v) > 1e-12 for v in diff.values()):
                shift = 0.0
            elif self.ys.estimable(diff):
                shift = self.ys.value(diff)
            else:
                continue
            bottom = run.bounds[3]
            if bottom < run.y:
                continue
            return run.y + shift, float(bottom - run.y + 1)
        return None

    def _overflow_end(self, rc, t: CellText, lo: float, hi: float, L: float, new_texts: dict) -> float:
        """左寄せの字がセルからはみ出す：右の空のセルに掛けて描く（Excel と同じ）。字の幅が正確に分かるときだけ。"""
        r, c = rc
        if not self.x_fixed:
            raise PictureError(f"{addr(*rc)}: 「{t.text}」がセルに入りきらず、隣に掛かる長さが表の絵から確かめられません（列の境が決まりません）")
        if lo != hi:
            raise PictureError(f"{addr(*rc)}: 「{t.text}」がセルに入りきらず、隣に掛かる長さが表の絵から確かめられません")
        end = None
        for cc in range(c + 1, self.model.range[3] + 2):
            right = self.bx(cc)
            if L + self.ml + hi <= right - self.mr:
                end = right
                break
            nxt = new_texts.get((r, cc))
            if cc > self.model.range[3] or (nxt is not None and nxt.text) or self.model.sheet.merged_at(r, cc):
                break
        if end is None:
            raise PictureError(f"{addr(*rc)}: 「{t.text}」がセルに入りきらず、右のセルにも字があるので、Excel では途中で切れます")
        # はみ出した字の下の、セルの境の線：Excel は、字が掛かる所の目盛り線を描かない。絵の線は消さないので、線があれば書かない
        _T, B_ = self.by(r), self.by(r + 1)
        R0 = self.bx(c + 1)
        for ln in self.emf.lines:
            # 行の境が決まらないときは、どの高さの線も、この行に掛かるものとみなす
            if ln.orient == "v" and R0 - 1 <= ln.at <= end - 1 and (not self.y_fixed or (ln.lo < B_ - 1 and ln.hi > _T + 1)):
                raise PictureError(f"{addr(*rc)}: 「{t.text}」がセルに入りきらず右のセルに掛かりますが、その所に描かれている"
                                   "セルの境の線を、Excel と同じに消せません")
        return end


# ====================================================================== 入口：絵を書き直す
@dataclass
class PictureEdit:
    data: bytes
    removed: list[tuple[tuple[int, int], str]] = field(default_factory=list)
    added: list[Placed] = field(default_factory=list)


def changed_cells(old: SheetModel, new: SheetModel) -> list[tuple[tuple[int, int], CellText, CellText]]:
    """絵に映る範囲で、見える字か描き方が変わったセル。"""
    out = []
    # 要素のないマスは、どちらのブックでも字がない。片方にだけ要素のあるマス（空のマスに書いた所）は、新しいブックの側で見つかる
    for rc in sorted({*old.display_cells(), *(x for x in new.display_cells() if old.visible(*x))}):
        a, b = old.text(*rc), new.text(*rc)
        if (a.text, a.parts, a.color, a.halign, a.pad_left, a.pad_right) != (b.text, b.parts, b.color, b.halign, b.pad_left,
                                                                           b.pad_right):
            out.append((rc, a, b))
    return out


def edit_picture(data: bytes, old: SheetModel, new: SheetModel, what: str = "表の絵") -> PictureEdit:
    """表の絵（EMF）を、old のブックから new のブックに変わったセルの字だけ描き直す。変わらなければ元のまま返す。"""
    changes = changed_cells(old, new)
    if not changes:
        return PictureEdit(data)
    try:
        emf = Emf(data)
        lay = Layout(emf, old)
        at = emf.insertion_point()
    except (MetafileError, PictureError) as e:
        raise PictureError(f"{what}: {e}") from None
    xform_ok = emf.state_at[at]["xform"] == lay.xform
    if not xform_ok:
        raise PictureError(f"{what}: 字を足す所の座標の決め方が、セルの字を描いた所と違います")
    new_texts = {rc: new.text(*rc) for rc in new.display_cells()}
    removed: set[int] = set()
    out = PictureEdit(data)
    blocks: list[bytes] = []
    ih = emf.n_handles
    try:
        for rc, a, b in changes:
            if a.problem:
                raise PictureError(f"{addr(*rc)}: いまの字を、Excel と同じに求められません（{a.problem}）")
            if any(rule.covers(*rc) for rule in old.sheet.cond_rules):
                raise PictureError(f"{addr(*rc)}: 条件付き書式のあるセルなので、値で変わる色や見え方を描き直せません")
            if a.text:
                runs = lay.runs_of(rc)
                want = sorted(a.lines)
                if sorted(r.text for r in runs) != want:
                    raise PictureError(f"{addr(*rc)}: いまの字「{a.text}」が、表の絵の中に見つかりません（絵がブックと合っていません）")
                _L0, L1, R0, _R1 = lay.x_range(*rc)  # 境が決まらないときは、ありうる中でいちばん狭いセル
                for r in runs:
                    x0, x1 = lay.span(r)
                    if x0 < L1 - 1 or x1 > R0 + 1:
                        # Excel は、はみ出した字が掛かる所の目盛り線を描いていない。字を変えると、その線の見え方も変わる
                        raise PictureError(f"{addr(*rc)}: いまの字「{r.text}」がセルからはみ出して描かれているので、"
                                           "書いた後の見え方（はみ出しが無くなった所の線など）を描き直せません")
                removed |= {r.index for r in runs}
                out.removed += [(rc, r.text) for r in runs]
            if b.text:
                p = lay.draw(rc, b, ih, new_texts)
                # 境が決まらないときは、ありうる中でいちばん広いセル（縦は、描く字の上端から下端）で確かめる
                L, _L1, _R0, R = lay.x_range(*rc)
                _L, T, _R, B = lay.rect(*rc)
                if not lay.y_fixed:
                    T, B = p.band
                for r in lay.runs:
                    if r.index in removed or lay.anchors.get(r.index) == rc:
                        continue
                    x0, x1 = lay.span(r)
                    bh = r.box_height or abs(r.font.height)
                    if x0 < R - 1 and x1 > L + 1 and r.y < B - 1 and r.y + bh > T + 1:
                        raise PictureError(f"{addr(*rc)}: 隣のセルの字「{r.text}」がこのセルに掛かって描かれているので、"
                                           "書いた後の見え方（隣の字が切れる）を描き直せません")
                blocks += p.records
                out.added.append(p)
            elif a.text:
                _check_neighbors(lay, rc)
    except PictureError as e:
        raise PictureError(f"{what}の {e}") from None
    out.data = emf.rebuild(removed, blocks, ih + 1 if blocks else emf.n_handles)
    return out


def _check_neighbors(lay: Layout, rc: tuple[int, int]) -> None:
    """空にしたセル：左のセルの字がはみ出していたら（Excel では、空になると続きが見えるようになる）書かない。"""
    r, c = rc
    for idx, cell in lay.anchors.items():
        if cell[0] != r or cell == rc:
            continue
        run = lay._run(idx)
        x0, x1 = lay.span(run)
        _L0, L, R, _R1 = lay.x_range(*cell)  # 境が決まらないときは、ありうる中でいちばん狭いセル
        if x1 > R + 1 or x0 < L - 1:
            raise PictureError(f"{addr(*rc)}: 隣のセルの字「{run.text}」がはみ出しているので、空にした後の見え方を描き直せません")
