"""Excel（.xlsx）：セルの字（選択肢の「有」「2.更新」など）を、○の図形で囲む。

指定（fill の値）：{"@c1": {"circle": 囲む字, "cell": "シート!B3"（省くと、ブックの字から探す）, "near": 近くの見出しの字,
"sheet": シート, "occurrence": 何番目か, "pad": 字と○のすき間（pt）, "style": "ellipse"（既定）か "border"}}。

style "border"（○でなくてよいと本人が確かめたとき。楕円がほかの字に触れる詰まった所のため）：Excel には、セルの中の字の一部だけを
囲む罫線（Word の字の囲み線にあたるもの）がない。セルの罫線は、セル全体（同じセルのほかの選択肢も）を囲んでしまい、セルの書式も
変えるので使わない。代わりに、見積もった字の箱をすき間だけ広げた四角の図形（prstGeom rect・塗りなし）を、楕円と同じ書き方で足す。

書き方：シートの図形（DrawingML）に、楕円（prstGeom ellipse・塗りなし・黒い線）を1つずつ足す（xdr:oneCellAnchor）。
シートに図形の部品がなければ、部品（xl/drawings/drawingN.xml）と、シートからの関係・[Content_Types].xml の登録・シートの
<drawing> を足す。セルの値・書式・ほかの部品は変えない。

位置：Excel はセルの中の字の位置を保存しないので、セルの大きさ（列の幅・行の高さ）・書式（揃え・字の大きさ）・字の幅
（全角は字の大きさ、半角はその半分と見積もる）から決める。決め方の分からない揃え（均等割り付け・両端揃え・縮小して全体を
表示・折り返して自動で改行される字・字の回転）のセルは、置かずに理由と代わりの書き方（字の前に「○」を足す・字を書き換える。
本人に確かめてから）を返す。位置は見積もりなので、結果には「確かめが要る」ことを添える。

見積もりの元：列の幅は、既定の書体の数字の幅（ＭＳ Ｐゴシック・游ゴシック 11pt は 8 ピクセル、Calibri 11pt は 7 ピクセル。
Excel が図形の位置を保存した様式の列の幅で確かめた値）× 保存された字の数。揃え（左・中央・右）・字下げ（1 段で数字 3 つ分）・
セルの余白（約 2 ピクセル）・結合の範囲・縦の揃え・行の高さの比（書体の種類ごとの Excel の既定の行の高さ）・字の幅（全角は
字の大きさ、ほかは半分）。書体の幅の種類（「Ｐ」の書体はかな・記号の幅が字ごとに変わる、欧文の書体は半角の字の幅が変わる）
から、囲む字の中心の位置の誤差を見積もり（error_estimate）、字の幅の半分を超えるなら描かずに断る（「位置を確かめられない」）。
折り返す・縮小して全体を表示するセルで、行の字の幅がセルの幅に近く、改行・縮小されるかが見積もりの誤差の中なら、同じく断る。
"""

from __future__ import annotations

import math
import posixpath
import re
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass

from ..formula.parser import addr, parse_cell
from ..textwidth import cells, east_asian_context
from ..pdf.circle import _pattern, shape, touches, face, inside
from ..pdf.circle import around as _around
from ..pdf.circle import border as _border
from ..pdf.circle import encloses as _encloses
from ..pdf.circle import hits as _hits

EMU = 12700  # 1 pt の EMU
TAG = "TdkCircle"
NS_X = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
UNCERTAIN = ("Excel はセルの中の字の位置を保存しないので、○の位置はセルの大きさ・揃え・字の幅から見積もったものです。"
             "ずれていないか、印刷した紙か画面で確かめてください")


class CircleError(ValueError):
    pass


@dataclass(eq=False)
class G:
    """見積もった字の箱（シートの左上からの pt）。pdf.circle の shape・face に渡せる形。"""

    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size: float
    vertical: bool = False


# ---------------------------------------------------------------------- 書式
def _styles(wb) -> dict:
    memo = getattr(wb, "_circle_styles", None)
    if memo is not None:
        return memo
    out = {"xfs": [], "fonts": []}
    target = None
    for typ, t in wb.rels(wb.wb_path).values():
        if typ == "styles":
            target = t
    if target and wb.zip.has(target):
        root = ET.fromstring(wb.zip.read(target))
        m = "{%s}" % root.tag[1:].split("}")[0] if root.tag.startswith("{") else ""
        for f in root.find(m + "fonts") if root.find(m + "fonts") is not None else []:
            sz = f.find(m + "sz")
            nm = f.find(m + "name")
            out["fonts"].append((nm.get("val") if nm is not None else "", float(sz.get("val")) if sz is not None else 11.0))
        xfs = root.find(m + "cellXfs")
        for xf in xfs if xfs is not None else []:
            al = xf.find(m + "alignment")
            g = (lambda k, d=None: al.get(k, d)) if al is not None else (lambda k, d=None: d)
            out["xfs"].append({"font": int(xf.get("fontId", "0")), "h": g("horizontal", "general"), "v": g("vertical", "bottom"),
                               "wrap": g("wrapText") in ("1", "true"), "shrink": g("shrinkToFit") in ("1", "true"),
                               "indent": int(g("indent", "0") or 0), "rot": int(g("textRotation", "0") or 0)})
    wb._circle_styles = out
    return out


def _font(wb, style: int) -> tuple[str, float]:
    st = _styles(wb)
    xf = st["xfs"][style] if 0 <= style < len(st["xfs"]) else {"font": 0}
    fonts = st["fonts"] or [("", 11.0)]
    return fonts[xf["font"]] if 0 <= xf["font"] < len(fonts) else fonts[0]


@dataclass(frozen=True)
class Family:
    """書体の幅の種類（名前から決める。書体のファイルは読まない）。
    cjk_prop：かな・全角の記号の幅が字ごとに変わる（ＭＳ Ｐゴシックなど「Ｐ」の書体・欧文の書体の代わりの字）。
    latin_mono：半角の字が字の大きさの半分にそろう（ＭＳ ゴシックなど）。
    pitch：Excel の行の高さ（1 行）÷ 字の大きさ。pitch_known：その値が Excel の既定の行の高さから分かる書体か。
    digit：数字の幅（字の大きさに対する比。列の幅の単位を決める）。known：名前の分かる書体か。"""

    cjk_prop: bool
    latin_mono: bool
    pitch: float
    pitch_known: bool
    digit: float
    known: bool


_P_FONT = re.compile(r"(ＭＳ\s*Ｐ|MS\s*P(?=[A-Z])|^HG[PS]|UDP|^IPAP|\bUI$|\sUI\b)", re.I)
_MONO_CJK = re.compile(r"^(ＭＳ\s*(ゴシック|明朝)|MS\s*(Gothic|Mincho)|HG(?![PS])|IPA(?!P)|BIZ\s*UD(?!P)|Osaka.等幅)", re.I)
_YU = re.compile(r"^(游|Yu\s)", re.I)
_MEIRYO = re.compile(r"^(メイリオ|Meiryo)", re.I)
_LATIN = {"calibri": (1.364, 0.507), "arial": (1.275, 0.556), "times new roman": (1.25, 0.5), "century": (1.3, 0.556),
          "cambria": (1.3, 0.556), "verdana": (1.3, 0.636), "tahoma": (1.3, 0.546), "segoe ui": (1.36, 0.559)}


def family(name: str) -> Family:
    n = (name or "").strip()
    if _P_FONT.search(n):
        return Family(True, False, 1.705 if _YU.search(n) or _MEIRYO.search(n) else 1.227,
                      bool(re.search(r"ＭＳ|MS\s*P|^HG", n)), 0.5, True)
    if _MONO_CJK.search(n):
        return Family(False, True, 1.227, bool(re.search(r"ＭＳ|^MS|^HG", n)), 0.5, True)
    if _YU.search(n) or _MEIRYO.search(n):
        return Family(False, False, 1.705, True, 0.5, True)
    if re.search(r"[^\x00-\x7f]", n) or re.search(r"(Noto|Source Han|Hiragino|Gothic|Mincho|Sans CJK|SimSun|Malgun|"
                                                    r"PMingLiU|MingLiU|Batang|Gulim)", n, re.I):
        return Family(False, False, 1.3, False, 0.5, False)  # 東アジアの書体（名前から幅の種類が分からない）
    lat = _LATIN.get(n.lower())
    if lat:
        return Family(True, False, lat[0], True, lat[1], True)  # 欧文の書体：かな・漢字は代わりの書体で描かれる
    return Family(True, False, 1.3, False, 0.5, False)


def _px(size: float) -> int:
    """字の大きさ（pt）→ 96 dpi の字の高さ（ピクセル。四捨五入）。"""
    return int(size * 96 / 72 + 0.5)


def mdw_of(wb) -> tuple[int, bool]:
    """列の幅の単位（既定の書体の数字の幅、ピクセル）と、その値が確かか。和文の書体は数字が字の大きさの半分で、
    ＭＳ Ｐゴシック 11pt は 8 ピクセル（Excel が図形の位置を保存した様式の列の幅で確かめた値）。Calibri 11pt は 7。"""
    name, size = (_styles(wb)["fonts"] or [("", 11.0)])[0]
    return mdw_for(name, size)


def mdw_for(name: str, size: float) -> tuple[int, bool]:
    """書体の名前と字の大きさ（pt）から、数字の幅（ピクセル）と、その値が確かか。"""
    fam = family(name)
    ppem = _px(size)
    if fam.digit == 0.5:
        return max(3, int(ppem * 0.5 + 0.5)), fam.known
    return max(3, int(ppem * fam.digit)), fam.known


def width_px(w: float | None, mdw: int, base: int = 8) -> int:
    """列の幅（保存された字の数。None は幅の指定のない列：baseColWidth と余白 5 ピクセルを 8 の倍数に切り上げた幅）の、ピクセル。"""
    if w is None:
        return -(-(base * mdw + 5) // 8) * 8
    return int(((256 * w + int(128 / mdw)) / 256) * mdw)


def _mdw(wb) -> int:
    return mdw_of(wb)[0]


def col_px(wb, sh, c: int) -> int:
    """列の幅（ピクセル。Excel の決め方：保存された字の数 × 数字の幅。幅のない列は defaultColWidth、それもなければ
    baseColWidth の字の数と余白 5 ピクセルを 8 の倍数に切り上げた幅）。"""
    if c in sh.hidden_cols:
        return 0
    return width_px(sh.col_widths.get(c, sh.default_col_width), _mdw(wb), getattr(sh, "base_col_width", 8) or 8)


def col_pt(wb, sh, c: int) -> float:
    return col_px(wb, sh, c) * 0.75


def row_pt(sh, r: int) -> float:
    if r in sh.hidden_rows:
        return 0.0
    return sh.row_heights.get(r, sh.default_row_height or 15.0)


def width(ch: str, size: float, ambiguous_wide: bool = False) -> float:
    """字の幅の見積もり（textwidth：Unicode の東アジアの幅）：全角（W・F）の字は字の大きさ、曖昧（A：○・①・ギリシャ文字・
    キリル文字など）は東アジアの文脈（ambiguous_wide）なら字の大きさ・ほかは半分、結合の記号は 0、ほかは半分。"""
    if ch == "\n":
        return 0.0
    return 0.5 * size * cells(ch, ambiguous_wide)


def _ideograph(ch: str) -> bool:
    o = ord(ch)
    return 0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF or 0x20000 <= o <= 0x3FFFF or 0xF900 <= o <= 0xFAFF or ch in "々〆〇"


def spread(ch: str, fam: Family, wide: bool) -> float:
    """字の幅の見積もりの、ありうるずれ（字の大きさに対する比）。漢字は、どの書体でも字の大きさ。かな・全角の記号は、
    「Ｐ」の書体と欧文の書体（代わりの書体）では字ごとに幅が変わる。半角の字は、等幅の和文の書体だけがそろう。"""
    if ch == "\n" or unicodedata.category(ch) in ("Mn", "Me", "Cf"):
        return 0.0
    n = cells(ch, wide)
    if n == 2:
        if _ideograph(ch) or ch == "　":  # 漢字と全角の空白は、「Ｐ」の書体でも字の大きさ
            return 0.0
        if unicodedata.east_asian_width(ch) == "A":
            return 0.25 if fam.cjk_prop else 0.05
        return 0.25 if fam.cjk_prop else 0.0
    if fam.latin_mono:
        return 0.0
    if ch.isdigit():
        return 0.06
    if ch == " ":
        return 0.15
    return 0.2


# ---------------------------------------------------------------------- 囲む所
def _text(wb, cell) -> str:
    from .window import display

    return display(wb, cell) if cell is not None else ""


def _cells(wb, sheet: str | None):
    for sh in wb.sheets:
        if sh.state.startswith("non-worksheet") or (sheet and sh.name != sheet):
            continue
        for (r, c), cell in sorted(sh.cells.items()):
            t = _text(wb, cell)
            if t:
                yield sh, r, c, t


def find(wb, spec: dict):
    """(シート, 行, 列, セルの字, 始め, 終わり)。"""
    target = spec.get("circle")
    if not isinstance(target, str) or not target.strip():
        raise CircleError('○で囲むときは {"circle": 囲む字, "cell": "シート!B3"（または "near": 近くの見出し）} の形で指定してください')
    pat = _pattern(target)
    sheet = spec.get("sheet")
    if sheet is not None and not wb.has_sheet(sheet):
        raise CircleError(f"シート「{sheet}」がありません")
    if spec.get("cell"):
        from .edit import EditError, resolve_target

        ref = spec["cell"] if "!" in spec["cell"] or not sheet else f"{sheet}!{spec['cell']}"
        try:
            sn, r, c = resolve_target(wb, ref)
        except EditError as e:
            raise CircleError(str(e)) from None
        sh = wb.sheet(sn)
        t = _text(wb, sh.cells.get((r, c)))
        ms = [m for m in pat.finditer(t) if m.end() > m.start()]
        if not ms:
            raise CircleError(f"{sn}!{addr(r, c)} の字「{t[:40]}」に、「{target}」がありません")
        occ = spec.get("occurrence")
        if len(ms) > 1 and occ is None:
            raise CircleError(f"{sn}!{addr(r, c)} に「{target}」が {len(ms)} か所あります。occurrence（セルの中の何番目か）で決めてください")
        m = ms[(occ or 1) - 1] if occ is None or (isinstance(occ, int) and 1 <= occ <= len(ms)) else None
        if m is None:
            raise CircleError(f"occurrence は 1〜{len(ms)} で指定してください")
        return sh, r, c, t, m.start(), m.end()
    hits = [(sh, r, c, t, m.start(), m.end()) for sh, r, c, t in _cells(wb, sheet) for m in pat.finditer(t) if m.end() > m.start()]
    if not hits:
        raise CircleError(f"「{target}」がブックのセルに見つかりません（窓の字のとおりに、または cell でセルを指定してください）")
    from ..docx.circle import CircleError as _WordCircleError
    from ..docx.circle import choose

    def anchors_of(near: str):
        np_ = _pattern(near)
        return [((sh.name, r, c), m.start(), m.end(), (sh, r, c)) for sh, r, c, t in _cells(wb, sheet)
                for m in np_.finditer(t) if m.end() > m.start()]

    def rank_of(h, anchors):
        best = (9, 10 ** 9)
        for _key, _a0, _a1, (sh, r, c) in anchors:
            if sh is not h[3][0]:
                continue
            hr, hc = h[3][1], h[3][2]
            if r == hr:
                k = (0, abs(c - hc))  # 同じ行（同じセルは 0）
            elif c == hc:
                k = (1, abs(r - hr))  # 同じ列
            else:
                k = (2, abs(r - hr) + abs(c - hc))
            best = min(best, k)
        return best

    wrapped = [((h[0].name, h[1], h[2]), h[4], h[5], h) for h in hits]
    try:
        got = choose(wrapped, spec.get("near"), spec.get("occurrence"), anchors_of, rank_of,
                     lambda w: (w[3][3], w[1], w[2]), target)
    except _WordCircleError as e:
        raise CircleError(str(e)) from None
    return got[3]


@dataclass
class Layout:
    """セルの字の置き方の見積もり（glyphs の中身）：誤差の見積もりにも使う。"""

    glyphs: list
    notes: list
    lines: list  # 行ごとの (始めの字の番号, 終わりの字の番号)
    hal: str
    val: str
    size: float
    fam: Family
    wide: bool
    pitch: float
    rng: tuple  # 結合の範囲 (r1, c1, r2, c2)
    x0: float
    w: float


_PAD = 2 * 0.75  # セルの縁から字までの余白（Excel は約 2 ピクセル）


def layout(wb, sh, r: int, c: int, text: str) -> Layout:
    """セルの字の箱の見積もり（シートの左上からの pt）。決め方の分からないセルは CircleError。
    揃え（左・中央・右。均等割り付け・両端揃えは断る）・字下げ・字の大きさと書体の幅の種類・セルの余白（約 2 ピクセル）・
    結合の範囲・縦の揃え・字の幅（textwidth の東アジアの幅）から決める。"""
    cell = sh.cells.get((r, c))
    style = cell.style if cell is not None else sh.style_of(r, c)
    st = _styles(wb)
    xf = st["xfs"][style] if 0 <= style < len(st["xfs"]) else {"h": "general", "v": "bottom", "wrap": False, "shrink": False,
                                                              "indent": 0, "rot": 0}
    name, size = _font(wb, style)
    fam = family(name)
    notes = []
    if xf["rot"]:
        raise CircleError("字を回して表示するセルなので、字の位置を決められません")
    if xf["h"] in ("distributed", "justify", "fill", "centerContinuous") or xf["v"] in ("distributed", "justify"):
        raise CircleError(f"セルの揃え（{xf['h'] if xf['h'] not in ('general', 'left', 'center', 'right') else xf['v']}）では、"
                          "字の位置を決められません")
    m = sh.merged_at(r, c)
    r1, c1, r2, c2 = m if m is not None else (r, c, r, c)
    x0 = sum(col_pt(wb, sh, k) for k in range(1, c1))
    y0 = sum(row_pt(sh, k) for k in range(1, r1))
    w = sum(col_pt(wb, sh, k) for k in range(c1, c2 + 1))
    h = sum(row_pt(sh, k) for k in range(r1, r2 + 1))
    lines = text.split("\n") if xf["wrap"] else [text]  # 折り返さないセルの改行は、幅のない字として
    # 曖昧な幅の字は、セルの字に全角の字があるか、書体の名前が東アジアの字（ＭＳ ゴシックなど）のときだけ全角
    wide = east_asian_context(text) or bool(re.search(r"[^\x00-\x7f]", name or ""))
    lw = [sum(width(ch, size, wide) for ch in ln) for ln in lines]
    lu = [math.sqrt(sum((spread(ch, fam, wide) * size) ** 2 for ch in ln)) for ln in lines]
    indent = xf["indent"] * 3 * _mdw(wb) * 0.75 if xf["h"] in ("left", "right", "general") else 0.0
    avail = w - 2 * _PAD - indent
    if xf["wrap"] and max(lw) > avail:
        raise CircleError("折り返して表示するセルで、字が自動で改行される（改行の位置が分からない）ので、字の位置を決められません")
    if xf["wrap"] and any(a + 2 * u > avail for a, u in zip(lw, lu)):
        raise CircleError("位置を確かめられない：折り返して表示するセルで、行の字の幅がセルの幅に近く、自動で改行されるかが"
                          "見積もりでは分かりません。字の前に「○」を足す書き方（old/new。本人に確かめてから）にしてください")
    if xf["shrink"] and max(lw) > avail:
        raise CircleError("縮小して全体を表示するセルで、字の大きさが分からないので、字の位置を決められません")
    if xf["shrink"] and any(a + 2 * u > avail for a, u in zip(lw, lu)):
        raise CircleError("位置を確かめられない：縮小して全体を表示するセルで、字の幅がセルの幅に近く、縮むかが見積もりでは"
                          "分かりません。字の前に「○」を足す書き方（old/new。本人に確かめてから）にしてください")
    pitch = size * fam.pitch
    n = len(lines)
    if xf["v"] == "top":
        top = y0 + _PAD / 2
    elif xf["v"] == "center":
        top = y0 + (h - n * pitch) / 2
    else:
        top = y0 + h - _PAD / 2 - n * pitch
    numeric = cell is not None and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
    boolean = cell is not None and isinstance(cell.value, bool)
    hal = xf["h"] if xf["h"] != "general" else ("right" if numeric else "center" if boolean else "left")
    if xf["indent"]:
        notes.append("字下げ（インデント）のあるセル")
    if any(spread(ch, fam, wide) for ch in text if ch != "\n"):
        # ラテン文字・キリル文字・ギリシャ文字・アラビア文字・「Ｐ」の書体のかななどは、書体で字ごとに幅が変わる（プロポーショナル）
        notes.append("半角の字・幅の変わる書体の字（ラテン文字・キリル文字・「Ｐ」の書体のかななど）を含む（字の幅は見積もり）")
    out: list[G] = []
    spans = []
    k0 = 0
    for k, (ln, wd) in enumerate(zip(lines, lw)):
        if hal == "left":
            x = x0 + _PAD + indent
        elif hal == "center":
            x = x0 + (w - wd) / 2
        else:
            x = x0 + w - _PAD - indent - wd
        y = top + k * pitch + (pitch - size) / 2
        for ch in ln:
            cw = width(ch, size, wide)
            out.append(G(ch, x, y, x + cw, y + size, size))
            x += cw
        spans.append((k0, k0 + len(ln)))
        k0 += len(ln)
        if k < len(lines) - 1:
            out.append(G("\n", x, y, x, y + size, size))
            k0 += 1
    return Layout(out, notes, spans, hal, xf["v"] if xf["v"] in ("top", "center") else "bottom", size, fam, wide, pitch,
                  (r1, c1, r2, c2), x0, w)


def glyphs(wb, sh, r: int, c: int, text: str) -> tuple[list[G], list[str]]:
    """セルの字の箱の見積もり（シートの左上からの pt）と、見積もりの確かさに関わる注意。決め方の分からないセルは CircleError。"""
    lay = layout(wb, sh, r, c, text)
    return lay.glyphs, lay.notes


def error_estimate(wb, sh, lay: Layout, text: str, a: int, b: int, box_x: float) -> tuple[float, float, list[str]]:
    """囲む字（text[a:b]）の中心の位置の、見積もりの誤差（横・縦、pt）と、その元。
    横：囲む字より前（左揃え）・後ろ（右揃え）・前後の差の半分（中央揃え）の字の幅のずれ（字ごとに独立として二乗和の平方根）、
    セルの余白（1 ピクセル）、字下げの単位（1 段ごとに数字 1 つ分）、列の幅の単位が分からない書体なら、図形を置く列までの
    列の幅のずれ（列の幅の字の数 × 1 ピクセル）。縦：行の高さの比が分からない書体では、基準の行から囲む字の行までの行の数 ×
    字の大きさの 2 割、行の中の字の位置（字の大きさの 1 割）。"""
    size, fam, wide = lay.size, lay.fam, lay.wide
    line = next(k for k, (s, e) in enumerate(lay.lines) if s <= a <= e)
    s, e = lay.lines[line]
    before = [spread(ch, fam, wide) * size for ch in text[s:a]]
    after = [spread(ch, fam, wide) * size for ch in text[b:e]]
    rss = lambda xs: math.sqrt(sum(x * x for x in xs))  # noqa: E731
    why = []
    if lay.hal == "left":
        ex = rss(before) + 0.75
    elif lay.hal == "right":
        ex = rss(after) + 0.75
    else:
        ex = 0.5 * math.sqrt(sum(x * x for x in before + after))
    if ex > 0.75 + 0.5:
        why.append("前後の字の幅が書体で変わる")
    st = _styles(wb)
    cell = sh.cells.get(lay.rng[:2])
    style = cell.style if cell is not None else sh.style_of(*lay.rng[:2])
    xf = st["xfs"][style] if 0 <= style < len(st["xfs"]) else {"indent": 0}
    if xf.get("indent") and lay.hal != "center":
        ex += xf["indent"] * _mdw(wb) * 0.75
        why.append("字下げの幅")
    mdw, known = mdw_of(wb)
    if not known:
        # 列の幅の単位（数字の幅）が 1 ピクセル違うと、列ごとに「保存された字の数」ピクセルずれる
        r1, c1, r2, c2 = lay.rng
        anchor_col = _anchor_cell(wb, sh, box_x, 0)[0]
        chars = lambda lo, hi: sum((sh.col_widths.get(k, sh.default_col_width) or 8.43) for k in range(lo, hi)  # noqa: E731
                                   if k not in sh.hidden_cols)
        to_anchor = chars(c1, anchor_col) if anchor_col > c1 else -chars(anchor_col, c1)
        whole = chars(c1, c2 + 1)
        ref = {"left": 0.0, "center": whole / 2, "right": whole}[lay.hal if lay.hal in ("left", "right") else "center"]
        ex += abs(ref - to_anchor) * 0.75
        why.append("既定の書体が分からず、列の幅の単位が見積もり")
    n = len(lay.lines)
    lines_from = {"top": line + 0.5, "bottom": n - line - 0.5, "center": abs(line - (n - 1) / 2)}[lay.val]
    ey = (0.03 if fam.pitch_known else 0.2) * size * lines_from + 0.1 * size + (0.75 if lay.val != "center" else 0.0)
    if not fam.pitch_known and lines_from >= 1:
        why.append("行の高さの比が分からない書体で、複数の行")
    return ex, ey, why


def resolve(wb, spec: dict, used: dict) -> dict:
    """○の位置：{id, sheet, row, col, text, ellipse（シートの左上からの pt）, label, notes}。書き込み・証明・読み戻しで同じ決め方。"""
    if getattr(wb, "kind", "xlsx") == "xls":
        raise CircleError("古い Excel（.xls）のシートには、○の図形を足す書き方をまだ作っていません")
    sh, r, c, text, a, b = find(wb, spec)
    lay = layout(wb, sh, r, c, text)
    gl, notes = lay.glyphs, list(lay.notes)
    flat = gl  # text の1字ごとの箱（改行も、幅のない箱として入る）
    if len(flat) != len(text):
        raise CircleError("セルの字の位置を、字ごとに決められません")
    mine = [g for g in flat[a:b] if not g.text.isspace()]
    if not mine:
        raise CircleError(f"「{spec['circle']}」の字の位置が分かりません")
    ids = {id(g) for g in mine}
    others = [g for g in flat if id(g) not in ids and not g.text.isspace()]
    rect = (min(g.x0 for g in mine), min(g.y0 for g in mine), max(g.x1 for g in mine), max(g.y1 for g in mine))
    size = max(g.size for g in mine)
    pad = spec.get("pad")
    if pad is not None and not (isinstance(pad, (int, float)) and not isinstance(pad, bool) and 0 <= pad <= 3 * size):
        raise CircleError(f"pad は、字と○のすき間（pt、0〜{3 * size:.0f}）で指定してください")
    style = spec.get("style", "ellipse")
    if style not in ("ellipse", "border"):
        raise CircleError('style は "ellipse"（楕円の○。既定）か "border"（字を囲む四角の図形。○でなくてよいと本人が確かめたとき）です')
    if style == "border":
        g, hit = _border(rect, size, others, None, pad)
        if g is None:
            raise CircleError(f"「{spec['circle']}」を四角で囲むと、同じセルのほかの字" + (f"「{hit.text}」" if hit else "") +
                              "に触れます（字の間が詰まっています）")
        # 楕円と同じ形（中心・半径）で持つ：図形の位置と大きさは、四角の外枠と同じ
        e = ((g[0] + g[2]) / 2, (g[1] + g[3]) / 2, (g[2] - g[0]) / 2, (g[3] - g[1]) / 2)
    else:
        e, hit = shape(rect, size, others, None, pad)
        if e is None:
            raise CircleError(f"「{spec['circle']}」を○で囲むと、同じセルのほかの字" + (f"「{hit.text}」" if hit else "") +
                              "に触れます。pad（字と○のすき間、pt）を小さくするか、○でなくてよいと本人が確かめたら"
                              " \"style\": \"border\"（字を囲む四角の図形）にしてください")
    # 見積もりの誤差：囲む字の幅の半分を超えるなら、囲みがほかの字にずれるおそれがあるので描かない
    ex, ey, why = error_estimate(wb, sh, lay, text, a, b, e[0] - e[2])
    half_w = 0.5 * sum(g.x1 - g.x0 for g in mine) / len(mine)
    mm = 25.4 / 72
    if ex > half_w or ey > 0.5 * size:
        over = (f"横 約 {ex * mm:.1f} mm（字の幅の半分 {half_w * mm:.1f} mm）" if ex > half_w
                else f"縦 約 {ey * mm:.1f} mm（字の高さの半分 {0.5 * size * mm:.1f} mm）")
        raise CircleError(f"位置を確かめられない：「{spec['circle']}」（{sh.name}!{addr(r, c)}）の位置の見積もりの誤差が{over}を超えます"
                          + (f"（{'・'.join(why)}）" if why else "") + "。囲みがほかの字にずれるおそれがあるので、描いていません。"
                          "字の前に「○」を足す書き方（old/new。本人に確かめてから）を使ってください"
                          + ("（style \"border\" も同じ見積もりで置くので、変わりません）" if style == "ellipse" else ""))
    notes.append(f"位置の見積もりの誤差 約 {max(ex, ey) * mm:.1f} mm 以内（目安）")
    if lay.hal in ("center", "right"):
        # 中央・右揃えの字の位置はセルの幅で決まる。列の幅の測り方が Excel と違うソフト（LibreOffice で 1 割半ほど狭く
        # 測られた様式があった）では、その分ずれる。位置は Excel の列の幅で決めている
        shift = (0.5 if lay.hal == "center" else 1.0) * 0.15 * lay.w
        if shift > half_w:
            notes.append(f"中央・右揃えのセル：列の幅の測り方が Excel と違うソフト（LibreOffice など）では、約 {shift * mm:.1f} mm "
                         "ずれて見えることがあります（位置は Excel の列の幅で決めています）")
    from ..pdf.circle import ellipses_meet

    def outer(x):  # 重なりを比べる楕円：四角は、外接する楕円
        return (x[0], x[1], x[2] * 2 ** 0.5, x[3] * 2 ** 0.5) if style == "border" else x

    for _k, e2, sn in used.get("all", []):
        if sn == sh.name and ellipses_meet(outer(e), e2):
            raise CircleError(f"「{spec['circle']}」の囲みが、ほかの囲みと重なります")
    k = 1 + sum(1 for x in used.get("all", []) if x[2] == sh.name) + _existing(wb, sh)
    used.setdefault("all", []).append((k, outer(e), sh.name))
    label = spec.get("near") or _row_label(wb, sh, r, c) or text[:30]
    return {"id": f"{sh.name}!○{k}", "sheet": sh.name, "row": r, "col": c, "cell": addr(r, c), "text": "".join(g.text for g in mine),
            "ellipse": [round(v, 3) for v in e], "label": label, "context": text[:60], "notes": notes, "size": size,
            "style": style, "around": _around(text, a, b), "error_mm": round(max(ex, ey) * mm, 2)}


def _row_label(wb, sh, r: int, c: int) -> str:
    """同じ行の左の、字のあるセル（見出し）。"""
    for k in range(c - 1, 0, -1):
        t = _text(wb, sh.cells.get((r, k))).strip()
        if t:
            return re.sub(r"\s+", " ", t)[:30]
    return ""


def _existing(wb, sh) -> int:
    """シートの図形に、すでにある とどくの○の数（番号を重ねないため）。"""
    part = drawing_part(wb, sh)
    if part is None:
        return 0
    return len(re.findall(rb'name="' + TAG.encode() + rb" ", wb.zip.read(part)))


# ---------------------------------------------------------------------- 書く
def drawing_part(wb, sh) -> str | None:
    for typ, t in wb.rels(sh.path).values():
        if typ == "drawing" and wb.zip.has(t):
            return t
    return None


def _anchor_cell(wb, sh, x: float, y: float) -> tuple[int, float, int, float]:
    """シートの左上からの位置 (x, y) pt → (列, 列の中の pt, 行, 行の中の pt)。"""
    c, acc = 1, 0.0
    while c < 16384 and acc + col_pt(wb, sh, c) <= x:
        acc += col_pt(wb, sh, c)
        c += 1
    r, accy = 1, 0.0
    while r < 1048576 and accy + row_pt(sh, r) <= y:
        accy += row_pt(sh, r)
        r += 1
    return c, x - acc, r, y - accy


def anchor_xml(wb, sh, sd: dict, shape_id: int, prefix: str = "xdr:", a: str = "a:") -> bytes:
    cx, cy, ra, rb = sd["ellipse"]
    lw = max(0.6, min(1.2, 0.07 * sd["size"]))
    x, y = cx - ra, cy - rb
    col, cof, row, rof = _anchor_cell(wb, sh, x, y)
    w, h = 2 * ra, 2 * rb
    xns = "" if prefix == "xdr:" and a == "a:" else ""
    desc = (f"四角で囲んだ字「{sd['text']}」（{sd['cell']}）" if sd.get("style") == "border"
            else f"○で囲んだ字「{sd['text']}」（{sd['cell']}）")
    x_ = prefix
    return (f"<{x_}oneCellAnchor><{x_}from><{x_}col>{col - 1}</{x_}col><{x_}colOff>{round(cof * EMU)}</{x_}colOff>"
            f"<{x_}row>{row - 1}</{x_}row><{x_}rowOff>{round(rof * EMU)}</{x_}rowOff></{x_}from>"
            f'<{x_}ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/>'
            f'<{x_}sp macro="" textlink=""><{x_}nvSpPr><{x_}cNvPr id="{shape_id}" name="{TAG} {_esc(sd["id"])}" descr="{_esc(desc)}"/>'
            f"<{x_}cNvSpPr/></{x_}nvSpPr><{x_}spPr><{a}xfrm><{a}off x=\"0\" y=\"0\"/><{a}ext cx=\"{round(w * EMU)}\" cy=\"{round(h * EMU)}\"/>"
            f'</{a}xfrm><{a}prstGeom prst="{"rect" if sd.get("style") == "border" else "ellipse"}"><{a}avLst/></{a}prstGeom><{a}noFill/><{a}ln w="{round(lw * EMU)}">'
            f'<{a}solidFill><{a}srgbClr val="000000"/></{a}solidFill></{a}ln></{x_}spPr></{x_}sp><{x_}clientData/></{x_}oneCellAnchor>'
            + xns).encode("utf-8")


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _prefixes(xml: bytes) -> tuple[str, str | None]:
    """図形の部品の、spreadsheetDrawing と drawingml の名前空間の接頭辞（既定の名前空間なら ""）。"""
    def pre(ns: str) -> str | None:
        m = re.search(rb'xmlns(?::([\w.-]+))?="' + re.escape(ns.encode()) + rb'"', xml[:4000])
        if m is None:
            return None
        return (m[1].decode() + ":") if m[1] else ""
    xp = pre(NS_X)
    return ("xdr:" if xp is None else xp), pre(NS_A)


def plan(wb, circles: list[dict]) -> dict[str, list[tuple[int, int, bytes]]]:
    """書き換える部品 → [(位置, 消すバイト数, 入れるバイト)]。足す部品は (0, 0, 中身)。"""
    out: dict[str, list] = {}
    new_parts: dict[str, bytes] = {}
    by_sheet: dict[str, list[dict]] = {}
    for sd in circles:
        by_sheet.setdefault(sd["sheet"], []).append(sd)
    names = set(wb.zip.names())
    for sn, sds in by_sheet.items():
        sh = wb.sheet(sn)
        part = drawing_part(wb, sh)
        if part is not None:
            xml = wb.zip.read(part)
            xp, ap = _prefixes(xml)
            ids = [int(x) for x in re.findall(rb'<(?:\w+:)?cNvPr\b[^>]*?\bid="(\d+)"', xml)]
            nid = max(ids, default=0) + 1
            body = b""
            for sd in sds:
                body += anchor_xml(wb, sh, sd, nid, xp, "a:" if ap is None else ap)
                nid += 1
            if ap is None:
                # drawingml の名前空間が宣言されていない部品：足す図形の要素で宣言する
                body = body.replace(b"<" + xp.encode() + b"sp ", b"<" + xp.encode() + b'sp xmlns:a="' + NS_A.encode() + b'" ')
            end = xml.rfind(b"</" + xp.encode() + b"wsDr>")
            if end < 0:
                raise CircleError(f"{part}: 図形の部品の終わりが見つかりません")
            out.setdefault(part, []).append((end, 0, body))
            continue
        # 図形の部品がない：部品を足し、シートから関係を張る
        k = 1
        while f"xl/drawings/drawing{k}.xml" in names or f"xl/drawings/drawing{k}.xml" in new_parts:
            k += 1
        dpart = f"xl/drawings/drawing{k}.xml"
        body = b"".join(anchor_xml(wb, sh, sd, i + 2) for i, sd in enumerate(sds))
        new_parts[dpart] = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                            b'<xdr:wsDr xmlns:xdr="' + NS_X.encode() + b'" xmlns:a="' + NS_A.encode() + b'">' + body + b"</xdr:wsDr>")
        d, base = posixpath.split(sh.path)
        rels = posixpath.join(d, "_rels", base + ".rels")
        target = posixpath.relpath(dpart, d)
        if wb.zip.has(rels):
            rx = wb.zip.read(rels)
            used_ids = set(re.findall(rb'Id="([^"]+)"', rx))
            n = 1
            while f"rId{n}".encode() in used_ids:
                n += 1
            rid = f"rId{n}"
            rel = f'<Relationship Id="{rid}" Type="{REL_DRAWING}" Target="{target}"/>'.encode()
            out.setdefault(rels, []).append((rx.rfind(b"</Relationships>"), 0, rel))
        else:
            rid = "rId1"
            new_parts[rels] = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                               b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                               + f'<Relationship Id="{rid}" Type="{REL_DRAWING}" Target="{target}"/>'.encode() + b"</Relationships>")
        sx = wb.zip.read(sh.path)
        pos, tag = _drawing_slot(sx, rid)
        out.setdefault(sh.path, []).append((pos, 0, tag))
        ct = wb.zip.read("[Content_Types].xml")
        ov = f'<Override PartName="/{dpart}" ContentType="{CT_DRAWING}"/>'.encode()
        out.setdefault("[Content_Types].xml", []).append((ct.rfind(b"</Types>"), 0, ov))
    for p, data in new_parts.items():
        out[p] = [(0, 0, data)]
        out[p + "\0new"] = []
    return out


# シートの XML で、<drawing> より後ろに来る要素（ECMA-376 の CT_Worksheet の順）
_AFTER_DRAWING = ("legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems",
                  "tableParts", "extLst")


def _drawing_slot(sx: bytes, rid: str) -> tuple[int, bytes]:
    m = re.match(rb"(?:<\?xml[^>]*\?>\s*)?<(\w+:)?worksheet\b[^>]*>", sx)
    if m is None:
        raise CircleError("シートの XML の始まりが分かりません")
    pre = (m[1] or b"").decode()
    rpre = re.search(rb'xmlns:([\w.-]+)="' + NS_R.encode() + rb'"', m[0])
    if rpre is not None:
        tag = f'<{pre}drawing {rpre[1].decode()}:id="{rid}"/>'.encode()
    else:
        tag = f'<{pre}drawing xmlns:r="{NS_R}" r:id="{rid}"/>'.encode()
    for name in _AFTER_DRAWING:
        k = re.search(rb"<" + re.escape(pre.encode()) + name.encode() + rb"\b", sx)
        if k is not None:
            return k.start(), tag
    end = sx.rfind(f"</{pre}worksheet>".encode())
    if end < 0:
        raise CircleError("シートの XML の終わりが分かりません")
    return end, tag


def apply(wb, circles: list[dict]) -> bytes:
    """○を書いたブックのバイト列（ほかの部品はバイト一致）。"""
    edits = plan(wb, circles)
    new = {p.split("\0")[0] for p in edits if p.endswith("\0new")}
    like = wb.sheets[0].path
    for part, ops_ in edits.items():
        if part.endswith("\0new"):
            continue
        if part in new:
            wb.zip.add(part, ops_[0][2], like)
            continue
        data = wb.zip.read(part)
        for pos, cut, ins in sorted(ops_, key=lambda x: -x[0]):
            data = data[:pos] + ins + data[pos + cut:]
        wb.zip.replace(part, data)
    blob = wb.zip.build()
    for e in wb.zip.entries:
        e._replacement = None
    if getattr(wb.zip, "added", None):
        wb.zip.drop_added()
    return blob


# ---------------------------------------------------------------------- 証明と読み戻し
def prove(before: bytes, after: bytes, name: str, specs: list[tuple[str, dict]]) -> list[str]:
    """○の書き込みの証明：元のブックで同じ決め方で○を決め直し、変わった部品がそれぞれ、決め直した所に決め直したバイトを
    入れただけか（ほかの部品はバイト一致か）を確かめる。想定外を返す。"""
    from ..zipraw import RawZip
    from .workbook import Workbook

    wb = Workbook(before, name)
    used: dict = {}
    out = []
    circles = []
    for fid, spec in specs:
        try:
            sd = resolve(wb, spec, used)
        except CircleError as e:
            return [f"○ {fid} の所が、元のブックで決まりません：{e}"]
        if sd["id"] != fid:
            out.append(f"○ {fid} が、決め直した所（{sd['id']}）と違います")
        circles.append(sd)
    edits = plan(wb, circles)
    za, zb = RawZip(before), RawZip(after)
    na, nb = za.names(), zb.names()
    new = [p.split("\0")[0] for p in edits if p.endswith("\0new")]
    if nb[:len(na)] != na or sorted(nb[len(na):]) != sorted(new):
        out.append(f"ZIP のエントリの並びが、○の部品を足すほかにも変わっています：{sorted(set(na) ^ set(nb))}")
    for n in nb:
        want = None
        if n in edits:
            if n in new:
                want = edits[n][0][2]
            else:
                want = za.read(n)
                for pos, cut, ins in sorted(edits[n], key=lambda x: -x[0]):
                    want = want[:pos] + ins + want[pos + cut:]
        elif n in na:
            ea, eb = za.get(n), zb.get(n)
            if (ea.local, ea.data, ea.descriptor) != (eb.local, eb.data, eb.descriptor):
                out.append(f"{n}: ○を描くのに変わってよい部品ではありません")
            continue
        if want is not None and zb.read(n) != want:
            out.append(f"{n}: ○を描くために足したバイトのほかにも変わっています")
    return out


def read_back(blob: bytes, name: str, circled: list[dict]) -> list[dict]:
    """書いた後のブックを開き直し、名前の付いた○の図形が、セルの字（書いた後のブックの書式で見積もった位置）のうち、
    言った字だけを囲んでいるかを確かめる。"""
    from .workbook import Workbook

    wb = Workbook(blob, name)
    out = []
    for c in circled:
        sh = wb.sheet(c["sheet"])
        part = drawing_part(wb, sh)
        xml = wb.zip.read(part) if part else b""
        m = re.search(rb"<(?:\w+:)?oneCellAnchor>((?:(?!</(?:\w+:)?oneCellAnchor>).)*?name=\"" + re.escape(f"{TAG} {_esc(c['id'])}".encode())
                      + rb"\".*?)</(?:\w+:)?oneCellAnchor>", xml, re.S)
        if m is None:
            out.append({"id": c["id"], "written": f"○「{c['text']}」", "read_back": "書いた後のブックに、この○の図形が見つかりません"})
            continue
        g = m[1]
        num = lambda t: int(re.search(rb"<(?:\w+:)?" + t + rb">(-?\d+)<", g)[1])  # noqa: E731
        col, cof, row, rof = num(b"col") + 1, num(b"colOff") / EMU, num(b"row") + 1, num(b"rowOff") / EMU
        ext = re.search(rb'<(?:\w+:)?ext cx="(\d+)" cy="(\d+)"', g)
        w, h = int(ext[1]) / EMU, int(ext[2]) / EMU
        x = sum(col_pt(wb, sh, k) for k in range(1, col)) + cof
        y = sum(row_pt(sh, k) for k in range(1, row)) + rof
        e = (x + w / 2, y + h / 2, w / 2, h / 2)
        r_, c_ = parse_cell(c["cell"])
        text = _text(wb, sh.cells.get((r_, c_)))
        try:
            gl, _n = glyphs(wb, sh, r_, c_, text)
        except CircleError as ex:
            out.append({"id": c["id"], "written": f"○「{c['text']}」", "read_back": str(ex)})
            continue
        rect_shape = re.search(rb'prstGeom prst="rect"', g) is not None
        if rect_shape != (c.get("style") == "border"):
            out.append({"id": c["id"], "written": f"○「{c['text']}」",
                        "read_back": f"書いた後の図形の形が{'四角' if rect_shape else '楕円'}で、書いた形と違います"})
            continue
        if rect_shape:
            box = (x, y, x + w, y + h)
            inner = "".join(x_.text for x_ in gl if not x_.text.isspace() and _encloses("border", box, x_))
            touch = [x_ for x_ in gl if not x_.text.isspace() and not _encloses("border", box, x_) and _hits("border", box, x_)]
        else:
            inner = "".join(x_.text for x_ in gl if not x_.text.isspace() and inside(e, (x_.x0, x_.y0, x_.x1, x_.y1)))
            touch = [x_ for x_ in gl if not x_.text.isspace() and not inside(e, (x_.x0, x_.y0, x_.x1, x_.y1)) and touches(e, face(x_))]
        if inner != re.sub(r"\s", "", c["text"]) or touch:
            out.append({"id": c["id"], "written": f"○「{c['text']}」",
                        "read_back": f"囲んでいる字「{inner}」" + (f"・触れている字「{''.join(t.text for t in touch)}」" if touch else "")})
    return out
