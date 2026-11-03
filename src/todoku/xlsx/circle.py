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
"""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass

from ..formula.parser import addr, parse_cell
from ..textwidth import cells, east_asian_context, fixed_width
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


def _mdw(wb) -> int:
    """既定の書体の数字の幅（ピクセル）：列の幅の単位。和文の書体・Calibri とも、数字はおよそ字の大きさの半分。"""
    size = (_styles(wb)["fonts"] or [("", 11.0)])[0][1]
    return max(5, round(size * 96 / 72 * 0.5))


def col_pt(wb, sh, c: int) -> float:
    if c in sh.hidden_cols:
        return 0.0
    w = sh.col_widths.get(c, sh.default_col_width if sh.default_col_width is not None else 8.43)
    mdw = _mdw(wb)
    px = int(((256 * w + int(128 / mdw)) / 256) * mdw)
    return px * 0.75


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


def glyphs(wb, sh, r: int, c: int, text: str) -> tuple[list[G], list[str]]:
    """セルの字の箱の見積もり（シートの左上からの pt）と、見積もりの確かさに関わる注意。決め方の分からないセルは CircleError。"""
    cell = sh.cells.get((r, c))
    style = cell.style if cell is not None else sh.style_of(r, c)
    st = _styles(wb)
    xf = st["xfs"][style] if 0 <= style < len(st["xfs"]) else {"h": "general", "v": "bottom", "wrap": False, "shrink": False,
                                                              "indent": 0, "rot": 0}
    name, size = _font(wb, style)
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
    pad = 2 * 0.75  # セルの縁から字までの余白（2 ピクセル）
    # 曖昧な幅の字は、セルの字に全角の字があるか、書体の名前が東アジアの字（ＭＳ ゴシックなど）のときだけ全角
    wide = east_asian_context(text) or bool(re.search(r"[^\x00-\x7f]", name or ""))
    lw = [sum(width(ch, size, wide) for ch in ln) for ln in lines]
    if xf["wrap"] and max(lw) > w - 2 * pad:
        raise CircleError("折り返して表示するセルで、字が自動で改行される（改行の位置が分からない）ので、字の位置を決められません")
    if xf["shrink"] and max(lw) > w - 2 * pad:
        raise CircleError("縮小して全体を表示するセルで、字の大きさが分からないので、字の位置を決められません")
    pitch = size * 1.2
    block = pitch * (len(lines) - 1) + size
    if xf["v"] == "top":
        top = y0 + pad
    elif xf["v"] == "center":
        top = y0 + (h - block) / 2
    else:
        top = y0 + h - pad - block
    numeric = cell is not None and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
    hal = xf["h"] if xf["h"] != "general" else ("right" if numeric else "left")
    if xf["indent"]:
        notes.append("字下げ（インデント）のあるセル")
    if any(not fixed_width(ch, wide) for ch in text if ch != "\n") or "P" in name.upper() or not re.search(r"[^\x00-\x7f]", name or "a"):
        # ラテン文字・キリル文字・ギリシャ文字・アラビア文字などは、書体で字ごとに幅が変わる（プロポーショナル）
        notes.append("半角の字・幅の変わる書体の字（ラテン文字・キリル文字など）を含む（字の幅は見積もり）")
    out: list[G] = []
    for k, (ln, wd) in enumerate(zip(lines, lw)):
        if hal == "left":
            x = x0 + pad + xf["indent"] * 3 * _mdw(wb) * 0.75
        elif hal == "center":
            x = x0 + (w - wd) / 2
        else:
            x = x0 + w - pad - wd
        y = top + k * pitch
        for ch in ln:
            cw = width(ch, size, wide)
            out.append(G(ch, x, y, x + cw, y + size, size))
            x += cw
        if k < len(lines) - 1:
            out.append(G("\n", x, y, x, y + size, size))
    return out, notes


def resolve(wb, spec: dict, used: dict) -> dict:
    """○の位置：{id, sheet, row, col, text, ellipse（シートの左上からの pt）, label, notes}。書き込み・証明・読み戻しで同じ決め方。"""
    if getattr(wb, "kind", "xlsx") == "xls":
        raise CircleError("古い Excel（.xls）のシートには、○の図形を足す書き方をまだ作っていません")
    sh, r, c, text, a, b = find(wb, spec)
    gl, notes = glyphs(wb, sh, r, c, text)
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
            "style": style, "around": _around(text, a, b)}


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
