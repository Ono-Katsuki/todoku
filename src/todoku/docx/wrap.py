"""Word（.docx）：書いたことで段落の行が折り返すかの見積もり（目安。wrapcheck）。

書いた段落（書いた所の印のしおりのある段落）ごとに、書く前と書いた後の字を、同じ幅で並べて行の数を比べる。幅は：
- 表のマスの段落：表の列の幅（w:tblGrid の w:gridCol を、マスの w:gridSpan の分だけ足す。なければ w:tcW）から、マスの左右の
  余白（w:tcMar → 表の w:tblCellMar → 表のスタイル（w:tblStyle と、その w:basedOn）の w:tblCellMar → 既定の表のスタイル
  （w:type="table" w:default="1"）の w:tblCellMar → どこにもなければ 108 twip と仮に置き、見積もりの元に「仮に」と書く。
  ECMA-376 第1部 17.4.42・17.7.6）と段落の字下げを引いた幅
- ほかの段落：その段落の区切りの w:sectPr のページの幅（w:pgSz）から左右の余白（w:pgMar）・段組みの分・段落の字下げを引いた幅
字の大きさは文字の要素の書式とスタイル（circle.run_size）。文書のグリッドで字を並べる設定（w:docGrid の linesAndChars・
snapToChars）は、全角の字の送りを、グリッドの字の送り（既定の字の大きさ + w:charSpace / 4096 pt）にそろえる。
字の囲み線（w:bdr。○の既定の書き方も）は、囲みの左右に、線の太さ（w:sz。1/8 pt）と字との間（w:space。pt）の分だけ幅を足す
（同じ囲み線の字が続く所は 1 つの囲み。ECMA-376 第1部 17.3.2.4）。囲み線を付けただけで字の変わらない段落も見る。
テキストボックスの中の段落は、幅が分からないので見ない。
"""

from __future__ import annotations

from ..wrapcheck import say, wraps
from .circle import _paras, _styles, run_size

_DEFAULT_MAR = 108  # 表のマスの左右の余白が、どこにも書かれていないときに仮に置く値（twip。Word の既定の表のスタイルの値）


def _table_styles(doc) -> dict:
    """styles.xml の表のスタイル：{"based": id → basedOn, "mar": id → {辺: twip}, "default": 既定の表のスタイルの id}。"""
    memo = getattr(doc, "_wrap_table_styles", None)
    if memo is not None:
        return memo
    out: dict = {"based": {}, "mar": {}, "default": ""}
    root = doc._load("word/styles.xml") if hasattr(doc, "_load") else None
    top = root.kids[0] if root is not None and root.kids else None
    for st in (top.findall("style") if top is not None else []):
        if st.get("type") != "table":
            continue
        sid = st.get("styleId") or ""
        b = st.find("basedOn")
        if b is not None:
            out["based"][sid] = b.get("val")
        if st.get("default") in ("1", "true", "on"):
            out["default"] = sid
        tp = st.find("tblPr")
        cm = tp.find("tblCellMar") if tp is not None else None
        if cm is not None:
            sides = {}
            for k in cm.kids:
                if k.get("type") in (None, "dxa") and k.get("w") not in (None, ""):
                    sides[{"start": "left", "end": "right"}.get(k.local, k.local)] = _tw(k.get("w"))
            out["mar"][sid] = sides
    try:
        doc._wrap_table_styles = out
    except AttributeError:
        pass
    return out


def _style_margin(doc, tbl, side: str) -> tuple[float | None, str]:
    """表のスタイルの階層から、マスの余白（twip）と、どこで決まったか。"""
    if doc is None:
        return None, ""
    ts = _table_styles(doc)
    tp = tbl.find("tblPr") if tbl is not None else None
    sref = tp.find("tblStyle") if tp is not None else None
    chain, sid, seen = [], (sref.get("val") if sref is not None else None), 0
    while sid and seen < 20:
        chain.append(sid)
        sid, seen = ts["based"].get(sid), seen + 1
    if ts["default"] and ts["default"] not in chain:
        chain.append(ts["default"])
    for sid in chain:
        v = ts["mar"].get(sid, {}).get(side)
        if v is not None:
            return v, f"表のスタイル「{sid}」"
    return None, ""


def _tw(v) -> float:
    try:
        return float(str(v))
    except (TypeError, ValueError):
        return 0.0


def _root(n):
    while n.parent is not None:
        n = n.parent
    return n


def _sect(p):
    """段落の区切りの w:sectPr（段落より後ろの最初のもの。なければ本文の最後のもの）。"""
    memo = [n for n in _root(p.node).iter("sectPr")]
    after = [n for n in memo if n.s >= p.node.s]
    return after[0] if after else (memo[-1] if memo else None)


def _ind(p, size: float) -> tuple[float, float, float]:
    """段落の (左の字下げ, 右の字下げ, 1 行目の字下げ（ぶら下げは負）)（pt）。段落の書式に直接あるものだけ。"""
    ppr = p.node.find("pPr")
    ind = ppr.find("ind") if ppr is not None else None
    if ind is None:
        return 0.0, 0.0, 0.0

    def val(*names, chars: str | None = None) -> float:
        if chars and ind.get(chars) not in (None, ""):
            return _tw(ind.get(chars)) / 100 * size
        for n in names:
            if ind.get(n) not in (None, ""):
                return _tw(ind.get(n)) / 20
        return 0.0

    left = val("left", "start", chars="leftChars")
    right = val("right", "end", chars="rightChars")
    first = val("firstLine", chars="firstLineChars") - val("hanging", chars="hangingChars")
    return left, right, first


def _cell_width(p, doc=None) -> tuple[float | None, str]:
    """表のマスの、字を置ける幅（pt。字下げは引く前）と、その元の説明。"""
    tc = p.cell.node if p.cell is not None else None
    if tc is None:
        return None, ""
    tr = tc.parent
    tbl = tr.parent if tr is not None else None
    tcpr = tc.find("tcPr")
    span = 1
    if tcpr is not None and tcpr.find("gridSpan") is not None:
        span = max(1, int(_tw(tcpr.find("gridSpan").get("val", "1"))))
    # 行の中で、このマスより前のマスの列の数（w:gridBefore も）
    idx = 0
    trpr = tr.find("trPr") if tr is not None else None
    if trpr is not None and trpr.find("gridBefore") is not None:
        idx += int(_tw(trpr.find("gridBefore").get("val", "0")))
    for k in tr.kids if tr is not None else []:
        if k is tc:
            break
        if k.local == "tc":
            kp = k.find("tcPr")
            gs = kp.find("gridSpan") if kp is not None else None
            idx += max(1, int(_tw(gs.get("val", "1")))) if gs is not None else 1
    width = None
    grid = tbl.find("tblGrid") if tbl is not None else None
    cols = [_tw(g.get("w")) for g in grid.findall("gridCol")] if grid is not None else []
    if cols and idx + span <= len(cols):
        width = sum(cols[idx:idx + span])
    elif tcpr is not None and tcpr.find("tcW") is not None and tcpr.find("tcW").get("type") in (None, "dxa"):
        width = _tw(tcpr.find("tcW").get("w"))
    if not width:
        return None, ""

    def margin(side: str) -> float | None:
        for holder in (tcpr.find("tcMar") if tcpr is not None else None,
                       (tbl.find("tblPr").find("tblCellMar") if tbl is not None and tbl.find("tblPr") is not None else None)):
            if holder is None:
                continue
            for name in ((side, "start") if side == "left" else (side, "end")):
                m = holder.find(name)
                if m is not None and m.get("type") in (None, "dxa"):
                    return _tw(m.get("w"))
        return None

    ml, mr = margin("left"), margin("right")
    notes = []
    if ml is None:
        ml, where = _style_margin(doc, tbl, "left")
        notes += [where] if where else []
    if mr is None:
        mr, where = _style_margin(doc, tbl, "right")
        notes += [where] if where else []
    assumed = ml is None or mr is None
    w = width - (_DEFAULT_MAR if ml is None else ml) - (_DEFAULT_MAR if mr is None else mr)
    basis = "表のマスの幅"
    if notes:
        basis += f"（余白は{'・'.join(dict.fromkeys(notes))}）"
    if assumed:
        basis += f"（余白の書かれていない辺は、仮に {_DEFAULT_MAR / 20:g} ポイント）"
    return w / 20, basis


def _page_width(p) -> tuple[float | None, str, object]:
    sp = _sect(p)
    if sp is None:
        return None, "", None
    sz, mar = sp.find("pgSz"), sp.find("pgMar")
    if sz is None or not _tw(sz.get("w")):
        return None, "", sp
    w = _tw(sz.get("w")) - (_tw(mar.get("left")) + _tw(mar.get("right")) + _tw(mar.get("gutter")) if mar is not None else 0)
    cols = sp.find("cols")
    n = int(_tw(cols.get("num"))) if cols is not None and cols.get("num") else 1
    if n > 1:
        w = (w - _tw(cols.get("space") or 720) * (n - 1)) / n
    return w / 20, "ページの幅（余白を引いた幅）", sp


def _grid_pitch(doc, sp) -> float | None:
    """文書のグリッドで字を並べる設定のときの、全角の字の送り（pt）。"""
    g = sp.find("docGrid") if sp is not None else None
    if g is None or g.get("type") not in ("linesAndChars", "snapToChars"):
        return None
    return _styles(doc)["default"] + _tw(g.get("charSpace")) / 4096


def _in_textbox(p) -> bool:
    n = p.node.parent
    while n is not None:
        if n.local in ("txbxContent", "textbox"):
            return True
        n = n.parent
    return False


def wrap_warnings(before, after, rename: dict | None = None) -> list[str]:
    """書いた段落のうち、書いたことで行が増える見込みのものの警告。before・after は DocxDocument（書く前・書いた後）。
    rename：書いた所の印（しおり）の名前 → 欄の id。"""
    from .window import _bookmarks

    pa = list(_paras(after))
    pb = list(_paras(before))
    if len(pa) != len(pb):
        return []
    index = {id(p.node): i for i, p in enumerate(pa)}
    written: dict[int, list[str]] = {}
    edges = {i: (_border_extra(after, a), _border_extra(before, b)) for i, (a, b) in enumerate(zip(pa, pb))}
    for i, (ea, eb) in edges.items():
        if ea != eb:
            written.setdefault(i, [])  # 囲み線が変わった段落（○の既定の書き方は、書いた所の印を置かない）
    for name, (q, _x, _y) in _bookmarks(after).items():
        i = index.get(id(q.node))
        if i is not None:
            written.setdefault(i, []).append(name)
    out = []
    for i, names in sorted(written.items()):
        a, b = pa[i], pb[i]
        ea, eb = edges[i]
        if _in_textbox(a) or (a.text == b.text and ea == eb):
            continue
        sizes_a = _sizes(after, a)
        sizes_b = _sizes(before, b)
        size = max(sizes_a) if sizes_a else 10.5
        left, right, first = _ind(a, size)
        width, basis = _cell_width(a, after)
        sp = None
        if width is None:
            width, basis, sp = _page_width(a)
        else:
            _w, _b, sp = _page_width(a)
        if width is None:
            continue
        avail = width - left - right
        pitch = _grid_pitch(after, sp)
        sa = _snap(sizes_a, a.text, pitch)
        sb = _snap(sizes_b, b.text, pitch)
        got = wraps(b.text, a.text, avail, lambda k, s=sb: s[k] if k < len(s) else size,
                    lambda k, s=sa: s[k] if k < len(s) else size, first,
                    extra_before=lambda k: eb[k] if k < len(eb) else 0.0, extra_after=lambda k: ea[k] if k < len(ea) else 0.0)
        if got is None:
            continue
        place = "・".join(dict.fromkeys((rename or {}).get(n, n) for n in names)) or "字を囲んだ段落"
        out.append(say(place, a.text, got[0], got[1],
                       f"{basis} 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント" + ("・文書のグリッド" if pitch else "")
                       + ("・字の囲み線" if any(ea) else ""), got[2]))
    return out


def _border_styles(doc) -> dict:
    """スタイル → (basedOn, 囲み線の要素)。"""
    memo = getattr(doc, "_wrap_border_styles", None)
    if memo is not None:
        return memo
    out = {}
    root = doc._load("word/styles.xml") if hasattr(doc, "_load") else None
    top = root.kids[0] if root is not None and root.kids else None
    for st in (top.findall("style") if top is not None else []):
        b, rp = st.find("basedOn"), st.find("rPr")
        out[st.get("styleId") or ""] = (b.get("val") if b is not None else None, rp.find("bdr") if rp is not None else None)
    try:
        doc._wrap_border_styles = out
    except AttributeError:
        pass
    return out


def _run_border(doc, p, run) -> tuple[float, str] | None:
    """文字の要素の囲み線：(左右それぞれに足す幅（pt）, 同じ囲みかを見る印)。文字の要素の書式 → 文字のスタイル → 段落のスタイル。"""
    rpr = run.find("rPr") if run is not None else None
    bd = rpr.find("bdr") if rpr is not None else None
    if bd is None:
        styles = _border_styles(doc)
        rs = rpr.find("rStyle") if rpr is not None else None
        for sid in (rs.get("val") if rs is not None else None, p.style or None):
            seen = 0
            while sid and bd is None and seen < 20:
                sid, bd = styles.get(sid, (None, None))
                seen += 1
            if bd is not None:
                break
    if bd is None or bd.get("val") in (None, "nil", "none"):
        return None
    return _tw(bd.get("sz")) / 8 + _tw(bd.get("space")), f'{bd.get("val")}/{bd.get("sz")}/{bd.get("space")}/{bd.get("color")}'


def _border_extra(doc, p) -> list[float]:
    """段落の字ごとの、囲み線で足す幅（pt）：同じ囲み線の字が続く所の、最初の字に左の分、最後の字に右の分。"""
    pos = p.positions()
    marks = [_run_border(doc, p, p.segs[pos[k][0]].run if k < len(pos) else None) for k in range(len(p.text))]
    out = [0.0] * len(marks)
    for k, m in enumerate(marks):
        if m is None:
            continue
        if k == 0 or marks[k - 1] is None or marks[k - 1][1] != m[1]:
            out[k] += m[0]
        if k + 1 == len(marks) or marks[k + 1] is None or marks[k + 1][1] != m[1]:
            out[k] += m[0]
    return out


def _sizes(doc, p) -> list[float]:
    pos = p.positions()
    out = []
    for k in range(len(p.text)):
        sg = p.segs[pos[k][0]] if k < len(pos) else None
        out.append(run_size(doc, p, sg.run if sg is not None else None))
    return out


def _snap(sizes: list[float], text: str, pitch: float | None) -> list[float]:
    """文書のグリッド：全角の字は、字の大きさの代わりにグリッドの送りで並ぶ（count_lines の字の大きさとして渡す）。"""
    if not pitch:
        return sizes
    from ..textwidth import cells, east_asian_context

    wide = east_asian_context(text)
    return [max(s, pitch) if k < len(text) and cells(text[k], wide) == 2 else s for k, s in enumerate(sizes)]
