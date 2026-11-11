"""PowerPoint（.pptx）：書いたことで段落の行が折り返すかの見積もり（目安。wrapcheck）。

書いた段落（字が変わった段落）ごとに、書く前と書いた後の字を同じ幅で並べて、行の数を比べる（docx.wrap と同じ）。
幅と字の大きさは、書いてあるものを、PresentationML の受け継ぎの順（ECMA-376 第 1 部 19.3・21.1.2）でたどって決める：
- 図形（p:sp）の段落：図形の幅（p:spPr の a:xfrm の a:ext の cx。なければ、同じプレースホルダーの、レイアウト → マスターの図形）に、
  グループの拡大（a:ext ÷ a:chExt）を掛け、字の枠の左右の余白（a:bodyPr の lIns・rIns。既定 91440 EMU）を引いた幅。縦書き
  （vert）は高さ（cy）と上下の余白（tIns・bIns。既定 45720 EMU）。段組み（numCol・spcCol）は段の数で分ける。折り返さない設定
  （wrap="none"）の図形と、幅がどこにも書かれていない図形は、見積もらない
- 表のマスの段落：列の幅（a:tblGrid の a:gridCol の w。結合 gridSpan の分）からマスの左右の余白（a:tcPr の marL・marR。
  既定 91440 EMU）を引いた幅。縦書きのマスは見積もらない
- 段落の字下げ（a:pPr の marL・marR・indent）と字の大きさ（a:rPr の sz）は、書かれていなければ、段落のレベル（lvl）の書式を、
  図形の a:lstStyle → レイアウト・マスターの同じプレースホルダーの a:lstStyle → マスターの p:txStyles（タイトルは titleStyle、
  ほかのプレースホルダーは bodyStyle、そのほかは otherStyle。ノートは notesStyle）→ presentation.xml の p:defaultTextStyle の順に
  たどる。どこにもない字の大きさは、仮に 18 ポイント（PowerPoint の既定）と置き、見積もりの元にそう書く
- 字を縮めて収める設定（a:normAutofit の fontScale）は、いまの縮め方を字の大きさに掛ける。表のスタイルの字の大きさは見ない
"""

from __future__ import annotations

from ..docx.document import _join
from ..wrapcheck import place_name, say, wraps
from .document import rels_of

EMU = 12700.0  # 1 ポイントの EMU
_INS_X, _INS_Y = 91440, 45720  # 字の枠の余白の既定（左右・上下。ECMA-376 第 1 部 21.1.2.1.1）
_DEFAULT_SZ = 1800  # 字の大きさがどこにも書かれていないときに仮に置く値（1/100 ポイント。PowerPoint の既定）
_PARENT = {"sld": "/slideLayout", "sldLayout": "/slideMaster", "notes": "/notesMaster"}
_FAMILY = {"ctrTitle": "title", "subTitle": "body", "obj": "body"}


def _int(node, name: str) -> int | None:
    v = node.get(name) if node is not None else None
    try:
        return int(v) if v is not None else None
    except ValueError:
        return None


def _ph(sp):
    nv = sp.find("nvSpPr")
    pr = nv.find("nvPr") if nv is not None else None
    return pr.find("ph") if pr is not None else None


class _Geom:
    def __init__(self, doc):
        self.doc = doc
        self.where: dict[int, tuple] = {}  # id(a:p) → ("sp", 図形, 拡大, 部品) / ("tc", マス, 列の幅, 列, 拡大, 部品)
        self.done: set[str] = set()

    # ------------------------------------------------------------------ 段落 → 入れ物
    def scan(self, part: str) -> None:
        if part in self.done:
            return
        self.done.add(part)
        top = self._top(part)
        csld = top.find("cSld") if top is not None else None
        tree = csld.find("spTree") if csld is not None else None
        if tree is not None:
            self._walk(tree, 1.0, part)

    def _top(self, part: str):
        root = self.doc.load(part) if part else None
        return root.kids[0] if root is not None and root.kids else None

    def _walk(self, node, scale: float, part: str) -> None:
        for ch in node.kids:
            loc = ch.local
            if loc == "AlternateContent":
                choice = ch.find("Choice")
                if choice is not None:
                    self._walk(choice, scale, part)
            elif loc == "grpSp":
                pr = ch.find("grpSpPr")
                xf = pr.find("xfrm") if pr is not None else None
                a, b = (_int(xf.find("ext"), "cx"), _int(xf.find("chExt"), "cx")) if xf is not None else (None, None)
                self._walk(ch, scale * (a / b if a and b else 1.0), part)
            elif loc == "sp":
                body = ch.find("txBody")
                for p in (body.findall("p") if body is not None else []):
                    self.where[id(p)] = ("sp", ch, scale, part)
            elif loc == "graphicFrame":
                for tbl in ch.iter("tbl"):
                    g = tbl.find("tblGrid")
                    grid = [_int(c, "w") for c in g.findall("gridCol")] if g is not None else []
                    for tr in tbl.findall("tr"):
                        for col, tc in enumerate(tr.findall("tc")):
                            body = tc.find("txBody")
                            for p in (body.findall("p") if body is not None else []):
                                self.where[id(p)] = ("tc", tc, grid, col, scale, part)

    # ------------------------------------------------------------------ 受け継ぎ
    def parent(self, part: str) -> str:
        top = self._top(part)
        want = _PARENT.get(top.local if top is not None else "")
        if not want:
            return ""
        for typ, target in rels_of(self.doc.zip, part).values():
            if typ.endswith(want) and target:
                name = _join(part, target)
                return name if self.doc.zip.has(name) else ""
        return ""

    def chain(self, sp, part: str) -> tuple[list, str, str]:
        """図形と、同じプレースホルダーのレイアウト・マスターの図形。いちばん上の部品（マスター）と、プレースホルダーの種類。"""
        out, ph = [sp], _ph(sp)
        kind = (ph.get("type") or "obj") if ph is not None else ""
        up = self.parent(part)
        top_part = part
        by_idx = True  # スライド → レイアウトは番号（idx）、レイアウト → マスターは種類（type）で合わせる
        while up:
            top_part = up
            if ph is not None:
                found = self._match(ph, up, by_idx)
                if found is not None:
                    out.append(found)
                    ph = _ph(found)
                else:
                    ph = None
            by_idx = False
            up = self.parent(up)
        return out, top_part, kind

    def _match(self, ph, part: str, by_idx: bool):
        top = self._top(part)
        cands = [(sp, _ph(sp)) for sp in (top.iter("sp") if top is not None else [])]
        cands = [(sp, q) for sp, q in cands if q is not None]
        typ = ph.get("type") or "obj"
        tests = [lambda q: (q.get("type") or "obj") == typ, lambda q: (q.get("type") or "obj") == _FAMILY.get(typ)]
        idx = ph.get("idx")
        if idx is not None:
            same = lambda q: q.get("idx") == idx  # noqa: E731
            tests = [same] + tests if by_idx else tests + [same]
        for t in tests:
            for sp, q in cands:
                if t(q):
                    return sp
        return None

    def levels(self, chain: list, top_part: str, kind: str, lvl: int) -> list:
        """段落のレベル lvl（1〜9）の書式（a:lvlNpPr）を、近い順に。"""
        name = f"lvl{lvl}pPr"
        out = []
        for sp in chain:
            body = sp.find("txBody")
            ls = body.find("lstStyle") if body is not None else None
            n = ls.find(name) if ls is not None else None
            if n is not None:
                out.append(n)
        top = self._top(top_part)
        if top is not None:
            style = None
            if top.local == "notesMaster":
                style = top.find("notesStyle")
            elif top.local == "sldMaster":
                tx = top.find("txStyles")
                which = ("titleStyle" if kind in ("title", "ctrTitle") else
                         "otherStyle" if kind in ("", "dt", "ftr", "sldNum", "hdr") else "bodyStyle")
                style = tx.find(which) if tx is not None else None
            n = style.find(name) if style is not None else None
            if n is not None:
                out.append(n)
        pres = self._top(self.doc.main)
        dts = pres.find("defaultTextStyle") if pres is not None else None
        n = dts.find(name) if dts is not None else None
        if n is not None:
            out.append(n)
        return out


def _first(nodes: list, name: str) -> int | None:
    for n in nodes:
        v = _int(n, name)
        if v is not None:
            return v
    return None


def _measure(g: _Geom, p) -> dict | None:
    """段落の幅（pt）・字下げ・字ごとの大きさ。見積もれなければ None。"""
    got = g.where.get(id(p.node))
    if got is None:
        return None
    if got[0] == "sp":
        _k, sp, scale, part = got
        chain, top_part, kind = g.chain(sp, part)
        bodies = [b for b in (s.find("txBody") for s in chain) if b is not None]
        bprs = [b for b in (x.find("bodyPr") for x in bodies) if b is not None]
        attr = lambda name: next((b.get(name) for b in bprs if b.get(name) is not None), None)  # noqa: E731
        if attr("wrap") == "none":
            return None
        ext = next((e for e in ((s.find("spPr").find("xfrm").find("ext") if s.find("spPr") is not None
                                 and s.find("spPr").find("xfrm") is not None else None) for s in chain) if e is not None), None)
        upright = (attr("vert") or "horz") != "horz"
        size = _int(ext, "cy" if upright else "cx")
        if not size:
            return None
        ins = [_first(bprs, n) for n in (("tIns", "bIns") if upright else ("lIns", "rIns"))]
        dflt = _INS_Y if upright else _INS_X
        width = size * scale - sum(dflt if v is None else v for v in ins)
        ncol = _first(bprs, "numCol") or 1
        if ncol > 1:
            width = (width - (ncol - 1) * (_first(bprs, "spcCol") or 0)) / ncol
        basis = "図形の字の枠の幅" + ("（縦書き：高さ）" if upright else "") + (f"（{ncol} 段の 1 段）" if ncol > 1 else "")
        fit = bprs[0].find("normAutofit") if bprs and bprs[0] is bodies[0].find("bodyPr") else None
        shrink = (_int(fit, "fontScale") or 100000) / 100000
    else:
        _k, tc, grid, col, scale, part = got
        pr = tc.find("tcPr")
        span = _int(tc, "gridSpan") or 1
        cols = grid[col:col + span]
        if len(cols) < span or any(not w for w in cols) or (pr is not None and (pr.get("vert") or "horz") != "horz"):
            return None
        pads = [_int(pr, n) for n in ("marL", "marR")]
        width = sum(cols) * scale - sum(_INS_X if v is None else v for v in pads)
        basis, shrink = "表のマスの幅", 1.0
        body = tc.find("txBody")
        ls = body.find("lstStyle") if body is not None else None
        top_part = part
        while g.parent(top_part):
            top_part = g.parent(top_part)
        chain, kind = [], ""
    ppr = p.node.find("pPr")
    lvl = min(max((_int(ppr, "lvl") or 0) + 1, 1), 9)
    levels = g.levels(chain, top_part, kind, lvl)
    if got[0] == "tc" and ls is not None and ls.find(f"lvl{lvl}pPr") is not None:
        levels.insert(0, ls.find(f"lvl{lvl}pPr"))
    here = ([ppr] if ppr is not None else []) + levels
    left, right, first = (_first(here, n) or 0 for n in ("marL", "marR", "indent"))
    inherited = _first([d for d in (n.find("defRPr") for n in levels) if d is not None], "sz")
    sizes: list[float] = []
    assumed = False
    for sg in p.segs:
        if sg.hidden:
            continue
        run = getattr(sg, "run", None)
        rpr = run.find("rPr") if run is not None else None
        sz = _int(rpr, "sz") or inherited
        if not sz:
            sz, assumed = _DEFAULT_SZ, True
        sizes += [sz / 100 * shrink] * len(sg.text)
    return {"avail": (width - left - right) / EMU, "first": first / EMU, "sizes": sizes, "basis": basis, "assumed": assumed}


def wrap_warnings(before, after, names: dict[int, list[str]] | None = None) -> list[str]:
    """書いた段落のうち、書いたことで行が増える見込みのものの警告。before・after は PptxDocument（書く前・書いた後）。
    names：段落の番号 → 書いた所の id（PowerPoint には書いた所の印を置けないので、書き手から受け取る）。"""
    from ..docx.document import iter_paras

    def paras(doc):
        return [p for _k, _p, blocks in doc.all_blocks() for p in iter_paras(blocks)]

    pa, pb = paras(after), paras(before)
    if len(pa) != len(pb):
        return []
    ga, gb = _Geom(after), _Geom(before)
    out = []
    for i, (a, b) in enumerate(zip(pa, pb)):
        if a.text == b.text:
            continue
        ga.scan(a.part)
        gb.scan(b.part)
        ma, mb = _measure(ga, a), _measure(gb, b)
        if ma is None or mb is None or ma["avail"] <= 0:
            continue
        sa, sb = ma["sizes"], mb["sizes"]
        size = max(sa) if sa else _DEFAULT_SZ / 100
        got = wraps(b.text, a.text, ma["avail"], lambda k, s=sb: s[k] if k < len(s) else size,
                    lambda k, s=sa: s[k] if k < len(s) else size, ma["first"])
        if got is None:
            continue
        out.append(say(place_name((names or {}).get(i) or []), a.text, got[0], got[1],
                       f"{ma['basis']} 約 {ma['avail']:.0f} ポイント・字の大きさ {size:g} ポイント"
                       + ("（字の大きさは書かれていないので、仮に PowerPoint の既定の 18 ポイント）" if ma["assumed"] else ""), got[2]))
    return out
