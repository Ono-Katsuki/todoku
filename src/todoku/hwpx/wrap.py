"""HWPX：書いたことで段落の行が折り返すかの見積もり（目安。wrapcheck）。

書いた段落（字が変わった段落：wrapcheck.written_places）ごとに、書く前と書いた後の字を同じ幅で並べて、行の数を比べる（docx.wrap と同じ）。
幅（HWPUNIT＝1/7200 インチ＝1/100 ポイント。KS X 6101）：
- 入れ物の幅：表のマスの段落は、マスの幅（hp:cellSz の width）からマスの左右の余白（マスが自分の余白を持つ hasMargin="1" なら
  hp:cellMargin、そうでなければ表の hp:inMargin）を引いた幅。本文・ヘッダー・フッター・脚注の段落は、その節の用紙の幅
  （hp:pagePr の width）から左右の余白ととじしろ（hp:margin の left・right・gutter）を引き、段組み（hp:colPr の colCount・sameGap）
  の数で分けた幅。用紙が横向き（landscape が WIDELY でない）の節・図形の中・見出し（キャプション）・メモの段落は、入れ物の幅を決めない
- 한글が書く前の段落に残した行の配置の控え（hp:linesegarray の horzsize）があり、入れ物の幅を超えていなければ（入れ物の幅が
  決まらないときも）、その行の幅を使う：한글が実際に並べた幅で、段落の字下げも入っている
- 控えがなければ、入れ物の幅から段落の形（hh:paraPr の hh:margin の left・right）を引き、1 行目の字下げ（intent。負はぶら下げ）
  を見る。hp:switch の中では、単位が HWPUNIT の枝（hp:case）を読み、古い単位の枝（hp:default）は読まない
- どちらも決まらない段落は、見積もらない
字の大きさは、字の形（hh:charPr の height）に、その字の種類（ハングル・英字・漢字・かな・そのほか）の장평（hh:ratio）と
자간（hh:spacing）を掛けた幅で数える。
"""

from __future__ import annotations

from ..wrapcheck import place_name, say, wraps, written_places
from .package import L, kid, kids, top

HEADER = "Contents/header.xml"
_STOP = {"drawText", "caption", "hiddenComment", "fieldBegin", "masterPage"}  # 入れ物の幅を決めない所


def _int(el, name: str) -> int | None:
    try:
        return int(el.attrs.get(name, "")) if el is not None else None
    except ValueError:
        return None


def _script(ch: str) -> str:
    c = ord(ch)
    if 0x1100 <= c <= 0x11FF or 0x3130 <= c <= 0x318F or 0xAC00 <= c <= 0xD7AF:
        return "hangul"
    if c < 0x250:
        return "latin"
    if 0x4E00 <= c <= 0x9FFF or 0x3400 <= c <= 0x4DBF or 0xF900 <= c <= 0xFAFF:
        return "hanja"
    if 0x3040 <= c <= 0x30FF:
        return "japanese"
    return "other"


class _Shapes:
    """header.xml の字の形（大きさ・장평・자간）と段落の形（左右・1 行目の字下げ）。"""

    def __init__(self, doc):
        self.chars: dict[str, tuple[float, dict, dict]] = {}
        self.paras: dict[str, tuple[int, int, int]] = {}
        if not doc.pkg.has(HEADER):
            return
        for el in top(doc.pkg.tree(HEADER)).iter():
            loc = L(el)
            if loc == "charPr":
                h = _int(el, "height")
                if h and h > 0:
                    ratio, spacing = kid(el, "ratio"), kid(el, "spacing")
                    self.chars[el.attrs.get("id", "")] = (h / 100, dict(ratio.attrs) if ratio is not None else {},
                                                          dict(spacing.attrs) if spacing is not None else {})
            elif loc == "paraPr":
                m = kid(el, "margin")
                if m is None:
                    for sw in kids(el, "switch"):
                        for case in kids(sw, "case"):
                            m = m if m is not None else kid(case, "margin")
                if m is not None:
                    get = lambda n, m=m: _int(kid(m, n), "value") or 0  # noqa: E731
                    self.paras[el.attrs.get("id", "")] = (get("left"), get("right"), get("intent"))

    def sizes(self, p) -> tuple[list[float], float]:
        """字ごとの、幅を数える大きさ（pt。장평・자간を掛けた値）と、いちばん大きい字の大きさ（pt）。"""
        out: list[float] = []
        top = 0.0
        for sg in p.segs:
            if sg.hidden:
                continue
            run = getattr(sg, "run", None)
            got = self.chars.get(run.attrs.get("charPrIDRef", "")) if run is not None else None
            if got is None:
                out += [0.0] * len(sg.text)
                continue
            h, ratio, spacing = got
            top = max(top, h)
            for ch in sg.text:
                k = _script(ch)
                try:
                    out.append(h * (int(ratio.get(k, "100")) + int(spacing.get(k, "0"))) / 100)
                except ValueError:
                    out.append(h)
        return out, top


def _page(sec, cache: dict) -> float | None:
    if id(sec) not in cache:
        w = None
        pp = next((el for el in sec.iter() if L(el) == "pagePr"), None)
        width = _int(pp, "width")
        if width and pp.attrs.get("landscape", "WIDELY") == "WIDELY":
            m = kid(pp, "margin")
            w = width - sum(_int(m, n) or 0 for n in ("left", "right", "gutter"))
            col = next((el for el in sec.iter() if L(el) == "colPr"), None)
            n = max(_int(col, "colCount") or 1, 1)
            w = (w - (n - 1) * (_int(col, "sameGap") or 0)) / n
        cache[id(sec)] = w if w and w > 0 else None
    return cache[id(sec)]


def _box(p, cache: dict) -> tuple[float | None, str]:
    """段落の入れ物の幅（HWPUNIT。決まらなければ None）と、その名前。"""
    el = getattr(p.el, "parent", None)
    while el is not None:
        loc = L(el)
        if loc == "tc":
            w = _int(kid(el, "cellSz"), "width")
            if not w:
                return None, "表のマスの幅"
            tbl = getattr(getattr(el, "parent", None), "parent", None)
            m = kid(el, "cellMargin") if el.attrs.get("hasMargin") == "1" else kid(tbl, "inMargin") if tbl is not None else None
            return w - (_int(m, "left") or 0) - (_int(m, "right") or 0), "表のマスの幅"
        if loc in _STOP:
            return None, ""
        if loc == "sec":
            return _page(el, cache), "ページの幅（余白を引いた幅）"
        el = getattr(el, "parent", None)
    return None, ""


def wrap_warnings(before, after) -> list[str]:
    """書いた段落のうち、書いたことで行が増える見込みのものの警告。before・after は HwpxDocument（書く前・書いた後）。"""
    from ..docx.document import iter_paras

    def paras(doc):
        return [p for _k, _p, blocks in doc.all_blocks() for p in iter_paras(blocks)]

    pa, pb = paras(after), paras(before)
    if len(pa) != len(pb):
        return []
    from ..lazy.blocks import _control_ids

    written = written_places(pa, pb, after.marks(), _control_ids(after))
    if not written:
        return []
    sa_, sb_ = _Shapes(after), _Shapes(before)
    cache: dict = {}
    out = []
    for i, (names, _at) in sorted(written.items()):
        a, b = pa[i], pb[i]
        if getattr(b, "el", None) is None:
            continue
        (sa, size), (sb, _s) = sa_.sizes(a), sb_.sizes(b)
        if not size:
            continue
        box, basis = _box(b, cache)
        lsa = kid(b.el, "linesegarray")
        segs = [w for w in (_int(ls, "horzsize") for ls in kids(lsa, "lineseg")) if w and w > 0] if lsa is not None else []
        left, right, first = sb_.paras.get(b.el.attrs.get("paraPrIDRef", ""), (0, 0, 0))
        if segs and (box is None or max(segs) <= box):
            # 1 行目の控えは 1 行目の幅：2 行目からの幅は、控えが 2 行あればその幅、なければ 1 行目の字下げの分を戻した幅
            rest = max(segs[1:]) if len(segs) > 1 else segs[0] + first
            avail, indent = rest / 100, (rest - segs[0]) / 100
            basis = "한글が残した行の幅（行の配置の控え）"
        elif box:
            rest = box - left - right + min(first, 0)
            avail, indent = rest / 100, first / 100
        else:
            continue
        if avail <= 0:
            continue
        got = wraps(b.text, a.text, avail, lambda k, s=sb: (s[k] if k < len(s) else 0.0) or size,
                    lambda k, s=sa: (s[k] if k < len(s) else 0.0) or size, indent)
        if got is None:
            continue
        out.append(say(place_name(names), a.text, got[0], got[1],
                       f"{basis} 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント", got[2]))
    return out
