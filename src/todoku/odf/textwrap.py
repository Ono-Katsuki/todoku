"""OpenDocument の文書（.odt）：書いたことで段落の行が折り返すかの見積もり（目安。wrapcheck）。

書いた段落（字が変わった段落：wrapcheck.written_places）ごとに、書く前と書いた後の字を同じ幅で並べて、行の数を比べる（docx.wrap と同じ）。
幅：表のマスの段落は、表の列（table:table-column の style:column-width。マスの table:number-columns-spanned の分）から
マスの左右の余白（fo:padding-left・fo:padding-right・fo:padding。なければ LibreOffice の既定 0.097cm）を引いた幅。
ほかの段落は、最初のページの体裁（style:page-layout-properties の fo:page-width・fo:margin-left・fo:margin-right）。
どちらも、段落のスタイルの fo:margin-left・fo:margin-right（字下げ）と fo:text-indent（1 行目）を引く。
字の大きさは、字のスタイル（text:span）→ 段落のスタイル → 既定のスタイルの fo:font-size（なければ 12 pt）。
"""

from __future__ import annotations

from ..wrapcheck import place_name, say, wraps, written_places
from .sheet import _columns, _pt

# マスの余白（fo:padding）が書かれていないときに仮に置く値：LibreOffice の既定の 0.097 cm（ODF の仕様では、書かれていない
# padding は 0。どちらで描くかはソフトによるので、見積もりの元に「仮に」と書く）
_DEFAULT_PAD = 0.097 * 28.3465
_DEFAULT_SIZE = 12.0  # 字の大きさが書かれていないときに仮に置く値（LibreOffice の既定。ODF の仕様の値ではない）


def _len(v: str | None) -> float | None:
    if not v:
        return None
    try:
        return _pt(v)
    except (ValueError, TypeError):
        return None


def _size(doc, p, run) -> float:
    st = doc.styles
    if run is not None and run is not p.el:
        v = _len(st.prop(p.part, "text", run.get("text:style-name"), "style:text-properties", "fo:font-size"))
        if v:
            return v
    v = st.prop(p.part, "paragraph", p.style or None, "style:text-properties", "fo:font-size")
    if v and v.endswith("%"):
        return _DEFAULT_SIZE * float(v[:-1]) / 100
    return _len(v) or _DEFAULT_SIZE


def _sizes(doc, p) -> list[float]:
    out = []
    for sg in p.segs:
        if sg.hidden:
            continue
        s = _size(doc, p, sg.run)
        out += [s] * len(sg.text)
    return out


def _indents(doc, p) -> tuple[float, float, float]:
    st = doc.styles
    get = lambda a: _len(st.prop(p.part, "paragraph", p.style or None, "style:paragraph-properties", a)) or 0.0  # noqa: E731
    return get("fo:margin-left"), get("fo:margin-right"), get("fo:text-indent")


def _cell_width(doc, p, facts: dict | None = None) -> float | None:
    c = p.cell
    if c is None or c.table is None or c.table.node is None:
        return None
    st = doc.styles
    widths = []
    for col in _columns(c.table.node):
        rep = min(int(col.get("table:number-columns-repeated", "1") or 1), 1000)
        w = _len(st.prop(p.part, "table-column", col.get("table:style-name"), "style:table-column-properties", "style:column-width"))
        widths += [w] * rep
    span = getattr(c, "span", 1) or 1
    if c.col + span > len(widths) or any(w is None for w in widths[c.col:c.col + span]):
        return None
    w = sum(widths[c.col:c.col + span])
    style = c.node.get("table:style-name") if c.node is not None else None
    pad = _len(st.prop(p.part, "table-cell", style, "style:table-cell-properties", "fo:padding"))
    pl = _len(st.prop(p.part, "table-cell", style, "style:table-cell-properties", "fo:padding-left"))
    pr = _len(st.prop(p.part, "table-cell", style, "style:table-cell-properties", "fo:padding-right"))
    dflt = _DEFAULT_PAD if pad is None else pad
    if facts is not None:
        facts["pad_assumed"] = pad is None and (pl is None or pr is None)
    return w - (dflt if pl is None else pl) - (dflt if pr is None else pr)


def _page_width(doc) -> float | None:
    part = doc.pkg.part_of("styles.xml")
    for sec in ("automatic-styles", "styles"):
        s = doc.pkg.section(part, sec)
        if s is None:
            continue
        for pl in s.findall("style:page-layout"):
            props = pl.find("style:page-layout-properties")
            if props is None:
                continue
            w = _len(props.get("fo:page-width"))
            if w:
                return w - (_len(props.get("fo:margin-left")) or 0.0) - (_len(props.get("fo:margin-right")) or 0.0)
    return None


def wrap_warnings(before, after) -> list[str]:
    """書いた段落のうち、書いたことで行が増える見込みのものの警告。before・after は OdtDocument（書く前・書いた後）。"""
    from ..docx.document import iter_paras

    def paras(doc):
        return [p for _k, _p, blocks in doc.all_blocks() for p in iter_paras(blocks)]

    pa, pb = paras(after), paras(before)
    if len(pa) != len(pb):
        return []
    from ..lazy.blocks import _control_ids

    out = []
    for i, (names, _at) in sorted(written_places(pa, pb, after.marks(), _control_ids(after)).items()):
        a, b = pa[i], pb[i]
        sa, sb = _sizes(after, a), _sizes(before, b)
        size = max(sa) if sa else _DEFAULT_SIZE
        left, right, first = _indents(after, a)
        facts: dict = {}
        width = _cell_width(after, a, facts)
        basis = "表のマスの幅" + ("（マスの余白は書かれていないので、仮に LibreOffice の既定の 0.097 cm）" if facts.get("pad_assumed") else "")
        if width is None:
            if a.cell is not None:
                continue
            width, basis = _page_width(after), "ページの幅（余白を引いた幅）"
        if not width:
            continue
        avail = width - left - right
        got = wraps(b.text, a.text, avail, lambda k, s=sb: s[k] if k < len(s) else size,
                    lambda k, s=sa: s[k] if k < len(s) else size, first)
        if got is None:
            continue
        out.append(say(place_name(names), a.text, got[0], got[1],
                       f"{basis} 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント", got[2]))
    return out
