"""PDF：窓に欄として示していない所に、エージェントが見出しの字を目印にして書く。

PDF の字は消せないので、書くのは見出しの右（または下）の空いている所だけ。空いている所は、見出しの行のマスの縁・
同じ高さの次の字・縦の罫線の手前まで（下なら、次の行かマスの下の縁の手前まで）。書き込みと証明で同じ決め方を使い、
証明では、足した欄の位置が、目印から決め直した位置と同じであることを確かめる。
"""

from __future__ import annotations

import itertools
import re

from .content import page_content
from .layout import PageLayout


class AnchorError(ValueError):
    pass


def _lines(lay: PageLayout):
    for c in lay.cells:
        for ln in c.lines:
            if not ln.vertical:
                yield c, ln
    for ln in lay.free_lines:
        if not ln.vertical:
            yield None, ln


def resolve_write(doc, spec: dict, used: dict | None = None, layouts: dict | None = None,
                  taken_names: set | None = None) -> dict:
    """見出しの字で指した書き込み → 書く所（窓の欄と同じ形の辞書：id「p ページ.w 番号」・ページ・矩形・字の大きさ）。

    spec：{"after": 見出しの字（窓の字のとおり）, "value": 書く字, "page": ページ（省くと全ページから）,
    "where": "right"（既定。見出しの右）/ "below"（見出しの下）, "occurrence": 何番目か（読む順。1 から）}、
    または {"cover": 覆う字（伏せ字の「○○」など）, "value": 書く字, …}：その字の上を白い下地の欄で覆って書く（元の字は消さない）。
    見つからない・いくつもあって決まらない・空いている所がないときは、理由（と候補）を添えて断る。"""
    if spec.get("box") is not None:
        return _resolve_box(doc, spec, used, layouts, taken_names)
    cover = spec.get("cover")
    after = cover if cover is not None else spec.get("after")
    value = spec.get("value")
    if not isinstance(after, str) or not after.strip() or value is None:
        raise AnchorError('PDF で欄のない所に書くときは {"after": 見出しの字, "value": 書く字}（字を覆って書くときは '
                          '{"cover": 覆う字, "value": 書く字}）の形で指定してください')
    where = "cover" if cover is not None else spec.get("where", "right")
    if where not in ("right", "below", "cover"):
        raise AnchorError('where は "right"（見出しの右）か "below"（見出しの下）です')
    pages = doc.pages()
    page_no = spec.get("page")
    if page_no is not None and not (isinstance(page_no, int) and 1 <= page_no <= len(pages)):
        raise AnchorError(f"page は 1〜{len(pages)} で指定してください")
    layouts = layouts if layouts is not None else {}
    hits = []
    for n in ([page_no] if page_no else range(1, len(pages) + 1)):
        if n not in layouts:
            pg = pages[n - 1]
            layouts[n] = PageLayout(doc, pg, page_content(doc, pg), n)
        lay = layouts[n]
        for cell, ln in _lines(lay):
            for m in re.finditer(re.escape(after), ln.text):
                hits.append((n, round(ln.y0, 1), round(ln.x0, 1), m.start(), cell, ln, lay))
    hits.sort(key=lambda h: (h[0], h[1], h[2], h[3]))
    if not hits:
        raise AnchorError(f"「{after}」が{f' {page_no} ページ' if page_no else ''}に見つかりません（窓の字のとおりに指定してください）")
    occ = spec.get("occurrence")
    if occ is not None:
        if not isinstance(occ, int) or not 1 <= occ <= len(hits):
            raise AnchorError(f"occurrence は 1〜{len(hits)} で指定してください")
        hits = [hits[occ - 1]]
    if len(hits) > 1:
        listing = "／".join(f"{k}番目：{h[0]}ページ「{h[5].text.strip()[:30]}」" for k, h in enumerate(hits[:10], 1))
        raise AnchorError(f"「{after}」が {len(hits)} か所にあります。page か occurrence（何番目か）で1つに決めてください：{listing}")
    n, _y, _x, i, cell, ln, lay = hits[0]
    sz = ln.size
    text = ln.text
    j = i + len(after)
    if where == "cover":
        x0, y0, x1, y1 = ln.span_rect(i, j)
        rect = (x0 - 0.3, y0 - 0.1 * sz, x1 + 0.3, y1 + 0.1 * sz)
        covered = {id(c) for c in ln.chars if x0 - 0.5 <= (c.x0 + c.x1) / 2 <= x1 + 0.5}  # 覆う字（その行の、範囲の中の字）
        others = [c for c in lay.chars if not c.text.isspace() and id(c) not in covered and rect[0] < (c.x0 + c.x1) / 2 < rect[2]
                  and rect[1] < (c.y0 + c.y1) / 2 < rect[3]]
        if others:
            raise AnchorError(f"「{after}」を覆うと、ほかの字（「{''.join(c.text for c in others[:6])}」）も隠れます")
    elif where == "right":
        x0 = ln.span_rect(i, j)[2] + 0.3 * sz
        rb = cell.x1 if cell is not None else lay.right
        rest = re.search(r"\S", text[j:])
        if rest:
            rb = min(rb, ln.span_rect(j + rest.start(), j + rest.start() + 1)[0])  # 同じ行の、見出しの後ろの字
        for o in (cell.lines if cell is not None else lay.free_lines):
            if o is not ln and not o.vertical and o.y0 < ln.y1 - 0.2 * sz and o.y1 > ln.y0 + 0.2 * sz and o.x0 >= x0 - 0.5:
                rb = min(rb, o.x0)
        for r in lay.vs:
            if r.a - 1 <= ln.y0 and r.b + 1 >= ln.y1 and x0 <= r.pos < rb:
                rb = r.pos
        rect = (x0, ln.y0 - 0.1 * sz, rb - 1.5, ln.y1 + 0.1 * sz)
        if rect[2] - rect[0] < 1.5 * sz:
            raise AnchorError(f"「{after}」の右に、書く空きがありません（where を \"below\" にすると、見出しの下に書けます）")
    else:
        x0 = ln.span_rect(i, j)[0]
        y0 = ln.y1 + 0.15 * sz
        bottom = cell.y1 if cell is not None else lay.view.height
        right = cell.x1 if cell is not None else lay.right
        for o in (cell.lines if cell is not None else lay.free_lines):
            if o is not ln and not o.vertical and o.y0 >= ln.y1 - 0.2 * sz and o.x0 < right and o.x1 > x0:
                bottom = min(bottom, o.y0)
        for r in getattr(lay, "hs", []):
            if r.a - 1 <= x0 and r.b + 1 >= x0 + sz and ln.y1 <= r.pos < bottom:
                bottom = r.pos
        rect = (x0, y0, right - 1.5, min(bottom - 0.5, y0 + 2.2 * sz))
        if rect[3] - rect[1] < 0.9 * sz or rect[2] - rect[0] < 1.5 * sz:
            raise AnchorError(f"「{after}」の下に、書く空きがありません")
    # 空いている所に、ほかの字がないこと（覆って書くときは、上で確かめた）
    inside = [] if where == "cover" else [c for c in lay.chars if not c.text.isspace() and rect[0] < (c.x0 + c.x1) / 2 < rect[2]
                                          and rect[1] < (c.y0 + c.y1) / 2 < rect[3]]
    if inside:
        raise AnchorError(f"「{after}」の{'右' if where == 'right' else '下'}の所に、ほかの字（「{''.join(c.text for c in inside[:6])}」）があります")
    return _place(n, rect, used, taken_names, {"id": None, "page": n, "kind": "inline" if where == "cover" else "text",
                                              "size": None, "anchored": dict(spec)}, where, after, sz)


def _place(n, rect, used, taken_names, out, where, after, sz, prefix: str = "w", label: str | None = None) -> dict:
    """書く所に名前（p ページ.w 番号。置いた画像は p ページ.s 番号）を付け、ほかの書き込みと重ならないことを確かめる。
    番号は、字の書き込みと画像で通しにする（同じページの書き込みの重なりも、両方で確かめる）。"""
    used = used if used is not None else {}
    taken = {int(m[1]) for nm in (taken_names or set()) for m in [re.fullmatch(rf"p{n}\.[ws](\d+)", nm)] if m}
    for k_, r_ in used.get(n, []):
        taken.add(k_)
        if rect[0] < r_[2] and r_[0] < rect[2] and rect[1] < r_[3] and r_[1] < rect[3]:
            raise AnchorError(f"{n} ページの、ほかの書き込みと重なります")
    k = next(x for x in itertools.count(1) if x not in taken)
    used.setdefault(n, []).append((k, rect))
    if where == "seal":
        size = sz  # 置いた画像の名前（label）は、呼ぶ側が決める
    elif where == "box":
        size = sz
        label = f"{n}ページの枠（左 {rect[0] / MM:.0f}・上 {rect[1] / MM:.0f} mm）"
    else:
        size = min(sz if sz >= 6 else 10.5, (rect[3] - rect[1]) * 0.8) if where == "below" else sz
        label = f"「{after}」を覆って" if where == "cover" else f"{after}（の{'右' if where == 'right' else '下'}）"
    out = {**out, "id": f"p{n}.{prefix}{k}", "rect": [round(v, 2) for v in rect], "size": round(size, 2), "label": label}
    if where == "cover":
        out["cover"] = True
        out["covered"] = after
    return out


MM = 72 / 25.4  # 1 mm のポイント


def _resolve_box(doc, spec: dict, used, layouts, taken_names) -> dict:
    """{"page": ページ, "box": [左, 上, 右, 下]（ページの左上からの mm）, "value": 書く字}：エージェントがページの画像
    （view_page）を見て決めた枠に書く。枠の中にページの字があれば断る（字は消せない）。"""
    from .layout import View

    pages = doc.pages()
    n = spec.get("page")
    box = spec.get("box")
    if not (isinstance(n, int) and 1 <= n <= len(pages)):
        raise AnchorError(f"枠で書くときは page（1〜{len(pages)}）を指定してください")
    if not (isinstance(box, list) and len(box) == 4 and all(isinstance(v, (int, float)) for v in box)):
        raise AnchorError("box は [左, 上, 右, 下]（ページの左上からの mm）で指定してください")
    if spec.get("value") is None:
        raise AnchorError('枠で書くときは {"page": …, "box": […], "value": 書く字} の形で指定してください')
    view = View(doc, pages[n - 1])
    x0, y0, x1, y1 = (v * MM for v in box)
    if not (0 <= x0 < x1 <= view.width + 0.01 and 0 <= y0 < y1 <= view.height + 0.01):
        raise AnchorError(f"box がページ（{view.width / MM:.0f}×{view.height / MM:.0f} mm）の外か、左＜右・上＜下になっていません")
    rect = (x0, y0, x1, y1)
    layouts = layouts if layouts is not None else {}
    if n not in layouts:
        layouts[n] = PageLayout(doc, pages[n - 1], page_content(doc, pages[n - 1]), n)
    lay = layouts[n]
    inside = [c for c in lay.chars if not c.text.isspace() and x0 < (c.x0 + c.x1) / 2 < x1 and y0 < (c.y0 + c.y1) / 2 < y1]
    if inside:
        raise AnchorError(f"枠の中に、ページの字（「{''.join(c.text for c in inside[:6])}」）があります")
    sz = min(10.5, (y1 - y0) * 0.75)
    return _place(n, rect, used, taken_names, {"id": None, "page": n, "kind": "text", "size": None, "anchored": dict(spec)},
                  "box", "", sz)
