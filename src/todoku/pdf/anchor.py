"""PDF：窓に欄として示していない所に、エージェントが見出しの字を目印にして書く。

PDF の字は消せないので、書くのは見出しの右（または下）の空いている所だけ。空いている所は、見出しの行のマスの縁・
同じ高さの次の字・縦の罫線の手前まで（下なら、次の行かマスの下の縁の手前まで）。書き込みと証明で同じ決め方を使い、
証明では、足した欄の位置が、目印から決め直した位置と同じであることを確かめる。
"""

from __future__ import annotations

import itertools
import re
import unicodedata

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


def _segments(text: str) -> list[tuple[int, int]]:
    """字の並びを、基の字と、それに続く結合の字（濁点・アクセントなど）のまとまりに分ける：(始まり, 終わり)。"""
    out: list[tuple[int, int]] = []
    for k, ch in enumerate(text):
        if out and unicodedata.combining(ch):
            out[-1] = (out[-1][0], k + 1)
        else:
            out.append((k, k + 1))
    return out


def find_spans(text: str, needle: str) -> list[tuple[int, int]]:
    """text の中の needle の位置（元の字の位置で (始まり, 終わり)）。字のとおりに見つからなければ、Unicode の正規化
    （NFC、次に NFKC）をした形どうしで探す（PDF の字が分解された形「e + ̂」で、指定が合成済みの「ê」のときなど）。
    正規化した字の位置は、まとまり（基の字と結合の字）ごとに、元の字の位置へ戻す。"""
    out = [(m.start(), m.end()) for m in re.finditer(re.escape(needle), text)
           if not (m.end() < len(text) and unicodedata.combining(text[m.end()]) and not unicodedata.combining(needle[-1]))]
    if out or not needle:
        return out
    for form in ("NFC", "NFKC"):
        segs = _segments(text)
        norm, owner = [], []
        for si, (a, b) in enumerate(segs):
            t = unicodedata.normalize(form, text[a:b])
            norm.append(t)
            owner += [si] * len(t)
        flat = "".join(norm)
        want = unicodedata.normalize(form, needle)
        for m in re.finditer(re.escape(want), flat):
            if m.end() == m.start():
                continue
            s0, s1 = owner[m.start()], owner[m.end() - 1]
            # まとまりの途中で始まる・終わる所は採らない（字の一部だけに合うのは、同じ字ではない）
            if (m.start() and owner[m.start() - 1] == s0) or (m.end() < len(flat) and owner[m.end()] == s1):
                continue
            out.append((segs[s0][0], segs[s1][1]))
        if out:
            return out
    return out


def _empty_cell_right(lay, cell, ln):
    """見出しのマス cell の右の縁に接し（1.5pt 以内）、見出しの行の高さ（行の中ほど）を含む、字のないマス（いちばん左のもの）。"""
    cy = (ln.y0 + ln.y1) / 2
    cand = [c for c in lay.cells if c is not cell and abs(c.x0 - cell.x1) <= 1.5 and c.y0 - 0.5 <= cy <= c.y1 + 0.5
            and not any(not ch.text.isspace() for ch in c.chars)]
    if cand:
        return min(cand, key=lambda c: (c.x0, c.y0))
    # 右のマスが閉じていない（下の罫線が見出しのマスの下にしか引かれていない）：見出しのマスの右の縁から、次の縦の罫線
    # （見出しのマスの高さにわたるもの）までの、見出しのマスと同じ高さの所。中に字がなければ、そこを空のマスとみなす
    from .layout import Cell

    xs = [r.pos for r in lay.vs if r.pos > cell.x1 + 1.5 * ln.size and r.a - 1 <= cell.y0 and r.b + 1 >= cell.y1]
    if not xs:
        return None
    box = Cell(cell.x1, cell.y0, min(xs), cell.y1)
    if any(not ch.text.isspace() and box.x0 < ch.cx < box.x1 and box.y0 < ch.cy < box.y1 for ch in lay.chars):
        return None
    return box


MIN_READABLE_PT = 6.0  # fill.MIN_READABLE_PT と同じ：これより小さい字は、印刷すると読みにくい


def _fit_size(value, rect, sz: float, unit: float) -> float:
    """rect に value を書くときの字の大きさ（fill.add_widget と同じ決め方）。入らなければ 0。"""
    from .fill import FillError, layout_text

    w, h = rect[2] - rect[0], rect[3] - rect[1]
    if w <= 0 or h <= 0:
        return 0.0
    try:
        return layout_text(str(value), w, h, min(sz, h * 0.8), multiline=h > 2.4 * min(sz, 10.5 / unit), unit=unit)[1]
    except FillError:
        return 0.0


def _right_cell_band(lay, cell, ln):
    """見出しのマス cell の右の縁に接し（1.5pt 以内）、見出しの行の高さを含むマス（いちばん左のもの）の、見出しの行の高さの
    帯のうち、左の縁から、帯にかかる字・縦の罫線の手前まで（字のある所には書かない）。なければ None。"""
    cy = (ln.y0 + ln.y1) / 2
    sz = ln.size
    cand = [c for c in lay.cells if c is not cell and abs(c.x0 - cell.x1) <= 1.5 and c.y0 - 0.5 <= cy <= c.y1 + 0.5]
    if not cand:
        return None
    c = min(cand, key=lambda c: (c.x0, c.y0))
    y0, y1 = max(c.y0 + 0.5, ln.y0 - 0.15 * sz), min(c.y1 - 0.5, ln.y1 + 0.15 * sz)
    x1 = c.x1
    for ch in c.chars:
        if not ch.text.isspace() and ch.y0 < y1 and ch.y1 > y0 and ch.x0 > c.x0:
            x1 = min(x1, ch.x0)
    for r in lay.vs:
        if r.a - 1 <= y0 and r.b + 1 >= y1 and c.x0 + 1.5 < r.pos < x1:
            x1 = r.pos
    return (c.x0 + 1.5, y0, x1 - 1.5, y1)


def resolve_write(doc, spec: dict, used: dict | None = None, layouts: dict | None = None,
                  taken_names: set | None = None) -> dict:
    """見出しの字で指した書き込み → 書く所（窓の欄と同じ形の辞書：id「p ページ.w 番号」・ページ・矩形・字の大きさ）。

    spec：{"after": 見出しの字（窓の字のとおり）, "value": 書く字, "page": ページ（省くと全ページから）,
    "where": "right"（既定。見出しの右）/ "before"（見出しの左。右から左へ書く字の見出し）/ "below"（見出しの下）, "occurrence": 何番目か（読む順。1 から）}、
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
    if where not in ("right", "below", "before", "cover"):
        raise AnchorError('where は "right"（見出しの右）・"before"（見出しの左：右から左へ書く字の見出し）・"below"（見出しの下）です')
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
            for a_, b_ in find_spans(ln.text, after):
                hits.append((n, round(ln.y0, 1), round(ln.x0, 1), a_, cell, ln, lay, b_))
    hits.sort(key=lambda h: (h[0], h[1], h[2], h[3]))
    if not hits:
        raise AnchorError(f"「{after}」が{f' {page_no} ページに' if page_no else ''}見つかりません（窓の字のとおりに指定してください）")
    occ = spec.get("occurrence")
    if occ is not None:
        if not isinstance(occ, int) or not 1 <= occ <= len(hits):
            raise AnchorError(f"occurrence は 1〜{len(hits)} で指定してください")
        hits = [hits[occ - 1]]
    if len(hits) > 1:
        listing = "／".join(f"{k}番目：{h[0]}ページ「{h[5].text.strip()[:30]}」" for k, h in enumerate(hits[:10], 1))
        raise AnchorError(f"「{after}」が {len(hits)} か所にあります。page か occurrence（何番目か）で1つに決めてください：{listing}")
    n, _y, _x, i, cell, ln, lay, j = hits[0]
    sz = ln.size
    text = ln.text
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
        unit = lay.view.unit
        here = _fit_size(value, rect, sz, unit) if rect[2] - rect[0] >= 1.5 * sz else 0.0
        if cell is not None and not rest and here * unit < max(MIN_READABLE_PT, 0.6 * sz * unit):
            # 見出しのマスに空きがない、または、そこでは字が読める大きさで入らない（「氏　　名」の後ろの 1 字ぶんの空き）：
            # すぐ右の、同じ行の高さにあるマス（見出しのマスの右の縁に接する）の、字のない所に書く。そこでも大きく
            # ならなければ、見出しのマスのまま（字が小さくなることは、書くときに伝える）
            band = _right_cell_band(lay, cell, ln)
            if band is None:
                nxt = _empty_cell_right(lay, cell, ln)
                if nxt is not None:
                    # 高さは見出しの行の高さ（マスが何行ぶんもあるとき、ほかの行の見出しの高さに書かない）
                    band = (nxt.x0 + 1.5, max(nxt.y0 + 0.5, ln.y0 - 0.15 * sz), nxt.x1 - 1.5, min(nxt.y1 - 0.5, ln.y1 + 0.15 * sz))
            if band is not None and band[2] - band[0] >= 1.5 * sz and _fit_size(value, band, sz, unit) > here:
                rect = band
        if rect[2] - rect[0] < 1.5 * sz:
            raise AnchorError(f"「{after}」の右に、書く空きがありません（where を \"below\" にすると、見出しの下に書けます）")
    elif where == "before":
        # 右から左へ書く字（アラビア文字など）の見出し：見出しの左の空き（右の空きを左右に映したもの）
        x1 = ln.span_rect(i, j)[0] - 0.3 * sz
        lb = cell.x0 if cell is not None else lay.left
        for c in ln.chars:
            if not c.text.isspace() and c.x1 <= x1 + 0.3 * sz - 0.5:
                lb = max(lb, c.x1)  # 同じ行の、見出しの左の字
        for o in (cell.lines if cell is not None else lay.free_lines):
            if o is not ln and not o.vertical and o.y0 < ln.y1 - 0.2 * sz and o.y1 > ln.y0 + 0.2 * sz and o.x1 <= x1 + 0.5:
                lb = max(lb, o.x1)
        for r in lay.vs:
            if r.a - 1 <= ln.y0 and r.b + 1 >= ln.y1 and lb < r.pos <= x1:
                lb = r.pos
        rect = (lb + 1.5, ln.y0 - 0.1 * sz, x1, ln.y1 + 0.1 * sz)
        if rect[2] - rect[0] < 1.5 * sz:
            raise AnchorError(f"「{after}」の左に、書く空きがありません")
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
        side = {"right": "右", "before": "左", "below": "下"}[where]
        raise AnchorError(f"「{after}」の{side}の所に、ほかの字（「{''.join(c.text for c in inside[:6])}」）があります")
    return _place(n, rect, used, taken_names, {"id": None, "page": n, "kind": "inline" if where == "cover" else "text",
                                              "size": None, "anchored": dict(spec)}, where, after, sz,
                  unit=lay.view.unit)


def _place(n, rect, used, taken_names, out, where, after, sz, prefix: str = "w", label: str | None = None,
           unit: float = 1.0) -> dict:
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
        label = f"{n}ページの枠（左 {rect[0] * unit / MM:.0f}・上 {rect[1] * unit / MM:.0f} mm）"
    else:
        # 6pt・10.5pt は紙の上の大きさ：UserUnit のあるページでは、ページの座標の単位に直す
        size = min(sz if sz * unit >= 6 else 10.5 / unit, (rect[3] - rect[1]) * 0.8) if where == "below" else sz
        label = f"「{after}」を覆って" if where == "cover" else f"{after}（の{ {'right': '右', 'before': '左', 'below': '下'}[where] }）"
    out = {**out, "id": f"p{n}.{prefix}{k}", "rect": [round(v, 2) for v in rect], "size": round(size, 2), "label": label}
    if where == "cover":
        out["cover"] = True
        out["covered"] = after
    return out


MM = 72 / 25.4  # 1 mm のポイント


A4_SHORT = 595.28  # A4 の短い辺（pt、ISO 216）


def box_base_size(view) -> float:
    """枠（mm）に書く字の、大きさの上限：A4 の文書の標準の 10.5pt を、紙の大きさ（A4 との短い辺の比、ISO 216 の紙は
    √2 倍ずつ）に合わせる。A3 のスキャンでは 14.8pt。枠の高さが低ければ、その高さに合わせて小さくする（_resolve_box）。"""
    unit = getattr(view, "unit", 1.0)  # UserUnit（ページの座標の 1 単位が何ポイントか）
    short = min(view.width, view.height) * unit  # 紙の上の長さ（pt）
    return round(10.5 * max(1.0, short / A4_SHORT) / unit, 2)  # ページの座標の単位で


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
    x0, y0, x1, y1 = (v * view.mm for v in box)  # mm → ページの座標（UserUnit も反映）
    if not (0 <= x0 < x1 <= view.width + 0.01 and 0 <= y0 < y1 <= view.height + 0.01):
        raise AnchorError(f"box がページ（{view.width / view.mm:.0f}×{view.height / view.mm:.0f} mm）の外か、左＜右・上＜下になっていません")
    rect = (x0, y0, x1, y1)
    layouts = layouts if layouts is not None else {}
    if n not in layouts:
        layouts[n] = PageLayout(doc, pages[n - 1], page_content(doc, pages[n - 1]), n)
    lay = layouts[n]
    inside = [c for c in lay.chars if not c.text.isspace() and x0 < (c.x0 + c.x1) / 2 < x1 and y0 < (c.y0 + c.y1) / 2 < y1]
    if inside:
        raise AnchorError(f"枠の中に、ページの字（「{''.join(c.text for c in inside[:6])}」）があります")
    sz = min(box_base_size(view), (y1 - y0) * 0.75)
    return _place(n, rect, used, taken_names, {"id": None, "page": n, "kind": "text", "size": None, "anchored": dict(spec)},
                  "box", "", sz, unit=view.unit)
