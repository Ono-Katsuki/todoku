"""PDF：印刷された字（選択肢の「有」「2.更新」「昭和」など）を、○で囲む。

指定（fill の値）：{"@c1": {"circle": 囲む字, "near": 近くの見出しの字, "page": ページ, "occurrence": 何番目か,
"pad": 字と○のすき間（pt）, "box": [左, 上, 右, 下]（mm。○の外枠をエージェントが決めるとき）}}。
- 囲む字は、窓の字のとおりに（字の間の空白は、あってもなくてもよい）。同じ字がいくつもあるときは、near（いちばん近い所）か
  occurrence（読む順）で1つに決める。決まらなければ、候補を添えて断る
- ○の形：字の箱（字の送りの幅 × 字の高さ）を、すき間（pad、既定は字の大きさの 0.12 倍）だけ広げた四角の 4 つの角を通る
  楕円のうち、ほかの字に触れないもので、上下と左右のはみ出しがいちばん近いもの。ほかの字は、字の箱の縁から字の大きさの
  1/12 だけ内側（字の面。和文の字は、字の箱いっぱいには描かれない）を、触れてはいけない所とする
- どの楕円もほかの字に触れる・囲む字がページの外にかかるときは、書かずに理由を返す（pad を小さくするか、box で外枠を決める）

書き方：ページの描画（/Contents）の前に「q」だけのストリームを、後ろに楕円を描くストリームを足す（増分の保存）。後ろの
ストリームは、元の描画の q の残り（閉じていない q）を Q で戻してから、ページの初めの座標で描く。元のストリームは1バイトも
変えない。楕円は印（/Artifact、TdkCircle の名前つき）の中に描くので、証明（proof.py）と読み戻しで、どれを足したかが分かる。
"""

from __future__ import annotations

import math
import re

from .anchor import MM, AnchorError
from .content import page_content, tokens
from .layout import PageLayout, View

K = 0.5522847498  # 4 つの 3 次ベジエ曲線で楕円を描くときの、制御点の比
FACE = 1 / 12  # 字の箱の縁から、字の面までの内側の幅（字の大きさに対する比）
TAG = "TdkCircle"


def _lines(lay: PageLayout):
    for c in lay.cells:
        for ln in c.lines:
            yield c, ln
    for ln in lay.free_lines:
        yield None, ln


def _pattern(text: str) -> re.Pattern:
    """字の間の空白（全角・半角）は、あってもなくてもよい。"""
    chars = [re.escape(ch) for ch in text if not ch.isspace()]
    return re.compile(r"[ 　]*".join(chars))


def _gap(r, s) -> float:
    dx = max(0.0, max(r[0], s[0]) - min(r[2], s[2]))
    dy = max(0.0, max(r[1], s[1]) - min(r[3], s[3]))
    return math.hypot(dx, dy)


def _rank(r, s) -> tuple[int, float]:
    """目印 s から字 r への近さ：同じ行の高さ（上下が重なる）にあるものを先に、次に同じ列（左右が重なる）、ほかは距離で。"""
    if min(r[3], s[3]) - max(r[1], s[1]) > 0:
        return 0, max(0.0, max(r[0], s[0]) - min(r[2], s[2]))
    if min(r[2], s[2]) - max(r[0], s[0]) > 0:
        return 1, max(0.0, max(r[1], s[1]) - min(r[3], s[3]))
    return 2, _gap(r, s)


def _hits(lay: PageLayout, text: str):
    pat = _pattern(text)
    out = []
    for cell, ln in _lines(lay):
        for m in pat.finditer(ln.text):
            if m.end() == m.start():
                continue
            out.append((cell, ln, m.start(), m.end()))
    return out


def find(doc, spec: dict, layouts: dict) -> tuple[int, object, object, int, int]:
    """囲む字の所：(ページ, マス, 行, 始め, 終わり)。"""
    target = spec.get("circle")
    if not isinstance(target, str) or not target.strip():
        raise AnchorError('○で囲むときは {"circle": 囲む字, "near": 近くの見出し（省ける）, "page": ページ（省ける）} の形で指定してください')
    pages = doc.pages()
    page_no = spec.get("page")
    if page_no is not None and not (isinstance(page_no, int) and not isinstance(page_no, bool) and 1 <= page_no <= len(pages)):
        raise AnchorError(f"page は 1〜{len(pages)} で指定してください")
    hits = []
    for n in ([page_no] if page_no else range(1, len(pages) + 1)):
        lay = layout(doc, n, layouts)
        for cell, ln, a, b in _hits(lay, target):
            hits.append((n, cell, ln, a, b))
    hits.sort(key=lambda h: (h[0], round(h[2].y0, 1), round(h[2].x0, 1), h[3]))
    within = spec.get("within")
    if within is not None:
        # 窓の選択肢の所（見た目の座標の矩形）：その中に字の中心があるもの
        if not (isinstance(within, list) and len(within) == 4 and page_no):
            raise AnchorError("within は page と [左, 上, 右, 下]（pt）で指定してください")
        def _in(h):
            r = h[2].span_rect(h[3], h[4])
            cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
            return within[0] - 0.5 <= cx <= within[2] + 0.5 and within[1] - 0.5 <= cy <= within[3] + 0.5
        hits = [h for h in hits if _in(h)]
        if len(hits) != 1:
            raise AnchorError(f"選択肢「{target}」の字の所が、ページの字から1つに決まりません")
        return hits[0]
    where = f" {page_no} ページ" if page_no else ""
    if not hits:
        raise AnchorError(f"「{target}」が{where}に見つかりません（窓の字のとおりに指定してください）")
    occ = spec.get("occurrence")
    if occ is not None:
        if not isinstance(occ, int) or isinstance(occ, bool) or not 1 <= occ <= len(hits):
            raise AnchorError(f"occurrence は 1〜{len(hits)} で指定してください（「{target}」は{where} {len(hits)} か所）")
        return hits[occ - 1]
    near = spec.get("near")
    if len(hits) > 1 and isinstance(near, str) and near.strip():
        anchors = []
        for n in sorted({h[0] for h in hits}):
            lay = layout(doc, n, layouts)
            for cell, ln, a, b in _hits(lay, near):
                anchors.append((n, ln.span_rect(a, b), cell))
        if not anchors:
            raise AnchorError(f"near の「{near}」が、「{target}」のあるページに見つかりません")
        scored = []
        for h in hits:
            r = h[2].span_rect(h[3], h[4])
            d = min((_rank(r, ar) for hn, ar, _ac in anchors if hn == h[0]), default=(9, float("inf")))
            scored.append((d, h))
        scored.sort(key=lambda x: x[0])
        best, second = scored[0], scored[1] if len(scored) > 1 else None
        tol = 0.25 * best[1][2].size
        if second is not None and second[0][0] == best[0][0] and second[0][1] - best[0][1] <= tol:
            listing = "／".join(f"{k}番目：{h[0]}ページ「{h[2].text.strip()[:24]}」" for k, h in enumerate(hits[:10], 1))
            raise AnchorError(f"「{near}」に同じくらい近い「{target}」が、2 か所以上あります。occurrence で1つに決めてください：{listing}")
        return best[1]
    if len(hits) > 1:
        listing = "／".join(f"{k}番目：{h[0]}ページ「{h[2].text.strip()[:24]}」" for k, h in enumerate(hits[:10], 1))
        raise AnchorError(f"「{target}」が {len(hits)} か所にあります。near（近くの見出しの字）か occurrence（何番目か）で"
                          f"1つに決めてください：{listing}")
    return hits[0]


def layout(doc, n: int, layouts: dict) -> PageLayout:
    if n not in layouts:
        pg = doc.pages()[n - 1]
        layouts[n] = PageLayout(doc, pg, page_content(doc, pg), n)
    return layouts[n]


def target_chars(lay: PageLayout, ln, a: int, b: int) -> list:
    """行の text[a:b] に当たる、ページの字（空白でないもの）。"""
    r = ln.span_rect(a, b)
    out = []
    for c in ln.chars:
        if c.text.isspace():
            continue
        if ln.vertical:
            if r[1] - 0.01 <= c.cy <= r[3] + 0.01:
                out.append(c)
        elif r[0] - 0.01 <= c.cx <= r[2] + 0.01:
            out.append(c)
    return out


# JIS X 4051（日本語文書の組版方法）の約物の字の面：全角の箱のうち、字が描かれるのは半分（二分）。横書きで、
# 始め括弧類は右半分、終わり括弧類・句点類・読点類は左半分、中点類は中央の半分
OPEN_BRACKETS = set("‘“（〔［｛〈《「『【⦅〘〖«〝")
CLOSE_BRACKETS = set("’”）〕］｝〉》」』】⦆〙〗»〟")
FULL_STOPS = set("。．、，")
MIDDLE_DOTS = set("・：；")


def face(c) -> tuple[float, float, float, float]:
    """字の面：字の箱の縁から、字の大きさの 1/12 だけ内側。全角の約物（JIS X 4051 の括弧類・句読点・中点類）は、字の描かれる
    半分の箱で（横書きの字のみ）。"""
    m = FACE * c.size
    bx0, bx1 = c.x0, c.x1
    if len(c.text) == 1 and not c.vertical and bx1 - bx0 > 0.75 * c.size:
        half = (bx1 - bx0) / 2
        if c.text in OPEN_BRACKETS:
            bx0 = bx0 + half
        elif c.text in CLOSE_BRACKETS or c.text in FULL_STOPS:
            bx1 = bx1 - half
        elif c.text in MIDDLE_DOTS:
            bx0, bx1 = bx0 + half / 2, bx1 - half / 2
    x0, y0, x1, y1 = bx0 + m, c.y0 + m, bx1 - m, c.y1 - m
    if x1 <= x0:
        x0 = x1 = (c.x0 + c.x1) / 2
    if y1 <= y0:
        y0 = y1 = (c.y0 + c.y1) / 2
    return x0, y0, x1, y1


def touches(e: tuple[float, float, float, float], box) -> bool:
    """楕円（中心 cx, cy・半径 a, b）の内側（縁を含む）と、四角が重なるか。"""
    cx, cy, a, b = e
    # 楕円を単位円にする縮め方で、四角も縮める（軸にそろった四角のまま）
    x0, y0, x1, y1 = (box[0] - cx) / a, (box[1] - cy) / b, (box[2] - cx) / a, (box[3] - cy) / b
    dx = 0.0 if x0 <= 0 <= x1 else min(abs(x0), abs(x1))
    dy = 0.0 if y0 <= 0 <= y1 else min(abs(y0), abs(y1))
    return dx * dx + dy * dy <= 1.0


def inside(e, box) -> bool:
    """四角が、楕円の内側にすっかり入るか（4 つの角が内側）。"""
    cx, cy, a, b = e
    return all(((x - cx) / a) ** 2 + ((y - cy) / b) ** 2 <= 1.0 + 1e-9 for x in (box[0], box[2]) for y in (box[1], box[3]))


def ellipses_meet(e1, e2) -> bool:
    """2 つの楕円が重なるか（片方の縁の点が、もう片方の内側にある）。"""
    for e, f in ((e1, e2), (e2, e1)):
        cx, cy, a, b = e
        for k in range(72):
            t = 2 * math.pi * k / 72
            x, y = cx + a * math.cos(t), cy + b * math.sin(t)
            if ((x - f[0]) / f[2]) ** 2 + ((y - f[1]) / f[3]) ** 2 <= 1.0:
                return True
    return False


def shape(rect, size: float, others: list, page: tuple[float, float] | None, pad: float | None = None):
    """囲む字の四角 rect を囲む楕円 (cx, cy, a, b)：rect をすき間だけ広げた四角の角を通り、ほかの字の面に触れないもので、
    左右と上下のはみ出しがいちばん近いもの。見つからなければ (None, 触れる字)。"""
    m = (0.12 * size) if pad is None else pad
    w, h = rect[2] - rect[0] + 2 * m, rect[3] - rect[1] + 2 * m
    cx, cy = (rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2
    best = None
    blockers: dict[int, list] = {}
    for k in range(1, 200):
        t = k / 200  # 角を通る楕円：(w/2a)^2 = t、(h/2b)^2 = 1 - t
        a, b = (w / 2) / math.sqrt(t), (h / 2) / math.sqrt(1 - t)
        if a - w / 2 > 1.5 * size or b - h / 2 > 1.5 * size:
            continue  # はみ出しが字の大きさの 1.5 倍を超える楕円は、どこを囲んだか分かりにくい
        e = (cx, cy, a, b)
        if page is not None and (cx - a < 0 or cy - b < 0 or cx + a > page[0] or cy + b > page[1]):
            continue
        hit = next((o for o in others if touches(e, face(o))), None)
        if hit is not None:
            blockers.setdefault(id(hit), [hit, 0])[1] += 1
            continue
        score = abs((a - w / 2) - (b - h / 2))
        if best is None or score < best[0]:
            best = (score, e)
    if best:
        return best[1], None
    return None, max(blockers.values(), key=lambda x: x[1])[0] if blockers else None


def resolve(doc, spec: dict, layouts: dict, used: dict, taken: set | None = None) -> dict:
    """○で囲む所を決める（書き込みと証明で同じ決め方）。返す：{id, page, text, ellipse（見た目の座標）, label, context, size}。"""
    n, cell, ln, a, b = find(doc, spec, layouts)
    lay = layout(doc, n, layouts)
    mine = target_chars(lay, ln, a, b)
    if not mine:
        raise AnchorError(f"「{spec['circle']}」の字の位置が分かりません")
    ids = {id(c) for c in mine}
    rect = (min(c.x0 for c in mine), min(c.y0 for c in mine), max(c.x1 for c in mine), max(c.y1 for c in mine))
    size = max(c.size for c in mine)
    others = [c for c in lay.chars if id(c) not in ids and not c.text.isspace()
              and c.x1 > rect[0] - 4 * size and c.x0 < rect[2] + 4 * size and c.y1 > rect[1] - 4 * size and c.y0 < rect[3] + 4 * size]
    pad = spec.get("pad")
    if pad is not None and not (isinstance(pad, (int, float)) and not isinstance(pad, bool) and 0 <= pad <= 3 * size):
        raise AnchorError(f"pad は、字と○のすき間（pt、0〜{3 * size:.0f}）で指定してください")
    page = (lay.view.width, lay.view.height)
    box = spec.get("box")
    if box is not None:
        if not (isinstance(box, list) and len(box) == 4 and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in box)):
            raise AnchorError("box は [左, 上, 右, 下]（ページの左上からの mm。○の外枠）で指定してください")
        x0, y0, x1, y1 = (v * MM for v in box)
        if not (0 <= x0 < x1 <= page[0] + 0.01 and 0 <= y0 < y1 <= page[1] + 0.01):
            raise AnchorError("box がページの外か、左＜右・上＜下になっていません")
        e = ((x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2)
        if not inside(e, rect):
            raise AnchorError(f"box の○が、「{spec['circle']}」の字を囲みきれません（字の箱は 左 {rect[0] / MM:.1f}・上 "
                              f"{rect[1] / MM:.1f}・右 {rect[2] / MM:.1f}・下 {rect[3] / MM:.1f} mm）")
        hit = next((o for o in others if touches(e, face(o))), None)
        if hit is not None:
            raise AnchorError(f"box の○が、ほかの字「{hit.text}」に触れます")
    else:
        e, hit = shape(rect, size, others, page, pad)
        if e is None:
            near_txt = hit.text if hit is not None else ""
            raise AnchorError(f"「{spec['circle']}」を○で囲むと、ほかの字" + (f"「{near_txt}」" if near_txt else "") +
                              "に触れます（字の間が詰まっています）。pad（字と○のすき間、pt）を小さくするか、box（○の外枠、mm）で"
                              "決めてください。囲めないときは、本人に伝えてください")
    for k_, e2 in used.get(n, []):
        if ellipses_meet(e, e2):
            raise AnchorError(f"「{spec['circle']}」の○が、{n} ページのほかの○と重なります")
    taken_nums = {int(m[1]) for nm in (taken or set()) for m in [re.fullmatch(rf"p{n}\.c(\d+)", nm)] if m}
    taken_nums |= {k_ for k_, _e in used.get(n, [])}
    k = next(x for x in range(1, 10 ** 6) if x not in taken_nums)
    used.setdefault(n, []).append((k, e))
    context = re.sub(r"[ 　]+", " ", ln.text).strip()
    label = spec.get("near") or (re.sub(r"[ 　]+", " ", cell.text()).strip()[:30] if cell is not None else context[:30])
    return {"id": f"p{n}.c{k}", "page": n, "text": "".join(c.text for c in mine), "ellipse": [round(v, 3) for v in e],
            "label": label, "context": context[:60], "size": size, "rect": [round(v, 3) for v in rect]}


# ---------------------------------------------------------------------- 描く
def line_width(size: float) -> float:
    return round(min(1.2, max(0.6, 0.07 * size)), 2)


def open_depth(data: bytes) -> int:
    """描画の終わりで閉じていない q の数。"""
    d = 0
    for kind, val in tokens(data):
        if kind == "o":
            if val == "q":
                d += 1
            elif val == "Q" and d > 0:
                d -= 1
    return d


def ops(view: View, circles: list[dict], depth: int) -> bytes:
    """後ろに足すストリーム：閉じていない q を戻し（前に足した q の分も）、ページの初めの座標で楕円を描く。"""
    out = bytearray(b"Q\n" * (depth + 1))
    for sd in circles:
        cx, cy, a, b = sd["ellipse"]
        px, py = view.back(cx, cy)
        ra, rb = (b, a) if view.rot in (90, 270) else (a, b)
        lw = line_width(sd["size"])
        out += (f"q\n/Artifact <</Type /Layout /{TAG} ({sd['id']})>> BDC\n0 G {lw} w [] 0 d 0 J 0 j\n"
                f"{px + ra:.3f} {py:.3f} m\n"
                f"{px + ra:.3f} {py + K * rb:.3f} {px + K * ra:.3f} {py + rb:.3f} {px:.3f} {py + rb:.3f} c\n"
                f"{px - K * ra:.3f} {py + rb:.3f} {px - ra:.3f} {py + K * rb:.3f} {px - ra:.3f} {py:.3f} c\n"
                f"{px - ra:.3f} {py - K * rb:.3f} {px - K * ra:.3f} {py - rb:.3f} {px:.3f} {py - rb:.3f} c\n"
                f"{px + K * ra:.3f} {py - rb:.3f} {px + ra:.3f} {py - K * rb:.3f} {px + ra:.3f} {py:.3f} c\n"
                f"h S\nEMC\nQ\n").encode()
    return bytes(out)


PREFIX = b"q\n"


# ---------------------------------------------------------------------- 読み戻し
def drawn(doc, page_no: int) -> list[tuple[str, tuple[float, float, float, float]]]:
    """書いた後のファイルのページに描いてある、とどくの○：(id, 楕円（見た目の座標）)。描画を読み直して、印の中の曲線の外枠から。"""
    pg = doc.pages()[page_no - 1]
    view = View(doc, pg)
    data = doc.page_content(pg)
    out = []
    for m in re.finditer(rb"/" + TAG.encode() + rb"\s*\(([^)]*)\)>>\s*BDC(.*?)EMC", data, re.S):
        pts, stack = [], []
        for kind, val in tokens(m[2]):
            if kind == "n":
                stack.append(float(val))
                continue
            if kind == "o" and val == "m" and len(stack) >= 2:
                pts.append((stack[-2], stack[-1]))
            elif kind == "o" and val == "c" and len(stack) >= 6:
                pts += [(stack[-6], stack[-5]), (stack[-4], stack[-3]), (stack[-2], stack[-1])]
            stack = []
        if len(pts) < 13:
            continue  # 楕円（4 つの曲線）でない
        xs = [view.pt(x, y)[0] for x, y in pts]
        ys = [view.pt(x, y)[1] for x, y in pts]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        out.append((m[1].decode("latin-1"), ((x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2)))
    return out


def read_back(doc, circled: list[dict]) -> list[dict]:
    """書いた後のファイルで、○がそれぞれ、言った字を囲み、ほかの字に触れていないか。違う所を返す。"""
    out = []
    layouts: dict = {}
    by_page: dict[int, dict] = {}
    for c in circled:
        n = c["page"]
        if n not in by_page:
            by_page[n] = dict(drawn(doc, n))
        e = by_page[n].get(c["id"])
        if e is None:
            out.append({"id": c["id"], "written": f"○「{c['text']}」", "read_back": "書いた後のファイルのページに、この○が見つかりません"})
            continue
        lay = layout(doc, n, layouts)
        inner = [ch for ch in lay.chars if not ch.text.isspace() and inside(e, (ch.x0, ch.y0, ch.x1, ch.y1))]
        touch = [ch for ch in lay.chars if not ch.text.isspace() and ch not in inner and touches(e, face(ch))]
        got = "".join(ch.text for ch in sorted(inner, key=lambda ch: (ch.y0, ch.x0) if ch.vertical else (ch.x0, ch.y0)))
        want = re.sub(r"\s", "", c["text"])
        if re.sub(r"\s", "", got) != want or touch:
            what = f"囲んでいる字「{got}」" + (f"・触れている字「{''.join(ch.text for ch in touch[:6])}」" if touch else "")
            out.append({"id": c["id"], "written": f"○「{c['text']}」", "read_back": what})
    return out
