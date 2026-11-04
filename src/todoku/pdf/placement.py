"""PDF に書いた字・○の、ページの上の重なりを確かめる（書き出す前に、書いた後の文書で）。

証明（proof.py）と読み戻しは、書いた値と見た目の字が正しいかを確かめるが、どこに描かれたか（印刷された字・罫線・同じ書き込みで
書いたほかの値に重なっていないか）は見ない。ここでは、書いた後の文書から形を取り出して、重なりを確かめる：

- 書いた字の形：欄の見た目（/AP）を、仕様どおりの行列（ISO 32000-1 12.5.5 Algorithm 8.1）で置いた、字ごとの墨の四角
  （字の原点から送りの向きに送りの幅、基線から上に字の大きさの 0.7 倍。字の下がり（ディセンダー）は入れない：下線の上に
  書く字が、下線に触れたとしないため）。字の送りの幅の両端は、字の大きさの 0.06 倍だけ内側
- ○の形：楕円（中心と半径）か四角の囲み線（circle.resolve の決めたもの）と、その線の太さ

断るのは、はっきり重なるときだけ（いまのエージェントは書く所を正しく決め、書いた後のページを view_page で見て確かめるので、
きわどい所は、所と位置（mm）を添えて知らせ、エージェントが見て決める）：
- 断る：(a) とどくが置いた字の墨が、印刷された字の墨（circle.INK）に大きく重なる（小さいほうの面の 1/4 以上）
        (c) 同じ書き込みで書いた値どうしの墨が重なる（字の面の四角 written_ink で比べる）・○の線が、書いた値の墨を髪の毛の太さより深く切る
- 知らせる（warnings）：字の面（circle.face）に少しかかる・罫線に触れる／横切る（○の線も）・曲線（「見本」の印など）に
  かかる・画像（スキャン）に重なる（画像に描かれた罫線・字は確かめられない）
（覆って書く所（cover）は、覆う字に重なってよい）
作り手の決めたフォームの欄（AcroForm）の中の字は、欄の位置を作り手が決めたので、(a)・(b) は見ない。
"""

from __future__ import annotations

import math
import unicodedata

from .content import Interpreter, apply, mul
from .layout import View

INK_UP = 0.7  # 基線から上の墨の高さ（字の大きさに対する比：和文の字面・欧文の大文字の高さの目安）
INK_SIDE = 0.06  # 字の送りの幅の両端から、内側に入れる幅（字の大きさに対する比：字の左右の余白）
EPS = 0.15  # 重なりとみなす、食い込みの下限（pt）：座標の丸めの分
CLEAR = 0.25  # はっきり重なるとみなす、重なりの面積の比（小さいほうの面に対して）
HAIR = 0.3  # 髪の毛の太さ（pt）：○の線がこれより深く書いた値の墨を切るときは断る


def _area(r) -> float:
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def _ovl_ratio(r, s) -> float:
    """二つの四角の重なりの面積の、小さいほうの面積に対する比。"""
    w, h = min(r[2], s[2]) - max(r[0], s[0]), min(r[3], s[3]) - max(r[1], s[1])
    if w <= 0 or h <= 0:
        return 0.0
    small = min(_area(r), _area(s))
    return (w * h) / small if small > 0 else 1.0


def _curve_hits(pts: list, box) -> bool:
    """曲線（制御点を結んだ折れ線で近づけたもの）が、四角の中を通るか。"""
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        for k in range(9):
            x, y = x0 + (x1 - x0) * k / 8, y0 + (y1 - y0) * k / 8
            if box[0] + EPS < x < box[2] - EPS and box[1] + EPS < y < box[3] - EPS:
                return True
    return False


def reopen(doc, blob: bytes):
    """書いた後のバイト列を、元の文書と同じ鍵で開く（暗号化された文書も、パスワードを持たずに）。"""
    from .doc import PdfDocument

    if doc.decryptor is None:
        return PdfDocument(blob, doc.filename)
    d = PdfDocument.__new__(PdfDocument)
    d.blob, d.filename = blob, doc.filename
    d.xref, d.trailer, d.cache, d.objstm_cache = {}, {}, {}, {}
    d.decryptor, d.encrypt_ref, d.repaired = None, None, False
    d._read_xref()
    enc = d.trailer.get("Encrypt")
    d.encrypt_ref = enc if hasattr(enc, "num") else None
    keep = d.encrypt_ref.num if d.encrypt_ref is not None else None
    d.cache = {k: v for k, v in d.cache.items() if k == keep}
    d.objstm_cache = {}
    d.decryptor = doc.decryptor
    return d


def appearance_glyphs(doc, ad: dict) -> list:
    """注釈の見た目（/AP の N。状態のある見た目は /AS の状態）の字（content.Glyph。ページの座標）。読めなければ []。"""
    from .objects import Stream

    apd = doc.resolve(ad.get("AP"))
    n = doc.resolve(apd.get("N")) if isinstance(apd, dict) else None
    if isinstance(n, dict) and not isinstance(n, Stream):
        st = doc.resolve(ad.get("AS"))
        n = doc.resolve(n.get(getattr(st, "name", "")))
    if not isinstance(n, Stream):
        return []
    try:
        rect = [float(doc.resolve(v)) for v in doc.resolve(ad.get("Rect"))]
        bbox = [float(doc.resolve(v)) for v in doc.resolve(n.dict.get("BBox"))]
        m = n.dict.get("Matrix")
        m = tuple(float(doc.resolve(v)) for v in doc.resolve(m)) if m is not None else (1, 0, 0, 1, 0, 0)
    except (TypeError, ValueError):
        return []
    if len(rect) != 4 or len(bbox) != 4 or len(m) != 6:
        return []
    pts = [apply(m, x, y) for x in (bbox[0], bbox[2]) for y in (bbox[1], bbox[3])]
    bx0, by0 = min(p[0] for p in pts), min(p[1] for p in pts)
    bx1, by1 = max(p[0] for p in pts), max(p[1] for p in pts)
    if bx1 - bx0 < 1e-6 or by1 - by0 < 1e-6:
        return []
    rx0, ry0, rx1, ry1 = min(rect[0], rect[2]), min(rect[1], rect[3]), max(rect[0], rect[2]), max(rect[1], rect[3])
    sx, sy = (rx1 - rx0) / (bx1 - bx0), (ry1 - ry0) / (by1 - by0)
    fit = (sx, 0, 0, sy, rx0 - bx0 * sx, ry0 - by0 * sy)
    try:
        pc = Interpreter(doc).run_stream(doc.decode(n), doc.resolve(n.dict.get("Resources")) or {}, mul(m, fit))
    except Exception:  # noqa: BLE001  読めない見た目は、形を確かめられない（証明が別に確かめる）
        return []
    return [g for g in pc.glyphs if not g.invisible and g.text.strip()]


def ink_box(g, view: View, up_ratio: float = INK_UP, down_ratio: float = 0.0) -> tuple[float, float, float, float] | None:
    """字の墨の四角（見た目の座標）：基線から上に up_ratio 字ぶん、下に down_ratio 字ぶん。"""
    ln = math.hypot(g.dx, g.dy)
    adv = g.adv or ln
    if ln <= 0 or adv <= 0 or g.size <= 0:
        return None
    ux, uy = g.dx / ln, g.dy / ln
    vx, vy = -uy, ux  # 送りの向きを反時計回りに 90 度：字の上の向き（ページの座標）
    side = min(INK_SIDE * g.size, 0.25 * adv)
    p0 = (g.ox + ux * side, g.oy + uy * side)
    p1 = (g.ox + ux * (adv - side), g.oy + uy * (adv - side))
    up, dn = up_ratio * g.size, down_ratio * g.size
    corners = [(p[0] + vx * k, p[1] + vy * k) for p in (p0, p1) for k in (up, -dn)]
    vs = [view.pt(x, y) for x, y in corners]
    return min(p[0] for p in vs), min(p[1] for p in vs), max(p[0] for p in vs), max(p[1] for p in vs)


def written_ink(g, view: View) -> tuple[float, float, float, float] | None:
    """書いたものどうしを比べる墨の四角：字の面の高さ（全角の字は字の箱いっぱいに近い：基線の少し下から 0.84 字ぶん上まで。
    ほかの字は基線から大文字の高さまで）。INK_UP の四角（下線・罫線に触れたとしないための低い四角）で比べると、行の間が 1 字に
    満たない 2 つの値の、実際に重なる字を見のがす。"""
    wide = any(unicodedata.east_asian_width(ch) in ("W", "F") for ch in g.text)
    return ink_box(g, view, 0.84, 0.08) if wide else ink_box(g, view, 0.72)


def _overlap(r, s, eps: float = EPS) -> bool:
    return min(r[2], s[2]) - max(r[0], s[0]) > eps and min(r[3], s[3]) - max(r[1], s[1]) > eps


def _rule_crosses_box(rule, horizontal: bool, box) -> bool:
    """罫線（Rule：横線は y・x0〜x1、縦線は x・y0〜y1）が、墨の四角の中を通るか。"""
    h = rule.width / 2
    if horizontal:
        return box[1] + EPS < rule.pos + h and rule.pos - h < box[3] - EPS and min(rule.b, box[2]) - max(rule.a, box[0]) > EPS
    return box[0] + EPS < rule.pos + h and rule.pos - h < box[2] - EPS and min(rule.b, box[3]) - max(rule.a, box[1]) > EPS


def _inside(kind: str, g, x: float, y: float) -> float:
    """囲みの内側なら負、外側なら正（境の目安。楕円は (x/a)^2+(y/b)^2 - 1、四角は縁までの距離）。"""
    if kind == "border":
        d = min(x - g[0], g[2] - x, y - g[1], g[3] - y)
        return -d
    cx, cy, a, b = g
    return ((x - cx) / a) ** 2 + ((y - cy) / b) ** 2 - 1.0


def rule_touches_shape(rule, horizontal: bool, kind: str, g, lw: float) -> bool:
    """罫線（線の太さの帯）が、囲みの線の外側の縁に触れるか（横切らなくても、太い線の縁に触れるだけで）。"""
    from .circle import touches

    h, e = rule.width / 2, lw / 2
    band = (rule.a, rule.pos - h, rule.b, rule.pos + h) if horizontal else (rule.pos - h, rule.a, rule.pos + h, rule.b)
    if kind == "border":
        return g[0] - e <= band[2] and band[0] <= g[2] + e and g[1] - e <= band[3] and band[1] <= g[3] + e
    return touches((g[0], g[1], g[2] + e, g[3] + e), band)


def _rule_crosses_shape(rule, horizontal: bool, kind: str, g, lw: float) -> bool:
    """罫線が、囲みの線を横切るか（線の上に、囲みの内側の点と外側の点がどちらもある）。"""
    if kind == "border":
        x0, y0, x1, y1 = g
        tol_in, tol_out = -(lw / 2 + EPS), lw / 2 + EPS
    else:
        cx, cy, a, b = g
        x0, y0, x1, y1 = cx - a, cy - b, cx + a, cy + b
        r = min(a, b)
        tol_in, tol_out = -((lw / 2 + EPS) / r) * 2, ((lw / 2 + EPS) / r) * 2
    if horizontal:
        if not (y0 < rule.pos < y1):
            return False
        lo, hi = max(rule.a, x0 - 2), min(rule.b, x1 + 2)
        pts = [(lo + (hi - lo) * k / 80, rule.pos) for k in range(81)] if hi > lo else []
    else:
        if not (x0 < rule.pos < x1):
            return False
        lo, hi = max(rule.a, y0 - 2), min(rule.b, y1 + 2)
        pts = [(rule.pos, lo + (hi - lo) * k / 80) for k in range(81)] if hi > lo else []
    vals = [_inside(kind, g, x, y) for x, y in pts]
    return bool(vals) and min(vals) < tol_in and max(vals) > tol_out


def _shape_meets_box(kind: str, g, lw: float, box) -> bool:
    from .circle import touches

    h = lw / 2
    if kind == "border":
        return _overlap((g[0] - h, g[1] - h, g[2] + h, g[3] + h), box)
    return touches((g[0], g[1], g[2] + h, g[3] + h), box)


def _shape_holds_box(kind: str, g, lw: float, box) -> bool:
    """書いた字の墨の四角が、○の線の内側にすっかり入るか（○の中に書いた字は、線を切らない）。"""
    from .circle import inside

    h = lw / 2
    if kind == "border":
        return g[0] + h <= box[0] and g[1] + h <= box[1] and box[2] <= g[2] - h and box[3] <= g[3] - h
    return g[2] > h and g[3] > h and inside((g[0], g[1], g[2] - h, g[3] - h), box)


def check(orig, new, patches: list, created: set[str], covers: set[str]) -> tuple[list[str], list[str]]:
    """返す：(断る理由, 知らせること)。orig：書く前の文書（印刷された字・罫線）、new：書いた後の文書（書いた字の形）。
    created：とどくが置いた欄の名前（窓の推定した欄・目印・枠）。covers：覆って書く所の名前。"""
    from . import circle as _circle
    from .form import read_fields

    problems: list[str] = []
    notes: list[str] = []
    names = {p.id for p in patches if p.kind in ("input", "control", "calc", "xfa")}
    pages = new.pages()
    items: dict[int, list] = {}  # ページ → [(名前, "text", 墨の四角の一覧) か (名前, 形, 囲み, 線の太さ)]
    views: dict[int, View] = {}

    def view_of(n: int) -> View:
        if n not in views:
            views[n] = View(new, pages[n - 1])
        return views[n]

    for f in read_fields(new):
        if f.name not in names:
            continue
        for w in f.widgets:
            if w.ref is None or not isinstance(w.page, int) or not 0 <= w.page < len(pages):
                continue
            wd = new.resolve(w.ref)
            if not isinstance(wd, dict):
                continue
            n = w.page + 1
            gs = appearance_glyphs(new, wd)
            boxes = [b for b in (ink_box(g, view_of(n)) for g in gs) if b]
            if boxes:
                items.setdefault(n, []).append((f.name, "text", boxes, f.name in created,
                                                [b for b in (written_ink(g, view_of(n)) for g in gs) if b]))
    for p in patches:
        if p.kind == "circle" and p.circle:
            sd = p.circle
            g = sd.get("border") if sd.get("style") == "border" else sd.get("ellipse")
            if g:
                items.setdefault(sd["page"], []).append((p.id, sd.get("style", "ellipse"), tuple(g),
                                                          _circle.line_width(sd.get("size") or 10.5)))
    layouts: dict = {}
    for n, its in sorted(items.items()):
        lay = _circle.layout(orig, n, layouts)
        mm = lay.view.mm
        pics = []
        for pic in lay.pc.pictures:
            r = lay.view.rect(pic.x0, pic.y0, pic.x1, pic.y1)
            if (r[2] - r[0]) * (r[3] - r[1]) > 4:
                pics.append(r)

        def where(box) -> str:
            return f"{n}ページ（左 {box[0] / mm:.1f}・上 {box[1] / mm:.1f}・右 {box[2] / mm:.1f}・下 {box[3] / mm:.1f} mm）"

        def outline(it) -> tuple:
            if it[1] == "text":
                bs = it[2]
                return min(b[0] for b in bs), min(b[1] for b in bs), max(b[2] for b in bs), max(b[3] for b in bs)
            k, g, lw = it[1], it[2], it[3]
            if k == "border":
                return g[0] - lw / 2, g[1] - lw / 2, g[2] + lw / 2, g[3] + lw / 2
            return g[0] - g[2] - lw / 2, g[1] - g[3] - lw / 2, g[0] + g[2] + lw / 2, g[1] + g[3] + lw / 2

        for it in its:
            name = it[0]
            ob = outline(it)
            placed = it[1] != "text" or it[3]  # とどくが置いたもの（作り手の欄の中の字は、(a)・(b) を見ない）
            if placed:
                if it[1] == "text" and name not in covers:
                    clear = next(((c, b) for b in it[2] for c in lay.chars if not c.text.isspace()
                                  and _ovl_ratio(b, _circle.face(c, _circle.INK)) >= CLEAR), None)
                    if clear is not None:
                        problems.append(f"{name}：書いた字が、印刷された字「{clear[0].text}」の上に重なります（{where(clear[1])}）")
                    else:
                        hit = next(((c, b) for b in it[2] for c in lay.chars if not c.text.isspace()
                                    and _overlap(b, _circle.face(c))), None)
                        if hit is not None:
                            notes.append(f"{name}：書いた字が、印刷された字「{hit[0].text}」の縁に少しかかる見込みです"
                                         f"（{where(hit[1])}）。view_page で見て確かめてください")
                rules = [(r, True) for r in lay.hs] + [(r, False) for r in lay.vs]
                if it[1] == "text":
                    cross = next(((r, hz, b) for b in it[2] for r, hz in rules if _rule_crosses_box(r, hz, b)), None)
                    if cross is not None:
                        notes.append(f"{name}：書いた字に、ページの{'横' if cross[1] else '縦'}の罫線がかかる見込みです"
                                     f"（{where(cross[2])}）。view_page で見て、字が線に重なっていれば、線のあいだの空いた所に"
                                     "書き直してください（box で枠を決めて書けます）")
                else:
                    cross = next(((r, hz) for r, hz in rules if _rule_crosses_shape(r, hz, it[1], it[2], it[3])), None)
                    touch = None if cross is not None else next(
                        ((r, hz) for r, hz in rules if rule_touches_shape(r, hz, it[1], it[2], it[3])), None)
                    if touch is not None:
                        notes.append(f"{name}：{'四角' if it[1] == 'border' else '○'}の線が、ページの{'横' if touch[1] else '縦'}の罫線"
                                     f"（{'上' if touch[1] else '左'} {touch[0].pos / mm:.1f} mm の線）に触れる見込みです（{where(ob)}）。"
                                     "view_page で見て確かめてください（pad・box で小さくできます）")
                    if cross is not None:
                        alt = "box（○の外枠、mm）" if it[1] == "border" else "\"style\": \"border\"（字に沿った四角）か box（○の外枠、mm）"
                        notes.append(f"{name}：{'四角' if it[1] == 'border' else '○'}の線が、ページの{'横' if cross[1] else '縦'}の罫線に"
                                     f"かかる見込みです（線のあいだに収まる形はありませんでした。{where(ob)}）。view_page で見て、"
                                     f"どの字を囲んだか分かりにくければ、{alt}で書き直してください")
                curves = getattr(lay.pc, "curve_paths", [])
                boxes_ = it[2] if it[1] == "text" else [ob]
                if it[1] == "text":
                    cv = next((b for b in boxes_ for pts in curves if _curve_hits([lay.view.pt(x, y) for x, y in pts], b)), None)
                    if cv is not None:
                        notes.append(f"{name}：書いた字に、ページの曲線（印・図など）がかかる見込みです（{where(cv)}）。"
                                     "view_page で見て確かめてください")
                if any(_overlap(ob, pr, 0.5) for pr in pics):
                    notes.append(f"{name}：書いた所は、ページの画像（スキャン）に重なります。画像に描かれた罫線や字に重なって"
                                 f"いないかは確かめられないので、view_page で見て確かめてください（{where(ob)}）")
        # (c) 同じ書き込みで書いたものどうし
        for i, a in enumerate(its):
            for b in its[i + 1:]:
                if a[1] != "text" and b[1] != "text":
                    continue  # ○どうしは circle.resolve が確かめる
                if a[1] == "text" and b[1] == "text":
                    # 書いたものどうしは、墨が少しでも重なれば断る（字の面の四角で比べる）
                    clear = next((x for x in a[4] for y in b[4] if _overlap(x, y)), None)
                    hit = clear
                else:
                    t, s_ = (a, b) if a[1] == "text" else (b, a)
                    clear = next((x for x in t[2] if x[2] - x[0] > 2 * HAIR and x[3] - x[1] > 2 * HAIR and
                                  _shape_meets_box(s_[1], s_[2], s_[3], (x[0] + HAIR, x[1] + HAIR, x[2] - HAIR, x[3] - HAIR))
                                  and not _shape_holds_box(s_[1], s_[2], s_[3], x)), None)
                    hit = clear or next((x for x in t[2] if _shape_meets_box(s_[1], s_[2], s_[3], x)
                                         and not _shape_holds_box(s_[1], s_[2], s_[3], x)), None)
                if clear is not None:
                    problems.append(f"{a[0]} と {b[0]}：同じ書き込みで書いたものどうしが重なります（{where(clear)}）")
                elif hit is not None:
                    notes.append(f"{a[0]} と {b[0]}：同じ書き込みで書いたものどうしが少しかかる見込みです（{where(hit)}）。"
                                 "view_page で見て確かめてください")
    return problems, notes
