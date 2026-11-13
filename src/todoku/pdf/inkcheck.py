"""書いた字が、ページの上で本当に見えるかを確かめる（書いた後のファイルを、書く前と同じ描画で描いて比べる）。

証明は「変わったのは書いた欄と見た目だけ」を、読み戻しは「見た目の字が値と同じ」を確かめるが、どちらも描いた結果を見ていない。
黒い帯の上の黒い字のように、値も見た目も正しいのに見えない所は、この 2 つでは分からない。

- 書いた所の四角を、書く前と書いた後で同じ解像度に描き、画素を比べる（とどくの描画：pdf/render。ビューアーと同じ順で、注釈の見た目も描く）
- 変わった画素がない：読み戻しの違いとして返す（ファイルは書き出されるが read_back_ok は false）
- 変わった画素のうち、いちばん差のある所でも字と下地の差が小さい：WCAG 2 の明るさの比が 3:1 未満（大きな字・図の基準）なら知らせる
- 欄の一部だけが下地に紛れる所（黒い帯に半分かかる字・濃い写真の上の字）は、欄の全体の画素では分からない：書いた欄の字を描くときに、
  字（グリフ）ごとに、墨の画素と、その下の色（この欄の字を描く前の色）を控える。墨のどの画素も下と同じ色の字は「見えない字」
  （読み戻しの違い）、墨の半分より多くが 3:1 未満の下地の上にある字は「見えにくい字」（知らせる）。位置は mm で伝える
- 字を描いた後に、その上に別のものが描かれる所（後ろの注釈の見た目・重なったほかの欄）：字を塗った直後の色も控え、描き終えた絵と比べる。
  墨の半分より多くの色が変わった字は「上に別のものが重なった字」として知らせる（透ける色の印の下の字は読めるので、止めずに、見て確かめてもらう）
- 字でないもの（入れたチェック・ラジオの印、字を囲む ○、印影・署名の画像）は、その四角（印は部品、○ は外枠、画像は置いた枠）の全体の
  画素で、字の欄の全体と同じ決まりで見る：変わった画素がなければ読み戻しの違い、いちばん差のある所でも 3:1 未満なら知らせる
- 描ききれなかった（時間切れ・書体がない）ときは、確かめていないことを知らせる（見えるとは言わない）
"""

from __future__ import annotations

import time

from .form import read_fields
from .layout import View

DPI = 100.0  # 字の形を見るのではなく、墨が出たかを見る解像度（6 pt の字でも 8 画素ある）
RENDER_LIMIT = 15.0  # 1 回の描画の時間の上限（秒）
TOTAL_LIMIT = 45.0  # 1 回の書き込みで、確かめに使う時間の上限（秒）
LOW_CONTRAST = 3.0  # WCAG 2.x 1.4.3・1.4.11：大きな字・図の、明るさの比の下限


def _luminance(r: int, g: int, b: int) -> float:
    """WCAG 2.x の相対輝度。"""

    def ch(v: int) -> float:
        c = v / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    la, lb = _luminance(*a), _luminance(*b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def _ink(before: dict, after: dict, box: tuple[float, float, float, float]) -> tuple[int, float]:
    """四角（画像の左上からの mm）の中で変わった画素の数と、その中でいちばん大きい明るさの比。"""
    s = before["scale"]
    ox, oy = before["origin_mm"]
    w, h = before["w"], before["h"]
    x0, y0 = max(0, int((box[0] - ox) * s)), max(0, int((box[1] - oy) * s))
    x1, y1 = min(w, int((box[2] - ox) * s) + 2), min(h, int((box[3] - oy) * s) + 2)
    a, b = before["rgb"], after["rgb"]
    n = 0
    best = 1.0
    for y in range(y0, y1):
        i0, i1 = (y * w + x0) * 3, (y * w + x1) * 3
        if a[i0:i1] == b[i0:i1]:
            continue
        for i in range(i0, i1, 3):
            pa, pb = a[i:i + 3], b[i:i + 3]
            if pa != pb:
                n += 1
                best = max(best, contrast(tuple(pa), tuple(pb)))
    return n, best


def _stand_ins(img: dict, box) -> tuple[int, int, float]:
    """描画が灰色の四角で示した字（字の形がこのパソコンにない字）の、四角（画像の左上からの mm）の中の画素について：
    (画素の数, 字の色が、描く前の下の色と同じ画素の数, 字の色と下の色の、いちばん大きい明るさの比)。
    字の形は分からないが、字の色は分かる：下地がどこも字と同じ色なら、どんな形の字も見えない。"""
    s = img["scale"]
    ox, oy = img["origin_mm"]
    w, h = img["w"], img["h"]
    x0, y0 = max(0, int((box[0] - ox) * s)), max(0, int((box[1] - oy) * s))
    x1, y1 = min(w, int((box[2] - ox) * s) + 2), min(h, int((box[3] - oy) * s) + 2)
    seen: dict = {}
    n = same = 0
    best = 1.0
    for px, under, rgb, alpha in img.get("box_inks") or []:
        k = int(alpha * 255 + 0.5)
        for q, u in zip(px, under):
            if not (x0 <= q % w < x1 and y0 <= q // w < y1):
                continue
            c = seen.get((u, rgb, k))
            if c is None:
                ink = rgb if k >= 255 else tuple((u[i] * (255 - k) + rgb[i] * k + 127) // 255 for i in range(3))
                c = seen[(u, rgb, k)] = 0.0 if ink == tuple(u) else contrast(tuple(u), ink)
            n += 1
            if c == 0.0:
                same += 1
            elif c > best:
                best = c
    return n, same, best


class _Glyphs:
    """書いた 1 つの欄の字（グリフ）ごとの、墨の画素と、その下の色。描画（render/paint）が、字を塗る前に glyph を、塗った後に
    painted を呼ぶ。下の色は、この欄の字が初めてその画素に触れる前の色（同じ字を 2 回重ねて描く太字も、1 回目の前の色と比べる）。
    塗った後の色は、この欄の字が終わりにその画素を塗った直後の色（描き終えた絵と違えば、後から別のものが上に描かれた）。"""

    def __init__(self) -> None:
        self.under: dict[int, bytes] = {}
        self.after: dict[int, bytes] = {}
        self.glyphs: list[tuple[list, tuple, float]] = []

    def painted(self, cv) -> None:
        if self.glyphs:
            buf, after = cv.buf, self.after
            for q, _a in self.glyphs[-1][0]:
                after[q] = bytes(buf[q * 3:q * 3 + 3])

    def glyph(self, cv, px: list, rgb, alpha: float) -> None:
        if not px:
            self.glyphs.append(([], (0, 0, 0), 0.0))  # painted が、前の字の画素を塗り直したと見ないように
            return
        buf, under = cv.buf, self.under
        for q, _a in px:
            if q not in under:
                under[q] = bytes(buf[q * 3:q * 3 + 3])
        self.glyphs.append((px, tuple(int(v) for v in rgb), alpha))

    def judge(self, final=None) -> tuple[int, list, list, float, list]:
        """返す：(墨のある字の数, 見えない字ごとの画素の番号の並び, 見えにくい字ごとの画素の番号の並び, 見えにくい字の墨の下でいちばん小さい比,
        上に別のものが重なった字ごとの画素の番号の並び)。
        見えない字：墨のどの画素も、下と同じ色。見えにくい字：墨（濃さで量る）の半分より多くが、比 3:1 未満の下地の上。
        重なった字（final：描き終えた絵の RGB があるとき）：見える字のうち、墨の半分より多くが、塗った直後と違う色になった字。"""
        seen: dict = {}
        n = 0
        hidden: list = []
        dim: list = []
        over: list = []
        lowest = LOW_CONTRAST
        after = self.after
        for px, rgb, alpha in self.glyphs:
            k = int(alpha * 255 + 0.5)
            total = same = low = 0
            least = LOW_CONTRAST
            for q, a in px:
                u = self.under[q]
                c = seen.get((u, rgb, k))
                if c is None:
                    ink = rgb if k >= 255 else tuple((u[i] * (255 - k) + rgb[i] * k + 127) // 255 for i in range(3))
                    c = seen[(u, rgb, k)] = 0.0 if ink == tuple(u) else contrast(tuple(u), ink)
                total += a
                if c < LOW_CONTRAST:
                    low += a
                    if c == 0.0:
                        same += a
                    elif c < least:
                        least = c
            if not total:
                continue
            n += 1
            if same == total:
                hidden.append([q for q, _a in px])
                continue
            if low * 2 > total:
                dim.append([q for q, _a in px])
                lowest = min(lowest, least if least < LOW_CONTRAST else 1.0)
            if final is not None:
                changed = sum(a for q, a in px if q in after and final[q * 3:q * 3 + 3] != after[q])
                if changed * 2 > total:
                    over.append([q for q, _a in px])
        return n, hidden, dim, lowest, over


def _place(n: int, glyphs: list, img: dict) -> str:
    """字ごとの画素の番号の並びを囲む四角を、ページの左上からの mm で。"""
    w, s = img["w"], img["scale"]
    ox, oy = img["origin_mm"]
    xs, ys = [q % w for px in glyphs for q in px], [q // w for px in glyphs for q in px]
    return f"{n}ページ（左 {ox + min(xs) / s:.1f}・上 {oy + min(ys) / s:.1f}・右 {ox + (max(xs) + 1) / s:.1f}・下 {oy + (max(ys) + 1) / s:.1f} mm）"


def _old_text(field) -> str:
    v = field.value_text() if field is not None else None
    return v if isinstance(v, str) else ""


def check(orig, new, patches: list) -> tuple[list[dict], list[str]]:
    """返す：(読み戻しの違い, 知らせること)。orig：書く前の文書、new：書いた後の文書。"""
    from .pageimage import render

    from .objects import decode_text

    old = {f.name: f for f in read_fields(orig)}
    by_page: dict[int, list] = {}
    annots: dict[int, str] = {}  # 書いた部品の注釈の辞書 → 欄の名前
    probes: dict[str, _Glyphs] = {}

    def probe_of(kind: str, key):
        """描画が、書いた欄の字を描く所で呼ぶ：注釈の見た目（annot）か、とどくがページに描いた値の印の範囲（span：TdkField）。"""
        if kind == "annot":
            name = annots.get(id(key))
        else:
            name = decode_text(key["TdkField"]) if "TdkField" in key else None
        return probes.get(name)

    marks: dict[int, list] = {}  # ページ → [(id, 何か, 書いたもの, 四角, 四角がページの座標か（違えば、ページの左上からの mm）)]：字でないもの
    for p in patches:
        if p.kind == "circle" and p.circle and p.circle.get("box_mm"):
            marks.setdefault(p.page, []).append((p.id, "○", f"「{p.value}」を囲む ○", list(p.circle["box_mm"]), False))
        elif p.kind == "seal" and p.seal and p.seal.get("rect"):
            mm = p.seal.get("mm") or 72 / 25.4
            marks.setdefault(p.page, []).append((p.id, "画像", str(p.value), [v / mm for v in p.seal["rect"]], False))
    for f in read_fields(new):
        p = next((x for x in patches if x.id == f.name and x.kind in ("input", "calc", "xfa")), None)
        if p is None or (p.kind == "xfa" and f.kind in ("checkbox", "radio")):
            # 静的な XFA に入れた印は、同じ名前の PDF の欄の部品で見る
            if f.kind in ("checkbox", "radio") and any(x.id == f.name and x.kind in ("control", "xfa") for x in patches) and not f.visibility()[0]:
                # 入れた印：「入」になった部品（書く前と同じ状態の部品は、絵が変わらなくてよい）
                was = old.get(f.name)
                prev = {(w.ref.num if w.ref is not None else None): w.state for w in was.widgets} if was is not None else {}
                for w in f.widgets:
                    if w.state in (None, "Off") or prev.get(w.ref.num if w.ref is not None else None) == w.state or not w.on_screen:
                        continue
                    if isinstance(w.page, int) and w.page >= 0:
                        marks.setdefault(w.page + 1, []).append((f.name, "印", f"印（{w.state}）", w.rect, True))
            continue
        if p.kind == "xfa":
            # 静的な XFA：同じ名前の PDF の欄にも書いた。描いた字は、欄の値（XFA の書式で整えた字）
            shown = _old_text(f) if f.kind in ("text", "combo", "list") else ""
        else:
            shown = str(p.shown if p.shown is not None else p.value or "")
        if not shown.strip():
            continue
        before = old.get(f.name)
        if before is not None and _old_text(before) == shown:
            continue  # 同じ値を書き直した：絵は変わらなくてよい
        if f.visibility()[0]:
            continue  # 隠れた欄（allow_hidden で書いた）：fill が知らせた
        for w in f.widgets:
            if not w.on_screen:  # 画面に出ない部品（NoView）は、画面の描画では確かめられない（fill が知らせた）
                continue
            if isinstance(w.page, int) and w.page >= 0:
                by_page.setdefault(w.page + 1, []).append((f.name, shown, w.rect))
                probes[f.name] = _Glyphs()
                if w.ref is not None:
                    annots[id(new.resolve(w.ref))] = f.name
    bad: list[dict] = []
    notes: list[str] = []
    deadline = time.monotonic() + TOTAL_LIMIT
    pages = new.pages()
    for n in sorted(set(by_page) | set(marks)):
        spots = by_page.get(n, [])
        ids = "・".join(dict.fromkeys([s[0] for s in spots] + [m[0] for m in marks.get(n, [])]))
        if time.monotonic() > deadline:
            notes.append(f"{n} ページの {ids}：字が見えるかの確かめを時間内にできませんでした。view_page で見て確かめてください")
            continue
        view = View(new, pages[n - 1])
        boxes = []  # 部品ごと（同じ欄の部品が、このページにいくつかあるときも、どれも見る）
        for name, shown, r in spots:
            x0, y0, x1, y1 = view.rect(*r)
            boxes.append((name, shown, [x0 / view.mm, y0 / view.mm, x1 / view.mm, y1 / view.mm], r))
        others = []  # 字でないもの：(id, 何か, 書いたもの, 四角（mm）)
        for name, what, wrote, r, in_page in marks.get(n, []):
            if in_page:
                x0, y0, x1, y1 = view.rect(*r)
                r = [x0 / view.mm, y0 / view.mm, x1 / view.mm, y1 / view.mm]
            others.append((name, what, wrote, r))
        pw, ph = view.width / view.mm, view.height / view.mm
        every = [b[2] for b in boxes] + [o[3] for o in others]
        region = [max(0.0, min(b[0] for b in every)), max(0.0, min(b[1] for b in every)),
                  min(pw, max(b[2] for b in every)), min(ph, max(b[3] for b in every))]
        if region[2] - region[0] < 1 or region[3] - region[1] < 1:
            # 四角がページの外か、とても小さい：描く所がない
            for name, shown, _b, r in boxes:
                bad.append({"id": name, "written": shown, "read_back": f"書いた所（{[round(v, 1) for v in r]}）がページの外で、字が見えません"})
            for name, what, wrote, r in others:
                bad.append({"id": name, "written": wrote, "read_back": f"書いた所（{[round(v, 1) for v in r]} mm）がページの外で、{what}が見えません"})
            continue
        for name, _shown, _b, _r in boxes:
            probes[name] = _Glyphs()  # 画素の番号は、このページの絵のもの（同じ名前の欄が、ほかのページにもあるとき）
        try:
            a = render(orig, n, region, dpi=DPI, time_limit=RENDER_LIMIT)
            b = render(new, n, region, dpi=DPI, time_limit=RENDER_LIMIT, probe_of=probe_of)
        except Exception as e:  # noqa: BLE001  描けないページは、確かめられなかったと言う（書き込みは止めない）
            notes.append(f"{n} ページの {ids}：字が見えるかを確かめられませんでした（ページを描けません：{type(e).__name__}）。view_page で見て確かめてください")
            continue
        if a["stopped"] or b["stopped"] or (a["w"], a["h"]) != (b["w"], b["h"]):
            notes.append(f"{n} ページの {ids}：字が見えるかの確かめを時間内にできませんでした。view_page で見て確かめてください")
            continue
        unknown = bool(b["missing_fonts"])
        told = set()  # 部品の全体の画素で、もう伝えた欄
        for name, shown, box, _r in boxes:
            count, best = _ink(a, b, box)
            where = f"{n}ページ（左 {box[0]:.1f}・上 {box[1]:.1f}・右 {box[2]:.1f}・下 {box[3]:.1f} mm）"
            if not count:
                told.add(name)
                if unknown:
                    notes.append(f"{name}：字を描く書体がこのパソコンにないので、字が見えるかを確かめられませんでした（{where}）。view_page で見て確かめてください")
                else:
                    bad.append({"id": name, "written": shown,
                                "read_back": f"値は入りましたが、描いても字が見えません（{where} の画素が 1 つも変わりません。"
                                             "同じ色の下地の上か、後から描かれるものの下かもしれません）。view_page で見て、別の所に書き直してください"})
            elif best < LOW_CONTRAST:
                told.add(name)
                notes.append(f"{name}：書いた字と下地の明るさの差が小さく（比 {best:.1f}:1。読みやすさの下限は 3:1）、見えにくいかもしれません"
                             f"（{where}）。view_page で見て確かめてください")
        unsure = []
        # 印の字の形がこのパソコンにないと、描画は、字の位置に灰色の四角を描く：その四角の明るさは、印の明るさではない
        boxed = unknown or (b.get("stats") or {}).get("glyphs_box", 0) > (a.get("stats") or {}).get("glyphs_box", 0)
        for name, what, wrote, box in others:
            count, best = _ink(a, b, box)
            where = f"{n}ページ（左 {box[0]:.1f}・上 {box[1]:.1f}・右 {box[2]:.1f}・下 {box[3]:.1f} mm）"
            if not count:
                if unknown and what == "印":
                    unsure.append(name)  # 印の字を描く書体がないのかもしれない
                else:
                    bad.append({"id": name, "written": wrote,
                                "read_back": f"{what}は入りましたが、描いても見えません（{where} の画素が 1 つも変わりません。"
                                             "同じ色の下地の上か、後から描かれるものの下かもしれません）。view_page で見て確かめてください"})
                continue
            stood, same, top = _stand_ins(b, box) if what == "印" else (0, 0, 1.0)
            if stood:
                # 印の字の形がこのパソコンにない（描画は灰色の四角を描いた）：四角の明るさではなく、印の色と、その下の色で決める
                if same == stood:
                    bad.append({"id": name, "written": wrote,
                                "read_back": f"{what}は入りましたが、描いても見えません（{where} の下地が、印と同じ色です。印の字の形が"
                                             "このパソコンにないので、印の色と下地の色で確かめました）。view_page で見て確かめてください"})
                elif top < LOW_CONTRAST:
                    notes.append(f"{name}：書いた{what}と下地の明るさの差が小さく（比 {top:.1f}:1。読みやすさの下限は 3:1）、見えにくいかもしれません"
                                 f"（{where}）。view_page で見て確かめてください")
            elif best < LOW_CONTRAST and not (boxed and what == "印"):
                notes.append(f"{name}：書いた{what}と下地の明るさの差が小さく（比 {best:.1f}:1。読みやすさの下限は 3:1）、見えにくいかもしれません"
                             f"（{where}）。view_page で見て確かめてください")
        if unsure:
            notes.append(f"{n} ページの {'・'.join(dict.fromkeys(unsure))}：描くのに要る書体がこのパソコンにないので、印が見えるかを"
                         "確かめられませんでした。view_page で見て確かめてください")
        if unknown:
            continue
        for name, shown in dict.fromkeys((x[0], x[1]) for x in boxes):
            if name in told:
                continue
            # 欄のどの部品も、どこかは見える：字ごとに、下地に紛れた字がないかを見る
            total, hidden, dim, lowest, over = probes[name].judge(b["rgb"])
            if hidden:
                bad.append({"id": name, "written": shown,
                            "read_back": f"値は入りましたが、描いた字 {total} 個のうち {len(hidden)} 個は、描いても見えません"
                                         f"（{_place(n, hidden, b)} の字は、下地と同じ色です。ほかの字は見えます）。"
                                         "view_page で見て、別の所に書き直してください"})
            if dim:
                notes.append(f"{name}：描いた字 {total} 個のうち {len(dim)} 個は、字と下地の明るさの差が小さく（比 {lowest:.1f}:1 から。"
                             f"読みやすさの下限は 3:1）、見えにくいかもしれません（{_place(n, dim, b)}）。view_page で見て確かめてください")
            if over:
                notes.append(f"{name}：描いた字 {total} 個のうち {len(over)} 個の上に、後から描かれるもの（ほかの注釈・重なった欄など）が"
                             f"重なっています（{_place(n, over, b)}）。字が隠れていないか、view_page で見て確かめてください")
    return bad, notes

