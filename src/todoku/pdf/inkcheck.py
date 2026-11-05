"""書いた字が、ページの上で本当に見えるかを確かめる（書いた後のファイルを、書く前と同じ描画で描いて比べる）。

証明は「変わったのは書いた欄と見た目だけ」を、読み戻しは「見た目の字が値と同じ」を確かめるが、どちらも描いた結果を見ていない。
黒い帯の上の黒い字のように、値も見た目も正しいのに見えない所は、この 2 つでは分からない。

- 書いた所の四角を、書く前と書いた後で同じ解像度に描き、画素を比べる（とどくの描画：pdf/render。ビューアーと同じ順で、注釈の見た目も描く）
- 変わった画素がない：読み戻しの違いとして返す（ファイルは書き出されるが read_back_ok は false）
- 変わった画素のうち、いちばん差のある所でも字と下地の差が小さい：WCAG 2 の明るさの比が 3:1 未満（大きな字・図の基準）なら知らせる
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


def _old_text(field) -> str:
    v = field.value_text() if field is not None else None
    return v if isinstance(v, str) else ""


def check(orig, new, patches: list) -> tuple[list[dict], list[str]]:
    """返す：(読み戻しの違い, 知らせること)。orig：書く前の文書、new：書いた後の文書。"""
    from .pageimage import render

    old = {f.name: f for f in read_fields(orig)}
    by_page: dict[int, list] = {}
    for f in read_fields(new):
        p = next((x for x in patches if x.id == f.name and x.kind in ("input", "calc")), None)
        if p is None or not str(p.shown if p.shown is not None else p.value or "").strip():
            continue
        shown = str(p.shown if p.shown is not None else p.value)
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
    bad: list[dict] = []
    notes: list[str] = []
    deadline = time.monotonic() + TOTAL_LIMIT
    pages = new.pages()
    for n, spots in sorted(by_page.items()):
        ids = "・".join(dict.fromkeys(s[0] for s in spots))
        if time.monotonic() > deadline:
            notes.append(f"{n} ページの {ids}：字が見えるかの確かめを時間内にできませんでした。view_page で見て確かめてください")
            continue
        view = View(new, pages[n - 1])
        boxes = {}
        for name, shown, r in spots:
            x0, y0, x1, y1 = view.rect(*r)
            boxes.setdefault(name, (shown, [x0 / view.mm, y0 / view.mm, x1 / view.mm, y1 / view.mm], r))
        pw, ph = view.width / view.mm, view.height / view.mm
        region = [max(0.0, min(b[1][0] for b in boxes.values())), max(0.0, min(b[1][1] for b in boxes.values())),
                  min(pw, max(b[1][2] for b in boxes.values())), min(ph, max(b[1][3] for b in boxes.values()))]
        if region[2] - region[0] < 1 or region[3] - region[1] < 1:
            # 四角がページの外か、とても小さい：描く所がない
            for name, (shown, _b, r) in boxes.items():
                bad.append({"id": name, "written": shown, "read_back": f"書いた所（{[round(v, 1) for v in r]}）がページの外で、字が見えません"})
            continue
        try:
            a = render(orig, n, region, dpi=DPI, time_limit=RENDER_LIMIT)
            b = render(new, n, region, dpi=DPI, time_limit=RENDER_LIMIT)
        except Exception as e:  # noqa: BLE001  描けないページは、確かめられなかったと言う（書き込みは止めない）
            notes.append(f"{n} ページの {ids}：字が見えるかを確かめられませんでした（ページを描けません：{type(e).__name__}）。view_page で見て確かめてください")
            continue
        if a["stopped"] or b["stopped"] or (a["w"], a["h"]) != (b["w"], b["h"]):
            notes.append(f"{n} ページの {ids}：字が見えるかの確かめを時間内にできませんでした。view_page で見て確かめてください")
            continue
        unknown = bool(b["missing_fonts"])
        for name, (shown, box, _r) in boxes.items():
            count, best = _ink(a, b, box)
            where = f"{n}ページ（左 {box[0]:.1f}・上 {box[1]:.1f}・右 {box[2]:.1f}・下 {box[3]:.1f} mm）"
            if not count:
                if unknown:
                    notes.append(f"{name}：字を描く書体がこのパソコンにないので、字が見えるかを確かめられませんでした（{where}）。view_page で見て確かめてください")
                else:
                    bad.append({"id": name, "written": shown,
                                "read_back": f"値は入りましたが、描いても字が見えません（{where} の画素が 1 つも変わりません。"
                                             "同じ色の下地の上か、後から描かれるものの下かもしれません）。view_page で見て、別の所に書き直してください"})
            elif best < LOW_CONTRAST:
                notes.append(f"{name}：書いた字と下地の明るさの差が小さく（比 {best:.1f}:1。読みやすさの下限は 3:1）、見えにくいかもしれません"
                             f"（{where}）。view_page で見て確かめてください")
    return bad, notes

