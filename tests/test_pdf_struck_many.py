"""斜線で消したマス：斜めの線が何万本もあるページ（図面・網かけ）でも、欄ごとに全部の線を調べない。

前は、欄の四角ごとに、ページの斜めの線の全部を調べていた：斜めの線が 52,736 本・欄の候補が 4,839 個のページ（集めた PDF の 1 件）は、
欄の一覧に 60 秒かかった。いまは、線の端を x の順に並べておき、角の近くに端のある線だけを調べる。決まり（片方の端が角の近く、
もう片方の端が向かいの角の近く）は同じ。
"""
import random

from test_pdf_scripts import glyphs_pdf, window
from todoku.pdf.layout import PageLayout


def _all_lines(diagonals, r) -> bool:
    """前の調べ方：全部の線。"""
    x0, y0, x1, y1 = r
    tol = max(2.0, 0.12 * min(x1 - x0, y1 - y0))

    def near(p, q):
        return abs(p[0] - q[0]) <= tol and abs(p[1] - q[1]) <= tol

    for a, b in diagonals:
        for p, q in (((x0, y0), (x1, y1)), ((x1, y0), (x0, y1))):
            if (near(a, p) and near(b, q)) or (near(a, q) and near(b, p)):
                return True
    return False


def _layout(diagonals) -> PageLayout:
    lay = PageLayout.__new__(PageLayout)
    lay.diagonals = diagonals
    return lay


def test_the_same_answer_as_looking_at_every_line():
    rnd = random.Random(11)
    diagonals = []
    for _ in range(4000):
        x, y, w, h = rnd.uniform(0, 600), rnd.uniform(0, 800), rnd.uniform(1, 120), rnd.uniform(1, 60)
        a, b = (x, y), (x + w, y + h)
        if rnd.random() < 0.5:
            a, b = (x + w, y), (x, y + h)  # 右上から左下
        diagonals.append((b, a) if rnd.random() < 0.5 else (a, b))  # 描く向きは、どちらでも
    lay = _layout(diagonals)
    rects = [(x, y, x + rnd.uniform(3, 150), y + rnd.uniform(3, 70)) for x, y in
             ((rnd.uniform(0, 600), rnd.uniform(0, 800)) for _ in range(1500))]
    for a, b in rnd.sample(diagonals, 400):  # 線の角に合わせたマスと、決まりの境目（2 pt・短い辺の 0.12 倍）の前後にずらしたマス
        x0, x1 = sorted((a[0], b[0]))
        y0, y1 = sorted((a[1], b[1]))
        edge = 0.12 * min(x1 - x0, y1 - y0)
        for d in (0.0, 1.99, 2.0, 2.01, -1.99, -2.01, edge - 0.01, edge + 0.01, -edge + 0.01, -edge - 0.01):
            rects += [(x0 + d, y0, x1 + d, y1), (x0, y0 + d, x1, y1 + d), (x0 - d, y0 - d, x1 + d, y1 + d)]
    rects = [r for r in rects if r[2] > r[0] and r[3] > r[1]]
    got = [lay._struck(r) for r in rects]
    assert got == [_all_lines(diagonals, r) for r in rects]
    assert 400 < sum(got) < len(rects) - 400  # 消したマスも、消していないマスもある


def test_no_diagonal_lines():
    assert _layout([])._struck((0, 0, 10, 10)) is False


def test_struck_cells_on_a_page_full_of_diagonal_lines():
    # 見出しのマスと、空のマス 3 つ（2 つ目は左下から右上、3 つ目は右下から左上の斜線で消してある）。下に、網かけの斜めの線 3,000 本
    lines = ["40 700 m 440 700 l S", "40 680 m 440 680 l S"] + [f"{x} 680 m {x} 700 l S" for x in (40, 140, 240, 340, 440)]
    lines += ["240 680 m 340 700 l S", "440 680 m 340 700 l S"]
    lines += [f"{40 + (i % 100) * 5} {100 + (i // 100) * 15} m {44 + (i % 100) * 5} {110 + (i // 100) * 15} l S" for i in range(3000)]
    blob = glyphs_pdf([("氏", 50, 685), ("名", 62, 685)], extra="\n".join(lines))
    rects = [f["rect"] for f in window(blob)["pages"][0]["fields"]]
    assert any(abs(r[0] - 140) < 2 for r in rects)
    assert not any(abs(r[0] - 240) < 2 or abs(r[0] - 340) < 2 for r in rects)
