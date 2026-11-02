"""速さのための仕組みが、答えを変えないこと。"""

from __future__ import annotations

import random

from todoku.xlsx.workbook import Validation


def test_many_ranges_index_gives_the_same_answer():
    rnd = random.Random(7)
    sq = []
    for _ in range(400):
        r1, c1 = rnd.randint(1, 300), rnd.randint(1, 40)
        sq.append((r1, c1, r1 + rnd.choice([0, 0, 3, 80, 500]), c1 + rnd.randint(0, 5)))
    v = Validation(sq, "list", None, None, None, True, None, None, None)
    for _ in range(3000):
        r, c = rnd.randint(1, 900), rnd.randint(1, 50)
        assert v.covers(r, c) == any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in sq)
    v.sqref = sq[:5]  # 範囲が変われば、索引も作り直す
    assert v.covers(*sq[0][:2]) and not v.covers(10_000, 1)


def test_window_cache_returns_independent_copies(tmp_path):
    from todoku import api
    from test_docx import SAMPLE_BODY, make_docx

    p = tmp_path / "a.docx"
    p.write_bytes(make_docx(SAMPLE_BODY))
    k, doc, parts = api.window(str(p))
    doc["fields"].clear()  # 呼んだ側が変えても
    k2, doc2, parts2 = api.window(str(p))
    assert doc2["fields"] and parts2 == parts  # 控えは変わらない
    p.write_bytes(make_docx(SAMPLE_BODY + "<w:p><w:r><w:t>追記</w:t></w:r></w:p>"))
    assert api.window(str(p))[1] != doc2  # 中身が変われば作り直す
