"""書く欄の多い書類で、同じ部品を欄ごとに展開し直さない（展開の合計の上限は、爆弾を止めるためのもの。大きさの上限は変えない）。"""

from __future__ import annotations

import pytest

from todoku import api, limits


def _many_fields(path, n: int = 300):
    """折り返すセルと、字の入ったセル（文字の中の空欄に書く）がたくさんあるブック。"""
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, n + 1):
        ws.cell(r, 1, f"見出し{r}")
        ws.cell(r, 2).alignment = openpyxl.styles.Alignment(wrap_text=True)
        ws.cell(r, 2).border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(style="thin"))
        ws.cell(r, 3, "令和　　年")
        ws.row_dimensions[r].height = 300  # 折り返しても切れない：書いたことで行が増えるかの見積もりまで進む
    wb.save(path)
    return path


def _count_inflates(monkeypatch) -> dict:
    got: dict = {}
    orig = limits.inflate

    def inflate(data, wbits=15, what="", part=None, tolerant=False):
        got[what] = got.get(what, 0) + 1
        return orig(data, wbits, what, part, tolerant)

    monkeypatch.setattr(limits, "inflate", inflate)
    return got


def test_parts_are_not_inflated_again_for_every_written_cell(tmp_path, monkeypatch):
    p = _many_fields(tmp_path / "a.xlsx")
    got = _count_inflates(monkeypatch)
    values = {f"S!B{i}": "長めの試しの値です。" * 4 for i in range(1, 301)}
    values.update({f"S!C{i}@2": "7" for i in range(1, 301)})
    r = api.fill(str(p), values, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"], r
    many = {part: n for part, n in got.items() if part.startswith("xl/") and n >= 40}
    assert not many, many  # 書く欄の数（600）によらない


def test_total_inflate_limit_still_stops_a_bomb():
    """上限は弱めていない：同じ大きさの展開を繰り返せば、合計で止まる。"""
    import zlib

    data = zlib.compress(b"x" * (1 << 20))
    with limits.call_budget(limits.Budget(total=3 << 20)):
        limits.inflate(data)
        limits.inflate(data)
        limits.inflate(data)
        with pytest.raises(limits.LimitExceeded):
            limits.inflate(data)
