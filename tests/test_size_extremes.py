"""大きさの極端な書類：「大きいから読み減らすは良いけど読めないは困る」。小さくした同じ形で、答えが変わらないことと、2 乗にならないことを確かめる。"""

from __future__ import annotations

import random
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from todoku.limits import Budget, call_budget


# ---------------------------------------------------------------- 1 ページの字が多い PDF
def _ch(x, y, text, seq, size=10.0):
    from todoku.pdf.pagetext import Ch

    return Ch(text=text, x0=x, y0=y, x1=x + size, y1=y + size, size=size, vertical=False, base=y, invisible=False, seq=seq)


def _naive(row):
    """前の形（すべての字の組を比べる）。"""
    from todoku.pdf.pagetext import _clash

    ink = [c for c in row if not c.text.isspace()]
    if not any(_clash(a, b) for a, b in zip(ink, ink[1:])) and not any(_clash(a, b) for a, b in zip(ink, ink[2:])):
        return [row]
    segs = []
    for c in sorted(row, key=lambda c: c.seq):
        if segs:
            p = segs[-1][-1]
            back = c.x0 < p.x0 - 0.1 * max(c.size, p.size)
            if not (back and any(_clash(c, o) for o in segs[-1])):
                segs[-1].append(c)
                continue
        segs.append([c])
    if len(segs) < 2:
        return [row]
    groups = []
    for sg in segs:
        for g in groups:
            if not any(_clash(x, y) for x in sg for y in g):
                g += sg
                break
        else:
            groups.append(list(sg))
    return groups


@pytest.mark.parametrize("seed", range(8))
def test_overprint_split_is_the_same_as_comparing_every_pair(seed):
    from todoku.pdf.pagetext import _split_overprinted

    rng = random.Random(seed)
    row, seq = [], 0
    for _line in range(rng.randint(1, 5)):
        x = rng.choice([0.0, 3.0, 5.0, 40.0])
        for _ in range(rng.randint(2, 25)):
            row.append(_ch(x, rng.choice([0.0, 0.0, 1.0]), chr(0x3042 + seq % 40), seq, rng.choice([8.0, 10.0, 14.0])))
            x += rng.choice([6.0, 10.0, 14.0])
            seq += 1
    got = [[c.seq for c in g] for g in _split_overprinted(list(row))]
    want = [[c.seq for c in g] for g in _naive(list(row))]
    assert got == want


def test_many_overprinted_lines_stay_inside_the_budget():
    from todoku.pdf.pagetext import _split_overprinted

    row = sorted((_ch(float(i % 500) * 10, 0.0, chr(0x3042 + (i // 500) % 40), i) for i in range(30_000)), key=lambda c: (c.x0, c.seq))  # 500 字の行を 60 本重ねる
    with call_budget(Budget(total=None, items=2_000_000)):
        groups = _split_overprinted(row)
    assert sum(len(g) for g in groups) == 30_000 and len(groups) > 1


# ---------------------------------------------------------------- 長い行の .txt・CSV
def test_byte_of_matches_encoding_each_prefix():
    from todoku.datafile.textenc import decode

    text = "ab申請\u3000é" * 700 + "x"
    for enc, data in (("utf-8", text.encode("utf-8")), ("cp932", text.replace("é", "e").encode("cp932")), ("utf-16-le", text.encode("utf-16-le"))):
        d = decode(data, enc)
        for i in (0, 1, 3, 1023, 1024, 1025, 2048, len(d.text)):
            assert d.byte_of(i) == len(d.bom) + len(d.text[:i].encode(d.encoding)), (enc, i)


def test_byte_of_keeps_no_table_per_character():
    from todoku.datafile.textenc import decode

    d = decode(("申請" * 5000).encode("utf-8"), "utf-8")
    assert d.byte_of(9999) == 9999 * 3
    assert len(d._prefix) <= 10000 // 1024 + 2


def test_list_fields_on_a_long_line_goes_in_steps(tmp_path, monkeypatch):
    from todoku import api, lazy
    from todoku.lazy import textdocs

    monkeypatch.setattr(textdocs, "SEG_CHARS", 1000)
    monkeypatch.setattr(textdocs, "SCAN_SEGS", 3)
    p = tmp_path / "t.txt"
    p.write_text("本文です。" * 2000 + "氏名：＿＿＿＿＿", encoding="utf-8")
    doc = lazy.open_doc(str(p), None, None, force=True)
    got = doc.fields({"paragraphs": "1"}, False, 100)
    assert got["next"] and "cursor" in got["next"] and not got["fields"]
    seen = []
    n = 0
    while got["next"]:
        got = doc.fields(got["next"], False, 100)
        seen += got["fields"]
        n += 1
        assert n < 20
    assert len(seen) >= 1 and any("氏名" in str(f.get("label")) for f in seen)


def test_csv_long_cell_list_fields_says_it_cut_and_fill_works(tmp_path, monkeypatch):
    from todoku import api, lazy
    from todoku.lazy.textdocs import FIELD_VALUE_CHARS

    p = tmp_path / "c.csv"
    long = "x" * (200_000)  # 標準の csv の既定の上限 131072 字より長い
    p.write_text(f"a,b\n1,{long}\n", encoding="utf-8")
    doc = lazy.open_doc(str(p), None, None, force=True)
    got = {f["id"]: f for f in doc.fields({"rows": "2-2"}, False, 100)["fields"]}
    assert got["R2.C2"]["value_chars"] == 200_000 and len(got["R2.C2"]["value"]) == FIELD_VALUE_CHARS and "value_cut" in got["R2.C2"]
    assert len(got["R2.C1"]["context"]) < FIELD_VALUE_CHARS + 50
    r = doc.fill({"R2.C1": "9"}, str(tmp_path / "o.csv"), False, False, None)
    assert r["proof_ok"] and r["read_back_ok"]
    assert (tmp_path / "o.csv").read_text(encoding="utf-8") == f"a,b\n9,{long}\n"
