"""PDF の取り消し線：字の書式ではなく、字の上に描いた線。線と字の位置から決めて、字は消さずに伝える。

- 行の向きに沿った、白くない直線（線・細い塗りの四角）が、字の四角の中央の半分を通る：その線のかかる字の後ろに
  「〔取り消し線：「字」〕」。ページの struck に、字・その行・位置。記法の説明は、取り消し線のあるときだけ
- 下線・字の上の線・太い帯・白い線・字より 1 字ぶんより長く続く線・両端がほかの線につながる線（枠の辺・罫線）・
  同じ道のほかの部分と重なる塗り・縞の帯は、取り消し線と言わない
- 破線は、空きが 1 字ぶんまでならひと続き。字の間の空いた並び（「譲　受　人」）は、空白ごと 1 つの範囲
- 字のすぐ隣に続く空白の字（描かれた空白のグリフ）も線の下にあれば、線は字より長くてよい（「　　本人　　」を空白ごと消した所）
- 縦書き（縦書きの書体・横書きの書体の字を積んだ行）は、縦の線
- 取り消し線と決めた線は、下線の罫線（書く所）にしない
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402

from todoku import api, lazy  # noqa: E402

ROW = [("申請者氏名", 100, 700)]  # 12 pt・全角：字は x 100〜160、基線 700（em の中心は 704.56、中央の半分は 701.56〜707.56）
MARK = "〔取り消し線：「%s」〕"
KEY = "〔取り消し線：「字」〕"


def _open(tmp_path, blob: bytes, name: str = "a.pdf"):
    p = tmp_path / name
    p.write_bytes(blob)
    api._WINDOWS.clear()
    lazy.clear()
    _kind, win, parts = api.window(str(p))
    return win, win["pages"][0], "\n".join(md for _n, md in parts), str(p)


def _text(tmp_path, rows, extra: str, size: float = 12):
    return _open(tmp_path, pdfmini.text_pdf(rows, size=size, extra_content=extra))


def _struck(pg: dict) -> list[str]:
    return [s["text"] for s in pg.get("struck", [])]


def _line(y: float, x0: float = 100, x1: float = 160, w: float = 0.5) -> str:
    return f"{w} w {x0} {y} m {x1} {y} l S"


# ---- 取り消し線と伝える


def test_stroked_line_over_text(tmp_path):
    win, pg, md, path = _text(tmp_path, ROW, _line(704.5))
    assert "申請者氏名" + MARK % "申請者氏名" in md
    (s,) = pg["struck"]
    assert s["text"] == "申請者氏名" and s["context"] == "申請者氏名"
    assert s["rect"] == pytest.approx([100.0, 842 - 712, 160.0, 842 - 700], abs=3.0)
    assert KEY in win["notation"] and KEY in md
    assert all(c["ok"] for c in api.audit(path)["checks"] if c["name"].startswith("読み落とし"))


def test_thin_filled_rect_is_a_line(tmp_path):
    _win, pg, md, _p = _text(tmp_path, ROW, "100 704.2 60 0.6 re f")
    assert _struck(pg) == ["申請者氏名"] and MARK % "申請者氏名" in md


def test_coloured_line(tmp_path):
    _win, pg, _md, _p = _text(tmp_path, ROW, "1 0 0 RG " + _line(704.5))
    assert _struck(pg) == ["申請者氏名"]


def test_only_the_characters_under_the_line(tmp_path):
    _win, pg, md, _p = _text(tmp_path, ROW, _line(704.5, 112, 148))
    assert _struck(pg) == ["請者氏"]
    assert "申請者氏" + MARK % "請者氏" + "名" in md  # 字は消さず、その字のすぐ後ろに


def test_character_half_covered_is_not_struck(tmp_path):
    # 「請」と「氏」は、幅の半分より少ししか線の下にない
    _win, pg, _md, _p = _text(tmp_path, ROW, _line(704.5, 119, 141))
    assert _struck(pg) == ["者"]


def test_double_line_is_one_mark(tmp_path):
    _win, pg, md, _p = _text(tmp_path, ROW, _line(703.5) + " " + _line(705.5))
    assert _struck(pg) == ["申請者氏名"] and md.count(MARK % "申請者氏名") == 1


@pytest.mark.parametrize("y", [702, 704.5, 707])
def test_line_in_the_middle_half(tmp_path, y):
    _win, pg, _md, _p = _text(tmp_path, ROW, _line(y))
    assert _struck(pg) == ["申請者氏名"]


def test_two_words_two_marks(tmp_path):
    rows = [("合併", 100, 700), ("分割", 200, 700)]
    _win, pg, md, _p = _text(tmp_path, rows, _line(704.5, 100, 124) + " " + _line(704.5, 200, 224))
    assert _struck(pg) == ["合併", "分割"]
    assert md.count("〔取り消し線：「") == 3  # 2 つの印と、記法の説明


def test_only_one_of_two_words(tmp_path):
    rows = [("氏名", 100, 700), ("印", 200, 700)]
    _win, pg, md, _p = _text(tmp_path, rows, _line(704.5, 100, 124))
    assert _struck(pg) == ["氏名"] and MARK % "印" not in md


def test_spaced_out_text_is_one_range(tmp_path):
    rows = [("譲", 100, 700), ("受", 140, 700), ("人", 180, 700)]
    _win, pg, _md, _p = _text(tmp_path, rows, _line(704.5, 100, 192))
    (t,) = _struck(pg)
    assert "".join(t.split()) == "譲受人" and t[0] == "譲" and t[-1] == "人"


def test_dashed_line_over_text(tmp_path):
    dash = "0.5 w " + " ".join(f"{x} 704.5 m {x + 3} 704.5 l S" for x in range(100, 160, 5))
    _win, pg, _md, _p = _text(tmp_path, ROW, dash)
    assert len(pg["struck"]) == 1 and _struck(pg)[0].startswith("申請者氏")


def test_slightly_longer_line(tmp_path):
    # 字の並びから、1 字ぶん（12 pt）までは外へ出てよい
    _win, pg, _md, _p = _text(tmp_path, ROW, _line(704.5, 90, 170))
    assert _struck(pg) == ["申請者氏名"]


def test_line_over_the_blank_glyphs_next_to_the_text(tmp_path):
    # 空白ごと選んで消した所：空白のグリフ（100〜124・148〜172）も線の下。線は、字（124〜148）より 2 字ぶん長い
    _win, pg, md, _p = _text(tmp_path, [("　　本人　　", 100, 700)], _line(704.5, 100, 172))
    assert _struck(pg) == ["本人"] and MARK % "本人" in md


def test_line_over_the_blank_glyphs_between_other_words(tmp_path):
    # 氏名 100〜124・空白 124〜148・本人 148〜172・空白 172〜196・印 196〜208
    _win, pg, md, _p = _text(tmp_path, [("氏名　　本人　　印", 100, 700)], _line(704.5, 124, 196))
    assert _struck(pg) == ["本人"]
    assert MARK % "氏名" not in md and MARK % "印" not in md


def test_mark_goes_after_a_place_in_the_struck_text(tmp_path):
    # 字の形の空欄（○○）は所：取り消し線の印は、所の印の中に入れない（○ は半角の幅：124〜136）
    _win, pg, md, _p = _text(tmp_path, [("令和○○年", 100, 700)], _line(704.5, 124, 136))
    assert _struck(pg) == ["○○"]
    i, j = md.index("＝○○］"), md.index(MARK % "○○")
    assert i < j and md[j + len(MARK % "○○")] == "年"


# ---- 取り消し線と言わない


@pytest.mark.parametrize("extra", [
    _line(698.5),  # 下線
    _line(701),  # 基線のすぐ上
    _line(708.2),  # 字の上の端の近く
    _line(711),  # 字の上の線
    _line(704.5, w=3.5),  # 字を塗る帯（太さが字の大きさの 1/4 より太い）
    "1 1 1 RG " + _line(704.5),  # 白い線
    _line(704.5, 60, 200),  # 字より 1 字ぶんより長く続く線
    "0.5 w " + " ".join(f"{x} 704.5 m {x + 3} 704.5 l S" for x in range(40, 300, 5)),  # 字より長く続く破線
    _line(704.5) + " 100 690 m 100 720 l S 160 690 m 160 720 l S",  # 両端が縦の線につながる（マスの中の線）
    _line(704.5) + " 160 690 m 160 720 l S",  # 片方の端が縦の線につながる
    "100 703.5 60 2 re 101 704 58 1 re f",  # 同じ道のほかの部分と重なる塗り（枠の内と外）
    "0.6 w " + " ".join(f"100 {702 + 0.6 * k:.1f} m 160 {702 + 0.6 * k:.1f} l S" for k in range(9)),  # 縞で塗った帯
    "0.5 w 100 698 m 160 711 l S",  # 斜めの線
    "0.5 w 130 690 m 130 720 l S",  # 行を横切る縦の線
    "",
])
def test_not_a_strike(tmp_path, extra):
    win, pg, md, _p = _text(tmp_path, ROW, extra)
    assert "struck" not in pg and "取り消し線" not in md
    assert KEY not in win["notation"]
    assert "申請者氏名" in md


@pytest.mark.parametrize("rows, extra", [
    ([("本人", 124, 700)], _line(704.5, 100, 172)),  # 同じ線でも、隣に空白のグリフがなければ、字より長く続く線
    ([("　　本人　　", 100, 700)], _line(704.5, 60, 212)),  # 空白のグリフより、さらに 1 字ぶんより長く続く線
    ([("　本人　", 112, 700), ("　　", 100, 716), ("　　", 148, 716)], _line(704.5, 88, 184)),  # 線が中央の半分を通らない空白（上の行の空白）は、並びに入れない
])
def test_long_line_without_blank_glyphs_under_it(tmp_path, rows, extra):
    _win, pg, md, _p = _text(tmp_path, rows, extra)
    assert "struck" not in pg and "取り消し線" not in md


def test_same_rect_alone_is_a_strike(tmp_path):
    # 上の「同じ道のほかの部分と重なる塗り」と同じ四角でも、ひとつだけなら線
    _win, pg, _md, _p = _text(tmp_path, ROW, "100 703.5 60 2 re f")
    assert _struck(pg) == ["申請者氏名"]


def test_thickness_up_to_a_quarter_of_the_size(tmp_path):
    _win, pg, _md, _p = _text(tmp_path, ROW, _line(704.5, w=2.5))
    assert _struck(pg) == ["申請者氏名"]


# ---- 縦書き


def _vertical(tmp_path, extra: str):
    chars = "申請者の"
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    objs = [pdfmini.type0("TestFont", 3, vertical=True), pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in code.items()}))]
    blob = pdfmini.build("BT /F1 12 Tf 300 800 Td <0001000200030004> Tj ET\n" + extra, "<< /Font << /F1 2 0 R >> >>", objs)
    return _open(tmp_path, blob)


def test_vertical_font_vertical_line(tmp_path):
    _win, pg, md, _p = _vertical(tmp_path, "0.5 w 300 800 m 300 752 l S")
    assert _struck(pg) == ["申請者の"] and MARK % "申請者の" in md


def test_vertical_font_partial(tmp_path):
    _win, pg, md, _p = _vertical(tmp_path, "0.5 w 300 788 m 300 764 l S")
    assert _struck(pg) == ["請者"] and "申請者" + MARK % "請者" + "の" in md


def test_vertical_font_side_line_is_not_a_strike(tmp_path):
    _win, pg, md, _p = _vertical(tmp_path, "0.5 w 307 800 m 307 752 l S")  # 字の右の端の線（傍線）
    assert "struck" not in pg and "取り消し線" not in md


def test_stacked_glyphs_vertical_line(tmp_path):
    rows = [(ch, 300, 800 - 12 * k) for k, ch in enumerate("申請者の")]  # 横書きの書体の字を、上から下へ積んだ行
    _win, pg, _md, _p = _text(tmp_path, rows, "0.5 w 306 765 m 306 811 l S")
    assert _struck(pg) == ["申請者の"]


def test_stacked_glyphs_horizontal_line_is_not_a_strike(tmp_path):
    rows = [(ch, 300, 800 - 12 * k) for k, ch in enumerate("申請者の")]
    _win, pg, _md, _p = _text(tmp_path, rows, "0.5 w 300 792.5 m 312 792.5 l S")  # 縦の行の 1 字を横切る線
    assert "struck" not in pg


# ---- 書く所


def test_strike_line_is_not_an_underline_place(tmp_path):
    # 字の中ほどより少し上を通る取り消し線は、「すぐ上に字のない横線」でもあるが、下線の罫線（書く所）にしない
    _win, pg, _md, _p = _text(tmp_path, ROW, _line(707))
    assert _struck(pg) == ["申請者氏名"]
    assert not [f for f in pg["fields"] if f["id"].startswith("p1.u")]


def test_underline_place_elsewhere_stays_and_can_be_written(tmp_path):
    _win, pg, _md, path = _text(tmp_path, ROW, _line(704.5) + " 0.5 w 100 650 m 300 650 l S")
    (u,) = [f for f in pg["fields"] if f["id"].startswith("p1.u")]
    out = str(tmp_path / "o.pdf")
    r = api.fill(path, {u["id"]: "山田太郎"}, out)
    assert r["read_back_ok"]
    api._WINDOWS.clear()
    lazy.clear()
    _kind, win, parts = api.window(out)
    md = "\n".join(m for _n, m in parts)
    assert "山田太郎" in md and MARK % "申請者氏名" in md  # 書いた後も、取り消し線の印は同じ
    assert [s["text"] for s in win["pages"][0]["struck"]] == ["申請者氏名"]
