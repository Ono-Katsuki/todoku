"""第5回の見直し（実際の様式での記入）で見つかった、証明と読み戻しを通るのに描くと誤る書き込みの直し（合成した書類で確かめる）。

- 書いた字・○の重なり（placement）：印刷された字・ページの罫線・同じ書き込みのほかの値に重なれば、書き出さずに断る。
  画像（スキャン）に重なる所は、確かめられないことを知らせる
- 文字の中の広い空欄（「区長宛て［　　　　］年」）：中央でなく、後ろの字の手前にそろえる
- ○の選び方：断るときの一覧の番号は、occurrence に渡す番号と同じ。near の字が何か所にもあり、それぞれに近い候補が違えば断る。
  occurrence で指した字は、字の並びの中の字でも囲む（結果に添える）
"""

from __future__ import annotations

import re

import pytest

from todoku import api
from todoku.pdf.fill import inline_align

from test_circle import _pdf, _write


# ====================================================================== 重なり
def test_box_write_crossed_by_a_rule_is_refused(tmp_path):
    # 横線（y 700pt = 上 50.8 mm）をまたぐ枠：字を線が横切るので断る。線の上の枠なら書ける
    data = _pdf([(70, 760, 10, "氏名")], b"0.8 w 60 700 m 540 700 l S\n")
    top = (842 - 700) * 25.4 / 72
    with pytest.raises(api.ApiError, match="罫線が横切ります"):
        _write(tmp_path, data, "t.pdf", {"@1": {"page": 1, "box": [40, top - 2.5, 100, top + 2.5], "value": "山田 花子"}})
    r, _ = _write(tmp_path, data, "t.pdf", {"@1": {"page": 1, "box": [40, top - 6, 100, top - 0.5], "value": "山田 花子"}})
    assert r["proof_ok"] and r["read_back_ok"]


def test_written_values_overlapping_each_other_are_refused(tmp_path):
    data = _pdf([(70, 760, 10, "氏名")])
    with pytest.raises(api.ApiError, match="重なります"):
        _write(tmp_path, data, "t.pdf", {"@1": {"page": 1, "box": [40, 60, 100, 66], "value": "ヤマダ ハナコ"},
                                         "@2": {"page": 1, "box": [40, 61, 100, 70], "value": "山田 花子"}})


def test_circle_crossing_table_rules_or_written_value_is_refused(tmp_path):
    # 高さ 12pt の行（上下に罫線）の中の 10pt の字：楕円は行の罫線を横切るので断る
    rules = b"0.8 w 60 690 m 540 690 l S 60 678 m 540 678 l S\n"
    data = _pdf([(70, 681, 10, "有・無")], rules)
    with pytest.raises(api.ApiError, match="○の線が、ページの横の罫線を横切ります"):
        _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "有"}})
    # 書いた値に重なる○
    data = _pdf([(70, 700, 10, "有・無")])
    left = 70 * 25.4 / 72
    with pytest.raises(api.ApiError, match="同じ書き込みで書いたものどうしが重なります"):
        _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "無"},
                                         "@1": {"page": 1, "box": [left + 9, (842 - 712) * 25.4 / 72, left + 40,
                                                                    (842 - 698) * 25.4 / 72], "value": "本店"}})


# ====================================================================== 広い空欄のそろえ方
@pytest.mark.parametrize("rect, text_w, want", [([0, 0, 288, 12], 30, "right"), ([0, 0, 40, 12], 30, "center")])
def test_wide_gap_between_words_aligns_to_the_following_text(rect, text_w, want):
    assert inline_align({"label": "目黒区長宛て［ ］年 月 日", "rect": rect, "size": 12}, text_w) == want
    assert inline_align({"label": "目黒区長宛て［ ］年 月 日"}) == "center"  # 幅の分からないときは、これまでどおり


# ====================================================================== ○の選び方
def _choices_pdf() -> bytes:
    return _pdf([(70, 700, 10, "入　浴"), (150, 700, 10, "１　自立　２　介助"),
                 (70, 680, 10, "浴　槽"), (150, 680, 10, "１　和式　２　洋式"),
                 (70, 640, 10, "都道府県名"), (70, 620, 10, "見出し")])


def test_ambiguous_near_is_refused(tmp_path):
    with pytest.raises(api.ApiError, match="near の「浴」が 2 か所"):
        _write(tmp_path, _choices_pdf(), "t.pdf", {"@c1": {"circle": "２", "near": "浴", "style": "border"}})
    r, _ = _write(tmp_path, _choices_pdf(), "t.pdf", {"@c1": {"circle": "２", "near": "浴　槽", "style": "border"}})
    assert "洋式" in r["circled"][0]["around"]


def _tie_pdf() -> bytes:
    # 見出し「入浴」の左右に、同じだけ離れた「１」（と、ほかの行の「１」）
    return _pdf([(120, 700, 10, "１"), (150, 700, 10, "入浴"), (190, 700, 10, "１"), (70, 600, 10, "１　自立")])


@pytest.mark.parametrize("near", ["入浴", None])
def test_refusal_numbers_are_the_occurrence_numbers(tmp_path, near):
    # 断りの一覧の番号を occurrence に渡すと、その字を囲む（near のあり・なしのどちらでも）
    spec = {"circle": "１", "style": "border", **({"near": near} if near else {})}
    with pytest.raises(api.ApiError) as e:
        _write(tmp_path, _tie_pdf(), "t.pdf", {"@c1": spec})
    listed = re.findall(r"(\d+)番目：1ページ・([^「]*)「([^」]*)」", str(e.value))
    assert len(listed) >= 2, str(e.value)
    for k, pos, around in listed:
        r, _ = _write(tmp_path, _tie_pdf(), f"t{k}.pdf", {"@c1": {**spec, "occurrence": int(k)}})
        assert r["circled"][0]["around"] == around
        left = float(re.search(r"左 (\d+)", pos)[1])
        assert abs(r["circled"][0]["border_pt"][0] * 25.4 / 72 - left) < 1.5


def test_explicit_occurrence_circles_a_character_inside_a_word(tmp_path):
    with pytest.raises(api.ApiError, match="occurrence"):
        _write(tmp_path, _choices_pdf(), "t.pdf", {"@c1": {"circle": "都", "near": "見出し", "style": "border"}})
    r, _ = _write(tmp_path, _choices_pdf(), "t.pdf", {"@c1": {"circle": "都", "near": "見出し", "occurrence": 1,
                                                              "style": "border"}})
    assert r["proof_ok"] and r["read_back_ok"]
    assert r["circled"][0]["around"].startswith("【都】") and r["circled"][0]["inside_word"]
    assert "長い字の並びの中" in r["say_to_user"]
