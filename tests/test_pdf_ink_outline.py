"""PDF に書いた字の墨：字形の輪郭の四角で比べ、輪郭が枠（見た目の切り抜き）を出ない高さに書く。

- 書いた後の確かめ（placement）は、書いたファイルに埋め込んだ書体の、その字の輪郭の四角を使う：基線より下に下がる字（g・j・p・q・y）、
  記号を重ねた字（Ấ・ệ）、字を組み立てる文字（タイ文字の上下の記号）の墨を、あるとおりに見る。埋め込まない書体（見る道具が自分の
  書体で描く）の字は、字の大きさからの見積もりのまま
- 埋め込まない Helvetica（標準の 14 書体）の字は、Adobe の AFM の字形の四角で比べる
- 書いたものどうし：輪郭の四角どうしは、はっきり重なれば断り、少しかかるだけなら知らせる
- 書くとき：輪郭が枠の上下を出るなら、枠に入る高さに寄せる。墨の高さが枠より高ければ、字を小さくする。輪郭が字の箱（基線の下 0.12・
  上 0.88 字ぶん）に入る字の書き方は、前と同じ
"""

from __future__ import annotations

import re

import pytest

from todoku import api
from todoku.pdf import glyphs, placement, shape
from todoku.pdf.doc import PdfDocument
from todoku.pdf.form import read_fields
from todoku.pdf.pagetext import View

from test_circle import _pdf, _write

MM = 72 / 25.4


def _written(out: str):
    """書いた欄の (見た目の座標の欄の四角, 字の一覧, View)。"""
    doc = PdfDocument(open(out, "rb").read(), out)
    view = View(doc, doc.pages()[0])
    for f in read_fields(doc):
        for w in f.widgets:
            return view.rect(*w.rect), placement.appearance_glyphs(doc, doc.resolve(w.ref)), view
    raise AssertionError("書いた欄がない")


def _ink(out: str) -> tuple[tuple, list, float]:
    """(欄の四角, 墨の四角の一覧（見た目の座標）, 字の大きさ)。"""
    rect, gs, view = _written(out)
    return rect, [view.rect(*b) for g in gs for b in g.ink or ()], gs[0].size


def _fill(tmp_path, value: str, box: list, name: str = "t.pdf", lines=None):
    return _write(tmp_path, _pdf(lines or [(70, 760, 10, "氏名")]), name, {"@1": {"page": 1, "box": box, "value": value}})


# ====================================================================== 輪郭の四角
def test_outline_box_shows_descender_and_stacked_marks(tmp_path):
    _r, out = _fill(tmp_path, "Ấg", [40, 60, 140, 70])
    _rect, gs, view = _written(out)
    a, g = [x for x in gs if x.text.strip()]
    base = view.pt(a.ox, a.oy)[1]  # 基線（見た目の y、下向き）
    (ab,), (gb,) = [view.rect(*b) for b in a.ink], [view.rect(*b) for b in g.ink]
    assert base - ab[1] > 0.93 * a.size, "記号を重ねた大文字は、字の大きさに近い高さまで墨がある"
    assert abs(ab[3] - base) < 0.03 * a.size
    assert 0.18 * g.size < gb[3] - base < 0.32 * g.size, "g は基線より下に墨がある"
    assert 0.45 * g.size < base - gb[1] < 0.65 * g.size, "g の上の端は、小文字の高さ"
    # 字の大きさからの見積もりは、どちらも見ていなかった
    est = placement.written_ink(g, view)
    assert est[3] <= base + 0.01 and base - placement.written_ink(a, view)[1] < 0.75 * a.size


def test_font_not_in_the_file_keeps_the_estimate(tmp_path):
    # ハングルは、埋め込まない書体（見る道具の書体）で描く：輪郭はファイルにないので、見積もりで比べる
    if glyphs.choose("한")[0] in ("TdkB", "TdkE"):
        pytest.skip("ハングルを埋め込む書体で描く環境")
    r, out = _fill(tmp_path, "한글", [40, 60, 140, 70])
    assert r["proof_ok"] and r["read_back_ok"]
    _rect, gs, _view = _written(out)
    assert all(g.ink is None for g in gs)


def test_shaped_line_has_one_box_per_glyph(tmp_path):
    if shape.font_for("Thai") is None:
        pytest.skip("このパソコンに、タイ文字を組み立てられる書体がない")
    _r, out = _fill(tmp_path, "ปู่ ที่", [40, 60, 140, 72])
    rect, gs, view = _written(out)
    assert len(gs) == 1, "組み立てた行は、1 つの字にまとめて読む"
    boxes = [view.rect(*b) for b in gs[0].ink]
    assert len(boxes) >= 6, "上下に重ねた記号も、1 つずつの四角"
    base = view.pt(gs[0].ox, gs[0].oy)[1]
    assert max(b[3] for b in boxes) - base > 0.1 * gs[0].size, "下の母音は基線より下"
    assert base - min(b[1] for b in boxes) > 0.85 * gs[0].size, "上に重ねた声調の記号"
    # 語の間の空白には、墨の四角がない
    xs = sorted((b[0], b[2]) for b in boxes)
    assert max(b[0] - a[1] for a, b in zip(xs, xs[1:])) > 0.15 * gs[0].size
    assert rect[1] - 0.01 <= min(b[1] for b in boxes) and max(b[3] for b in boxes) <= rect[3] + 0.01


# ====================================================================== 書くとき：枠に入る高さ
@pytest.mark.parametrize("value", ["Ấn Độ", "gjpqy", "Ấg", "Nguyễn Thị Hồng"])
@pytest.mark.parametrize("height_mm", [3.5, 4.2, 5.0, 8.0])
def test_ink_stays_inside_the_box(tmp_path, value, height_mm):
    r, out = _fill(tmp_path, value, [40, 60, 140, 60 + height_mm])
    assert r["proof_ok"] and r["read_back_ok"]
    rect, boxes, _size = _ink(out)
    assert boxes and rect[1] - 0.02 <= min(b[1] for b in boxes) and max(b[3] for b in boxes) <= rect[3] + 0.02


def test_ordinary_text_is_written_as_before(tmp_path):
    # 輪郭が字の箱に入る字：枠の上下の中ほど（基線は、枠の下の端から 高さ/2 - 0.3 字ぶん）
    for i, (value, h) in enumerate([("山田花子", 5.0), ("ABC 123", 5.0), ("テスト", 3.5), ("東京都千代田区", 9.0)]):
        _r, out = _fill(tmp_path, value, [40, 60, 140, 60 + h], f"t{i}.pdf")
        rect, gs, view = _written(out)
        s = gs[0].size
        base = view.pt(gs[0].ox, gs[0].oy)[1]
        assert abs((rect[3] - base) - ((rect[3] - rect[1]) / 2 - 0.3 * s)) < 0.02, value


def test_tall_ink_shrinks_when_the_box_is_lower_than_the_ink(tmp_path):
    # 3 行：行の高さ（字の大きさの 1.15 倍）では入るが、上に重ねた記号と下に下がる字の墨は枠を出る → 小さくして、枠に入れる
    _r0, out0 = _fill(tmp_path, "山田\n花子\n山田", [40, 60, 140, 72.4], "t0.pdf")
    _r, out = _fill(tmp_path, "Ấg\nẤg\nẤg", [40, 60, 140, 72.4])
    rect, boxes, size = _ink(out)
    assert size < _ink(out0)[2]
    assert rect[1] - 0.02 <= min(b[1] for b in boxes) and max(b[3] for b in boxes) <= rect[3] + 0.02


def test_ink_extent_from_the_embedded_outlines():
    top, bottom = glyphs.ink_extent("山田")
    assert 0.7 < top < 0.9 and -0.12 < bottom < 0
    top, bottom = glyphs.ink_extent("Ấg")
    assert top > 0.93 and bottom < -0.18
    assert glyphs.ink_extent("   ") is None
    if glyphs.choose("한")[0] not in ("TdkB", "TdkE"):
        assert glyphs.ink_extent("한글") is None, "埋め込まない書体の字は数えない"


# ====================================================================== 書いた後の確かめ
def test_descender_touching_print_is_told(tmp_path):
    # 枠（上 60〜下 65 mm）の下の端に、印刷された字の上の端が 3 pt 入っている：下に下がる字の墨だけが届く
    lines = [(150, 842 - 65 * MM - 8.8 + 3.0, 10, "申請書類申請書類")]
    r, _ = _fill(tmp_path, "gjpqy gjpqy", [50, 60, 120, 65], "g.pdf", lines)
    assert any(re.search("印刷された字「.」の縁に少しかかる見込み", w) for w in r["warnings"]), r["warnings"]
    for i, v in enumerate(["ABCDE", "山田花子"]):
        r, _ = _fill(tmp_path, v, [50, 60, 120, 65], f"n{i}.pdf", lines)
        assert not any("印刷された字" in w for w in r["warnings"]), (v, r["warnings"])


def test_real_height_meets_a_rule_above(tmp_path):
    # 枠の上の端の 0.75 mm 下を通る横線：片仮名・漢字の上の端（基線から 0.8 字ぶんほど）は線に届く。小文字（x の高さ）は届かない
    top = 60.0
    y = 842 - (top + 0.75) * MM
    rules = f"0.5 w 100 {y:.2f} m 400 {y:.2f} l S\n".encode()
    data = _pdf([(70, 760, 10, "氏名")], rules)
    r, _ = _write(tmp_path, data, "k.pdf", {"@1": {"page": 1, "box": [50, top, 120, top + 4.2], "value": "テスト"}})
    assert any("横の罫線がかかる見込み" in w for w in r["warnings"]), r["warnings"]
    r, _ = _write(tmp_path, data, "x.pdf", {"@1": {"page": 1, "box": [50, top, 120, top + 4.2], "value": "xxxx"}})
    assert not any("罫線" in w for w in r["warnings"]), r["warnings"]


# ====================================================================== マスの並び（1 字ずつ書く欄）
def _comb(tmp_path, cw: float, ch: float, value: str, name: str):
    from test_pdf_places import _fields, _rect, _row, glyphs_pdf

    cells = [_rect(60 + cw * k, 680, 60 + cw * (k + 1), 680 + ch) for k in range(5)]
    blob = glyphs_pdf(_row("番号", 20, 685), extra="\n".join(cells))
    run = next(f for f in _fields(blob) if f["kind"] == "comb")
    r, out = _write(tmp_path, blob, name, {run["id"]: value})
    rect, gs, view = _written(out)
    return r, rect, [(g.text, view.rect(*g.ink[0])) for g in gs if g.ink], gs[0].size


def test_comb_characters_fit_the_cell_width(tmp_path):
    # 幅 5.4 pt・高さ 31 pt の細いマス：マスの高さから決めた大きさでは、字の送り幅がマスより広い → マスの幅に入る大きさにする
    cw = 5.4
    r, rect, inks, size = _comb(tmp_path, cw, 31, "12345", "narrow.pdf")
    assert size == pytest.approx(cw / glyphs.width("1"), abs=0.01) and size < 12
    for k, (ch, b) in enumerate(inks):
        assert rect[0] + k * cw - 0.01 <= b[0] and b[2] <= rect[0] + (k + 1) * cw + 0.01, (ch, b, rect)
    assert not any("欄の四角を出て" in w for w in r["warnings"])


def test_comb_in_wide_cells_is_written_as_before(tmp_path):
    _r, _rect_, inks, size = _comb(tmp_path, 20, 20, "山田太郎", "wide.pdf")
    assert size == 12.0 and [ch for ch, _b in inks] == list("山田太郎")


# ====================================================================== 埋め込まない Helvetica
def test_helvetica_boxes_come_from_the_afm(tmp_path):
    # ş・Č は、同梱の書体になく、Helvetica（埋め込まない）で描く：字形の四角は Adobe の AFM から
    if glyphs.choose("ş")[0] != "TdkL":
        pytest.skip("ş を埋め込む書体で描く環境")
    assert glyphs.helv_box("ş") == (0.032, -0.225, 0.464, 0.538) and glyphs.helv_box(" ") == () and glyphs.helv_box("あ") is None
    assert glyphs.ink_extent("şČ") == (0.929, -0.225)
    r, out = _fill(tmp_path, "şČ", [40, 60, 100, 66])
    assert r["proof_ok"] and r["read_back_ok"]
    _rect, gs, view = _written(out)
    s_, c_ = [g for g in gs if g.text.strip()]
    base = view.pt(s_.ox, s_.oy)[1]
    (sb,), (cb,) = [view.rect(*b) for b in s_.ink], [view.rect(*b) for b in c_.ink]
    assert sb[3] - base == pytest.approx(0.225 * s_.size, abs=0.01), "セディーユは基線より下"
    assert base - cb[1] == pytest.approx(0.929 * c_.size, abs=0.01), "ハーチェクは大文字の上"
    assert sb[2] - sb[0] == pytest.approx(0.432 * s_.size, abs=0.01)


def test_helvetica_widths_of_accented_i():
    # í・ì・î・ï の幅は、点のない i の幅（278）：基の字 i（222）ではない
    assert [glyphs.helv_width(ch) for ch in "íìîïi"] == [278, 278, 278, 278, 222]


# ====================================================================== 書いたものどうし
def _two_lines(tmp_path, pitch: float, upper: str, lower: str, name: str):
    from test_pdf_places import _fields, _row, glyphs_pdf

    # 12 pt の字の 2 行。行の送りが字の高さより狭いと、行の中の空きの四角（高さ 1.2 字ぶん）は上下に重なる
    blob = glyphs_pdf(_row("名前（        ）様", 60, 700) + _row("住所（        ）様", 60, 700 - pitch))
    ids = [f["id"] for f in _fields(blob)]
    assert len(ids) == 2
    return _write(tmp_path, blob, name, dict(zip(ids, (upper, lower))))


def test_descender_touching_the_next_written_line_is_told(tmp_path):
    # 上の行の g・j の下の端が、下の行の字の上の端に少しかかる：断らずに、位置を添えて知らせる
    r, _out = _two_lines(tmp_path, 11.0, "gjpqy", "gjpqy", "a.pdf")
    assert r["proof_ok"] and r["read_back_ok"]
    assert any(re.search("同じ書き込みで書いたものどうしが少しかかる見込みです（1ページ（左 .* mm））", w) for w in r["warnings"])
    # 墨がかからない字（和文）は、何も知らせない
    r, _out = _two_lines(tmp_path, 11.0, "山田", "花子", "b.pdf")
    assert not any("書いたものどうし" in w for w in r["warnings"]), r["warnings"]


def test_written_lines_clearly_over_each_other_are_refused(tmp_path):
    # 下の行の É の上の部分が、上の行の g・j の下の部分にはっきり重なる
    with pytest.raises(api.ApiError, match="同じ書き込みで書いたものどうしが重なります"):
        _two_lines(tmp_path, 11.0, "gjpqy", "ÉÉÉ", "c.pdf")
    # 行の送りが字の高さだけあれば、重ならない
    r, _out = _two_lines(tmp_path, 14.4, "gjpqy", "ÉÉÉ", "d.pdf")
    assert not any("書いたものどうし" in w for w in r["warnings"]), r["warnings"]
