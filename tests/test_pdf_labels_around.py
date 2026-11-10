"""PDF の所の周りの字：下の字も見出しにし、1 字ぶん以上の空きで分かれた行は、上下では分かれた部分ごとに見る。
所をまたいで続く行（所が、行の中の空きにある）の前後の部分は、左右とは別に「同じ行の左・右」として示す。

どれも合成した様式（いくつかの言語・いくつかの組み方）で確かめる。実際の様式の数や形に合わせた値は使わない。
"""

from __future__ import annotations

from todoku.pdf.doc import PdfDocument
from todoku.pdf.window import build, to_markdown

from test_pdf_places import _fields, _rect, _row
from test_pdf_scripts import glyphs_pdf

RULE = "0.8 w 100 600 m 300 600 l S"  # 左 100〜300 の横線（下線の罫線）
RULE2 = "0.8 w 350 600 m 500 600 l S"


def _sides(f: dict) -> dict:
    return {k: v for k, v in f["facts"].items() if k in ("left", "above", "right", "below", "line_left", "line_right")}


def _rules(blob: bytes) -> list[dict]:
    return sorted((f for f in _fields(blob) if f.get("shape") == "underline"), key=lambda f: f["facts"]["rect_mm"][0])


def _cells(blob: bytes) -> list[dict]:
    return sorted((f for f in _fields(blob) if f.get("shape") == "cell"), key=lambda f: f["facts"]["rect_mm"][0])


# ====================================================================== 1. 下線の罫線：下の字
def test_rule_with_a_caption_only_below_takes_it_as_the_label():
    # 線の下に見出しを書く様式（署名の線）：どの言語でも、下の字が見出しになり、離れと一緒に facts にある
    for cap in ("Signature", "署名", "서명", "Подпись", "Υπογραφή", "חתימה"):
        rtl = cap[0] >= "\u0590" and cap[0] <= "\u08ff"  # 右から左へ書く字は、右から順に置く
        glyphs = [(ch, 280 - 12 * k, 585) for k, ch in enumerate(cap)] if rtl else _row(cap, 100, 585)
        (f,) = _rules(glyphs_pdf(glyphs, extra=RULE))
        assert f["label"] == cap, (cap, f["label"])
        assert set(_sides(f)) == {"below"} and set(f["facts"]["below"]) == {"text", "mm"}, _sides(f)
        assert f["facts"]["below"]["text"] == cap and 0 < f["facts"]["below"]["mm"] < 5


def test_rule_with_text_above_and_below_does_not_choose_between_them():
    # 上にも下にも字がある線：どちらが見出しかは決めない（どちらの字も、離れと一緒に示す）
    blob = glyphs_pdf(_row("Name", 100, 630) + _row("Address", 100, 585), extra=RULE)
    (f,) = _rules(blob)
    assert f["label"] == ""
    assert f["facts"]["above"]["text"] == "Name" and f["facts"]["below"]["text"] == "Address"
    assert f["facts"]["above"]["mm"] > f["facts"]["below"]["mm"] > 0
    md = to_markdown(build(PdfDocument(blob, "t.pdf")))
    line = next(ln for ln in md.split("\n") if ln.startswith(f"- ［{f['id']}］"))
    assert "上「Name」" in line and "下「Address」" in line, line


def test_rule_keeps_the_left_text_as_label_and_shows_the_note_below():
    blob = glyphs_pdf(_row("Date", 40, 602) + _row("(dd/mm/yyyy)", 100, 585), extra=RULE)
    (f,) = _rules(blob)
    assert f["label"] == "Date" and f["facts"]["left"]["text"] == "Date"
    assert f["facts"]["below"]["text"] == "(dd/mm/yyyy)"
    assert f["size_from"] == "線の左の字の大きさ"


def test_rule_with_only_text_on_the_right_takes_it_as_the_label():
    (f,) = _rules(glyphs_pdf(_row("円", 310, 602), extra=RULE))
    assert f["label"] == "円" and set(_sides(f)) == {"right"}


def test_rule_with_nothing_around_has_no_label():
    (f,) = _rules(glyphs_pdf(_row("A", 400, 100), extra=RULE))
    assert f["label"] == "" and not _sides(f)


# ====================================================================== 2. 空のマス：下の字
def test_empty_cell_with_a_caption_only_below_takes_it_as_the_label():
    for cap in ("Stamp", "印", "도장"):
        (f,) = _cells(glyphs_pdf(_row(cap, 110, 560), extra=_rect(100, 580, 200, 640)))
        assert f["label"] == cap and set(_sides(f)) == {"below"}, (cap, f["label"], _sides(f))
        # 下の字（マスの下に小さく添えた字のことがある）は、書く字の大きさの元にしない
        assert f["size_from"] == "ページのふつうの字の大きさ"


def test_empty_cell_prefers_left_then_above_then_right_over_below():
    box = _rect(100, 580, 200, 640)
    below = _row("Below", 110, 560)
    for extra_rows, want in ((_row("L", 80, 600), "L"), (_row("Up", 110, 650), "Up"), (_row("R", 210, 600), "R")):
        (f,) = _cells(glyphs_pdf(below + extra_rows, extra=box))
        assert f["label"] == want and f["facts"]["below"]["text"] == "Below", (want, f["label"])


# ====================================================================== 3. 1 字ぶん以上の空きで分かれた行
def test_one_caption_line_below_two_rules_is_read_by_its_pieces():
    blob = glyphs_pdf(_row("Signature", 100, 585) + _row("Date", 350, 585), extra=RULE + "\n" + RULE2)
    a, b = _rules(blob)
    assert a["label"] == "Signature" and b["label"] == "Date"
    for f, cap in ((a, "Signature"), (b, "Date")):
        below = f["facts"]["below"]
        assert below["text"] == cap and below["line"].split() == ["Signature", "Date"], below


def test_one_right_to_left_caption_line_below_two_rules_is_read_by_its_pieces():
    # 右から左へ書く行（字の順と左右の順が逆）も、部分の左右は字の位置で決める
    sig, name = "חתימה", "שם"
    glyphs = [(ch, 280 - 12 * k, 585) for k, ch in enumerate(sig)] + [(ch, 480 - 12 * k, 585) for k, ch in enumerate(name)]
    a, b = _rules(glyphs_pdf(glyphs, extra=RULE + "\n" + RULE2))
    assert (a["label"], b["label"]) == (sig, name)
    assert a["facts"]["below"]["line"] == b["facts"]["below"]["line"] != sig


def test_rule_inside_one_line_shows_the_pieces_before_and_after_it():
    # 「Name: ____ Date: ____」：字は 1 つの行。1 本目の線は行の途中にある（行が線をまたいで続く）
    rules = "0.8 w 125 600 m 300 600 l S\n0.8 w 385 600 m 520 600 l S"
    for title in ([], _row("Application", 60, 650)):  # 上に題の行があっても、名前は同じ行の字
        blob = glyphs_pdf(title + _row("Name:", 60, 602) + _row("Date:", 320, 602), extra=rules)
        a, b = _rules(blob)
        # 前後のどちらが見出しかは決めない：行の中の空きの所と同じに、所の印をはさんで示す
        assert a["label"] == "Name:［　］Date:", a["label"]
        assert a["facts"]["line_left"]["text"] == "Name:" and a["facts"]["line_right"]["text"] == "Date:"
        assert a["facts"]["line_left"]["line"].split() == ["Name:", "Date:"]
        assert a["facts"]["line_left"]["mm"] < a["facts"]["line_right"]["mm"]
        assert "left" not in a["facts"] and "right" not in a["facts"]
        # 2 本目の線は行の後ろ：左の字は、行の全体（どの部分が見出しかは決めない）
        assert b["label"].split() == ["Name:", "Date:"] and b["facts"]["left"]["text"].split() == ["Name:", "Date:"]
        assert set(_sides(b)) == {"left"}
        md = to_markdown(build(PdfDocument(blob, "t.pdf")))
        line = next(ln for ln in md.split("\n") if ln.startswith(f"- ［{a['id']}］"))
        assert "同じ行の左「Name:」" in line and "同じ行の右「Date:」" in line, line


def test_rule_stays_a_place_when_a_larger_line_continues_past_it():
    # 線をまたいで続く行の字が大きくても、線の上の帯（1 行ぶん）は、ページのふつうの字の大きさで見る：
    # 大きな字の 1 行ぶんの中にある小さな注の字で、所がなくならない
    big = [(ch, 40 + 24 * k, 602) for k, ch in enumerate("Name:")] + [(ch, 420 + 24 * k, 602) for k, ch in enumerate("Date:")]
    code = {ch: i + 1 for i, ch in enumerate(sorted({ch for ch, _x, _y in big}))}

    def small(text: str, x: float, y: float) -> str:  # 6pt の字（ページの字の多く）
        return "\n".join(f"BT /F1 6 Tf {x + 6 * k:.3f} {y:.3f} Td <{code[ch]:04X}> Tj ET" for k, ch in enumerate(text))

    extra = "0.8 w 170 600 m 400 600 l S\n" + small("ate", 200, 614) + "\n" + small("NameDateNameDateNameDate", 40, 100)
    (f,) = _rules(glyphs_pdf(big, size=24, extra=extra))
    assert f["size"] == 6 and f["size_from"] == "ページのふつうの字の大きさ"
    assert f["facts"]["above"]["text"] == "ate" and f["label"] == "Name:［　］Date:"


BOXES = _rect(140, 665, 170, 695) + "\n" + _rect(200, 665, 230, 695)  # 行の中の空きに置いた 2 つのマス


def test_cell_inside_a_line_keeps_the_row_header_and_shows_the_line_apart():
    # 「変更 |［　］年［　］月」：見出しは左のマスの字のまま。またいで続く行の前後の字は、別の事実
    blob = glyphs_pdf(_row("変更", 50, 675) + [("年", 178, 675), ("月", 238, 675)], extra=_rect(40, 660, 120, 700) + "\n" + BOXES)
    a, b = _cells(blob)
    assert a["label"] == b["label"] == "変更" and b["facts"]["left"]["text"] == "変更"
    assert b["facts"]["line_left"]["text"] == "年" and b["facts"]["line_right"]["text"] == "月"
    assert b["facts"]["line_left"]["line"].replace("　", "") == "年月"
    assert b["size_from"] == "左の字の大きさ"
    # 1 つ目のマスは行の左の外：行の全体が右の字
    assert a["facts"]["right"]["text"].replace("　", "") == "年月" and "line_right" not in a["facts"]


def test_cell_inside_a_line_with_no_other_text_is_named_by_the_line_around_it():
    # 同じ空きでつないだ「年　月」も、その空きに所があれば、前後に分けて示す。書く字の大きさは、その行の字
    for a_, b_ in (("年", "月"), ("년", "월")):
        _first, f = _cells(glyphs_pdf([(a_, 178, 675), (b_, 238, 675)], extra=BOXES))
        assert f["label"] == f"{a_}［　］{b_}", f["label"]
        assert set(_sides(f)) == {"line_left", "line_right"} and f["size_from"] == "同じ行の字の大きさ"


def test_cell_inside_a_right_to_left_line_reads_the_pieces_in_reading_order():
    # 右から左へ書く行：前の部分は、所の右
    _first, f = _cells(glyphs_pdf([("\u05d0", 178, 675), ("\u05d1", 238, 675)], extra=BOXES))
    assert f["facts"]["line_right"]["text"] == "\u05d1" and f["facts"]["line_left"]["text"] == "\u05d0"
    assert f["label"] == "\u05d1［　］\u05d0"


def test_cell_prefers_text_above_over_the_line_around_it():
    blob = glyphs_pdf(_row("Title", 200, 720) + [("年", 178, 675), ("月", 238, 675)], extra=BOXES)
    _first, f = _cells(blob)
    assert f["label"] == "Title" and f["facts"]["line_left"]["text"] == "年" and f["size_from"] == "上の字の大きさ"


def test_header_line_above_cells_is_read_by_its_pieces():
    for heads in (("Name", "Date"), ("氏名", "日付"), ("Имя", "Дата")):
        rows = _row(heads[0], 105, 705) + _row(heads[1], 305, 705)
        a, b = _cells(glyphs_pdf(rows, extra=_rect(100, 660, 300, 700) + "\n" + _rect(300, 660, 500, 700)))
        assert (a["label"], b["label"]) == heads
        assert a["facts"]["above"]["text"] == heads[0] and a["facts"]["above"]["line"].split() == list(heads)
        assert a["size_from"] == "上の字の大きさ"


def test_caption_line_below_two_boxes_is_read_by_its_pieces():
    rows = _row("Stamp", 105, 645) + _row("Sign", 305, 645)
    a, b = _cells(glyphs_pdf(rows, extra=_rect(100, 660, 250, 700) + "\n" + _rect(300, 660, 450, 700)))
    assert (a["label"], b["label"]) == ("Stamp", "Sign")
    assert a["facts"]["below"]["line"] == b["facts"]["below"]["line"]


def test_line_left_of_a_cell_is_shown_whole():
    # 左の行は、分かれていても全体を示す（どの部分が見出しかは決めない：「フリガナ　○○ショウジ」の近いほうは、見本の値）
    (f,) = _cells(glyphs_pdf(_row("申請者", 40, 675) + _row("住所", 110, 675), extra=_rect(140, 660, 400, 700)))
    left = f["facts"]["left"]
    assert f["label"].split() == ["申請者", "住所"] and left["text"].split() == ["申請者", "住所"] and "line" not in left, left


def test_piece_that_does_not_reach_the_place_falls_back_to_the_whole_line():
    # 上の行のどの部分も、マスの左右の範囲に重ならない（空きの真下のマス）：行の全体を示す
    rows = _row("AB", 100, 705) + _row("CD", 400, 705)
    (f,) = _cells(glyphs_pdf(rows, extra=_rect(200, 660, 300, 700)))
    above = f["facts"]["above"]
    assert above["text"].split() == ["AB", "CD"] and "line" not in above


# ====================================================================== 4. 分けない行
def test_ordinary_word_gaps_do_not_split_a_line():
    cap = "Full name of the applicant"
    (f,) = _rules(glyphs_pdf(_row(cap, 100, 585, 6), {ch: 500 for ch in set(cap)}, extra=RULE))
    assert f["label"] == cap and "line" not in f["facts"]["below"]


def test_evenly_spaced_label_stays_one_label():
    # 字間を空けた「氏　　名」は 1 つの見出し（部分に分けない）
    for a, b in (("氏", "名"), ("성", "명")):
        (f,) = _cells(glyphs_pdf(_row(a, 105, 705) + _row(b, 165, 705), extra=_rect(100, 660, 300, 700)))
        above = f["facts"]["above"]
        assert above["text"].replace("　", "").replace(" ", "") == a + b and "line" not in above


# ====================================================================== 5. 書いた後に伝える字
def test_say_to_user_names_a_rule_by_the_caption_below(tmp_path):
    from todoku import api

    src = tmp_path / "a.pdf"
    src.write_bytes(glyphs_pdf(_row("Signature", 100, 585), extra=RULE))
    fid = next(f["id"] for f in api.fields(str(src)) if f.get("shape") == "underline")
    r = api.fill(str(src), {fid: "Yamada"}, str(tmp_path / "b.pdf"))
    assert f"{fid}（近くの字「Signature」）：Yamada" in r["say_to_user"], r["say_to_user"]
