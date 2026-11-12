"""PDF のフォームの欄（AcroForm）の見出し：説明（/TU）のない欄の見出しを、所の見出しと同じ決まりで周りの字から選ぶ。
チェックの欄は、同じ行のすぐ後ろの字（箱の字と同じ）。印刷された箱の字の上に重ねた欄は、箱の字を見出しにしない。
書き換えられない字の欄（読み取り専用）に出ている値は、周りの字として見る。

どれも合成した様式で確かめる。実際の様式の数や形に合わせた値は使わない。
"""

from __future__ import annotations

from pdfmini import build as build_pdf, stream, tounicode, type0

from todoku.pdf.doc import PdfDocument
from todoku.pdf.window import build

AP = 4  # 空の見た目（チェック・ひとつ選ぶ欄の、入・切の状態の名前のため）


def _pdf(rows: list[tuple[str, float, float]], fields: list[tuple], extra: str = "", size: float = 10) -> bytes:
    """rows：(字, x, y)。字の幅は 0.5em（CJK は 1em）。fields：
    ("text", 名前, "x0 y0 x1 y1", 足す字) / ("check", 名前, "x0 y0 x1 y1") / ("radio", 名前, [("x0 y0 x1 y1", 入の状態), …])。"""
    chars = sorted({ch for t, _x, _y in rows for ch in t})
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    w = " ".join(f"{code[ch]} [500]" for ch in chars if ord(ch) < 0x2E80)
    content = "\n".join(f"BT /F1 {size} Tf {x} {y} Td <{''.join(f'{code[ch]:04X}' for ch in t)}> Tj ET" for t, x, y in rows)
    objs: list = [type0("TestFont", 3, f"/W [{w}]" if w else ""), stream(tounicode({c: ch for ch, c in code.items()})),
                  stream("", "/Type /XObject /Subtype /Form /BBox [0 0 10 10]")]
    top, annots = [], []
    for f in fields:
        n = len(objs) + 2
        if f[0] == "text":
            objs.append(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T ({f[1]}) /Rect [{f[2]}] /F 4 /DA (/Helv 10 Tf 0 g) "
                        f"{f[3] if len(f) > 3 else ''} >>")
            top.append(n), annots.append(n)
        elif f[0] == "check":
            objs.append(f"<< /Type /Annot /Subtype /Widget /FT /Btn /T ({f[1]}) /Rect [{f[2]}] /F 4 /AS /Off /V /Off "
                        f"/AP << /N << /Yes {AP} 0 R /Off {AP} 0 R >> >> >>")
            top.append(n), annots.append(n)
        else:
            kids = [n + 1 + k for k in range(len(f[2]))]
            objs.append(f"<< /FT /Btn /Ff 32768 /T ({f[1]}) /V /Off /Kids [{' '.join(f'{k} 0 R' for k in kids)}] >>")
            for rect, state in f[2]:
                objs.append(f"<< /Type /Annot /Subtype /Widget /Parent {n} 0 R /Rect [{rect}] /F 4 /AS /Off "
                            f"/AP << /N << /{state} {AP} 0 R /Off {AP} 0 R >> >> >>")
            top.append(n), annots.extend(kids)
    return build_pdf(content + "\n" + extra, "<< /Font << /F1 2 0 R >> >>", objs,
                     annots=f"/Annots [{' '.join(f'{k} 0 R' for k in annots)}]",
                     acroform=f"/AcroForm << /Fields [{' '.join(f'{k} 0 R' for k in top)}] /DA (/Helv 0 Tf 0 g) >>")


def _forms(blob: bytes) -> dict[str, dict]:
    return {f["id"]: f for pg in build(PdfDocument(blob, "t.pdf"))["pages"] for f in pg["fields"] if f.get("source") == "acroform"}


def _box(x0, y0, x1, y1) -> str:
    return f"0.8 w {x0} {y0} {x1 - x0} {y1 - y0} re S"


# ====================================================================== 1. チェックの欄：同じ行のすぐ後ろの字
def test_checkbox_takes_the_text_after_it_on_the_same_line():
    # 「［欄］ Yes　　［欄］ No」：それぞれの欄の後ろの字。1 字ぶん以上の空きの先（次の選択肢）は入れない
    for yes, no in (("Yes", "No"), ("はい", "いいえ"), ("예", "아니오"), ("Да", "Нет")):
        f = _forms(_pdf([(yes, 116, 700), (no, 216, 700)], [("check", "a", "100 698 112 710"), ("check", "b", "200 698 212 710")]))
        assert (f["a"]["label"], f["b"]["label"]) == (yes, no), (f["a"]["label"], f["b"]["label"])


def test_checkbox_with_text_only_before_it_takes_that_text():
    # 「I agree to the terms ［欄］」：後ろに字がなければ、前の字
    f = _forms(_pdf([("I agree to the terms", 100, 700)], [("check", "a", "230 698 242 710")]))
    assert f["a"]["label"] == "I agree to the terms"
    assert f["a"]["neighbours"]["left"]["text"] == "I agree to the terms"


def test_checkbox_between_two_texts_takes_the_text_after_and_says_both():
    # 「Question ［欄］ Option」：後ろの字が見出し。前の字は neighbours に残る
    f = _forms(_pdf([("Question", 100, 700), ("Option", 176, 700)], [("check", "a", "160 698 172 710")]))
    assert f["a"]["label"] == "Option"
    nb = f["a"]["neighbours"]  # 欄をまたいで続く 1 つの行：同じ行の左・右
    assert nb["line_left"]["text"] == "Question" and nb["line_right"]["text"] == "Option", nb


def test_checkbox_in_a_right_to_left_line_takes_the_text_at_its_left():
    # 右から左へ書く行：欄の後ろの字は、欄の左。右の字（質問）ではない
    q, opt = "שאלה", "כן"
    rows = [(ch, 300 - 5 * k, 700) for k, ch in enumerate(q)] + [(ch, 235 - 5 * k, 700) for k, ch in enumerate(opt)]
    f = _forms(_pdf(rows, [("check", "a", "244 698 256 710")]))
    assert f["a"]["label"] == opt, f["a"]["label"]


def test_checkbox_over_a_printed_box_takes_the_option_text_not_the_box():
    # マスの中に「☐ 2026/27」と印刷し、箱の字の上にチェックの欄を重ねた様式：見出しは選択肢の字。箱の字ではない
    for opt in ("2026/27 (first year)", "令和8年度", "제1차"):
        blob = _pdf([("☐", 104, 700), (opt, 120, 700)], [("check", "a", "102 698 114 710")], extra=_box(98, 694, 400, 714))
        f = _forms(blob)["a"]
        assert f["label"] == opt, (opt, f["label"])
    # 選択肢の字が箱の字の前にある様式（「I confirm … ☐」）
    f = _forms(_pdf([("I confirm the above", 104, 700), ("☐", 380, 700)], [("check", "a", "378 698 390 710")],
                    extra=_box(98, 694, 400, 714)))["a"]
    assert f["label"] == "I confirm the above"


def test_checkbox_beside_only_other_boxes_does_not_take_a_box_as_its_label():
    # 上下に並ぶ箱の字（欄が重なる）と、左のマスの選択肢の字：上下の箱の字は見出しにしない
    rows = [("First option", 104, 720), ("☐", 304, 720), ("Second option", 104, 700), ("☐", 304, 700)]
    extra = " ".join(_box(*r) for r in ((98, 714, 298, 734), (298, 714, 320, 734), (98, 694, 298, 714), (298, 694, 320, 714)))
    f = _forms(_pdf(rows, [("check", "a", "302 718 314 730"), ("check", "b", "302 698 314 710")], extra=extra))
    assert (f["a"]["label"], f["b"]["label"]) == ("First option", "Second option")


# ====================================================================== 2. 字の欄：左 → 上 → 右 → 下、離れの上限なし
def test_text_field_takes_the_left_text_however_far():
    # 見出しと欄の間が 200pt より広い様式：左の字が見出し（離れは neighbours に）
    f = _forms(_pdf([("Name", 60, 700)], [("text", "t1", "400 696 540 712")]))["t1"]
    assert f["label"] == "Name"
    assert f["neighbours"]["left"]["text"] == "Name" and f["neighbours"]["left"]["mm"] > 100


def test_text_field_takes_right_or_below_text_when_nothing_is_left_or_above():
    f = _forms(_pdf([("kg", 250, 700)], [("text", "t1", "100 696 240 712")]))["t1"]
    assert f["label"] == "kg"
    f = _forms(_pdf([("Signature", 100, 680)], [("text", "t1", "100 696 240 712")]))["t1"]
    assert f["label"] == "Signature"


def test_text_field_prefers_left_over_above():
    f = _forms(_pdf([("Section heading", 100, 730), ("Name", 60, 700)], [("text", "t1", "100 696 240 712")]))["t1"]
    assert f["label"] == "Name" and f["neighbours"]["above"]["text"] == "Section heading"


def test_text_field_in_the_gap_of_a_line_takes_the_text_before_it_not_the_line_above():
    # 「氏名［欄］年　月　日」：同じ行の、欄より前の字
    f = _forms(_pdf([("申請書", 100, 730), ("氏名", 60, 700), ("年月日", 260, 700)], [("text", "t1", "100 696 240 712")]))["t1"]
    assert f["label"] == "氏名"


def test_widget_over_a_thin_cell_uses_its_own_surroundings():
    # 2 本の罫線の間の細いマスに中心のある、大きな欄：欄はそのマスの中にはない。マスの上の字ではなく、欄の左の字
    extra = _box(100, 650, 400, 653)
    f = _forms(_pdf([("Above the rules", 100, 760), ("Remarks", 40, 648)], [("text", "t1", "100 600 400 700")], extra=extra))["t1"]
    assert f["label"] == "Remarks", f["label"]


def test_widget_in_an_empty_cell_takes_the_text_of_the_cell_at_its_left():
    extra = _box(60, 694, 160, 714) + " " + _box(160, 694, 400, 714)
    for lab in ("Address", "住所", "주소", "Адрес"):
        f = _forms(_pdf([(lab, 64, 700)], [("text", "t1", "164 696 396 712")], extra=extra))["t1"]
        assert f["label"] == lab


# ====================================================================== 3. 書き換えられない字の欄の値
RO = "/Ff 1 /V (Funding ID)"


def test_read_only_field_value_is_the_label_of_the_field_beside_it():
    # 見出しを、印刷された字ではなく読み取り専用の欄の値として置く様式：右の書く欄の見出しは、その値。どの欄の値かも示す
    f = _forms(_pdf([("Form", 60, 760)], [("text", "lab", "60 696 160 712", RO), ("text", "inp", "200 696 400 712")]))
    assert f["inp"]["label"] == "Funding ID"
    assert f["inp"]["neighbours"]["left"] == {"text": "Funding ID", "mm": f["inp"]["neighbours"]["left"]["mm"], "field": "lab"}
    assert 13 < f["inp"]["neighbours"]["left"]["mm"] < 15  # 40pt
    # 上に置いた読み取り専用の欄
    f = _forms(_pdf([], [("text", "lab", "200 720 400 736", RO), ("text", "inp", "200 696 400 712")]))
    assert f["inp"]["label"] == "Funding ID" and f["inp"]["neighbours"]["above"]["field"] == "lab"


def test_writable_or_hidden_field_values_are_not_surrounding_text():
    # 書ける欄の値（前に書いた字）・画面にも紙にも出ない読み取り専用の欄の値は、見出しにしない
    f = _forms(_pdf([], [("text", "lab", "60 696 160 712", "/V (Funding ID)"), ("text", "inp", "200 696 400 712")]))
    assert f["inp"]["label"] == "inp" and "neighbours" not in f["inp"]
    hidden = "/Ff 1 /V (Funding ID) /F 2"
    f = _forms(_pdf([], [("text", "lab", "60 696 160 712", hidden), ("text", "inp", "200 696 400 712")]))
    assert f["inp"]["label"] == "inp"


def test_read_only_label_fields_do_not_label_each_other():
    # 並んだ読み取り専用の欄（見出しの並び）：隣の見出しの値を、見出しの欄の見出しにしない。書く欄は、いちばん近い値
    f = _forms(_pdf([], [("text", "h1", "60 696 160 712", "/Ff 1 /V (Activity)"), ("text", "h2", "164 696 264 712", "/Ff 1 /V (Court)"),
                         ("text", "inp", "164 670 264 686")]))
    assert f["h2"]["label"] == "h2" and f["h2"]["value"] == "Court"
    assert f["inp"]["label"] == "Court" and f["inp"]["neighbours"]["above"]["field"] == "h2"


def test_calculated_read_only_value_is_not_a_label():
    # 値を計算で出す読み取り専用の欄（合計）：値は見出しではない
    calc = "/Ff 1 /V (1,200) /AA << /C << /S /JavaScript /JS (AFSimple_Calculate\\(\"SUM\", new Array\\(\"inp\"\\)\\);) >> >>"
    f = _forms(_pdf([], [("text", "total", "60 696 160 712", calc), ("text", "inp", "200 696 400 712")]))
    assert f["inp"]["label"] == "inp" and "neighbours" not in f["inp"]
    assert f["total"]["calculated"]


def test_nearer_printed_text_wins_over_a_read_only_value_on_the_same_side():
    f = _forms(_pdf([("Printed", 160, 700)], [("text", "lab", "40 696 140 712", RO), ("text", "inp", "200 696 400 712")]))
    assert f["inp"]["label"] == "Printed" and "field" not in f["inp"]["neighbours"]["left"]


def test_read_only_value_nearer_than_the_same_line_text_wins():
    # 同じ高さの行の字（節の見出し）より、欄のすぐ左の読み取り専用の欄の値が近い
    f = _forms(_pdf([("2: REQUEST", 40, 700)], [("text", "lab", "150 696 250 712", RO), ("text", "inp", "260 696 400 712")]))
    assert f["inp"]["label"] == "Funding ID"


def test_checkbox_takes_a_read_only_value_at_its_left():
    f = _forms(_pdf([], [("text", "lab", "60 696 260 712", "/Ff 1 /V (Please tick to confirm)"), ("check", "c", "280 698 292 710")]))
    assert f["c"]["label"] == "Please tick to confirm"


# ====================================================================== 4. ひとつ選ぶ欄の組の見出し
def test_radio_group_takes_the_question_in_the_cell_at_its_left():
    # 質問の字がマスの中、選択肢の部品がその右：組の見出しは質問の字（前は、マスの中の字を探さず、欄の名前になっていた）
    extra = _box(60, 690, 300, 716) + " " + _box(300, 690, 500, 716)
    rows = [("Is your project within scope?", 64, 700), ("Yes", 330, 700), ("No", 410, 700)]
    f = _forms(_pdf(rows, [("radio", "Group9", [("312 698 324 710", "Yes"), ("392 698 404 710", "No")])], extra=extra))["Group9"]
    assert f["label"] == "Is your project within scope?", f["label"]
    assert f["options_labels"] == ["Yes", "No"]


def test_radio_group_takes_the_line_above_however_far_when_nothing_is_left():
    rows = [("Organisation type", 100, 760), ("Company", 116, 700), ("Charity", 216, 700)]
    f = _forms(_pdf(rows, [("radio", "g", [("100 698 112 710", "a"), ("200 698 212 710", "b")])]))["g"]
    assert f["label"] == "Organisation type" and f["options_labels"] == ["Company", "Charity"]


def test_radio_group_without_a_question_keeps_its_name_not_an_option_text():
    f = _forms(_pdf([("Male", 116, 700), ("Female", 216, 700)], [("radio", "sex", [("100 698 112 710", "m"), ("200 698 212 710", "f")])]))
    assert f["sex"]["label"] == "sex" and f["sex"]["options_labels"] == ["Male", "Female"]


def test_radio_group_in_the_same_cell_as_its_question_keeps_the_cell_text():
    # 質問と選択肢が同じマスにあり、マスの左にも上にも字がない：マスの字（質問と選択肢）のまま。欄の名前にはしない
    rows = [("Organisation type:", 64, 700), ("Company", 216, 700), ("Charity", 316, 700)]
    f = _forms(_pdf(rows, [("radio", "g", [("200 698 212 710", "a"), ("300 698 312 710", "b")])], extra=_box(60, 690, 500, 716)))["g"]
    assert f["options_labels"] == ["Company", "Charity"]
    assert f["label"].startswith("Organisation type:"), f["label"]


def test_single_char_option_does_not_swallow_the_first_char_of_the_next_word():
    # 「［欄］ 可　　不可」：1 字の選択肢の字の後ろの、次の語の初めの 1 字までを、字間を空けた 1 つの見出し（「氏　　名」）と読まない
    from todoku.pdf.layout import _option_text

    assert _option_text("可　　不可") == "可"
    assert _option_text("예          아니오") == "예"
    assert _option_text("氏　　名") == "氏　　名" and _option_text("事　業　所　　ほか") == "事　業　所"


def test_tooltip_stays_the_label_and_the_near_text_is_told_beside_it():
    f = _forms(_pdf([("Name", 60, 700)], [("text", "t1", "100 696 240 712", "/TU (Full legal name)")]))["t1"]
    assert f["label"] == "Full legal name" and f["nearby_text"] == "Name"
