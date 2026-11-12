"""ActiveX の部品の控えの絵（WMF）と、8 ビットの文字コードでは近い字のバイトになる字（〜 U+301C・£ など）。

WMF の字は 8 ビットの文字コードで書く：CP932 のエンコーダは 〜 を ～ のバイトに、£ を ￡ のバイトにする。
前は、こういう字を WMF の控えの絵のある部品に書くと、証明が「控えの絵に、書いた文字が描かれていません」と誤って伝えていた
（部品の値は、頼まれた字のまま書けていた）。いまは、絵の中でどの字になるかを伝え、証明は、その字が描かれていることを確かめる。
"""
import io
import zipfile

import pytest

from test_activex import ax_docx, ax_xlsx, form_controls, morph, word_controls
from todoku import activex as ax
from todoku import api, axdraw


def _docx(value: str):
    from todoku.docx.document import DocxDocument
    from todoku.docx.fill import fill, written_of
    from todoku.docx.proof import prove
    from todoku.docx.window import build

    blob = ax_docx(word_controls())
    res = fill(DocxDocument(blob, "a.docx"), {"TextBox1": value})
    pr = prove(blob, res.blob, written_of(res.patches))
    after = {x["id"]: x["value"] for x in build(DocxDocument(res.blob, "a.docx"))["fields"]}
    z = zipfile.ZipFile(io.BytesIO(res.blob))
    drawn = [t for n in z.namelist() if n.endswith(".wmf") for t in axdraw.wmf_texts(z.read(n))]
    return pr, after["TextBox1"], drawn, res.warnings


@pytest.mark.parametrize("value, drawn, pairs", [
    ("10〜20", "10～20", ["「〜」（U+301C）が「～」（U+FF5E）"]),
    ("£5 ¢ ¬", "￡5 ￠ ￢", ["「£」（U+00A3）が「￡」（U+FFE1）", "「¢」（U+00A2）が「￠」（U+FFE0）", "「¬」（U+00AC）が「￢」（U+FFE2）"]),
    ("a‖b−c", "a∥b－c", ["「‖」（U+2016）が「∥」（U+2225）", "「−」（U+2212）が「－」（U+FF0D）"]),
])
def test_a_near_character_in_a_wmf_preview_is_told_and_proved(value, drawn, pairs):
    pr, after, shown, warnings = _docx(value)
    assert pr.ok, pr.unexpected + pr.lint_errors
    assert after == value  # 部品の値は、頼まれた字のまま
    assert shown == [drawn]
    note = [w for w in warnings if "控えの絵（WMF）" in w]
    assert len(note) == 1 and note[0].startswith("テキストボックス「TextBox1」") and "部品の値は、書いた字のままです" in note[0]
    assert all(p in note[0] for p in pairs)


@pytest.mark.parametrize("value", ["山田　花子", "10～20", "Tanaka 5"])
def test_an_exact_character_in_a_wmf_preview_has_no_note(value):
    pr, after, shown, warnings = _docx(value)
    assert pr.ok and after == value and shown == [value] and not [w for w in warnings if "控えの絵" in w]


def test_the_note_reaches_the_result_of_fill(tmp_path):
    (tmp_path / "a.docx").write_bytes(ax_docx(word_controls()))
    res = api.fill(str(tmp_path / "a.docx"), {"TextBox1": "9 時〜17 時"}, str(tmp_path / "out.docx"))
    assert res["proof_ok"] and res["read_back_ok"] and not res["unexpected"]
    assert any("「〜」（U+301C）が「～」（U+FF5E）" in w for w in res["warnings"])
    assert {f["id"]: f["value"] for f in api.fields(res["out_path"])}["TextBox1"] == "9 時〜17 時"


def test_an_emf_preview_carries_the_character_itself():
    """EMF の字は Unicode：〜 は 〜 のまま描かれ、知らせることはない（Excel の部品の控えの絵）。"""
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove
    from todoku.xlsx.workbook import Workbook

    blob = ax_xlsx(form_controls())
    res = fill(Workbook(blob, "a.xlsx"), {"申請!TextBox1": "10〜20 £5"})
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"})
    assert pr.ok, pr.unexpected + pr.lint_errors
    z = zipfile.ZipFile(io.BytesIO(res.blob))
    assert axdraw.emf_texts(z.read("xl/media/image1.emf")) == ["10〜20 £5"]
    assert not [w for w in res.warnings if "控えの絵" in w]


def test_drawn_as_is_strict_where_the_picture_is_unicode_and_for_other_characters():
    def pic(text: str, wmf: bool) -> bytes:
        cv = axdraw.render(ax.parse(morph(1, text), "TextBox"), 120, 24, 96.0)
        return axdraw.to_wmf(cv, (0, 0, 1800, 360), 1440) if wmf else axdraw.to_emf(cv)

    assert axdraw.drawn_as(pic("10〜20", True), "10〜20") == "10～20" and axdraw.near_drawn(pic("10〜20", True), "10〜20") == [("〜", "～")]
    assert axdraw.drawn_as(pic("10～20", True), "10～20") == "10～20" and axdraw.near_drawn(pic("10～20", True), "10～20") == []
    # WMF でも、近い字でない字は、描かれていない（証明は、前と同じに止める）
    assert axdraw.drawn_as(pic("10～20", True), "10-20") is None and axdraw.drawn_as(pic("山田", True), "鈴木") is None
    assert axdraw.drawn_as(pic("10～20", True), "10～200") is None
    # EMF の字は Unicode：近い字でも、違う字は「描かれていない」
    assert axdraw.drawn_as(pic("10〜20", False), "10〜20") == "10〜20"
    assert axdraw.drawn_as(pic("10～20", False), "10〜20") is None
    # 近い字は、文字コードの表から決まる：全角の ～ は、どの表でも ～ のまま
    assert axdraw._near("～") == frozenset() and "～" in axdraw._near("〜") and axdraw._near("山") == frozenset()
