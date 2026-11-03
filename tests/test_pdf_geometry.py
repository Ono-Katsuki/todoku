"""PDF の窓の書く所を、字の意味（語の表）でなく、字の形と測った空きだけで示す（どの言語の書類にも同じ決め方）。

- 空き：1.5 字ぶん以上の測った空き（行の中では 2 字以上の空白）。前後の字を、そのまま添えて「？」の候補にする
  （中華民國・韓国・中国の日付と住所、点線「Name ......」「Nom : ……」「नाम ......」、下線の字、かっこ「[ ] Yes」）
- 字間を空けた見出し（「氏　　名」）は、前に空きがないか中央の 2 字のとき、字の間を空欄にしない
- 右から左へ書く行（アラビア文字）は、読む順の行の頭が右
- フォームの欄（AcroForm）のある書類：本物の欄に重なる・すぐ隣に接する所は、別の所として示さない
- ルビは漢字の行にだけつなぐ。タイ文字の折り返しは空白なしでつなぐ。キリル文字は語の切れ目で折り返す
- 漢字の字形は、書類の言語（/Lang）の字形の集まりから先に選ぶ
"""

from __future__ import annotations

from todoku.pdf import glyphs
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import wrap
from todoku.pdf.layout import evenly_spaced, no_word_space, gap_blanks
from todoku.pdf.window import build

from pdf_forms_maker import build_all
from test_pdf_scripts import glyphs_pdf, window


def _kinds(text: str, lead: int = 0, trail: int = 0) -> list[tuple[str, str]]:
    return [(text[a:b], x.why) for x in gap_blanks(text, lead, trail) for a, b in [x]]


def test_gaps_by_geometry_in_any_language():
    assert [w for _t, w in _kinds("中華民國　　年　　月　　日")] == ["gap", "gap", "gap"]
    assert [w for _t, w in _kinds("　　년　　월　　일")] == ["lead", "gap", "gap"]
    assert len(gap_blanks("　　省　　市　　县")) == 3
    assert _kinds("Name ..............................")[0][1] == "leader"
    assert _kinds("Nom : …………………………")[0][1] == "leader"
    assert _kinds("नाम ...................")[0][1] == "leader"
    yes = gap_blanks("[ ] Yes [ ] No")
    assert len(yes) == 2 and all(x.structural and x.why == "bracket" for x in yes)
    assert gap_blanks("Signature ______")[0].structural  # 下線の字（Unicode の Pc）の並び
    assert not gap_blanks("Name____")[0].structural  # 英数字に接した下線の字：名前の一部かもしれない（？）
    assert gap_blanks("الاسم: ____")[0].structural
    assert gap_blanks("Phone ‿‿‿‿")[0].why == "connector"  # 下線の字は「_」「＿」だけでなく Pc の字
    assert gap_blanks("────────────────") == []  # 線の字だけの行（飾り）


def test_spaced_headings_are_not_blanks():
    assert evenly_spaced("氏　　名") and evenly_spaced("事　業　所") and not evenly_spaced("：　　～　　：")
    assert gap_blanks("氏　　名") == [] and gap_blanks("事　業　所") == []
    # 前に空きのある並び（右にそろえた日付）は、字の間も空欄の並び
    assert len(gap_blanks("　　　年　　月　　日", 3, 0)) == 3
    # 中央にそろえた見出し（左右の空きがほぼ同じで、行の中に空欄がない）：前後の空きは示さない
    assert gap_blanks("　　　申請書　　　", 3, 3) == []
    # 文の終わりの字で終わる行の後ろは、文の終わりの余白
    assert gap_blanks("上記のとおり申請します。　　　", 0, 3) == []


def test_rtl_line_head_is_on_the_right():
    # 右から左へ書く表：右のマスに見出し「الاسم」、左のマスが空。見出しの左の空き（読む順の行の終わりの後ろ）は、すぐ左の空のマスに書く
    from test_pdf_order import _visual

    lines = ["40 700 m 340 700 l S", "40 680 m 340 680 l S", "40 680 m 40 700 l S", "240 680 m 240 700 l S",
             "340 680 m 340 700 l S"]
    word = _visual("الاسم")
    blob = glyphs_pdf([(ch, 300 + 7 * k, 685) for k, ch in enumerate(word)], {ch: 583 for ch in set(word)}, extra="\n".join(lines))
    w = window(blob)["pages"][0]
    assert w["reading"][0].startswith("الاسم")
    heads = [f for f in w["fields"] if f["kind"] == "inline" and "行の頭の前" in f.get("why", "")]
    assert not heads, heads  # 行の頭（右）はマスの縁に接している


def test_places_next_to_real_form_fields_are_kept_as_agent_places():
    # 本物のフォームの欄（AcroForm）に重なる・すぐ隣の所（見出し「Family name」の後ろの空き）は、消さずに「？」にして、
    # その欄の名前を添える（同じ種類なら alternative_to：どちらか一方に書く。種類が違えば理由に欄の名前）
    win = build(PdfDocument(build_all(), "a.pdf"))
    pg = win["pages"][0]
    real = {f["id"]: f for f in pg["fields"] if f["source"] == "acroform"}
    layout = [f for f in pg["fields"] if f["source"] == "layout"]
    assert real
    for f in layout:
        for r in real.values():
            h = min(f["rect"][3] - f["rect"][1], r["rect"][3] - r["rect"][1])
            vov = min(f["rect"][3], r["rect"][3]) - max(f["rect"][1], r["rect"][1])
            gap = max(f["rect"][0] - r["rect"][2], r["rect"][0] - f["rect"][2])
            if vov >= 0.5 * h and gap <= r["rect"][3] - r["rect"][1]:
                assert f.get("agent"), (f, r)
                assert f.get("alternative_to") in real or "フォームの欄 " in f["why"], f
    fam = [f for f in layout if f.get("label") == "Family name"]
    assert fam and fam[0]["alternative_to"] == "family" and fam[0]["agent"], fam
    agree = [f for f in layout if "フォームの欄 agree" in f.get("why", "")]
    assert agree and not agree[0].get("alternative_to"), agree  # 余白とチェックの欄：種類が違うので、別の読み方にはしない
    text = "\n".join(pg["reading"])
    assert f"［{fam[0]['id']}？］" in text


def test_ruby_needs_han_and_thai_joins_without_space():
    from todoku.pdf.layout import _ruby_base

    # ルビ（振り仮名・注音）は漢字につく：漢字のない行の上の小さな字（「(office use)」）は、ルビとしてつながない
    assert _ruby_base("葛") and not _ruby_base("N") and not _ruby_base("か")
    # タイ文字・和文は、行の折り返しを空白なしでつなぐ（ハングル・ラテン文字は空白でつなぐ）
    assert no_word_space("ก") and no_word_space("漢") and not no_word_space("a") and not no_word_space("한")


def test_wrap_breaks_cyrillic_at_spaces():
    lines = wrap("Иванов Пётр Сергеевич", 10, 80)
    assert all(not ln.endswith(("-",)) for ln in lines)
    assert all(w in ("Иванов", "Пётр", "Сергеевич") for ln in lines for w in ln.split()), lines


def test_cjk_glyph_set_follows_document_language():
    assert glyphs.order_for_lang("zh-Hant-TW")[0] == "Adobe-CNS1"
    assert glyphs.order_for_lang("zh-TW")[0] == "Adobe-CNS1"
    assert glyphs.order_for_lang("zh-CN")[0] == "Adobe-GB1"
    assert glyphs.order_for_lang("ko-KR")[0] == "Adobe-Korea1"
    assert glyphs.order_for_lang("") == glyphs.ORDER and glyphs.order_for_lang("fr")[0] == "Adobe-Japan1"
    assert glyphs.choose("中")[0] == "TdkJ"
    with glyphs.cjk_order("zh-TW"):
        assert glyphs.choose("中")[0] == "TdkC"
    assert glyphs.choose("中")[0] == "TdkJ"
