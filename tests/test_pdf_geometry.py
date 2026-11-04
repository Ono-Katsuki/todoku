"""PDF の窓の書く所を、字の意味（語の表）でなく、字の形と測った空きだけで示す（どの言語の書類にも同じ決め方）。

- 空き：1 字ぶん（1 em）以上の測った空き（行の中では全角の空白か 2 字以上の半角の空白）。前後の字を、そのまま添えて「？」の
  候補にする（中華民國・韓国・中国の日付と住所、点線「Name ......」「Nom : ……」「नाम ......」、下線の字、かっこ「[ ] Yes」）
- 字間を空けた見出し（「氏　　名」）の字の間も、1 字ぶん以上の空きなら所（形は facts の spaced。見出しかはエージェントが見る）
- 右から左へ書く行（アラビア文字）は、読む順の行の頭が右
- フォームの欄（AcroForm）のある書類：本物の欄に重なる・すぐ隣に接する所は、別の所として示さない
- ルビは漢字の行にだけつなぐ。タイ文字の折り返しは空白なしでつなぐ。キリル文字は語の切れ目で折り返す
- 漢字の字形は、書類の言語（/Lang）の字形の集まりから先に選ぶ
"""

from __future__ import annotations

import pytest

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


def test_spaced_headings_are_offered_with_their_own_key():
    assert evenly_spaced("氏　　名") and evenly_spaced("事　業　所") and not evenly_spaced("：　　～　　：")
    # 字間を空けた見出しの形（「氏　　名」と「第　　号」「年　　月　　日」は形が同じ）：形では分けない。どれも 1 字ぶん以上の空き
    assert [x.why for x in gap_blanks("氏　　名")] == ["gap"] and [x.why for x in gap_blanks("事　業　所")] == ["gap"] * 2
    # 行の中に空きのある行は、左右の空きが等しくても、前後の空きも所（中央そろえとして除くのは、行の中に空きのない行だけ）
    assert [x.why for x in gap_blanks("　　　氏　　名　　　", 3, 3)] == ["lead", "gap", "trail"]
    assert [x.why for x in gap_blanks("　　　年　　月　　日", 3, 0)] == ["lead", "gap", "gap"]
    # 点線だけの行も点線の並び（同じ字の 2 字以上）。線の字（罫線素片 U+2500）だけの行は、字で描いた線
    assert [x.why for x in gap_blanks("..............................")] == ["leader"]
    assert gap_blanks("────────────────") == []
    # 1 字だけの点線・ダッシュ（「…」「―」）は句読点・区切りの記号：所にしない
    assert gap_blanks("以下…") == [] and gap_blanks("03―1234") == []
    # 中央にそろえた見出し（左右の空きがほぼ同じで、行の中に空欄がない）：前後の空きは示さない
    assert gap_blanks("　　　申請書　　　", 3, 3) == []
    # 文の終わりの字で終わる行の後ろは、文の終わりの余白
    assert gap_blanks("上記のとおり申請します。　　　", 0, 3) == []


def _row(text: str, x: float, y: float, size: float, em: float, word_gap: float) -> list[tuple[str, float, float]]:
    """字を左から 1 字ずつ置く：字の送りは em（/1000）、空白の所は送りの代わりに word_gap（字の大きさに対する割合）だけ空ける。"""
    out = []
    for ch in text:
        if ch == " ":
            x += word_gap * size
            continue
        out.append((ch, x, y))
        x += em * size / 1000
    return out


def _gap_places(blob: bytes) -> list[dict]:
    return [f for pg in window(blob)["pages"] for f in pg["fields"] if f.get("shape") == "gap"]


def test_word_spaces_in_justified_paragraphs_are_not_places():
    # 両端そろえで広げた語の間の空き（0.6 em まで）は、和字の空白 1 字ぶん（1 em）に届かない：ラテン・キリル・アラビア文字
    paras = ["Applicants must submit the signed form together with proof of residence before the stated deadline",
             "Заявитель должен представить подписанную форму вместе с документом о месте жительства",
             "يجب على مقدم الطلب تقديم النموذج الموقع مع إثبات الإقامة قبل الموعد المحدد"]
    for para in paras:
        words = para.split()
        glyphs = []
        for row in range(4):  # 4 行：語の間を 0.25・0.33・0.5・0.6 em に広げる
            glyphs += _row(" ".join(words[row * 3:row * 3 + 4]), 40, 760 - 16 * row, 10, 500, (0.25, 0.33, 0.5, 0.6)[row])
        blob = glyphs_pdf(glyphs, {ch: 500 for ch in set(para) if ch != " "}, size=10)
        assert not _gap_places(blob), (para[:20], _gap_places(blob))


def test_one_ideographic_space_between_units_is_a_place():
    # 「第　号」「令和　年　月　日」：単位の字の間の、和字の空白 1 字ぶん（U+3000 は 1 em）の空き。空きが U+3000 の字でも、
    # 半角の空白の字の並びでも、字の位置を動かしただけでも、測った幅で「？」の所にする
    lines = ["40 700 m 400 700 l S", "40 640 m 400 640 l S", "40 640 m 40 700 l S", "400 640 m 400 700 l S",
             "40 670 m 400 670 l S", "120 640 m 120 700 l S"]
    g = _row("氏 名", 50, 680, 12, 1000, 1.0)  # 字間を空けた見出し（マスの縁に接する 2 字）：所（形は facts の spaced）
    g += _row("第 号", 300, 680, 12, 1000, 1.0)  # 字の位置を動かしただけの 1 em の空き
    g += [("令", 200, 650), ("和", 212, 650), ("　", 224, 650), ("年", 236, 650)]  # U+3000 の字で描いた空き
    g += [(" ", 248 + 3 * k, 650) for k in range(4)] + [("月", 260, 650)]  # 半角の空白の字（0.25 em）4 つの空き
    g += _row(" 日", 272, 650, 12, 1000, 1.0)
    blob = glyphs_pdf(g, {"　": 1000, " ": 250}, extra="\n".join(lines))
    places = _gap_places(blob)
    hints = sorted((f["facts"]["before"], f["facts"]["after"]) for f in places)
    assert len(places) == 5, hints  # 氏□名・第□号・令和□年・年□月・月□日
    assert all(f["facts"]["width_em"] >= 1 for f in places)
    shi = next(f for f in places if f["facts"]["before"] == "氏")
    assert shi["facts"]["spaced"] and shi["agent"], shi
    # 1 em に届かない空き（0.6 em）は所にしない
    blob = glyphs_pdf(_row("第 号", 300, 680, 12, 1000, 0.6), extra="\n".join(lines))
    assert not _gap_places(blob)
    # プロポーショナルな書体（U+3000 の送りが 0.66 em）：字の大きさに対して 1 em に届かないので、所にしない
    g = [("第", 300, 680), ("　", 312, 680), ("号", 319.92, 680)]
    assert not _gap_places(glyphs_pdf(g, {"　": 660}, extra="\n".join(lines)))


def test_only_closed_curves_are_circles_and_shadowed_squares_are_left_to_the_agent():
    # 曲線だけで閉じた図形（○）は字を囲む○。直線を含む曲線の図形（角の丸い四角）は○としない
    k = 0.5523 * 10
    oval = (f"306 690 m 306 {690 + k} {316 - k} 700 316 700 c {316 + k} 700 326 {690 + k} 326 690 c "
            f"326 {690 - k} {316 + k} 680 316 680 c {316 - k} 680 306 {690 - k} 306 690 c S")
    rounded = "306 682 m 306 698 l 306 699 307 700 308 700 c 324 700 l 325 700 326 699 326 698 c 326 682 l " \
              "326 681 325 680 324 680 c 308 680 l 307 680 306 681 306 682 c S"
    g = [("有", 310, 685), ("・", 322, 685), ("無", 334, 685)]
    assert [c["text"] for c in window(glyphs_pdf(g, extra=oval))["pages"][0]["circled"]] == ["有"]
    # 角の丸い四角は○と決めないが、落とさずに「○かもしれない」所（［○？］）として示す
    pg = window(glyphs_pdf(g, extra=rounded))["pages"][0]
    assert pg["circled"] == [] and [c["text"] for c in pg["maybe_circled"]] == ["有"], pg["maybe_circled"]
    assert pg["maybe_circled"][0]["agent"] and "有［○？］" in "\n".join(pg["reading"]), pg["reading"]
    # 長さ 0 の直線（「x0 y0 l」）で閉じた○（作る道具がよく足す）・対角線の 1% に足りない直線を含む○は、○のまま
    zero = oval.replace(" S", " 306 690 l S")
    tiny = oval.replace("306 690 m", "306 690 m 306 690.1 l 306 690 l")
    for path in (zero, tiny):
        pg = window(glyphs_pdf(g, extra=path))["pages"][0]
        assert [c["text"] for c in pg["circled"]] == ["有"] and not pg["maybe_circled"], (path, pg["circled"])
    # 箱の字はどれも同じ扱い（❏ と □ を使い方で分けない）。字と Unicode の名前は facts に
    pg = window(glyphs_pdf([("❏", 40, 700), ("A", 52, 700), ("□", 40, 680), ("B", 52, 680)]))["pages"][0]
    boxes = {f["label"]: f for f in pg["fields"] if f["kind"] == "checkbox"}
    assert boxes["A"].get("agent") and boxes["B"].get("agent"), boxes  # どの箱の字も「？」（箱か箇条の印かは決めない）
    assert boxes["A"]["facts"]["char_name"].startswith("U+274F") and boxes["B"]["facts"]["char"] == "□"


def test_rtl_line_head_is_on_the_right():
    # 右から左へ書く表：右のマスに見出し「الاسم」、左のマスが空。見出しの左の空き（読む順の行の終わりの後ろ）は、すぐ左の空のマスに書く
    from test_pdf_order import _visual

    lines = ["40 700 m 340 700 l S", "40 680 m 340 680 l S", "40 680 m 40 700 l S", "240 680 m 240 700 l S",
             "340 680 m 340 700 l S"]
    word = _visual("الاسم")
    blob = glyphs_pdf([(ch, 300 + 7 * k, 685) for k, ch in enumerate(word)], {ch: 583 for ch in set(word)}, extra="\n".join(lines))
    w = window(blob)["pages"][0]
    assert w["reading"][0].startswith("الاسم")
    heads = [f for f in w["fields"] if f.get("shape") == "lead"]
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
    fam = [f for f in layout if f.get("shape") == "trail" and f["facts"]["line"].startswith("Family name")]
    # 見出しの右の余白は、欄 family に 1 字ぶん（10pt）に足りない 7.8pt だけかかる：重なりではなく「すぐ隣」。
    # 「？」で理由に欄の名前を添えるが、同じ所の別の読み方（alternative_to）にはしない
    assert fam and not fam[0].get("alternative_to") and fam[0]["agent"] and "フォームの欄 family" in fam[0]["why"], fam
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


@pytest.mark.fonts
def test_wrap_breaks_cyrillic_at_spaces(embedding_font):
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
