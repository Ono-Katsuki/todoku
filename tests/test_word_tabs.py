"""タブで作った空欄（Word・RTF・OpenDocument の文書・HWPX）：下線を引いたタブ・かっこの中のタブ・線（リーダー）を引く
タブ位置のある段落のタブを、所として示す。書くときはタブを残し、値をタブの前に、タブと同じ書式で差し込む
（タブを消すと、書類の線が短くなる）。手で組んだ文書で、読む・書く・証明・読み戻しを確かめる。"""

from __future__ import annotations

import io
import zipfile

import pytest
from test_docx import W, make_docx
from test_hwpx import NS as HWPX_NS
from test_hwpx import XML as HWPX_XML
from test_hwpx import make_hwpx, para, run_, t
from test_openformats import B, make_odt, make_rtf, sj

from todoku import api

U = '<w:rPr><w:u w:val="single"/></w:rPr>'
VALUE = "山田太郎"


@pytest.fixture(autouse=True)
def clean():
    api._WINDOWS.clear()
    yield
    api._WINDOWS.clear()


def r(text: str, u: bool = False) -> str:
    return f'<w:r>{U if u else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


def tab(u: bool = False, n: int = 1) -> str:
    return f'<w:r>{U if u else ""}<w:tab/></w:r>' * n


def wp(*runs: str, ppr: str = "") -> str:
    return f"<w:p>{ppr}{''.join(runs)}</w:p>"


def _fields(path) -> dict:
    api._WINDOWS.clear()
    return {f["id"]: f for f in api.fields(str(path))}


def _lines(path) -> list[str]:
    api._WINDOWS.clear()
    win = api.window(str(path))[1]
    return [ln for st in win["stories"] for ln in st["reading"]]


def _fill(src, out, values: dict) -> dict:
    res = api.fill(str(src), values, str(out))
    assert res["proof_ok"] and res["read_back_ok"], res
    return res


def _doc_xml(path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


# ====================================================================== Word：下線を引いたタブ
def test_docx_underlined_tabs_are_a_rule_blank_and_stay_after_the_value(tmp_path):
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(wp(r("事業所名称："), tab(True, 4), r("　　", True), r("印"))))
    fl = _fields(src)
    assert list(fl) == ["P1.c6"]
    f = fl["P1.c6"]
    assert not f.get("agent") and f["facts"]["tabs"] == 4 and f["facts"]["len"] == 0
    assert _lines(src)[0].startswith("事業所名称：［P1.c6］") and _lines(src)[0].endswith("　　印")  # 印は、値の書かれる位置（タブの前）
    res = _fill(src, tmp_path / "o.docx", {"P1.c6": VALUE})
    assert any("タブ 4 個" in w and "タブは残しました" in w for w in res["warnings"])
    xml = _doc_xml(tmp_path / "o.docx")
    assert xml.count("<w:tab/>") == 4  # タブは 1 つも消さない
    # 値は、タブと同じ下線つきの字の要素で、初めのタブの前
    assert f'<w:r>{U}<w:t xml:space="preserve">{VALUE}</w:t></w:r>' in xml
    assert xml.index(VALUE) < xml.index("<w:tab/>")
    # 書いた後：残したタブ（とその後ろの下線つきの空白）は、書いた所の線の残り。新しい空欄として出さない
    after = _fields(tmp_path / "o.docx")
    assert list(after) == ["P1.c6"] and after["P1.c6"]["value"] == VALUE


def test_docx_rest_of_a_written_line_is_not_offered_again(tmp_path):
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(wp(r("Name:"), tab(n=2), r(" Date:"), tab(), ppr=LEADER) + wp(r("電話（"), tab(), r("　"), tab(), r("）まで。"))))
    assert list(_fields(src)) == ["P1.c5", "P1.c10", "P2.c3"]
    _fill(src, tmp_path / "o.docx", {"P1.c5": "Ono", "P2.c3": "03-0000"})
    after = _fields(tmp_path / "o.docx")
    # 書いた所の後ろのタブ・空白は出さない（「Date:」の後ろのタブは、書いていない所なので残る）
    assert sorted(after) == ["P1.c10", "P1.c5", "P2.c3"]
    assert after["P1.c5"]["value"] == "Ono" and after["P2.c3"]["value"] == "03-0000" and after["P1.c10"].get("agent")


def test_docx_spaces_then_underlined_tabs_write_only_the_spaces(tmp_path):
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(wp(r("Signature: "), r("    ", True), tab(True, 2), r(" Date: "), tab(True))))
    fl = _fields(src)
    # 空白 4 つ＋タブ 2 つ（6 字）は規則の空欄、タブ 1 つだけは「？」
    assert len(fl) == 2
    sig = next(k for k, f in fl.items() if f["facts"]["tabs"] == 2)
    day = next(k for k, f in fl.items() if f["facts"]["tabs"] == 1)
    assert not fl[sig].get("agent") and fl[sig]["facts"]["printed"].strip(" ") == ""
    assert fl[day]["agent"] and fl[day]["why_code"] == "underlined_tab"
    _fill(src, tmp_path / "o.docx", {sig: "Taro", day: "2026-11-11"})
    xml = _doc_xml(tmp_path / "o.docx")
    assert xml.count("<w:tab/>") == 3
    got = _fields(tmp_path / "o.docx")
    assert got[sig]["value"] == "Taro" and got[day]["value"] == "2026-11-11"


def test_docx_two_tab_blanks_in_one_run_are_both_written(tmp_path):
    # 1 つの字の要素の中に、タブが 2 か所（あいだに字）：どちらにも書ける
    src = tmp_path / "a.docx"
    body = f'<w:p><w:r>{U}<w:tab/></w:r><w:r><w:t xml:space="preserve"> / </w:t></w:r><w:r>{U}<w:tab/><w:tab/><w:tab/></w:r></w:p>'
    src.write_bytes(make_docx(body))
    fl = _fields(src)
    assert len(fl) == 2
    one = next(k for k, f in fl.items() if f["facts"]["tabs"] == 1)
    three = next(k for k, f in fl.items() if f["facts"]["tabs"] == 3)
    _fill(src, tmp_path / "o.docx", {one: "11", three: "2026"})
    xml = _doc_xml(tmp_path / "o.docx")
    assert xml.count("<w:tab/>") == 4
    got = _fields(tmp_path / "o.docx")
    assert got[one]["value"] == "11" and got[three]["value"] == "2026"


def test_docx_plain_tab_is_not_a_blank(tmp_path):
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(wp(r("氏名"), tab(), r("山田"), tab(), r("年齢"), tab(), r("30。"))))
    assert not _fields(src)


# ====================================================================== Word：かっこの中のタブ
def test_docx_bracket_holding_only_tabs_is_a_rule_blank(tmp_path):
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(wp(r("電話（"), tab(), r("　"), tab(), r("）まで。"))))
    fl = _fields(src)
    assert list(fl) == ["P1.c3"] and not fl["P1.c3"].get("agent") and fl["P1.c3"]["facts"]["tabs"] == 2
    _fill(src, tmp_path / "o.docx", {"P1.c3": "03-0000-0000"})
    xml = _doc_xml(tmp_path / "o.docx")
    assert xml.count("<w:tab/>") == 2 and xml.index("03-0000-0000") < xml.index("<w:tab/>")
    assert _fields(tmp_path / "o.docx")["P1.c3"]["value"] == "03-0000-0000"


# ====================================================================== Word：線（リーダー）を引くタブ位置
LEADER = '<w:pPr><w:tabs><w:tab w:val="right" w:leader="underscore" w:pos="8000"/></w:tabs></w:pPr>'


def _docx_with_styles(body: str, styles: str) -> bytes:
    buf = io.BytesIO(make_docx(body))
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:styles {W}>{styles}</w:styles>')
    return buf.getvalue()


def test_docx_leader_tab_is_an_agent_place_with_the_leader_kinds(tmp_path):
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(wp(r("Name:"), tab(n=2), r(" Date:"), tab(), ppr=LEADER)))
    fl = _fields(src)
    assert set(fl) == {"P1.c5", "P1.c10"}  # 続けて並ぶタブは 1 つの所
    f = fl["P1.c5"]
    assert f["agent"] and f["why_code"] == "leader_tab"
    assert f["facts"]["leaders"] == ["underscore"] and f["facts"]["tabs"] == 2
    res = _fill(src, tmp_path / "o.docx", {"P1.c5": " Taro"})
    assert any("タブは残しました" in w for w in res["warnings"])
    xml = _doc_xml(tmp_path / "o.docx")
    assert xml.count("<w:tab/>") == 3 and xml.index(" Taro") < xml.index("<w:tab/>")


def test_docx_leader_from_a_style_chain_and_cleared_stops(tmp_path):
    styles = ('<w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/>'
              '</w:tabs></w:pPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="Line"><w:basedOn w:val="Base"/></w:style>'
              '<w:style w:type="paragraph" w:styleId="Plain"><w:basedOn w:val="Base"/><w:pPr><w:tabs>'
              '<w:tab w:val="clear" w:pos="9000"/><w:tab w:val="left" w:pos="2000"/></w:tabs></w:pPr></w:style>')
    body = (wp(r("第 1 章"), tab(), r("3"), ppr='<w:pPr><w:pStyle w:val="Line"/></w:pPr>')
            + wp(r("氏名"), tab(), r("山田。"), ppr='<w:pPr><w:pStyle w:val="Plain"/></w:pPr>')
            + wp(r("年齢"), tab(), r("30。"), ppr='<w:pPr><w:pStyle w:val="Line"/><w:tabs><w:tab w:val="clear" w:pos="9000"/></w:tabs></w:pPr>'))
    src = tmp_path / "a.docx"
    src.write_bytes(_docx_with_styles(body, styles))
    fl = _fields(src)
    # 基のスタイルの線つきのタブ位置。消した段落・スタイルには、タブの所がない
    assert [k for k, f in fl.items() if f.get("why_code") == "leader_tab"] == ["P1.c3"]
    assert fl["P1.c3"]["facts"]["leaders"] == ["dot"]
    assert not any("tabs" in (f.get("facts") or {}) for k, f in fl.items() if k != "P1.c3")


# ====================================================================== RTF
def _w(word: str) -> str:
    return B + word


def test_rtf_underlined_tabs_are_kept(tmp_path):
    src = tmp_path / "a.rtf"
    src.write_bytes(make_rtf(_w("pard") + _w("f1 ") + sj("氏名：") + _w("ul") + _w("tab") + _w("tab") + _w("tab") + _w("tab")
                             + _w("ulnone ") + sj("印") + _w("par")))
    fl = _fields(src)
    assert list(fl) == ["P1.c3"] and fl["P1.c3"]["facts"]["tabs"] == 4 and not fl["P1.c3"].get("agent")
    res = _fill(src, tmp_path / "o.rtf", {"P1.c3": VALUE})
    assert any("タブ 4 個" in w for w in res["warnings"])
    out = (tmp_path / "o.rtf").read_bytes()
    assert out.count(b"\\tab") == 4
    assert _fields(tmp_path / "o.rtf")["P1.c3"]["value"] == VALUE
    assert _lines(tmp_path / "o.rtf")[0].startswith(f"氏名：［P1.c3＝{VALUE}］")


def test_rtf_list_number_tab_is_part_of_the_number(tmp_path):
    # 段落番号の字（\\listtext）の中のタブは、番号の字のうち：本文の字にしない（下線つきでも、空欄にしない）
    src = tmp_path / "a.rtf"
    src.write_bytes(make_rtf("{" + _w("listtext") + _w("pard") + _w("plain") + _w("f1") + _w("ul ") + sj("○") + _w("tab") + "}"
                             + _w("pard") + _w("f1") + _w("ul ") + sj("世帯に関しての質問") + _w("par")))
    assert not any((f.get("facts") or {}).get("tabs") for f in _fields(src).values())


def test_rtf_leader_tab_stop(tmp_path):
    src = tmp_path / "a.rtf"
    src.write_bytes(make_rtf(
        _w("pard") + _w("tlul") + _w("tqr") + _w("tx8000") + _w("f1 ") + "Name:" + _w("tab") + _w("par"),
        _w("pard") + _w("tx2000") + _w("f1 ") + "Age" + _w("tab") + "30." + _w("par")))
    fl = _fields(src)
    assert list(fl) == ["P1.c5"]  # \pard で、前の段落のタブ位置は消える
    assert fl["P1.c5"]["why_code"] == "leader_tab" and fl["P1.c5"]["facts"]["leaders"] == ["underscore"]
    _fill(src, tmp_path / "o.rtf", {"P1.c5": " Taro"})
    assert (tmp_path / "o.rtf").read_bytes().count(b"\\tab") == 2
    assert _fields(tmp_path / "o.rtf")["P1.c5"]["value"] == " Taro"


# ====================================================================== OpenDocument の文書
ODT_U = '<style:style style:name="T1" style:family="text"><style:text-properties style:text-underline-style="solid"/></style:style>'


def test_odt_underlined_tabs_are_kept(tmp_path):
    src = tmp_path / "a.odt"
    src.write_bytes(make_odt('<text:p>氏名：<text:span text:style-name="T1"><text:tab/><text:tab/><text:tab/></text:span>印</text:p>', ODT_U))
    fl = _fields(src)
    assert list(fl) == ["P1.c3"] and fl["P1.c3"]["facts"]["tabs"] == 3
    res = _fill(src, tmp_path / "o.odt", {"P1.c3": VALUE})
    assert any("タブ 3 個" in w for w in res["warnings"])
    with zipfile.ZipFile(tmp_path / "o.odt") as z:
        xml = z.read("content.xml").decode("utf-8")
    assert xml.count("<text:tab/>") == 3
    # 値は、下線のある所（text:span T1）の中、初めのタブの前
    span = xml[xml.index('<text:span text:style-name="T1">'):]
    assert VALUE in span[:span.index("<text:tab/>")]
    assert _fields(tmp_path / "o.odt")["P1.c3"]["value"] == VALUE


# ====================================================================== HWPX
def test_hwpx_underlined_tabs_are_kept(tmp_path):
    tabs = '<hp:tab width="4000" leader="0" type="1"/>' * 3
    sec = HWPX_XML + f"<hs:sec {HWPX_NS}>" + para(run_(t("성명："), "0"), run_(f"<hp:t>{tabs}</hp:t>", "1"), run_(t("인"), "0")) + "</hs:sec>"
    src = tmp_path / "a.hwpx"
    src.write_bytes(make_hwpx(sec))
    fl = _fields(src)
    spot = next(k for k, f in fl.items() if (f.get("facts") or {}).get("tabs") == 3)
    res = _fill(src, tmp_path / "o.hwpx", {spot: "홍길동"})
    assert any("タブ 3 個" in w for w in res["warnings"])
    with zipfile.ZipFile(tmp_path / "o.hwpx") as z:
        xml = z.read("Contents/section0.xml").decode("utf-8")
    assert xml.count("<hp:tab ") == 3 and xml.index("홍길동") < xml.index("<hp:tab ")
    assert _fields(tmp_path / "o.hwpx")[spot]["value"] == "홍길동"


# ====================================================================== 線（リーダー）を引くタブ位置：.odt・.doc・HWPX
FIX = __import__("pathlib").Path(__file__).parent / "fixtures"


def _leader_places(path) -> dict:
    return {k: f["facts"]["leaders"] for k, f in _fields(path).items() if f.get("why_code") == "leader_tab"}


def _leader_places_in_range(path, rng: str) -> dict:
    """大きな書類の道（段落の範囲で読む道）で見た、同じ所。"""
    return {f["id"]: f["facts"]["leaders"] for f in api.fields_range(str(path), {"paragraphs": rng})["fields"]
            if f.get("why_code") == "leader_tab"}


def _odt_with_styles(body: str, auto: str, styles: str) -> bytes:
    from test_openformats import NS as ODF_NS

    buf = io.BytesIO(make_odt(body, auto))
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED) as z:
        z.writestr("styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><office:document-styles {ODF_NS}>'
                   f'<office:styles>{styles}</office:styles></office:document-styles>')
    return buf.getvalue()


def _odt_stops(*stops: str) -> str:
    return f'<style:paragraph-properties><style:tab-stops>{"".join(stops)}</style:tab-stops></style:paragraph-properties>'


def test_odt_leader_tab_stop_is_an_agent_place(tmp_path):
    auto = ('<style:style style:name="P1" style:family="paragraph">'
            + _odt_stops('<style:tab-stop style:position="14cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>')
            + '</style:style><style:style style:name="P2" style:family="paragraph">'
            + _odt_stops('<style:tab-stop style:position="14cm" style:leader-style="none" style:leader-text="."/>',
                         '<style:tab-stop style:position="3cm"/>') + "</style:style>")
    src = tmp_path / "a.odt"
    src.write_bytes(make_odt('<text:p text:style-name="P1">氏名<text:tab/>印</text:p>'
                             '<text:p text:style-name="P2">住所<text:tab/>まで。</text:p>', auto))
    assert _leader_places(src) == {"P1.c2": ["dotted（.）"]}  # leader-style が none の位置・leader-style のない位置は、線なし
    res = _fill(src, tmp_path / "o.odt", {"P1.c2": VALUE})
    assert any("タブは残しました" in w for w in res["warnings"])
    with zipfile.ZipFile(tmp_path / "o.odt") as z:
        xml = z.read("content.xml").decode("utf-8")
    assert xml.count("<text:tab/>") == 2 and xml.index(VALUE) < xml.index("<text:tab/>")


def test_odt_leader_from_the_parent_chain_and_the_default_style(tmp_path):
    lead = '<style:tab-stop style:position="14cm" style:type="right" style:leader-style="solid" style:leader-text="_"/>'
    styles = ('<style:default-style style:family="paragraph">'
              + _odt_stops('<style:tab-stop style:position="10cm" style:leader-style="dash"/>') + "</style:default-style>"
              '<style:style style:name="Base" style:family="paragraph">' + _odt_stops(lead) + "</style:style>"
              '<style:style style:name="Line" style:family="paragraph" style:parent-style-name="Base"/>'
              '<style:style style:name="Plain" style:family="paragraph" style:parent-style-name="Base">'
              + _odt_stops('<style:tab-stop style:position="3cm"/>') + "</style:style>"
              '<style:style style:name="Empty" style:family="paragraph" style:parent-style-name="Base">'
              "<style:paragraph-properties><style:tab-stops/></style:paragraph-properties></style:style>"
              '<style:style style:name="Free" style:family="paragraph"/>')
    auto = '<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Line"/>'
    body = "".join(f'<text:p text:style-name="{s}">項目<text:tab/>値。</text:p>' for s in ("P1", "Plain", "Empty", "Free"))
    src = tmp_path / "a.odt"
    src.write_bytes(_odt_with_styles(body, auto, styles))
    # style:tab-stops は 1 つの性質：親の連なりで初めにそれを持つスタイルが決める（子が持てば、親の位置は使わない）。
    # どのスタイルも持たなければ、既定のスタイル（style:default-style）
    assert _leader_places(src) == {"P1.c2": ["solid（_）"], "P4.c2": ["dash"]}


def test_odt_and_doc_written_by_libreoffice_read_like_the_docx_they_were_made_from():
    # LibreOffice が同じ .docx から書き出した .odt と .doc（スタイルの連なり・消した位置・段落の位置）
    assert _leader_places(FIX / "leader_tab_styles.odt") == {"P1.c8": ["dotted（.）"], "P4.c3": ["solid（-）"]}
    assert _leader_places(FIX / "leader_tab_styles.doc") == {"P1.c8": ["dot"], "P4.c3": ["hyphen"]}
    # 段落ごとのタブ位置（点・ハイフン・下線。LibreOffice は .doc に太線を下線、中点を線なしとして書き出す）
    assert _leader_places(FIX / "leader_tabs.doc") == {"P1.c4": ["dot"], "P2.c7": ["hyphen"], "P3.c11": ["underscore"],
                                                         "P4.c6": ["underscore"]}
    for name in ("leader_tab_styles.odt", "leader_tab_styles.doc", "leader_tabs.doc"):  # 範囲で読む道でも同じ
        assert _leader_places_in_range(FIX / name, "1-6") == _leader_places(FIX / name)


def test_doc_leader_tab_is_written_before_the_tab(tmp_path):
    res = _fill(FIX / "leader_tab_styles.doc", tmp_path / "o.doc", {"P1.c8": " Intro"})
    assert any("タブは残しました" in w for w in res["warnings"])
    fl = _fields(tmp_path / "o.doc")
    assert fl["P1.c8"]["value"] == " Intro" and _leader_places(tmp_path / "o.doc") == {"P4.c3": ["hyphen"]}


def test_doc_tab_stop_sprms_follow_the_spec():
    import struct

    from todoku.doc.book import lay_tabs, sprm_spans, tab_leaders

    def add(*stops: tuple) -> bytes:  # (位置, jc, tlc)
        return bytes([len(stops)]) + struct.pack(f"<{len(stops)}h", *[s[0] for s in stops]) + bytes(s[1] | s[2] << 3 for s in stops)

    stops: dict = {}
    lay_tabs(stops, b"\x00" + add((2000, 0, 0), (5000, 2, 4), (8000, 2, 5), (9000, 4, 1)), False)
    assert tab_leaders(stops) == ("heavy", "middleDot")  # 縦線（jc 4）の位置の tlc は読まない。tlc 0 は線なし
    assert tab_leaders(stops, 4) == ("heavy",)  # Word 6.0/95 の仕様書の tlc は 4 まで
    lay_tabs(stops, bytes([1]) + struct.pack("<h", 5020) + add((3000, 0, 7)), False)  # 25 twip 以内の位置を消す
    assert sorted(stops) == [2000, 3000, 8000, 9000] and tab_leaders(stops) == ("middleDot",)
    # sprmPChgTabs：消す位置ごとの幅（rgdxaClose）
    lay_tabs(stops, bytes([1]) + struct.pack("<hh", 7900, 100) + add(), True)
    assert sorted(stops) == [2000, 3000, 9000] and tab_leaders(stops) == ()
    lay_tabs(stops, b"\x05\x00", False)  # 途中で終わる値は、重ねない
    assert sorted(stops) == [2000, 3000, 9000]
    # cb が 255 の sprmPChgTabs：長さは、消す数・足す数から数え、後ろの sprm を読み違えない
    op = bytes([1]) + struct.pack("<hh", 100, 30) + add((6000, 2, 1))
    grp = struct.pack("<H", 0xC615) + b"\xff" + op + struct.pack("<HB", 0x2416, 1)
    assert [(sp, size) for sp, _at, size in sprm_spans(grp)] == [(0xC615, len(op)), (0x2416, 1)]


def _word6_stsh(styles: list[tuple]) -> bytes:
    """Word 6.0/95 のスタイルの表（仕様書「Style Sheet (STSH)」の形）。styles：(名前, 基のスタイル, 段落の書式の grpprl)。"""
    import struct

    out = struct.pack("<H", 14) + struct.pack("<7H", len(styles), 8, 1, 0, 15, 0, 0)
    for k, (name, base, grpprl) in enumerate(styles):
        nm = name.encode("ascii")
        std = struct.pack("<4H", k, 1 | base << 4, 2 | 0x0FFF << 4, 0) + bytes([len(nm)]) + nm + b"\x00"
        std += bytes(len(std) & 1)
        upx = struct.pack("<H", k) + grpprl
        std += struct.pack("<H", len(upx)) + upx + bytes(len(upx) & 1) + struct.pack("<H", 0)
        out += struct.pack("<H", len(std)) + std
    return out


def test_word6_leader_tab_stops_from_the_paragraph_and_the_style(tmp_path):
    import struct

    from test_legacy import make_word6

    def chg(sprm: int, dels: list, adds: list, wide: bool = False) -> bytes:
        op = bytes([len(dels)]) + struct.pack(f"<{len(dels)}h", *dels) + (struct.pack(f"<{len(dels)}h", *[25] * len(dels)) if sprm == 23 else b"")
        op += bytes([len(adds)]) + struct.pack(f"<{len(adds)}h", *[a[0] for a in adds]) + bytes(a[1] | a[2] << 3 for a in adds)
        return bytes([sprm, 255 if wide else len(op)]) + op

    text = "Name:\tend\rDate:\tend\rPlain:\tend\rStyle:\tend\rChild:\tend\rCleared:\tend\rWide:\tend\r"
    ends = [i for i, ch in enumerate(text) if ch == "\r"]
    stsh = _word6_stsh([("Normal", 0x0FFF, b""), ("Lead", 0, chg(15, [], [(8000, 2, 1)])), ("Child", 1, b""),
                        ("Off", 1, chg(15, [8010], [(2000, 0, 0)]))])
    paras = {ends[0]: chg(15, [], [(8000, 2, 3)]), ends[1]: chg(23, [], [(8000, 2, 4)]), ends[2]: chg(15, [], [(8000, 2, 0)]),
             ends[3]: (1, b""), ends[4]: (2, b""), ends[5]: (3, b""), ends[6]: chg(23, [], [(8000, 2, 2)], wide=True)}
    src = tmp_path / "a.doc"
    src.write_bytes(make_word6(text, paras, enc="cp1252", lid=0x0409, stsh=stsh))
    assert _leader_places(src) == {"P1.c5": ["underscore"], "P2.c5": ["heavy"], "P4.c6": ["dot"], "P5.c6": ["dot"], "P7.c5": ["hyphen"]}
    assert _leader_places_in_range(src, "1-7") == _leader_places(src)
    # Word 6.0/95 は字のバイト数を変えずに上書きするので、タブの前（字のない所）には書けない：所は、線のあることを伝えるだけ
    with pytest.raises(api.ApiError, match="字のバイト数を変えずに上書き"):
        api.fill(str(src), {"P1.c5": "A"}, str(tmp_path / "o.doc"))


# ---------------------------------------------------------------------- HWPX
def _hwpx_with_tabs(sec: str, tab_prs: str, para_prs: str) -> bytes:
    """header.xml に、タブの定義（hh:tabProperties）と、それを指す段落の形（hh:paraPr）を足した HWPX。"""
    src = zipfile.ZipFile(io.BytesIO(make_hwpx(sec)))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "Contents/header.xml":
                head = data.decode("utf-8")
                assert "<hh:paraProperties" in head and "</hh:paraProperties>" in head
                head = head.replace("<hh:paraProperties", f"<hh:tabProperties>{tab_prs}</hh:tabProperties><hh:paraProperties", 1)
                data = head.replace("</hh:paraProperties>", para_prs + "</hh:paraProperties>", 1).encode("utf-8")
            z.writestr(info, data, compress_type=info.compress_type)
    src.close()
    return buf.getvalue()


def test_hwpx_leader_tab_stop_is_an_agent_place(tmp_path):
    item = '<hh:tabItem pos="{}" type="RIGHT" leader="{}"/>'
    tab_prs = ('<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>'
               f'<hh:tabPr id="1" autoTabLeft="0" autoTabRight="0">{item.format(40000, "DOT")}{item.format(9000, "NONE")}</hh:tabPr>'
               '<hh:tabPr id="2" autoTabLeft="0" autoTabRight="0"><hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">'
               f'{item.format(80000, "DASH")}</hp:case><hp:default>{item.format(40000, "DASH")}</hp:default></hp:switch></hh:tabPr>'
               f'<hh:tabPr id="3" autoTabLeft="0" autoTabRight="1">{item.format(9000, "NONE")}</hh:tabPr>')
    para_prs = "".join(f'<hh:paraPr id="{10 + k}" tabPrIDRef="{k}"><hh:heading type="NONE" idRef="0" level="0"/></hh:paraPr>' for k in range(4))
    tab_ = '<hp:tab width="4000" leader="{}" type="2"/>'
    line = lambda label, pr, lead="0": para(run_(f"<hp:t>{label}{tab_.format(lead)}끝.</hp:t>"), pr=pr)  # noqa: E731
    sec = (HWPX_XML + f"<hs:sec {HWPX_NS}>" + line("성명", "11") + line("주소", "12") + line("날짜", "13") + line("기타", "10")
           + line("전화", "10", "3") + "</hs:sec>")
    src = tmp_path / "a.hwpx"
    src.write_bytes(_hwpx_with_tabs(sec, tab_prs, para_prs))
    # 段落の形 → タブの定義 → hh:tabItem の leader（NONE のほか）。タブそのものの leader（0 のほか）は、書いてあるとおりに示す
    assert _leader_places(src) == {"P1.c2": ["DOT"], "P2.c2": ["DASH"], "P5.c2": ['hp:tab leader="3"']}
    assert _leader_places_in_range(src, "1-5") == _leader_places(src)
    res = _fill(src, tmp_path / "o.hwpx", {"P1.c2": "홍길동"})
    assert any("タブは残しました" in w for w in res["warnings"])
    with zipfile.ZipFile(tmp_path / "o.hwpx") as z:
        xml = z.read("Contents/section0.xml").decode("utf-8")
    assert xml.count("<hp:tab ") == 5 and xml.index("홍길동") < xml.index("<hp:tab ")
