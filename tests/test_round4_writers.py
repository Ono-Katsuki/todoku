"""第4回の見直しで見つかった、書き込みの誤り（証明と読み戻しを通ってしまう誤った書き込み）の確かめ（合成の文書）。

- Word：XML のデータに結び付いたコンテンツコントロール（w:dataBinding）は、結び付けた customXml の値も書く。たどれなければ断る
- Word：1行の書式なしテキストのコントロールに改行を書かない。複数行なら改行の要素（w:br）で書く
- Word：警告の欄の名前は、欄の id（T1.R2.C2）で伝える
- ODT・RTF・PowerPoint・HWPX：書く所に残る取り消し線・字の色・非表示などを、書いた字に引き継がない
- PowerPoint：空のマス（<a:p/>）への書き込みを、証明が誤って止めない
- .txt：str.splitlines が行を分ける字（U+2028 など）は書かない
- CSV：表計算が数式として読む先頭の字（= + - @ など）を、値を変えずに警告で伝える
"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku import api
from todoku.docx.document import DocxDocument
from todoku.docx.fill import FillError, fill
from todoku.docx.proof import prove

import test_word_findings as wf

STORE = "{11111111-2222-3333-4444-555555555555}"


def bound_docx(body: str, item: str, *, store: str = STORE, props: bool = True) -> bytes:
    """wf.docx に、ユーザー定義の XML の部品（customXml/item1.xml と性質の部品）を足したもの。"""
    base = wf.docx(body)
    src = zipfile.ZipFile(io.BytesIO(base))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            data = src.read(n)
            if n == "word/_rels/document.xml.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rC" Type="http://schemas.openxmlformats.org/'
                                    b'officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>')
            z.writestr(n, data)
        z.writestr("customXml/item1.xml", item)
        if props:
            z.writestr("customXml/_rels/item1.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/'
                       'package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                       'relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')
            z.writestr("customXml/itemProps1.xml", f'<?xml version="1.0"?><ds:datastoreItem ds:itemID="{store}" '
                       'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>')
    return buf.getvalue()


def sdt(alias: str, pr: str, content: str = "ここをクリック", plc: bool = True) -> str:
    return (f'<w:sdt><w:sdtPr><w:alias w:val="{alias}"/>{"<w:showingPlcHdr/>" if plc else ""}{pr}</w:sdtPr>'
            f"<w:sdtContent><w:r><w:t>{content}</w:t></w:r></w:sdtContent></w:sdt>")


def bind(xpath: str, pm: str = "") -> str:
    pmx = f' w:prefixMappings="{pm}"' if pm else ""
    return f'<w:dataBinding{pmx} w:xpath="{xpath}" w:storeItemID="{STORE}"/>'


def _custom(blob: bytes) -> bytes:
    return zipfile.ZipFile(io.BytesIO(blob)).read("customXml/item1.xml")


def _fields(path: str) -> dict:
    return {f["id"]: f for f in api.fields(path)}


# ---------------------------------------------------------------------- A2. w:dataBinding
NSX = "urn:example:form"
ITEM = f'<?xml version="1.0" encoding="UTF-8"?><ns0:root xmlns:ns0="{NSX}"><ns0:tel>000-0000</ns0:tel><ns0:mail/></ns0:root>'
BODY = (wf.p(wf.r("電話 "), sdt("電話", bind("/ns0:root[1]/ns0:tel[1]", f"xmlns:ns0='{NSX}'") + "<w:text/>"))
        + wf.p(wf.r("メール "), sdt("メール", bind("/ns0:root[1]/ns0:mail[1]", f"xmlns:ns0='{NSX}'") + "<w:text/>")))


def test_bound_control_window_shows_the_bound_value(tmp_path):
    path = wf.write(tmp_path, bound_docx(BODY, ITEM))
    f = _fields(path)
    assert f["電話"]["value"] == "000-0000"  # Word が開くときに表示する値（customXml の値）
    win = {x["id"]: x for x in wf.build(DocxDocument(bound_docx(BODY, ITEM)))["fields"]}
    assert win["電話"]["bound_to"].startswith("customXml/item1.xml")
    assert f["メール"]["value"] is None  # 空の値でプレースホルダー


def test_bound_control_fill_writes_custom_xml_and_proves(tmp_path):
    blob = bound_docx(BODY, ITEM)
    path = wf.write(tmp_path, blob)
    r = api.fill(path, {"電話": "090-1111-2222", "メール": "a&b@example.jp"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    out = (tmp_path / "o.docx").read_bytes()
    want = ITEM.replace("000-0000", "090-1111-2222").replace("<ns0:mail/>", "<ns0:mail>a&amp;b@example.jp</ns0:mail>")
    assert _custom(out) == want.encode()  # 値の所のほかは、バイトのまま
    assert _fields(str(tmp_path / "o.docx"))["電話"]["value"] == "090-1111-2222"


def test_bound_control_proof_catches_custom_xml_tampering():
    blob = bound_docx(BODY, ITEM)
    res = fill(DocxDocument(blob), {"電話": "090"})
    from todoku.docx.fill import written_of
    from todoku.zipraw import RawZip

    assert prove(blob, res.blob, written_of(res.patches)).ok
    z = RawZip(res.blob)
    z.replace("customXml/item1.xml", _custom(res.blob).replace(b"<ns0:mail/>", b"<ns0:mail>x</ns0:mail>"))
    bad = prove(blob, z.build(), written_of(res.patches))
    assert not bad.ok and any("customXml" in u for u in bad.unexpected)
    z = RawZip(res.blob)
    z.replace("customXml/item1.xml", _custom(res.blob).replace(b"090", b"091"))
    bad = prove(blob, z.build(), written_of(res.patches))
    assert not bad.ok  # 結び付けた値が、書いた値と違う


@pytest.mark.parametrize("xpath,why", [
    ("/ns0:root[1]/ns0:fax[1]", "ありません"),
    ("/ns0:root[1]/ns9:tel[1]", "接頭辞"),
    ("//ns0:tel", "根からたどる"),
    ("/ns0:root[1]/ns0:tel[position()=1]", "確かめられる形"),
])
def test_bound_control_unresolvable_is_refused(xpath, why):
    body = wf.p(wf.r("電話 "), sdt("電話", bind(xpath, f"xmlns:ns0='{NSX}'") + "<w:text/>"))
    with pytest.raises(FillError) as e:
        fill(DocxDocument(bound_docx(body, ITEM)), {"電話": "090"})
    assert "w:dataBinding" in str(e.value) and why in str(e.value)


def test_bound_control_missing_store_is_refused():
    body = wf.p(wf.r("電話 "), sdt("電話", bind("/ns0:root[1]/ns0:tel[1]", f"xmlns:ns0='{NSX}'") + "<w:text/>"))
    with pytest.raises(FillError, match="storeItemID"):
        fill(DocxDocument(bound_docx(body, ITEM, store="{99999999-2222-3333-4444-555555555555}")), {"電話": "090"})


def test_bound_controls_sharing_a_node(tmp_path):
    pr = bind("/ns0:root[1]/ns0:tel[1]", f"xmlns:ns0='{NSX}'") + "<w:text/>"
    body = wf.p(sdt("電話", pr)) + wf.p(sdt("電話（控え）", pr))
    blob = bound_docx(body, ITEM)
    with pytest.raises(FillError, match="違う値"):
        fill(DocxDocument(blob), {"電話": "1", "電話（控え）": "2"})
    res = fill(DocxDocument(blob), {"電話": "1"})
    assert any("電話（控え）" in w and "同じ値" in w for w in res.warnings)


def test_bound_checkbox_and_dropdown_and_date(tmp_path):
    item = (f'<?xml version="1.0"?><r xmlns="{NSX}"><ok>false</ok><pref>13</pref><day></day></r>')
    pm = f"xmlns:ns0='{NSX}'"
    body = (wf.p(sdt("同意", bind("/ns0:r[1]/ns0:ok[1]", pm) + '<w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>', "☐", plc=False))
            + wf.p(sdt("県", bind("/ns0:r[1]/ns0:pref[1]", pm) + '<w:dropDownList><w:listItem w:displayText="東京都" w:value="13"/>'
                       '<w:listItem w:displayText="大阪府" w:value="27"/></w:dropDownList>', "東京都", plc=False))
            + wf.p(sdt("日", bind("/ns0:r[1]/ns0:day[1]", pm) + '<w:date><w:dateFormat w:val="yyyy/MM/dd"/><w:lid w:val="ja-JP"/>'
                       '<w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>')))
    path = wf.write(tmp_path, bound_docx(body, item))
    f = _fields(path)
    assert f["同意"]["value"] is False and f["県"]["value"] == "東京都"
    r = api.fill(path, {"同意": True, "県": "大阪府", "日": "2026-11-04"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    got = _custom((tmp_path / "o.docx").read_bytes()).decode()
    assert "<ok>true</ok>" in got and "<pref>27</pref>" in got and "<day>2026-11-04T00:00:00</day>" in got


def test_bound_attribute(tmp_path):
    item = f'<?xml version="1.0"?><a:r xmlns:a="{NSX}"><a:x a:v="old" w="1"/></a:r>'
    body = wf.p(sdt("属性", bind("/ns0:r[1]/ns0:x[1]/@ns0:v", f"xmlns:ns0='{NSX}'") + "<w:text/>"))
    path = wf.write(tmp_path, bound_docx(body, item))
    assert _fields(path)["属性"]["value"] == "old"
    r = api.fill(path, {"属性": 'a"b'}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert _custom((tmp_path / "o.docx").read_bytes()) == item.replace('a:v="old"', 'a:v="a&quot;b"').encode()


# ---------------------------------------------------------------------- A11. 1行のテキストのコントロール
def test_single_line_text_control_refuses_newline():
    body = wf.p(sdt("メール", '<w:text w:multiLine="0"/>')) + wf.p(sdt("住所", "<w:text/>"))
    for key in ("メール", "住所"):  # w:multiLine の既定は「いいえ」
        for nl in ("\n", " ", "\r"):
            with pytest.raises(FillError, match="1行の欄"):
                fill(DocxDocument(wf.docx(body)), {key: f"a{nl}b"})


def test_multi_line_text_control_uses_br(tmp_path):
    body = wf.p(sdt("備考", '<w:text w:multiLine="1"/>')) + wf.p(sdt("自由", ""))
    path = wf.write(tmp_path, wf.docx(body))
    r = api.fill(path, {"備考": "1行目\n2行目", "自由": "x\ny"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    xml = zipfile.ZipFile(tmp_path / "o.docx").read("word/document.xml").decode()
    assert "1行目</w:t><w:br/><w:t" in xml and "1行目\n" not in xml and "x</w:t><w:br/>" in xml


# ---------------------------------------------------------------------- C6. 警告の欄の名前
def test_dropped_formatting_warning_uses_field_id(tmp_path):
    body = ('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>旧住所</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:strike/></w:rPr>'
            '<w:t xml:space="preserve">　　　　</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
    res = fill(DocxDocument(wf.docx(body)), {"T1.R1.C2": "東京"})
    assert any(w.startswith("T1.R1.C2:") for w in res.warnings), res.warnings
    assert not any("T1_R1_C2" in w for w in res.warnings)


# ---------------------------------------------------------------------- A3. ODT
from test_openformats import make_odt  # noqa: E402

ODT_AUTO = ('<style:style style:name="T1" style:family="text"><style:text-properties fo:color="#ff0000" '
            'style:text-line-through-style="solid" fo:font-size="14pt"/></style:style>'
            '<style:style style:name="T2" style:family="text"><style:text-properties fo:color="#ffffff"/></style:style>'
            '<style:style style:name="T3" style:family="text"><style:text-properties text:display="none"/></style:style>'
            '<style:style style:name="T4" style:family="text"><style:text-properties fo:background-color="#ffff00"/></style:style>')


def _odt_fill(tmp_path, body: str, values: dict):
    path = tmp_path / "a.odt"
    path.write_bytes(make_odt(body, ODT_AUTO))
    r = api.fill(str(path), values, str(tmp_path / "o.odt"))
    return r, zipfile.ZipFile(tmp_path / "o.odt").read("content.xml").decode()


def test_odt_written_text_does_not_inherit_strike_or_colour(tmp_path):
    body = ('<text:p>住所：<text:span text:style-name="T1">　　　　</text:span></text:p>'
            '<text:p>電話：<text:span text:style-name="T2">＿＿＿＿</text:span></text:p>'
            '<text:p>番号：<text:span text:style-name="T4">　　　　</text:span></text:p>')
    r, xml = _odt_fill(tmp_path, body, {"P1.1": "大阪", "P2.1": "03", "P3.1": "12"})
    assert r["proof_ok"] and r["read_back_ok"], r
    assert not r.get("written_formatting"), r.get("written_formatting")
    assert any("P1.c3" in w and "取り消し線" in w and "字の色" in w for w in r["warnings"])
    assert 'style:text-line-through-style="none"' in xml and 'style:use-window-font-color="true"' in xml
    assert 'fo:background-color="transparent"' in xml
    # 字の大きさ（14pt の span）は、外の span のまま
    # （空白だけの長い空きは、見出しにくっつかないよう、初めの 1 字を残して中に書く）
    assert '<text:span text:style-name="T1">　<text:bookmark-start text:name="_todoku_P1.c3"/><text:span text:style-name="todoku_w' in xml


def test_odt_proof_catches_other_style_change(tmp_path):
    from todoku.odf.text import OdtDocument
    from todoku.odf.textfill import fill as ofill
    from todoku.odf.textfill import written_of
    from todoku.odf.textproof import prove as oprove
    from todoku.odf.package import OdfPackage

    blob = make_odt('<text:p>住所：<text:span text:style-name="T1">　　　　</text:span></text:p>', ODT_AUTO)
    res = ofill(OdtDocument(blob), {"P1.1": "大阪"})
    assert oprove(blob, res.blob, written_of(res.patches)).ok
    pkg = OdfPackage(res.blob)
    bad = pkg.xml("content.xml").replace(b'style:text-line-through-type="none"', b'style:text-line-through-type="none" fo:font-size="30pt"')
    bad_blob = pkg.rebuilt({"content.xml": bad})
    assert not oprove(blob, bad_blob, written_of(res.patches)).ok


def test_odt_plain_span_unchanged(tmp_path):
    r, xml = _odt_fill(tmp_path, "<text:p>氏名：　　　　</text:p>", {"P1.1": "山田"})
    assert r["proof_ok"] and r["read_back_ok"] and "todoku_w" not in xml and not r["warnings"]


# ---------------------------------------------------------------------- A3. RTF
B = "\\"


def _rtf(body: str, colors: str = ";\\red0\\green0\\blue0;\\red255\\green0\\blue0;\\red255\\green255\\blue255;") -> bytes:
    return ("{\\rtf1\\ansi\\ansicpg932\\deff0{\\fonttbl{\\f0\\fcharset0 Century;}}{\\colortbl" + colors + "}\r\n"
            + body + "\r\n}").encode("latin-1")


def _u(s: str) -> str:
    return "".join(f"\\u{ord(c)}?" for c in s)


@pytest.mark.parametrize("fmt,name", [("\\strike", "取り消し線"), ("\\cf2", "字の色"), ("\\cf3", "字の色"), ("\\striked1", "二重"),
                                      ("\\highlight2", "蛍光ペン")])
def test_rtf_written_text_does_not_inherit(tmp_path, fmt, name):
    blob = _rtf("\\pard " + _u("住所：") + "{" + fmt + " " + _u("　　　　") + "}\\par")
    path = tmp_path / "a.rtf"
    path.write_bytes(blob)
    r = api.fill(str(path), {"P1.1": "大阪"}, str(tmp_path / "o.rtf"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert not r.get("written_formatting"), r.get("written_formatting")
    assert any("P1.c3" in w and name in w for w in r["warnings"]), r["warnings"]
    from todoku.rtf.fill import written_formatting

    assert written_formatting((tmp_path / "o.rtf").read_bytes()) == []


def test_rtf_written_formatting_reads_back_strike():
    from todoku.rtf.fill import written_formatting

    blob = _rtf("\\pard " + _u("住所：") + "{\\*\\bkmkstart _todoku_P1_1}{\\strike " + _u("大阪") + "}{\\*\\bkmkend _todoku_P1_1}\\par")
    got = written_formatting(blob)
    assert got and "取り消し線" in got[0]["formatting"]


# ---------------------------------------------------------------------- A3 + C4. PowerPoint
import pptx_maker as pm  # noqa: E402


def _pptx_fill(tmp_path, shapes: str, values: dict):
    path = tmp_path / "a.pptx"
    path.write_bytes(pm.make([shapes]))
    r = api.fill(str(path), values, str(tmp_path / "o.pptx"))
    return r, zipfile.ZipFile(tmp_path / "o.pptx").read("ppt/slides/slide1.xml").decode()


def test_pptx_written_text_drops_strike_and_white(tmp_path):
    para = ('<a:p><a:r><a:rPr lang="ja-JP"/><a:t>氏名：</a:t></a:r><a:r><a:rPr lang="ja-JP" sz="1400" strike="sngStrike">'
            '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Meiryo"/></a:rPr><a:t>　　　　　　</a:t></a:r></a:p>')
    r, xml = _pptx_fill(tmp_path, pm.sp(para), {"P1.1": "山田"})
    assert r["proof_ok"] and r["read_back_ok"], r
    assert not r.get("written_formatting")
    assert any("P1.c3" in w and "取り消し線" in w for w in r["warnings"])
    assert '<a:rPr lang="ja-JP" sz="1400" strike="noStrike"><a:latin typeface="Meiryo"/></a:rPr><a:t>山田</a:t>' in xml


def test_pptx_empty_cell_paragraph_proof(tmp_path):
    tbl = pm.table([["住所", ""], ["氏名", ""]])
    tbl = tbl.replace('<a:p><a:endParaRPr lang="ja-JP" sz="1400"/></a:p>', "<a:p/>", 1)
    r, xml = _pptx_fill(tmp_path, tbl, {"T1.R1.C2": "東京", "T1.R2.C2": "山田"})
    assert r["proof_ok"] and r["read_back_ok"], r


def test_pptx_written_formatting_detects_bad_run():
    from todoku.pptx.fill import written_formatting

    a = pm.make([pm.sp('<a:p><a:r><a:rPr lang="ja-JP"/><a:t>氏名：　　　</a:t></a:r></a:p>')])
    b = a.replace(b"x", b"x")
    b = pm.make([pm.sp('<a:p><a:r><a:rPr lang="ja-JP"/><a:t>氏名：</a:t></a:r><a:r><a:rPr lang="ja-JP" strike="dblStrike"/>'
                       '<a:t>山田</a:t></a:r></a:p>')])
    got = written_formatting(a, b, {"P1.c3": "山田"})
    assert got and got[0]["id"] == "P1.c3" and "取り消し線" in got[0]["formatting"]


# ---------------------------------------------------------------------- A3. HWPX
import test_hwpx as th  # noqa: E402


def _hwpx_with_struck_run() -> bytes:
    hdr = th.HEADER.replace('<hh:charProperties itemCnt="2">', '<hh:charProperties itemCnt="3">').replace(
        "</hh:charProperties>",
        '<hh:charPr id="2" height="1200" textColor="#FF0000" shadeColor="#FFFF00" borderFillIDRef="1"><hh:underline type="NONE" '
        'shape="SOLID" color="#000000"/><hh:strikeout shape="SOLID" color="#000000"/></hh:charPr></hh:charProperties>')
    sec = th.section().replace(th.para(th.run_(th.t("생년월일 ")) + th.run_(th.t("            "), "1")),
                               th.para(th.run_(th.t("생년월일：")) + th.run_(th.t("　　　　　　"), "2")))
    orig = th.HEADER
    th.HEADER = hdr
    try:
        return th.make_hwpx(sec)
    finally:
        th.HEADER = orig


def test_hwpx_written_text_drops_strike_colour_shade(tmp_path):
    from todoku.hwpx.document import HwpxDocument
    from todoku.hwpx.fill import fill as hfill
    from todoku.hwpx.fill import written_formatting, written_of
    from todoku.hwpx.proof import prove as hprove

    blob = _hwpx_with_struck_run()
    path = tmp_path / "a.hwpx"
    path.write_bytes(blob)
    key = next(f["id"] for f in api.fields(str(path)) if "생년월일" in (f.get("label") or ""))
    r = api.fill(str(path), {key: "1990"}, str(tmp_path / "o.hwpx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert not r.get("written_formatting")
    assert any(key in w and "取り消し線" in w and "字の色" in w and "網かけ" in w for w in r["warnings"])
    out = (tmp_path / "o.hwpx").read_bytes()
    hdr = zipfile.ZipFile(io.BytesIO(out)).read("Contents/header.xml").decode()
    assert '<hh:charProperties itemCnt="4">' in hdr
    assert ('<hh:charPr id="3" height="1200" textColor="#000000" shadeColor="none" borderFillIDRef="1">'
            '<hh:underline type="NONE" shape="SOLID" color="#000000"/><hh:strikeout shape="NONE" color="#000000"/></hh:charPr>') in hdr
    assert written_formatting(out) == []
    # 足した文字の形を書き換えると、証明が止める
    res = hfill(HwpxDocument(blob), {key: "1990"})
    from todoku.zipraw import RawZip

    z = RawZip(res.blob)
    h = z.read("Contents/header.xml").replace(b'id="3" height="1200"', b'id="3" height="2400"')
    z.replace("Contents/header.xml", h)
    assert not hprove(blob, z.build(), written_of(res.patches)).ok


# ---------------------------------------------------------------------- D3. .txt
@pytest.mark.parametrize("sep", [" ", " ", "\x0b", "\x0c", "\x1c", "\x1d", "\x1e", "\x85", "\r", "\n"])
def test_txt_refuses_all_line_separators(tmp_path, sep):
    path = tmp_path / "a.txt"
    path.write_bytes("氏名：　　　　\n住所：　　　　\n".encode("utf-8"))
    key = next(f["id"] for f in api.fields(str(path)) if "氏名" in (f.get("label") or ""))
    with pytest.raises(api.ApiError, match="改行"):
        api.fill(str(path), {key: f"山{sep}田"}, str(tmp_path / "o.txt"))


# ---------------------------------------------------------------------- A10. CSV
def test_csv_formula_like_values_are_warned_not_changed(tmp_path):
    path = tmp_path / "a.csv"
    path.write_bytes("氏名,電話,金額,メモ,通貨\r\n,,,,\r\n".encode("utf-8"))
    vals = {"R2.C1": "=HYPERLINK(\"x\")", "R2.C2": "+81-90", "R2.C3": "-5", "R2.C4": "@SUM(A1)", "R2.C5": "\tx"}
    r = api.fill(str(path), vals, str(tmp_path / "o.csv"))
    assert r["proof_ok"] and r["read_back_ok"], r
    risky = {x["id"]: x for x in r["formula_like"]}
    assert set(risky) == {"R2.C1", "R2.C2", "R2.C4", "R2.C5"}  # ただの負の数は数式として読まれない
    assert risky["R2.C2"]["safe_form"] == "'+81-90"
    assert "CSV インジェクション" in r["say_to_user"] and "'=HYPERLINK" in r["say_to_user"]
    out = (tmp_path / "o.csv").read_bytes().decode("utf-8")
    assert '=HYPERLINK(""x"")' in out and "+81-90" in out and "'" not in out  # 値は変えない
