"""Word のほかの形式でも、書いたことで行が折り返す・セルに収まらない・ページが増える見込みを伝える。

前は、行が折り返す見込み（wrapcheck）とページの増え（pagecount）は Word（.docx・.doc）と .odt の一部だけ、セルに収まるかの
見積もり（cellcheck）は .xlsx・.xls だけだった。いまは RTF・HWPX・PowerPoint・.xlsb・.ods にも、同じ決め方で伝える。
どの形式も、書いた段落は「書く前と書いた後で字が変わった段落」（頼み方ではなく、文書に起きたこと）で決める。
"""

from __future__ import annotations

import io
import struct
import zipfile

import pytest

import pptx_maker as pm
import xlsb_maker as xm
from test_hwpx import make_hwpx
from test_hwpx import para as hpara
from test_hwpx import run_ as hrun
from test_hwpx import t as ht
from test_hwpx import NS as HNS
from test_hwpx import XML as HXML
from test_openformats import NS as ONS
from todoku import api

MSG = "書いたことで行が折り返す見込み"
PAGES = "ページが増えるかもしれません"
NAME = "架空太郎"


def _fill(tmp_path, name: str, blob: bytes, values: dict) -> dict:
    src = tmp_path / name
    src.write_bytes(blob)
    r = api.fill(str(src), values, str(tmp_path / ("o_" + name)), overwrite=True)
    assert r["verified"], (r.get("unexpected"), r.get("read_back_differences"))
    return r


def _wrap(r: dict) -> list[str]:
    return [w for w in r["warnings"] if MSG in w]


# ====================================================================== 見積もりのずれは、広い側にも狭い側にもある
def test_wrap_estimate_says_which_side_it_is_unsure_about():
    from todoku.wrapcheck import say, wraps

    assert wraps("あ" * 30, "あ" * 45, 400, 10, 10) == (1, 2, True)  # 450 pt は 400 pt に入らない
    assert wraps("あ" * 30, "あ" * 38, 400, 10, 10) == (1, 2, False)  # 380 pt：字の幅を少し広く測ると入らない
    # 書く前の 410 pt は、見積もりでは 2 行。字の幅を少し狭く測ると 1 行で、書いた後の 600 pt はどちらでも 2 行
    assert wraps("あ" * 41, "架" * 60, 400, 10, 10) == (1, 2, None)
    assert wraps("あ" * 41, "あ" * 44, 400, 10, 10) is None  # どちらの測り方でも、行は増えない
    assert wraps("あ" * 30, "あ" * 31, 400, 10, 10) is None
    wide, narrow = say("P1", "架空", 1, 2, "x", False), say("P1", "架空", 1, 2, "x", None)
    assert "字の幅を少し広く測るソフト" in wide and "書く前の字が幅いっぱいに近く、字の幅を少し狭く測るソフト" in narrow
    assert MSG in wide and MSG in narrow and narrow.added == 1


def test_rtf_paragraph_that_was_already_at_the_edge_is_still_told(tmp_path):
    # 書く前：半角の空白 61 個と全角 6 字（12 pt で 438 pt）は、見積もりでは 425 pt を少し出て 2 行。書いた後（648 pt）も 2 行
    old = " " * 61
    tail = "申請の目的に"
    r = _fill(tmp_path, "a.rtf", _rtf("\\pard\\fs24 " + old + _u(tail) + "\\par"), {"@1": {"old": old, "new": NAME * 12}})
    hit = _wrap(r)
    assert hit and "書く前の字が幅いっぱいに近く" in hit[0] and "1 行から 2 行" in hit[0], r["warnings"]


# ====================================================================== RTF
def _u(s: str) -> str:
    return "".join(c if ord(c) < 128 else f"\\u{ord(c)}?" for c in s)


def _rtf(body: str, head: str = "\\paperw11906\\margl1701\\margr1701") -> bytes:
    return ("{\\rtf1\\ansi\\deff0\\uc1{\\fonttbl{\\f0 Arial;}}" + head + body + "}").encode("ascii")


LABEL = "氏名：　　　　"
LONG = "氏名：" + NAME * 10  # 43 字：12 pt で 516 pt、8 pt で 344 pt


def test_rtf_paragraph_wraps_at_the_page_width(tmp_path):
    # A4（11906 twip）から左右の余白 1701 twip ずつ＝425 pt。12 pt の 43 字は 2 行
    r = _fill(tmp_path, "a.rtf", _rtf("\\pard\\fs24 " + _u(LABEL) + "\\par"), {"@1": {"old": LABEL, "new": LONG}})
    hit = _wrap(r)
    assert hit and "1 行から 2 行" in hit[0] and "ページの幅（余白を引いた幅） 約 425 ポイント・字の大きさ 12 ポイント" in hit[0], r["warnings"]
    assert r["wrap"] and MSG in r["say_to_user"]


def test_rtf_short_value_and_small_font_do_not_wrap(tmp_path):
    r = _fill(tmp_path, "a.rtf", _rtf("\\pard\\fs24 " + _u(LABEL) + "\\par"), {"@1": {"old": LABEL, "new": "氏名：" + NAME}})
    assert not _wrap(r) and "wrap" not in r and "page_growth" not in r
    # \fs16（8 pt）なら 344 pt で、425 pt に入る
    r = _fill(tmp_path, "b.rtf", _rtf("\\pard\\fs16 " + _u(LABEL) + "\\par"), {"@1": {"old": LABEL, "new": LONG}})
    assert not _wrap(r), r["warnings"]


def test_rtf_paragraph_indents_narrow_the_line(tmp_path):
    # \li2000・\ri2000（100 pt ずつ）：425 − 200 = 225 pt
    r = _fill(tmp_path, "a.rtf", _rtf("\\pard\\li2000\\ri2000\\fs16 " + _u(LABEL) + "\\par"), {"@1": {"old": LABEL, "new": LONG}})
    hit = _wrap(r)
    assert hit and "約 225 ポイント・字の大きさ 8 ポイント" in hit[0], r["warnings"]


def test_rtf_columns_split_the_page_width(tmp_path):
    # 2 段・段の間 720 twip：（8504 − 720）÷ 2 = 3892 twip ≒ 195 pt
    r = _fill(tmp_path, "a.rtf", _rtf("\\sectd\\cols2\\colsx720\\pard\\fs16 " + _u(LABEL) + "\\par"), {"@1": {"old": LABEL, "new": LONG}})
    hit = _wrap(r)
    assert hit and "段組み（2 段）の 1 段の幅 約 195 ポイント" in hit[0], r["warnings"]


RTF_ROW = ("\\trowd\\trgaph100\\cellx2000\\cellx5000\\pard\\intbl\\fs24 " + _u("氏名") + "\\cell\\pard\\intbl\\fs24 \\cell\\row\\pard\\par")


def test_rtf_table_cell_uses_the_cell_width(tmp_path):
    # マス 2000〜5000 twip（150 pt）から、左右の余白 \trgaph 100 twip（5 pt）ずつ＝140 pt。12 pt の 12 字（144 pt）は 2 行
    src = tmp_path / "a.rtf"
    src.write_bytes(_rtf(RTF_ROW))
    cid = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("T1"))
    r = api.fill(str(src), {cid: NAME * 3}, str(tmp_path / "o.rtf"))
    hit = _wrap(r)
    assert r["verified"] and hit and hit[0].startswith(cid) and "表のマスの幅 約 140 ポイント" in hit[0], r["warnings"]
    r = api.fill(str(src), {cid: NAME * 2}, str(tmp_path / "o2.rtf"))
    assert not _wrap(r), r["warnings"]


def test_rtf_page_growth_states_the_recorded_pages(tmp_path):
    body = "\\pard\\fs24 " + _u(LABEL) + "\\par"
    r = _fill(tmp_path, "a.rtf", _rtf(body, "\\paperw11906\\margl1701\\margr1701{\\info{\\nofpages2}}"), {"@1": {"old": LABEL, "new": LONG}})
    assert r["page_growth"] == {"paragraphs": 1, "lines_added": 1, "recorded_pages_before": 2}
    assert PAGES in r["say_to_user"] and "ページ数の記録は 2 ページです" in r["say_to_user"]
    r = _fill(tmp_path, "b.rtf", _rtf(body), {"@1": {"old": LABEL, "new": LONG}})
    assert r["page_growth"]["recorded_pages_before"] is None and "ページ数の記録はありません" in r["say_to_user"]


# ====================================================================== OpenDocument の文書（.odt）
def _odf(mime: str, body: str, auto: str = "", meta: str = "", styles: str = "") -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), mime)
        parts = "".join(f'<manifest:file-entry manifest:full-path="{n}" manifest:media-type="text/xml"/>'
                        for n, x in (("content.xml", "1"), ("meta.xml", meta), ("styles.xml", styles)) if x)
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest '
                   'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
                   f'<manifest:file-entry manifest:full-path="/" manifest:media-type="{mime}"/>{parts}</manifest:manifest>',
                   zipfile.ZIP_DEFLATED)
        z.writestr("content.xml", f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {ONS}><office:automatic-styles>{auto}'
                   f'</office:automatic-styles><office:body>{body}</office:body></office:document-content>', zipfile.ZIP_DEFLATED)
        if meta:
            z.writestr("meta.xml", '<?xml version="1.0" encoding="UTF-8"?><office:document-meta '
                       'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
                       'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3">'
                       f"<office:meta>{meta}</office:meta></office:document-meta>", zipfile.ZIP_DEFLATED)
    return buf.getvalue()


ODT_AUTO = ('<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>'
            '<style:style style:name="P1" style:family="paragraph"><style:text-properties fo:font-size="12pt"/></style:style>')
ODT_TABLE = ('<table:table table:name="t1"><table:table-column table:style-name="co1" table:number-columns-repeated="2"/>'
             '<table:table-row><table:table-cell><text:p text:style-name="P1">氏名</text:p></table:table-cell>'
             '<table:table-cell><text:p text:style-name="P1"/></table:table-cell></table:table-row></table:table>')


def _odt(body: str = ODT_TABLE, meta: str = "") -> bytes:
    return _odf("application/vnd.oasis.opendocument.text", f"<office:text>{body}</office:text>", ODT_AUTO, meta)


def test_odt_table_cell_wraps_and_states_recorded_pages(tmp_path):
    # 列 3 cm（85 pt）から左右の余白（仮に 0.097 cm ずつ）を引いた約 80 pt。12 pt の 8 字（96 pt）は 2 行
    src = tmp_path / "a.odt"
    src.write_bytes(_odt(meta='<meta:document-statistic meta:page-count="3" meta:paragraph-count="2"/>'))
    cid = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("T1"))
    r = api.fill(str(src), {cid: NAME * 2}, str(tmp_path / "o.odt"))
    hit = _wrap(r)
    assert r["verified"] and hit and hit[0].startswith(cid) and "1 行から 2 行" in hit[0], r["warnings"]
    assert r["page_growth"]["recorded_pages_before"] == 3 and "ページ数の記録は 3 ページです" in r["say_to_user"]
    r = api.fill(str(src), {cid: NAME}, str(tmp_path / "o2.odt"))
    assert not _wrap(r) and "page_growth" not in r


def test_odt_without_a_page_record_says_so(tmp_path):
    src = tmp_path / "a.odt"
    src.write_bytes(_odt())
    cid = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("T1"))
    r = api.fill(str(src), {cid: NAME * 2}, str(tmp_path / "o.odt"))
    assert r["page_growth"]["recorded_pages_before"] is None and "ページ数の記録はありません" in r["say_to_user"]


def test_odt_text_control_in_a_cell_is_named_by_its_id(tmp_path):
    # 入力の仕組み（text:placeholder）に書いた字で、マスの段落が折り返す：目印（しおり）がなくても、その仕組みの id で伝える
    body = ODT_TABLE.replace('<text:p text:style-name="P1"/>',
                             '<text:p text:style-name="P1"><text:placeholder text:placeholder-type="text" '
                             'text:description="氏名">&lt;氏名&gt;</text:placeholder></text:p>')
    src = tmp_path / "a.odt"
    src.write_bytes(_odt(body))
    assert "氏名" in [f["id"] for f in api.fields(str(src))]
    r = api.fill(str(src), {"氏名": NAME * 3}, str(tmp_path / "o.odt"))
    hit = _wrap(r)
    assert r["verified"] and hit and hit[0].startswith("氏名: "), r["warnings"]


def test_odt_boxed_form_control_gets_no_wrap_estimate(tmp_path):
    # 枠のある入力の仕組み（form:text）は、字が段落に流れない：段落の行は増えない（枠に収まるかは見積もらない）
    forms = ('<office:forms><form:form form:name="Standard"><form:text form:name="住所" xml:id="c2" form:id="c2" '
             'form:current-value=""/></form:form></office:forms>')
    body = forms + ODT_TABLE.replace('<text:p text:style-name="P1"/>', '<text:p text:style-name="P1"><draw:control draw:control="c2"/></text:p>')
    src = tmp_path / "a.odt"
    src.write_bytes(_odt(body))
    r = api.fill(str(src), {"住所": NAME * 6}, str(tmp_path / "o.odt"))
    assert r["verified"] and "wrap" not in r and not _wrap(r), r["warnings"]


# ====================================================================== HWPX
def _hsec(*paras: str) -> str:
    first = hpara(hrun('<hp:secPr id="" textDirection="HORIZONTAL" outlineShapeIDRef="1"/>' + ht("신청서")))
    return HXML + f"<hs:sec {HNS}>" + first + "".join(paras) + "</hs:sec>"


def _htc(col: int, inner: str = "", width: int = 10000, margin: str = "") -> str:
    body = hpara(hrun(ht(inner))) if inner else hpara(hrun())
    has = "1" if margin else "0"
    margin = margin or '<hp:cellMargin left="0" right="0" top="0" bottom="0"/>'
    return (f'<hp:tc name="" header="0" hasMargin="{has}" protect="0" editable="0" dirty="0" borderFillIDRef="1">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER">{body}</hp:subList>'
            f'<hp:cellAddr colAddr="{col}" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{width}" height="1000"/>{margin}</hp:tc>')


def _htbl(*cells: str, inmargin: str = "") -> str:
    return hpara(hrun('<hp:tbl id="1" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" rowCnt="1" colCnt="2" cellSpacing="0" '
                      f'borderFillIDRef="1" noAdjust="0">{inmargin}<hp:tr>' + "".join(cells) + "</hp:tr></hp:tbl><hp:t/>"))


HANGUL = "가나다라마바사아자차카타"  # 12 字：10 pt で 120 pt


def test_hwpx_table_cell_uses_the_cell_width(tmp_path):
    # マスの幅 10000 HWPUNIT（100 pt）。10 pt の 12 字（120 pt）は 2 行、8 字（80 pt）は 1 行
    src = tmp_path / "a.hwpx"
    src.write_bytes(make_hwpx(_hsec(_htbl(_htc(0, "성명"), _htc(1)))))
    cid = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("T1"))
    r = api.fill(str(src), {cid: HANGUL}, str(tmp_path / "o.hwpx"))
    hit = _wrap(r)
    assert r["verified"] and hit and hit[0].startswith(cid) and "表のマスの幅 約 100 ポイント・字の大きさ 10 ポイント" in hit[0], r["warnings"]
    assert "page_growth" in r and r["page_growth"]["recorded_pages_before"] is None  # HWPX には、ページ数の記録がない
    r = api.fill(str(src), {cid: HANGUL[:8]}, str(tmp_path / "o2.hwpx"))
    assert not _wrap(r), r["warnings"]


def test_hwpx_cell_margins_come_from_the_cell_or_the_table(tmp_path):
    # マスが自分の余白を持つ（hasMargin="1"）：100 − 15 − 15 = 70 pt
    own = '<hp:cellMargin left="1500" right="1500" top="0" bottom="0"/>'
    src = tmp_path / "a.hwpx"
    src.write_bytes(make_hwpx(_hsec(_htbl(_htc(0, "성명"), _htc(1, margin=own)))))
    cid = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("T1"))
    hit = _wrap(api.fill(str(src), {cid: HANGUL[:8]}, str(tmp_path / "o.hwpx")))
    assert hit and "約 70 ポイント" in hit[0], hit
    # 持たないマスは、表の余白（hp:inMargin）：100 − 5 − 5 = 90 pt
    src.write_bytes(make_hwpx(_hsec(_htbl(_htc(0, "성명"), _htc(1), inmargin='<hp:inMargin left="500" right="500" top="0" bottom="0"/>'))))
    hit = _wrap(api.fill(str(src), {cid: HANGUL[:10]}, str(tmp_path / "o2.hwpx")))
    assert hit and "約 90 ポイント" in hit[0], hit


def test_hwpx_body_paragraph_uses_the_line_width_hangul_left(tmp_path):
    # 本文の段落：한글が残した行の幅（hp:lineseg の horzsize 42520＝425 pt）。10 pt の 48 字（480 pt）は 2 行
    old = "주소:    "
    r = _fill(tmp_path, "a.hwpx", make_hwpx(_hsec(hpara(hrun(ht(old))))), {"@1": {"old": old, "new": "주소:" + HANGUL * 4}})
    hit = _wrap(r)
    assert hit and "한글が残した行の幅（行の配置の控え） 約 425 ポイント" in hit[0], r["warnings"]


def test_hwpx_click_here_field_is_named_by_its_id_and_boxed_edit_gets_none(tmp_path):
    src = tmp_path / "a.hwpx"
    src.write_bytes(make_hwpx())
    r = api.fill(str(src), {"전화번호": HANGUL * 5}, str(tmp_path / "o.hwpx"))  # 누름틀：字が段落に流れる
    hit = _wrap(r)
    assert r["verified"] and hit and hit[0].startswith("전화번호: "), r["warnings"]
    r = api.fill(str(src), {"나이": "123"}, str(tmp_path / "o2.hwpx"))  # 枠のある入力（hp:edit）
    assert r["verified"] and "wrap" not in r


# ====================================================================== PowerPoint
def _psp(paras: str, cx: int | None = 2540000, body: str = "<a:bodyPr/>", ph: str = "", lst: str = "<a:lstStyle/>", sid: int = 2) -> str:
    xfrm = f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="500000"/></a:xfrm>' if cx else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="枠"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr>{xfrm}</p:spPr>'
            f"<p:txBody>{body}{lst}{paras}</p:txBody></p:sp>")


def _ppara(text: str, sz: int | None = 1800) -> str:
    rpr = f'<a:rPr lang="ja-JP" sz="{sz}"/>' if sz else '<a:rPr lang="ja-JP"/>'
    return f"<a:p><a:r>{rpr}<a:t>{text}</a:t></a:r></a:p>"


def _repart(blob: bytes, changes: dict[str, str]) -> bytes:
    """パッケージの部品の中の字を置き換える（レイアウト・マスター・presentation.xml を作り変える）。"""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename in changes:
                data = changes[info.filename](data.decode("utf-8")).encode("utf-8")
            zout.writestr(info.filename, data)
    return out.getvalue()


PP_OLD = "氏名：　"
PP_NEW = "氏名：" + NAME * 3  # 15 字：18 pt で 270 pt


def _pp(tmp_path, slide: str, change: dict | None = None, new: str = PP_NEW) -> dict:
    blob = pm.make([slide])
    if change:
        blob = _repart(blob, change)
    return _fill(tmp_path, "a.pptx", blob, {"@1": {"old": PP_OLD, "new": new}})


def test_pptx_shape_width_minus_insets(tmp_path):
    # 図形の幅 2540000 EMU（200 pt）から、左右の余白（既定 91440 EMU＝7.2 pt ずつ）＝約 186 pt。18 pt の 15 字（270 pt）は 2 行
    r = _pp(tmp_path, _psp(_ppara(PP_OLD)))
    hit = _wrap(r)
    assert hit and "1 行から 2 行" in hit[0] and "約 186 ポイント・字の大きさ 18 ポイント" in hit[0], r["warnings"]
    assert "page_growth" not in r and PAGES not in r["say_to_user"]  # スライドは、行が増えてもページは増えない
    assert not _wrap(_pp(tmp_path, _psp(_ppara(PP_OLD)), new="氏名：" + NAME))


def test_pptx_no_wrap_shape_and_wide_shape_give_no_warning(tmp_path):
    assert not _wrap(_pp(tmp_path, _psp(_ppara(PP_OLD), body='<a:bodyPr wrap="none"/>')))
    assert not _wrap(_pp(tmp_path, _psp(_ppara(PP_OLD), cx=5080000)))


def test_pptx_shape_insets_and_autofit_scale(tmp_path):
    # 左右の余白 lIns・rIns 0：200 pt。字を 50% に縮める（normAutofit fontScale）と 9 pt の 15 字（135 pt）で 1 行
    hit = _wrap(_pp(tmp_path, _psp(_ppara(PP_OLD), body='<a:bodyPr lIns="0" rIns="0"/>')))
    assert hit and "約 200 ポイント" in hit[0], hit
    assert not _wrap(_pp(tmp_path, _psp(_ppara(PP_OLD), body='<a:bodyPr><a:normAutofit fontScale="50000"/></a:bodyPr>')))


def test_pptx_group_scales_the_shape(tmp_path):
    # グループが中の図形を半分に縮めて置く（ext 1270000 ／ chExt 2540000）：幅 100 pt − 余白
    grp = ('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="グループ"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
           '<a:off x="0" y="0"/><a:ext cx="1270000" cy="250000"/><a:chOff x="0" y="0"/><a:chExt cx="2540000" cy="500000"/></a:xfrm></p:grpSpPr>'
           + _psp(_ppara(PP_OLD, 900)) + "</p:grpSp>")
    # 9 pt の 15 字は 135 pt：縮めなければ 186 pt に入るが、半分の幅（約 86 pt）には入らない
    hit = _wrap(_pp(tmp_path, grp))
    assert hit and "約 86 ポイント" in hit[0], hit
    assert not _wrap(_pp(tmp_path, _psp(_ppara(PP_OLD, 900))))


def test_pptx_placeholder_takes_width_and_size_from_the_layout(tmp_path):
    # スライドのプレースホルダーに幅も字の大きさもない：レイアウトの同じプレースホルダーの幅（200 pt）と、その lstStyle の大きさ（24 pt）
    lay = _psp(_ppara("本文"), ph='<p:ph type="body" idx="1"/>',
               lst='<a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>')
    change = {"ppt/slideLayouts/slideLayout1.xml": lambda x: x.replace("</p:spTree>", lay + "</p:spTree>")}
    r = _pp(tmp_path, _psp(_ppara(PP_OLD, None), cx=None, ph='<p:ph type="body" idx="1"/>'), change, new="氏名：" + NAME * 2)
    hit = _wrap(r)
    assert hit and "約 186 ポイント・字の大きさ 24 ポイント" in hit[0], r["warnings"]  # 11 字 × 24 pt = 264 pt


def test_pptx_size_falls_back_to_master_text_styles_then_presentation(tmp_path):
    slide = _psp(_ppara(PP_OLD, None))
    # マスターの txStyles の otherStyle（プレースホルダーでない図形）：28 pt
    tx = '<p:txStyles><p:otherStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></p:otherStyle></p:txStyles>'
    change = {"ppt/slideMasters/slideMaster1.xml": lambda x: x.replace("</p:sldMaster>", tx + "</p:sldMaster>")}
    hit = _wrap(_pp(tmp_path, slide, change, new="氏名：" + NAME))
    assert hit and "字の大きさ 28 ポイント" in hit[0] and "仮" not in hit[0], hit  # 7 字 × 28 pt = 196 pt
    # presentation.xml の defaultTextStyle：32 pt
    dts = '<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:defaultTextStyle>'
    change = {"ppt/presentation.xml": lambda x: x.replace("</p:presentation>", dts + "</p:presentation>")}
    hit = _wrap(_pp(tmp_path, slide, change, new="氏名：" + NAME))
    assert hit and "字の大きさ 32 ポイント" in hit[0], hit


def test_pptx_unstated_size_is_assumed_and_said(tmp_path):
    hit = _wrap(_pp(tmp_path, _psp(_ppara(PP_OLD, None))))
    assert hit and "字の大きさ 18 ポイント" in hit[0] and "仮" in hit[0], hit


def test_pptx_table_cell_uses_the_grid_width(tmp_path):
    # 表の列 1000000 EMU（78.7 pt）から左右の余白 7.2 pt ずつ＝約 64 pt
    src = tmp_path / "a.pptx"
    src.write_bytes(pm.make([pm.table([["氏名", ""]])]))
    cid = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("T1"))
    r = api.fill(str(src), {cid: NAME * 3}, str(tmp_path / "o.pptx"))
    hit = _wrap(r)
    assert r["verified"] and hit and hit[0].startswith(cid) and "表のマスの幅 約 64 ポイント" in hit[0], r["warnings"]


# ====================================================================== Word：入力の仕組みに書いた段落の名前
def test_docx_content_control_wrap_is_named_by_the_control(tmp_path):
    from test_docx import make_docx

    sdt = ('<w:sdt><w:sdtPr><w:alias w:val="氏名"/><w:tag w:val="name"/><w:id w:val="7"/><w:showingPlcHdr/><w:text/></w:sdtPr>'
           '<w:sdtContent><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>入力</w:t></w:r></w:sdtContent></w:sdt>')
    body = ('<w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="2136"/></w:tblGrid><w:tr>'
            '<w:tc><w:p><w:r><w:t>氏名</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="2136" w:type="dxa"/></w:tcPr><w:p>{sdt}</w:p></w:tc></w:tr></w:tbl>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1701" w:right="1701" w:top="1985" w:bottom="1701"/></w:sectPr>')
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(body))
    assert "氏名" in [f["id"] for f in api.fields(str(src))]
    r = api.fill(str(src), {"氏名": NAME * 3}, str(tmp_path / "o.docx"))
    hit = _wrap(r)
    assert r["verified"] and hit and hit[0].startswith("氏名: "), r["warnings"]


# ====================================================================== .xlsb：セルに収まるか
def _xlsb(tmp_path, rows: bytes, head: bytes = b"") -> str:
    src = tmp_path / "a.xlsb"
    src.write_bytes(xm.package({"S": xm.sheet(rows, head=head)}))
    return str(src)


def test_xlsb_long_text_is_clipped_by_a_filled_neighbour(tmp_path):
    src = _xlsb(tmp_path, xm.row(0) + xm.st(0, "氏名") + xm.blank(1, 1) + xm.st(2, "備考"))
    r = api.fill(src, {"S!B1": NAME * 5}, str(tmp_path / "o.xlsb"))
    assert r["verified"] and r["fit"] and "隣の C1 の字で切れて見える見込み" in r["fit"][0]["message"], r["warnings"]
    assert r["fit"][0]["message"] in r["warnings"]
    r = api.fill(src, {"S!B1": "架空"}, str(tmp_path / "o2.xlsb"))
    assert "fit" not in r


def test_xlsb_long_text_next_to_empty_cells_has_no_warning(tmp_path):
    src = _xlsb(tmp_path, xm.row(0) + xm.st(0, "氏名") + xm.blank(1, 1))
    r = api.fill(src, {"S!B1": NAME * 5}, str(tmp_path / "o.xlsb"))
    assert r["verified"] and "fit" not in r, r["warnings"]


def test_xlsb_wide_number_shows_hashes(tmp_path):
    src = _xlsb(tmp_path, xm.row(0) + xm.st(0, "金額") + xm.blank(1, 1))
    r = api.fill(src, {"S!B1": 123456789012345}, str(tmp_path / "o.xlsb"))
    assert r["verified"] and r["fit"] and "####" in r["fit"][0]["message"], r["warnings"]


def test_xlsb_right_to_left_sheet_checks_both_sides_and_says_so(tmp_path):
    # BrtBeginWsView（MS-XLSB 2.4.304）の fRightToLeft：はみ出す向きを決めず、両方の隣を見る
    view = xm.rec(133) + xm.rec(137, struct.pack("<HIIIBBHHHHHI", 0x20, 0, 0, 0, 64, 0, 0, 100, 0, 0, 0, 0)) + xm.rec(138) + xm.rec(134)
    src = _xlsb(tmp_path, xm.row(0) + xm.st(0, "氏名") + xm.blank(1, 1), head=view)
    r = api.fill(src, {"S!B1": NAME * 5}, str(tmp_path / "o.xlsb"))
    assert r["verified"] and r["fit"], r["warnings"]
    m = r["fit"][0]["message"]
    assert "隣の A1 の字で切れて見えるかもしれません" in m and "はみ出す向きを決めていない" in m, m


# ====================================================================== .xls：右から左へ表示するシート
def _xls(tmp_path, rtl: bool) -> str:
    import test_xls_controls as xc

    data = xc.workbook([{"name": "S", "cells": {(0, 0): "氏名", (1, 1): "備考"}}])
    win = struct.pack("<HHH", 0x023E, 18, 0x06B6)  # WINDOW2（MS-XLS 2.4.346）の頭と grbit
    assert data.count(win) == 1
    if rtl:
        data = data.replace(win, struct.pack("<HHH", 0x023E, 18, 0x06B6 | 0x40))  # fRightToLeft
    src = tmp_path / "a.xls"
    src.write_bytes(xc.cfb("Workbook", data))
    return str(src)


def test_xls_right_to_left_sheet_checks_both_sides_and_says_so(tmp_path):
    r = api.fill(_xls(tmp_path, False), {"S!B1": NAME * 5}, str(tmp_path / "o.xls"), force=True)
    assert r["verified"] and "fit" not in r, r["warnings"]  # 左から右：右の C1 は空
    src = _xls(tmp_path, True)
    assert "右から左へ表示するシート" in "\n".join(str(x[1]) for x in api.window(src)[2])
    r = api.fill(src, {"S!B1": NAME * 5}, str(tmp_path / "o2.xls"), force=True)
    assert r["verified"] and r["fit"], r["warnings"]
    m = r["fit"][0]["message"]
    assert "隣の A1 の字で切れて見えるかもしれません" in m and "はみ出す向きを決めていない" in m, m


# ====================================================================== .xlsx：はみ出す向き
def _xlsx(tmp_path, build) -> str:
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    build(ws)
    p = tmp_path / "a.xlsx"
    wb.save(p)
    return str(p)


def test_xlsx_left_to_right_text_is_checked_on_its_side_only(tmp_path):
    def build(ws):
        ws["A1"] = "氏名"

    src = _xlsx(tmp_path, build)
    r = api.fill(src, {"S!B1": NAME * 5}, str(tmp_path / "o.xlsx"), force=True)
    assert "fit" not in r, r["warnings"]  # 左寄せ：右の C1 は空。左の A1 の字にはかからない


def test_xlsx_right_to_left_sheet_checks_both_sides(tmp_path):
    def build(ws):
        ws["A1"] = "氏名"
        ws.sheet_view.rightToLeft = True

    src = _xlsx(tmp_path, build)
    r = api.fill(src, {"S!B1": NAME * 5}, str(tmp_path / "o.xlsx"), force=True)
    assert r["fit"] and "隣の A1 の字で切れて見えるかもしれません" in r["fit"][0]["message"], r["warnings"]


def test_xlsx_right_to_left_text_in_a_general_cell_checks_both_sides(tmp_path):
    def build(ws):
        ws["A1"] = "الاسم"

    src = _xlsx(tmp_path, build)
    r = api.fill(src, {"S!B1": "محمد بن عبد الله بن عبد المطلب بن هاشم"}, str(tmp_path / "o.xlsx"), force=True)
    assert r["fit"] and "かもしれません" in r["fit"][0]["message"] and "A1" in r["fit"][0]["message"], r["warnings"]


def test_overflow_sides_rule():
    from todoku.xlsx.cellcheck import blocked_by, overflow_sides

    assert overflow_sides("general", False, "架空") == ([1], True)
    assert overflow_sides("left", False, "محمد") == ([1], True)  # 寄せ方が決まっていれば、その側
    assert overflow_sides("right", False, "abc") == ([-1], True)
    assert overflow_sides("center", True, "abc") == ([1, -1], True)
    assert overflow_sides("general", False, "123 محمد") == ([1, -1], False)  # 初めの、向きの強い字が右から左
    assert overflow_sides("general", False, "abc محمد") == ([1], True)
    # 右から左へ並ぶシート：画面の左が、列の番号の増える向き
    assert overflow_sides("right", True, "abc") == ([1], True) and overflow_sides("left", True, "abc") == ([-1], True)
    assert overflow_sides("general", True, "abc") == ([1, -1], False)
    # 読む向きが右から左のセル（readingOrder 2）は、標準の寄せ方でも右
    assert overflow_sides("general", False, "abc", 2) == ([-1], True) and overflow_sides("general", True, "abc", 2) == ([1], True)
    assert overflow_sides("general", False, "محمد", 1) == ([1, -1], False) and overflow_sides("left", False, "abc", 2) == ([1], True)
    # 中央寄せは両側へ半分ずつ：10 はみ出すとき、幅 6 の空の隣の先の字にはかからない
    filled = {1, 5}
    assert blocked_by([1, -1], True, 10, 3, lambda c: 6.0, lambda c: c not in filled) is None
    assert blocked_by([1, -1], False, 10, 3, lambda c: 6.0, lambda c: c not in filled) == 5
    assert blocked_by([1], True, 10, 3, lambda c: 6.0, lambda c: c != 4) == 4


# ====================================================================== .ods：セルに収まるか
ODS_STYLES = (
    '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2cm"/></style:style>'
    '<style:style style:name="ta1" style:family="table"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>'
    '<style:style style:name="ta2" style:family="table"><style:table-properties table:display="true" style:writing-mode="rl-tb"/></style:style>'
    '<style:style style:name="ceW" style:family="table-cell"><style:table-cell-properties fo:wrap-option="wrap"/></style:style>'
    '<style:style style:name="ceS" style:family="table-cell"><style:table-cell-properties style:shrink-to-fit="true"/></style:style>'
    '<style:style style:name="ceR" style:family="table-cell"><style:table-cell-properties style:text-align-source="fix"/>'
    '<style:paragraph-properties fo:text-align="end"/></style:style>'
    '<style:style style:name="ceT" style:family="table-cell"><style:table-cell-properties style:rotation-angle="90"/></style:style>'
    '<style:style style:name="ceB" style:family="table-cell"><style:text-properties fo:font-size="20pt"/></style:style>')


def _oc(text: str = "", style: str = "", span: int = 0) -> str:
    st = f' table:style-name="{style}"' if style else ""
    sp = f' table:number-columns-spanned="{span}" table:number-rows-spanned="1"' if span else ""
    if not text:
        return f"<table:table-cell{st}{sp}/>"
    return f'<table:table-cell{st}{sp} office:value-type="string" calcext:value-type="string"><text:p>{text}</text:p></table:table-cell>'


def _ods(tmp_path, *rows: str, table_style: str = "ta1") -> str:
    body = (f'<office:spreadsheet><table:table table:name="S" table:style-name="{table_style}">'
            '<table:table-column table:style-name="co1" table:number-columns-repeated="6"/>'
            + "".join(f"<table:table-row>{r}</table:table-row>" for r in rows) + "</table:table></office:spreadsheet>")
    src = tmp_path / "a.ods"
    src.write_bytes(_odf("application/vnd.oasis.opendocument.spreadsheet", body, ODS_STYLES))
    return str(src)


TEN = NAME * 2 + "架空"  # 10 字：10 pt で 100 pt。列 2 cm（56.7 pt）から左右の余白を引いた約 55 pt に入らない


def _ods_fit(tmp_path, src: str, values: dict) -> list[str]:
    r = api.fill(src, values, str(tmp_path / "o.ods"), force=True, overwrite=True)
    assert r["verified"], (r.get("unexpected"), r.get("read_back_differences"))
    assert all(x["message"] in r["warnings"] for x in r.get("fit") or [])
    return [x["message"] for x in r.get("fit") or []]


def test_ods_long_text_is_clipped_by_a_filled_neighbour(tmp_path):
    src = _ods(tmp_path, _oc("氏名") + _oc() + _oc("備考") + _oc() * 3, _oc("住所") + _oc() * 5)
    got = _ods_fit(tmp_path, src, {"S!B1": TEN})
    assert got and "隣の C1 の字で切れて見える見込み" in got[0], got
    assert not _ods_fit(tmp_path, src, {"S!B1": NAME})  # 4 字（40 pt）は入る
    assert not _ods_fit(tmp_path, src, {"S!B2": TEN})  # 右の隣が空なら、はみ出して見える
    # 同じときに書く隣のセルも、字のある隣
    got = _ods_fit(tmp_path, src, {"S!B2": TEN, "S!C2": "あ"})
    assert got and "隣の C2 の字" in got[0], got


def test_ods_far_neighbour_only_counts_within_the_overflow(tmp_path):
    # 10 字（100 pt）は、B と C の 2 列（約 111 pt）に入る：D1 の字にはかからない。20 字なら D1 で切れる
    src = _ods(tmp_path, _oc("氏名") + _oc() + _oc() + _oc("備考") + _oc() * 2)
    assert not _ods_fit(tmp_path, src, {"S!B1": TEN})
    got = _ods_fit(tmp_path, src, {"S!B1": TEN * 2})
    assert got and "隣の D1 の字" in got[0], got


def test_ods_font_size_of_the_cell_style_counts(tmp_path):
    src = _ods(tmp_path, _oc("氏名") + _oc(style="ceB") + _oc("備考") + _oc() * 3)
    got = _ods_fit(tmp_path, src, {"S!B1": NAME})  # 20 pt の 4 字は 80 pt
    assert got and "隣の C1 の字" in got[0], got


def test_ods_wide_number_shows_hashes(tmp_path):
    src = _ods(tmp_path, _oc("金額") + _oc() + _oc() * 4)
    got = _ods_fit(tmp_path, src, {"S!B1": 123456789012345})
    assert got and "「###」と表示される見込み" in got[0], got
    assert not _ods_fit(tmp_path, src, {"S!B1": 1234})


def test_ods_merged_cell_clips_at_its_own_edge(tmp_path):
    src = _ods(tmp_path, _oc("氏名") + _oc(span=2) + "<table:covered-table-cell/>" + _oc() * 3)
    got = _ods_fit(tmp_path, src, {"S!B1": TEN * 2})
    assert got and "結合したセルの幅に収まらず" in got[0], got
    assert not _ods_fit(tmp_path, src, {"S!B1": TEN})  # 2 列ぶん（約 111 pt）に入る


def test_ods_shrink_to_fit_cell_says_the_scale(tmp_path):
    src = _ods(tmp_path, _oc("氏名") + _oc(style="ceS") + _oc() * 4)
    got = _ods_fit(tmp_path, src, {"S!B1": TEN})
    assert got and "字が縮んで表示されます" in got[0] and "約 54%" in got[0], got


def test_ods_right_aligned_text_overflows_to_the_left(tmp_path):
    src = _ods(tmp_path, _oc("氏名") + _oc(style="ceR") + _oc() * 4)
    got = _ods_fit(tmp_path, src, {"S!B1": TEN})
    assert got and "隣の A1 の字で切れて見える見込み" in got[0], got


def test_ods_wrap_cell_gets_the_wrap_warning_and_rotated_cell_none(tmp_path):
    src = _ods(tmp_path, _oc("氏名") + _oc(style="ceW") + _oc("備考") + _oc(style="ceT") + _oc("印") + _oc())
    r = api.fill(src, {"S!B1": TEN}, str(tmp_path / "o.ods"), force=True)
    assert "fit" not in r and _wrap(r), r["warnings"]
    assert not _ods_fit(tmp_path, src, {"S!D1": TEN})  # 回した字は、見積もらない


def test_ods_right_to_left_sheet_and_text_check_both_sides(tmp_path):
    src = _ods(tmp_path, _oc("氏名") + _oc() + _oc() * 4, table_style="ta2")
    got = _ods_fit(tmp_path, src, {"S!B1": TEN})
    assert got and "隣の A1 の字で切れて見えるかもしれません" in got[0] and "はみ出す向きを決めていない" in got[0], got
    src = _ods(tmp_path, _oc("الاسم") + _oc() + _oc() * 4)
    got = _ods_fit(tmp_path, src, {"S!B1": "محمد بن عبد الله بن عبد المطلب بن هاشم"})
    assert got and "かもしれません" in got[0], got
