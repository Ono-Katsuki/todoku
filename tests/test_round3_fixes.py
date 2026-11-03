"""第3回の見直しで見つかった、証明と読み戻しを通ってしまう誤った書き込みなどの直し（合成した書類で確かめる）。"""

from __future__ import annotations

import io
import struct
import zipfile

import pytest

from todoku import api

openpyxl = pytest.importorskip("openpyxl")


def _xlsx(tmp_path, build, name="a.xlsx"):
    wb = openpyxl.Workbook()
    build(wb.active)
    p = tmp_path / name
    wb.save(p)
    return p


def _cell(path, ref):
    return openpyxl.load_workbook(path)["Sheet"][ref].value


# ====================================================================== 1. 字のセルに true/false
def test_bool_on_text_cell_checks_the_box_glyph_not_replaces_text(tmp_path):
    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("B3", "   □移動支援（介護あり）"))
    r = api.fill(str(p), {"Sheet!B3": True}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"]
    assert _cell(r["out_path"], "B3") == "   ■移動支援（介護あり）"
    assert any("「□」を「■」" in w for w in r["warnings"])
    assert "■移動支援（介護あり）" in r["say_to_user"] and "［" not in r["say_to_user"].split("\n")[1]


def test_bool_on_text_without_or_with_many_boxes_is_refused(tmp_path):
    p = _xlsx(tmp_path, lambda ws: (ws.__setitem__("B3", "移動支援"), ws.__setitem__("B4", "□有　□無")))
    with pytest.raises(api.ApiError, match="箱の字（□・☐ など）がありません"):
        api.fill(str(p), {"Sheet!B3": True}, str(tmp_path / "o.xlsx"))
    with pytest.raises(api.ApiError, match="2 つあります"):
        api.fill(str(p), {"Sheet!B4": True}, str(tmp_path / "o.xlsx"))


def test_replacing_printed_text_fails_proof_unless_explicit(tmp_path):
    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("B3", "□移動支援（介護あり）"))
    r = api.fill(str(p), {"Sheet!B3": "山田"}, str(tmp_path / "o.xlsx"))
    assert not r["proof_ok"] and not r["verified"] and ".unverified" in r["out_path"]
    assert any("印刷された字「□移動支援（介護あり）」" in u for u in r["unexpected"])
    r = api.fill(str(p), {"Sheet!B3": {"value": "山田", "overwrite_text": True}}, str(tmp_path / "o2.xlsx"))
    assert r["proof_ok"] and r["verified"] and _cell(r["out_path"], "B3") == "山田"
    # 空欄に書き足しただけ（元の字が順に残る）は置き換えではない
    r = api.fill(str(p), {"Sheet!B3": "☑移動支援（介護あり）"}, str(tmp_path / "o3.xlsx"))
    assert r["proof_ok"], r["unexpected"]


# ====================================================================== 5. id：結合の範囲・字を目印
def test_merged_range_id_is_the_top_left_cell(tmp_path):
    def build(ws):
        ws.merge_cells("E3:J3")
    p = _xlsx(tmp_path, build)
    r = api.fill(str(p), {"Sheet!E3:J3": "12"}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and _cell(r["out_path"], "E3") == 12
    with pytest.raises(api.ApiError, match="結合したセルの範囲ではありません"):
        api.fill(str(p), {"Sheet!E3:F3": "12"}, str(tmp_path / "o2.xlsx"))


def test_xlsx_old_new_anchor(tmp_path):
    p = _xlsx(tmp_path, lambda ws: (ws.__setitem__("A1", "令和　　年　　月　　日"), ws.__setitem__("A2", "第　　号"),
                                    ws.__setitem__("A3", "第　　号")))
    r = api.fill(str(p), {"@1": {"old": "令和　　年", "new": "令和8年"}}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and r["anchored"] == {"@1": "Sheet!A1"}
    assert _cell(r["out_path"], "A1") == "令和8年　　月　　日"
    with pytest.raises(api.ApiError, match="2 か所にあり"):
        api.fill(str(p), {"@1": {"old": "第　　号", "new": "第12号"}}, str(tmp_path / "o2.xlsx"))
    r = api.fill(str(p), {"@1": {"old": "第　　号", "new": "第12号", "cell": "Sheet!A3"}}, str(tmp_path / "o3.xlsx"))
    assert r["verified"] and _cell(r["out_path"], "A3") == "第12号" and _cell(r["out_path"], "A2") == "第　　号"


# ====================================================================== 4. Excel の囲み：border
def test_xlsx_circle_border_draws_a_rectangle(tmp_path):
    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("B2", "有・無"))
    r = api.fill(str(p), {"@c1": {"circle": "無", "cell": "Sheet!B2", "style": "border"}}, str(tmp_path / "o.xlsx"))
    assert r["verified"], (r["unexpected"], r["read_back_differences"])
    assert r["circled"][0]["style"] == "border" and "四角" in r["circled"][0]["how"]
    z = zipfile.ZipFile(r["out_path"])
    xml = b"".join(z.read(n) for n in z.namelist() if n.startswith("xl/drawings/"))
    assert b'prst="rect"' in xml and b'prst="ellipse"' not in xml


# ====================================================================== 3. 囲む字の決め方（Word・Excel）
def _docx(body: str) -> bytes:
    from test_docx import make_docx

    return make_docx(body)


def _p(t: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


def test_word_circle_occurrence_counts_within_near_and_skips_the_heading(tmp_path):
    p = tmp_path / "a.docx"
    p.write_bytes(_docx(_p("生年月日　昭和　年　月　日") + _p("通院頻度　（月・週）　回") + _p("治療の有無　有・無")))
    r = api.fill(str(p), {"@c1": {"circle": "月", "near": "通院頻度", "occurrence": 1, "style": "border"}},
                 str(tmp_path / "o.docx"))
    assert "（【月】・週）" in r["circled"][0]["around"], r["circled"]
    r = api.fill(str(p), {"@c1": {"circle": "無", "near": "治療の有無", "style": "border"}}, str(tmp_path / "o2.docx"))
    assert "有・【無】" in r["circled"][0]["around"]
    with pytest.raises(api.ApiError, match="near の「治療の有無」の字の中にしか"):
        q = tmp_path / "b.docx"
        q.write_bytes(_docx(_p("治療の有無")))
        api.fill(str(q), {"@c1": {"circle": "無", "near": "治療の有無"}}, str(tmp_path / "o3.docx"))


def test_word_circle_in_hanging_indent_is_marked_uncertain(tmp_path):
    p = tmp_path / "a.docx"
    p.write_bytes(_docx('<w:p><w:pPr><w:ind w:left="722" w:hanging="722"/></w:pPr><w:r><w:t>１　有・無</w:t></w:r></w:p>'))
    r = api.fill(str(p), {"@c1": {"circle": "無"}}, str(tmp_path / "o.docx"))
    assert "ぶら下げ" in r["circled"][0]["uncertain"] and "border" in r["circled"][0]["uncertain"]


# ====================================================================== 2. 書いた字に、取り消し線・赤を引き継がない
def test_docx_empty_cell_does_not_inherit_strike_and_red(tmp_path):
    tbl = ('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>氏名</w:t></w:r></w:p></w:tc><w:tc><w:p><w:pPr><w:rPr><w:strike/>'
           '<w:color w:val="FF0000"/><w:sz w:val="24"/></w:rPr></w:pPr></w:p></w:tc></w:tr></w:tbl>')
    p = tmp_path / "a.docx"
    p.write_bytes(_docx(tbl))
    fid = next(f["id"] for f in api.fields(str(p)) if f["id"].startswith("T1.R1.C2"))
    r = api.fill(str(p), {fid: "山田　花子"}, str(tmp_path / "o.docx"))
    assert r["verified"], r["unexpected"]
    xml = zipfile.ZipFile(r["out_path"]).read("word/document.xml")
    run = xml[xml.index(b"\xe5\xb1\xb1\xe7\x94\xb0") - 200:xml.index(b"\xe5\xb1\xb1\xe7\x94\xb0")]
    assert b"strike" not in run.split(b"<w:r>")[-1] and b"FF0000" not in run.split(b"<w:r>")[-1]
    assert b'<w:sz w:val="24"/>' in run  # 字の大きさは引き継ぐ
    assert any("取り消し線・字の色" in w for w in r["warnings"])
    assert not r.get("written_formatting")


def test_doc_clean_chpx_drops_strike_and_colour():
    from todoku.doc.fill import clean_chpx

    grpprl = struct.pack("<HB", 0x0837, 1) + struct.pack("<HB", 0x2A42, 6) + struct.pack("<HH", 0x4A43, 21)
    out, names = clean_chpx(grpprl)
    assert out == struct.pack("<HH", 0x4A43, 21) and names == ["取り消し線", "字の色"]
    out, names = clean_chpx(struct.pack("<HB", 0x0837, 0))  # 「なし」の値は、名前を伝えない
    assert names == []


def test_doc_anchored_key_is_the_given_key():
    from todoku.docx.fill import Patch

    p = Patch("P1.w1", "input", "x", "doc", anchor={"old": "a", "new": "x"}, key="@1")
    assert {(q.key or q.id): q.id for q in [p] if q.anchor} == {"@1": "P1.w1"}


# ====================================================================== 6. 本人に伝えること
def test_hand_only_places_from_grade_reach_say_to_user(tmp_path):
    p = _xlsx(tmp_path, lambda ws: (ws.__setitem__("A1", "氏名"), ws.__setitem__("A2", "□同意する")))
    api.grade(str(p), findings=[{"check": "cannot_do", "title": "自署と押印", "detail": "申請者の署名欄に自署・押印", "where": ["A5"]}])
    r = api.fill(str(p), {"Sheet!A2": True}, str(tmp_path / "o.xlsx"))
    assert r["hand_only"][0]["title"] == "自署と押印"
    assert "手で行う所" in r["say_to_user"] and "自署と押印" in r["say_to_user"]


def test_verbatim_label_drops_a_guessed_prefix_not_in_the_row():
    from todoku.api import _verbatim_label

    f = {"label": "個人番号：身体障害者 手帳番号［▶］療育手帳［　］", "context": "身体障害者 手帳番号 ｜ ［▶E18:H18］ ｜ 療育手帳 ｜ ［K18］"}
    assert _verbatim_label(f) == "身体障害者 手帳番号［▶］療育手帳［　］"
    f = {"label": "電話：［▶］", "context": "電話 ｜ ［▶B2］"}
    assert _verbatim_label(f) == "電話：［▶］"


def test_overlapping_rects():
    from todoku.api import _overlaps

    assert _overlaps((1, [0, 0, 10, 10]), (1, [1, 1, 9, 9]))
    assert not _overlaps((1, [0, 0, 10, 10]), (2, [1, 1, 9, 9]))
    assert not _overlaps((1, [0, 0, 10, 10]), (1, [9, 9, 20, 20]))


# ====================================================================== 7. Word の画像だけのページ
def test_word_page_of_only_images_is_e():
    from test_grade_followup import _docx as gdocx
    from test_grade_followup import _pic

    from todoku.grade import grade_blob

    body = ('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Dashboard Q3</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' + _pic(1)
            + '<w:p><w:r><w:lastRenderedPageBreak/></w:r></w:p>' + _pic(2) + _pic(3))
    g = grade_blob(gdocx(body), "a.docx")
    assert g.grade == "E" and g.reasons[0].title == "画像だけのページ" and "2 ページ目" in g.reasons[0].detail
    # 区切りのない文書・字のあるページは、ページごとには決めない（注意だけ）
    g = grade_blob(gdocx('<w:p><w:r><w:t>本文の説明です。</w:t></w:r></w:p>' + _pic(1)), "a.docx")
    assert g.grade == "A"


# ====================================================================== 8. 壊れた複合ファイル
def test_broken_cfb_is_a_clean_error(tmp_path):
    from todoku.cfb import build_cfb

    b = build_cfb({"WordDocument": b"\xec\xa5" + b"\0" * 100, "1Table": b"\0" * 100})
    p = tmp_path / "a.doc"
    p.write_bytes(b[:700])
    with pytest.raises(api.ApiError, match="壊れた Office の旧形式"):
        api.window(str(p))
    g = api.grade(str(p))
    assert g["grade"] == "F" and "壊れた Office の旧形式" in g["reasons"][0]["detail"]


# ====================================================================== 9. ZIP の中の HTML・XML
def test_html_and_xml_in_zip_are_read_as_text(tmp_path):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("data.txt", "a\tb\n1\t2\n")
        z.writestr("meta/x.html", "<html><head><title>t</title><script>var x=1;</script></head>"
                                  "<body><h1>Survey &amp; data</h1><p>Line one</p></body></html>")
        z.writestr("meta/x.xml", '<?xml version="1.0"?><r><a>Alberta</a><b><![CDATA[2013]]></b></r>')
    p = tmp_path / "a.zip"
    p.write_bytes(buf.getvalue())
    g = api.grade(str(p))
    assert g["grade"] == "B" and any("HTML/XML は字だけを読む" in r["title"] for r in g["reasons"])
    k, doc, parts = api.window(str(p) + "!/meta/x.html")
    md = parts[0][1]
    assert k == "markup" and "Survey & data" in md and "Line one" in md and "var x" not in md
    k, doc, parts = api.window(str(p) + "!/meta/x.xml")
    assert doc["lines"] == ["Alberta", "2013"]
    assert api.audit(str(p) + "!/meta/x.xml")["ok"]
    with pytest.raises(api.ApiError, match="HTML・XML には書きません"):
        api.fill(str(p) + "!/meta/x.xml", {"x": "1"}, str(tmp_path / "o.zip"))


# ====================================================================== 10. 外字
def test_pua_in_xlsx_and_docx_is_read_b(tmp_path):
    from todoku.grade import grade_blob

    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("A1", "氏名 山太郎"))
    g = grade_blob(p.read_bytes(), "a.xlsx")
    r = next(r for r in g.reasons if r.title.startswith("外字"))
    assert g.grade == "B" and r.where[0].startswith("Sheet!A1「") and "U+E000" in r.where[0]
    g = grade_blob(_docx(_p("住所 東京都区")), "a.docx")
    assert g.grade == "B" and any(r.title.startswith("外字") for r in g.reasons)
