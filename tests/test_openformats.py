"""CSV・TSV・JSON・JSON Lines・RTF・OpenDocument（.odt・.ods）の読み書きのテスト。

テストの中で小さなファイルを組み立てて確かめ、公開の様式（corpus/wide_csv・wide_rtf・wide_odf）でも確かめる（なければ飛ばす）。
"""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

import pytest

from todoku import api
from todoku.book import kind_of
from todoku.grade import grade_blob

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
B = chr(92)  # 逆斜線（RTF の制御語を組み立てる）


# ====================================================================== 形式の見分け
def test_kind_of_by_content_and_name():
    assert kind_of(b"a,b\n1,2\n") == "unknown"  # 名前を渡さなければ、いままでと同じ
    assert kind_of(b"a,b\n1,2\n", "x.csv") == "csv"
    assert kind_of(b"a\tb\n", "x.TSV") == "tsv"
    assert kind_of(b'{"a": 1}', "x.json") == "json"
    assert kind_of(b'{"a": 1}\n', "x.jsonl") == "jsonl"
    assert kind_of((B + "rtf1 x}").join(["{", ""]).encode()) == "rtf"
    assert kind_of(make_odt("<text:p>a</text:p>")) == "odt"
    assert kind_of(make_ods(SHEET)) == "ods"
    assert kind_of(b"\x00\x01garbage", "x.bin") == "unknown"


# ====================================================================== CSV・TSV
def test_csv_reads_and_writes_only_written_bytes():
    from todoku.datafile.csvfile import CsvFile, fill, prove

    # Shift_JIS（CP932）・CRLF・必要なマスだけ引用符。NEC 選定 IBM 拡張文字（0xED40）は、読み直して書いても同じバイトのまま
    blob = '氏名,住所,"電話"\r\n山田,"東京都,千代田区",\r\n'.encode("cp932") + b"\xed\x40,,\r\n"
    f = CsvFile(blob, "a.csv")
    assert f.dec.encoding == "cp932" and f.newline == "\r\n" and f.rows[1] == ["山田", "東京都,千代田区", ""]
    res = fill(f, {"R2.電話": "090", "行3.列「住所」": 'a"b,c', "R4.C1": "佐藤"})
    assert res.blob == '氏名,住所,"電話"\r\n山田,"東京都,千代田区",090\r\n'.encode("cp932") + b'\xed\x40,"a""b,c",\r\n' \
        + "佐藤,,\r\n".encode("cp932")
    pr = prove(blob, res.blob, {(p.row, p.col): p.value for p in res.patches}, "a.csv")
    assert pr.ok, pr.unexpected
    with pytest.raises(ValueError):
        fill(f, {"R6.C1": "x"})  # 行を飛ばしては足さない


def test_csv_bom_quote_all_and_no_final_newline():
    from todoku.datafile.csvfile import CsvFile, fill

    blob = '﻿"a","b"\n"1","2"'.encode("utf-8")
    f = CsvFile(blob, "q.csv")
    assert f.quote_all and not f.ends_with_newline
    res = fill(f, {"R2.C2": "x", "R3.C1": "y"})
    assert res.blob == '﻿"a","b"\n"1","x"\n"y",'.encode("utf-8")


def test_csv_audit_and_grade_and_api(tmp_path):
    from todoku.datafile.audit import run

    blob = "項目,値\n氏名,\n住所,\n".encode("utf-8")
    au = run(blob, "f.csv", "csv")
    assert au.ok, [(c.name, c.items) for c in au.checks if not c.ok]
    assert grade_blob(blob, "f.csv").grade == "S"
    src = tmp_path / "f.csv"
    src.write_bytes(blob)
    r = api.fill(str(src), {"R2.値": "山田", "R3.C2": "東京都"}, str(tmp_path / "g.csv"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert "値（2行目）：山田" in r["say_to_user"]


def test_tsv_rejects_tab_in_value():
    from todoku.datafile.csvfile import CsvError, CsvFile, fill

    f = CsvFile(b"a\tb\n1\t2\n", "t.tsv", "tsv")
    assert f.rows == [["a", "b"], ["1", "2"]]
    with pytest.raises(CsvError):
        fill(f, {"R2.C1": "x\ty"})


# ====================================================================== JSON・JSON Lines
JSON_SRC = '{\n  "title": "申請書",\n  "applicant": {\n    "name": null,\n    "age": 0,\n    "agree": false\n  },\n  "items": [\n    {"no": 1, "what": ""}\n  ]\n}\n'


def test_json_fill_types_insert_and_proof():
    from todoku.datafile.jsonfile import JsonFile, fill, prove

    blob = JSON_SRC.encode()
    res = fill(JsonFile(blob, "a.json"), {"/applicant/name": "山田", "/applicant/age": "34", "/applicant/agree": True,
                                         "/applicant/tel": "03", "/items/-": {"no": 2, "what": "b"}})
    out = json.loads(res.blob)
    assert out["applicant"] == {"name": "山田", "age": 34, "agree": True, "tel": "03"}
    assert out["items"][1] == {"no": 2, "what": "b"}
    assert '    "tel": "03"\n' in res.blob.decode()  # 足した名前は、ほかの名前と同じ字下げ
    pr = prove(blob, res.blob, {p.id: p.value for p in res.patches}, "a.json")
    assert pr.ok, pr.unexpected


def test_json_keeps_escape_style_and_jsonl():
    from todoku.datafile.audit import run
    from todoku.datafile.jsonfile import JsonFile, fill

    blob = json.dumps({"a": "日本"}).encode()  # 非 ASCII を \\u でエスケープしたファイル
    assert fill(JsonFile(blob, "e.json"), {"/a": "東京"}).blob == json.dumps({"a": "東京"}).encode()
    lines = b'{"id": 1, "name": "a"}\n{"id": 2, "name": ""}\n'
    res = fill(JsonFile(lines, "c.jsonl", "jsonl"), {"L2/name": "b"})
    assert res.blob == b'{"id": 1, "name": "a"}\n{"id": 2, "name": "b"}\n'
    assert run(lines, "c.jsonl", "jsonl").ok
    assert run(JSON_SRC.encode(), "a.json", "json").ok


def test_json_duplicate_names_are_c():
    assert grade_blob(b'{"a": 1, "a": 2}', "d.json").grade == "C"
    assert grade_blob(b'{"a": ', "d.json").grade == "F"


# ====================================================================== OpenDocument の文書
NS = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
      'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
      'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" '
      'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" '
      'xmlns:xlink="http://www.w3.org/1999/xlink" '
      'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:version="1.3"')


def _package(mime: str, body: str, auto: str = "") -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), mime)
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
                   f'<manifest:file-entry manifest:full-path="/" manifest:media-type="{mime}"/>'
                   '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>', zipfile.ZIP_DEFLATED)
        z.writestr("content.xml", f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {NS}><office:automatic-styles>{auto}'
                   f'</office:automatic-styles><office:body>{body}</office:body></office:document-content>', zipfile.ZIP_DEFLATED)
    return buf.getvalue()


def make_odt(body: str, auto: str = "", forms: str = "") -> bytes:
    return _package("application/vnd.oasis.opendocument.text", f"<office:text>{forms}{body}</office:text>", auto)


ODT_AUTO = '<style:style style:name="T1" style:family="text"><style:text-properties style:text-underline-style="solid"/></style:style>'
ODT_BODY = ('<text:p>申請書</text:p><text:p>令和　　年　　月　　日</text:p>'
            '<table:table table:name="t1"><table:table-column table:number-columns-repeated="2"/>'
            '<table:table-row><table:table-cell><text:p>氏名</text:p></table:table-cell><table:table-cell><text:p/></table:table-cell></table:table-row>'
            '</table:table>'
            '<text:p>□新規　　□変更</text:p>'
            '<text:p>代表者 <text:s text:c="5"/>印</text:p>'
            '<text:p>氏名：<text:placeholder text:placeholder-type="text" text:description="氏名">&lt;氏名&gt;</text:placeholder></text:p>'
            '<text:p>同意 <draw:control draw:control="c1"/>する</text:p>'
            '<text:section text:name="s" text:protected="true"><text:p>受付番号：</text:p></text:section>')
ODT_FORMS = ('<office:forms><form:form form:name="Standard"><form:checkbox form:name="同意" xml:id="c1" form:id="c1" '
             'form:current-state="unchecked"/></form:form></office:forms>')


def test_odt_window_fill_proof_and_stable_ids():
    from todoku.odf.text import OdtDocument
    from todoku.odf.textfill import fill, written_of
    from todoku.odf.textproof import prove
    from todoku.odf.textwindow import build

    blob = make_odt(ODT_BODY, ODT_AUTO, ODT_FORMS)
    win = build(OdtDocument(blob, "a.odt"))
    lines = [ln for st in win["stories"] for ln in st["reading"]]
    # 「令和　年　月　日」の空きは字の並びからの推定（？）。マスの見出し「氏名」の後ろも、右のマスの代わりの所として示す（消さない）
    assert "令和［P2.c2？　　］年［P2.c3？　　］月［P2.c4？　　］日" in lines and "氏名［P3.c2？］ ｜ ［T1.R1.C2］" in lines, lines
    by = {f["id"]: f for f in win["fields"]}
    assert by["同意"]["label"] == "する" and by["同意"]["kind"] == "checkbox"
    assert by["P9.c5"]["readonly"] and by["P9.c5"]["protected"]  # 保護された節の中
    vals = {"P2.c2": "8", "T1.R1.C2": "山田 太郎", "P5.c0": True, "氏名": "山田", "同意": True, "P6.c3": "代表 一郎"}
    res = fill(OdtDocument(blob, "a.odt"), vals)
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected
    after = build(OdtDocument(res.blob, "a.odt"))
    got = {f["id"]: f.get("value") for f in after["fields"]}
    assert got["P2.c2"] == "8" and got["T1.R1.C2"] == "山田 太郎" and got["氏名"] == "山田" and got["同意"] is True
    assert got["P6.c3"] == "代表 一郎" and got["P5.c0"] is True
    with pytest.raises(ValueError):
        fill(OdtDocument(blob, "a.odt"), {"P9.c5": "1"})  # 保護された節には書かない
    content = zipfile.ZipFile(io.BytesIO(res.blob)).read("content.xml").decode()
    assert "山田<text:s/>太郎" in content  # 値の中の空白は text:s（ODF の空白の決まりで消えない）


def test_odt_anchored_write_and_audit(tmp_path):
    from todoku.odf.textaudit import run

    blob = make_odt(ODT_BODY, ODT_AUTO, ODT_FORMS)
    au = run(blob, "a.odt")
    assert au.ok, [(c.name, c.items) for c in au.checks if not c.ok]
    src = tmp_path / "a.odt"
    src.write_bytes(blob)
    r = api.fill(str(src), {"@1": {"old": "申請書", "new": "申請書（新規）"}, "同意": True}, str(tmp_path / "b.odt"))
    assert r["proof_ok"] and r["read_back_ok"] and r["anchored"] == {"@1": "P1.w1"}, r
    g = api.grade(str(src))
    assert g["grade"] == "A" and g["write"]["grade"] == "B"  # 規則では決められない所は、書く段階の B


def test_odt_whitespace_rules():
    from todoku.odf.text import OdtDocument

    doc = OdtDocument(make_odt("<text:p>  a  b<text:s text:c=\"2\"/>c <text:tab/> d  </text:p>"), "w.odt")
    assert doc.body[0].text == "a b  c \t d"


def test_flat_odt():
    from todoku.odf.text import OdtDocument

    flat = (f'<?xml version="1.0" encoding="UTF-8"?><office:document {NS} office:mimetype="application/vnd.oasis.opendocument.text">'
            '<office:body><office:text><text:p>氏名：</text:p></office:text></office:body></office:document>').encode()
    assert kind_of(flat) == "fodt"
    assert OdtDocument(flat, "f.fodt").body[0].text == "氏名："


# ====================================================================== OpenDocument の表計算
ODS_AUTO = ('<number:number-style style:name="N3"><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true"/></number:number-style>'
            '<number:date-style style:name="N40"><number:year number:style="long"/><number:text>/</number:text><number:month/><number:text>/</number:text><number:day/></number:date-style>'
            '<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:border="0.06pt solid #000000"/></style:style>'
            '<style:style style:name="ce2" style:family="table-cell" style:data-style-name="N3"><style:table-cell-properties fo:border="0.06pt solid #000000"/></style:style>'
            '<style:style style:name="ce3" style:family="table-cell" style:data-style-name="N40"><style:table-cell-properties fo:border="0.06pt solid #000000"/></style:style>')
ODS_VAL = ('<table:content-validations><table:content-validation table:name="val1" table:condition="of:cell-content-is-in-list(&quot;男&quot;;&quot;女&quot;)" '
           'table:allow-empty-cell="true"><table:error-message table:message-type="stop" table:display="true"/></table:content-validation></table:content-validations>')
SHEET = ('<table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="4"/>'
         '<table:table-row><table:table-cell office:value-type="string" calcext:value-type="string"><text:p>氏名</text:p></table:table-cell>'
         '<table:table-cell table:style-name="ce1" table:number-columns-spanned="2"/><table:covered-table-cell table:style-name="ce1"/><table:table-cell/></table:table-row>'
         '<table:table-row><table:table-cell office:value-type="string" calcext:value-type="string"><text:p>性別</text:p></table:table-cell>'
         '<table:table-cell table:style-name="ce1" table:content-validation-name="val1"/><table:table-cell table:number-columns-repeated="2"/></table:table-row>'
         '<table:table-row><table:table-cell office:value-type="string" calcext:value-type="string"><text:p>金額</text:p></table:table-cell>'
         '<table:table-cell table:style-name="ce2" table:number-columns-repeated="2"/><table:table-cell table:style-name="ce2" table:formula="of:=[.B3]+[.C3]" '
         'office:value-type="float" office:value="0" calcext:value-type="float"><text:p>0</text:p></table:table-cell></table:table-row>'
         '<table:table-row><table:table-cell office:value-type="string" calcext:value-type="string"><text:p>日付</text:p></table:table-cell>'
         '<table:table-cell table:style-name="ce3"/><table:table-cell table:number-columns-repeated="2"/></table:table-row>'
         '<table:table-row table:number-rows-repeated="3"><table:table-cell table:style-name="ce1" table:number-columns-repeated="4"/></table:table-row>'
         '<table:table-row table:number-rows-repeated="1048570"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>'
         '</table:table>')


def make_ods(tables: str, auto: str = ODS_AUTO, validations: str = ODS_VAL) -> bytes:
    return _package("application/vnd.oasis.opendocument.spreadsheet", f"<office:spreadsheet>{validations}{tables}</office:spreadsheet>", auto)


def test_ods_window_fill_recalc_and_proof():
    from todoku.formula.parser import parse_cell
    from todoku.odf.sheet import OdsWorkbook
    from todoku.odf.sheetfill import fill, prove
    from todoku.odf.sheetwindow import build, grade_sheets

    blob = make_ods(SHEET)
    doc = build(blob, "a.ods")
    fields = {f["id"]: f for f in doc["sheets"][0]["fields"]}
    assert fields["Sheet1!B2"]["options"] == ["男", "女"] and fields["Sheet1!B3"]["kind"] == "number"
    assert fields["Sheet1!B4"]["kind"] == "date" and "D3" in fields["Sheet1!B3"]["write_as"]
    cell_d3 = next(c for c in doc["sheets"][0]["cells"] if c["cell"] == "D3")
    assert cell_d3["formula"] == "of:=[.B3]+[.C3]" and cell_d3["formula_excel"] == "=B3+C3"
    # 枠・塗りのある空のマスは、見出しが近くにあってもなくても、ファイルが入力欄と宣言していない（入力規則・保護の解除がない）
    # ので、規則では欄と決められない所［B1？］［B5？］として窓に出す（書く段階の B：エージェントが判断する）。
    # 中身は規則で読めるので、読む段階は A
    g = grade_sheets(blob, "a.ods")
    titles = {r.title: r.grade for r in g.write_reasons}
    assert g.grade == "A" and g.write == "B" and titles.get("規則では決められない所") == "B", g.headline()
    assert fields["Sheet1!B1"]["agent"] and fields["Sheet1!B1"]["reason_codes"] == ["blank_box"]
    assert fields["Sheet1!B5"]["agent"] and fields["Sheet1!B5"]["evidence"]["borders"] == ["左", "右", "上", "下"]
    res = fill(OdsWorkbook(blob, "a.ods"), {"Sheet1!B1": "山田 太郎", "Sheet1!B2": "女", "Sheet1!B3": "1,000", "Sheet1!C3": 234,
                                             "Sheet1!B4": "2026/4/1", "Sheet1!C6": "x"})
    recalc = {(p.sheet, *parse_cell(p.address)): p.after for p in res.patches if p.kind == "formula"}
    written = {(p.sheet, *parse_cell(p.address)): p.after for p in res.patches if p.kind == "input"}
    assert recalc == {("Sheet1", 3, 4): 1234.0}
    pr = prove(blob, res.blob, written, recalc)
    assert pr.ok, pr.unexpected
    content = zipfile.ZipFile(io.BytesIO(res.blob)).read("content.xml").decode()
    assert 'office:value="1234" calcext:value-type="float"><text:p>1,234</text:p>' in content  # 計算し直した結果と見える字
    assert 'office:date-value="2026-04-01"><text:p>2026/4/1</text:p>' in content
    # 繰り返しの行（3行）を、書いた行の前・その行・後に分けた
    assert content.count('<table:table-row><table:table-cell table:style-name="ce1" table:number-columns-repeated="4"/></table:table-row>') == 2
    with pytest.raises(ValueError):
        fill(OdsWorkbook(blob, "a.ods"), {"Sheet1!B2": "その他"})  # 入力規則（停止）に合わない


def test_ods_audit_and_api(tmp_path):
    from todoku.odf.sheetaudit import run

    blob = make_ods(SHEET)
    au = run(blob, "a.ods")
    assert au.ok, [(c.name, c.items) for c in au.checks if not c.ok]
    src = tmp_path / "a.ods"
    src.write_bytes(blob)
    r = api.fill(str(src), {"Sheet1!B3": 5, "Sheet1!C3": 6}, str(tmp_path / "b.ods"))
    assert r["proof_ok"] and r["read_back_ok"] and r["recalculated"] == {"Sheet1!D3": 11.0}, r
    assert [f["id"] for f in api.fields(str(src))][:2] == ["Sheet1!B1", "Sheet1!B2"]


def test_openformula_translation():
    from todoku.odf.formula import OdfFormulaError, to_excel

    assert to_excel("of:=SUM([.A1:.B2];['シート 1'.$C$3])") == "SUM(A1:B2,'シート 1'!$C$3)"
    assert to_excel("of:={1;2|3;4}") == "{1,2;3,4}"
    assert to_excel('of:=IF(ISERROR([.A1]);"";#N/A)') == 'IF(ISERROR(A1),"",#N/A)'
    with pytest.raises(OdfFormulaError):
        to_excel("of:=SUM([.A1]~[.B1])")


# ====================================================================== RTF
def sj(s: str) -> str:
    return "".join(B + "'" + f"{x:02x}" for x in s.encode("cp932"))


def make_rtf(*paras: str) -> bytes:
    head = ("{" + B + "rtf1" + B + "ansi" + B + "ansicpg932" + B + "deff0{" + B + "fonttbl{" + B + "f0" + B + "fcharset0 Century;}{"
            + B + "f1" + B + "fcharset128 " + sj("ＭＳ 明朝") + ";}}")
    return ("\r\n".join([head, *paras, "}"])).encode("latin-1")


def _w(word: str) -> str:
    return B + word


RTF = make_rtf(
    _w("pard") + _w("f1 ") + sj("申請書") + _w("par"),
    _w("pard") + _w("f1 ") + sj("令和　　年　　月　　日") + _w("par"),
    _w("trowd") + _w("cellx2000") + _w("cellx6000") + _w("pard") + _w("intbl") + _w("f1 ") + sj("氏名") + _w("cell") + _w("pard")
    + _w("intbl") + _w("cell") + _w("row"),
    _w("pard") + _w("f1") + _w("uc1") + _w("u12354") + "?" + _w("u12356") + "?ABC" + _w("par"),
    _w("pard") + _w("f1 ") + sj("電話") + "{" + _w("field") + "{" + _w("*") + _w("fldinst {FORMTEXT {") + _w("*") + _w("formfield{")
    + _w("fftype0{") + _w("*") + _w("ffname Tel}}}}}{") + _w("fldrslt {") + _w("u8194") + _w("'20") + _w("u8194") + _w("'20}}}") + _w("par"),
    _w("pard") + _w("f1 ") + "{" + _w("field") + "{" + _w("*") + _w("fldinst {FORMCHECKBOX {") + _w("*") + _w("formfield{")
    + _w("fftype1") + _w("ffres25") + _w("ffdefres0{") + _w("*") + _w("ffname Agree}}}}}{") + _w("fldrslt }}") + sj("同意") + _w("par"),
    _w("pard") + _w("f1 ") + sj("□新規　　□変更") + _w("par"),
    "{" + _w("header ") + _w("pard") + _w("f1 ") + sj("表紙") + _w("par}"),
)


def test_rtf_window_fill_and_proof():
    from todoku.rtf.document import RtfDocument
    from todoku.rtf.fill import fill, prove, written_of
    from todoku.rtf.window import build

    win = build(RtfDocument(RTF, "a.rtf"))
    lines = [ln for st in win["stories"] for ln in st["reading"]]
    # 「令和　年　月　日」の空きは字の並びからの推定（？）。マスの見出し「氏名」の後ろも、右のマスの代わりの所として示す（消さない）
    assert "令和［P2.c2？　　］年［P2.c3？　　］月［P2.c4？　　］日" in lines and "氏名［P3.c2？］ ｜ ［T1.R1.C2］" in lines, lines
    by = {f["id"]: f for f in win["fields"]}
    assert by["Tel"]["value"] == "" and by["Agree"]["kind"] == "checkbox" and by["Agree"]["value"] is False
    vals = {"P2.c2": "8", "T1.R1.C2": "山田 太郎", "Tel": "03-1234", "Agree": True, "P8.c0": True}
    res = fill(RtfDocument(RTF, "a.rtf"), vals)
    pr = prove(RTF, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected
    got = {f["id"]: f.get("value") for f in build(RtfDocument(res.blob, "a.rtf"))["fields"]}
    assert got["P2.c2"] == "8" and got["T1.R1.C2"] == "山田 太郎" and got["Tel"] == "03-1234" and got["Agree"] is True and got["P8.c0"] is True
    assert (B + "u23665?" + B + "u30000?").encode() in res.blob  # 「山田」：\\uN と、\\uc1 の数だけの代わりの字
    assert (B + "ffres1").encode() in res.blob


def test_rtf_no_control_word_merging():
    from todoku.rtf.lexer import control_word_ends_at

    assert control_word_ends_at(b"x" + (B + "loch").encode(), 6)
    assert not control_word_ends_at((B + "loch ").encode(), 6)
    assert not control_word_ends_at((B + B + "b").encode(), 3)


def test_rtf_audit_and_api(tmp_path):
    from todoku.rtf.audit import run

    au = run(RTF, "a.rtf")
    assert au.ok, [(c.name, c.items) for c in au.checks if not c.ok]
    src = tmp_path / "a.rtf"
    src.write_bytes(RTF)
    r = api.fill(str(src), {"T1.R1.C2": "山田", "@1": {"old": "申請書", "new": "申請書（新規）"}}, str(tmp_path / "b.rtf"))
    assert r["proof_ok"] and r["read_back_ok"], r
    g = api.grade(str(src))
    assert g["grade"] == "A" and g["write"]["grade"] == "B"
    assert api.summary(str(src))["kind_ja"] == "RTF（リッチテキスト）"


# ====================================================================== 公開の様式（corpus があるときだけ）
def _corpus(sub: str, exts: tuple[str, ...]) -> list[Path]:
    base = CORPUS / sub
    if not base.exists():
        return []
    return sorted(p for p in base.rglob("*") if p.is_file() and p.suffix.lower() in exts)


@pytest.mark.corpus
@pytest.mark.parametrize("path", _corpus("wide_csv", (".csv", ".tsv", ".json", ".jsonl")), ids=lambda p: p.name)
def test_corpus_data(path):
    from todoku.openkinds import audit_runner

    blob = path.read_bytes()
    au = audit_runner(kind_of(blob, path.name))(blob, path.name)
    assert au.ok, [(c.name, c.items[:5]) for c in au.checks if not c.ok]
    g = grade_blob(blob, path.name)
    # データそのもの（S）。ただし、いくつもの文字コードで仕様どおりに読めるもの（Shift_JIS と GBK の両方で読めるものなど）は、
    # 規則では文字コードを決められないので B（中身は窓にすべて出ていて、窓に候補を示し、エージェントが決める）
    if g.grade == "B":
        assert [r.title for r in g.reasons if r.grade == "B"] == ["文字コードを規則では決められない"], g.reasons
    else:
        assert g.grade == "S"


@pytest.mark.corpus
@pytest.mark.parametrize("path", _corpus("wide_odf", (".odt", ".ods", ".ott", ".ots")), ids=lambda p: p.name)
def test_corpus_odf(path):
    from todoku.openkinds import audit_runner

    blob = path.read_bytes()
    au = audit_runner(kind_of(blob, path.name))(blob, path.name)
    assert au.ok, [(c.name, c.items[:5]) for c in au.checks if not c.ok]


@pytest.mark.corpus
@pytest.mark.parametrize("path", _corpus("wide_rtf", (".rtf",)), ids=lambda p: p.name)
def test_corpus_rtf(path):
    from todoku.rtf.audit import run

    au = run(path.read_bytes(), path.name)
    assert au.ok, [(c.name, c.items[:5]) for c in au.checks if not c.ok]
