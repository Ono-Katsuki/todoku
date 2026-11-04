"""第5回の見直しで見つかった、Office・OpenDocument・HWPX の書き込みの誤り・黙った失いの確かめ（合成の文書）。

- Excel：テーブルの見出し・集計の行の見出しのセルは、ふつうは書かない。force なら列の名前と構造化参照を合わせて書き換え、証明する
- Office・OpenDocument：パッケージの電子署名は、書くと無効になることを伝える（警告・本人に伝えること・書く段階・確かめ）
- Word：テキストボックスの代わりの中身（mc:Fallback）にも同じ字を写す。写せなければ取り除いて伝える
- Word：変更履歴の中には書かない（挿入は分けて外に置く）。しおりの番号を変更履歴の番号と重ねない。記録がオンでも変更履歴にしない
- Excel：JSON の数は数として、字は字として書く（12 桁の番号）
- Excel：選択肢の範囲が空の入力規則（リスト）は、どの値も受け付けないので断る
- HWPX：누름틀の中身のタブなどは、書き換えるときに一緒に除いてよい
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from todoku import api

X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------- Excel
def _xlsx(sheet_rows: str, *, table: str | None = None, validations: str = "", sst: list[str] | None = None) -> bytes:
    sst = sst or []
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          + ('<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' if table else "")
          + '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = (f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{X}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>')
    wbr = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
    ss = f'<?xml version="1.0" encoding="UTF-8"?><sst xmlns="{X}">' + "".join(f"<si><t>{t}</t></si>" for t in sst) + "</sst>"
    sh = (f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{X}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          f"<sheetData>{sheet_rows}</sheetData>{validations}"
          + ('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' if table else "") + "</worksheet>")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wbr)
        z.writestr("xl/sharedStrings.xml", ss)
        z.writestr("xl/worksheets/sheet1.xml", sh)
        if table:
            z.writestr("xl/worksheets/_rels/sheet1.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>')
            z.writestr("xl/tables/table1.xml", table)
    return buf.getvalue()


TABLE_ROWS = ('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
              '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>3</v></c><c r="C2"><f>T1[[#This Row],[金額]]*2</f><v>6</v></c></row>'
              '<row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3"><v>5</v></c><c r="C3"><f>T1[[#This Row],[金額]]*2</f><v>10</v></c></row>'
              '<row r="4"><c r="A4" t="s"><v>4</v></c><c r="B4"><f>SUBTOTAL(109,T1[金額])</f><v>8</v></c>'
              '<c r="D4"><f>SUM(T1[[#Totals],[金額]])*3</f><v>24</v></c></row>')
TABLE = (f'<?xml version="1.0" encoding="UTF-8"?><table xmlns="{X}" id="1" name="T1" displayName="T1" ref="A1:B4" totalsRowCount="1">'
         '<autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="品目" totalsRowLabel="合計"/>'
         '<tableColumn id="2" name="金額" totalsRowFunction="sum"/></tableColumns><tableStyleInfo name="TableStyleMedium2"/></table>')
TABLE_SST = ["品目", "金額", "a", "b", "合計"]


@pytest.fixture
def tbl(tmp_path):
    p = tmp_path / "tbl.xlsx"
    p.write_bytes(_xlsx(TABLE_ROWS, table=TABLE, sst=TABLE_SST))
    return p


@pytest.mark.parametrize("cell", ["S!B1", "S!A4"])
def test_table_label_cells_refused(tbl, tmp_path, cell):
    with pytest.raises(api.ApiError) as e:
        api.fill(str(tbl), {cell: {"value": "新しい名前", "overwrite_text": True}}, str(tmp_path / "o.xlsx"))
    assert "書き換えると表の定義とずれる" in str(e.value)


def test_table_header_force_renames_column_and_refs(tbl, tmp_path):
    r = api.fill(str(tbl), {"S!B1": {"value": "金 額", "overwrite_text": True}}, str(tmp_path / "o.xlsx"), force=True)
    assert r["proof_ok"] and r["read_back_ok"], r["unexpected"]
    z = zipfile.ZipFile(r["out_path"])
    t = z.read("xl/tables/table1.xml").decode()
    assert 'name="金 額"' in t and 'name="金額"' not in t
    sh = z.read("xl/worksheets/sheet1.xml").decode()
    assert "T1[[#This Row],[金 額]]*2" in sh and "SUBTOTAL(109,T1[[金 額]])" in sh and "T1[[#Totals],[金 額]]" in sh
    assert "[金額]" not in sh
    assert any("列の名前" in w for w in r["warnings"])


def test_table_totals_label_force(tbl, tmp_path):
    r = api.fill(str(tbl), {"S!A4": {"value": "総計", "overwrite_text": True}}, str(tmp_path / "o.xlsx"), force=True)
    assert r["proof_ok"], r["unexpected"]
    assert 'totalsRowLabel="総計"' in zipfile.ZipFile(r["out_path"]).read("xl/tables/table1.xml").decode()


def test_table_duplicate_name_refused(tbl, tmp_path):
    with pytest.raises(api.ApiError) as e:
        api.fill(str(tbl), {"S!B1": {"value": "品目", "overwrite_text": True}}, str(tmp_path / "o.xlsx"), force=True)
    assert "重なります" in str(e.value)


def test_proof_catches_stale_table_column(tbl):
    """見出しのセルだけを書き換え、列の名前を残した書き込み（前の誤り）は、証明が止める。"""
    from todoku.xlsx.proof import prove

    blob = tbl.read_bytes()
    src = zipfile.ZipFile(io.BytesIO(blob))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            d = src.read(n)
            if n == "xl/worksheets/sheet1.xml":
                d = d.replace(b'<c r="B1" t="s"><v>1</v></c>', '<c r="B1" t="inlineStr"><is><t>額</t></is></c>'.encode())
            z.writestr(n, d)
    pr = prove(blob, buf.getvalue(), {("S", "B1")})
    assert not pr.ok and any("列の名前" in u for u in pr.unexpected)


def test_twelve_digits_number_and_text(tmp_path):
    p = tmp_path / "n.xlsx"
    p.write_bytes(_xlsx('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', sst=["番号"]))
    r = api.fill(str(p), {"S!B1": 123456789012, "S!B2": "123456789012"}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and r["read_back_ok"]
    sh = zipfile.ZipFile(r["out_path"]).read("xl/worksheets/sheet1.xml").decode()
    assert '<c r="B1"><v>123456789012</v></c>' in sh
    assert '<c r="B2" t="inlineStr"><is><t>123456789012</t></is></c>' in sh
    w = " ".join(r["warnings"])
    assert "1.23457E+11" in w and "字（\"123456789012\"）で指定" in w
    assert "数として書くなら 123456789012 と指定" not in w


def test_list_validation_empty_range_refused(tmp_path):
    p = tmp_path / "v.xlsx"
    dv = ('<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B1">'
          "<formula1>$H$1:$H$2</formula1></dataValidation></dataValidations>")
    p.write_bytes(_xlsx('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', validations=dv, sst=["区分"]))
    with pytest.raises(api.ApiError) as e:
        api.fill(str(p), {"S!B1": "x"}, str(tmp_path / "o.xlsx"))
    assert "空で、選べる値がありません" in str(e.value) and "連動する欄" not in str(e.value)


# ---------------------------------------------------------------- Word
def _docx(body: str, *, settings: str = "", extra: dict | None = None) -> bytes:
    ns = (f'xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        rels = ('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                + (extra or {}).pop("_pkgrel", ""))
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {ns}><w:body>{body}<w:sectPr/></w:body></w:document>')
        z.writestr("word/settings.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="{W}">{settings}</w:settings>')
        for k, v in (extra or {}).items():
            z.writestr(k, v)
    return buf.getvalue()


TXBX_PARA = '<w:p><w:r><w:t xml:space="preserve">代表者名：　　　　　　</w:t></w:r></w:p>'


def _txbx(fallback_para: str = TXBX_PARA) -> str:
    return ('<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" '
            'simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="page"><wp:posOffset>3000000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page">'
            '<wp:posOffset>1000000</wp:posOffset></wp:positionV><wp:extent cx="2000000" cy="500000"/><wp:docPr id="1" name="TB1"/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr>'
            '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx>'
            f'<w:txbxContent>{TXBX_PARA}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor>'
            '</w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape style="position:absolute;width:150pt;height:40pt"><v:textbox>'
            f'<w:txbxContent>{fallback_para}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>')


def _box_id(path) -> str:
    return next(f["id"] for f in api.fields(str(path)) if "代表者名" in (f.get("label") or ""))


def test_textbox_fallback_mirrored(tmp_path):
    p = tmp_path / "tb.docx"
    p.write_bytes(_docx(_txbx()))
    r = api.fill(str(p), {_box_id(p): "山田"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r["unexpected"]
    x = zipfile.ZipFile(r["out_path"]).read("word/document.xml").decode()
    fb = re.search(r"<mc:Fallback>.*</mc:Fallback>", x, re.S)[0]
    assert "山田" in fb and "_todoku_" not in fb


def test_textbox_fallback_dropped_when_unmatched(tmp_path):
    p = tmp_path / "tb2.docx"
    p.write_bytes(_docx(_txbx('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">代表者名：　　　　　　</w:t></w:r></w:p>')))
    r = api.fill(str(p), {_box_id(p): "山田"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r["unexpected"]
    x = zipfile.ZipFile(r["out_path"]).read("word/document.xml").decode()
    assert "mc:Fallback" not in x and "mc:Choice" in x
    assert any("Fallback" in w for w in r["warnings"])


def test_proof_catches_blank_fallback(tmp_path):
    """Choice だけに書いて Fallback を空のままにした書き込み（前の誤り）は、証明が止める。"""
    from todoku.docx.document import DocxDocument
    from todoku.docx.proof import prove
    from todoku.docx.window import Spots

    blob = _docx(_txbx())
    doc = zipfile.ZipFile(io.BytesIO(blob)).read("word/document.xml").decode()
    old = "代表者名：　　　　　　</w:t></w:r></w:p></w:txbxContent></wps:txbx>"
    assert old in doc
    bad = doc.replace(old, '代表者名：</w:t></w:r><w:bookmarkStart w:id="9" w:name="_todoku_X"/><w:r><w:t>山田</w:t></w:r>'
                           '<w:bookmarkEnd w:id="9"/></w:p></w:txbxContent></wps:txbx>')
    src = zipfile.ZipFile(io.BytesIO(blob))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            z.writestr(n, bad if n == "word/document.xml" else src.read(n))
    fid = next(k for k, it in Spots(DocxDocument(blob)).items.items() if it.get("para") is not None and "代表者名" in it["para"].text)
    pr = prove(blob, buf.getvalue(), {fid: "山田"})
    assert not pr.ok


REV_TABLE = ('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>氏名</w:t>'
             '</w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>')
REV = ('<w:p><w:r><w:t>連絡先：</w:t></w:r><w:del w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>旧番号</w:delText>'
       '</w:r></w:del><w:ins w:id="2" w:author="x" w:date="2026-01-01T00:00:00Z"><w:r><w:t>（　　　）</w:t></w:r></w:ins></w:p>')


def test_write_inside_insertion_is_placed_outside(tmp_path):
    p = tmp_path / "rev.docx"
    p.write_bytes(_docx(REV_TABLE + REV))
    fid = next(f["id"] for f in api.fields(str(p)) if "連絡先" in (f.get("label") or ""))
    r = api.fill(str(p), {fid: "03-1234"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r["unexpected"]
    x = zipfile.ZipFile(r["out_path"]).read("word/document.xml").decode()
    ins = re.findall(r"<w:ins\b.*?</w:ins>", x, re.S)
    assert all("03-1234" not in i for i in ins) and "03-1234" in x
    ids = re.findall(r'w:id="(\d+)"', x)
    bm = re.search(r'<w:bookmarkStart w:id="(\d+)" w:name="_todoku_', x)[1]
    assert ids.count(bm) == 2  # しおりの始まりと終わりだけ（変更履歴の番号と重ならない）
    w = " ".join(r["warnings"])
    assert "変更履歴" in w and "挿入 1" in w and "削除 1" in w


def test_write_inside_nested_revision_refused(tmp_path):
    from todoku.docx.document import DocxDocument
    from todoku.docx.fill import FillError, Filler

    body = REV_TABLE + ('<w:p><w:r><w:t>連絡先：</w:t></w:r><w:ins w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z"><w:hyperlink>'
                        '<w:r><w:t>（　　　）</w:t></w:r></w:hyperlink></w:ins></w:p>')
    doc = DocxDocument(_docx(body))
    f = Filler(doc)
    para = next(p for _k, _part, bl in doc.all_blocks() for p in __import__("todoku.docx.document", fromlist=["iter_paras"]).iter_paras(bl)
                if "連絡先" in p.text)
    a = para.text.index("（") + 1
    f.replace_span({"id": "X"}, para, a, a + 3, "03")
    with pytest.raises(FillError) as e:
        f.finish_runs()
    assert "変更履歴" in str(e.value)


def test_track_revisions_on_is_told(tmp_path):
    p = tmp_path / "tr.docx"
    p.write_bytes(_docx(REV_TABLE, settings="<w:trackRevisions/>"))
    r = api.fill(str(p), {"T1.R1.C2": "山田"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"]
    assert any("変更履歴の記録" in w for w in r["warnings"])


# ---------------------------------------------------------------- 電子署名
def _signed_docx() -> bytes:
    rel = ('<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" '
           'Target="_xmlsignatures/origin.sigs"/>')
    return _docx(REV_TABLE, extra={"_pkgrel": rel, "_xmlsignatures/origin.sigs": "",
                                   "_xmlsignatures/sig1.xml": '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>'})


def test_signature_warned_graded_and_audited(tmp_path):
    p = tmp_path / "s.docx"
    p.write_bytes(_signed_docx())
    r = api.fill(str(p), {"T1.R1.C2": "山田"}, str(tmp_path / "o.docx"))
    assert r["warnings"][0].startswith("この書類には電子署名があります。書き込むと署名は無効になり、もう一度署名が要ります")
    assert "電子署名" in r["say_to_user"] and r["digital_signature"]["invalidated"]
    g = api.grade(str(p))
    assert g["write"]["grade"] in ("C", "D") and any(x["title"] == "電子署名" for x in g["write"]["reasons"])
    au = api.audit(str(p))
    assert au["digital_signature"]["kind"] == "opc"


def test_odf_signature_found():
    from todoku import pkgsig

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.text")
        z.writestr("META-INF/manifest.xml", "<manifest/>")
        z.writestr("META-INF/documentsignatures.xml", '<document-signatures><Signature xmlns="x"/></document-signatures>')
    s = pkgsig.find(buf.getvalue())
    assert s is not None and s.kind == "odf" and s.count == 1
    # マクロだけの署名は数えない
    buf2 = io.BytesIO()
    with zipfile.ZipFile(buf2, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.text")
        z.writestr("META-INF/macrosignatures.xml", "<x/>")
    assert pkgsig.find(buf2.getvalue()) is None


# ---------------------------------------------------------------- HWPX
def test_hwpx_click_here_with_tab_replaced(tmp_path):
    import test_hwpx as T

    cmd = "Clickhere:set:66:Direction:wstring:2:입력 HelpState:wstring:0: "
    begin = ('<hp:ctrl><hp:fieldBegin id="22" type="CLICK_HERE" name="주소" editable="1" dirty="1" zorder="-1" fieldid="627272811">'
             f'<hp:parameters cnt="2" name=""><hp:integerParam name="Prop">9</hp:integerParam><hp:stringParam name="Command" '
             f'xml:space="preserve">{cmd}</hp:stringParam></hp:parameters></hp:fieldBegin></hp:ctrl>')
    end = '<hp:ctrl><hp:fieldEnd beginIDRef="22" fieldid="627272811"/></hp:ctrl>'
    runs = T.run_(begin + T.t("서울시 "), "0") + T.run_(T.t("<hp:tab/>"), "1") + T.run_(T.t("중구") + end + "<hp:t/>", "0")
    sec = (T.XML + f"<hs:sec {T.NS}>" + T.para(T.run_('<hp:secPr id="" textDirection="HORIZONTAL" outlineShapeIDRef="1"/>' + T.t("신청서")))
           + T.para(T.run_(T.t("주소: ")) + runs) + "</hs:sec>")
    p = tmp_path / "cs.hwpx"
    p.write_bytes(T.make_hwpx(sec))
    r = api.fill(str(p), {"주소": "부산시"}, str(tmp_path / "o.hwpx"))
    assert r["proof_ok"] and r["read_back_ok"], r["unexpected"]
    x = zipfile.ZipFile(r["out_path"]).read("Contents/section0.xml").decode()
    assert "부산시" in x and "중구" not in x and "<hp:tab/>" not in x


# ---------------------------------------------------------------- 第5回の実際の書類での見直し（Word・.doc・Excel の段落とマス）
def _para_doc(*texts: str) -> bytes:
    return _docx("".join(f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>' for t in texts))


def test_long_gap_keeps_space_after_heading(tmp_path):
    p = tmp_path / "g.docx"
    p.write_bytes(_para_doc("氏　名　　　　　　　　", "通院先：　　　　　　　　　通院頻度："))
    ids = {f.get("label_key", f["label"]): f["id"] for f in api.fields(str(p))}  # 見出しは字のまま、鍵は字間を詰めた写し
    name_id = next(i for l_, i in ids.items() if l_.startswith("氏名［"))
    clinic_id = next(i for l_, i in ids.items() if l_.startswith("通院先：［"))
    r = api.fill(str(p), {name_id: "山田　花子", clinic_id: "架空医院"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r["unexpected"]
    from todoku.docx.document import DocxDocument, iter_paras

    doc = DocxDocument(open(r["out_path"], "rb").read())
    texts = [q.text for _k, _p, bl in doc.all_blocks() for q in iter_paras(bl)]
    assert any(t.startswith("氏　名　山田　花子") for t in texts), texts
    assert any(t.startswith("通院先：　架空医院") and t.endswith("　通院頻度：") for t in texts), texts
    assert any("くっつかない" in w for w in r["warnings"])


def test_spaced_two_char_heading_offers_place_after():
    from todoku.docx.document import DocxDocument
    from todoku.docx.window import Spots

    sp = Spots(DocxDocument(_para_doc("住　所", "電　話")))
    after = [it for it in sp.items.values() if it.get("why_key") == "after_label"]
    assert len(after) == 2 and all(it["span"][0] == 3 for it in after)


def test_word_circle_is_resolved_on_the_original_text(tmp_path):
    """○の所は、書く前の字で決める：同じ書き込みで「5月」と書いても、occurrence 1 の「月」は元の文書の 1 つ目の月のまま。"""
    p = tmp_path / "c.docx"
    p.write_bytes(_para_doc("受診日：　　　　　　", "頻度：回／週・月"))
    gap = next(f["id"] for f in api.fields(str(p)) if (f.get("label") or "").startswith("受診日"))
    alone = api.fill(str(p), {"@c": {"circle": "月", "occurrence": 1}}, str(tmp_path / "a.docx"))
    both = api.fill(str(p), {gap: "5月1日", "@c": {"circle": "月", "occurrence": 1}}, str(tmp_path / "b.docx"))
    assert alone["proof_ok"] and both["proof_ok"], both["unexpected"]
    assert alone["circled"][0]["around"] == both["circled"][0]["around"] and "週・【月】" in alone["circled"][0]["around"]


def test_word_border_is_placed_in_schema_order():
    from todoku.docx import circle as dc
    from todoku.docx.document import DocxDocument

    doc = DocxDocument(_docx('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:lang w:eastAsia="ja-JP"/></w:rPr>'
                             '<w:t>都道府県</w:t></w:r></w:p>'))
    sd = dc.resolve(doc, {"circle": "都", "style": "border", "occurrence": 1}, {})
    out = dc.apply(doc, [sd])
    x = zipfile.ZipFile(io.BytesIO(out)).read("word/document.xml").decode()
    assert re.search(r'<w:rFonts w:hint="eastAsia"/><w:bdr [^>]*/><w:lang ', x), x
