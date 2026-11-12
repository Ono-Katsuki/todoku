"""Word（.docx）の監査で見つかったことの確かめ（合成の文書）：プレースホルダーの書式・差し込みのフィールド・箇条書きの □・
図形の字・日付の欄の読み戻し・見出し。"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku import api
from todoku.docx.document import DocxDocument
from todoku.docx.fill import FillError, fill
from todoku.docx.proof import prove
from todoku.docx.window import build

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"')

STYLES = ('<w:style w:type="character" w:styleId="a3"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="List"><w:name w:val="List Check"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>')

NUMBERING = ('<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="□"/>'
             '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/></w:rPr></w:lvl></w:abstractNum>'
             '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A8;"/>'
             '<w:lvlJc w:val="left"/><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
             '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="■"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
             '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>')


def docx(body: str, *, numbering: bool = False) -> bytes:
    buf = io.BytesIO()
    over = '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    rels = '<Relationship Id="rS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    if numbering:
        rels += '<Relationship Id="rN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   f'<Default Extension="xml" ContentType="application/xml"/>{over}</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships '
                   f'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}</w:body></w:document>')
        z.writestr("word/styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:styles {NS}>{STYLES}</w:styles>')
        if numbering:
            z.writestr("word/numbering.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:numbering {NS}>{NUMBERING}</w:numbering>')
    return buf.getvalue()


def p(*runs: str, ppr: str = "") -> str:
    return f"<w:p>{ppr}{''.join(runs)}</w:p>"


def r(t: str, rpr: str = "") -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>'


def write(tmp_path, blob: bytes, name: str = "a.docx") -> str:
    path = tmp_path / name
    path.write_bytes(blob)
    return str(path)


# ---------------------------------------------------------------------- 1. プレースホルダーの書式
PLACEHOLDER = p(r("氏名："), '<w:sdt><w:sdtPr><w:alias w:val="name"/><w:id w:val="1"/><w:placeholder><w:docPart w:val="X"/></w:placeholder>'
                '<w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="a3"/><w:sz w:val="21"/></w:rPr>'
                '<w:t>ここをクリックして入力</w:t></w:r></w:sdtContent></w:sdt>')


def test_placeholder_style_removal_is_part_of_the_fill():
    blob = docx(PLACEHOLDER)
    res = fill(DocxDocument(blob), {"name": "山田 花子"})
    pr = prove(blob, res.blob, {"name": "山田 花子"})
    assert pr.ok, pr.unexpected
    xml = zipfile.ZipFile(io.BytesIO(res.blob)).read("word/document.xml")
    assert b"showingPlcHdr" not in xml and b'w:rStyle w:val="a3"' not in xml and b'<w:sz w:val="21"/>' in xml


def test_other_new_formatting_is_still_caught():
    """プレースホルダーのスタイルを外すほかの書式の変更は、これまでどおり想定外とする。"""
    blob = docx(PLACEHOLDER)
    res = fill(DocxDocument(blob), {"name": "山田 花子"})
    bad = res.blob.replace(b"", b"")
    z = zipfile.ZipFile(io.BytesIO(bad))
    xml = z.read("word/document.xml").replace(b'<w:sz w:val="21"/>', b'<w:sz w:val="40"/>')
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z2:
        for n in z.namelist():
            z2.writestr(n, xml if n == "word/document.xml" else z.read(n))
    assert not prove(blob, out.getvalue(), {"name": "山田 花子"}).ok


# ---------------------------------------------------------------------- 2. 差し込み・計算のフィールドの結果
MERGE = p(r("宛名："), '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD 宛名 </w:instrText></w:r>'
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«宛名»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
SIMPLE = p(r("日付："), '<w:fldSimple w:instr=" DATE \\@ &quot;yyyy年M月d日&quot; "><w:r><w:t>2026年1月1日</w:t></w:r></w:fldSimple>')


@pytest.mark.parametrize("body,old", [(MERGE, "«宛名»"), (SIMPLE, "2026年1月1日")])
def test_field_result_is_refused_with_reason(body, old):
    with pytest.raises(FillError, match="フィールド"):
        fill(DocxDocument(docx(body)), {"@1": {"old": old, "new": "千代田区長"}})


def test_text_outside_field_still_writable():
    blob = docx(MERGE)
    res = fill(DocxDocument(blob), {"@1": {"old": "宛名：", "new": "宛名（様）："}})
    assert prove(blob, res.blob, {"P1.w1": {"@": {"old": "宛名：", "new": "宛名（様）："}, "value": "（様）"}}).ok


# ---------------------------------------------------------------------- 3. 箇条書きの □
BULLETS = (p(r("Ⅰ　負担上限月額に関する認定"), ppr='<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>')
           + p(r("Ⅱ　食費の減免"), ppr='<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>')
           + p(r("郵送を希望する"), ppr='<w:pPr><w:pStyle w:val="List"/></w:pPr>')
           + p(r("見出しの印"), ppr='<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr>'))


def test_bullet_boxes_are_checkboxes(tmp_path):
    path = write(tmp_path, docx(BULLETS, numbering=True))
    fl = {f["id"]: f for f in api.fields(path)}
    assert fl["P1.b"]["kind"] == "checkbox" and fl["P1.b"]["value"] is False
    assert "Ⅰ" in fl["P1.b"]["label"]
    assert "P3.b" in fl  # スタイルで箇条書きにした段落（Wingdings の □）
    assert "P4.b" not in fl  # ■ の箇条書きは、チェックの欄ではない
    _k, _d, parts = api.window(path)
    assert "☐［P1.b］" in parts[0][1]


def test_bullet_fill_proof_readback_and_only_that_paragraph(tmp_path):
    path = write(tmp_path, docx(BULLETS, numbering=True))
    out = str(tmp_path / "o.docx")
    r_ = api.fill(path, {"P1.b": True, "P3.b": True}, out)
    assert r_["proof_ok"], r_["unexpected"]
    assert r_["read_back_ok"], r_["read_back_differences"]
    d = DocxDocument(open(out, "rb").read())
    labels = [q.label for q in d.body]
    # Ⅱ の段落の印は □ のまま。Wingdings の □（0xA8）は Wingdings の 0xFE：U+1F5F9 BALLOT BOX WITH BOLD CHECK
    assert labels == ["■", "□", "\U0001F5F9", "■"]
    num = d.xml("word/numbering.xml")
    assert b"&#xF0A8;" in num or "".encode() in num  # 元の定義はそのまま
    fl = {f["id"]: f for f in api.fields(out)}
    assert fl["P1.b"]["value"] is True and fl["P3.b"]["value"] is True and fl["P2.b"]["value"] is False
    with pytest.raises(api.ApiError, match="すでに書き込んだ"):
        api.fill(out, {"P1.b": True}, str(tmp_path / "o2.docx"))


def test_bullet_proof_catches_change_of_other_definition(tmp_path):
    blob = docx(BULLETS, numbering=True)
    res = fill(DocxDocument(blob), {"P1.b": True})
    z = zipfile.ZipFile(io.BytesIO(res.blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z2:
        for n in z.namelist():
            data = z.read(n)
            if n == "word/numbering.xml":
                data = data.replace('<w:lvlText w:val="■"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'.encode(),
                                    '<w:lvlText w:val="●"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'.encode())
            z2.writestr(n, data)
    assert not prove(blob, out.getvalue(), {"P1.b": True}).ok


def test_plain_text_box_still_writable(tmp_path):
    path = write(tmp_path, docx(p(r("□申請者と同じ　□別住所"))))
    out = str(tmp_path / "o.docx")
    r_ = api.fill(path, {"P1.c0": True}, out)
    assert r_["proof_ok"] and r_["read_back_ok"]
    assert DocxDocument(open(out, "rb").read()).body[0].text.startswith("■申請者と同じ")


# ---------------------------------------------------------------------- 4. 図形・テキストボックスの字
def _anchor(inner: str, n: int) -> str:
    return ('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph">'
            f'<wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2000000" cy="400000"/><wp:wrapNone/><wp:docPr id="{n}" name="S{n}"/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' + inner + '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor>'
            '</w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect><v:textbox><w:txbxContent>'
            + p(r("代わりの中身")) + '</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r>')


def test_shape_text_at_anchor_and_empty_shapes_are_not_fields(tmp_path):
    note = '<wps:txbx><w:txbxContent>' + p(r("※この委任状は、本人がすべて記入してください。")) + '</w:txbxContent></wps:txbx>'
    empty_box = '<wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>'
    vml = ('<w:r><w:pict><v:shape style="width:100pt;height:20pt"><v:textbox><w:txbxContent>' + p(r("（注）押印が必要です"))
           + '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>')
    body = (p(r("委任状")) + p(r("代理人"), _anchor(note, 1)) + p(_anchor("", 2), _anchor("", 3), _anchor(empty_box, 4))
            + p(r("私は"), vml) + p(r("以上")))
    path = write(tmp_path, docx(body))
    _k, _d, parts = api.window(path)
    lines = [ln for ln in parts[0][1].splitlines() if ln.startswith("- ")]
    i = next(k for k, ln in enumerate(lines) if ln.startswith("- 代理人"))
    assert "本人がすべて記入" in lines[i + 1]  # 置いた段落のすぐ後ろ
    j = next(k for k, ln in enumerate(lines) if ln.startswith("- 私は"))
    assert "押印が必要" in lines[j + 1]
    assert "代わりの中身" not in parts[0][1]
    ids = [f["id"] for f in api.fields(path)]
    # 図形だけの段落（P4）・空のテキストボックスの段落（P5）は、欄の候補にしない
    assert not any(i.startswith(("P4.", "P5.")) for i in ids), ids


# ---------------------------------------------------------------------- 5. 日付の欄の読み戻し
DATE = p(r("申請日："), '<w:sdt><w:sdtPr><w:tag w:val="date"/><w:id w:val="4"/><w:showingPlcHdr/><w:date><w:dateFormat w:val="ggge年M月d日"/>'
         '<w:lid w:val="ja-JP"/><w:calendar w:val="japan"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>日付を選択</w:t></w:r></w:sdtContent></w:sdt>')


def test_date_control_read_back_uses_the_date_format(tmp_path):
    path = write(tmp_path, docx(DATE))
    out = str(tmp_path / "o.docx")
    r_ = api.fill(path, {"date": "2026-11-03"}, out)
    assert r_["read_back_ok"], r_["read_back_differences"]
    assert r_["written_as"] == {"date": "令和8年11月3日"}
    assert "date（近くの字「申請日」）：令和8年11月3日" in r_["say_to_user"]


# ---------------------------------------------------------------------- 6. 見出し
LABELS = (p(r("同意する："), '<w:sdt><w:sdtPr><w:tag w:val="agree"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>'
            '<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>')
          + p(r("人数："), '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="ninzu"/><w:textInput/></w:ffData></w:fldChar></w:r>'
              '<w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>     </w:t></w:r>'
              '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
          + '<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>'
          '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>' + p(r("住所")) + '</w:tc><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr>'
          '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc>' + p(r("電話")) + '</w:tc><w:tc><w:p/></w:tc></w:tr>'
          '<w:tr><w:tc>' + p(r("事業所名")) + '</w:tc><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>')


def test_labels_come_from_printed_text(tmp_path):
    path = write(tmp_path, docx(LABELS))
    fl = {f["id"]: f["label"] for f in api.fields(path)}
    assert fl["agree"] == "同意する"
    assert fl["ninzu"] == "人数"
    assert fl["T1.R3.C2"] == "事業所名"  # 2列にまたがるマスの上の、1列だけの「電話」は列の見出しではない
    out = str(tmp_path / "o.docx")
    r_ = api.fill(path, {"agree": True, "ninzu": "12", "T1.R1.C2": "東京都千代田区1-1"}, out)
    say = r_["say_to_user"]
    assert "agree（近くの字「同意する」）：☑" in say and "ninzu（近くの字「人数」）：12" in say and "（見出しなし）" not in say and "：：" not in say
    assert "東京都千代田区1-1" not in say.split("まだ空いている欄")[1]  # 書いた値を、ほかの欄の見出しにしない
    after = {f["id"]: f["label"] for f in api.fields(out)}
    assert after["T1.R2.C3"] == "電話"


# ---------------------------------------------------------------------- 変更履歴
def test_tracked_deletion_shown_as_deleted(tmp_path):
    body = p(r("備考："), '<w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>旧い注意</w:delText></w:r></w:del>', r("新しい注意"))
    path = write(tmp_path, docx(body))
    _k, _d, parts = api.window(path)
    # 変更した人・日時（w:author・w:date）は、書いてある値のまま添える
    assert "備考：新しい注意〔変更履歴で削除された字（いまの本文ではない。変更した人：a・日時：2026-01-01T00:00:00Z）: 旧い注意〕" in parts[0][1]
    assert api.audit(path)["checks"][0]["ok"]
