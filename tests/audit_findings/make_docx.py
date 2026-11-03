"""監査用の合成 Word（.docx）を、XML を手で書いて作る。

  python3 tests/audit_findings/make_docx.py 出力フォルダ

word_all.docx：
  - コンテンツコントロール：文字（plainText）・チェック（w14:checkbox）・ドロップダウン（表示と値が別）・日付（yyyy年M月d日）
  - 旧式のフォームの欄：FORMTEXT（最大 5 字・数）・FORMCHECKBOX・FORMDROPDOWN
  - 表：横に結合（gridSpan）・縦に結合（vMerge）・空のマス
  - テキストボックス（wps:txbx）の中の「氏名：＿＿＿＿」
  - ヘッダー・フッターの「受付番号：　　」
  - 変更履歴（w:ins・w:del）が残っている段落
  - MERGEFIELD（差し込みの欄）
word_protected.docx：上と同じで、documentProtection w:edit="forms"（フォームの欄だけ書ける）
"""
import sys
import zipfile
from pathlib import Path

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
     'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
     'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"')


def p(*runs, ppr=""):
    return f"<w:p>{ppr}{''.join(runs)}</w:p>"


def r(t):
    return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'


def body() -> str:
    parts = []
    parts.append(p(r("申請書（監査用）")))
    # コンテンツコントロール
    parts.append(p(r("氏名："), '<w:sdt><w:sdtPr><w:alias w:val="氏名"/><w:tag w:val="name"/><w:id w:val="101"/><w:showingPlcHdr/><w:text/></w:sdtPr>'
                   '<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>ここをクリックして入力</w:t></w:r></w:sdtContent></w:sdt>'))
    parts.append(p(r("同意する："), '<w:sdt><w:sdtPr><w:tag w:val="agree"/><w:id w:val="102"/><w14:checkbox><w14:checked w14:val="0"/>'
                   '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>'
                   '<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>'))
    parts.append(p(r("区分："), '<w:sdt><w:sdtPr><w:tag w:val="kubun"/><w:id w:val="103"/><w:showingPlcHdr/><w:dropDownList>'
                   '<w:listItem w:displayText="選んでください" w:value=""/><w:listItem w:displayText="新規" w:value="1"/>'
                   '<w:listItem w:displayText="変更" w:value="2"/><w:listItem w:displayText="廃止" w:value="3"/></w:dropDownList></w:sdtPr>'
                   '<w:sdtContent><w:r><w:t>選んでください</w:t></w:r></w:sdtContent></w:sdt>'))
    parts.append(p(r("申請日："), '<w:sdt><w:sdtPr><w:tag w:val="date"/><w:id w:val="104"/><w:showingPlcHdr/><w:date><w:dateFormat w:val="yyyy年M月d日"/>'
                   '<w:lid w:val="ja-JP"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr>'
                   '<w:sdtContent><w:r><w:t>日付を選択</w:t></w:r></w:sdtContent></w:sdt>'))
    # 旧式のフォームの欄
    parts.append(p(r("人数："), '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="ninzu"/><w:enabled/><w:calcOnExit w:val="0"/>'
                   '<w:textInput><w:type w:val="number"/><w:maxLength w:val="5"/></w:textInput></w:ffData></w:fldChar></w:r>'
                   '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                   '<w:r><w:t>     </w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'))
    parts.append(p(r("郵送を希望："), '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="yuso"/><w:enabled/><w:calcOnExit w:val="0"/>'
                   '<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r>'
                   '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'))
    parts.append(p(r("性別："), '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="sex"/><w:enabled/><w:calcOnExit w:val="0"/>'
                   '<w:ddList><w:listEntry w:val="男"/><w:listEntry w:val="女"/><w:listEntry w:val="回答しない"/></w:ddList></w:ffData></w:fldChar></w:r>'
                   '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'))
    # 表（結合）
    tc = lambda t, extra="": f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>{extra}</w:tcPr>{p(r(t)) if t else "<w:p/>"}</w:tc>'  # noqa: E731
    rows = [
        "<w:tr>" + tc("住所", '<w:vMerge w:val="restart"/>') + tc("", '<w:gridSpan w:val="2"/>') + "</w:tr>",
        "<w:tr>" + tc("", "<w:vMerge/>") + tc("電話") + tc("") + "</w:tr>",
        "<w:tr>" + tc("事業所名") + tc("", '<w:gridSpan w:val="2"/>') + "</w:tr>",
    ]
    parts.append('<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
                 '<w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
                 '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' + "".join(rows) + "</w:tbl>")
    # テキストボックス
    parts.append(p('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
                   'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
                   '<wp:positionH relativeFrom="column"><wp:posOffset>3000000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph">'
                   '<wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2000000" cy="400000"/><wp:wrapNone/><wp:docPr id="5" name="TB1"/>'
                   '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr>'
                   '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                   '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>'
                   + p(r("担当者：＿＿＿＿＿＿")) + '</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor>'
                   '</w:drawing></mc:Choice></mc:AlternateContent></w:r>', r("（右の枠に担当者）")))
    # 変更履歴
    parts.append(p(r("備考："), '<w:del w:id="1" w:author="市役所" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>旧い注意書き</w:delText></w:r></w:del>',
                   '<w:ins w:id="2" w:author="市役所" w:date="2026-01-01T00:00:00Z"><w:r><w:t>新しい注意書き　</w:t></w:r></w:ins>', r("＿＿＿＿＿＿")))
    # MERGEFIELD
    parts.append(p(r("宛名："), '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD 宛名 </w:instrText></w:r>'
                   '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«宛名»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'))
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'
            '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(parts)}{sect}</w:body></w:document>'


def build(path: Path, protected: bool) -> None:
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
             '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
    prot = '<w:documentProtection w:edit="forms" w:enforcement="1"/>' if protected else ""
    settings = f'<?xml version="1.0" encoding="UTF-8"?><w:settings {W}>{prot}<w:trackRevisions/></w:settings>'
    hdr = f'<?xml version="1.0" encoding="UTF-8"?><w:hdr {W}>{p(r("受付番号：　　　　　"))}</w:hdr>'
    ftr = f'<?xml version="1.0" encoding="UTF-8"?><w:ftr {W}>{p(r("記入日：　　年　　月　　日"))}</w:ftr>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", body())
        z.writestr("word/settings.xml", settings)
        z.writestr("word/header1.xml", hdr)
        z.writestr("word/footer1.xml", ftr)


if __name__ == "__main__":
    o = Path(sys.argv[1])
    build(o / "word_all.docx", False)
    build(o / "word_protected.docx", True)
    print("ok")
