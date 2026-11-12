"""試しの Excel（.xlsx）を作る：記号の書体（Wingdings・Symbol・Wingdings 2）のセルのあるシート。既定の書体は Arial。

A 列は見出し（Agree・Decline・Alpha・Direct・Japanese・W2・PUA・Number・Empty・Rich）、B 列が記号の書体のセル：
B1「þ」・B2「¨」（Wingdings の、印のついた箱・印のない箱）、B3「abg」（Symbol）、B4「☑」（U+2611 を Wingdings で）、B5「山田」（Wingdings）、
B6「RT」（Wingdings 2）、B7 U+F0FC（Wingdings）、B8 数 1（Wingdings）、B9 空（Wingdings）、B10「Rich 」+ Wingdings の「ü」+「 done」。
extra：11 行目からの行の数（A は字、B は数、C は Wingdings の「þ」「¨」：大きな表の道の試し）。formula：D1 に =COUNTIF(B1:B2,"þ")。"""

from __future__ import annotations

import io
import zipfile

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension='
      '"rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartN'
      'ame="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sh'
      'eet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="ap'
      'plication/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxml'
      'formats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationsh'
      'ip Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
WB = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://sch'
      'emas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
WBR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationsh'
      'ip Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rI'
      'd2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schema'
      's.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
ST = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4">'
      '<font><sz val="11"/><name val="Arial"/></font><font><sz val="11"/><name val="Wingdings"/><charset val="2"/></font><font><sz val="11"/><name val="Symbo'
      'l"/><charset val="2"/></font><font><sz val="11"/><name val="Wingdings 2"/><charset val="2"/></font></fonts><fills count="2"><fill><patternFill pattern'
      'Type="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></bor'
      'der></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="'
      '0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0'
      '" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>')
RICH = ('<si><r><t xml:space="preserve">Rich </t></r><r><rPr><sz val="11"/><rFont val="Wingdings"/><charset val="2"/></rPr><t>ü</t></r><r><rPr><sz val="11"/><r'
      'Font val="Arial"/></rPr><t xml:space="preserve"> done</t></r></si>')
STRINGS = ["Agree", "þ", "Decline", "¨", "Alpha", "abg", "Direct", "☑", "Japanese", "山田", "W2", "RT", "PUA", "\uf0fc", "Number", "Empty", "Rich"]
ROWS = [(0, 0, 1, 1), (2, 0, 3, 1), (4, 0, 5, 2), (6, 0, 7, 1), (8, 0, 9, 1), (10, 0, 11, 3), (12, 0, 13, 1)]
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def make(extra: int = 0, formula: bool = False) -> bytes:
    sst = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{_MAIN}">'
           + "".join(f'<si><t xml:space="preserve">{s}</t></si>' for s in STRINGS) + RICH + "</sst>")
    body = ""
    for n, (a, sa, b, sb) in enumerate(ROWS, 1):
        more = '<c r="D1"><f>COUNTIF(B1:B2,"þ")</f><v>1</v></c>' if formula and n == 1 else ""
        body += f'<row r="{n}"><c r="A{n}" s="{sa}" t="s"><v>{a}</v></c><c r="B{n}" s="{sb}" t="s"><v>{b}</v></c>{more}</row>'
    body += '<row r="8"><c r="A8" t="s"><v>14</v></c><c r="B8" s="1"><v>1</v></c></row>'
    body += '<row r="9"><c r="A9" t="s"><v>15</v></c><c r="B9" s="1"/></row>'
    body += '<row r="10"><c r="A10" t="s"><v>16</v></c><c r="B10" t="s"><v>17</v></c></row>'
    for n in range(11, 11 + extra):
        body += (f'<row r="{n}"><c r="A{n}" t="inlineStr"><is><t>利用者{n} の記録の行です</t></is></c><c r="B{n}"><v>{n}</v></c>'
                 f'<c r="C{n}" s="1" t="inlineStr"><is><t>{"þ" if n % 2 else "¨"}</t></is></c></row>')
    sheet = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{_MAIN}"><sheetData>' + body
             + "</sheetData></worksheet>")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in (("[Content_Types].xml", CT), ("_rels/.rels", RELS), ("xl/workbook.xml", WB), ("xl/_rels/workbook.xml.rels", WBR),
                           ("xl/styles.xml", ST), ("xl/sharedStrings.xml", sst), ("xl/worksheets/sheet1.xml", sheet)):
            z.writestr(name, data)
    return buf.getvalue()
