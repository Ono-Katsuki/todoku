"""埋め込んだもの（グラフ・SmartArt・OLE の部品・暗号化した要約のストリーム）を平文で読むテスト。

テストの中で小さなファイル（グラフの部品・.xlsx・.docx・.xls のグラフの部分・複合ファイルの入れ物・数式・PowerPoint の
ストリーム）を組み立てて確かめ、公開の様式（corpus）があればそれでも確かめる（なければ飛ばす）。
"""

from __future__ import annotations

import glob
import io
import struct
import zipfile
import zlib
from dataclasses import replace
from pathlib import Path

import pytest

from todoku.bincrypto import decrypt_binary, encrypt_binary, new_key, summary_entries
from todoku.cfb import CFB
from todoku.docx.document import DocxDocument
from todoku.docx.window import build as word_window, to_markdown as word_md
from todoku.embeds.chart import chart_dict, chart_markdown, read_chart
from todoku.embeds.mtef import equation_text
from todoku.embeds.ole import compobj, describe, ole10native
from todoku.embeds.ppt import ppt_slides
from todoku.embeds.smartart import smartart_lines
from todoku.embeds.word import fobjh_storage
from todoku.xls.book import XlsWorkbook
from todoku.xlsx import window as xw
from todoku.xlsx.workbook import Workbook

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' \
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ====================================================================== 組み立てる道具
def make_cfb(tree: dict) -> bytes:
    """入れ物（dict）とストリーム（bytes）の木から、複合ファイル（版 3・512 バイトのセクタ）を作る。
    4096 バイト未満のストリームはミニストリームに置く。入れ物の CLSID は、名前に「|{16 バイトの hex}」を付けて指定できる。"""
    ss, ms = 512, 64
    END, FREE, FATSECT, NOS = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF
    sectors: list[bytes] = []
    fat: list[int] = []

    def put(data: bytes) -> int:
        if not data:
            return END
        start, k = len(sectors), -(-len(data) // ss)
        for i in range(k):
            sectors.append(data[i * ss:(i + 1) * ss].ljust(ss, b"\0"))
            fat.append(start + i + 1 if i < k - 1 else END)
        return start

    mini, minifat = bytearray(), []
    ents: list[list] = []  # [名前, 種類, 始まり, 大きさ, 子, 右, clsid]

    def add(name: str, node) -> int:
        clsid = b""
        if "|" in name:
            name, h = name.split("|")
            clsid = bytes.fromhex(h)
        idx = len(ents)
        if isinstance(node, dict):
            ents.append([name, 1, 0, 0, NOS, NOS, clsid])
            kids = [add(k, v) for k, v in node.items()]
            for a, b in zip(kids, kids[1:]):
                ents[a][5] = b
            ents[idx][4] = kids[0] if kids else NOS
        else:
            if len(node) >= 4096:
                start = put(node)
            else:
                start = len(mini) // ms
                k = -(-len(node) // ms)
                mini.extend(node.ljust(k * ms, b"\0"))
                minifat.extend([start + i + 1 for i in range(k - 1)] + [END] if k else [])
                start = start if k else END
            ents.append([name, 2, start, len(node), NOS, NOS, clsid])
        return idx

    ents.append(["Root Entry", 5, 0, 0, NOS, NOS, b""])
    kids = [add(k, v) for k, v in tree.items()]
    for a, b in zip(kids, kids[1:]):
        ents[a][5] = b
    ents[0][4] = kids[0] if kids else NOS
    ents[0][2] = put(bytes(mini))
    ents[0][3] = len(mini)
    minifat_at = put(struct.pack(f"<{len(minifat)}I", *minifat)) if minifat else END
    raw = b""
    for name, kind, start, size, child, right, clsid in ents:
        nm = (name + "\0").encode("utf-16-le")
        raw += (nm.ljust(64, b"\0") + struct.pack("<HBB3I", len(nm), kind, 1, NOS, right, child)
                + clsid.ljust(16, b"\0") + bytes(20) + struct.pack("<IQ", start, size))
    raw = raw.ljust(-(-len(raw) // ss) * ss, b"\0")
    dir_at = put(raw)
    n_fat = -(-(len(fat) + 2) // (ss // 4))
    fat_at = len(sectors)
    fat.extend([FATSECT] * n_fat)
    table = fat + [FREE] * (n_fat * ss // 4 - len(fat))
    for k in range(n_fat):
        sectors.append(struct.pack(f"<{ss // 4}I", *table[k * ss // 4:(k + 1) * ss // 4]))
    head = (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(16) + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
            + struct.pack("<9I", 0, n_fat, dir_at, 0, 4096, minifat_at, -(-len(minifat) * 4 // ss), END, 0)
            + struct.pack("<109I", *([fat_at + k for k in range(n_fat)] + [FREE] * (109 - n_fat))))
    return head + b"".join(sectors)


def compobj_stream(user_type: str, prog: str) -> bytes:
    """\\x01CompObj（MS-OLEDS 2.3.8）：見出し 28 バイト・表示名・クリップボードの形式（なし）・ProgID。"""
    def lp(s: str) -> bytes:
        b = s.encode("cp932") + b"\0"
        return struct.pack("<I", len(b)) + b

    return bytes(28) + lp(user_type) + struct.pack("<I", 0) + lp(prog) + struct.pack("<I", 0x71B239F4) + struct.pack("<I", 0)


def zip_of(parts: dict[str, str | bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, v in parts.items():
            z.writestr(n, v)
    return buf.getvalue()


def rels(*items: tuple[str, str, str]) -> str:
    body = "".join(f'<Relationship Id="{i}" Type="{REL}/{t}" Target="{target}"/>' for i, t, target in items)
    return f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>'


CT = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'


def chart_xml(title: str = "売上の推移") -> str:
    def str_cache(vals):
        return f'<c:ptCount val="{len(vals)}"/>' + "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))

    ser = lambda i, name, vals: (  # noqa: E731
        f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>{str_cache([name])}</c:strCache></c:strRef></c:tx>'
        f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache>{str_cache(["4月", "5月", "6月"])}</c:strCache></c:strRef></c:cat>'
        f'<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>#,##0</c:formatCode>{str_cache(vals)}</c:numCache></c:numRef></c:val></c:ser>')
    return (f'<c:chartSpace {C}><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>{title}</a:t></a:r></a:p></c:rich></c:tx></c:title>'
            f'<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'
            + ser(0, "東京", [1200, 3400.5, 0]) + ser(1, "大阪", [10, 20, 30]).replace(
                "<c:val>", '<c:dLbls><c:dLbl><c:idx val="1"/><c:tx><c:rich><a:p><a:r><a:t>最高</a:t></a:r></a:p></c:rich></c:tx></c:dLbl></c:dLbls><c:val>')
            + '<c:axId val="1"/><c:axId val="2"/></c:barChart>'
            '<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="2"/><c:order val="2"/><c:tx><c:v>目標</c:v></c:tx>'
            '<c:val><c:numLit><c:formatCode>0%</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>0.5</c:v></c:pt>'
            '<c:pt idx="2"><c:v>0.75</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>'
            '<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>月</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>'
            '<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:r><a:t>金額（円）</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>'
            '</c:plotArea></c:chart><c:userShapes r:id="rId1"/></c:chartSpace>')


USER_SHAPES = ('<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
               'xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               '<cdr:relSizeAnchor><cdr:sp><cdr:txBody><a:p><a:r><a:t>出典：市の統計</a:t></a:r></a:p></cdr:txBody></cdr:sp></cdr:relSizeAnchor></c:userShapes>')


# ====================================================================== グラフの部品
def test_chart_reads_titles_series_cached_values_and_labels():
    pkg = zip_of({"xl/charts/chart1.xml": chart_xml(), "xl/charts/_rels/chart1.xml.rels": rels(("rId1", "chartUserShapes", "../drawings/drawing9.xml")),
                  "xl/drawings/drawing9.xml": USER_SHAPES})
    from todoku.zipraw import RawZip

    z = RawZip(pkg)
    ch = read_chart(z.read("xl/charts/chart1.xml"), "xl/charts/chart1.xml", z)
    assert ch.title == "売上の推移" and ch.kinds == ["縦棒グラフ", "折れ線グラフ"]
    assert [s.name for s in ch.series] == ["東京", "大阪", "目標"]
    assert ch.series[0].vals == {0: "1,200", 1: "3,401", 2: "0"}  # 控えの表示形式（#,##0）で見えている字
    assert ch.series[2].vals == {0: "50%", 2: "75%"} and ch.series[1].labels == {1: "最高"}
    assert ch.axes == [("横軸（項目）", "月"), ("縦軸（数値）", "金額（円）")]
    assert ch.texts == ["グラフの中の文字：出典：市の統計"]
    md = "\n".join(chart_markdown(chart_dict(ch, 1)))
    assert "### グラフ1：売上の推移（縦棒グラフ・折れ線グラフ）" in md
    assert "| 項目 | 東京 | 大阪 | 目標 |" in md and "| 4月 | 1,200 | 10 | 50% |" in md and "| 5月 | 3,401 | 20 |  |" in md
    assert "データラベル（大阪）：5月「最高」" in md


def test_chart_without_cache_reads_cells_and_multi_level_categories():
    xml = (f'<c:chartSpace {C}><c:chart><c:plotArea><c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/>'
           '<c:cat><c:multiLvlStrRef><c:f>S!$A$1:$B$3</c:f><c:multiLvlStrCache><c:ptCount val="3"/>'
           '<c:lvl><c:pt idx="0"><c:v>1月</c:v></c:pt><c:pt idx="1"><c:v>2月</c:v></c:pt><c:pt idx="2"><c:v>1月</c:v></c:pt></c:lvl>'
           '<c:lvl><c:pt idx="0"><c:v>2025年</c:v></c:pt><c:pt idx="2"><c:v>2026年</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>'
           '<c:val><c:numRef><c:f>S!$C$1:$C$3</c:f></c:numRef></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>')
    ch = read_chart(xml.encode(), "chart.xml", resolve=lambda ref: ["1", "2", "3"] if ref == "S!$C$1:$C$3" else None)
    s = ch.series[0]
    assert ch.kinds == ["円グラフ"] and s.name == "系列1"
    assert s.cats == {0: "2025年 ／ 1月", 1: "2025年 ／ 2月", 2: "2026年 ／ 1月"}
    assert s.vals == {0: "1", 1: "2", 2: "3"} and "参照先のセル" in s.notes[0]


def test_chartex_and_smartart():
    cx = ('<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0">'
          '<cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">申請</cx:pt><cx:pt idx="1">審査</cx:pt></cx:lvl></cx:strDim>'
          '<cx:numDim type="val"><cx:lvl ptCount="2" formatCode="General"><cx:pt idx="0">100</cx:pt><cx:pt idx="1">40</cx:pt></cx:lvl></cx:numDim>'
          '</cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>手続きの件数</cx:v></cx:txData></cx:tx></cx:title>'
          '<cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"><cx:tx><cx:txData><cx:v>件数</cx:v></cx:txData></cx:tx>'
          '<cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>')
    ch = read_chart(cx.encode(), "chartEx1.xml")
    assert ch.title == "手続きの件数" and ch.kinds == ["じょうごグラフ"]
    assert ch.series[0].cats == {0: "申請", 1: "審査"} and ch.series[0].vals == {0: "100", 1: "40"}
    dm = ('<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
          '<dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>申請</a:t></a:r></a:p></dgm:t></dgm:pt>'
          '<dgm:pt modelId="2"><dgm:t><a:p><a:r><a:t>書類をそろえる</a:t></a:r></a:p></dgm:t></dgm:pt>'
          '<dgm:pt modelId="3"><dgm:t><a:p><a:r><a:t>審査</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="9" type="pres"/></dgm:ptLst>'
          '<dgm:cxnLst><dgm:cxn modelId="10" srcId="0" destId="3" srcOrd="1"/><dgm:cxn modelId="11" srcId="0" destId="1" srcOrd="0"/>'
          '<dgm:cxn modelId="12" srcId="1" destId="2" srcOrd="0"/><dgm:cxn modelId="13" type="presOf" srcId="1" destId="9"/></dgm:cxnLst></dgm:dataModel>')
    assert smartart_lines(dm.encode()) == ["申請", "　書類をそろえる", "審査"]


# ====================================================================== Excel（.xlsx）：グラフ・グラフのシート・埋め込み
SHEET = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{REL}" '
         'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
         '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>月</t></is></c><c r="B1" t="inlineStr"><is><t>東京</t></is></c></row>'
         '<row r="2"><c r="A2" t="inlineStr"><is><t>4月</t></is></c><c r="B2"><v>1200</v></c></row></sheetData>'
         '<drawing r:id="rId1"/><oleObjects><mc:AlternateContent><mc:Choice Requires="x14"><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId2">'
         '<objectPr defaultSize="0" r:id="rId3"><anchor moveWithCells="1"><from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row>'
         '<xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor>'
         '</objectPr></oleObject></mc:Choice><mc:Fallback><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId2"/></mc:Fallback>'
         '</mc:AlternateContent></oleObjects></worksheet>').replace("{REL}", REL)
DRAWING = ('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           f'xmlns:r="{REL}"><xdr:twoCellAnchor><xdr:from><xdr:col>14</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
           '<xdr:to><xdr:col>20</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
           '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="グラフ 1" descr="月ごとの売上"/></xdr:nvGraphicFramePr>'
           '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>'
           '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>')
INNER_DOCX = zip_of({"[Content_Types].xml": CT, "_rels/.rels": rels(("rId1", "officeDocument", "word/document.xml")),
                     "word/document.xml": '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
                                          '<w:p><w:r><w:t>埋め込んだ文書の字</w:t></w:r></w:p></w:body></w:document>'})


def make_xlsx() -> bytes:
    wbx = (f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{REL}"><sheets>'
           '<sheet name="データ" sheetId="1" r:id="rId1"/><sheet name="グラフ1" sheetId="2" r:id="rId2"/></sheets></workbook>')
    return zip_of({
        "[Content_Types].xml": CT, "_rels/.rels": rels(("rId1", "officeDocument", "xl/workbook.xml")), "xl/workbook.xml": wbx,
        "xl/_rels/workbook.xml.rels": rels(("rId1", "worksheet", "worksheets/sheet1.xml"), ("rId2", "chartsheet", "chartsheets/sheet1.xml")),
        "xl/worksheets/sheet1.xml": SHEET,
        "xl/worksheets/_rels/sheet1.xml.rels": rels(("rId1", "drawing", "../drawings/drawing1.xml"), ("rId2", "package", "../embeddings/Microsoft_Word_Document.docx"),
                                                    ("rId3", "image", "../media/image1.emf")),
        "xl/drawings/drawing1.xml": DRAWING, "xl/drawings/_rels/drawing1.xml.rels": rels(("rId1", "chart", "../charts/chart1.xml")),
        "xl/charts/chart1.xml": chart_xml().replace('<c:userShapes r:id="rId1"/>', ""),
        "xl/chartsheets/sheet1.xml": f'<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{REL}"><drawing r:id="rId1"/></chartsheet>',
        "xl/chartsheets/_rels/sheet1.xml.rels": rels(("rId1", "drawing", "../drawings/drawing2.xml")),
        "xl/drawings/drawing2.xml": DRAWING.replace("twoCellAnchor", "absoluteAnchor"),
        "xl/drawings/_rels/drawing2.xml.rels": rels(("rId1", "chart", "../charts/chart2.xml")),
        "xl/charts/chart2.xml": chart_xml("グラフのシートの題名").replace('<c:userShapes r:id="rId1"/>', ""),
        "xl/embeddings/Microsoft_Word_Document.docx": INNER_DOCX,
    })


def test_xlsx_window_shows_charts_chartsheet_and_embedded_word():
    from todoku.xlsx.audit import run

    blob = make_xlsx()
    doc = xw.build(Workbook(blob, "グラフ.xlsx"))
    md = xw.to_markdown(doc)
    assert "### グラフ1：売上の推移（縦棒グラフ・折れ線グラフ、O6〜U21）" in md and "代替テキスト「月ごとの売上」" in md
    assert "## シート「グラフ1」（グラフのシート）" in md and "### グラフ2：グラフのシートの題名" in md
    assert "### 埋め込み1：Word の文書（Word.Document.12・xl/embeddings/Microsoft_Word_Document.docx）" in md
    assert "位置：D5〜G10" in md and "- 埋め込んだ文書の字" in md
    checks = {c.name: c for c in run(blob, "グラフ.xlsx").checks}
    assert checks["読み落とし（グラフ）"].ok and "グラフ 2 個" in checks["読み落とし（グラフ）"].detail
    assert checks["読み落とし（埋め込み）"].ok, checks["読み落とし（埋め込み）"].items


# ====================================================================== Word（.docx）：グラフ・SmartArt・埋め込み
DOCX_BODY = ('<w:p><w:r><w:t>実績は次のとおり。</w:t></w:r><w:r><w:drawing><wp:inline><wp:docPr id="1" name="グラフ 1" descr="実績のグラフ"/>'
             '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId5"/></a:graphicData></a:graphic>'
             '</wp:inline></w:drawing></w:r></w:p>'
             '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="2" name="図表 2"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
             '<dgm:relIds r:dm="rId6" r:lo="rId7" r:qs="rId8" r:cs="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
             '<w:p><w:r><w:object><v:shape id="_x0000_i1025" alt="手順の資料"/><o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" r:id="rId10"/>'
             '</w:object></w:r><w:r><w:object><o:OLEObject Type="Embed" ProgID="Equation.3" r:id="rId11"/></w:object></w:r>'
             '<w:r><w:object><o:OLEObject Type="Embed" ProgID="PowerPoint.Show.8" r:id="rId12"/></w:object></w:r>'
             '<w:r><w:object><o:OLEObject Type="Embed" ProgID="Visio.Drawing.11" r:id="rId13"/></w:object></w:r></w:p>')
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:v="urn:schemas-microsoft-com:vml" '
      f'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="{REL}"')


def tiny_pdf(text: str) -> bytes:
    """1ページに text（ASCII）を書いた小さな PDF。"""
    content = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode()
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
            b"<< /Length %d >>\nstream\n" % len(content) + content + b"\nendstream",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
    out = bytearray(b"%PDF-1.4\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1) + b"".join(b"%010d 00000 n \n" % o for o in offs)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)
    return bytes(out)


def package_native(fname: str, data: bytes) -> bytes:
    """\\x01Ole10Native のパッケージ（ファイルを埋め込んだもの）。"""
    nm = fname.encode("cp932") + b"\0"
    tmp = b"C:\\tmp\\" + nm
    native = struct.pack("<H", 2) + nm + nm + struct.pack("<HH", 0, 3) + struct.pack("<I", len(tmp)) + tmp + struct.pack("<I", len(data)) + data
    return struct.pack("<I", len(native)) + native


def mtef_fraction() -> bytes:
    """数式「x＝(a＋1)/(2)」の Equation Native（見出し 28 バイト ＋ MTEF 版 3）。"""
    def ch(c: str, face: int = 3) -> bytes:
        return bytes([0x02, face + 128]) + struct.pack("<H", ord(c))

    line = lambda *objs: b"\x01" + b"".join(objs) + b"\x00"  # noqa: E731
    mtef = bytes([3, 1, 1, 3, 0]) + line(ch("x"), ch("="), bytes([0x03, 14, 0, 0]) + line(ch("a"), ch("+", 6), ch("1", 8)) + line(ch("2", 8)) + b"\x00") + b"\x00"
    return struct.pack("<H", 28) + bytes(26) + mtef


def ppt_document() -> tuple[bytes, bytes]:
    """古い形の PowerPoint のストリーム（スライド1枚：題名はプレースホルダーの字、本文はテキストボックスの字）。"""
    def rec(ver_inst: int, t: int, body: bytes) -> bytes:
        return struct.pack("<HHI", ver_inst, t, len(body)) + body

    slwt = rec(0x000F, 0x0FF0, rec(0, 0x03F3, struct.pack("<IIIII", 2, 0, 1, 256, 0)) + rec(0, 0x0F9F, struct.pack("<I", 0))
               + rec(0, 0x0FA0, "会議の資料".encode("utf-16-le")))
    doc = rec(0x000F, 0x03E8, rec(0x0001, 0x03E9, bytes(40)) + slwt)
    textbox = rec(0x000F, 0xF00D, rec(0, 0x0F9F, struct.pack("<I", 1)) + rec(0, 0x0FA8, b"Agenda 2026"))
    title = rec(0x000F, 0xF00D, rec(0, 0x0F9E, struct.pack("<I", 0)))
    drawing = rec(0x000F, 0x040C, rec(0x000F, 0xF002, rec(0x000F, 0xF003, rec(0x000F, 0xF004, title) + rec(0x000F, 0xF004, textbox))))
    slide = rec(0x000F, 0x03EE, rec(0x0002, 0x03EF, bytes(24)) + drawing)
    d = bytearray(doc)
    slide_at = len(d)
    d += slide
    pdir_at = len(d)
    d += rec(0, 0x1772, struct.pack("<I", 1 | (2 << 20)) + struct.pack("<II", 0, slide_at))
    edit_at = len(d)
    d += rec(0, 0x0FF5, struct.pack("<IHBBIIIIHH", 256, 0, 0, 3, 0, pdir_at, 1, 2, 1, 0))
    cu = rec(0, 0x0FF6, struct.pack("<IIIHHBBH", 20, 0xE391C05F, edit_at, 0, 0x03F4, 3, 0, 0))
    return bytes(d), cu


def make_docx_with_objects() -> bytes:
    pdf_pkg = make_cfb({"\x01Ole10Native": package_native("手順.pdf", tiny_pdf("Procedure A")),
                        "\x01CompObj": compobj_stream("Package", "Package")})
    eq = make_cfb({"Equation Native": mtef_fraction(), "\x01CompObj": compobj_stream("Microsoft 数式 3.0", "Equation.3")})
    d, cu = ppt_document()
    ppt = make_cfb({"PowerPoint Document": d, "Current User": cu, "\x01CompObj": compobj_stream("Microsoft PowerPoint スライド", "PowerPoint.Show.8")})
    visio = make_cfb({"VisioDocument": b"V" * 100, "\x01CompObj": compobj_stream("Microsoft Visio 図面", "Visio.Drawing.11")})
    dm = smart = ('<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                  '<dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>受付</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>'
                  '<dgm:cxnLst><dgm:cxn modelId="5" srcId="0" destId="1" srcOrd="0"/></dgm:cxnLst></dgm:dataModel>')
    del smart
    return zip_of({
        "[Content_Types].xml": CT, "_rels/.rels": rels(("rId1", "officeDocument", "word/document.xml")),
        "word/document.xml": f"<w:document {NS}><w:body>{DOCX_BODY}</w:body></w:document>",
        "word/_rels/document.xml.rels": rels(("rId5", "chart", "charts/chart1.xml"), ("rId6", "diagramData", "diagrams/data1.xml"),
                                             ("rId10", "oleObject", "embeddings/oleObject1.bin"), ("rId11", "oleObject", "embeddings/oleObject2.bin"),
                                             ("rId12", "oleObject", "embeddings/oleObject3.bin"), ("rId13", "oleObject", "embeddings/oleObject4.bin")),
        "word/charts/chart1.xml": chart_xml().replace('<c:userShapes r:id="rId1"/>', ""),
        "word/diagrams/data1.xml": dm,
        "word/embeddings/oleObject1.bin": pdf_pkg, "word/embeddings/oleObject2.bin": eq,
        "word/embeddings/oleObject3.bin": ppt, "word/embeddings/oleObject4.bin": visio,
    })


def test_docx_window_shows_chart_smartart_and_every_embedding():
    from todoku.docx.audit import run

    blob = make_docx_with_objects()
    win = word_window(DocxDocument(blob, "埋め込み.docx"))
    md = word_md(win)
    body = "\n".join(win["stories"][0]["reading"])
    assert "〔グラフ1〕" in body and "〔SmartArt1〕" in body and "〔埋め込み1：ファイル（パッケージ）〕" in body
    assert "## グラフ1：売上の推移（縦棒グラフ・折れ線グラフ）" in md and "| 4月 | 1,200 | 10 | 50% |" in md
    assert "## SmartArt1（word/diagrams/data1.xml）" in md and "- 受付" in md
    assert "ファイル名「手順.pdf」" in md and "Procedure A" in md and "代替テキスト「手順の資料」" in md
    assert "数式：x=(a+1)/(2)" in md
    assert "スライド1：" in md and "会議の資料" in md and "Agenda 2026" in md
    assert "Microsoft Visio 図面・Visio.Drawing.11" in md and "とどくでは読めない形" in md
    checks = {c.name: c for c in run(blob, "埋め込み.docx").checks}
    for name in ("読み落とし（グラフ）", "読み落とし（SmartArt）", "読み落とし（埋め込み）"):
        assert checks[name].ok, (name, checks[name].items)


def test_ole_pieces():
    info = compobj(compobj_stream("Microsoft Excel ワークシート", "Excel.Sheet.8"))
    assert info == {"user_type": "Microsoft Excel ワークシート", "prog": "Excel.Sheet.8"}
    assert ole10native(package_native("a.txt", b"hello")) == ("a.txt", b"hello")
    assert equation_text(mtef_fraction()) == ("x=(a+1)/(2)", "")
    d, cu = ppt_document()
    assert ppt_slides(CFB(make_cfb({"PowerPoint Document": d, "Current User": cu}))) == [("スライド1", ["会議の資料", "Agenda 2026"])]
    # Word の文書を入れた OLE の入れ物（Package ストリームに .docx）
    o = describe(make_cfb({"Package": INNER_DOCX, "\x01CompObj": compobj_stream("Microsoft Word 文書", "Word.Document.12")}), "x.bin")
    assert o["kind"] == "word" and o["markdown"] == ["- 埋め込んだ文書の字"]


def test_fobjh_storage_in_data_stream():
    inner = make_cfb({"\x01CompObj": compobj_stream("Microsoft 数式 3.0", "Equation.3"), "Equation Native": mtef_fraction()})
    data = b"PAD!" + struct.pack("<hHi", 8, 0, 8 + len(inner)) + inner
    assert fobjh_storage(data, 4) == inner
    comp = zlib.compress(inner)
    data2 = struct.pack("<hHi", 8, 1, 8 + len(comp)) + comp
    assert fobjh_storage(data2, 0) == inner
    assert fobjh_storage(data, 0) is None
    assert describe(fobjh_storage(data, 4))["equation"] == "x=(a+1)/(2)"


# ====================================================================== 古い Excel（.xls）のグラフと埋め込み
def rec(t: int, body: bytes = b"") -> bytes:
    return struct.pack("<HH", t, len(body)) + body


def bof(dt: int) -> bytes:
    return rec(0x0809, struct.pack("<HHHHII", 0x0600, dt, 0x0DBB, 0x07CC, 0, 0x0206))


def short(s: str) -> bytes:
    return bytes([len(s), 1]) + s.encode("utf-16-le")


def area(r1: int, r2: int, c1: int, c2: int) -> bytes:
    """ptgArea（絶対参照）の ChartParsedFormula。"""
    rgce = bytes([0x25]) + struct.pack("<HHHH", r1, r2, c1, c2)
    return struct.pack("<H", len(rgce)) + rgce


def biff_chart(cached: bool) -> bytes:
    """縦棒グラフ1つ（系列「件数」、項目 A2:A3、値 B2:B3）。cached なら値の控え（SERIESDATA）を持つ。"""
    def brai(aid: int, rt: int, f: bytes = struct.pack("<H", 0)) -> bytes:
        return rec(0x1051, bytes([aid, rt]) + struct.pack("<HH", 0, 0) + f)

    series = (rec(0x1003, struct.pack("<6H", 3, 1, 2, 2, 1, 0)) + rec(0x1033)
              + brai(0, 1) + rec(0x100D, struct.pack("<H", 0) + short("件数"))
              + brai(1, 2, area(1, 2, 1 | 0xC000, 1 | 0xC000)) + brai(2, 2, area(1, 2, 0 | 0xC000, 0 | 0xC000)) + brai(3, 1)
              + rec(0x1045, struct.pack("<H", 0)) + rec(0x1034))
    title = (rec(0x1025, bytes(32)) + rec(0x1033) + brai(0, 1) + rec(0x100D, struct.pack("<H", 0) + short("受付の件数"))
             + rec(0x1027, struct.pack("<HHH", 1, 0, 0)) + rec(0x1034))
    axis_title = (rec(0x1025, bytes(32)) + rec(0x1033) + brai(0, 1) + rec(0x100D, struct.pack("<H", 0) + short("件"))
                  + rec(0x1027, struct.pack("<HHH", 2, 0, 0)) + rec(0x1034))
    crt = rec(0x1014, bytes(18) + struct.pack("<H", 0)) + rec(0x1033) + rec(0x1017, struct.pack("<hHH", 0, 150, 0)) + rec(0x1034)
    body = rec(0x1002, bytes(16)) + rec(0x1033) + series + rec(0x1041, bytes(18)) + rec(0x1033) + crt + rec(0x1034) + title + axis_title + rec(0x1034)
    data = rec(0x0200, bytes(14))
    if cached:
        data += (rec(0x1065, struct.pack("<H", 1)) + rec(0x0203, struct.pack("<HHHd", 0, 0, 0, 5.0)) + rec(0x0203, struct.pack("<HHHd", 1, 0, 0, 7.0))
                 + rec(0x1065, struct.pack("<H", 2)) + rec(0x0204, struct.pack("<HHH", 0, 0, 0) + struct.pack("<HB", 2, 1) + "窓口".encode("utf-16-le"))
                 + rec(0x0204, struct.pack("<HHH", 1, 0, 0) + struct.pack("<HB", 3, 1) + "郵送等".encode("utf-16-le")) + rec(0x1065, struct.pack("<H", 3)))
    return bof(0x0020) + body + data + rec(0x000A)


def obj(ot: int, oid: int, extra: bytes = b"") -> bytes:
    return rec(0x005D, struct.pack("<HHHHH", 0x15, 0x12, ot, oid, 0x6011) + bytes(12) + extra + bytes(4))


def pict_obj(oid: int, pos: int, cls: str) -> bytes:
    """埋め込みの OLE の部品の OBJ：FtCf・FtPioGrbit・FtPictFmla（PtgTbl ＋ strClass ＋ lPosInCtlStm）。"""
    info = bytes([0x03, len(cls), 0, 0]) + cls.encode("latin-1")
    fmla = struct.pack("<H", 5) + bytes(4) + bytes([0x02]) + bytes(4) + info
    fmla += b"\0" * (len(fmla) % 2)
    body = struct.pack("<H", len(fmla)) + fmla + struct.pack("<I", pos)
    return rec(0x005D, struct.pack("<HHHHH", 0x15, 0x12, 0x08, oid, 0x6011) + bytes(12) + struct.pack("<HHH", 0x07, 2, 0xFFFF)
               + struct.pack("<HHH", 0x08, 2, 0x0001) + struct.pack("<HH", 0x09, len(body)) + body + bytes(4))


def make_xls(marker_len: int = 0, with_object: bool = True) -> bytes:
    """シート「データ」（A1:B3 の表、その上にグラフ。控えなし）とグラフのシート「グラフ」（控えあり）のブック。"""
    head = bof(0x0005) + (rec(0x0000, bytes(marker_len)) if marker_len else b"") + rec(0x0042, b"\xb0\x04")
    names = ["データ", "グラフ"]
    sst = ["区分", "件数", "窓口", "郵送等"]
    tail = rec(0x00FC, struct.pack("<II", len(sst), len(sst)) + b"".join(struct.pack("<HB", len(s), 1) + s.encode("utf-16-le") for s in sst)) + rec(0x000A)
    bs_len = sum(4 + 8 + 2 * len(n) for n in names)
    sheet = bof(0x0010) + rec(0x0200, bytes(14))
    for r, c, k in ((0, 0, 0), (0, 1, 1), (1, 0, 2), (2, 0, 3)):
        sheet += rec(0x00FD, struct.pack("<HHHI", r, c, 15, k))
    sheet += rec(0x0203, struct.pack("<HHHd", 1, 1, 15, 12)) + rec(0x0203, struct.pack("<HHHd", 2, 1, 15, 3))
    sheet += obj(0x05, 1) + biff_chart(cached=False)
    if with_object:
        sheet += pict_obj(2, 0x0001A2B3, "Word.Document.8")
    sheet += rec(0x000A)
    chart_sheet = biff_chart(cached=True)
    first = len(head) + bs_len + len(tail)
    offs = [first, first + len(sheet)]
    bs = b"".join(rec(0x0085, struct.pack("<IBBBB", o, 0, dt, len(n), 1) + n.encode("utf-16-le")) for o, dt, n in zip(offs, (0, 2), names))
    return head + bs + tail + sheet + chart_sheet


def word97(text: str) -> dict[str, bytes]:
    """小さな Word 97 の文書（本文 text の1段落）のストリーム。"""
    text += "\r"
    wd = bytearray(4608)
    struct.pack_into("<HHHHHHHI", wd, 0, 0xA5EC, 0x00C1, 0, 0x0411, 0, 0x1200, 0x00BF, 0)
    struct.pack_into("<H", wd, 0x20, 14)
    struct.pack_into("<H", wd, 0x3E, 22)
    struct.pack_into("<22i", wd, 0x40, len(wd), 0, 0, len(text), *([0] * 18))
    struct.pack_into("<H", wd, 0x98, 93)
    wd[1024:1024 + 2 * len(text)] = text.encode("utf-16-le")
    clx = b"\x02" + struct.pack("<I", 16) + struct.pack("<II", 0, len(text)) + struct.pack("<HIH", 0, 1024, 0)
    struct.pack_into("<II", wd, 0x9A + 33 * 8, 0, len(clx))
    return {"WordDocument": bytes(wd), "1Table": clx}


def test_xls_charts_and_embedded_word_document():
    from todoku.xlsx.audit import run

    inner = {"\x01CompObj": compobj_stream("Microsoft Word 文書", "Word.Document.8"), **word97("中に入れた文書")}
    blob = make_cfb({"Workbook": make_xls(), "MBD0001A2B3": inner})
    wb = XlsWorkbook(blob, "グラフ.xls")
    md = xw.to_markdown(xw.build(wb))
    # シートに置いたグラフ：控えがないので、指しているセル（同じシートの A2:A3・B2:B3）を読む
    assert "### グラフ1：受付の件数（縦棒グラフ）" in md and "縦軸（数値）の題名：件" in md
    assert "| 項目 | 件数 |" in md and "| 窓口 | 12 |" in md and "| 郵送等 | 3 |" in md
    # グラフのシート：値の控え（SIIndex の後ろの Number・Label）
    assert "## シート「グラフ」（グラフのシート）" in md and "### グラフ2：受付の件数" in md and "| 窓口 | 5 |" in md
    assert "### 埋め込み1：Word の文書（Microsoft Word 文書・Word.Document.8・MBD0001A2B3）" in md and "- 中に入れた文書" in md
    checks = {c.name: c for c in run(blob, "グラフ.xls").checks}
    assert checks["読み落とし（グラフ）"].ok and "グラフ 2 個" in checks["読み落とし（グラフ）"].detail
    assert checks["読み落とし（埋め込み）"].ok, checks["読み落とし（埋め込み）"].items


# ====================================================================== 暗号化した要約のストリーム（encryption）
def summary_stream(entries: list[tuple[str, bool, bytes]]) -> bytes:
    """解いた形の encryption ストリーム（MS-OFFCRYPTO 2.3.5.4）：見出し 8 バイト・中身・記述子の並び。"""
    data = bytearray(8)
    desc = b""
    for k, (name, is_stream, body) in enumerate(entries):
        nm = name.encode("utf-16-le")
        desc += struct.pack("<IIHBBI", len(data), len(body), k, len(name), 1 if is_stream else 0, 0) + nm + b"\0\0"
        data += body
    arr = struct.pack("<I", len(entries)) + desc
    struct.pack_into("<II", data, 0, len(data), len(arr))
    return bytes(data + arr)


def test_encrypted_xls_exposes_embeddings_in_the_encryption_stream():
    key = new_key("xls", "cryptoapi", "とどく123", salt=bytes(range(16)), verifier=bytes(range(100, 116)))
    key = replace(key, summary=True)
    inner = make_cfb({"\x01CompObj": compobj_stream("Microsoft Word 文書", "Word.Document.8"), **word97("暗号の中の文書")})
    summ = summary_stream([("MBD0001A2B3", False, inner), ("\x05SummaryInformation", True, b"S" * 100),
                           ("\x05DocumentSummaryInformation", True, b"D" * 60)])
    plain = make_cfb({"Workbook": make_xls(len(key.marker)), "encryption": summ, "\x05DocumentSummaryInformation": b"\0" * 48})
    enc = encrypt_binary(key, plain)
    assert CFB(enc).read("encryption") != summ and CFB(enc).read("encryption")[:8] == summ[:8]  # 先頭の 8 バイトは平文
    got, key2 = decrypt_binary(enc, "とどく123")
    assert got == plain and key2 == key and key2.summary  # 解き直すと元の解いた形とバイト一致、暗号化し直すと元とバイト一致
    assert encrypt_binary(key2, got) == enc
    assert [(n, s) for n, s, _d in summary_entries(CFB(got).read("encryption"))] == [
        ("MBD0001A2B3", False), ("\x05SummaryInformation", True), ("\x05DocumentSummaryInformation", True)]
    md = xw.to_markdown(xw.build(XlsWorkbook(got, "暗号.xls")))
    assert "### 埋め込み1：Word の文書" in md and "- 暗号の中の文書" in md


def test_encryption_stream_that_does_not_parse_is_left_alone():
    key = new_key("xls", "cryptoapi", "pw", salt=bytes(16), verifier=bytes(16))
    plain = make_cfb({"Workbook": make_xls(len(key.marker), with_object=False), "encryption": b"\xff" * 300})
    enc = encrypt_binary(key, plain)  # 鍵に summary がないので、encryption ストリームは変えない
    assert CFB(enc).read("encryption") == b"\xff" * 300
    got, key2 = decrypt_binary(enc, "pw")
    assert not key2.summary and got == plain


# ====================================================================== 公開の様式（corpus）
def _corpus(*pats: str) -> list[str]:
    out = []
    for p in pats:
        out += glob.glob(str(CORPUS / p), recursive=True)
    return sorted(out)


ESTAT = CORPUS / "data.e-gov.go.jp" / "a5dc5ac2e8c3fa40.bin"


@pytest.mark.corpus
def test_corpus_estat_chart():
    if not ESTAT.exists():
        pytest.skip("corpus がありません")
    from todoku.xlsx.audit import run

    blob = ESTAT.read_bytes()
    md = xw.to_markdown(xw.build(Workbook(blob, "e-stat.xlsx")))
    assert "### グラフ1：（題名なし）（積み上げ横棒グラフ、O6〜AA66）" in md
    assert "| 北海道 | 5 | 49 | 3 | 6 | 48 | 18 | 14 | 1 | 182 | 326 |  |" in md
    assert next(c for c in run(blob, "e-stat.xlsx").checks if c.name == "読み落とし（グラフ）").ok


@pytest.mark.corpus
def test_corpus_xls_ole_object_of_unknown_kind():
    p = CORPUS / "wide_xls" / "www.pref.tottori.lg.jp" / "saigai-yousiki1.xls"
    if not p.exists():
        pytest.skip("corpus がありません")
    md = xw.to_markdown(xw.build(XlsWorkbook(p.read_bytes(), p.name)))
    assert "### 埋め込み1：Vnas.Drawing（Vnas.Drawing・MBD000AB20A）" in md and "とどくでは読めない形" in md


# 生駒市の調査票（Excel のブック 9 個を埋め込んだ .docx）は窓を作るのに 1 分ほどかかるので、ここでは省く（監査の一括の検査で確かめる）
WIDE_EMBED = [p for p in _corpus("wide_embed/**/*.xlsx", "wide_embed/**/*.xls", "wide_embed/**/*.docx", "wide_embed/**/*.doc")
              if not p.endswith("ikoutyousa.docx")]


@pytest.mark.corpus
@pytest.mark.parametrize("path", WIDE_EMBED or ["(なし)"])
def test_corpus_wide_embed_reads_everything(path):
    """集めた様式（グラフ・埋め込みを含むもの）：窓とは別の読み方で、グラフ・SmartArt・埋め込みの字がすべて窓に出ている。"""
    if not WIDE_EMBED:
        pytest.skip("corpus/wide_embed がありません")
    from todoku.book import open_book
    from todoku.doc.book import DocFile
    from todoku.embeds.audit import excel_checks, word_checks
    from todoku.unlock import unlock

    blob, _lock = unlock(Path(path).read_bytes(), None)
    low = path.lower()
    if low.endswith((".docx", ".doc")):
        doc = DocxDocument(blob, "") if low.endswith(".docx") else DocFile(blob, "")
        checks = word_checks(doc, word_window(doc))
    else:
        wb = open_book(blob, "")
        checks = excel_checks(wb, xw.build(wb))
    for c in checks:
        assert c.ok, (c.name, c.items[:5])


def _embed_file(name: str) -> Path:
    p = next(iter(glob.glob(str(CORPUS / "wide_embed" / "**" / name), recursive=True)), None)
    if p is None:
        pytest.skip("corpus/wide_embed がありません")
    return Path(p)


@pytest.mark.corpus
def test_corpus_chart_kinds_in_real_files():
    from todoku.api import window

    # 静岡市の .docx：テキストボックスの中のグラフも1回だけ数える（7 個）
    _k, doc, _p = window(str(_embed_file("prtrgaiyou2016.docx")))
    assert sum(1 for o in doc["objects"] if o["kind"] == "chart") == 7
    # 三田市の .xlsx：グラフのシート 4 枚。フィルターで隠した系列も落とさない
    _k, doc, parts = window(str(_embed_file("0926kansensyagraf.xlsx")))
    md = "\n".join(t for _n, t in parts)
    assert md.count("（グラフのシート）") == 4 and "兵庫県感染者数" in md and "フィルターで非表示にした系列" in md
    # 総務省の .xls（書き込みパスワードだけの暗号）：シートに置いた BIFF のグラフ 5 個。控えのエラーは #N/A
    _k, doc, parts = window(str(_embed_file("31.xls")))
    assert sum(len(sh.get("charts") or []) for sh in doc["sheets"]) == 5
    md = "\n".join(t for _n, t in parts)
    assert "| H18 | 116882 | 134452 |" in md and "#N/A" in md
    # 草津市の .doc：Excel のグラフ（Excel.Chart.8）を埋め込んだもの。グラフのシートの散布図とデータラベル
    _k, doc, parts = window(str(_embed_file("1844.doc")))
    md = "\n".join(t for _n, t in parts)
    assert "#### グラフ1：行財政マネジメントのための基本方針（散布図）" in md and "「市民から信頼される市政運営」" in md


@pytest.mark.corpus
def test_corpus_embedded_office_files():
    from todoku.api import window

    _k, doc, parts = window(str(_embed_file("word_with_embeded.doc")))
    md = "\n".join(t for _n, t in parts)
    assert "スライド1：" in md and "Sample PowerPoint file" in md
    _k, doc, parts = window(str(_embed_file("excel_with_embeded.xls")))
    kinds = sorted(o["kind"] for sh in doc["sheets"] for o in sh.get("objects") or [])
    assert kinds == ["excel", "excel", "powerpoint", "powerpoint", "word", "word"]
    _k, doc, parts = window(str(_embed_file("WithEmbeded.xlsx")))
    md = "\n".join(t for _n, t in parts)
    assert "ZIP の書庫の中のファイル" in md
