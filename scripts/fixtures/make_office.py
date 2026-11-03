"""テスト用の Office・ODF・RTF・HWPX・一太郎・文字だけのファイル・ZIP（架空の様式。個人情報なし）を作る。

- docx・xlsx・文字だけのファイル・ZIP は、標準ライブラリ（xlsx は openpyxl）で、ここで組み立てる（ZIP の時刻は固定）。
- doc・odt・rtf・xls・ods は、LibreOffice（soffice --headless --convert-to）で docx・xlsx から変換する
  （本物の書き出し側が作ったファイルになる）。soffice がないときは、それらを作らない。
- pptx・hwpx・jtd は、テストの中の組み立て（tests/pptx_maker.py・test_hwpx.py・test_legacy.py）を使う。
"""

from __future__ import annotations

import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

TESTS = Path(__file__).resolve().parents[2] / "tests"
FIXED = (2026, 1, 1, 0, 0, 0)


def zip_bytes(files: dict[str, bytes | str], stored: tuple[str, ...] = ()) -> bytes:
    """時刻を固定した ZIP（同じ中身なら同じバイト）。"""
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w") as z:
        for name, data in files.items():
            zi = zipfile.ZipInfo(name, FIXED)
            zi.compress_type = zipfile.ZIP_STORED if name in stored else zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            z.writestr(zi, data.encode("utf-8") if isinstance(data, str) else data)
    return bio.getvalue()


def rezip(blob: bytes) -> bytes:
    """ほかの組み立てが作った ZIP を、順番はそのままに時刻だけ固定して詰め直す。"""
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        files = {i.filename: z.read(i) for i in z.infolist()}
        stored = tuple(i.filename for i in z.infolist() if i.compress_type == zipfile.ZIP_STORED)
    return zip_bytes(files, stored)


# ====================================================================== Word（.docx）
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="w14"')
XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def _r(text: str, rpr: str = "") -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _p(*inner: str, ppr: str = "") -> str:
    return f"<w:p>{ppr}{''.join(inner)}</w:p>"


def _tc(inner: str, w: int = 2400) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{inner}</w:tc>'


def _sdt_text(alias: str, placeholder: str) -> str:
    grey = _p(_r(placeholder, '<w:rPr><w:color w:val="808080"/></w:rPr>'))
    return (f'<w:sdt><w:sdtPr><w:alias w:val="{alias}"/><w:tag w:val="{alias}"/><w:showingPlcHdr/><w:text/></w:sdtPr>'
            f'<w:sdtContent>{grey}</w:sdtContent></w:sdt>')


def _sdt_drop(alias: str, items: list[str]) -> str:
    li = "".join(f'<w:listItem w:displayText="{i}" w:value="{i}"/>' for i in items)
    return (f'<w:sdt><w:sdtPr><w:alias w:val="{alias}"/><w:tag w:val="{alias}"/><w:showingPlcHdr/>'
            f'<w:dropDownList><w:listItem w:displayText="選択してください" w:value=""/>{li}</w:dropDownList></w:sdtPr>'
            f'<w:sdtContent>{_p(_r("選択してください"))}</w:sdtContent></w:sdt>')


def _sdt_check(alias: str) -> str:
    return (f'<w:sdt><w:sdtPr><w:alias w:val="{alias}"/><w:tag w:val="{alias}"/><w14:checkbox><w14:checked w14:val="0"/>'
            '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
            '</w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>')


def _formtext(name: str, maxlen: int) -> str:
    return (f'<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="{name}"/><w:enabled/><w:calcOnExit w:val="0"/>'
            f'<w:textInput><w:maxLength w:val="{maxlen}"/></w:textInput></w:ffData></w:fldChar></w:r>'
            '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>     </w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')


def _formcheck(name: str) -> str:
    return (f'<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="{name}"/><w:enabled/><w:calcOnExit w:val="0"/>'
            '<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r>'
            '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')


def _textbox(text: str) -> str:
    inner = f'<w:txbxContent>{_p(_r(text))}</w:txbxContent>'
    return ('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" '
            'simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>3600000</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="1800000" cy="360000"/><wp:wrapNone/><wp:docPr id="1" name="テキスト ボックス 1"/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="360000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'</wps:spPr><wps:txbx>{inner}</wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'
            '</mc:Choice><mc:Fallback><w:pict><v:shape style="position:absolute;width:142pt;height:28pt"><v:textbox>'
            f'{inner}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>')


DOCX_TEXTS = ["届出書", "氏名", "区分", "電話番号", "受付番号", "提出先", "市役所"]


def docx() -> bytes:
    """内容のコントロール（字・一覧・チェック）、古いフォームの欄（FORMTEXT・FORMCHECKBOX）、テキストボックス、
    変更の記録（挿入と削除）、表のある、架空の届出書。"""
    jc = '<w:pPr><w:jc w:val="center"/></w:pPr>'
    rows = [("氏名", _sdt_text("氏名", "ここに氏名を入力")), ("区分", _sdt_drop("区分", ["新規", "変更", "廃止"])),
            ("同意", _p(_sdt_check("同意"), _r(" 内容に同意します")))]
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblBorders>'
           + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in
                     ("top", "left", "bottom", "right", "insideH", "insideV"))
           + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="4800"/></w:tblGrid>')
    for label, cell in rows:
        tbl += "<w:tr>" + _tc(_p(_r(label))) + _tc(cell, 4800) + "</w:tr>"
    tbl += "</w:tbl>"
    body = (_p(_r("届出書", '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'), ppr=jc)
            + _p(_r("受付欄"), _textbox("受付番号"))
            + tbl
            + _p(_r("電話番号："), _formtext("Tel", 13))
            + _p(_formcheck("Mail"), _r(" 結果を郵送で受け取る"))
            + _p(_r("提出先："),
                 '<w:del w:id="1" w:author="todoku" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>町役場</w:delText></w:r></w:del>',
                 '<w:ins w:id="2" w:author="todoku" w:date="2026-01-01T00:00:00Z">' + _r("市役所") + "</w:ins>")
            + _p(_r("この様式は、とどくのテストのための架空のものです。"))
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
              'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    styles = (XMLDECL + f'<w:styles {W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho"/>'
              '<w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    return zip_bytes({
        "[Content_Types].xml": XMLDECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '</Types>',
        "_rels/.rels": XMLDECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": XMLDECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/></Relationships>',
        "word/document.xml": XMLDECL + f"<w:document {W_NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": styles,
        "word/settings.xml": XMLDECL + f'<w:settings {W_NS}><w:trackRevisions/><w:defaultTabStop w:val="840"/></w:settings>',
    })


# ====================================================================== Excel（.xlsx）
ITEMS = [("用紙", 3, 1200), ("封筒", 10, 85), ("切手", 4, 110)]


def xlsx() -> bytes:
    """セルの結合・入力規則（一覧）・数式（計算結果つき）・書式つきの字（リッチテキスト）・1904 年の日付の方式・
    シートの保護（入力のセルだけ鍵を外す）のある、架空の利用申請書。"""
    from openpyxl import Workbook
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.styles import Border, Protection, Side
    from openpyxl.utils.datetime import CALENDAR_MAC_1904
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    wb.epoch = CALENDAR_MAC_1904
    ws = wb.active
    ws.title = "申請書"
    thin = Side(style="thin")
    edge = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells("A1:D1")
    ws["A1"] = "利用申請書"
    ws["A3"], ws["A4"], ws["A5"] = "氏名", "区分", "利用開始日"
    ws.merge_cells("B3:D3")
    for ref in ("B3", "B4", "B5"):
        ws[ref].protection = Protection(locked=False)
        ws[ref].border = edge
    dv = DataValidation(type="list", formula1='"個人,法人"', allow_blank=True)
    dv.add("B4")
    ws.add_data_validation(dv)
    ws["B5"] = date(2026, 4, 1)
    ws["B5"].number_format = "yyyy/m/d"
    ws["A6"] = CellRichText(TextBlock(InlineFont(b=True, color="FF0000"), "※"), "太枠の中を記入してください")
    ws["A8"], ws["B8"], ws["C8"], ws["D8"] = "品目", "数量", "単価", "金額"
    for i, (name, n, price) in enumerate(ITEMS, start=9):
        ws.cell(i, 1, name)
        ws.cell(i, 2, n)
        ws.cell(i, 3, price)
        ws.cell(i, 4, f"=B{i}*C{i}")
    last = 8 + len(ITEMS)
    ws.cell(last + 1, 3, "合計")
    ws.cell(last + 1, 4, f"=SUM(D9:D{last})")
    ws.protection.sheet = True
    bio = io.BytesIO()
    wb.save(bio)
    # 計算結果（Excel が保存するときに書く <v>）を、ここで計算した値で入れる
    values = {f"D{i}": n * price for i, (_name, n, price) in enumerate(ITEMS, start=9)}
    values[f"D{last + 1}"] = sum(values.values())
    with zipfile.ZipFile(io.BytesIO(bio.getvalue())) as z:
        files = {i.filename: z.read(i) for i in z.infolist()}
    sheet = files["xl/worksheets/sheet1.xml"].decode()

    def put(m: re.Match) -> str:
        ref = m.group(1)
        return f'<c r="{ref}"{m.group(2)}><f>{m.group(3)}</f><v>{values[ref]}</v></c>'

    sheet = re.sub(r'<c r="(D\d+)"([^>]*)><f>([^<]*)</f><v\s*/?>(?:</v>)?</c>', put, sheet)
    assert sheet.count("<v>") >= len(values), "数式の計算結果を入れられなかった"
    files["xl/worksheets/sheet1.xml"] = sheet.encode()
    return zip_bytes(files)


XLSX_TOTAL = sum(n * p for _name, n, p in ITEMS)


# ====================================================================== そのほか
def pptx() -> bytes:
    sys.path.insert(0, str(TESTS))
    from pptx_maker import make, para, sp, table

    s1 = sp(para('<a:r><a:rPr lang="ja-JP"/><a:t>説明会の参加申込</a:t></a:r>'), "タイトル 1", "title") \
        + table([["氏名", ""], ["所属", ""], ["参加日", "□ 10日　□ 11日"]])
    return rezip(make([s1], notes={1: sp(para('<a:r><a:t>この申込は架空のものです</a:t></a:r>'), "ノート", "body", 3)}))


def hwpx() -> bytes:
    sys.path.insert(0, str(TESTS))
    from test_hwpx import make_hwpx

    return rezip(make_hwpx())


def jtd() -> bytes:
    sys.path.insert(0, str(TESTS))
    from test_legacy import make_jtd_full

    return make_jtd_full()


TXT = "申請書の書き方\n\n1. 氏名を記入する\n2. 住所を記入する\n\n氏名：\n住所：\n"
CSV = "番号,品目,数量,備考\n1,用紙,3,\n2,封筒,10,\"窓あき, 長形3号\"\n"
JSON = {"様式": "参加申込", "項目": [{"見出し": "氏名", "値": ""}, {"見出し": "電話番号", "値": ""}]}


def texts() -> dict[str, bytes]:
    return {"notes.txt": TXT.encode("utf-8"), "items.csv": CSV.encode("utf-8"),
            "items_sjis.csv": CSV.encode("cp932"),
            "form.json": json.dumps(JSON, ensure_ascii=False, indent=1).encode("utf-8")}


def bundle(members: dict[str, bytes]) -> bytes:
    """いくつかの様式をまとめた ZIP（名前は UTF-8）。"""
    return zip_bytes(members)


def soffice_convert(src_name: str, blob: bytes, ext: str, flt: str | None = None) -> bytes | None:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if exe is None:
        return None
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / src_name
        src.write_bytes(blob)
        target = f"{ext}:{flt}" if flt else ext
        subprocess.run([exe, f"-env:UserInstallation=file://{td}/profile", "--headless", "--convert-to", target,
                        "--outdir", td, str(src)], check=True, capture_output=True, timeout=180)
        out = src.with_suffix("." + ext)
        return out.read_bytes() if out.exists() else None
