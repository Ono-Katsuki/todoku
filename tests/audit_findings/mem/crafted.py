"""監査用：宣言された数で大きなメモリ・長い時間を引き出す、手で作った書類（1 属性の変異では届かない組み合わせ）。

使い方：crafted.py 出力フォルダ [名前の一部 ...]
各ケースを作って guard（nice 10・1.5 GB・60 秒）で全部の道具にかけ、results_crafted.jsonl に足す。
環境変数なしで作るだけなら CRAFT_ONLY=1。
"""
import io
import json
import os
import re
import struct
import sys
import zipfile
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))
from guard import guarded  # noqa: E402

PUB = os.path.join(ROOT, "tests", "fixtures", "public")
CASES = {}


def case(name, tools=None, env=None):
    def deco(fn):
        CASES[name] = (fn, tools, env or {})
        return fn
    return deco


def rezip(blob, changes, add=None):
    src = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for i in src.infolist():
            ni = zipfile.ZipInfo(i.filename, i.date_time)
            ni.compress_type = zipfile.ZIP_DEFLATED if i.compress_type else zipfile.ZIP_STORED
            z.writestr(ni, changes.get(i.filename, src.read(i.filename)))
        for k, v in (add or {}).items():
            z.writestr(zipfile.ZipInfo(k), v, compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()


# ====================================================================== xlsx
XL = open(os.path.join(PUB, "sheet.xlsx"), "rb").read()
SH = "xl/worksheets/sheet1.xml"
X0 = zipfile.ZipFile(io.BytesIO(XL)).read(SH).decode()
HEAD = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
# 保護を外した、入力 B3 と数式だけの小さなシート（順序は CT_Worksheet のとおり）


def sheet(cells="", tail="", cols="", pre=""):
    return (HEAD + pre + cols + '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>品目</t></is></c><c r="B1" t="inlineStr"><is><t>氏名</t></is></c></row>'
            + '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>' + cells + "</sheetData>" + tail + "</worksheet>")


def xlsx(xml, changes=None, add=None, ct_extra=""):
    ch = {SH: xml.encode() if isinstance(xml, str) else xml}
    ch.update(changes or {})
    if ct_extra:
        ct = zipfile.ZipFile(io.BytesIO(XL)).read("[Content_Types].xml").decode().replace("</Types>", ct_extra + "</Types>")
        ch["[Content_Types].xml"] = ct.encode()
    return rezip(XL, ch, add)


def fcell(formula, ref="C3"):
    return f'<row r="3"><c r="{ref}"><f>{formula}</f></c></row>'


@case("xlsx_merge_all")
def _():
    return xlsx(sheet(tail='<mergeCells count="1"><mergeCell ref="D5:XFD1048576"/></mergeCells>'))


@case("xlsx_merge_many_tall")
def _():
    m = "".join(f'<mergeCell ref="{chr(68 + (i % 20))}{5 + i}:{chr(68 + (i % 20))}1048576"/>' for i in range(3000))
    return xlsx(sheet(tail=f"<mergeCells>{m}</mergeCells>"))


@case("xlsx_validation_all")
def _():
    return xlsx(sheet(tail='<dataValidations count="1"><dataValidation type="whole" sqref="A1:XFD1048576"><formula1>1</formula1></dataValidation></dataValidations>'))


@case("xlsx_merge_plus_validation")
def _():
    return xlsx(sheet(tail='<mergeCells count="1"><mergeCell ref="D5:XFD1048576"/></mergeCells>'
                           '<dataValidations count="1"><dataValidation type="whole" sqref="A1:XFD1048576"><formula1>1</formula1></dataValidation></dataValidations>'))


@case("xlsx_lastcell_plus_validation")
def _():
    # 右下の端のセルが 1 つあるだけで「使っている範囲」は全シートになる
    return xlsx(sheet(cells='<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>',
                      tail='<dataValidations count="1"><dataValidation type="whole" sqref="A1:XFD1048576"><formula1>1</formula1></dataValidation></dataValidations>'))


@case("xlsx_lastcell_plus_formula_all")
def _():
    return xlsx(sheet(cells='<row r="3"><c r="C3"><f>SUM(A:XFD)</f></c></row><row r="1048576"><c r="XFD1048576"><v>1</v></c></row>'))


@case("xlsx_merge_plus_formula_all")
def _():
    return xlsx(sheet(cells=fcell("SUM(D5:XFD1048576)"), tail='<mergeCells count="1"><mergeCell ref="D5:XFD1048576"/></mergeCells>'))


@case("xlsx_array_ref_all")
def _():
    return xlsx(sheet(cells='<row r="3"><c r="C3"><f t="array" ref="C3:XFD1048576">A1</f></c></row>'))


@case("xlsx_shared_ref_all")
def _():
    return xlsx(sheet(cells='<row r="3"><c r="C3"><f t="shared" ref="C3:XFD1048576" si="0">A1</f></c><c r="D3"><f t="shared" si="0"/></c></row>'))


@case("xlsx_cols_many")
def _():
    c = "".join(f'<col min="{1 + 16384 * i}" max="{16384 + 16384 * i}" width="9" hidden="1" style="1"/>' for i in range(6000))
    return xlsx(sheet(cols=f"<cols>{c}</cols>"))


@case("xlsx_print_area_all")
def _():
    wbx = zipfile.ZipFile(io.BytesIO(XL)).read("xl/workbook.xml").decode()
    name = re.search(r'<sheet [^>]*name="([^"]+)"', wbx)[1]
    dn = f'<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">\'{name}\'!$A$1:$XFD$1048576</definedName></definedNames>'
    return xlsx(sheet(tail='<dataValidations count="1"><dataValidation type="whole" sqref="A1:XFD1048576"><formula1>1</formula1></dataValidation></dataValidations>'),
                {"xl/workbook.xml": wbx.replace("</sheets>", "</sheets>" + dn).encode()})


@case("xlsx_table_all")
def _():
    t = ('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:XFD1048576" totalsRowShown="0">'
         '<tableColumns count="2"><tableColumn id="1" name="品目"/><tableColumn id="2" name="氏名"/></tableColumns></table>')
    rel = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
           'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>')
    return xlsx(sheet(tail='<tableParts count="1"><tablePart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></tableParts>'),
                add={"xl/tables/table1.xml": t, "xl/worksheets/_rels/sheet1.xml.rels": rel},
                ct_extra='<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>')


@case("xlsx_control_anchor_all")
def _():
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ctl = (f'<controls xmlns:r="{R}" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><control shapeId="1025" r:id="rId1" name="c1"><controlPr>'
           '<anchor><from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>'
           '<to><xdr:col>16383</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1048575</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control></controls>')
    rel = (f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>')
    pr = '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox"/>'
    return xlsx(sheet(tail=ctl), add={"xl/ctrlProps/ctrlProp1.xml": pr, "xl/worksheets/_rels/sheet1.xml.rels": rel})


@case("xlsx_cf_all")
def _():
    return xlsx(sheet(tail='<conditionalFormatting sqref="A1:XFD1048576"><cfRule type="cellIs" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>'))


@case("xlsx_hyperlink_autofilter_all")
def _():
    return xlsx(sheet(tail='<autoFilter ref="A1:XFD1048576"/><hyperlinks><hyperlink ref="A1:XFD1048576" location="A1"/></hyperlinks>'))


@case("xlsx_sqref_many_all")
def _():
    sq = " ".join(f"A{i}:A{i + 63}" for i in range(1, 60000, 1))
    return xlsx(sheet(tail=f'<dataValidations count="1"><dataValidation type="whole" sqref="{sq}"><formula1>1</formula1></dataValidation></dataValidations>'))


@case("xlsx_formula_deep_parens")
def _():
    return xlsx(sheet(cells=fcell("(" * 100000 + "1" + ")" * 100000)))


@case("xlsx_formula_chain_deep")
def _():
    rows = "".join(f'<row r="{i}"><c r="C{i}"><f>C{i + 1}+1</f></c></row>' for i in range(3, 200000))
    return xlsx(sheet(cells=rows))


# 式：計算（試算・書き込みの再計算）で評価される。入力 B2 → C3 の式
FORMULAS = {
    "sequence": "SUM(SEQUENCE(1048576,16384))",
    "sequence_col": "SUM(SEQUENCE(1048576))",
    "rept": 'LEN(REPT("x",4000000000))',
    "rept_rept": 'LEN(REPT(REPT("x",1000000),5000))',
    "expand": "SUM(EXPAND(A1,1048576,16384,0))",
    "munit": "SUM(MUNIT(30000))",
    "makearray": "SUM(RANDARRAY(1048576,16384))",
    "sumproduct_col": "SUMPRODUCT(A:A,B:B)",
    "sumproduct_all": "SUMPRODUCT(--(A:XFD=1))",
    "sum_all": "SUM(A:XFD)",
    "countif_all": 'COUNTIF(A:XFD,"")',
    "countblank_all": "COUNTBLANK(A:XFD)",
    "vlookup_all": "VLOOKUP(B2,A:XFD,2,FALSE)",
    "match_col": "MATCH(9,A:A,0)",
    "index_all": "SUM(INDEX(A:XFD,0,0))",
    "indirect_all": 'SUM(INDIRECT("A1:XFD1048576"))',
    "indirect_col": 'SUMPRODUCT(INDIRECT("A:A")*1)',
    "offset_all": "SUM(OFFSET(A1,0,0,1048576,16384))",
    "row_col": "SUMPRODUCT(ROW(A:A))",
    "row_rows": "SUMPRODUCT(ROW(1:1048576)*COLUMN(A:XFD))",
    "column_all": "SUMPRODUCT(COLUMN(A:XFD))",
    "textjoin_all": 'LEN(TEXTJOIN("",FALSE,A:XFD))',
    "concat_all": "LEN(CONCAT(A:XFD))",
    "sort_all": "SUM(SORT(A:XFD))",
    "unique_all": "SUM(UNIQUE(A:XFD))",
    "filter_all": "SUM(FILTER(A:XFD,A:A=1))",
    "transpose_all": "SUM(TRANSPOSE(A:XFD))",
    "mmult": "SUM(MMULT(A1:C5000,TRANSPOSE(A1:C5000)))",
    "mmult2": "SUM(MMULT(SEQUENCE(4000,4000),SEQUENCE(4000,4000)))",
    "subtotal_all": "SUBTOTAL(9,A:XFD)",
    "aggregate_all": "AGGREGATE(9,6,A:XFD)",
    "networkdays": "NETWORKDAYS(1,2958465)",
    "workday": "WORKDAY(1,2000000)",
    "substitute": 'LEN(SUBSTITUTE(REPT("a",100000),"a",REPT("b",100000)))',
    "lambda_like_if": "SUM(IF(A:XFD=1,1,0))",
    "phonetic_all": "LEN(PHONETIC(A:XFD))",
    "add_all_all": "SUM(A:XFD+A:XFD)",
    "unary_all": "SUM(-A:XFD)",
    "wrap": "SUM(WRAPROWS(SEQUENCE(1048576),1))",
    "tocol": "SUM(TOCOL(A:XFD))",
    "hstack": "SUM(HSTACK(A:XFD,A:XFD))",
    "vstack": "SUM(VSTACK(A:XFD,A:XFD))",
    "take_all": "SUM(TAKE(A:XFD,1048576))",
    "xlookup_all": "XLOOKUP(9,A:A,A:A)",
    "large_all": "LARGE(A:XFD,1)",
    "median_all": "MEDIAN(A:XFD)",
    "rank_all": "RANK(1,A:XFD)",
    "frequency_all": "SUM(FREQUENCY(A:XFD,A:XFD))",
}
for _k, _f in FORMULAS.items():
    def _make(_f=_f):
        return lambda: xlsx(sheet(cells=fcell(_f.replace("&", "&amp;").replace('"', "&quot;")),
                                  pre=""))
    CASES["xlsxf_" + _k] = (_make(), ["summary", "audit", "calc_model", "calculate", "fill"], {"PROBE_CELL": "B2"})


# ====================================================================== ODS
OD = open(os.path.join(PUB, "sheet.ods"), "rb").read()
C0 = zipfile.ZipFile(io.BytesIO(OD)).read("content.xml").decode()
TB = re.search(r"<table:table [^>]*>.*?</table:table>", C0, re.S)


def ods(body):
    head = re.match(r"<table:table [^>]*>", TB[0])[0].replace('table:protected="true"', "")
    return rezip(OD, {"content.xml": (C0[:TB.start()] + head + body + "</table:table>" + C0[TB.end():]).encode()})


CELL = '<table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell>'
RCELL = '<table:table-cell office:value-type="string" table:number-columns-repeated="{n}"><text:p>x</text:p></table:table-cell>'


@case("ods_rows_cols_content")
def _():
    return ods(f'<table:table-column/><table:table-row table:number-rows-repeated="1048576">{RCELL.format(n=16384)}</table:table-row>')


@case("ods_rows_cols_content_many")
def _():
    rows = "".join(f'<table:table-row table:number-rows-repeated="10000">{RCELL.format(n=1024)}</table:table-row>' for _ in range(100))
    return ods("<table:table-column/>" + rows)


@case("ods_rows_big_styled")
def _():
    # 中身のない・書式つきのセル：n<=2000 行 × 列（BIG まで）
    st = re.search(r'table:style-name="(ce\d+)"', C0)[1]
    rows = "".join(f'<table:table-row table:number-rows-repeated="2000"><table:table-cell table:style-name="{st}" table:number-columns-repeated="1000"/></table:table-row>' for _ in range(50))
    return ods("<table:table-column/>" + rows)


@case("ods_span_all")
def _():
    return ods('<table:table-column/><table:table-row><table:table-cell table:number-columns-spanned="16384" table:number-rows-spanned="1048576" office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row>')


@case("ods_span_huge_plus_validation")
def _():
    # 入力規則つきの巨大な繰り返し＋巨大な結合
    return ods('<table:table-column table:number-columns-repeated="16384"/><table:table-row table:number-rows-repeated="1048576"><table:table-cell table:content-validation-name="val1" table:number-columns-repeated="16384" table:number-columns-spanned="2147483647" table:number-rows-spanned="2147483647"/></table:table-row>')


@case("ods_column_hidden_many")
def _():
    return ods('<table:table-column table:visibility="collapse" table:number-columns-repeated="1000"/>' * 20000 + f"<table:table-row>{CELL}</table:table-row>")


# ====================================================================== ODT
ODT = open(os.path.join(PUB, "form.odt"), "rb").read()
T0 = zipfile.ZipFile(io.BytesIO(ODT)).read("content.xml").decode()


def odt(tbl):
    return rezip(ODT, {"content.xml": T0.replace("</office:text>", tbl + "</office:text>").encode()})


@case("odt_rows_cols_amp")
def _():
    row = '<table:table-row table:number-rows-repeated="50"><table:table-cell table:number-columns-repeated="200"><text:p>x</text:p></table:table-cell></table:table-row>'
    return odt('<table:table table:name="T9"><table:table-column table:number-columns-repeated="200"/>' + row * 500 + "</table:table>")


@case("odt_span_all")
def _():
    return odt('<table:table table:name="T9"><table:table-column/><table:table-row><table:table-cell table:number-columns-spanned="2147483647" table:number-rows-spanned="2147483647"><text:p>x</text:p></table:table-cell></table:table-row></table:table>')


# ====================================================================== Word
DX = open(os.path.join(PUB, "form.docx"), "rb").read()
D0 = zipfile.ZipFile(io.BytesIO(DX)).read("word/document.xml").decode()


def docx(body):
    return rezip(DX, {"word/document.xml": D0.replace("<w:body>", "<w:body>" + body).encode()})


TC = '<w:tc><w:tcPr>{}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>'


@case("docx_gridspan")
def _():
    return docx(f'<w:tbl><w:tr>{TC.format("""<w:gridSpan w:val="2147483647"/>""")}{TC.format("")}</w:tr><w:tr>{TC.format("")}{TC.format("")}</w:tr></w:tbl>')


@case("docx_gridspan_many")
def _():
    rows = "".join(f'<w:tr>{TC.format("""<w:gridSpan w:val="1000000"/>""")}{TC.format("""<w:vMerge w:val="restart"/>""")}</w:tr>' for _ in range(5000))
    return docx(f"<w:tbl>{rows}</w:tbl>")


@case("docx_grid_cols_many")
def _():
    return docx("<w:tbl><w:tblGrid>" + '<w:gridCol w:w="1"/>' * 500000 + f"</w:tblGrid><w:tr>{TC.format('')}</w:tr></w:tbl>")


@case("docx_wide_table")
def _():
    # 行 2000 × 列 500（列ごとの計算が行 × 列 × 何か、にならないか）
    row = "<w:tr>" + TC.format("") * 500 + "</w:tr>"
    return docx("<w:tbl>" + row * 2000 + "</w:tbl>")


@case("docx_vmerge_chain")
def _():
    rows = f'<w:tr>{TC.format("""<w:vMerge w:val="restart"/>""")}</w:tr>' + f'<w:tr>{TC.format("<w:vMerge/>")}</w:tr>' * 100000
    return docx(f"<w:tbl>{rows}</w:tbl>")


@case("docx_nested_tables")
def _():
    s = "<w:p/>"
    for _ in range(3000):
        s = f"<w:tbl><w:tr><w:tc>{s}</w:tc></w:tr></w:tbl><w:p/>"
    return docx(s)


@case("docx_many_runs_para")
def _():
    # 1 段落に 20 万の字のかたまり（段落の中で 2 乗にならないか）
    return docx("<w:p>" + "<w:r><w:t>あ</w:t></w:r>" * 200000 + "</w:p>")


@case("docx_many_blanks_para")
def _():
    return docx("<w:p>" + "<w:r><w:t>氏名＿＿＿（　　）</w:t></w:r>" * 30000 + "</w:p>")


# ====================================================================== HWPX
HX = open(os.path.join(PUB, "form.hwpx"), "rb").read()
H0 = zipfile.ZipFile(io.BytesIO(HX)).read("Contents/section0.xml").decode()


@case("hwpx_rowspan")
def _():
    return rezip(HX, {"Contents/section0.xml": re.sub(r'rowSpan="\d+"', 'rowSpan="100000000"', H0, count=1).encode()})


@case("hwpx_colspan")
def _():
    return rezip(HX, {"Contents/section0.xml": re.sub(r'colSpan="\d+"', 'colSpan="2147483647"', H0).encode()})


@case("hwpx_addr")
def _():
    return rezip(HX, {"Contents/section0.xml": re.sub(r'(colAddr|rowAddr)="\d+"', r'\1="2147483647"', H0).encode()})


# ====================================================================== RTF
RTF = open(os.path.join(PUB, "form.rtf"), "rb").read()


@case("rtf_deep_groups")
def _():
    return RTF[:-1] + b"{\\pard " * 300000 + b"x" + b"}" * 300000 + b"}"


@case("rtf_cellx_many")
def _():
    cols = b"".join(b"\\cellx%d" % (i * 10) for i in range(1, 3000))
    row = b"\\trowd" + cols + b"\\intbl x\\cell\\row\n"
    return RTF[:-1] + row * 2000 + b"}"


@case("rtf_cellx_vary")
def _():
    # 行ごとに違う \cellx：格子の列の数が 行の数 × 列の数 になる
    rows = bytearray()
    for r in range(3000):
        rows += b"\\trowd" + b"".join(b"\\cellx%d" % (100000 + r * 300 + i) for i in range(1, 200)) + b"\\intbl x\\cell\\row\n"
    return RTF[:-1] + bytes(rows) + b"}"


@case("rtf_hexbin")
def _():
    return RTF[:-1] + b"{\\pict\\jpegblip\\picw100000000\\pich100000000\\bin2147483647 " + b"\xff" * 1000 + b"}}"


# ====================================================================== テキスト・データ
@case("csv_one_quote")
def _():
    return b'a,b\n"' + b"x," * 2000000 + b"\n"


@case("csv_wide")
def _():
    return (b"," * 200000 + b"\n") * 100


@case("csv_many_empty_lines")
def _():
    return b"a,b\n" + b"\n" * 5000000


@case("json_deep")
def _():
    return b'{"a":' * 200000 + b"1" + b"}" * 200000


@case("json_deep_list")
def _():
    return b"[" * 1000000 + b"]" * 1000000


@case("txt_long_line")
def _():
    return "氏名：＿＿＿（　　）".encode() * 1000000


@case("txt_blanks_many")
def _():
    return ("氏名：\n" * 1000000).encode()


# ====================================================================== zip
@case("zip_bomb_csv")
def _():
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("a.csv", b"," * (600 * 1048576))
    return out.getvalue()


@case("zip_many_members")
def _():
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_STORED) as z:
        for i in range(200000):
            z.writestr(f"{i}.txt", b"")
    return out.getvalue()


@case("xlsx_bomb_sheet")
def _():
    return xlsx(sheet(cells="".join(f'<row r="{i}"/>' for i in range(3, 1000000))))


@case("xlsx_bomb_comment_xml")
def _():
    # 部品が 256 MB に近い（XML の木に開くと何倍か）
    pad = "<!--" + "x" * (240 * 1048576) + "-->"
    return xlsx(sheet(cells=pad))


@case("xlsx_attr_bomb_tree")
def _():
    # 空のセル要素だけ 3000 万字
    row = "<c/>" * 16384
    return xlsx(sheet(cells="".join(f'<row r="{i}">{row}</row>' for i in range(3, 700))))


# ====================================================================== PDF
def pdf(objs, root=1, trailer=b""):
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = {}
    for n in sorted(objs):
        offs[n] = len(out)
        out += b"%d 0 obj\n" % n + objs[n] + b"\nendobj\n"
    x = len(out)
    size = max(objs) + 1
    out += b"xref\n0 %d\n0000000000 65535 f \n" % size
    for n in range(1, size):
        out += (b"%010d 00000 n \n" % offs[n]) if n in offs else b"0000000000 00000 f \n"
    out += b"trailer\n<< /Size %d /Root %d 0 R %s >>\nstartxref\n%d\n%%%%EOF\n" % (size, root, trailer, x)
    return bytes(out)


def stream(d, data):
    return b"<< " + d + b" /Length %d >>\nstream\n" % len(data) + data + b"\nendstream"


HELV = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"


def page(content=b"BT /F1 12 Tf 72 720 Td (hello) Tj ET", res=b"/Font << /F1 5 0 R >>", extra=b"", box=b"[0 0 595 842]", more=None, cat=b"", cfilter=b""):
    o = {1: b"<< /Type /Catalog /Pages 2 0 R " + cat + b" >>",
         2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
         3: b"<< /Type /Page /Parent 2 0 R /MediaBox " + box + b" /Contents 4 0 R /Resources << " + res + b" >> " + extra + b" >>",
         4: stream(cfilter, content), 5: HELV}
    o.update(more or {})
    return pdf(o)


@case("pdf_mediabox_1e9")
def _():
    return page(box=b"[0 0 1000000000 1000000000]")


@case("pdf_mediabox_thin")
def _():
    return page(box=b"[0 0 1000000000 0.001]")


@case("pdf_userunit")
def _():
    return page(extra=b"/UserUnit 1000000")


@case("pdf_dpi_huge", env={"PROBE_DPI": "100000"})
def _():
    return page()


@case("pdf_xrefstm_index")
def _():
    # xref ストリーム：/Index の数は 20 億、中身は数行だけ
    objs = {1: b"<< /Type /Catalog /Pages 2 0 R >>", 2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>"}
    out = bytearray(b"%PDF-1.7\n")
    offs = {}
    for n, b in objs.items():
        offs[n] = len(out)
        out += b"%d 0 obj\n" % n + b + b"\nendobj\n"
    x = len(out)
    rows = b"\0\0\0\0" + b"".join(b"\x01" + struct.pack(">H", offs[n]) + b"\0" for n in objs)
    out += b"4 0 obj\n" + stream(b"/Type /XRef /W [1 2 1] /Size 2000000000 /Index [0 2000000000] /Root 1 0 R", rows) + b"\nendobj\n"
    out += b"startxref\n%d\n%%%%EOF\n" % x
    return bytes(out)


@case("pdf_objstm_n")
def _():
    # ObjStm の /N が 20 億（中身は短い）。ページの字のオブジェクトが ObjStm の中
    body = b"6 0 "
    objs = {1: b"<< /Type /Catalog /Pages 2 0 R >>", 2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Metadata 6 0 R >>",
            4: stream(b"/Type /ObjStm /N 2000000000 /First 4", body + b"<< /A 1 >>")}
    out = bytearray(b"%PDF-1.7\n")
    offs = {}
    for n, b in objs.items():
        offs[n] = len(out)
        out += b"%d 0 obj\n" % n + b + b"\nendobj\n"
    x = len(out)
    rows = b"\0\0\0\0" + b"".join(b"\x01" + struct.pack(">H", offs[n]) + b"\0" for n in objs) + b"\x01" + struct.pack(">H", x) + b"\0" + b"\x02\0\x04\0"
    out += b"5 0 obj\n" + stream(b"/Type /XRef /W [1 2 1] /Size 7 /Root 1 0 R", rows) + b"\nendobj\n"
    out += b"startxref\n%d\n%%%%EOF\n" % x
    return bytes(out)


@case("pdf_objstm_n_scan")
def _():
    # 目次が壊れていて、ファイルを読み直して ObjStm を探す道
    b = pdf({1: b"<< /Type /Catalog /Pages 2 0 R >>", 2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
             3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>",
             4: stream(b"/Type /ObjStm /N 2000000000 /First 4", b"6 0 << /A 1 >>")})
    return b.replace(b"startxref\n", b"startxref\n9")


@case("pdf_predictor_columns")
def _():
    return page(content=zlib.compress(b"BT /F1 12 Tf (x) Tj ET"), cfilter=b"/Filter /FlateDecode /DecodeParms << /Predictor 12 /Columns 40000000000 >>")


@case("pdf_predictor_columns_small")
def _():
    return page(content=zlib.compress(b"BT /F1 12 Tf (x) Tj ET"), cfilter=b"/Filter /FlateDecode /DecodeParms << /Predictor 12 /Colors 255 /BitsPerComponent 16 /Columns 100000000 >>")


@case("pdf_flate_many_big")
def _():
    # 上限の内側の大きな中身をいくつも：控えに残らないか
    c = zlib.compress(b" " * (200 * 1048576), 9)
    n = list(range(10, 18))
    more = {i: stream(b"/Filter /FlateDecode", c) for i in n}
    more[3] = b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents [" + b" ".join(b"%d 0 R" % i for i in n) + b"] >>"
    return page(more=more)


@case("pdf_lzw_bomb")
def _():
    # LZW の濃いデータ：同じ字の続き（コードが長くなっていく）
    bits = []
    buf = 0
    n = 0
    out = bytearray()
    w = 9
    size = 258

    def put(code):
        nonlocal buf, n, w
        buf = (buf << w) | code
        n += w
        while n >= 8:
            out.append((buf >> (n - 8)) & 255)
            n -= 8
        buf &= (1 << n) - 1
    for _rep in range(100):
        put(256)
        w = 9
        put(32)
        size = 258
        c = 258
        for _ in range(3800):
            put(c)
            c += 1
            size += 1
            if size + 1 >= (1 << w) and w < 12:
                w += 1
    put(257)
    if n:
        out.append((buf << (8 - n)) & 255)
    return page(content=bytes(out), cfilter=b"/Filter /LZWDecode")


@case("pdf_image_flate_full")
def _():
    # 上限ぎりぎりの大きな画像（画素 1 億近く）を本当に展開させる
    w, h = 9900, 9900
    data = zlib.compress(b"\x80" * (w * h), 9)
    more = {6: stream(b"/Type /XObject /Subtype /Image /Width %d /Height %d /BitsPerComponent 8 /ColorSpace /DeviceGray /Filter /FlateDecode" % (w, h), data)}
    return page(content=b"q 500 0 0 500 50 50 cm /I Do Q", res=b"/XObject << /I 6 0 R >>", more=more)


@case("pdf_image_declared_huge")
def _():
    more = {6: stream(b"/Type /XObject /Subtype /Image /Width 1000000 /Height 1000000 /BitsPerComponent 8 /ColorSpace /DeviceRGB", b"\x80" * 100)}
    return page(content=b"q 500 0 0 500 50 50 cm /I Do Q", res=b"/XObject << /I 6 0 R >>", more=more)


@case("pdf_image_many_draws")
def _():
    w, h = 3000, 3000
    data = zlib.compress(os.urandom(3000) * (w * h // 3000), 1)
    more = {6: stream(b"/Type /XObject /Subtype /Image /Width %d /Height %d /BitsPerComponent 8 /ColorSpace /DeviceGray /Filter /FlateDecode" % (w, h), data)}
    return page(content=b"q 500 0 0 500 50 50 cm /I Do Q\n" * 3000, res=b"/XObject << /I 6 0 R >>", more=more)


@case("pdf_images_many_distinct")
def _():
    # 別の画像オブジェクト 40 個（画像の控えの合計）
    w, h = 4000, 4000
    data = zlib.compress(b"\x80" * (w * h * 3), 9)
    more = {6 + i: stream(b"/Type /XObject /Subtype /Image /Width %d /Height %d /BitsPerComponent 8 /ColorSpace /DeviceRGB /Filter /FlateDecode" % (w, h), data) for i in range(40)}
    res = b"/XObject << " + b" ".join(b"/I%d %d 0 R" % (i, 6 + i) for i in range(40)) + b" >>"
    return page(content=b" ".join(b"q 10 0 0 10 %d 50 cm /I%d Do Q" % (i * 10, i) for i in range(40)), res=res, more=more)


@case("pdf_ccitt_huge")
def _():
    more = {6: stream(b"/Type /XObject /Subtype /Image /Width 10000 /Height 10000 /BitsPerComponent 1 /ImageMask true /Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns 10000 /Rows 10000 >>", b"\x00" * 200)}
    return page(content=b"q 500 0 0 500 50 50 cm /I Do Q", res=b"/XObject << /I 6 0 R >>", more=more)


@case("pdf_jpeg_sof_huge")
def _():
    j = open(os.path.join(ROOT, "tests", "fixtures", "scan_box.jpg"), "rb").read()
    k = j.find(b"\xff\xc0")
    if k < 0:
        k = j.find(b"\xff\xc2")
    j = j[:k + 5] + struct.pack(">HH", 65000, 65000) + j[k + 9:]
    more = {6: stream(b"/Type /XObject /Subtype /Image /Width 65000 /Height 65000 /BitsPerComponent 8 /ColorSpace /DeviceGray /Filter /DCTDecode", j)}
    return page(content=b"q 500 0 0 500 50 50 cm /I Do Q", res=b"/XObject << /I 6 0 R >>", more=more)


@case("pdf_w_many")
def _():
    # /W [a b w] の 1 つは 70000 未満に抑えられているが、並べると何億
    w = b" ".join(b"%d %d 500" % (i * 70000, i * 70000 + 69999) for i in range(30000))
    more = {5: b"<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding /Identity-H /DescendantFonts [6 0 R] >>",
            6: b"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /X /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /W [" + w + b"] >>"}
    return page(content=b"BT /F1 12 Tf 72 720 Td <0041> Tj ET", more=more)


@case("pdf_w_huge_range")
def _():
    more = {5: b"<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding /Identity-H /DescendantFonts [6 0 R] >>",
            6: b"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /X /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /W [0 4294967295 500] /W2 [0 4294967295 1 2 3] >>"}
    return page(content=b"BT /F1 12 Tf 72 720 Td <0041> Tj ET", more=more)


@case("pdf_w2_many")
def _():
    w = b" ".join(b"%d %d 1 2 3" % (i * 70000, i * 70000 + 69999) for i in range(30000))
    more = {5: b"<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding /Identity-V /DescendantFonts [6 0 R] >>",
            6: b"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /X /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /W2 [" + w + b"] >>"}
    return page(content=b"BT /F1 12 Tf 72 720 Td <0041> Tj ET", more=more)


@case("pdf_bfrange_huge")
def _():
    cm = b"1 begincodespacerange <0000> <FFFF> endcodespacerange 1 beginbfrange <00000000> <FFFFFFFF> <0041> endbfrange"
    more = {5: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>", 6: stream(b"", cm)}
    return page(more=more)


@case("pdf_bfrange_array_many")
def _():
    arr = b"[" + b"<0041>" * 1000000 + b"]"
    cm = b"1 beginbfrange <0000> <FFFFFFFF> " + arr + b" endbfrange"
    more = {5: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>", 6: stream(b"", cm)}
    return page(more=more)


@case("pdf_bfrange_many_ranges_lookup")
def _():
    # 範囲が 100 万個で、字ごとに順に探すと 字数 × 範囲
    cm = b"1 beginbfrange " + b"".join(b"<%08X> <%08X> <0041>\n" % (i * 1000 + 0x10000, i * 1000 + 0x10000 + 900) for i in range(300000)) + b" endbfrange"
    more = {5: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>", 6: stream(b"", cm)}
    return page(content=b"BT /F1 12 Tf 72 720 Td (" + b"hello " * 200000 + b") Tj ET", more=more)


@case("pdf_cidrange_many_lookup")
def _():
    cm = b"1 begincodespacerange <0000> <FFFF> endcodespacerange 1 begincidrange " + b"".join(b"<%04X> <%04X> %d\n" % (i, i, i) for i in range(1, 65000)) * 5 + b" endcidrange"
    more = {5: b"<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding 6 0 R /DescendantFonts [7 0 R] >>",
            6: stream(b"/CMapName /X", cm),
            7: b"<< /Type /Font /Subtype /CIDFontType0 /BaseFont /X /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> >>"}
    return page(content=b"BT /F1 12 Tf 72 720 Td <" + b"FFFE" * 200000 + b"> Tj ET", more=more)


@case("pdf_firstchar_huge")
def _():
    more = {5: b"<< /Type /Font /Subtype /TrueType /BaseFont /Arial /FirstChar 0 /LastChar 2000000000 /Widths [500] >>"}
    return page(more=more)


@case("pdf_differences_huge")
def _():
    more = {5: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Differences [2000000000 /a /b] >> >>"}
    return page(more=more)


def eexec(data):
    r = 55665
    out = bytearray()
    for b in b"\0\0\0\0" + data:
        c = b ^ (r >> 8)
        out.append(c)
        r = ((c + r) * 52845 + 22719) & 0xFFFF
    return bytes(out)


@case("pdf_type1_subrs")
def _():
    clear = b"%!PS-AdobeFont-1.0: X\n/FontMatrix [0.001 0 0 0.001 0 0] def\ncurrentfile eexec\n"
    priv = b"dup /Private 8 dict dup begin /lenIV 4 def /Subrs 4000000000 array\n noaccess def\n2 index /CharStrings 1 dict dup begin\n/a 4 RD xxxx ND\nend\n"
    ff = clear + eexec(priv)
    more = {5: b"<< /Type /Font /Subtype /Type1 /BaseFont /X /FirstChar 97 /LastChar 97 /Widths [500] /FontDescriptor 6 0 R >>",
            6: b"<< /Type /FontDescriptor /FontName /X /Flags 4 /FontBBox [0 0 1000 1000] /FontFile 7 0 R >>",
            7: stream(b"/Length1 %d /Length2 %d /Length3 0" % (len(clear), len(ff) - len(clear)), ff)}
    return page(content=b"BT /F1 12 Tf 72 720 Td (a) Tj ET", more=more)


def sfnt(tables):
    n = len(tables)
    head = struct.pack(">IHHHH", 0x00010000, n, 16, 0, 0)
    pos = 12 + 16 * n
    recs = b""
    body = b""
    for tag, d in sorted(tables.items()):
        recs += struct.pack(">4sIII", tag.encode(), 0, pos + len(body), len(d))
        body += d + b"\0" * (-len(d) % 4)
    return head + recs + body


def min_tables(cmap, num=2, extra=None):
    head = bytearray(54)
    struct.pack_into(">I", head, 0, 0x00010000)
    struct.pack_into(">I", head, 12, 0x5F0F3CF5)
    struct.pack_into(">H", head, 18, 1000)
    struct.pack_into(">h", head, 50, 1)  # long loca
    hhea = bytearray(36)
    struct.pack_into(">I", hhea, 0, 0x00010000)
    struct.pack_into(">H", hhea, 34, num)
    maxp = struct.pack(">IH", 0x00010000, num) + bytes(26)
    hmtx = struct.pack(">HH", 500, 0) * num
    loca = struct.pack(f">{num + 1}I", *([0] * (num + 1)))
    t = {"head": bytes(head), "hhea": bytes(hhea), "maxp": maxp, "hmtx": hmtx, "loca": loca, "glyf": b"\0\0\0\0", "cmap": cmap}
    t.update(extra or {})
    return sfnt(t)


def cmap12(groups):
    body = struct.pack(">HHIII", 12, 0, 16 + 12 * len(groups), 0, len(groups)) + b"".join(struct.pack(">III", *g) for g in groups)
    return struct.pack(">HH", 0, 1) + struct.pack(">HHI", 3, 10, 12) + body


def ttf_pdf(font, cid=True):
    desc = b"<< /Type /FontDescriptor /FontName /X /Flags 4 /FontBBox [0 0 1000 1000] /FontFile2 7 0 R >>"
    if cid:
        more = {5: b"<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding /Identity-H /DescendantFonts [8 0 R] >>",
                6: desc, 7: stream(b"", font),
                8: b"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /X /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /CIDToGIDMap /Identity /FontDescriptor 6 0 R >>"}
        return page(content=b"BT /F1 12 Tf 72 720 Td <0001> Tj ET", more=more)
    more = {5: b"<< /Type /Font /Subtype /TrueType /BaseFont /X /FirstChar 97 /LastChar 97 /Widths [500] /FontDescriptor 6 0 R >>",
            6: desc, 7: stream(b"", font)}
    return page(content=b"BT /F1 12 Tf 72 720 Td (a) Tj ET", more=more)


@case("pdf_cmap12_many")
def _():
    # 12 バイトの区切りが、1 つで 110 万個の字。3000 個で 30 億
    return ttf_pdf(min_tables(cmap12([(i << 21, (i << 21) + 0x10FFFF, 1) for i in range(1500)])))


@case("pdf_cmap12_many_simple")
def _():
    return ttf_pdf(min_tables(cmap12([(i << 21, (i << 21) + 0x10FFFF, 1) for i in range(1500)])), cid=False)


@case("pdf_cmap4_many_seg")
def _():
    seg = 8000
    ends = struct.pack(f">{seg}H", *([0xFFFE] * seg))
    starts = struct.pack(f">{seg}H", *([0] * seg))
    deltas = struct.pack(f">{seg}h", *([1] * seg))
    ro = struct.pack(f">{seg}H", *([0] * seg))
    body = struct.pack(">HHHHHHH", 4, 0, 0, seg * 2, 0, 0, 0) + ends + b"\0\0" + starts + deltas + ro
    return ttf_pdf(min_tables(struct.pack(">HH", 0, 1) + struct.pack(">HHI", 3, 1, 12) + body), cid=False)


@case("pdf_gsub_coverage_ranges")
def _():
    # GSUB の Coverage（形式 2）：範囲 1 つは 6 バイトで 65,536 個
    rngs = struct.pack(">HH", 2, 2000) + b"".join(struct.pack(">HHH", 0, 0xFFFF, 0) for _ in range(2000))
    sub = struct.pack(">HHh", 1, 6, 1) + rngs
    lookup = struct.pack(">HHHH", 1, 0, 1, 8) + sub
    ll = struct.pack(">HH", 1, 4) + lookup
    gsub = struct.pack(">IHHH", 0x00010000, 10, 12, 14) + struct.pack(">H", 0) + struct.pack(">H", 0) + ll
    return ttf_pdf(min_tables(cmap12([(0x41, 0x41, 1)]), extra={"GSUB": gsub}))


@case("pdf_hmtx_maxp_declared")
def _():
    f = bytearray(min_tables(cmap12([(0x41, 0x41, 1)])))
    # 宣言だけ 65535 字：maxp と hhea
    return ttf_pdf(bytes(f).replace(struct.pack(">IH", 0x00010000, 2), struct.pack(">IH", 0x00010000, 0xFFFF)))


@case("pdf_dash_tiny")
def _():
    return page(content=b"[0.0001 0.0001] 0 d 1 w 0 0 m 595 842 l S")


@case("pdf_dash_tiny_huge_line")
def _():
    return page(content=b"[0.001] 0 d 1 w -1000000000 0 m 1000000000 842 l S")


@case("pdf_dash_phase")
def _():
    return page(content=b"[0.000001 0.000001] 1000000000 d 0 0 m 595 842 l S")


@case("pdf_curve_many_stroke")
def _():
    return page(content=b"2 w 0 0 m " + b"100 100 500 700 595 842 c " * 100000 + b"S")


@case("pdf_fill_huge_coords")
def _():
    return page(content=b"-1000000000 -1000000000 2000000000 2000000000 re f 0 0 m 1000000000 1000000000 l 1000000000 0 l f")


@case("pdf_stroke_huge_width")
def _():
    return page(content=b"1000000000 w 0 0 m 595 842 l S")


@case("pdf_font_size_huge")
def _():
    return page(content=b"BT /F1 1000000000 Tf 0 0 Td (hello world) Tj ET")


@case("pdf_tiling_4000")
def _():
    pat = stream(b"/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 13 /Resources << >>",
                 b"0 0 10 10 re f " * 1000)
    return page(content=b"/Pattern cs /P scn 0 0 595 842 re f", res=b"/Pattern << /P 6 0 R >>", more={6: pat})


@case("pdf_form_chain_exp")
def _():
    # フォームが次のフォームを 10 回ずつ呼ぶ（深さ 12：10^12）
    more = {}
    for i in range(12):
        more[6 + i] = stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842] /Resources << /XObject << /N %d 0 R >> >>" % (7 + i), b"/N Do " * 10)
    more[18] = stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842]", b"0 0 595 842 re f")
    return page(content=b"/N Do", res=b"/XObject << /N 6 0 R >>", more=more)


@case("pdf_form_text_amp")
def _():
    # 字の見せ方が 100 万（1 ページの上限の内側）：窓の字の控えの大きさ
    f = stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >>", b"BT /F1 12 Tf 72 720 Td " + b"(ab) Tj 1 0 Td " * 1000 + b"ET")
    return page(content=b"/N Do " * 1000, res=b"/XObject << /N 6 0 R >>", more={6: f})


@case("pdf_kids_same_page")
def _():
    # 同じページを 2 万回：1 回の呼び出しの字句の上限（2 億）の内側
    cont = b"BT /F1 12 Tf 72 720 Td " + b"(ab) Tj 1 0 Td " * 10000 + b"ET"
    kids = b" ".join([b"3 0 R"] * 20000)
    return page(content=cont, more={2: b"<< /Type /Pages /Kids [" + kids + b"] /Count 20000 >>"})


@case("pdf_kids_count_huge")
def _():
    return page(more={2: b"<< /Type /Pages /Kids [3 0 R] /Count 2000000000 >>"})


@case("pdf_kids_cycle")
def _():
    return page(more={2: b"<< /Type /Pages /Kids [3 0 R 2 0 R] /Count 2 >>"})


@case("pdf_kids_deep")
def _():
    more = {}
    n = 200000
    more[2] = b"<< /Type /Pages /Kids [6 0 R] /Count 1 >>"
    for i in range(n):
        more[6 + i] = b"<< /Type /Pages /Kids [%d 0 R] /Count 1 >>" % (7 + i)
    more[6 + n] = b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>"
    return page(more=more)


@case("pdf_xfa_occur")
def _():
    tpl = ('<template xmlns="http://www.xfa.org/schema/xfa-template/3.3/"><subform name="form1" layout="tb">'
           '<pageSet><pageArea name="P"><contentArea w="8in" h="10in"/><medium stock="letter" short="8.5in" long="11in"/></pageArea></pageSet>'
           '<subform name="Row" layout="tb"><occur min="1" max="-1" initial="2000000000"/><field name="Item"><ui><textEdit/></ui></field></subform></subform></template>')
    return xfa_pdf(tpl)


@case("pdf_xfa_occur_min")
def _():
    tpl = ('<template xmlns="http://www.xfa.org/schema/xfa-template/3.3/"><subform name="form1" layout="tb">'
           '<pageSet><pageArea name="P"><contentArea w="8in" h="10in"/><medium stock="letter" short="8.5in" long="11in"/></pageArea></pageSet>'
           '<subform name="Row" layout="tb"><occur min="2000000000" max="-1"/><field name="Item"><ui><textEdit/></ui></field></subform></subform></template>')
    ds = '<xfa:datasets xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/"><xfa:data><form1><Row><Item>a</Item></Row></form1></xfa:data></xfa:datasets>'
    return xfa_pdf(tpl, ds)


@case("pdf_xfa_occur_nested")
def _():
    # 入れ子のくり返し（それぞれ 100 回、深さ 6：10^12）
    inner = '<field name="Item"><ui><textEdit/></ui></field>'
    for i in range(6):
        inner = f'<subform name="R{i}" layout="tb"><occur min="100" max="-1" initial="100"/>{inner}</subform>'
    tpl = ('<template xmlns="http://www.xfa.org/schema/xfa-template/3.3/"><subform name="form1" layout="tb">'
           '<pageSet><pageArea name="P"><contentArea w="8in" h="10in"/><medium stock="letter" short="8.5in" long="11in"/></pageArea></pageSet>' + inner + "</subform></template>")
    return xfa_pdf(tpl)


@case("pdf_xfa_transparent_occur")
def _():
    # 名前のない subform（データに透明）
    tpl = ('<template xmlns="http://www.xfa.org/schema/xfa-template/3.3/"><subform name="form1" layout="tb">'
           '<pageSet><pageArea name="P"><contentArea w="8in" h="10in"/><medium stock="letter" short="8.5in" long="11in"/></pageArea></pageSet>'
           '<subform layout="tb"><occur min="2000000000" max="-1"/><field name="Item"><ui><textEdit/></ui></field></subform></subform></template>')
    return xfa_pdf(tpl)


def xfa_pdf(tpl, ds=None):
    pre = '<?xml version="1.0" encoding="UTF-8"?>\n<xdp:xdp xmlns:xdp="http://ns.adobe.com/xdp/">'
    cfg = '<config xmlns="http://www.xfa.org/schema/xci/3.1/"><present><pdf><version>1.7</version></pdf></present><acrobat><acrobat7><dynamicRender>required</dynamicRender></acrobat7></acrobat></config>'
    pk = [("preamble", pre), ("config", cfg), ("template", tpl)] + ([("datasets", ds)] if ds else []) + [("postamble", "</xdp:xdp>")]
    more = {}
    arr = b""
    for i, (nm, tx) in enumerate(pk):
        more[10 + i] = stream(b"/Filter /FlateDecode", zlib.compress(tx.encode()))
        arr += b"(%s) %d 0 R " % (nm.encode(), 10 + i)
    more[6] = b"<< /Fields [] /XFA [" + arr + b"] /DA (/Helv 0 Tf 0 g) /DR << /Font << /Helv 5 0 R >> >> >>"
    return page(cat=b"/AcroForm 6 0 R /NeedsRendering true", more=more)


def widget(extra, fill=None):
    w = b"<< /Type /Annot /Subtype /Widget /T (a) /FT /Tx /Rect [100 700 400 720] /P 3 0 R /F 4 /DA (/Helv 10 Tf 0 g) " + extra + b" >>"
    return page(cat=b"/AcroForm << /Fields [6 0 R] /DA (/Helv 0 Tf 0 g) /DR << /Font << /Helv 5 0 R >> >> >>", extra=b"/Annots [6 0 R]", more={6: w})


@case("pdf_comb_maxlen")
def _():
    return widget(b"/Ff %d /MaxLen 2000000000" % (1 << 24))


@case("pdf_maxlen_only")
def _():
    return widget(b"/MaxLen 2000000000")


@case("pdf_rect_huge")
def _():
    return widget(b"/Rect [-1000000000 -1000000000 1000000000 1000000000]")


@case("pdf_field_kids_cycle")
def _():
    return page(cat=b"/AcroForm << /Fields [6 0 R] >>", more={6: b"<< /T (a) /Kids [6 0 R 6 0 R] >>"})


@case("pdf_field_kids_exp")
def _():
    # 欄の木：同じ子を 2 つずつ、深さ 40（2^40）
    more = {}
    for i in range(40):
        more[6 + i] = b"<< /T (a%d) /Kids [%d 0 R %d 0 R] >>" % (i, 7 + i, 7 + i)
    more[46] = b"<< /T (z) /FT /Tx /Rect [0 0 1 1] /Subtype /Widget >>"
    return page(cat=b"/AcroForm << /Fields [6 0 R] >>", more=more)


@case("pdf_struct_tree_exp")
def _():
    more = {}
    for i in range(40):
        more[6 + i] = b"<< /S /P /K [%d 0 R %d 0 R] >>" % (7 + i, 7 + i)
    more[46] = b"<< /S /P /K 0 >>"
    return page(cat=b"/StructTreeRoot << /Type /StructTreeRoot /K 6 0 R >> /MarkInfo << /Marked true >>", more=more)


@case("pdf_outlines_exp")
def _():
    more = {}
    for i in range(30):
        more[6 + i] = b"<< /Title (x) /First %d 0 R /Last %d 0 R /Next %d 0 R >>" % (7 + i, 7 + i, 7 + i)
    more[36] = b"<< /Title (z) >>"
    return page(cat=b"/Outlines 6 0 R", more=more)


@case("pdf_ref_chain_resolve")
def _():
    more = {6 + i: b"%d 0 R" % (7 + i) for i in range(100000)}
    more[100006] = b"<< /A 1 >>"
    return page(res=b"/Font 6 0 R", more=more)


@case("pdf_nested_arrays")
def _():
    return page(extra=b"/Z " + b"[" * 500000 + b"]" * 500000)


@case("pdf_nested_dicts_content")
def _():
    return page(content=b"/X " + b"<</A " * 500000 + b"1" + b">>" * 500000 + b" BDC EMC")


@case("pdf_shading_func_tiny")
def _():
    fn = b"<< /FunctionType 2 /Domain [0 1] /N 1 /C0 [0] /C1 [1] >>"
    return page(content=b"/S sh " * 20000, res=b"/Shading << /S << /ShadingType 2 /ColorSpace /DeviceGray /Coords [0 0 595 842] /Function " + fn + b" >> >>")


@case("pdf_shading_mesh_declared")
def _():
    more = {6: stream(b"/ShadingType 4 /ColorSpace /DeviceGray /BitsPerCoordinate 32 /BitsPerComponent 16 /BitsPerFlag 8 /Decode [0 595 0 842 0 1]", b"\0" * 1000000)}
    return page(content=b"/S sh", res=b"/Shading << /S 6 0 R >>", more=more)


@case("pdf_shading_lattice")
def _():
    more = {6: stream(b"/ShadingType 5 /VerticesPerRow 2000000000 /ColorSpace /DeviceGray /BitsPerCoordinate 8 /BitsPerComponent 8 /Decode [0 595 0 842 0 1]", b"\0" * 1000)}
    return page(content=b"/S sh", res=b"/Shading << /S 6 0 R >>", more=more)


@case("pdf_indexed_hival")
def _():
    more = {6: stream(b"/Type /XObject /Subtype /Image /Width 100 /Height 100 /BitsPerComponent 8 /ColorSpace [/Indexed /DeviceRGB 2000000000 (abc)]", b"\0" * 10000)}
    return page(content=b"q 500 0 0 500 50 50 cm /I Do Q", res=b"/XObject << /I 6 0 R >>", more=more)


@case("pdf_func_sampled_big")
def _():
    fn = stream(b"/FunctionType 0 /Domain [0 1 0 1 0 1 0 1] /Range [0 1 0 1 0 1 0 1] /Size [40 40 40 40] /BitsPerSample 8", b"\0" * 1000)
    more = {6: b"[/DeviceN [/A /B /C /D] /DeviceRGB 7 0 R]", 7: fn}
    return page(content=b"/C cs 0.5 0.5 0.5 0.5 scn 0 0 100 100 re f " * 5000, res=b"/ColorSpace << /C 6 0 R >>", more=more)


@case("pdf_many_annots_ap")
def _():
    ap = stream(b"/Type /XObject /Subtype /Form /BBox [0 0 10 10]", b"BT /F1 3 Tf (x) Tj ET")
    annots = b" ".join([b"6 0 R"] * 200000)
    return page(extra=b"/Annots [" + annots + b"]", more={6: b"<< /Type /Annot /Subtype /Stamp /Rect [0 0 10 10] /AP << /N 7 0 R >> >>", 7: ap})


@case("pdf_many_widgets_same_row")
def _():
    # 1 行に 5000 の欄（近くの字を探す処理が 2 乗にならないか）
    more = {}
    ids = []
    for i in range(5000):
        more[10 + i] = b"<< /Type /Annot /Subtype /Widget /T (f%d) /FT /Tx /Rect [%d 700 %d 720] /P 3 0 R /F 4 /DA (/Helv 10 Tf 0 g) >>" % (i, i % 500, i % 500 + 1)
        ids.append(b"%d 0 R" % (10 + i))
    arr = b" ".join(ids)
    return page(content=b"BT /F1 12 Tf 72 720 Td " + b"(ab) Tj 1 0 Td " * 5000 + b"ET", cat=b"/AcroForm << /Fields [" + arr + b"] >>", extra=b"/Annots [" + arr + b"]", more=more)


@case("pdf_ruled_lines_many")
def _():
    # 罫線 3 万本（格子を探す処理）
    lines = b"".join(b"%d 0 m %d 842 l S\n" % (i % 595, i % 595) for i in range(15000)) + b"".join(b"0 %d m 595 %d l S\n" % (i % 842, i % 842) for i in range(15000))
    return page(content=lines)


@case("pdf_rect_many")
def _():
    return page(content=b"".join(b"%d %d 5 5 re S\n" % (i % 590, (i // 590) % 840) for i in range(100000)))


@case("pdf_text_one_line_many")
def _():
    return page(content=b"BT /F1 1 Tf 0 700 Td " + b"(a) Tj 0.01 0 Td " * 300000 + b"ET")


@case("pdf_text_all_overlap")
def _():
    return page(content=b"BT /F1 12 Tf " + b"72 720 Td (a) Tj -72 -720 Td " * 300000 + b"ET")


@case("pdf_xref_prev_chain")
def _():
    # /Prev が長く続く（1 万の改訂）
    base = bytearray(page())
    x0 = int(re.search(rb"startxref\n(\d+)", bytes(base))[1])
    prev = x0
    for _ in range(100000):
        here = len(base)
        base += b"xref\n0 0\ntrailer\n<< /Size 6 /Root 1 0 R /Prev %d >>\n" % prev
        prev = here
    base += b"startxref\n%d\n%%%%EOF\n" % prev
    return bytes(base)


# ====================================================================== 古い Excel（.xls）
def xls_with(records, keep_tail=True):
    from todoku.cfb import CFB, build_cfb

    raw = open(os.path.join(PUB, "sheet.xls"), "rb").read()
    c = CFB(raw)
    names = {p: c.read(p) for p in c.streams()}
    wbn = next(p for p in names if p.lower().endswith("workbook") or p.lower().endswith("book"))
    d = names[wbn]
    # 最後の EOF の前に足す
    pos = 0
    last_eof = None
    while pos + 4 <= len(d):
        op, ln = struct.unpack_from("<HH", d, pos)
        if op == 0x000A:
            last_eof = pos
        pos += 4 + ln
    add = b"".join(struct.pack("<HH", op, len(b)) + b for op, b in records)
    names[wbn] = d[:last_eof] + add + d[last_eof:]
    return build_cfb(names)


def dv_rec(r1, r2, c1, c2):
    # DV：whole（between）。式は 1 つ（ptgInt 1）
    flags = 1 | 0x100
    s = b"\x01\x00\x00\x00"  # 空の文字列 4 つ（cch=1, flag=0, 字=0）
    b = struct.pack("<I", flags) + s * 4 + struct.pack("<HH", 3, 0) + b"\x1e\x01\x00" + struct.pack("<HH", 0, 0)
    b += struct.pack("<H", 1) + struct.pack("<HHHH", r1, r2, c1, c2)
    return (0x01BE, b)


DVAL = (0x01B2, struct.pack("<HIIII", 0, 0, 0, 0xFFFFFFFF, 1))


@case("xls_merge_all")
def _():
    return xls_with([(0x00E5, struct.pack("<H", 1) + struct.pack("<HHHH", 100, 65535, 5, 255))])


@case("xls_merge_many")
def _():
    # 1 つの記録に 1024 の結合。ファイルの大きさに比例しない
    recs = []
    for k in range(40):
        recs.append((0x00E5, struct.pack("<H", 1000) + b"".join(struct.pack("<HHHH", k * 1000 + i, 65535, 5 + i % 200, 255) for i in range(1000))))
    return xls_with(recs)


@case("xls_merge_plus_dv")
def _():
    return xls_with([(0x00E5, struct.pack("<H", 1) + struct.pack("<HHHH", 100, 65535, 5, 255)), DVAL, dv_rec(0, 65535, 0, 255)])


@case("xls_lastcell_plus_dv")
def _():
    return xls_with([(0x0201, struct.pack("<HHH", 65535, 255, 0)), DVAL, dv_rec(0, 65535, 0, 255)])


@case("xls_mulblank_rows")
def _():
    # 1 行 256 列の書式つきの空のセル × 65536 行：ファイルも大きい（それぞれ約 520 バイト）が、窓の格子に開くと何倍か
    row = lambda r: (0x00BE, struct.pack("<HH", r, 0) + struct.pack("<H", 1) * 256 + struct.pack("<H", 255))
    return xls_with([row(r) for r in range(20, 30000)])


@case("xls_condfmt_many")
def _():
    b = struct.pack("<HHHHHH", 1, 0, 0, 65535, 0, 255) + struct.pack("<H", 1000) + b"".join(struct.pack("<HHHH", i, 65535, 0, 255) for i in range(1000))
    return xls_with([(0x01B0, b)] * 40)


@case("xls_colinfo_many")
def _():
    return xls_with([(0x007D, struct.pack("<HHHHHH", 0, 65535, 100, 0, 1, 0))] * 200000)



# ====================================================================== グラフ（xlsx の図の中）
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def rels(*items):
    return ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(
        f'<Relationship Id="{i}" Type="{R_NS}/{t}" Target="{g}"/>' for i, t, g in items) + "</Relationships>")


def xlsx_chart(chart_xml, kind="chart", name="chart1"):
    drawing = ('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
               '<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
               '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="c"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
               f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="{R_NS}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>')
    return xlsx(sheet(tail=f'<drawing xmlns:r="{R_NS}" r:id="rId1"/>'), add={
        "xl/worksheets/_rels/sheet1.xml.rels": rels(("rId1", "drawing", "../drawings/drawing1.xml")),
        "xl/drawings/drawing1.xml": drawing,
        "xl/drawings/_rels/drawing1.xml.rels": rels(("rId1", kind, f"../charts/{name}.xml")),
        f"xl/charts/{name}.xml": chart_xml},
        ct_extra='<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                 f'<Override PartName="/xl/charts/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')


def chart_ser(val):
    return (f'<c:chartSpace {C_NS}><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>{val}</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>')


@case("xlsx_chart_pt_idx")
def _():
    return xlsx_chart(chart_ser('<c:val><c:numRef><c:f>S!$A$1</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="2000000000"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>'))


@case("xlsx_chart_ptcount")
def _():
    return xlsx_chart(chart_ser('<c:val><c:numRef><c:f>S!$A$1</c:f><c:numCache><c:ptCount val="2000000000"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>'))


@case("xlsx_chart_multilvl_ptcount")
def _():
    return xlsx_chart(chart_ser('<c:cat><c:multiLvlStrRef><c:f>S!$A$1</c:f><c:multiLvlStrCache><c:ptCount val="2000000000"/><c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>'
                                '<c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>'))


@case("xlsx_chart_multilvl_quad")
def _():
    # 段の点の数 × ptCount（lookups）で 2 乗
    pts = "".join(f'<c:pt idx="{i}"><c:v>a</c:v></c:pt>' for i in range(30000))
    return xlsx_chart(chart_ser(f'<c:cat><c:multiLvlStrRef><c:f>S!$A$1</c:f><c:multiLvlStrCache><c:ptCount val="30000"/><c:lvl>{pts}</c:lvl><c:lvl>{pts}</c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>'))


@case("xlsx_chart_ref_all")
def _():
    # 控えのない参照：ブックのセルを読む（10 万マスまでの上限の内側で、系列をたくさん）
    sers = "".join(f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:val><c:numRef><c:f>S!$A$1:$A$100000</c:f></c:numRef></c:val></c:ser>' for i in range(2000))
    return xlsx_chart(f'<c:chartSpace {C_NS}><c:chart><c:plotArea><c:barChart>{sers}</c:barChart></c:plotArea></c:chart></c:chartSpace>')


@case("xlsx_chartex_idx")
def _():
    cx = ('<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0">'
          '<cx:numDim type="val"><cx:f>S!$A$1</cx:f><cx:lvl ptCount="1"><cx:pt idx="2000000000">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>'
          '<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>')
    return xlsx_chart(cx, kind="chart", name="chart2")


# ====================================================================== 古い Word（.doc）
@case("doc_ccp_huge")
def _():
    from todoku.cfb import CFB, build_cfb

    c = CFB(open(os.path.join(PUB, "form.doc"), "rb").read())
    st = {p: c.read(p) for p in c.streams()}
    wd = bytearray(st["WordDocument"])
    struct.pack_into("<I", wd, 0x4C, 2000000000)  # ccpText
    st["WordDocument"] = bytes(wd)
    return build_cfb(st)


@case("doc_ccp_foot_huge")
def _():
    from todoku.cfb import CFB, build_cfb

    c = CFB(open(os.path.join(PUB, "form.doc"), "rb").read())
    st = {p: c.read(p) for p in c.streams()}
    wd = bytearray(st["WordDocument"])
    for off in (0x50, 0x54, 0x58, 0x5C, 0x60, 0x64):
        struct.pack_into("<I", wd, off, 1000000000)
    st["WordDocument"] = bytes(wd)
    return build_cfb(st)


# ====================================================================== 実行の入口
def main():
    out = sys.argv[1]
    want = sys.argv[2:]
    os.makedirs(out, exist_ok=True)
    resfile = os.path.join(out, "results_crafted.jsonl")
    done = set()
    if os.path.exists(resfile):
        done = {json.loads(line)["name"] for line in open(resfile) if line.strip()}
    for name, (fn, tools, env) in CASES.items():
        if want and not any(w in name for w in want):
            continue
        if name in done and not os.environ.get("CRAFT_AGAIN"):
            continue
        ext = {"xlsx": "xlsx", "xlsxf": "xlsx", "ods": "ods", "odt": "odt", "docx": "docx", "hwpx": "hwpx", "rtf": "rtf", "csv": "csv",
               "json": "json", "txt": "txt", "zip": "zip", "pdf": "pdf", "xls": "xls", "doc": "doc"}[name.split("_")[0]]
        path = os.path.join(out, name + "." + ext)
        try:
            blob = fn()
        except Exception as e:  # noqa: BLE001
            print("BUILD-FAIL", name, type(e).__name__, e, flush=True)
            continue
        with open(path, "wb") as f:
            f.write(blob)
        if os.environ.get("CRAFT_ONLY"):
            print("built", name, len(blob))
            continue
        g = guarded(path, tools, limit_s=float(os.environ.get("CRAFT_S", 60)), limit_gb=1.5, max_kills=2, env_extra=env, outdir=out)
        rec = {"name": name, "size": len(blob), "peak": g["peak"], "kills": g["kills"],
               "tools": {x["tool"]: {k: x.get(k) for k in ("s", "peak", "ok", "err", "at") if x.get(k) is not None}
                         for x in g["recs"] if "tool" in x}}
        with open(resfile, "a") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        slow = {k: v["s"] for k, v in rec["tools"].items() if v.get("s", 0) > 5}
        crashes = {k: (v["err"][:60], v.get("at")) for k, v in rec["tools"].items() if v.get("at")}
        print(f"{name}: {len(blob):,}B peak={g['peak']}MB kills={[(k['tool'], k['why'][:30]) for k in g['kills']]} slow={slow} crash={crashes}", flush=True)
        if os.path.exists(path) and os.path.getsize(path) > 20_000_000:
            os.unlink(path)
        for fnm in os.listdir(out):
            if fnm.startswith("out_"):
                os.unlink(os.path.join(out, fnm))


if __name__ == "__main__":
    main()
