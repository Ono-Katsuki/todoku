"""記入例と照らして、入力欄の見落としを確かめる仕組み（PDF・Word）と、そこで見つけた欄の規則のテスト。"""

from __future__ import annotations

from todoku.docx.document import DocxDocument
from todoku.docx.example import check as word_check
from todoku.docx.fill import fill as docx_fill
from todoku.docx.window import build as docx_build
from todoku.pdf.example import check as pdf_check
from todoku.pdf.layout import pdf_blanks
from todoku.xlsx.window import inline_blanks

from test_docx import SAMPLE_BODY, make_docx, p
from test_pdf import jp_font_objs, jp_text, make_pdf

CHARS = "申請書氏名郵便番号令和年月日新規変更□とどく花子12345678レ余白欄外"


def _form(values: bool, stray: bool = False) -> bytes:
    """氏名（空のマス）・郵便番号（7桁の枠）・令和　年　月　日・□ 新規 □ 変更 の申請書。values なら記入例。"""
    font, extra, codes = jp_font_objs(CHARS, 5)
    t = lambda s: jp_text(codes, s)  # noqa: E731
    c = bytearray()
    c += b"BT /F1 16 Tf 250 780 Td " + t("申請書") + b" Tj ET\n"
    c += b"0.8 w 60 700 m 540 700 l S 60 670 m 540 670 l S 60 640 m 540 640 l S\n"
    c += b"60 700 m 60 640 l S 160 700 m 160 640 l S 540 700 m 540 640 l S\n"
    c += b"BT /F1 10 Tf 70 682 Td " + t("氏名") + b" Tj ET\n"
    c += b"BT /F1 10 Tf 70 652 Td " + t("郵便番号") + b" Tj ET\n"
    for i in range(8):
        x = 160 + i * 20
        c += f"{x} 670 m {x} 640 l S\n".encode()
    c += b"BT /F1 10 Tf 380 600 Td " + t("令和") + b" Tj 40 0 Td " + t("年") + b" Tj 40 0 Td " + t("月") + b" Tj 40 0 Td " + t("日") + b" Tj ET\n"
    c += b"BT /F1 10 Tf 70 560 Td " + t("□新規") + b" Tj 80 0 Td " + t("□変更") + b" Tj ET\n"
    if values:
        c += b"BT /F1 10 Tf 200 682 Td " + t("とどく花子") + b" Tj ET\n"
        for i, d in enumerate("1234567"):
            c += f"BT /F1 10 Tf {165 + i * 20} 652 Td ".encode() + t(d) + b" Tj ET\n"
        c += b"BT /F1 10 Tf 405 600 Td " + t("8") + b" Tj 40 0 Td " + t("4") + b" Tj 40 0 Td " + t("1") + b" Tj ET\n"
        c += b"BT /F1 10 Tf 70 560 Td " + t("レ") + b" Tj ET\n"
    if stray:
        c += b"BT /F1 10 Tf 300 470 Td " + t("欄外余白") + b" Tj ET\n"
    return make_pdf(bytes(c), {"F1": font}, extra)


def test_pdf_example_all_values_in_fields():
    r = pdf_check(_form(False), _form(True))
    assert r.added >= 4 and r.covered == r.added and not r.missed, (r.headline(), r.missed)
    assert r.same_version


def test_pdf_example_reports_value_outside_fields():
    r = pdf_check(_form(False), _form(True, stray=True))
    assert [m.text for m in r.missed] == ["欄外余白"], r.missed
    assert r.covered == r.added - 1


def test_word_example_all_values_in_fields():
    blank = make_docx(SAMPLE_BODY)
    filled = (SAMPLE_BODY.replace("令和　　年　　月　　日", "令和　８年　４月　１日")
              .replace('<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr></w:p>', p("とどく 花子"))
              .replace("□新規", "■新規"))
    r = word_check(blank, make_docx(filled))
    assert r.added >= 2 and r.covered == r.added and not r.missed, (r.headline(), r.missed)


def test_word_example_reports_text_outside_fields():
    blank = make_docx(SAMPLE_BODY)
    filled = SAMPLE_BODY.replace(p("申請書"), p("申請書とどく花子"))
    r = word_check(blank, make_docx(filled))
    assert [m.text for m in r.missed] == ["とどく花子"], r.missed


# ====================================================================== 記入例で見つけた見落としの規則
def _blanks(text: str) -> int:
    lead = len(text) - len(text.lstrip("　"))
    trail = len(text) - len(text.rstrip("　"))
    return len(pdf_blanks(text, lead, trail))


def test_time_fields():
    # 「　　：　　～　　：　　」：時・分・時・分の4か所（マスの縁までの空きも）
    assert _blanks("　　　：　　　～　　　：　　　") == 4


def test_unit_fields_and_spaced_labels():
    assert _blanks("□身体障害者手帳　　種　　級") == 2  # 「種」「級」の前に書く
    assert _blanks("片道　　　　　分") == 1
    assert _blanks("氏　　名") == 0  # 字間を空けた見出しは欄でない
    assert _blanks("事　業　主　氏　名") == 0
    assert _blanks("　　分類") == 0  # 「分類」の「分」は単位でない
    assert len(inline_blanks("番　　号")) == 0 and len(inline_blanks("第　　号")) == 1


def test_word_address_cell_insertion_point():
    # 住所のマス「　　　都道　　　　郡　市」：都道府県の前、郡の前、市の後ろ（町名・番地）に書ける
    body = '<w:tbl><w:tr><w:tc>' + p("住所") + '</w:tc><w:tc>' + p("　　　都道　　　　郡　市") + '</w:tc></w:tr></w:tbl>'
    blob = make_docx(body)
    win = docx_build(DocxDocument(blob, "a.docx"))
    inl = [f for f in win["fields"] if f["kind"] == "inline"]
    assert len(inl) == 3, [f["label"] for f in inl]
    vals = dict(zip([f["id"] for f in inl], ["山梨", "甲府", "丸の内１－６－１"]))
    res = docx_fill(DocxDocument(blob, "a.docx"), vals)
    after = docx_build(DocxDocument(res.blob, "a.docx"))
    text = "\n".join(after["reading"]) if isinstance(after.get("reading"), list) else str(after)
    assert "丸の内１－６－１" in text
    got = {f["id"]: f.get("value") for f in after["fields"]}
    assert all(got.get(k) == v for k, v in vals.items()), got


# ====================================================================== Excel：表の行の空欄・名前の違うシートの対応
def make_xlsx(sheets: dict[str, tuple[dict[str, object], list[tuple[str, str]]]]) -> bytes:
    """最小の .xlsx。sheets：シート名 → ({"A1": 値}, [(範囲, 入力規則の種類)])。"""
    import io
    import re as _re
    import zipfile
    from xml.sax.saxutils import escape

    ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' \
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        over = "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                       'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                       for i in range(1, len(sheets) + 1))
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   + over + '</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        sh = "".join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(sheets, 1))
        z.writestr("xl/workbook.xml", f'<?xml version="1.0" encoding="UTF-8"?><workbook {ns}><sheets>{sh}</sheets></workbook>')
        rels = "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                       f'Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets) + 1))
        rels += (f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                 'Target="styles.xml"/>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
        z.writestr("xl/styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><styleSheet {ns}>'
                   '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
                   '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
                   '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                   '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                   '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>')
        for i, (cells, vals) in enumerate(sheets.values(), 1):
            rows: dict[int, list[str]] = {}
            for a, v in sorted(cells.items(), key=lambda kv: (int(_re.sub(r"\D", "", kv[0])), kv[0])):
                r = int(_re.sub(r"\D", "", a))
                if v is None:
                    rows.setdefault(r, []).append(f'<c r="{a}"/>')  # 値のないセル（表の枠などの書式だけがあるもの）
                elif isinstance(v, str):
                    rows.setdefault(r, []).append(f'<c r="{a}" t="inlineStr"><is><t>{escape(v)}</t></is></c>')
                else:
                    rows.setdefault(r, []).append(f'<c r="{a}"><v>{v}</v></c>')
            data = "".join(f'<row r="{r}">{"".join(cs)}</row>' for r, cs in sorted(rows.items()))
            dv = ""
            if vals:
                dv = f'<dataValidations count="{len(vals)}">' + "".join(
                    f'<dataValidation type="{t}" allowBlank="1" sqref="{ref}">'
                    + ('<formula1>"A,B"</formula1>' if t == "list" else '<formula1>0</formula1>') + '</dataValidation>'
                    for ref, t in vals) + '</dataValidations>'
            z.writestr(f"xl/worksheets/sheet{i}.xml", f'<?xml version="1.0" encoding="UTF-8"?><worksheet {ns}>'
                       f'<sheetData>{data}</sheetData>{dv}</worksheet>')
    return buf.getvalue()


HEAD = {"A1": "Type", "B1": "Description", "C1": "Amount", "D1": "Notes",
        **{f"{c}{r}": None for c in "ABCD" for r in range(2, 10)}}  # 表の体（値のないセル）
VALS = [("A2:A9", "list"), ("C2:C9", "decimal")]


def test_xlsx_unmarked_table_columns_are_left_to_the_agent():
    from todoku.xlsx.diagnose import diagnose
    from todoku.xlsx.window import input_cells, undecided_cells
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(make_xlsx({"Budget": (HEAD, VALS)}))
    sh = wb.sheet("Budget")
    got = input_cells(wb, sh)
    # 入力規則のある列だけが、規則で決めた入力欄
    assert (2, 1) in got and (2, 3) in got and (2, 2) not in got and (2, 4) not in got
    # 印のない「Description」「Notes」の列は、入力欄と言い切らず「規則では決められない」
    und = undecided_cells(wb, sh, got)
    assert und[(2, 2)] == "Description" and und[(9, 4)] == "Notes" and (1, 2) not in und
    sc = diagnose(wb).score()
    assert sc["rules_reached_name"] == "操作できるの手前" and sc["rules_reached"] < sc["reached"]


def test_xlsx_example_pairs_renamed_sheet():
    from todoku.xlsx.example import check as xlsx_check

    blank = make_xlsx({"Budget": (HEAD, VALS)})
    filled = make_xlsx({"Sample budget": ({**HEAD, "A2": "A", "B2": "Hall hire", "C2": 120, "D2": "30 people"}, VALS)})
    r = xlsx_check(blank, filled)
    assert r.inputs == 4 and not r.not_input_fields, (r.headline(), r.not_input_fields)
    assert sorted(r.agent_fields) == ["Budget!B2", "Budget!D2"]  # 印のない列に書かれた所は、エージェントの判断の所
