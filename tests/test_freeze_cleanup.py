"""版を固める前の片付け：見出しは字のまま（label_text）と探す鍵（label_key）、表の格子のマスは近くの字の事実、
文字の中の空きのそろえ方（align）、欄の一覧に出す事実（label_key・referenced_by・evidence）。"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku import api

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
U = '<w:rPr><w:u w:val="single"/></w:rPr>'


def _docx(tmp_path, body: str, name: str = "a.docx") -> str:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '</Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>')
    path = tmp_path / name
    path.write_bytes(buf.getvalue())
    return str(path)


def _p(text: str, rpr: str = "") -> str:
    return f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def _run(text: str, rpr: str = "") -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _book(tmp_path, cells: dict, name: str = "a.xlsx", setup=None) -> str:
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for a, v in cells.items():
        ws[a] = v
    if setup:
        setup(ws)
    path = tmp_path / name
    wb.save(path)
    return str(path)


def _paras(path: str) -> list[str]:
    from todoku.docx.document import DocxDocument, iter_paras

    d = DocxDocument(open(path, "rb").read(), path)
    return [q.text for _k, _part, blocks in d.all_blocks() for q in iter_paras(blocks)]


# ---------------------------------------------------------------------- 見出し：字のまま・探す鍵
def test_word_labels_are_verbatim_with_a_search_key(tmp_path):
    # 字間を空けた見出し「氏　　名」は詰めない（窓の字のまま）。探すための詰めた写しは label_key
    body = ('<w:tbl><w:tr><w:tc>' + _p("氏　　名") + '</w:tc><w:tc><w:p/></w:tc></w:tr>'
            '<w:tr><w:tc>' + _p("事 業 所 番 号") + '</w:tc><w:tc><w:p/></w:tc></w:tr>'
            '<w:tr><w:tc>' + _p("Given name") + '</w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>')
    by = {f["id"]: f for f in api.fields(_docx(tmp_path, body))}
    assert by["T1.R1.C2"]["label"] == "氏　　名" and by["T1.R1.C2"]["label_key"] == "氏名"
    assert by["T1.R2.C2"]["label"] == "事 業 所 番 号" and by["T1.R2.C2"]["label_key"] == "事業所番号"
    assert by["T1.R3.C2"]["label"] == "Given name" and "label_key" not in by["T1.R3.C2"]  # 語の区切りの空白は残す（鍵は同じ字）


def test_pdf_labels_are_verbatim_with_a_search_key():
    from todoku.pdf.layout import PageLayout

    nb = {"left": {"text": "氏　　名", "_rtl": False}, "above": {"text": "申請者", "_rtl": False}}
    assert PageLayout._label_of(nb) == ("氏　　名", "氏　　名", "申請者")
    from todoku.xlsx.window import with_label_keys

    f = with_label_keys([{"label": "氏　　名"}])[0]
    assert f["label_key"] == "氏名"


def test_norm_label_compaction_rule_is_gone():
    from todoku.xlsx import window

    assert not hasattr(window, "norm_label")  # 「2 字以下が 3 つ以上」で詰める決まりは使わない（label_key だけ）


# ---------------------------------------------------------------------- list_fields に出す事実
def test_fields_pass_label_key_referenced_by_and_agent_evidence():
    from todoku.api import _field

    f = {"id": "S!B3", "label": "氏　名", "label_key": "氏名", "kind": "text", "referenced_by": ["S!D9"],
         "agent": True, "evidence": {"borders": "上下左右", "referenced_by": ["S!D9"]}}
    d = _field(f, sheet="S")
    assert d["label_key"] == "氏名" and d["referenced_by"] == ["S!D9"]
    assert d["evidence"] == {"borders": "上下左右", "referenced_by": ["S!D9"]}
    assert "evidence" not in _field(dict(f, agent=False), sheet="S")  # 規則の欄は、理由（reason）で示す


def test_sheet_agent_cell_evidence_reaches_list_fields(tmp_path):
    def setup(ws):
        ws["D1"] = "=B1*2"

    src = _book(tmp_path, {"A1": "金額"}, setup=setup)
    fl = api.fields(src)
    got = [f for f in fl if f.get("agent") and f.get("evidence")]
    for f in got:
        assert isinstance(f["evidence"], dict)


# ---------------------------------------------------------------------- 離れた字：すぐ上の字を「1 行あけて」と言わない
def test_far_label_counts_rows_between_not_row_difference():
    from todoku.api import far_label

    near = {"above": {"text": "対象者", "label": "対象者", "cell": "C12", "rows": 1, "mm": 0}}
    assert far_label({"id": "S!C13", "label": "対象者", "near": near}) is None  # すぐ上のマス：見出しのまま
    near = {"above": {"text": "対象者", "label": "対象者", "cell": "C10", "rows": 3, "mm": 12.0}}
    assert far_label({"id": "S!C13", "label": "対象者", "near": near}) == "C13（上に 2 行あけて『対象者』）"


# ---------------------------------------------------------------------- 文字の中の空き：そろえ方（align）
def test_word_inline_align_right_keeps_the_gap_width_and_is_proved(tmp_path):
    body = '<w:p>' + _run("住所") + _run("　　　　　　　　", U) + '</w:p>'
    src = _docx(tmp_path, body)
    fid = next(f["id"] for f in api.fields(src) if f["kind"] == "inline")
    out = str(tmp_path / "o.docx")
    r = api.fill(src, {fid: {"value": "東京", "align": "right"}}, out)
    assert r["proof_ok"] and r["read_back_ok"], r
    # 空きの初めの 1 字は、見出しにくっつけないよう元のまま残す（gap_interior）。残りの 7 字（14）に「東京」（4）を右へ
    assert r["aligned"][fid] == {"align": "right", "pad": [10, 0]}
    text = _paras(out)[0]
    assert text == "住所　" + " " * 10 + "東京", repr(text)
    # 既定（そろえ方の指定なし）は左：下線の長さを保つ空白は後ろ
    r = api.fill(src, {fid: "東京"}, str(tmp_path / "o2.docx"))
    assert r["proof_ok"] and r["read_back_ok"] and "aligned" not in r
    assert _paras(str(tmp_path / "o2.docx"))[0] == "住所　東京" + " " * 10


def test_word_inline_align_center_on_plain_gap(tmp_path):
    # 下線のない空き（「？」の所）も、そろえ方を指定すれば空きの幅を保つ（指定しなければ、空きを値に置き換える）
    body = _p("番号　　　　　号")
    src = _docx(tmp_path, body)
    fid = next(f["id"] for f in api.fields(src) if f["kind"] == "inline")
    r = api.fill(src, {fid: {"value": "12", "align": "center"}}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    lead, trail = r["aligned"][fid]["pad"]
    text = _paras(str(tmp_path / "o.docx"))[0]
    assert "12" in text and text.endswith("号") and lead >= 1 and abs(lead - trail) <= 1, (text, lead, trail)


def test_word_inline_align_rejects_bad_options(tmp_path):
    body = '<w:p>' + _run("住所") + _run("　　　　", U) + '</w:p>'
    src = _docx(tmp_path, body)
    fid = next(f["id"] for f in api.fields(src) if f["kind"] == "inline")
    with pytest.raises(api.ApiError, match="align"):
        api.fill(src, {fid: {"value": "x", "align": "middle"}}, str(tmp_path / "o.docx"))
    with pytest.raises(api.ApiError, match="ambiguous_width"):
        api.fill(src, {fid: {"value": "x", "ambiguous_width": "half"}}, str(tmp_path / "o.docx"))


def test_word_underline_pad_counts_ambiguous_width_the_same_in_fill_and_proof(tmp_path):
    # 曖昧な幅の字（①）を含む値：書き込みと証明が同じ決まり（段落に全角の字があれば全角）で空白を数える
    body = '<w:p>' + _run("番号") + _run("　　　　　　", U) + '</w:p>'
    src = _docx(tmp_path, body)
    fid = next(f["id"] for f in api.fields(src) if f["kind"] == "inline")
    r = api.fill(src, {fid: "①2"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert _paras(str(tmp_path / "o.docx"))[0] == "番号　①2" + " " * 7  # 残りの 5 字（10）から ①（2）と 2（1）
    r = api.fill(src, {fid: {"value": "①2", "ambiguous_width": "narrow"}}, str(tmp_path / "o2.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert _paras(str(tmp_path / "o2.docx"))[0] == "番号　①2" + " " * 8


def test_sheet_inline_align_right_in_cell_text(tmp_path):
    src = _book(tmp_path, {"A1": "申請日", "B1": "令和　　年　　月　　日"})
    spans = [s["id"] for f in api.fields(src) for s in f.get("spans", [])]
    assert spans
    first = spans[0]
    r = api.fill(src, {first: {"value": "8", "align": "right"}}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    (key, info), = r["aligned"].items()
    assert key == first and info["align"] == "right" and info["pad"][1] == 0 and info["pad"][0] >= 1
    openpyxl = pytest.importorskip("openpyxl")
    got = openpyxl.load_workbook(tmp_path / "o.xlsx")["S"]["B1"].value
    assert got.startswith("令和") and (" " * info["pad"][0] + "8年") in got, repr(got)


def test_textwidth_gap_padding():
    from todoku.textwidth import gap_padding

    assert gap_padding("　　　", "12", "left") == ("", "    ")
    assert gap_padding("　　　", "12", "right") == ("    ", "")
    assert gap_padding("　　　", "12", "center") == ("  ", "  ")
    assert gap_padding("  ", "長い値", "right") == ("", "")  # 空きより広い値：足さない
    assert gap_padding("　　", "○", "left", context="氏名") == ("", "  ")  # 全角の文脈：○ は 2
    assert gap_padding("　　", "○", "left", ambiguous_wide=False) == ("", "   ")


# ---------------------------------------------------------------------- 段階：小さな片付け
def test_where_keeps_every_place_and_shows_the_rest_as_count():
    from todoku.grade import Grade, Reason

    r = Reason("B", "t", "d", [f"S!A{i}" for i in range(1, 31)])
    d = Grade("A", [r]).to_dict()["reasons"][0]
    assert d["where"][:12] == [f"S!A{i}" for i in range(1, 13)] and d["where"][12] == "ほか 18 か所" and len(d["where"]) == 13
    assert d["where_all"] == [f"S!A{i}" for i in range(1, 31)]
    small = Grade("A", [Reason("B", "t", "d", ["x"])]).to_dict()["reasons"][0]
    assert small["where"] == ["x"] and "where_all" not in small


def test_unlabeled_sheet_fields_are_counted():
    from todoku.grade import Reason, sheet_input_reasons

    doc = {"sheets": [{"fields": [{"id": "S!B2", "kind": "text", "label": "", "unlabeled": True, "reason_codes": ["border"]},
                                  {"id": "S!B3", "kind": "text", "label": "氏名", "reason_codes": ["validation"]}]}]}
    reasons: list[Reason] = []
    write: list[Reason] = []
    sheet_input_reasons(doc, reasons, write)
    got = {r.title: r for r in write}
    assert got["見出しのない欄"].where == ["S!B2"] and got["見出しのない欄"].grade == "B"
    assert "印のある入力欄" not in got  # 見出しのない欄があれば、すべて印で決まるとは言わない


def test_formula_mismatch_is_one_rule_with_the_saving_app_as_a_fact(tmp_path):
    from todoku import grade

    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 2
    ws["B1"] = "=A1*3"
    p = tmp_path / "m.xlsx"
    wb.save(p)
    # 保存された計算結果を、数式と合わない値にする（openpyxl は値を保存しないので、XML を直す）
    blob = p.read_bytes()
    zin = zipfile.ZipFile(io.BytesIO(blob))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in zin.namelist():
            data = zin.read(n)
            if n == "xl/worksheets/sheet1.xml":
                data = data.replace(b"<f>A1*3</f><v></v>", b"<f>A1*3</f><v>7</v>").replace(b"<f>A1*3</f>", b"<f>A1*3</f><v>7</v>", 1) \
                    if b"<v>7</v>" not in data else data
            if n == "docProps/app.xml":
                data = b'<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Google Sheets</Application></Properties>'
            z.writestr(n, data)
    g = grade.grade_blob(buf.getvalue(), "m.xlsx")
    mism = [r for r in g.reasons if r.title == "保存された値と合わない数式"]
    assert mism and mism[0].grade == "C" and "Google Sheets" in mism[0].detail, [(r.grade, r.title) for r in g.reasons]
    assert not any(r.title == "Excel 以外で保存された計算結果" for r in g.reasons)


def test_ooxml_package_without_office_parts_is_unknown_not_excel(tmp_path):
    from todoku.book import kind_of, unsupported_message

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                   'Target="visio/document.xml"/></Relationships>')
        z.writestr("visio/document.xml", "<VisioDocument/>")
    blob = buf.getvalue()
    assert kind_of(blob) == "ooxml-unknown"
    assert "application/vnd.ms-visio.drawing.main+xml" in unsupported_message("ooxml-unknown", blob)
    p = tmp_path / "d.vsdx"
    p.write_bytes(blob)
    with pytest.raises(api.Unsupported):
        api.window(str(p))
    g = api.grade(str(p)) if hasattr(api, "grade") else None
    if g is not None:
        assert g["grade"] == "F" and g["write"]["grade"] == "D"


def test_pptx_audit_uses_the_shared_box_glyphs():
    from todoku import chars
    from todoku.pptx import audit

    assert audit.BOXES == chars.BOXES | chars.MAYBE_CHECKED


def test_read_table_limit_returns_first_rows_and_counts(tmp_path):
    from todoku import mcp_server

    cells = {"A1": "番号", "B1": "名前"}
    for i in range(2, 13):
        cells[f"A{i}"] = i - 1
        cells[f"B{i}"] = f"名{i}"
    src = _book(tmp_path, cells)
    res = mcp_server.call_tool("read_table", {"path": src, "header": "番号", "limit": 4})["structuredContent"]
    assert res["rows"] == 11 and res["rows_returned"] == 4 and len(res["all_rows"]) == 4 and "limit" in res["note"]
    res = mcp_server.call_tool("read_table", {"path": src, "header": "番号"})["structuredContent"]
    assert len(res["all_rows"]) == 11 and "rows_returned" not in res


def test_box_mark_unchecks_every_checked_glyph():
    # 不具合：▣（MAYBE_CHECKED）を外すときの字がなく、KeyError になっていた
    from todoku import chars
    from todoku.docx.window import box_mark

    for ch in chars.CHECKED_BOXES | chars.MAYBE_CHECKED:
        assert box_mark(ch, False) in chars.EMPTY_BOXES and box_mark(ch, True) == ch
