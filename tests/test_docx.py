"""Word（.docx）の読み書きのテスト。テストの中で小さな文書を組み立てて確かめ、公開の様式（corpus/wide_docx）でも確かめる。"""

from __future__ import annotations

import glob
import io
import zipfile
from pathlib import Path

import pytest

from todoku.book import kind_of
from todoku.docx.audit import run as audit
from todoku.docx.document import DocxDocument, number_text
from todoku.docx.fill import FillError, fill, format_word_date
from todoku.docx.proof import prove
from todoku.docx.window import build, to_markdown

CORPUS = sorted(glob.glob(str(Path(__file__).parent.parent / "corpus" / "wide_docx" / "**" / "*.docx"), recursive=True))

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'


def make_docx(body: str, settings: str = "") -> bytes:
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
        if settings:
            z.writestr("word/settings.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>{settings}</w:settings>')
    return buf.getvalue()


def p(text: str, rpr: str = "") -> str:
    return f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


SAMPLE_BODY = (
    p("申請書")
    + p("令和　　年　　月　　日")
    + '<w:tbl><w:tr><w:tc>' + p("氏名") + '</w:tc><w:tc><w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr></w:p></w:tc></w:tr></w:tbl>'
    + p("□新規　　□変更")
    + p("住所　　　　　　　　", "<w:rPr><w:u w:val=\"single\"/></w:rPr>")
    + '<w:p><w:r><w:t>電話：</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Tel"/><w:enabled/>'
      '<w:textInput><w:maxLength w:val="13"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>'
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>     </w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    + '<w:p><w:sdt><w:sdtPr><w:alias w:val="同意"/><w14:checkbox><w14:checked w14:val="0"/>'
      '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>'
      '<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>同意します</w:t></w:r></w:p>'
)


def test_kind_of_docx():
    assert kind_of(make_docx(p("a"))) == "docx"


def test_number_and_date_formats():
    assert [number_text(3, f) for f in ("decimal", "decimalFullWidth", "decimalEnclosedCircle", "aiueoFullWidth", "iroha", "japaneseCounting")] == \
        ["3", "３", "③", "ウ", "ハ", "三"]
    import datetime as dt

    assert format_word_date(dt.date(2026, 4, 1), "ggge年M月d日") == "令和8年4月1日"
    assert format_word_date(dt.date(2019, 5, 1), "ggge年M月d日") == "令和元年5月1日"


def test_sample_window():
    win = build(DocxDocument(make_docx(SAMPLE_BODY), "s.docx"))
    by = {f["id"]: f for f in win["fields"]}
    kinds = sorted(f["kind"] for f in win["fields"])
    assert kinds.count("inline") >= 4  # 年・月・日の前、住所の下線
    assert kinds.count("checkbox") == 3  # □新規・□変更・同意のチェック
    assert by["Tel"]["maxlen"] == 13
    assert any(f["id"].startswith("T1.R1.C2") and f["label"] == "氏名" for f in win["fields"])
    md = to_markdown(win)
    assert "［T1.R1.C2］" in md and "☐［同意］" in md


def test_sample_fill_prove_readback():
    blob = make_docx(SAMPLE_BODY)
    doc = DocxDocument(blob, "s.docx")
    win = build(doc)
    vals = {}
    for f in win["fields"]:
        vals[f["id"]] = True if f["kind"] == "checkbox" else ("８" if f["kind"] == "inline" else "とどく 花子")
    vals["Tel"] = "03-1234-5678"
    res = fill(doc, vals)
    pr = prove(blob, res.blob, {x.id: x.value for x in res.patches})
    assert pr.ok, pr.unexpected
    after = build(DocxDocument(res.blob, "s.docx"))
    got = {f["id"]: f["value"] for f in after["fields"]}
    for k, v in vals.items():
        assert (bool(got[k]) == v) if isinstance(v, bool) else got[k] == v, (k, v, got[k])
    with pytest.raises(FillError):
        fill(doc, {"Tel": "0" * 20})  # 13字までの欄


def test_protection_forms_only():
    blob = make_docx(SAMPLE_BODY, '<w:documentProtection w:edit="forms" w:enforcement="1"/>')
    doc = DocxDocument(blob, "s.docx")
    win = build(doc)
    cell = next(f["id"] for f in win["fields"] if f["id"].startswith("T1"))
    assert next(f for f in win["fields"] if f["id"] == cell).get("protected")
    with pytest.raises(FillError):
        fill(doc, {cell: "書けないはず"})
    fill(doc, {"Tel": "03"})  # フォームの欄は書ける


def test_sample_audit():
    au = audit(make_docx(SAMPLE_BODY), "s.docx")
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]


@pytest.mark.parametrize("name", ["11ninshintodoke.docx", "torisage-word.docx", "sanko15_kakunin_seiyakusy.docx", "jigyoukeikaku.docx"])
def test_corpus_audit(name):
    path = next((x for x in CORPUS if Path(x).name == name), None)
    if path is None:
        pytest.skip(name)
    au = audit(Path(path).read_bytes(), name)
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]
