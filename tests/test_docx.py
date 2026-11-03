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
        ["3", "３", "③", "ウ", "ﾊ", "三"]  # iroha は半角のカタカナ（ECMA-376 17.18.59）
    import datetime as dt

    assert format_word_date(dt.date(2026, 4, 1), "ggge年M月d日", "ja-JP", "japan") == "令和8年4月1日"
    assert format_word_date(dt.date(2019, 5, 1), "ggge年M月d日", "ja-JP", "japan") == "令和元年5月1日"


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


U = '<w:rPr><w:u w:val="single"/></w:rPr>'

    # 書類の作りで決まる空欄（下線を引いた空白・「＿＿＿」）は規則の欄。2つ以上の空白の並びは、字の意味（年・月・日など）を見ずに「？」（同じ番号の並び）
def test_only_structural_blanks_are_rules_and_word_patterns_are_agent():
    # 書類の作りで決まる空欄（下線を引いた空白・「＿＿＿」）は規則の欄。年・月・日などの字の並びから推定した空欄は「？」（同じ番号の並び）
    body = (p("令和　　年　　月　　日") + p("氏名＿＿＿＿") + p("住所　　　　", U)
            + p("連絡先　　") + '<w:p><w:r><w:t>番号</w:t></w:r><w:r>' + U + '<w:t xml:space="preserve">  </w:t></w:r></w:p>'
            + p("https://example.jp/__icsFiles/a.docx"))
    win = build(DocxDocument(make_docx(body), "b.docx"))
    by = {f["id"]: f for f in win["fields"]}
    assert all(by[i].get("agent") and "2つ以上の空白" in by[i]["why"] for i in ("P1.1", "P1.2", "P1.3"))
    assert not by["P2.1"].get("agent") and not by["P3.1"].get("agent")  # 「＿＿＿＿」・下線を引いた空白
    assert by["P4.1"].get("agent") and "2つ以上の空白" in by["P4.1"]["why"]  # 段落の終わりの空白の並び：前は出さなかった
    assert by["P5.1"].get("agent") and "下線を引いた短い空白" in by["P5.1"]["why"]  # 2字の下線：前は出さなかった
    assert by["P6.1"].get("agent") and "URL" in by["P6.1"]["why"]  # 英数字にはさまれた「__」は名前の一部のことも
    md = to_markdown(win)
    assert "令和［P1.1？］年［P1.2？］月［P1.3？］日" in md and "氏名［P2.1］" in md
    # 「？」の所も id で書け、証明と読み戻しが通る
    blob = make_docx(body)
    vals = {"P1.1": "8", "P1.2": "4", "P1.3": "1", "P2.1": "山田", "P4.1": "03", "P5.1": "12"}
    res = fill(DocxDocument(blob, "b.docx"), vals)
    assert prove(blob, res.blob, {x.id: x.value for x in res.patches}).ok
    got = {f["id"]: f["value"] for f in build(DocxDocument(res.blob, "b.docx"))["fields"]}
    assert all(got[k] == v for k, v in vals.items()), got


def test_checked_marks_are_checkbox_candidates():
    # ☑・■・✓ は、チェック済みのチェックボックスかもしれない（■ は箇条の印のことも）：「？」の候補で、値は True。外すと印のない箱にする
    body = p("☑ 新規　　□ 変更") + p("■ 申請者の情報") + p("✓ 確認した")
    blob = make_docx(body)
    win = build(DocxDocument(blob, "m.docx"))
    by = {f["id"]: f for f in win["fields"]}
    assert by["P1.1"]["kind"] == "checkbox" and by["P1.1"].get("agent") and by["P1.1"]["value"] is True
    assert by["P1.2"]["kind"] == "checkbox" and not by["P1.2"].get("agent") and by["P1.2"]["value"] is False
    assert by["P2.1"].get("agent") and by["P3.1"].get("agent") and "箇条の印" in by["P2.1"]["why"]
    md = to_markdown(win)
    assert "☑［P1.1？］ 新規" in md and "■［P2.1？］ 申請者の情報" in md
    res = fill(DocxDocument(blob, "m.docx"), {"P1.1": False, "P1.2": True, "P2.1": False, "P3.1": True})
    assert prove(blob, res.blob, {x.id: x.value for x in res.patches}).ok
    after = build(DocxDocument(res.blob, "m.docx"))
    # 「申請者の情報」「確認した」の後ろ（P2.a1・P3.a1）も、見出しの後ろの候補（印の箱があっても落とさない）。書いたのは印だけ
    got = {f["id"]: f["value"] for f in after["fields"] if not f["id"].endswith(".a1")}
    assert got == {"P1.1": False, "P1.2": True, "P2.1": False, "P3.1": True}, got
    text = "\n".join(ln for st in after["stories"] for ln in st["reading"])
    assert "☐［P1.1］ 新規" in text and "☑［P1.2］ 変更" in text and "☐［P2.1］ 申請者の情報" in text


def test_label_beside_ruled_cell_is_kept_as_alternative():
    # マスの見出し「住所」の後ろは、すぐ右の空のマスがあっても消さない。代わりの所（alternative_to）として示し、どちらに書くかはエージェント
    body = '<w:tbl><w:tr><w:tc>' + p("住所") + '</w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    win = build(DocxDocument(make_docx(body), "t.docx"))
    by = {f["id"]: f for f in win["fields"]}
    assert not by["T1.R1.C2"].get("agent")
    assert by["P1.a1"].get("agent") and by["P1.a1"]["alternative_to"] == "T1.R1.C2"
    assert "T1.R1.C2 の代わりの所" in to_markdown(win)


def test_empty_paragraph_in_cell_with_long_text_is_a_candidate():
    # 見出しのあるマスの中の空の段落：マスのほかの字が長くても（説明の文のこともある）、字数では消さず「？」で示す
    long = "申請者の住所を記入してください（住民票の住所と同じ）"
    body = '<w:tbl><w:tr><w:tc>' + p(long) + '<w:p/></w:tc><w:tc>' + p("備考") + '</w:tc></w:tr></w:tbl>'
    win = build(DocxDocument(make_docx(body), "n.docx"))
    cand = [f for f in win["fields"] if f.get("agent") and "マスのほかの字が長く" in f["why"]]
    assert [f["id"] for f in cand] == ["P2.a1"], win["fields"]


@pytest.mark.parametrize("name", ["11ninshintodoke.docx", "torisage-word.docx", "sanko15_kakunin_seiyakusy.docx", "jigyoukeikaku.docx"])
def test_corpus_audit(name):
    path = next((x for x in CORPUS if Path(x).name == name), None)
    if path is None:
        pytest.skip(name)
    au = audit(Path(path).read_bytes(), name)
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]


def test_anchored_write_where_rules_found_nothing():
    # 規則でも「？」でも示していない所：エージェントが字を目印にして書く（{"old": 元の字, "new": 書き換えた後の字}）。
    # 変わるのは old と new の違う所だけで、証明は同じ決め方で位置を決め直して確かめる
    from todoku.docx.fill import FillError, fill, written_of
    from todoku.docx.proof import prove
    from todoku.docx.window import build

    blob = make_docx(SAMPLE_BODY + p("茨城県　　　農林事務所長　殿") + p("第２条　訓練は年１回以上行う。") + p("第３条　訓練は年１回以上行う。"))
    doc = DocxDocument(blob, "a.docx")
    res = fill(doc, {"@1": {"old": "茨城県　　　農林", "new": "茨城県　土浦　農林"},
                     "@2": {"old": "年１回", "new": "年２回", "occurrence": 2}})
    assert [(x.id, x.value) for x in res.patches] == [("P9.w1", "土浦"), ("P11.w1", "２")], res.patches
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected
    lines = [ln for st in build(DocxDocument(res.blob, "a.docx"))["stories"] for ln in st["reading"]]
    # 「殿」で終わる段落の後ろも、語を見ずに「？」の所として示す（字の意味では除かない）
    assert any(ln.startswith("茨城県　［P9.w1＝土浦］　農林事務所長　殿") for ln in lines), lines
    assert "第３条　訓練は年［P11.w1＝２］回以上行う。" in lines, lines
    # 決まらないときは、候補を添えて断る
    with pytest.raises(FillError, match="2 か所にあります"):
        fill(DocxDocument(blob, "a.docx"), {"@1": {"old": "年１回", "new": "年２回"}})
    with pytest.raises(FillError, match="見つかりません"):
        fill(DocxDocument(blob, "a.docx"), {"@1": {"old": "年３回", "new": "年２回"}})
    # 証明は、目印から決まる所と違う書き込みを見つける
    bad = written_of(res.patches)
    bad["P9.w1"] = {"@": {"old": "茨城県　　　農林", "new": "茨城県　水戸　農林"}, "value": "土浦"}
    assert not prove(blob, res.blob, bad).ok


def test_embedded_excel_table_is_shown_as_plain_text():
    # 文書に埋め込んだ Excel の表（word/embeddings の .xlsx）：本文の〔埋め込み1：Excel の表〕の所と、窓の節に、表を平文で出す
    from test_example import make_xlsx
    from todoku.docx.window import build, to_markdown

    ns = ('xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    body = (p("工賃の実績") + f'<w:p><w:r><w:object {ns}><v:shape id="_x0000_i1025" style="width:300pt;height:80pt"/>'
            '<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" '
            'ObjectID="_1774977835" r:id="rId9"/></w:object></w:r></w:p>')
    base = make_docx(body)
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(base)) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            z.writestr(n, src.read(n))
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
                   'Target="embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>')
        z.writestr("word/embeddings/Microsoft_Excel_Worksheet.xlsx", make_xlsx({"Sheet1": ({"A1": "年度", "B1": "工賃総額", "A2": "5年度"}, [])}))
    doc = DocxDocument(buf.getvalue(), "e.docx")
    win = build(doc)
    lines = [ln for st in win["stories"] for ln in st["reading"]]
    assert "〔埋め込み1：Excel の表〕" in lines, lines
    md = to_markdown(win)
    assert "## 埋め込み1：Excel の表（Excel.Sheet.12・word/embeddings/Microsoft_Excel_Worksheet.xlsx）" in md
    assert "工賃総額" in md and "5年度" in md
