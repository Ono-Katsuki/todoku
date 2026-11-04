"""大きな書類の続き：分けて読む道でも、窓の全体と同じものが、id も同じに読めること（小さな上限で、分けて読む道に入れる）。"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from todoku import api, lazy, limits


@pytest.fixture(autouse=True)
def _small_limits(monkeypatch, tmp_path):
    monkeypatch.setattr(limits, "LAZY_MIN_BYTES", 0)
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2000)
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    yield
    lazy.clear()


W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def _word_blob(body: str, numbering: bool = True, header: bool = True) -> bytes:
    buf = io.BytesIO()
    ct = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    if numbering:
        drels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    if header:
        drels += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    drels += "</Relationships>"
    num = (f'<?xml version="1.0"?><w:numbering {W_NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
           '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
    head = f'<?xml version="1.0"?><w:hdr {W_NS}><w:p><w:r><w:t>ヘッダーの字　氏名（　　　　）</w:t></w:r></w:p></w:hdr>'
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS}><w:body>{body}<w:sectPr/></w:body></w:document>'
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc.encode())
        z.writestr("word/_rels/document.xml.rels", drels)
        if numbering:
            z.writestr("word/numbering.xml", num)
        if header:
            z.writestr("word/header1.xml", head)
    return buf.getvalue()


def _para(i: int) -> str:
    num = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' if i % 7 == 0 else ""
    return f'<w:p>{num}<w:r><w:t xml:space="preserve">第{i}段落の字 氏名（　　　　）</w:t></w:r></w:p>'


def _table(i: int) -> str:
    return (f'<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表{i}の見出し</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>'
            f'<w:tr><w:tc><w:p><w:r><w:t>住所{i}</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>')


def _sdt(i: int) -> str:
    return (f'<w:sdt><w:sdtPr><w:alias w:val="欄"/><w:tag w:val="same"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>値{i}</w:t></w:r></w:p></w:sdtContent></w:sdt>')


def _box(i: int) -> str:
    return (f'<w:p><w:r><w:t>箱の前{i}</w:t></w:r><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape><v:textbox><w:txbxContent>'
            f'<w:p><w:r><w:t>箱の中{i}（　　）</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>')


def _heavy_word(tmp_path, n: int = 400):
    parts = []
    for i in range(n):
        parts.append(_para(i))
        if i % 11 == 0:
            parts.append(_table(i))
        if i % 37 == 0:
            parts.append(_sdt(i))
        if i % 53 == 0:
            parts.append(_box(i))
    p = tmp_path / "heavy.docx"
    p.write_bytes(_word_blob("".join(parts)))
    return p


def _follow(p, first: dict | None = None, max_chars=None) -> list:
    out = []
    rng = dict(first or {})
    for _ in range(100000):
        pg, _d = api.read_range(str(p), rng, max_chars)
        out.append(pg)
        if pg.next is None:
            return out
        rng = dict(pg.next)
    raise AssertionError("終わらない")


def test_word_is_lazy_and_lines_equal_whole_window(tmp_path, monkeypatch):
    from todoku.docx import document as dd
    from todoku.docx.document import DocxDocument
    from todoku.docx.window import build
    from todoku.lazy import docxdoc

    monkeypatch.setattr(docxdoc, "CHUNK_BYTES", 2000)
    p = _heavy_word(tmp_path)
    doc = lazy.open_doc(str(p), force=True)
    assert isinstance(doc, docxdoc.DocxLazy)
    win = build(DocxDocument(p.read_bytes(), "heavy.docx"))
    want = [ln for st in win["stories"] for ln in st["reading"]]
    pages = _follow(p, max_chars=1500)
    got = [re.sub(r"^〔段落 \d+(?:〜\d+)?〕", "", ln) for pg in pages for b in pg.body for ln in b.split("\n") if not b.startswith(("〔ヘッダー", "### ", "\n"))]
    got = [x for x in got if x and not x.startswith(("### ", "- ")) or x in want]
    assert got == want
    # id も、窓の全体と同じ（段落番号の番号も、前のまとまりから続く）
    all_ids = {f["id"]: f for f in win["fields"]}
    res = []
    rng: dict = {}
    while True:
        r = api.fields_range(str(p), rng, limit=50)
        res += r["fields"]
        if not r["next"]:
            break
        rng = dict(r["next"])
    assert sorted(f["id"] for f in res) == sorted(all_ids)
    assert {f["id"]: f.get("label") for f in res} == {k: f.get("label") for k, f in all_ids.items()}


def test_word_range_read_builds_only_the_needed_chunk(tmp_path, monkeypatch):
    from todoku.lazy import docxdoc

    monkeypatch.setattr(docxdoc, "CHUNK_BYTES", 2000)
    p = _heavy_word(tmp_path)
    doc = lazy.open_doc(str(p), force=True)
    assert len(doc.chunks) > 5
    seen = []
    real = docxdoc.DocxLazy.chunk_doc

    def spy(self, i, real_parts=()):
        seen.append(i)
        return real(self, i, real_parts)

    monkeypatch.setattr(docxdoc.DocxLazy, "chunk_doc", spy)
    pg, _ = api.read_range(str(p), {"paragraphs": "300-301"})
    assert len(seen) <= 2
    assert "〔段落 300" in pg.text


def test_word_audit_in_chunks_says_ok(tmp_path, monkeypatch):
    from todoku.lazy import docxdoc

    monkeypatch.setattr(docxdoc, "CHUNK_BYTES", 8000)
    p = _heavy_word(tmp_path, 120)
    doc = lazy.open_doc(str(p), force=True)
    st = doc.audit_begin()
    while not st.get("done"):
        st = doc.audit_step(st)
    names = {c["name"]: c for c in st["checks"]}
    assert names["読み落とし（文書の文字）"]["ok"], names["読み落とし（文書の文字）"]
    assert names["全入力欄への書き込み"]["ok"], names["全入力欄への書き込み"]
    assert names["字を目印にした書き込み"]["ok"], names["字を目印にした書き込み"]


def test_word_fill_keeps_working_on_lazy_ids(tmp_path, monkeypatch):
    from todoku.lazy import docxdoc

    monkeypatch.setattr(docxdoc, "CHUNK_BYTES", 2000)
    p = _heavy_word(tmp_path, 60)
    res = api.fields_range(str(p), {"paragraphs": "50-52"})
    fid = next(f["id"] for f in res["fields"] if f["id"].startswith("P"))
    r = api.fill(str(p), {fid: "山田"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r


# ---------------------------------------------------------------------- どの形式も：たどり終えると、窓の全体と同じものを読んだことになる
FIX = __import__("pathlib").Path(__file__).parent / "fixtures"
_LINE = re.compile(r"^〔段落 \d+(?:〜\d+)?〕")


def _word_like(p, monkeypatch, cls_name: str, max_chars=1200) -> None:
    """Word と同じ窓の形式：たどり終えた行が、窓の全体の行とちょうど同じ。欄の id も同じ。"""
    from todoku.lazy import blocks, seqdocs

    monkeypatch.setattr(seqdocs, "UNIT_CHARS", 200)
    monkeypatch.setattr(seqdocs, "UNIT_BYTES", 800)
    monkeypatch.setattr(blocks, "VIEW_CHUNK_PARAS", 5)
    doc = lazy.open_doc(str(p), force=True)
    assert type(doc).__name__ == cls_name, type(doc)
    k, win, _parts = api.window(str(p))
    want = [ln for st in win["stories"] if not st.get("template") for ln in st["reading"]]
    got = []
    for pg in _follow(p, max_chars=max_chars):
        for b in pg.body:
            for ln in b.split("\n"):
                if _LINE.match(ln):
                    got.append(_LINE.sub("", ln))
    assert got == want
    ids = sorted(f["id"] for f in api.fields(str(p)))
    res = []
    rng: dict = {}
    while True:
        r = api.fields_range(str(p), rng, limit=20)
        res += r["fields"]
        if not r["next"]:
            break
        rng = dict(r["next"])
    assert sorted(f["id"] for f in res) == ids
    # 窓の終わりの節（埋め込み・画像の一覧など）も、たどると全部読める
    text = "\n".join(pg.text for pg in _follow(p, max_chars=max_chars))
    whole = "\n".join(t for _n, t in _parts)
    for ln in whole.split("\n"):
        if ln.startswith(("### 書類の中の画像", "## 埋め込み", "## マクロ")):
            assert ln.split("（")[0] in text, ln


def _sheet_like(p, monkeypatch, cls_name: str, max_chars=1500) -> None:
    doc = lazy.open_doc(str(p), force=True)
    assert type(doc).__name__ == cls_name, type(doc)
    k, win, _parts = api.window(str(p))
    text = "\n".join(pg.text for pg in _follow(p, max_chars=max_chars))
    # 値は全部出る（ほかのシートから指されたセルの印「？」は、窓の全体のときだけ付くことがあるので、値で比べる）
    bare = re.compile(r"[［〈]([^＝］〉]*)＝([^］〉]*?)？?[］〉]")
    flat = bare.sub(r"\2", text)
    for sh in win["sheets"]:
        for ln in sh.get("reading", []):
            t = ln.get("text", "")
            if t:
                assert bare.sub(r"\2", t) in flat, (sh["name"], t)
    ids = {f["id"] for f in api.fields(str(p))}
    got = set()
    for d in doc.sheet_meta:
        rng: dict = {"sheet": d["name"]}
        while True:
            r = api.fields_range(str(p), rng, limit=200)
            got |= {f["id"] for f in r["fields"]}
            if not r.get("next"):
                break
            rng = dict(r["next"])
    # ほかのシートの数式に指されたセルの「？」は、そのシートを読まないと分からない（番地で書ける）
    only_whole = {f["id"] for f in api.fields(str(p)) if f.get("agent") and f.get("value") not in (None, "")}
    assert ids - only_whole <= got, sorted(ids - got)[:10]


def _copy(tmp_path, name):
    src = FIX / "public" / name if (FIX / "public" / name).exists() else FIX / name
    dst = tmp_path / name
    dst.write_bytes(src.read_bytes())
    return dst


def test_concat_doc(tmp_path, monkeypatch):
    _word_like(_copy(tmp_path, "form.doc"), monkeypatch, "DocLazy")


def test_concat_jtd(tmp_path, monkeypatch):
    _word_like(_copy(tmp_path, "form.jtd"), monkeypatch, "ViewLazy")


def test_concat_rtf(tmp_path, monkeypatch):
    _word_like(_copy(tmp_path, "form.rtf"), monkeypatch, "ViewLazy")


def test_concat_odt(tmp_path, monkeypatch):
    _word_like(_copy(tmp_path, "form.odt"), monkeypatch, "OdtLazy")


def test_concat_odt_many(tmp_path, monkeypatch):
    from test_openformats import ODT_AUTO, ODT_FORMS, make_odt

    body = "".join(f'<text:p>第{i}段落 氏名（　　）</text:p><text:h text:outline-level="1">見出し{i}</text:h>'
                   f'<table:table table:name="t{i}"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell><text:p>氏名{i}</text:p></table:table-cell><table:table-cell><text:p/></table:table-cell></table:table-row></table:table>'
                   for i in range(60))
    p = tmp_path / "many.odt"
    p.write_bytes(make_odt(body, ODT_AUTO, ODT_FORMS))
    _word_like(p, monkeypatch, "OdtLazy")


def test_concat_hwpx(tmp_path, monkeypatch):
    _word_like(_copy(tmp_path, "form.hwpx"), monkeypatch, "HwpxLazy")


def test_concat_hwp(tmp_path, monkeypatch):
    from test_hwpx import make_hwp5

    p = tmp_path / "a.hwp"
    p.write_bytes(make_hwp5())
    _word_like(p, monkeypatch, "HwpLazy")


def test_concat_pptx(tmp_path, monkeypatch):
    import pptx_maker

    sp = lambda t: f'<p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>{t}</a:t></a:r></a:p></p:txBody></p:sp>'
    slides = [sp(f"スライド{i}の字 氏名（　　）") + sp(f"二つ目{i}") for i in range(12)]
    p = tmp_path / "a.pptx"
    p.write_bytes(pptx_maker.make(slides, notes={2: sp("ノートの字")}, master_text="型の字"))
    _word_like(p, monkeypatch, "PptxLazy")
    text = "\n".join(pg.text for pg in _follow(p, max_chars=1200))
    assert "型の字" in text


def test_concat_xls(tmp_path, monkeypatch):
    _sheet_like(_copy(tmp_path, "r4_general.xls"), monkeypatch, "XlsLazy")


def test_concat_xls_biff5(tmp_path, monkeypatch):
    from test_legacy import make_biff5

    p = tmp_path / "b5.xls"
    p.write_bytes(make_biff5())
    _sheet_like(p, monkeypatch, "XlsLazy")


def test_concat_ods(tmp_path, monkeypatch):
    _sheet_like(_copy(tmp_path, "sheet.ods"), monkeypatch, "OdsLazy")
    _sheet_like(_copy(tmp_path, "r4_general.ods"), monkeypatch, "OdsLazy")


def test_concat_xlsb(tmp_path, monkeypatch):
    import xlsb_maker as m

    rows = b"".join(m.row(r) + m.real(0, r * 1.5) + m.isst(1, r % 3) + m.st(2, f"字{r}") for r in range(300))
    p = tmp_path / "a.xlsb"
    p.write_bytes(m.package({"データ": m.sheet(rows), "二枚目": m.sheet(m.row(4) + m.st(1, "別のシート"))}, strings=["一", "二", "三"]))
    _sheet_like(p, monkeypatch, "XlsbLazy")
    p2 = tmp_path / "sample.xlsb"
    p2.write_bytes(m.sample())
    lazy.clear()
    _sheet_like(p2, monkeypatch, "XlsbLazy")
    text = "\n".join(pg.text for pg in _follow(p, {"raw": True}, max_chars=500))
    assert "字299" in text and "別のシート" in text


def test_xls_row_range_uses_lazy_strings(tmp_path):
    """行の範囲を読むとき、共有文字列は使う分だけ読む（CONTINUE のまたぎも窓の読み手と同じ）。"""
    p = _copy(tmp_path, "r4_general.xls")
    doc = lazy.open_doc(str(p), force=True)
    wb = doc.book_of({0: [(1, 3)]})
    assert wb.sheets[0].cells
    import todoku.xls.book as xb

    whole = xb.XlsWorkbook(p.read_bytes(), p.name)
    for key, c in wb.sheets[0].cells.items():
        assert whole.sheets[0].cells[key].value == c.value


# ---------------------------------------------------------------------- 字のまま（raw）
@pytest.mark.parametrize("name", ["form.doc", "form.jtd", "r4_general.xls"])
def test_raw_reads_old_formats(tmp_path, name):
    p = _copy(tmp_path, name)
    pages = _follow(p, {"raw": True}, max_chars=500)
    text = "\n".join(pg.text for pg in pages)
    assert len(text) > 10
    k, win, parts = api.window(str(p))
    if k in ("doc", "jtd"):
        flat = text.replace("\n", "").replace("\t", "")
        for st in win["stories"]:
            for t in st["_plain"][:10]:
                t = t.strip(" 　\r\n\t\x07")
                if len(t) >= 4 and not re.search(r"[\x00-\x1f]", t):
                    assert t[:4] in flat, t


def test_raw_hwp(tmp_path):
    from test_hwpx import make_hwp5

    p = tmp_path / "a.hwp"
    p.write_bytes(make_hwp5())
    assert "복지 급여 신청서" in "\n".join(pg.text for pg in _follow(p, {"raw": True}, max_chars=500))


# ---------------------------------------------------------------------- JSON
def test_json_stream_pages_and_path(tmp_path, monkeypatch):
    import json

    data = {"申請": {"氏名": "山田", "住所": None}, "項目": [{"番号": i, "名前": f"名{i}", "値": i * 1.5} for i in range(400)], "長い字": "あ" * 5000,
            "入れ子": {"a": {"b": {"c": [1, 2, {"d": "深い"}]}}}, "空": {}, "空の配列": []}
    p = tmp_path / "a.json"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    doc = lazy.open_doc(str(p), force=True)
    assert type(doc).__name__ == "JsonLazy"
    pages = _follow(p, max_chars=800)
    text = "\n".join(pg.text for pg in pages)
    assert len(pages) > 5
    assert "| /項目/399 | 399 | \"名399\" | 598.5 |" in text
    assert "`/申請/氏名` ＝ \"山田\"" in text and "`/申請/住所` ＝ null" in text and "`/入れ子/a/b/c/2/d` ＝ \"深い\"" in text
    assert "`/空` ＝ {}" in text and "`/空の配列` ＝ []" in text
    long_ = "".join(ln for ln in text.split("\n") if ln.count("あ") > 10 or ln.startswith("- `/長い字`"))
    assert long_.count("あ") == 5000
    one, _ = api.read_range(str(p), {"path": "/項目/7"})
    assert "/項目/7/名前" in one.text and "/項目/8" not in one.text and one.next is None
    with pytest.raises(api.ApiError, match="ありません"):
        api.read_range(str(p), {"path": "/ない"})
    r = api.fields_range(str(p), {"path": "/申請"})
    assert {f["id"] for f in r["fields"]} == {"/申請/氏名", "/申請/住所"}
    res = api.fields_range(str(p), {}, limit=30)
    assert res["next"]
    # 全部の値を読むと、窓の全体の値の数と同じ
    k, win, parts = api.window(str(p))
    for v in win["values"]:
        if v["id"].startswith("/項目") or v["id"] == "/長い字":
            continue
        assert f"`{v['id']}`" in text, v


def test_big_file_is_not_refused_by_size(tmp_path, monkeypatch):
    """メモリの上限の半分を超えるファイルでも、断らない（写像で読む）。"""
    from test_lazy import _make_pdf

    monkeypatch.setattr(limits, "MAX_FILE", 100)
    p = _make_pdf(tmp_path / "a.pdf", 8)
    lazy.clear()
    pg, _ = api.read_range(str(p), {"pages": "3-3"})
    assert "Page 3 line" in pg.text
    out = tmp_path / "o.pdf"
    r = api.fill(str(p), {"@1": {"after": "Page 3 line 4 text sample", "value": "ABC", "where": "below"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert out.read_bytes().startswith(p.read_bytes())
    assert not list(tmp_path.glob(".todoku-append*"))
    # ZIP の書類も、断らずに開く
    z = tmp_path / "a.docx"
    z.write_bytes(_word_blob(_para(1)))
    assert api.fields(str(z))
