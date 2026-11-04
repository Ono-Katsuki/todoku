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
