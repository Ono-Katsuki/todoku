"""集めた書類の全部の所に書いて読み直す検査（audit_whole）で見つかった誤りの確かめ（合成の文書）。

- ふりがなつきの字（親字）・図の印・脚注の印・数式の中の空白は、書き換えられないので、所として示さない
- しおりの番号の重なり：元の文書に初めからある重なりは、書き込みの誤りにしない（書き込みで増えた重なりだけ）
- 文書のプロパティ（docProps/core.xml の題など）に結び付いたコンテンツコントロール：プロパティの値も書いて、そう伝える
- RTF の脚注の印のすぐ後ろへの書き込み：脚注の中にも、上付きの字のグループの中にも入れない
- .doc に埋め込んだ部品の種類は、入れ物が名乗る ProgID（\\x01CompObj）で決める（EMBED の命令の語は、ProgID でないことがある）
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest
import test_word_findings as wf
from test_embeds import compobj_stream
from test_openformats import B, make_rtf
from test_round4_writers import sdt

from todoku import api, lazy
from todoku.docx.proof import prove


@pytest.fixture(autouse=True)
def clean():
    api._WINDOWS.clear()
    lazy.clear()
    yield
    api._WINDOWS.clear()
    lazy.clear()


def _fields(path) -> list[dict]:
    api._WINDOWS.clear()
    lazy.clear()
    return api.fields(str(path))


def _zip_with(blob: bytes, change: dict[str, bytes]) -> bytes:
    src = zipfile.ZipFile(io.BytesIO(blob))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            z.writestr(n, change.pop(n, None) or src.read(n))
        for n, data in change.items():
            z.writestr(n, data)
    return buf.getvalue()


# ---------------------------------------------------------------------- ふりがなつきの字の中の空白
RUBY = ('<w:r><w:ruby><w:rubyPr/><w:rt><w:r><w:t>フリガナ</w:t></w:r></w:rt><w:rubyBase><w:r><w:t xml:space="preserve">氏　　名</w:t>'
        "</w:r></w:rubyBase></w:ruby></w:r>")


def test_blank_inside_ruby_base_is_not_a_place(tmp_path):
    path = wf.write(tmp_path, wf.docx(wf.p(wf.r("　　　　"), RUBY, wf.r("　　　　　　　　印"))))
    f = _fields(path)
    assert [x["id"] for x in f] == ["P1.c0", "P1.c8"]  # 親字「氏　　名」の中の空白（c5）は、所にしない
    assert "氏　　名（フリガナ）" in f[0]["context"]  # 字は、そのまま読める
    r = api.fill(path, {"P1.c8": "山田太郎"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r


# ---------------------------------------------------------------------- しおりの番号の重なり
BM = '<w:bookmarkStart w:id="0" w:name="{}"/><w:bookmarkEnd w:id="0"/>'


def test_duplicate_bookmark_ids_of_the_source_do_not_fail_the_proof(tmp_path):
    blob = wf.docx(wf.p(BM.format("a"), wf.r("氏名：　　　　　　")) + wf.p(BM.format("b"), wf.r("住所：　　　　　　")))
    path = wf.write(tmp_path, blob)
    fid = _fields(path)[0]["id"]
    r = api.fill(path, {fid: "山田"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    out = (tmp_path / "o.docx").read_bytes()
    x = zipfile.ZipFile(io.BytesIO(out)).read("word/document.xml").decode()
    mine = re.search(r'<w:bookmarkStart w:id="(\d+)" w:name="_todoku_', x)[1]
    assert mine != "0"
    # 書き込みで重なりが増えたら、証明が止める
    bad = _zip_with(out, {"word/document.xml": x.replace(f'w:id="{mine}"', 'w:id="0"').encode()})
    pr = prove(blob, bad, {fid: "山田"})
    assert not pr.ok and any("しおりの番号が重なっています" in e for e in pr.lint_errors)


# ---------------------------------------------------------------------- 文書のプロパティに結び付いた欄
CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
CORE = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="{CORE_NS}" xmlns:dc="{DC_NS}">'
        "<dc:title>前の題</dc:title><dc:creator>作った人</dc:creator></cp:coreProperties>")
CORE_STORE = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"


def _core_docx(xpath: str, core: str = CORE, rel: bool = True) -> bytes:
    pr = (f"<w:dataBinding w:prefixMappings=\"xmlns:ns0='{CORE_NS}' xmlns:ns1='{DC_NS}'\" w:xpath=\"{xpath}\" "
          f'w:storeItemID="{CORE_STORE}"/><w:text/>')
    base = wf.docx(wf.p(wf.r("題 "), sdt("タイトル", pr, content="前の題", plc=False)))
    rels = zipfile.ZipFile(io.BytesIO(base)).read("_rels/.rels")
    if rel:
        rels = rels.replace(b"</Relationships>", b'<Relationship Id="rCore" Type="http://schemas.openxmlformats.org/package/2006/'
                            b'relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')
    return _zip_with(base, {"_rels/.rels": rels, "docProps/core.xml": core.encode()})


def test_control_bound_to_core_properties_writes_the_property(tmp_path):
    blob = _core_docx("/ns0:coreProperties[1]/ns1:title[1]")
    path = wf.write(tmp_path, blob)
    assert {f["id"]: f for f in _fields(path)}["タイトル"]["value"] == "前の題"
    r = api.fill(path, {"タイトル": "新しい題 A&B"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert any("文書のプロパティ" in w and "docProps/core.xml" in w for w in r["warnings"]), r["warnings"]
    core = zipfile.ZipFile(str(tmp_path / "o.docx")).read("docProps/core.xml").decode()
    assert core == CORE.replace("前の題", "新しい題 A&amp;B")  # 値の所のほかは、バイトのまま
    assert {f["id"]: f for f in _fields(tmp_path / "o.docx")}["タイトル"]["value"] == "新しい題 A&B"


@pytest.mark.parametrize("xpath,rel", [
    ("/ns0:coreProperties[1]/ns1:subject[1]", True),  # プロパティの部品に、その要素がない
    ("/ns0:coreProperties[1]/ns1:title[1]", False),  # プロパティの部品として、文書につながっていない
])
def test_control_bound_to_a_missing_property_is_refused(tmp_path, xpath, rel):
    path = wf.write(tmp_path, _core_docx(xpath, rel=rel))
    with pytest.raises(api.ApiError) as e:
        api.fill(path, {"タイトル": "新しい題"}, str(tmp_path / "o.docx"))
    assert "w:storeItemID" in str(e.value)
    assert not (tmp_path / "o.docx").exists()


# ---------------------------------------------------------------------- RTF の脚注の印の後ろ
def _w(word: str) -> str:
    return B + word


NOTE = "{" + _w("footnote") + " " + _w("pard") + _w("plain") + " {" + _w("super") + " " + _w("chftn") + "} Does not include meals.}"
FN_END = "{" + _w("b") + " HOSPITALITY}{" + _w("super") + " " + _w("chftn") + " " + NOTE + "}" + _w("par")
FN_MID = "{" + _w("b") + " Name}{" + _w("super") + " " + _w("chftn") + "\r\n" + NOTE + "}{ : }" + _w("par")


def _rtf_text(path) -> str:
    api._WINDOWS.clear()
    lazy.clear()
    return "\n".join(t for _n, t in api.window(str(path))[2])


def test_rtf_write_at_paragraph_end_after_footnote_mark(tmp_path):
    p = tmp_path / "a.rtf"
    p.write_bytes(make_rtf(_w("pard") + " Title" + _w("par"), FN_END))
    fid = next(f["id"] for f in _fields(p) if "HOSPITALITY※" in f["context"])
    r = api.fill(str(p), {fid: "12 GBP"}, str(tmp_path / "o.rtf"))
    assert r["read_back_ok"] and r["proof_ok"], r
    out = (tmp_path / "o.rtf").read_bytes().decode("latin-1")
    at = out.index("12 GBP")
    assert out.index("Does not include meals.}}") < at  # 脚注のグループと、上付きの字のグループの外
    assert out[:at].count("{") - out[:at].count("}") == 1  # 文書のいちばん外のグループの中だけ
    text = _rtf_text(tmp_path / "o.rtf")
    assert f"HOSPITALITY※［{fid}＝12 GBP］" in text and "- Does not include meals." in text  # 脚注の字は、元のまま


def test_rtf_write_between_footnote_mark_and_text(tmp_path):
    p = tmp_path / "a.rtf"
    p.write_bytes(make_rtf(FN_MID))
    r = api.fill(str(p), {"@1": {"old": " :", "new": "（代表） :"}}, str(tmp_path / "o.rtf"))
    assert r["read_back_ok"] and r["proof_ok"], r
    out = (tmp_path / "o.rtf").read_bytes().decode("latin-1")
    body = out[out.index("Does not include meals.}}") + len("Does not include meals.}}"):]
    assert body.startswith("{{" + _w("*") + _w("bkmkstart"))  # 脚注と上付きの字のグループを閉じた後ろ、次の字のグループの頭
    text = _rtf_text(tmp_path / "o.rtf")
    assert "- Name※［P1.w1＝（代表）］ :" in text and "- Does not include meals." in text  # 脚注の字は、元のまま


# ---------------------------------------------------------------------- .doc に埋め込んだ部品の種類
class _Doc:
    """DocFile._add_embed が使うものだけ（複合ファイルと、埋め込みの一覧）。"""

    def __init__(self, streams: dict[str, bytes]):
        from todoku.cfb import CFB, build_cfb

        self.cfb = CFB(build_cfb({"WordDocument": bytes(4096), **streams}))
        self.embeds: list[dict] = []


@pytest.mark.parametrize("word,stored,want", [
    ("Γράφημα", "Excel.Chart.8", "Excel.Chart.8"),  # 命令の語が、その言語の Word の名前
    ("Excel.Sheet.8", "", "Excel.Sheet.8"),  # 入れ物が名乗らない：命令の語のまま
    ("Excel.Sheet.8", "表 の 名前", "Excel.Sheet.8"),  # ProgID の形でない字：命令の語のまま
])
def test_doc_embed_kind_comes_from_the_storage(word, stored, want):
    from todoku.doc.book import DocFile

    streams = {"ObjectPool/_1/Workbook": b"x" * 16}
    if stored:
        streams["ObjectPool/_1/\x01CompObj"] = compobj_stream("Microsoft Excel", stored)
    doc = _Doc(streams)
    e = DocFile._add_embed(doc, "ObjectPool/_1", word, None)
    assert e["prog"] == want
    assert DocFile._add_embed(doc, "ObjectPool/_1", word, None) is e  # 同じ入れ物は 1 度だけ
    assert DocFile._add_embed(doc, "ObjectPool/_9", word, None) is None  # 入れ物がない
