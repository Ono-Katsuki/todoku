"""書庫の中の PDF を、PDF そのものと同じ読み手で読む：ページの範囲（pages）・大きな PDF の概要・欄・段階・検査。

前は、書庫の中の PDF は、窓の全体を作ってから分けていた：pages で頼むと断り、大きな PDF は、開くたびに全部のページを配置していた。
いまは、取り出した中身を、ページの木だけを読む読み手（PdfLazy）に渡す：結果は、取り出したファイルを読んだ結果と同じ。
"""
import json
import zipfile
from pathlib import Path

import pytest

from test_lazy import _make_pdf
from todoku import api, lazy, limits, mcp_server

PUBLIC = Path(__file__).parent / "fixtures" / "public"


def _zip(tmp_path, name: str, blob: bytes, zname: str = "z.zip") -> str:
    with zipfile.ZipFile(tmp_path / zname, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(name, blob)
    return f"{tmp_path / zname}!/{name}"


@pytest.fixture()
def box(tmp_path):
    f = _make_pdf(tmp_path / "a.pdf", 12)
    yield str(f), _zip(tmp_path, "a.pdf", f.read_bytes())
    lazy.clear()  # Windows は、控えが開いたままのファイルを消せない


def _tool(name: str, f: str, m: str, path: str, **a) -> str:
    """道具の結果（書庫の中の道筋の字を、ファイルの道筋にそろえる）。"""
    r = mcp_server.call_tool(name, dict(a, path=path))
    return json.dumps(r, ensure_ascii=False, sort_keys=True, default=str).replace(json.dumps(m)[1:-1], json.dumps(f)[1:-1])


def _page(path: str, rng: dict, n: int = 4000, **kw):
    page, doc = api.read_range(path, rng, n, **kw)
    return page.text, page.next, page.covered, type(doc).__name__


def test_pages_of_a_pdf_in_an_archive(box):
    f, m = box
    for rng in ({"pages": "7-7"}, {"pages": "3-5"}, {"pages": "12"}, {}):
        inside, outside = _page(m, rng), _page(f, rng)
        assert inside == outside and inside[3] == "PdfLazy"
    assert "Page 7 line 0" in _page(m, {"pages": "7-7"})[0] and "Page 8" not in _page(m, {"pages": "7-7"})[0]
    # 続き（next）も、ファイルそのものと同じ所から
    first = api.read_range(m, {"pages": "1-12"}, 600)[0]
    assert first.next and first.next == api.read_range(f, {"pages": "1-12"}, 600)[0].next
    assert _page(m, first.next, 600) == _page(f, first.next, 600)
    assert api.fields_range(m, {"pages": "2-3"}) == api.fields_range(f, {"pages": "2-3"})
    with pytest.raises(api.ApiError, match="13 ページ目はありません"):
        api.read_range(m, {"pages": "13"}, 500)
    # PDF にない範囲の名前は、ファイルそのものと同じに断る
    for rng in ({"rows": "1-5"}, {"paragraphs": "1-5"}, {"part": 2}):
        with pytest.raises(api.ApiError) as a:
            api.read_range(m, rng, 500)
        with pytest.raises(api.ApiError) as b:
            api.read_range(f, rng, 500)
        assert str(a.value) == str(b.value)


def test_a_small_pdf_in_an_archive_is_read_by_pages(tmp_path):
    f = str(PUBLIC / "flat.pdf")  # 2 ページ：大きな書類ではない
    m = _zip(tmp_path, "flat.pdf", Path(f).read_bytes())
    assert _page(m, {"pages": "2"}) == _page(f, {"pages": "2"})
    assert api.fields_range(m, {"pages": "1"}) == api.fields_range(f, {"pages": "1"})
    for tool in ("open_document", "read_window", "list_fields", "grade_document"):
        assert _tool(tool, f, m, m) == _tool(tool, f, m, f)
    lazy.clear()


def test_a_large_pdf_in_an_archive_reads_like_the_file(box, monkeypatch):
    f, m = box
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    res = mcp_server.tool_result("open_document", {"path": m})
    assert res["structuredContent"]["large"] and res["structuredContent"]["pages"] == 12
    for tool, extra in (("open_document", {}), ("read_window", {}), ("list_fields", {}), ("grade_document", {}),
                        ("read_window", {"pages": "11-12"}), ("list_fields", {"pages": "4"})):
        assert _tool(tool, f, m, m, **extra) == _tool(tool, f, m, f, **extra), tool
    # 段階：全部のページを配置しない（未確認は未確認と伝える）
    g = api.grade(m)
    assert g == api.grade(f) and any("未確認" in x for x in g["deferred"])


def test_the_audit_of_a_pdf_in_an_archive_goes_page_by_page(box, monkeypatch):
    """検査（長い仕事）の中身：ページごとに確かめ、書き込みの試しは、書庫の中のファイルに書く道を通る。"""
    f, m = box
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)

    def run(path: str) -> dict:
        doc = lazy.open_doc(path, force=True)
        assert type(doc).__name__ == "PdfLazy"
        state = doc.audit_begin()
        steps = 0
        while not state.get("done"):
            state = doc.audit_step(state)
            steps += 1
        return {"steps": steps, **state}

    inside, outside = run(m), run(f)
    assert inside == outside and inside["steps"] >= 12 and inside["chars"] > 0


def test_a_pdf_in_an_archive_in_an_archive(tmp_path):
    f = _make_pdf(tmp_path / "a.pdf", 7)
    inner = tmp_path / "inner.zip"
    with zipfile.ZipFile(inner, "w") as z:
        z.write(f, "a.pdf")
    m = _zip(tmp_path, "inner.zip", inner.read_bytes(), "outer.zip") + "!/a.pdf"
    assert _page(m, {"pages": "6-7"}) == _page(str(f), {"pages": "6-7"})
    lazy.clear()


def test_a_password_pdf_in_an_archive(tmp_path):
    f = str(PUBLIC / "encrypted.pdf")
    m = _zip(tmp_path, "encrypted.pdf", Path(f).read_bytes())
    with pytest.raises(api.PasswordNeeded):
        api.read_range(m, {"pages": "1"}, 2000)
    with pytest.raises(api.ApiError):
        api.read_range(m, {"pages": "1"}, 2000, password="wrong")
    inside = _page(m, {"pages": "1"}, password="todoku")
    assert inside == _page(f, {"pages": "1"}, password="todoku") and "補助金交付申請書" in inside[0]
    # パスワードのある書類は控えない（パスワードを控えに残さない）
    assert not [k for k in lazy._CACHE if k and k[0] == "member"]


def test_a_pdf_in_an_archive_is_kept_only_while_small(box, monkeypatch):
    f, m = box
    doc = lazy.open_doc(m, force=True)
    assert lazy.open_doc(m, force=True) is doc  # 呼び出しのたびに取り出さない
    lazy.clear()
    monkeypatch.setattr(lazy, "_MEMBER_KEEP", 100)  # これより大きな中身は、メモリに控えない
    doc = lazy.open_doc(m, force=True)
    assert type(doc).__name__ == "PdfLazy" and not [k for k in lazy._CACHE if k and k[0] == "member"]
    assert _page(m, {"pages": "9"}) == _page(f, {"pages": "9"})


def test_writing_into_a_pdf_in_an_archive_is_unchanged(tmp_path, monkeypatch):
    """書き込みは、前と同じに、取り出したファイルに書いて書庫に入れ直す（大きな PDF でも）。"""
    f = PUBLIC / "acroform.pdf"
    m = _zip(tmp_path, "acroform.pdf", f.read_bytes())
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 200)
    fid = next(x["id"] for x in api.fields(str(f)) if x.get("kind") == "text")
    res = api.fill(m, {fid: "架空 太郎"}, str(tmp_path / "out.zip"))
    assert res["proof_ok"] and res["read_back_ok"] and not res["unexpected"]
    with zipfile.ZipFile(res["out_path"]) as z:
        written = z.read("acroform.pdf")
    (tmp_path / "w.pdf").write_bytes(written)
    assert written.startswith(f.read_bytes()) and any(x.get("value") == "架空 太郎" for x in api.fields(str(tmp_path / "w.pdf")))
    lazy.clear()
