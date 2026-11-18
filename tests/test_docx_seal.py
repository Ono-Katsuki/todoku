"""Word（.docx）に、本人の印影・署名の画像を置く（本人が同意したときだけ）：字を目印に置く・証明・読み戻し。

画像は、テストの中で作った図形（輪と四角）。実在の人の印影は使わない。
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from todoku import api, sealimage
from todoku.docx import seal as ds

from test_docx import make_docx, p
from test_seal import png_bytes, ring_rgba

BODY = (p("請求書") + p("株式会社とどく　代表取締役　山田太郎　印") + p("金額　100,000 円") + p("担当　鈴木花子")
        + "<w:tbl><w:tr><w:tc>" + p("確認者") + "</w:tc><w:tc>" + p("佐藤次郎　印") + "</w:tc></w:tr></w:tbl>")
MARK = "〔図: 本人の印影の画像（本人の同意を得て置いたもの）〕"


def _files(tmp_path, body: str = BODY, opaque: bool = False):
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(body))
    png = tmp_path / "ring.png"
    if opaque:  # アルファのない、赤い四角
        png.write_bytes(png_bytes(8, 8, 2, [bytes((200, 20, 20)) * 8] * 8))
    else:
        png.write_bytes(ring_rgba()[0])
    return src, png


def _spec(png, **kw) -> dict:
    return {"seal": str(png), "consent": True, "size": 12, **kw}


def _parts(path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def test_three_seals_share_one_image_part_and_only_add(tmp_path):
    src, png = _files(tmp_path)
    out = tmp_path / "b.docx"
    res = api.fill(str(src), {"@1": _spec(png, cover="印", near="山田太郎"), "@2": _spec(png, after="鈴木花子", size=10),
                              "@3": _spec(png, cover="印", near="佐藤次郎")}, str(out))
    assert res["proof_ok"] and res["read_back_ok"] and res["unexpected"] == []
    assert [pl["id"] for pl in res["placed"]] == ["S1", "S2", "S3"]
    assert [pl["where"] for pl in res["placed"]] == ["cover", "after", "cover"]
    assert res["placed"][1]["size_mm"] == [10.0, 10.0] and res["placed"][0]["size_mm"] == [12.0, 12.0]
    assert "「印」の上に、本人の印影の画像を置きました" in res["say_to_user"]
    assert "「鈴木花子」のすぐ後ろに、本人の印影の画像を置きました" in res["say_to_user"]
    assert ds.UNCERTAIN in res["warnings"] and sealimage.caution("印影") in res["warnings"]
    # 表のマスの中の字の上に置いた画像は、行の上にはみ出すことを伝える
    assert "表のマス" in res["placed"][2]["uncertain"] and "表のマス" not in res["placed"][0]["uncertain"]
    a, b = _parts(src), _parts(out)
    # 足したのは、画像の部品 1 つ（3 つの印で共有）と、その関係の部品。種類の表には png を足した
    assert set(b) - set(a) == {"word/media/tdkseal1.png", "word/_rels/document.xml.rels"}
    assert b["word/media/tdkseal1.png"] == png.read_bytes()
    assert b["word/_rels/document.xml.rels"].count(b"media/tdkseal1.png") == 1
    assert b'<Default Extension="png" ContentType="image/png"/>' in b["[Content_Types].xml"]
    assert b["_rels/.rels"] == a["_rels/.rels"]
    # 本文は、画像の図形を足しただけ：図形を除くと、字は元のまま
    assert b["word/document.xml"].count(b"<wp:anchor") == 3
    stripped = re.sub(rb"<w:r><w:drawing>.*?</w:drawing></w:r>", b"", b["word/document.xml"])
    assert re.sub(rb"<[^>]*>", b"", stripped) == re.sub(rb"<[^>]*>", b"", a["word/document.xml"])
    # 読み直した窓に、何の画像かが字で出る（目印の字の直前）
    _k, win, _p = api.window(str(out))
    reading = "\n".join(win["stories"][0]["reading"])
    assert reading.count(MARK) == 3 and f"{MARK}印" in reading and f"{MARK}鈴木花子" in reading


def test_opaque_image_goes_behind_the_text_and_keeps_ratio(tmp_path):
    src, _ = _files(tmp_path)
    wide = tmp_path / "wide.png"
    wide.write_bytes(png_bytes(20, 10, 2, [bytes((200, 20, 20)) * 20] * 10))
    out = tmp_path / "b.docx"
    res = api.fill(str(src), {"@1": {"signature": str(wide), "consent": True, "after": "鈴木花子", "size": 30}}, str(out))
    assert res["proof_ok"] and res["read_back_ok"]
    pl = res["placed"][0]
    assert pl["what"] == "署名" and pl["behind_text"] is True and pl["size_mm"] == [30.0, 15.0]
    assert "字の後ろに置きました" in pl["say"]
    xml = _parts(out)["word/document.xml"]
    assert b'behindDoc="1"' in xml and b'descr="\xe6\x9c\xac\xe4\xba\xba\xe3\x81\xae\xe7\xbd\xb2\xe5\x90\x8d' in xml
    m = re.search(rb'<wp:extent cx="(\d+)" cy="(\d+)"', xml)
    # 大きさは pt に丸めてから EMU にするので、数 EMU（1 EMU = 1/36000 mm）ずれる
    assert abs(int(m[1]) - 30 * 36000) <= 20 and abs(int(m[2]) - 15 * 36000) <= 20


def test_values_and_seal_in_one_call_and_a_second_seal_later(tmp_path):
    src, png = _files(tmp_path, p("氏名：［　　　］　印") + p("日付　令和　年"))
    mid, out = tmp_path / "b.docx", tmp_path / "c.docx"
    res = api.fill(str(src), {"@1": {"old": "令和　年", "new": "令和8年"}, "@2": _spec(png, cover="印")}, str(mid))
    assert res["proof_ok"] and res["read_back_ok"], (res["unexpected"], res["read_back_differences"])
    assert "令和8年" in re.sub(r"<[^>]*>", "", _parts(mid)["word/document.xml"].decode())
    # 置いた後のファイルに、もう 1 つ置く：番号も、画像の部品の名前も、関係の id も重ならない
    other = tmp_path / "sq.png"
    other.write_bytes(png_bytes(8, 8, 2, [bytes((20, 20, 200)) * 8] * 8))
    res2 = api.fill(str(mid), {"@1": {"seal": str(other), "consent": True, "after": "日付", "size": 8}}, str(out))
    assert res2["proof_ok"] and res2["read_back_ok"], (res2["unexpected"], res2["read_back_differences"])
    assert res2["placed"][0]["id"] == "S2"
    parts = _parts(out)
    assert {"word/media/tdkseal1.png", "word/media/tdkseal2.png"} <= set(parts)
    rels = parts["word/_rels/document.xml.rels"].decode()
    assert 'Id="rIdTdkSeal1"' in rels and 'Id="rIdTdkSeal2"' in rels
    ids = re.findall(rb'<wp:docPr id="(\d+)"', parts["word/document.xml"])
    assert len(ids) == 2 and len(set(ids)) == 2


def test_existing_rels_and_a_jpeg_named_default_are_kept(tmp_path):
    # 関係の部品がすでにあり、種類の表に png がほかの種類で登録されている文書：関係は足すだけ、種類は、その部品だけの Override
    blob = make_docx(p("氏名　山田太郎　印"))
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as zi, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if n == "[Content_Types].xml":
                d = d.replace(b"</Types>", b'<Default Extension="png" ContentType="application/x-other"/></Types>')
            zo.writestr(n, d)
        zo.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://example.com/x" Target="media/tdkseal1.png"/></Relationships>')
        zo.writestr("word/media/tdkseal1.png", b"not ours")
    src = tmp_path / "a.docx"
    src.write_bytes(buf.getvalue())
    png = tmp_path / "ring.png"
    png.write_bytes(ring_rgba()[0])
    out = tmp_path / "b.docx"
    res = api.fill(str(src), {"@1": _spec(png, cover="印")}, str(out))
    assert res["proof_ok"] and res["read_back_ok"], (res["unexpected"], res["read_back_differences"])
    parts = _parts(out)
    assert parts["word/media/tdkseal1.png"] == b"not ours" and parts["word/media/tdkseal2.png"] == png.read_bytes()
    rels = parts["word/_rels/document.xml.rels"].decode()
    assert 'Id="rId1" Type="http://example.com/x"' in rels and 'Target="media/tdkseal2.png"' in rels
    assert b'<Override PartName="/word/media/tdkseal2.png" ContentType="image/png"/>' in parts["[Content_Types].xml"]


@pytest.mark.parametrize("spec, said", [
    ({"cover": "印"}, "同意"),  # consent がない
    ({"consent": True, "cover": "印"}, '"size"'),
    ({"consent": True, "size": 12}, "字の目印"),
    ({"consent": True, "size": 12, "page": 1, "box": [10, 10, 30, 30]}, "ページと枠"),
    ({"consent": True, "size": 12, "after": "印", "where": "below"}, '"where" は使えません'),
    ({"consent": True, "size": 12, "cover": "印", "after": "印"}, "どちらか1つ"),
    ({"consent": True, "size": 12, "cover": "割印"}, "割印"),  # 文書にない字
])
def test_refusals_write_nothing(tmp_path, spec, said):
    src, png = _files(tmp_path)
    out = tmp_path / "b.docx"
    if spec.get("cover") == "印" and "after" not in spec and "size" in spec:
        spec = {**spec, "near": "山田太郎"}
    with pytest.raises(api.ApiError, match=said):
        api.fill(str(src), {"@1": {"seal": str(png), **spec}}, str(out))
    assert not out.exists()


def test_same_mark_twice_is_refused_and_ambiguous_mark_lists_candidates(tmp_path):
    src, png = _files(tmp_path)
    out = tmp_path / "b.docx"
    with pytest.raises(api.ApiError, match="同じ所に、画像を 2 つ"):
        api.fill(str(src), {"@1": _spec(png, cover="印", near="山田太郎"), "@2": _spec(png, cover="印", near="山田太郎")}, str(out))
    with pytest.raises(api.ApiError):  # 「印」は 2 か所にある：near か occurrence がなければ決めない
        api.fill(str(src), {"@1": _spec(png, cover="印")}, str(out))
    assert not out.exists()
    # 同じ字の上と後ろに 1 つずつは置ける（どちらの図形も入る）
    res = api.fill(str(src), {"@1": _spec(png, cover="鈴木花子"), "@2": _spec(png, after="鈴木花子")}, str(out))
    assert res["proof_ok"] and res["read_back_ok"]
    assert _parts(out)["word/document.xml"].count(b"<wp:anchor") == 2


def test_proof_and_read_back_catch_changes(tmp_path):
    src, png = _files(tmp_path)
    out = tmp_path / "b.docx"
    spec = _spec(png, cover="印", near="山田太郎")
    res = api.fill(str(src), {"@1": spec}, str(out))
    img = sealimage.decode(png.read_bytes())
    before, after = src.read_bytes(), out.read_bytes()
    assert ds.prove(before, after, "a.docx", [("S1", spec, "seal", img)]) == []

    def rewrite(change) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(after)) as zi, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in zi.namelist():
                zo.writestr(n, change(n, zi.read(n)))
        return buf.getvalue()

    # 本文のほかの字が変わった・画像の部品が別の画像になった：証明が見つける
    bad = rewrite(lambda n, d: d.replace("100,000".encode(), "900,000".encode()) if n == "word/document.xml" else d)
    assert any("word/document.xml" in x for x in ds.prove(before, bad, "a.docx", [("S1", spec, "seal", img)]))
    swapped = rewrite(lambda n, d: png_bytes(8, 8, 2, [bytes((0, 0, 0)) * 8] * 8) if n.endswith("tdkseal1.png") else d)
    assert any("tdkseal1.png" in x for x in ds.prove(before, swapped, "a.docx", [("S1", spec, "seal", img)]))
    # 読み戻し：画像が別の字の前に移った・画像の中身が違う・画像がない
    pl = res["placed"]
    assert ds.read_back(after, "a.docx", pl, {"S1": img}) == []
    assert "一致しません" in ds.read_back(swapped, "a.docx", pl, {"S1": img})[0]["read_back"]
    xml = _parts(out)["word/document.xml"]
    run = re.search(rb"<w:r><w:drawing>.*?</w:drawing></w:r>", xml)[0]
    moved = rewrite(lambda n, d: d.replace(run, b"").replace("<w:r><w:t xml:space=\"preserve\">金額".encode(),
                                                             run + "<w:r><w:t xml:space=\"preserve\">金額".encode())
                    if n == "word/document.xml" else d)
    assert "目印の字「印」ではない" in ds.read_back(moved, "a.docx", pl, {"S1": img})[0]["read_back"]
    gone = rewrite(lambda n, d: d.replace(run, b"") if n == "word/document.xml" else d)
    assert "見つかりません" in ds.read_back(gone, "a.docx", pl, {"S1": img})[0]["read_back"]


def test_image_in_a_header_gets_its_own_relationship(tmp_path):
    # ヘッダーの字を目印に置く：関係は、ヘッダーの部品の関係の部品に足す
    blob = make_docx(p("本文") + '<w:sectPr><w:headerReference w:type="default" r:id="rId9" '
                     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>')
    buf = io.BytesIO()
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    with zipfile.ZipFile(io.BytesIO(blob)) as zi, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if n == "[Content_Types].xml":
                d = d.replace(b"</Types>", b'<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-'
                                           b'officedocument.wordprocessingml.header+xml"/></Types>')
            zo.writestr(n, d)
        zo.writestr("word/header1.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {w}>{p("承認　部長")}</w:hdr>')
        zo.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
    src = tmp_path / "a.docx"
    src.write_bytes(buf.getvalue())
    png = tmp_path / "ring.png"
    png.write_bytes(ring_rgba()[0])
    out = tmp_path / "b.docx"
    res = api.fill(str(src), {"@1": _spec(png, after="部長")}, str(out))
    assert res["proof_ok"] and res["read_back_ok"], (res["unexpected"], res["read_back_differences"])
    a, b = _parts(src), _parts(out)
    assert set(b) - set(a) == {"word/media/tdkseal1.png", "word/_rels/header1.xml.rels"}
    assert b["word/_rels/document.xml.rels"] == a["word/_rels/document.xml.rels"] and b["word/document.xml"] == a["word/document.xml"]
    assert b"<wp:anchor" in b["word/header1.xml"]


def test_a_write_todoku_cannot_do_says_to_tell_the_user_before_switching_tools(tmp_path):
    # とどくにまだできない書き込み（Excel に画像を置く）を断るとき：ほかの道具で書くと証明がなくなることを、切り替える前に
    # 本人に伝えるよう、道具の結果に添える。頼み方の誤り（ない id）には添えない
    from pathlib import Path

    from todoku import mcp_server as ms

    src, png = _files(tmp_path)
    xlsx = tmp_path / "a.xlsx"
    xlsx.write_bytes((Path(__file__).parent / "fixtures" / "kami_meisai.xlsx").read_bytes())
    r = ms.tool_result("fill_document", {"path": str(xlsx), "out_path": str(tmp_path / "o.xlsx"),
                                       "values": {"@1": {"seal": str(png), "cell": "A1", "consent": True}}})
    said = r["content"][0]["text"]
    assert r["isError"] and "いまは PDF と Word（.docx）の書類だけ" in said and ms.SWITCH_NOTE in said
    assert "本人に伝えて" in ms.SWITCH_NOTE and "証明" in ms.SWITCH_NOTE
    r = ms.tool_result("fill_document", {"path": str(src), "out_path": str(tmp_path / "o.docx"), "values": {"T9.R9.C9": "x"}})
    assert r["isError"] and ms.SWITCH_NOTE not in r["content"][0]["text"]
    with pytest.raises(api.NoWay):
        api.fill(str(xlsx), {"@1": {"seal": str(png), "cell": "A1", "consent": True}}, str(tmp_path / "o2.xlsx"))
