"""印でも名前でも形式が決まらないファイル：中身が字だけなら、文字のファイル（txt）・HTML/XML（markup）として読み、
形式を中身から決めたことを窓と段階に残す（book.content_kind）。"""

from __future__ import annotations

import io
import os
import random
import zipfile

import pytest

from todoku import api, lazy, limits
from todoku.book import _ALWAYS_CONTROL, content_kind, kind_of
from todoku.datafile.textenc import decode
from todoku.lazy.textstream import MemorySource

FORM = "様式第１号\r\n\r\n申請者　氏名（　　　）\r\n電話番号：\r\n以上\r\n"


def _kind(blob: bytes) -> str | None:
    return content_kind(MemorySource(blob))


# ---------------------------------------------------------------------- 決まり
@pytest.mark.parametrize("blob", [
    b"hello\n",
    FORM.encode("utf-8"),
    b"\xef\xbb\xbf" + FORM.encode("utf-8"),
    b"\xff\xfe" + FORM.encode("utf-16-le"),
    b"\xfe\xff" + FORM.encode("utf-16-be"),
    ("name: a\r\n" + FORM).encode("utf-16-le"),  # BOM のない UTF-16（ASCII の字が NUL と組になる）
    FORM.encode("cp932"),
    FORM.encode("euc_jp"),
    "성명\r\n주소\r\n".encode("cp949"),
    "Nom (Position) étiquette\r\n".encode("cp1252"),
    FORM.encode("iso2022_jp"),  # ESC は字の集合の切り替え
    b"page 1\x0cpage 2\n",  # 改ページは空白
])
def test_text_in_any_supported_encoding_is_text(blob):
    assert _kind(blob) == "txt"
    assert kind_of(blob, "a.dat") == "txt" and kind_of(blob, "README") == "txt"


@pytest.mark.parametrize("blob", [
    b"",
    b"abc\x00def\n",  # NUL のある ASCII
    b"abc\x01def\n",  # 制御文字
    b"abc\x7fdef\n",  # DEL
    b"\x1b[31mred\x1b[0m\n",  # 色の指定の ESC は、切り替えではない：制御文字
    "日本語\x00\n".encode("utf-8"),  # UTF-8 の字の中の NUL
    b"\xff\xfe" + b"\x00\xd8\x41\x00",  # BOM の文字コードとして読めない
    bytes(range(256)) * 4,
    b"\x00" * 64,
])
def test_bytes_with_controls_are_not_text(blob):
    assert _kind(blob) is None
    assert kind_of(blob, "a.dat") == "unknown"


def test_without_a_name_the_kind_is_as_before():
    assert kind_of(b"hello") == "unknown" and kind_of(FORM.encode("cp932")) == "unknown"


@pytest.mark.parametrize("blob,kind", [
    (b"<?xml version='1.0'?>\n<a><b>x</b></a>", "markup"),
    (b"\xef\xbb\xbf\r\n  <?XML version='1.0'?><a/>", "markup"),
    (b"<!DOCTYPE html>\n<html><body><p>x</p></body></html>", "markup"),
    (b"<!doctype HTML PUBLIC \"-//W3C//DTD HTML 4.01//EN\"><html></html>", "markup"),
    (b"  <HTML xmlns:o=\"urn:schemas-microsoft-com:office:office\"><table><tr><td>1</td></tr></table></HTML>", "markup"),
    (b"\xff\xfe" + "<?xml version='1.0' encoding='UTF-16'?><a>字</a>".encode("utf-16-le"), "markup"),
    (b"<table><tr><td>1</td></tr></table>", "txt"),  # 宣言も <html もない：タグのある字として、そのまま読む
    (b"<!DOCTYPE plist PUBLIC \"x\"><plist/>", "txt"),
    (b"x <?xml version='1.0'?>", "txt"),
])
def test_markup_only_by_its_own_declaration(blob, kind):
    assert kind_of(blob, "a.xls") == kind


def test_marks_and_names_still_decide_first():
    assert kind_of(b"%PDF-1.4\n1 0 obj\n<<>>\nendobj\n", "a.dat") == "pdf"
    assert kind_of(b"{\\rtf1 abc}", "a.dat") == "rtf"
    assert kind_of(b"a,b\n1,2\n", "a.csv") == "csv" and kind_of(b"a,b\n1,2\n", "a.dat") == "txt"
    assert kind_of(b'{"a": 1}', "a.json") == "json" and kind_of(b"hello", "a.txt") == "txt"
    assert kind_of(b"<?xml version='1.0'?><svg xmlns='http://www.w3.org/2000/svg'/>", "a.dat") == "image"
    assert kind_of(b"hello", "a.xdw") == "docuworks"


def _plain_rule(blob: bytes) -> bool:
    """先の安い確かめを通さない、決まりそのもの：.txt の決め方で読めて、制御文字がない。"""
    if not blob:
        return False
    try:
        d = decode(blob)
    except UnicodeDecodeError:
        return False
    return d.basis != "none" and not d.controls and not d.binary


def test_cheap_checks_do_not_change_the_rule():
    rnd = random.Random(20261112)
    words = ["申請", "氏名", "Ａ", "abc", " ", "\r\n", "\t", "、", "é", "성명", "1"]
    seen = {True: 0, False: 0}
    for i in range(1500):
        text = "".join(rnd.choice(words) for _ in range(rnd.randint(1, 30)))
        enc = rnd.choice(["utf-8", "utf-16-le", "utf-16-be", "utf-32-le", "cp932", "euc_jp", "cp949", "gbk", "big5", "cp1252", "iso2022_jp"])
        try:
            blob = text.encode(enc)
        except UnicodeEncodeError:
            blob = text.encode("utf-8")
        if rnd.random() < 0.2 and enc.startswith("utf-16"):
            blob = (b"\xff\xfe" if enc.endswith("le") else b"\xfe\xff") + blob
        r = rnd.random()
        if r < 0.3:  # どこかに 1 バイトを足す（制御文字・NUL・高いバイト）
            at = rnd.randint(0, len(blob))
            blob = blob[:at] + bytes([rnd.choice([0, 1, 8, 0x0b, 0x0e, 0x1a, 0x1b, 0x7f, 0x80, 0xff, 0x41])]) + blob[at:]
        elif r < 0.35:
            blob = bytes(rnd.randrange(256) for _ in range(rnd.randint(1, 40)))
        want = _plain_rule(blob)
        assert (_kind(blob) is not None) == want, (i, enc, blob[:60])
        seen[want] += 1
    assert seen[True] > 300 and seen[False] > 300, seen


def test_always_control_bytes_are_never_part_of_a_character():
    # 先の確かめの前提：ASCII と同じバイトを使う文字コードの字に、いつも制御文字とするバイトは出ない
    import sys

    for enc in ("utf-8", "cp932", "euc_jp", "cp950", "big5hkscs", "gbk", "gb18030", "cp949", "johab", "cp1252", "iso8859_15", "koi8_r"):
        for cp in list(range(0x20, 0x7f)) + list(range(0xa0, 0x3000, 7)) + list(range(0x3000, 0x9fff, 13)) + list(range(0xac00, 0xd7a3, 17)):
            try:
                b = chr(cp).encode(enc)
            except UnicodeEncodeError:
                continue
            assert not _ALWAYS_CONTROL.search(b), (enc, hex(cp), b)
    assert sys.maxunicode == 0x10ffff


# ---------------------------------------------------------------------- 窓・段階・書き込み
def test_window_and_grade_say_the_kind_came_from_the_content(tmp_path):
    p = tmp_path / "form.doc"
    p.write_bytes(FORM.encode("utf-8"))
    s = api.summary(str(p))
    assert s["kind"] == "txt" and s["format"]["kind_by"] == "content" and ".doc" in s["format"]["kind_note"]
    k, doc, parts = api.window(str(p))
    assert k == "txt" and "名前の終わり（.doc）からは、形式が決まりません" in parts[0][1] and "\n3｜申請者" in parts[0][1]
    g = api.grade(str(p))
    assert g["grade"] == "B" and g["write"]["grade"] == "B"
    assert "形式を中身から決めた" in [r["title"] for r in g["reasons"]] and "形式を中身から決めた" in [r["title"] for r in g["write"]["reasons"]]
    # 同じ中身の .txt：名前で決まるので、知らせも、この理由もない
    q = tmp_path / "form.txt"
    q.write_bytes(FORM.encode("utf-8"))
    s2 = api.summary(str(q))
    assert "kind_by" not in s2["format"] and "形式が決まりません" not in api.window(str(q))[2][0][1]
    g2 = api.grade(str(q))
    assert "形式を中身から決めた" not in [r["title"] for r in g2["reasons"] + g2["write"]["reasons"]]
    assert [f["id"] for f in api.fields(str(p))] == [f["id"] for f in api.fields(str(q))]


def test_name_without_extension(tmp_path):
    p = tmp_path / "README"
    p.write_bytes(b"read me\n")
    s = api.summary(str(p))
    assert s["kind"] == "txt" and s["format"]["kind_note"].startswith("ファイルの頭の印と名前からは、形式が決まりません")


def test_fill_changes_only_the_written_bytes(tmp_path):
    p = tmp_path / "form.doc"
    blob = FORM.encode("cp932")
    p.write_bytes(blob)
    out = tmp_path / "out.doc"
    r = api.fill(str(p), {"@1": {"old": "電話番号：", "new": "電話番号：03-0000-0000"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert out.read_bytes() == blob.replace("電話番号：".encode("cp932"), "電話番号：03-0000-0000".encode("cp932"))
    assert api.summary(str(out))["format"]["kind_by"] == "content"


def test_markup_by_content_window_and_grade(tmp_path):
    p = tmp_path / "list.xls"
    p.write_bytes("<html><head><title>t</title></head><body><table><tr><td>氏名</td><td>山田</td></tr></table></body></html>".encode("utf-8"))
    k, doc, parts = api.window(str(p))
    assert k == "markup" and doc["format"]["kind_by"] == "content" and "氏名" in doc["lines"] and "山田" in doc["lines"]
    assert "名前の終わり（.xls）からは、形式が決まりません" in parts[0][1]
    g = api.grade(str(p))
    assert g["grade"] == "B" and g["write"]["grade"] == "D" and "形式を中身から決めた" in [r["title"] for r in g["reasons"]]


def test_binary_with_another_name_is_still_refused(tmp_path):
    p = tmp_path / "a.dat"
    p.write_bytes(b"abc\x00\x01\x02" * 50)
    with pytest.raises(api.ApiError, match="unknown"):
        api.summary(str(p))


def test_member_of_an_archive(tmp_path):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("README", "read me\n")
        z.writestr("notes.md", "# 手順\n")
        z.writestr("data.bin", b"\x00\x01\x02\x03")
    p = tmp_path / "a.zip"
    p.write_bytes(buf.getvalue())
    kinds = {m["name"]: m["kind"] for m in api.summary(str(p))["members"]}
    assert kinds["README"] == "txt" and kinds["notes.md"] == "txt" and kinds["data.bin"] == ""
    member = str(p) + "!/README"
    assert api.summary(member)["format"]["kind_by"] == "content"
    out = tmp_path / "out.zip"
    r = api.fill(member, {"@1": {"old": "read me", "new": "read me!"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    with zipfile.ZipFile(out) as z:
        assert z.read("README") == b"read me!\n" and z.read("data.bin") == b"\x00\x01\x02\x03"


# ---------------------------------------------------------------------- 大きなファイル（範囲で読む道）
def test_large_text_with_another_name_is_read_in_ranges(tmp_path, monkeypatch):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    monkeypatch.setattr(limits, "PAGE_CHARS", 1_000)
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    try:
        p = tmp_path / "a.log"
        p.write_text("".join(f"第{i}行 氏名（　　　）の申請です。\n" for i in range(1, 401)), encoding="utf-8")
        s = api.summary(str(p))
        assert s["kind"] == "txt" and s["large"] and s["format"]["kind_by"] == "content" and s["lines"] == 400
        doc = lazy.open_doc(str(p))
        assert "名前の終わり（.log）からは、形式が決まりません" in doc.overview_text()
        page, _d = api.read_range(str(p), {"paragraphs": "100-101"}, None)
        assert page.body[0].startswith("100｜第100行")
        g = api.grade(str(p))
        assert g["grade"] == "B" and "形式を中身から決めた" in [r["title"] for r in g["reasons"]]
        b = tmp_path / "b.bin"
        b.write_bytes(os.urandom(4096) * 4)
        with pytest.raises(api.ApiError):
            api.summary(str(b))
    finally:
        lazy.clear()
        api._WINDOWS.clear()
