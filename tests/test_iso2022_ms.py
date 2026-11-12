"""ISO-2022-JP の Windows の拡張（CP50220・CP50221・CP50222）：NEC 特殊文字（13 区：①・㈱）と NEC 選定 IBM 拡張文字
（89〜92 区：髙・﨑）のあるファイル・メールを読み、書く。

- codec（todoku.datafile.iso2022ms）：2 バイトの字を CP932 の表で読む。表の字のハッシュと、決まったバイトで確かめる
  （macOS の iconv の CP50221 と、字のある区点の全部で同じだった表：テストは iconv を呼ばない）。
- 見分け方：標準の ISO-2022-JP の仲間のどれでも読めないファイルだけを、Windows の決まりで読む。
- 書き込み：書いた所のバイトだけを変える。CP932 の表にない字は断る。標準の ISO-2022-JP のファイルに Windows の字を
  書こうとしたときは、断って、encoding に cp50221 を渡す道を伝える。
- メール（.eml）：ISO-2022-JP の名前で送られた Windows の字を、ヘッダー・本文・添付の名前で読む。
どれも、作ったファイルで確かめる。
"""

from __future__ import annotations

import base64
import codecs
import hashlib
import urllib.parse

import pytest

from todoku import api, lazy
from todoku.datafile import iso2022ms, textenc
from todoku.datafile.textenc import EncodingError, decode
from todoku.lazy.textstream import Source, detect
from todoku.legacy import mail

MS = "iso2022_jp_ms"
WIN = "番号,会社,氏名,備考\n①,㈱髙橋,髙﨑 花子,ｶﾅ ¥100\n②,㈲山田,山田 太郎,～∥－‾\n"
# iconv -t CP50221 が書いたバイト（¥ の後ろの数字も ESC ( J のまま）と、CP50222 のバイト（半角カナは SO・SI）
ICONV_21 = (b"\x1b$BHV9f\x1b(B,\x1b$B2q<R\x1b(B,\x1b$B;aL>\x1b(B,\x1b$BHw9M\x1b(B\n\x1b$B-!\x1b(B,\x1b$B-j|b66\x1b(B,"
            b"\x1b$B|byu\x1b(B \x1b$B2V;R\x1b(B,\x1b(I6E\x1b(B \x1b(J\\100\x1b(B\n\x1b$B-\"\x1b(B,\x1b$B-k;3ED\x1b(B,"
            b"\x1b$B;3ED\x1b(B \x1b$BB@O:\x1b(B,\x1b$B!A!B!]\x1b(J~\x1b(B\n")
ICONV_22 = ICONV_21.replace(b"\x1b(I6E\x1b(B", b"\x0e6E\x0f")
NAME = "ISO-2022-JP（Windows の CP50221：NEC 特殊文字・IBM 拡張文字つき）"


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _file(tmp_path, name: str, data: bytes) -> str:
    p = tmp_path / name
    p.write_bytes(data)
    return str(p)


def _fill(tmp_path, name, blob, values, encoding=None):
    p = _file(tmp_path, name, blob)
    out = str(tmp_path / ("out_" + name))
    r = api.fill(p, values, out, encoding=encoding)
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"] and not r["unexpected"], r
    return open(out, "rb").read()


# ---------------------------------------------------------------------- codec
def test_table_is_the_cp932_table_by_jis_row_and_cell():
    dec, enc = iso2022ms.tables()
    assert len(dec) == 94 * 94
    have = [c for c in sorted(dec) if dec[c] != iso2022ms._BAD]
    assert (len(have), len(enc)) == (7336, 7326)  # 同じ字が 2 つの区にあるものが 10
    # 区点の順に並べた字（字のない区点は U+FFFF）：iconv の CP50221 と全部の区点で同じだった表
    assert hashlib.sha256("".join(dec[c] for c in sorted(dec)).encode("utf-8")).hexdigest() == \
        "2a259f3f7dd7ffb9589d3d951f630db57ca5e9b3ebcef79b1c6c0b26d65e2cea"
    for ch, code in enc.items():
        assert dec[(code[0] << 8) | code[1]] == ch


@pytest.mark.parametrize("text, blob", [
    ("①", b"\x1b$B-!\x1b(B"), ("髙", b"\x1b$B|b\x1b(B"), ("ⅰ", b"\x1b$B|q\x1b(B"), ("Ⅰ", b"\x1b$B-5\x1b(B"),
    ("∵", b"\x1b$B\"h\x1b(B"), ("￢", b"\x1b$B\"L\x1b(B"),  # 2 つの区にある字は、区の番号の小さい方
    ("¥", b"\x1b(J\\\x1b(B"), ("‾", b"\x1b(J~\x1b(B"), ("ｶ", b"\x1b(I6\x1b(B"), ("～∥－", b"\x1b$B!A!B!]\x1b(B"),
    ("亜 亜\n亜", b"\x1b$B0!\x1b(B \x1b$B0!\x1b(B\n\x1b$B0!\x1b(B"),  # 空白・改行は ASCII の状態で書く
    ("a¥ｶ亜b", b"a\x1b(J\\\x1b(I6\x1b$B0!\x1b(Bb"), ("", b""), ("abc\t\n", b"abc\t\n"),
])
def test_fixed_bytes_both_ways(text, blob):
    assert text.encode(MS) == blob and blob.decode(MS) == text


@pytest.mark.parametrize("blob, text", [
    (b"\x1b$@-!\x1b(B", "①"),  # ESC $ @ も CP932 の表
    (b"\x1b$B0!", "亜"),  # 終わりの ESC ( B がなくても読む
    (b"\x0e6E\x0f", "ｶﾅ"), (b"\x1b$B\x0e6\x0f0!\x1b(B", "ｶ亜"),  # SO・SI（CP50222）：ESC の印で選んだ集合にかかわらず半角カナ
    (b"\x1b$B0!\n0! 0!\x1b(B", "亜\n亜 亜"),  # 制御文字と空白は、どの状態でもそのままの字
    (b"\x1b(J\\~\x1b(B\\~", "¥‾\\~"), (b"\x1b$B|b|q-5\"h\x1b(B", "髙ⅰⅠ∵"),
    (b"\x1b$B|r\x1b(B", "ⅱ"), (b"\x1b$By!\x1b(B", "纊"), (b"\x1b$B-b\x1b(B", "№"),
])
def test_reads_what_windows_writes(blob, text):
    assert blob.decode(MS) == text
    assert text.encode(MS).decode(MS) == text


@pytest.mark.parametrize("blob, why", [
    (b"\x1b$B.!\x1b(B", "字のない区点"), (b"\x1b$(D0!\x1b(B", "印として読めない"), (b"\x1b(I`\x1b(B", "半角カナにない"),
    (b"\x1b$B0\x1b(B", "途中で終わっている"), (b"\x1b[0m", "印として読めない"), (b"a\x80", "印として読めない"),
    (b"\x1b$B0!0", "途中で終わっている"), (b"\x1b", "印として読めない"),
])
def test_bytes_outside_the_rule_are_refused(blob, why):
    with pytest.raises(UnicodeDecodeError, match=why):
        blob.decode(MS)
    assert "\ufffd" in blob.decode(MS, "replace")


@pytest.mark.parametrize("text", ["〜", "‖", "−", "한", "\x1b", "\x0e", "\x0f", "a😀"])
def test_characters_outside_cp932_are_refused(text):
    with pytest.raises(UnicodeEncodeError):
        text.encode(MS)


def test_incremental_reading_and_writing_equal_the_whole():
    for blob in (ICONV_21, ICONV_22, WIN.encode(MS)):
        d = codecs.getincrementaldecoder(MS)()
        assert "".join(d.decode(blob[i:i + 1]) for i in range(len(blob))) + d.decode(b"", True) == WIN
        assert d.getstate()[0] == b""
        for cut in range(len(blob) + 1):  # どこで切っても、状態を渡せば続きが読める
            a = codecs.getincrementaldecoder(MS)()
            head = a.decode(blob[:cut])
            b = codecs.getincrementaldecoder(MS)()
            b.setstate(a.getstate())
            assert head + b.decode(blob[cut:], True) == WIN
    e = codecs.getincrementalencoder(MS)()
    assert b"".join(e.encode(ch) for ch in WIN) + e.encode("", True) == WIN.encode(MS)
    e = codecs.getincrementalencoder(MS)()
    assert e.encode("亜") == b"\x1b$B0!" and e.getstate() != 0 and e.encode("", True) == b"\x1b(B" and e.getstate() == 0
    assert codecs.lookup("ISO2022-JP-MS").name == MS and codecs.lookup(MS).name == MS


# ---------------------------------------------------------------------- 見分け方
@pytest.mark.parametrize("blob", [ICONV_21, ICONV_22, WIN.encode(MS)])
def test_windows_files_are_recognised_only_when_no_standard_reader_reads(tmp_path, blob):
    d = decode(blob)
    assert (d.encoding, d.basis, d.text, d.candidates) == (MS, "escape", WIN, [])
    src = Source(_file(tmp_path, "s.csv", blob))
    try:
        s = detect(src)
    finally:
        src.close()
    assert (s.encoding, s.basis) == (MS, "escape")
    assert textenc.stateful(MS) and textenc.NAME_JA[MS] == NAME


def test_standard_files_keep_the_standard_reading(tmp_path):
    """標準の読み手で読めるファイルは、前と同じ読み方（〜 は U+301C）。Windows の決まりで読むのは、頼まれたときだけ。"""
    blob = "氏名,〜\n山田,a\n".encode("iso2022_jp")
    assert (decode(blob).encoding, decode(blob).text) == ("iso2022_jp", "氏名,〜\n山田,a\n")
    assert (decode(blob, "cp50221").encoding, decode(blob, "cp50221").text) == (MS, "氏名,～\n山田,a\n")
    kana = "名前,ｶﾅ\n".encode("iso2022_jp_ext")
    assert decode(kana).encoding == "iso2022_jp_ext"
    src = Source(_file(tmp_path, "s.csv", blob))
    try:
        assert detect(src).encoding == "iso2022_jp"
    finally:
        src.close()


def test_names_for_the_windows_code_pages():
    names = ("cp50220", "CP50221", "cp50222", "windows-50220", "windows-50221", "Windows-50222", "ms50220", "ms50221",
             "ms50222", "ISO-2022-JP-MS", "iso2022_jp_ms", NAME)
    assert [textenc.canonical(x) for x in names] == [MS] * len(names)
    assert textenc.read_as(ICONV_21, MS) == WIN
    # 標準の名前で頼まれた字が標準の表で読めないとき：Windows の決まりで読み直す。どちらでも読めないなら None
    assert textenc.read_as(ICONV_21, "iso2022_jp") == WIN
    assert textenc.read_as("あ〜".encode("iso2022_jp"), "iso2022_jp") == "あ〜"
    assert textenc.read_as(b"\x1b$B.!\x1b(B", "iso2022_jp") is None
    assert textenc.read_as(b"\xff", "utf-8") is None and textenc.read_as("é".encode(), "utf-8") == "é"


# ---------------------------------------------------------------------- 窓・段階・範囲
def test_window_grade_fields_and_range_road(tmp_path):
    p = _file(tmp_path, "w.csv", ICONV_21)
    text = "\n".join(x for _n, x in api.window(p)[2])
    assert "| R2 | ① | ㈱髙橋 | 髙﨑 花子 | ｶﾅ ¥100 |" in text and "| R3 | ② | ㈲山田 | 山田 太郎 | ～∥－‾ |" in text
    assert f"文字コード {NAME}・BOM なし" in text and "決められません" not in text
    g = api.grade(p)
    assert (g["grade"], g["write"]["grade"]) == ("S", "A")
    f = {x["id"]: x.get("value") for x in api.fields(p)}
    assert f["R2.C2"] == "㈱髙橋" and f["R3.C4"] == "～∥－‾"
    page, doc = api.read_range(p, {"rows": "2-2"})
    assert "| R2 | ① | ㈱髙橋 | 髙﨑 花子 | ｶﾅ ¥100 |" in page.text and "R3" not in page.text
    assert doc.recoded and doc.told.encoding == MS and "UTF-8 に直した写し" in doc.source_line() and NAME in doc.source_line()
    t = _file(tmp_path, "w.txt", "申請書 №1\n\n氏名：髙橋\n".encode(MS))
    assert "髙橋" in "\n".join(x for _n, x in api.window(t)[2])


# ---------------------------------------------------------------------- 書き込み
def test_fill_changes_only_the_written_cells(tmp_path):
    new = _fill(tmp_path, "w.csv", ICONV_21, {"R2.C1": "⑳", "R3.C3": "﨑 ﾊﾅｺ", "R2.C4": "¥5 ｱ髙A", "R4.C1": "abc"})
    want = (ICONV_21.replace(b"\x1b$B-!\x1b(B,", b"\x1b$B-4\x1b(B,")
            .replace(b"\x1b(I6E\x1b(B \x1b(J\\100\x1b(B\n", b"\x1b(J\\\x1b(B5 \x1b(I1\x1b$B|b\x1b(BA\n")
            .replace(b"\x1b$B;3ED\x1b(B \x1b$BB@O:\x1b(B", b"\x1b$Byu\x1b(B \x1b(IJE:\x1b(B") + b"abc,,,\n")
    assert new == want
    assert new.decode(MS) == (WIN.replace("①", "⑳").replace("山田 太郎", "﨑 ﾊﾅｺ").replace("ｶﾅ ¥100", "¥5 ｱ髙A") + "abc,,,\n")
    assert decode(new).encoding == MS
    # SO・SI で半角カナを書いたファイル（CP50222）：書くカナは ESC ( I で書く（どちらの読み手も読む）
    new = _fill(tmp_path, "w22.csv", ICONV_22, {"R2.C4": "ﾊﾅ"})
    assert new == ICONV_22.replace(b"\x0e6E\x0f \x1b(J\\100\x1b(B\n", b"\x1b(IJE\x1b(B\n")
    new = _fill(tmp_path, "w.txt", "申請書\n\n氏名：\n".encode(MS) + b"\x1b$B-!\x1b(B\n", {"P3.c3": "髙橋 ｶﾅ"})
    assert new.decode(MS) == "申請書\n\n氏名：髙橋 ｶﾅ\n①\n"


@pytest.mark.parametrize("value", ["한", "〜", "😀"])
def test_characters_outside_the_windows_table_are_refused(tmp_path, value):
    p = _file(tmp_path, "w.csv", ICONV_21)
    with pytest.raises(api.ApiError, match="CP50221：NEC 特殊文字・IBM 拡張文字つき））では書けない字"):
        api.fill(p, {"R2.C1": value}, str(tmp_path / "o.csv"))
    assert not (tmp_path / "o.csv").exists()


def test_windows_characters_into_a_standard_file_are_refused_with_the_way_forward(tmp_path):
    """標準の ISO-2022-JP として読んだファイルに ①・髙 を書く：断る（黙って Windows の決まりに変えない）。
    そのファイルが CP50221 としても読めるなら、いまの字が同じかどうかと、encoding に cp50221 を渡す道を伝える。"""
    same = _file(tmp_path, "a.csv", "氏名,あ\n山田,a\n".encode("iso2022_jp"))
    with pytest.raises(api.ApiError) as e:
        api.fill(same, {"R2.C2": "①"}, str(tmp_path / "o.csv"))
    msg = str(e.value)
    assert "（ISO-2022-JP）では書けない字" in msg and "CP50221）にはある字" in msg
    assert "どちらで読んでも同じ" in msg and "encoding に cp50221" in msg
    differs = _file(tmp_path, "b.csv", "氏名,〜\n山田,a\n".encode("iso2022_jp"))
    with pytest.raises(api.ApiError, match="読み方の変わる字があります（〜 と ～ など）"):
        api.fill(differs, {"R2.C2": "髙"}, str(tmp_path / "o.csv"))
    with pytest.raises(api.ApiError) as e:  # CP50221 にもない字：道は伝えない
        api.fill(same, {"R2.C2": "한"}, str(tmp_path / "o.csv"))
    assert "cp50221" not in str(e.value)
    assert not (tmp_path / "o.csv").exists()
    # 伝えた道のとおりに頼むと書ける：できたファイルは、規則で Windows の決まりと見分ける
    new = _fill(tmp_path, "a.csv", "氏名,あ\n山田,a\n".encode("iso2022_jp"), {"R2.C2": "①髙"}, "cp50221")
    assert new == b"\x1b$B;aL>\x1b(B,\x1b$B$\"\x1b(B\n\x1b$B;3ED\x1b(B,\x1b$B-!|b\x1b(B\n"
    assert decode(new).encoding == MS


def test_rewrite_keeps_the_rest_of_a_windows_file():
    blob = "①a髙,ｶ\n".encode(MS)
    d = decode(blob)
    assert d.rewrite(blob, [(1, 2, "②"), (5, 5, "¥")]).decode(MS) == "①②髙,ｶ¥\n"
    with pytest.raises(EncodingError, match="書けない字"):
        d.rewrite(blob, [(0, 1, "〜")])


# ---------------------------------------------------------------------- メール（.eml）
def _b(s: str) -> bytes:
    return b"=?ISO-2022-JP?B?" + base64.b64encode(s.encode(MS)) + b"?="


def test_eml_with_windows_characters_under_the_iso2022jp_name():
    """Windows のメールソフトは、①・髙 を、ISO-2022-JP の名前のまま送る。標準の email は、ヘッダーの符号化語を解かずに残し、
    本文の字を「�」にする：元のバイトから読み直す。"""
    raw = (b"From: " + _b("髙橋 ㈱") + b" <a@example.com>\r\nTo: =?iso-2022-jp?Q?=1B$B;3ED=1B(B?= <x@example.com>\r\n"
           b"Subject: " + _b("①髙橋") + b"\r\n " + _b("です") + b" (re)\r\n"
           b"Content-Type: text/plain; charset=iso-2022-jp\r\n\r\n" + "①の件、髙橋です。～\r\n".encode(MS))
    m = mail.read_eml(raw)
    assert dict(m.headers)["From"] == "髙橋 ㈱ <a@example.com>" and dict(m.headers)["To"] == "山田 <x@example.com>"
    assert dict(m.headers)["Subject"] == "①髙橋です (re)"  # 符号化語の間の空白は字にしない（RFC 2047）
    assert m.body == "①の件、髙橋です。～\r\n" and m.notes == []
    # 標準の字だけのメールは、前と同じ読み方（〜 は U+301C のまま）
    std = (b"Subject: =?ISO-2022-JP?B?" + base64.b64encode("申請〜".encode("iso2022_jp")) + b"?=\r\n"
           b"Content-Type: text/plain; charset=iso-2022-jp\r\n\r\n" + "本文〜\r\n".encode("iso2022_jp"))
    m = mail.read_eml(std)
    assert dict(m.headers)["Subject"] == "申請〜" and m.body == "本文〜\r\n" and m.notes == []


def test_eml_bytes_no_reader_reads_are_told():
    raw = (b"Subject: x\r\nContent-Type: text/plain; charset=iso-2022-jp\r\n\r\n"
           b"\x1b$B0!.!0!\x1b(B\r\n")  # 14 区 1 点は、どちらの表にも字がない
    m = mail.read_eml(raw)
    assert "\ufffd" in m.body
    assert m.notes == ["本文に、宣言の文字コード（iso-2022-jp）では読めないバイトがあります。読めない所は「�」にしました"]
    # 解けない符号化語は、元の字のまま残す（字を作らない）
    raw = b"Subject: =?ISO-2022-JP?B?" + base64.b64encode(b"\x1b$B.!\x1b(B") + b"?=\r\n\r\nbody\r\n"
    assert dict(mail.read_eml(raw).headers)["Subject"].startswith("=?ISO-2022-JP?B?")


@pytest.mark.parametrize("enc, name", [("iso2022_jp", "様式.csv"), (MS, "①髙橋.csv")])
def test_eml_attachment_names(tmp_path, enc, name):
    """添付の名前：符号化語（name="=?…?="）と RFC 2231（filename*=iso-2022-jp''…）。Windows の字の名前では、標準の email は
    符号化語を解かず、RFC 2231 の名前でメール全体を読めなくなる：どちらも元の行から読み直す。"""
    ew = b"=?ISO-2022-JP?B?" + base64.b64encode(name.encode(enc)) + b"?="
    pct = urllib.parse.quote_from_bytes(name.encode(enc)).encode()
    raw = (b"Subject: x\r\nContent-Type: multipart/mixed; boundary=B\r\n\r\n--B\r\n"
           b"Content-Type: text/plain; charset=us-ascii\r\n\r\nbody\r\n--B\r\n"
           b"Content-Type: text/csv; name=\"" + ew + b"\"\r\nContent-Disposition: attachment; filename=\"" + ew + b"\"\r\n\r\n"
           b"a,b\r\n--B\r\nContent-Type: text/csv\r\nContent-Disposition: attachment; filename*=iso-2022-jp''" + pct +
           b"\r\n\r\nc,d\r\n--B--\r\n")
    m = mail.read_eml(raw)
    stem = name[:-4]
    assert m.attachments == [(name, b"a,b"), (f"{stem}(2).csv", b"c,d")] and m.body == "body" and m.notes == []
    p = _file(tmp_path, "a.eml", raw)
    text = "\n".join(x for _n, x in api.window(p)[2])
    assert name in text and f"{stem}(2).csv" in text
    assert "| R1（見出し） | a | b |" in "\n".join(x for _n, x in api.window(f"{p}!/{name}")[2])
