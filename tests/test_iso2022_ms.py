"""ISO-2022-JP の Windows の拡張（CP50220・CP50221・CP50222）：NEC 特殊文字（13 区：①・㈱）と NEC 選定 IBM 拡張文字
（89〜92 区：髙・﨑）、Windows が ESC $ ( D で書く IBM 拡張文字のあるファイル・メールを読み、書く。

- codec（todoku.datafile.iso2022ms）：JIS X 0208 にある区点は規格の表（Python の iso2022_jp と同じ字：〜 は U+301C）、
  ない区点（13 区・89〜92 区）は CP932 の表、ESC $ ( D は JIS X 0212 と Windows の 83・84 区。表の字のハッシュと、決まった
  バイトで確かめる。Windows（10.0.26100）の 50221・50222 が書いたバイト（WIN_WROTE）は、Windows の上で採ったもの。
  Windows そのものの表・読み書きとの照らしは、tests/test_windows_codepages.py（Windows でだけ動く）。
- 見分け方：標準の ISO-2022-JP の仲間のどれでも読めないファイルだけを、Windows の決まりで読む。
- 書き込み：書いた所のバイトだけを変える。表にない字は断る。標準の ISO-2022-JP のファイルに Windows の字を
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
WIN = "番号,会社,氏名,備考\n①,㈱髙橋,髙﨑 花子,ｶﾅ ¥100\n②,㈲山田,山田 太郎,〜‖−‾\n"
# iconv -t CP50221 が書いたバイト（¥ の後ろの数字も ESC ( J のまま）と、CP50222 のバイト（半角カナは SO・SI）。
# 1 区 33 点・34 点・61 点（!A !B !]）は、iconv とブラウザは ～∥－（CP932 の字）と読み、Windows の 50221 と Python の
# iso2022_jp は 〜‖−（JIS X 0208 の規格の字）と読む：とどくは、規格の ISO-2022-JP と同じ字に読む
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
def test_table_is_the_standard_table_with_the_cp932_rows_added():
    dec, enc = iso2022ms.tables()
    assert len(dec) == 94 * 94
    have = [c for c in sorted(dec) if dec[c] != iso2022ms._BAD]
    assert (len(have), len(enc)) == (7336, 7327)  # 同じ字が 2 つの区にあるものが 9
    # 区点の順に並べた字（字のない区点は U+FFFF）
    assert hashlib.sha256("".join(dec[c] for c in sorted(dec)).encode("utf-8")).hexdigest() == \
        "9e3ddbc40a93cf29740ddb8d7bc88db396f9bf1c460c275129f03f38e065e8b9"
    for ch, code in enc.items():
        assert dec[(code[0] << 8) | code[1]] == ch
    # JIS X 0208 にある区点は、Python の iso2022_jp と同じ字。ない区点は、CP932 の同じ区点の字
    std = cp = 0
    for row in range(1, 95):
        for cell in range(1, 95):
            kk = bytes((row + 0x20, cell + 0x20))
            mine = dec[(kk[0] << 8) | kk[1]]
            try:
                assert mine == (b"\x1b$B" + kk + b"\x1b(B").decode("iso2022_jp")
                std += 1
            except UnicodeDecodeError:
                try:
                    assert mine == iso2022ms._sjis(row, cell).decode("cp932")
                    cp += 1
                except UnicodeDecodeError:
                    assert mine == iso2022ms._BAD
    assert (std, cp) == (6879, 457)
    assert [r for r in range(1, 95) if any(dec[((r + 0x20) << 8) | (c + 0x20)] != iso2022ms._BAD
                                           and _not_std(r, c) for c in range(1, 95))] == [13, 89, 90, 91, 92]


def _not_std(row: int, cell: int) -> bool:
    try:
        (b"\x1b$B" + bytes((row + 0x20, cell + 0x20)) + b"\x1b(B").decode("iso2022_jp")
    except UnicodeDecodeError:
        return True
    return False


def test_supplementary_table_is_jis_x_0212_with_the_windows_rows():
    dec, enc = iso2022ms.tables212()
    have = [c for c in sorted(dec) if dec[c] != iso2022ms._BAD]
    assert (len(dec), len(have), len(enc)) == (94 * 94, 6159, 6158)  # 2 区 23 点の ~ は、ASCII の字として書く
    assert hashlib.sha256("".join(dec[c] for c in sorted(dec)).encode("utf-8")).hexdigest() == \
        "c0b60b1c8fe7c09eedf45745a77762f25372c6e6f7ff5095298406e3cc61dbb7"
    ibm = 0
    for row in range(1, 95):
        for cell in range(1, 95):
            kk = bytes((row + 0x20, cell + 0x20))
            mine = dec[(kk[0] << 8) | kk[1]]
            try:
                assert mine == (b"\x1b$(D" + kk + b"\x1b(B").decode("iso2022_jp_1")
            except UnicodeDecodeError:
                assert row in (83, 84) or mine == iso2022ms._BAD
                ibm += mine != iso2022ms._BAD
    assert ibm == 92
    # Windows が 83・84 区に置く字は、どれも CP932 の IBM 拡張文字（115〜119 区）にある字
    for row in (83, 84):
        for cell in range(1, 95):
            ch = dec[((row + 0x20) << 8) | (cell + 0x20)]
            if ch != iso2022ms._BAD:
                assert 0xFA <= ch.encode("cp932")[0] <= 0xFC or 0xED <= ch.encode("cp932")[0] <= 0xEE, (row, cell)


@pytest.mark.parametrize("text, blob", [
    ("①", b"\x1b$B-!\x1b(B"), ("髙", b"\x1b$B|b\x1b(B"), ("ⅰ", b"\x1b$B|q\x1b(B"), ("Ⅰ", b"\x1b$B-5\x1b(B"),
    ("∵", b"\x1b$B\"h\x1b(B"), ("¬", b"\x1b$B\"L\x1b(B"),  # 2 つの区にある字は、区の番号の小さい方
    ("￢", b"\x1b$B|{\x1b(B"),  # 規格の表にない ￢ は、92 区 91 点（2 区 44 点は、規格の字 ¬）
    ("¥", b"\x1b(J\\\x1b(B"), ("‾", b"\x1b(J~\x1b(B"), ("ｶ", b"\x1b(I6\x1b(B"), ("〜‖−", b"\x1b$B!A!B!]\x1b(B"),
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
    (b"\x1b$(Dt)sI\x1b(B", "髙\ufa11"), (b"\x1b$(D0!\x1b(B", "丂"),  # ESC $ ( D：Windows の 84 区 9 点・83 区 41 点、JIS X 0212
    (b"\x1b$(D\"7\x1b(B", "~"), (b"\x1b$(Ds!\x1b$B|q\x1b(B", "ⅰⅰ"),
    (b"\x1b(J\x0e6\x1b(Ba", "ｶa"), (b"\x0e6\x1b$B0!\x1b(B", "ｶ亜"),  # ESC の印は SO を終える（Windows は SI を書かない）
    (b"\x0e6\x0f\x0f6", "ｶ6"),  # SO の間でないときの SI は、何も変えない
])
def test_reads_what_windows_writes(blob, text):
    assert blob.decode(MS) == text
    if text != "丂":  # JIS X 0212 だけにある字は、ESC $ ( D の書き方でだけ書ける
        assert text.encode(MS).decode(MS) == text
    assert iso2022ms.encode(text, True).decode(MS) == text


# Windows（10.0.26100）の WideCharToMultiByte が書いたバイト：(字, 50221 のバイト, 50222 のバイト)。Windows の上で採った
WIN_WROTE = [
    ("~侄あｶ。X。", "7e1b242844307e1b244224221b2849361b244221231b2842581b244221231b2842",
     "7e1b242844307e1b244224221b284a0e361b244221231b2842581b244221231b2842"),
    ("あいﾅ-髙£aアb侄", "1b2442242224241b2849451b28422d1b24284474291b244221721b2842611b244225221b2842621b242844307e1b2842",
     "1b2442242224241b284a0e451b28422d1b24284474291b244221721b2842611b244225221b2842621b242844307e1b2842"),
    ("ﾅ‖侄X丂德㈱ﾅ\"鷗", "1b2849451b244221421b242844307e1b2842581b2428443021734a1b24422d6a1b2849451b2842221b2428446c3f1b2842",
     "1b284a0e451b244221421b242844307e1b2842581b2428443021734a1b24422d6a0e451b2842221b2428446c3f1b2842"),
    ("\ufa11ｶ〜,鷗丂\ufa11№", "1b24284473491b2849361b244221411b28422c1b2428446c3f302173491b24422d621b2842",
     "1b24284473491b284a0e361b244221411b28422c1b2428446c3f302173491b24422d621b2842"),
    ("㍻\n/いｶZ", None, "1b24422d5f1b28420a2f1b244224241b284a0e361b28425a"),
    ("\nｱ侄\\德", None, "0a1b284a0e311b242844307e1b28425c1b242844734a1b2842"),
    ("㈱‖¢\n々-イ侄\n、ｰ", None, "1b24422d6a214221711b28420a1b244221391b28422d1b244225241b242844307e1b28420a1b244221221b284a0e301b2842"),
    ("ｶ\nﾅ", None, "1b284a0e361b28420a0e451b2842"), ("あｶあ", None, "1b244224221b284a0e361b244224221b2842"),
]


@pytest.mark.parametrize("text, w21, w22", WIN_WROTE)
def test_bytes_windows_itself_wrote(text, w21, w22):
    """Windows の 50221 のバイトは、ESC $ ( D の書き方（x0212）で書いたバイトと同じ。50222 のバイト（SO で半角カナ、SI なし）も
    同じ字に読む。"""
    if w21 is not None:
        assert bytes.fromhex(w21).decode(MS) == text and iso2022ms.encode(text, True) == bytes.fromhex(w21)
    assert bytes.fromhex(w22).decode(MS) == text
    d = codecs.getincrementaldecoder(MS)()
    assert "".join(d.decode(bytes((b,))) for b in bytes.fromhex(w22)) + d.decode(b"", True) == text


def test_ibm_extension_characters_follow_the_file():
    """89〜92 区の字（髙・﨑）と、2 バイトの表にない字（JIS X 0212 の 侄）：ESC $ ( D の書き方では ESC $ ( D で、
    そうでなければ 89〜92 区で書く（侄 は書けない）。ESC $ ( D に点のない 昻 は、どちらでも 89〜92 区。"""
    assert iso2022ms.encode("髙ⅰ①侄昻", True) == b"\x1b$(Dt)s!\x1b$B-!\x1b$(D0~\x1b$Bz6\x1b(B"
    assert iso2022ms.encode("髙ⅰ①昻") == "髙ⅰ①昻".encode(MS) == b"\x1b$B|b|q-!z6\x1b(B"
    with pytest.raises(UnicodeEncodeError):
        iso2022ms.encode("侄")
    for x in (False, True):
        table = set(iso2022ms.tables()[1]) | (set(iso2022ms.tables212()[1]) if x else set())
        text = "".join(sorted(table))
        assert iso2022ms.encode(text, x).decode(MS) == text


@pytest.mark.parametrize("blob, why", [
    (b"\x1b$B.!\x1b(B", "字のない区点"), (b"\x1b$(O0!\x1b(B", "印として読めない"), (b"\x1b(I`\x1b(B", "半角カナにない"),
    (b"\x1b$(D.!\x1b(B", "字のない区点"), (b"\x1b$(Dt#\x1b(B", "字のない区点"), (b"\x1b$(", "印として読めない"),
    (b"\x1b$B0\x1b(B", "途中で終わっている"), (b"\x1b[0m", "印として読めない"), (b"a\x80", "印として読めない"),
    (b"\x1b$B0!0", "途中で終わっている"), (b"\x1b", "印として読めない"),
])
def test_bytes_outside_the_rule_are_refused(blob, why):
    with pytest.raises(UnicodeDecodeError, match=why):
        blob.decode(MS)
    assert "\ufffd" in blob.decode(MS, "replace")


@pytest.mark.parametrize("text", ["～", "∥", "－", "￠", "￡", "侄", "한", "\x1b", "\x0e", "\x0f", "a😀"])
def test_characters_outside_the_table_are_refused(text):
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
    """標準の読み手で読めるファイルは、標準の文字コードとして読む。Windows の決まりで読むのは、頼まれたときだけで、
    読んだ字は同じ（〜 は、どちらでも U+301C）。"""
    blob = "氏名,〜\n山田,a\n".encode("iso2022_jp")
    assert (decode(blob).encoding, decode(blob).text) == ("iso2022_jp", "氏名,〜\n山田,a\n")
    assert (decode(blob, "cp50221").encoding, decode(blob, "cp50221").text) == (MS, "氏名,〜\n山田,a\n")
    # 規格の読み手が読むバイトは、同じ字に読む
    for std, text in (("iso2022_jp", "氏名：山田〜‖−¢£¬ ¥‾"), ("iso2022_jp_1", "氏名：山田〜‖− ¥‾ 侄丂"),
                      ("iso2022_jp_ext", "氏名：山田〜‖− ¥‾ 侄 ｶﾅ")):
        assert text.encode(std).decode(MS) == text
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
    assert "| R2 | ① | ㈱髙橋 | 髙﨑 花子 | ｶﾅ ¥100 |" in text and "| R3 | ② | ㈲山田 | 山田 太郎 | 〜‖−‾ |" in text
    assert f"文字コード {NAME}・BOM なし" in text and "決められません" not in text
    g = api.grade(p)
    assert (g["grade"], g["write"]["grade"]) == ("S", "A")
    f = {x["id"]: x.get("value") for x in api.fields(p)}
    assert f["R2.C2"] == "㈱髙橋" and f["R3.C4"] == "〜‖−‾"
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


def test_fill_writes_ibm_extension_characters_the_way_the_file_does(tmp_path):
    """Windows の 50221 が書いたファイル（髙 は ESC $ ( D の 84 区 9 点）：書く 﨑・髙 も ESC $ ( D で書く（Windows は 89〜92 区を
    外字として読む）。ESC $ ( D のないファイル（メールの道具・iconv の書き方）には、89〜92 区で書く。"""
    win = b"\x1b$B;aL>\x1b(B,\x1b$(Dt)\x1b$B66\x1b(B\n\x1b$B-!\x1b(B,a\n"  # 氏名,髙橋 / ①,a
    assert decode(win).encoding == MS and decode(win).text == "氏名,髙橋\n①,a\n"
    new = _fill(tmp_path, "x.csv", win, {"R2.C2": "\ufa11 ｶ髙"})
    assert new == win.replace(b",a\n", b",\x1b$(DsI\x1b(B \x1b(I6\x1b$(Dt)\x1b(B\n")
    assert new.decode(MS) == "氏名,髙橋\n①,\ufa11 ｶ髙\n"
    new = _fill(tmp_path, "y.csv", ICONV_21, {"R3.C4": "\ufa11"})
    assert new == ICONV_21.replace(b"\x1b$B!A!B!]\x1b(J~\x1b(B\n", b"\x1b$Byu\x1b(B\n")
    # Windows の 50222 が書いた半角カナ（ESC ( J・SO、SI なし）のある行：ほかの欄に書いても、その欄のバイトは同じ
    so = b"\x1b$B-!\x1b(B,\x1b(J\x0e6E\x1b(B,a\n"
    assert decode(so).text == "①,ｶﾅ,a\n"
    assert _fill(tmp_path, "z.csv", so, {"R1.C3": "髙"}) == so.replace(b",a\n", b",\x1b$B|b\x1b(B\n")


@pytest.mark.parametrize("value", ["한", "～", "😀"])
def test_characters_outside_the_windows_table_are_refused(tmp_path, value):
    p = _file(tmp_path, "w.csv", ICONV_21)
    with pytest.raises(api.ApiError, match="CP50221：NEC 特殊文字・IBM 拡張文字つき））では書けない字") as e:
        api.fill(p, {"R2.C1": value}, str(tmp_path / "o.csv"))
    # ～（CP932 の字）は、規格の表が同じ所に置く 〜 なら書ける：そのことを伝える（字は置き換えない）
    assert ("「〜」なら、このファイルに書けます" in str(e.value)) == (value == "～")
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
    wave = _file(tmp_path, "b.csv", "氏名,〜\n山田,a\n".encode("iso2022_jp"))  # 〜 のあるファイルも、読む字は同じ
    with pytest.raises(api.ApiError, match="どちらで読んでも同じ"):
        api.fill(wave, {"R2.C2": "髙"}, str(tmp_path / "o.csv"))
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
        d.rewrite(blob, [(0, 1, "～")])
    assert d.rewrite(blob, [(0, 1, "〜")]).decode(MS) == "〜a髙,ｶ\n"


# ---------------------------------------------------------------------- メール（.eml）
def _b(s: str) -> bytes:
    return b"=?ISO-2022-JP?B?" + base64.b64encode(s.encode(MS)) + b"?="


def test_eml_with_windows_characters_under_the_iso2022jp_name():
    """Windows のメールソフトは、①・髙 を、ISO-2022-JP の名前のまま送る。標準の email は、ヘッダーの符号化語を解かずに残し、
    本文の字を「�」にする：元のバイトから読み直す。"""
    raw = (b"From: " + _b("髙橋 ㈱") + b" <a@example.com>\r\nTo: =?iso-2022-jp?Q?=1B$B;3ED=1B(B?= <x@example.com>\r\n"
           b"Subject: " + _b("①髙橋") + b"\r\n " + _b("です") + b" (re)\r\n"
           b"Content-Type: text/plain; charset=iso-2022-jp\r\n\r\n" + "①の件、髙橋です。〜\r\n".encode(MS))
    m = mail.read_eml(raw)
    assert dict(m.headers)["From"] == "髙橋 ㈱ <a@example.com>" and dict(m.headers)["To"] == "山田 <x@example.com>"
    assert dict(m.headers)["Subject"] == "①髙橋です (re)"  # 符号化語の間の空白は字にしない（RFC 2047）
    assert m.body == "①の件、髙橋です。〜\r\n" and m.notes == []
    # Windows の 50221 が書いた本文（髙 は ESC $ ( D、半角カナは ESC ( I）
    win = (b"Subject: x\r\nContent-Type: text/plain; charset=iso-2022-jp\r\n\r\n"
           b"\x1b$B-!\x1b$(Dt)\x1b$B66\x1b(I6E\x1b(B\r\n")
    assert mail.read_eml(win).body == "①髙橋ｶﾅ\r\n" and mail.read_eml(win).notes == []
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
