"""Windows そのものの読み書きと照らす（Windows でだけ動く：ほかの OS では飛ばす）。

とどくの iso2022_jp_ms（datafile/iso2022ms.py）と、切り替えの印のある文字コードへの書き込み（datafile/textenc.py）を、
Windows の kernel32 の MultiByteToWideChar・WideCharToMultiByte（コードページ 50220・50221・50222＝ISO-2022-JP、
50225＝ISO-2022-KR、65000＝UTF-7、52936＝HZ）と、その場で照らす：
- 表：ESC $ B と ESC $ ( D の 94 × 94 の区点の全部。違う所は、下の「分かっている違い」だけ。
- とどくが書いたバイトを Windows が読む：表の字の全部と、でたらめな字の並び。
- Windows が書いたバイトをとどくが読む：表の字の全部と、でたらめな字の並び（50220・50221・50222）。
- ISO-2022-KR・UTF-7・HZ のファイルに書き込んだ後のバイトを、Windows が同じ字に読む。

分かっている違い（Windows 10.0.26100 で確かめた。変わったら、このテストが落ちて分かる）：
- 89〜92 区（NEC 選定 IBM 拡張文字）：Windows は外字（U+E000〜）として読む。とどくは、メールの道具・iconv・ブラウザと同じに
  CP932 の字として読む。Windows は、その字を ESC $ ( D の 83・84 区で書く（とどくは、ESC $ ( D のあるファイルには同じに書く）。
- 13 区 64 点：Windows は 〞（U+301E）、とどくは CP932 と同じ 〝（U+301D）。
- 13 区のうち JIS X 0208 と同じ字の 9 点（≒ ≡ ∫ √ ⊥ ∠ ∵ ∩ ∪）：Windows は読まない。とどくは読む（書くのは JIS X 0208 の区点）。
- ESC $ ( D の 2 区 23 点・81 点（~・№）：Windows は読まない。とどくは JIS X 0212 のとおりに読む。
- 85〜94 区の外字：Windows は U+E000〜 として読む。とどくは読まない（89〜92 区を除く）。
- ESC ( J：Windows は 0x5C・0x7E を \\・~ と読む。とどくは ¥・‾（JIS X 0201）。
- ISO-2022-KR：Windows は、印 ESC $ ) C のすぐ後ろを、SO がなくても、SI が来るまで韓国語の字として読む（RFC 1557 の形の、
  印がファイルの頭にあって後ろが ASCII の字のファイルを、別の字に読む）。とどくは、印のすぐ後ろが ASCII の字になる
  書き込みでは SI を先に置く。元からその形のファイルの、書いていない所は変えない。
"""

from __future__ import annotations

import random
import sys

import pytest

from todoku import api
from todoku.datafile import iso2022ms, textenc  # noqa: F401  （textenc が codec を登録する）

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows のコードページの読み書きそのものと照らす")

MS = iso2022ms.NAME
NONE = "・"  # Windows は、字のない区点を ・（U+30FB）と読む（1 区 6 点は、本当の ・）
SAME_AS_0208 = "≒≡∫√⊥∠∵∩∪"
KANA = "".join(chr(c) for c in range(0xFF61, 0xFFA0))
ASCII = "aZ09 ,.\"'-/\t\n"


def _k32():
    import ctypes

    k = ctypes.windll.kernel32
    k.MultiByteToWideChar.argtypes = [ctypes.c_uint, ctypes.c_ulong, ctypes.c_char_p, ctypes.c_int, ctypes.c_void_p,
                                      ctypes.c_int]
    k.MultiByteToWideChar.restype = ctypes.c_int
    k.WideCharToMultiByte.argtypes = [ctypes.c_uint, ctypes.c_ulong, ctypes.c_wchar_p, ctypes.c_int, ctypes.c_void_p,
                                      ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p]
    k.WideCharToMultiByte.restype = ctypes.c_int
    return ctypes, k


def wdec(cp: int, data: bytes) -> str | None:
    """Windows がコードページ cp で読んだ字（読めなければ None）。"""
    ctypes, k = _k32()
    n = k.MultiByteToWideChar(cp, 0, data, len(data), None, 0)
    if n <= 0:
        return None
    buf = ctypes.create_unicode_buffer(n)
    k.MultiByteToWideChar(cp, 0, data, len(data), buf, n)
    return buf[:n]


def wenc(cp: int, s: str) -> bytes | None:
    """Windows がコードページ cp で書いたバイト（書けない字は「?」になる）。"""
    ctypes, k = _k32()
    wl = len(s.encode("utf-16-le")) // 2
    n = k.WideCharToMultiByte(cp, 0, s, wl, None, 0, None, None)
    if n <= 0:
        return None
    buf = ctypes.create_string_buffer(n)
    k.WideCharToMultiByte(cp, 0, s, wl, buf, n, None, None)
    return buf.raw[:n]


def _pua(ch) -> bool:
    return isinstance(ch, str) and len(ch) == 1 and 0xE000 <= ord(ch) <= 0xF8FF


def table_differences(win, table: dict, known) -> list:
    """94 × 94 の区点で、Windows の字（win(row, cell)）と、とどくの表の字が違う所のうち、known(row, cell, w, t) が
    「分かっている違い」と言わない所。"""
    out = []
    for row in range(1, 95):
        for cell in range(1, 95):
            w = win(row, cell)
            t = table[((row + 0x20) << 8) | (cell + 0x20)]
            same = w == t or (t == iso2022ms._BAD and w == NONE)
            if not same and not known(row, cell, w, t):
                out.append((row, cell, w, t))
    return out


def known_0208(row: int, cell: int, w, t: str) -> bool:
    if row >= 85:
        return _pua(w) and (t == iso2022ms._BAD or 89 <= row <= 92)
    if row == 13:
        return (cell == 64 and (w, t) == ("〞", "〝")) or (w == NONE and t in SAME_AS_0208)
    return False


def known_0212(row: int, cell: int, w, t: str) -> bool:
    if row >= 85:
        return _pua(w) and t == iso2022ms._BAD
    return w == NONE and (row, cell, t) in ((2, 23, "~"), (2, 81, "№"))


def _win_cell(mark: bytes):
    return lambda row, cell: wdec(50221, mark + bytes((row + 0x20, cell + 0x20)) + b"\x1b(B")


def test_two_byte_table_is_the_windows_table():
    assert table_differences(_win_cell(b"\x1b$B"), iso2022ms.tables()[0], known_0208) == []
    assert wdec(50221, b"\x1b$B!&\x1b(B") == NONE and iso2022ms.tables()[0][0x2126] == NONE  # 1 区 6 点は本当の ・


def test_supplementary_table_is_the_windows_table():
    assert table_differences(_win_cell(b"\x1b$(D"), iso2022ms.tables212()[0], known_0212) == []


@pytest.mark.parametrize("cp", [50220, 50222])
def test_the_three_code_pages_read_the_same(cp):
    for mark in (b"\x1b$B", b"\x1b$(D"):
        for row in range(1, 95):
            line = mark + b"".join(bytes((row + 0x20, c + 0x20)) for c in range(1, 95)) + b"\x1b(B"
            assert wdec(cp, line) == wdec(50221, line), (cp, mark, row)


def _written_low(ch: str) -> bool:
    """とどくが ESC $ B の 1〜84 区に書く字か。"""
    kk = iso2022ms.tables()[1].get(ch)
    return kk is not None and kk[0] - 0x20 <= 84


def test_windows_reads_every_character_todoku_writes():
    """表の字を 1 字ずつ書いて、Windows に読ませる。違うのは、分かっている違いだけ。"""
    enc, enc212 = iso2022ms.tables()[1], iso2022ms.tables212()[1]
    other = {ch: wdec(50221, iso2022ms.encode(ch)) for ch in enc if wdec(50221, iso2022ms.encode(ch)) != ch}
    assert {ch for ch in other if _written_low(ch)} == {"〝"} and other["〝"] == "〞"
    assert all(_pua(w) for ch, w in other.items() if ch != "〝")  # 89〜92 区に書いた字：Windows は外字として読む
    # ESC $ ( D のあるファイルの書き方：89〜92 区の字を、Windows と同じに ESC $ ( D で書く
    other = {}
    for ch in sorted(set(enc) | set(enc212)):
        got = wdec(50221, iso2022ms.encode(ch, True))
        if got != ch:
            other[ch] = got
    assert set(other) == {"〝", "昻", "￢", "￤"}, other  # 後ろの 3 字は、Windows の ESC $ ( D の表にない（Windows は「?」と書く）
    for ch in KANA:
        raw = iso2022ms.encode(ch)
        assert raw == b"\x1b(I" + bytes((ord(ch) - 0xFF61 + 0x21,)) + b"\x1b(B"
        assert [wdec(cp, raw) for cp in (50220, 50221, 50222)] == [ch] * 3
    assert wdec(50221, iso2022ms.encode("¥‾")) == "\\~"  # ESC ( J：Windows は ASCII として読む


def _pool(x0212: bool) -> str:
    enc, enc212 = iso2022ms.tables()[1], iso2022ms.tables212()[1]
    if x0212:
        chars = (set(enc) | set(enc212)) - set("〝昻￢￤")
    else:
        chars = {ch for ch in enc if _written_low(ch)} - {"〝"}
    return "".join(sorted(ch for ch in chars if ord(ch) >= 0x80))


@pytest.mark.parametrize("x0212", [False, True])
def test_windows_reads_random_text_todoku_writes(x0212):
    rnd = random.Random(20261113 + x0212)
    pool = _pool(x0212)
    for _ in range(400):
        s = "".join(rnd.choice(pool if rnd.random() < 0.5 else (KANA if rnd.random() < 0.4 else ASCII))
                    for _ in range(rnd.randrange(1, 40)))
        raw = iso2022ms.encode(s, x0212)
        assert raw.decode(MS) == s
        for cp in (50220, 50221, 50222):
            assert wdec(cp, raw) == s, (cp, s, raw)


def test_todoku_reads_every_character_windows_writes():
    enc, enc212 = iso2022ms.tables()[1], iso2022ms.tables212()[1]
    unwritten = []
    for ch in sorted(set(enc) | set(enc212) | set(KANA)):
        for cp in (50221, 50222):
            raw = wenc(cp, ch)
            if raw == b"?":
                unwritten.append(ch)
            else:
                assert raw.decode(MS) == ch, (cp, ch, raw)
    assert set(unwritten) == {"〝", "昻", "￢", "￤"}  # とどくの表の字のうち、Windows が書かない字
    assert all(wenc(cp, ch) == b"?" for cp in (50220, 50221, 50222) for ch in "～－∥￠￡")  # Windows は、JIS の字（〜 − ‖ ¢ £）の方だけを書く
    assert not set("〜−‖¢£¬髙﨑①㈱") & set(unwritten)
    assert wenc(50221, "髙") == b"\x1b$(Dt)\x1b(B" == iso2022ms.encode("髙", True)


@pytest.mark.parametrize("cp", [50220, 50221, 50222])
def test_todoku_reads_random_text_windows_writes(cp):
    """Windows が書いたバイトを、とどくが、書かせた字のとおりに読む。50220 は半角カナを全角にして書くので、Windows が
    読み直した字と比べる。50222 は SO で半角カナを書き、SI を書かずに ESC の印で終える。"""
    rnd = random.Random(cp)
    enc, enc212 = iso2022ms.tables()[1], iso2022ms.tables212()[1]
    pool = "".join(ch for ch in sorted(set(enc) | set(enc212)) if ord(ch) >= 0x80 and wenc(50221, ch) != b"?")
    for _ in range(400):
        s = "".join(rnd.choice(pool if rnd.random() < 0.5 else (KANA if rnd.random() < 0.4 else ASCII))
                    for _ in range(rnd.randrange(1, 40)))
        raw = wenc(cp, s)
        if cp == 50220:
            assert raw.decode(MS) == wdec(cp, raw), (s, raw)
        else:
            assert raw.decode(MS) == s, (s, raw)


def _fill(tmp_path, name: str, blob: bytes, values: dict, encoding=None) -> bytes:
    p = tmp_path / name
    p.write_bytes(blob)
    out = tmp_path / ("out_" + name)
    r = api.fill(str(p), values, str(out), encoding=encoding)
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"] and not r["unexpected"], r
    return out.read_bytes()


def test_windows_reads_files_after_todoku_wrote_into_them(tmp_path):
    """切り替えの印のある文字コードのファイルに書き込んだ後のバイトを、Windows が、Python と同じ字に読む。"""
    late = b"a,b\n1,\x1b$)C\x0eH+\x0f\n"  # ISO-2022-KR の印が、ファイルの頭でなく、初めの韓国語の字の前にある（Windows もこう書く）
    assert wenc(50225, "a,b\n1,홍\n") == late
    for i, (values, want) in enumerate([({"R1.C1": "성"}, "성,b\n1,홍\n"), ({"R1.C1": "성", "R2.C2": "길"}, "성,b\n1,길\n"),
                                        ({"R3.C1": "끝"}, "a,b\n1,홍\n끝,\n"), ({"R1.C1": "z", "R2.C2": "성"}, "z,b\n1,성\n")]):
        new = _fill(tmp_path, f"k{i}.csv", late, values)
        assert new.decode("iso2022_kr") == want and wdec(50225, new) == want, new
    # 印のすぐ後ろが ASCII の字になる書き込み：SI を先に置くので、Windows も同じ字に読む（SI がなければ、別の字に読む）
    top = "성,b\n홍,c\n".encode("iso2022_kr")
    assert wenc(50225, "성,b\n홍,c\n") == top and wdec(50225, top) == "성,b\n홍,c\n"
    assert wdec(50225, b"\x1b$)Cabc,b\n\x0eH+\x0f,c\n") != "abc,b\n홍,c\n"
    for i, (values, want) in enumerate([({"R1.C1": "abc"}, "abc,b\n홍,c\n"), ({"R1.C1": ""}, ",b\n홍,c\n"), ({"R1.C1": "a성"}, "a성,b\n홍,c\n"),
                                        ({"R1.C1": "abc", "R2.C1": "x"}, "abc,b\nx,c\n"), ({"R1.C2": "x길"}, "성,x길\n홍,c\n")]):
        new = _fill(tmp_path, f"t{i}.csv", top, values)
        assert new.decode("iso2022_kr") == want and wdec(50225, new) == want, new
    new = _fill(tmp_path, "k.json", '{"name": "김", "b": "홍"}\n'.encode("iso2022_kr"), {"/name": "Kim"})
    assert new.decode("iso2022_kr") == '{"name": "Kim", "b": "홍"}\n' == wdec(50225, new), new
    new = _fill(tmp_path, "k2.json", '{"name": "김"}\n'.encode("iso2022_kr"), {"/name": "Kim"})
    assert new.decode("iso2022_kr") == '{"name": "Kim"}\n' == wdec(50225, new), new
    u = "氏名,あいう\n住所,東京\n".encode("utf-7")
    for i, (values, want) in enumerate([({"R1.C2": "あxう"}, "氏名,あxう\n住所,東京\n"), ({"R2.C2": "a+b-c"}, "氏名,あいう\n住所,a+b-c\n"),
                                        ({"R1.C2": "-", "R3.C1": "終"}, "氏名,-\n住所,東京\n終,\n")]):
        new = _fill(tmp_path, f"u{i}.csv", u, values, "utf-7")
        assert new.decode("utf-7") == want and wdec(65000, new) == want, new
    z = "己所,不欲\n勿施,a~b\n".encode("hz")
    new = _fill(tmp_path, "z.csv", z, {"R1.C2": "一a二", "R2.C1": "~"}, "hz")
    want = "己所,一a二\n~,a~b\n"
    assert new.decode("hz") == want and wdec(52936, new) == want, new
    j = "氏名,住所\nｶﾅ,①髙\n".encode(MS)
    new = _fill(tmp_path, "j.csv", j, {"R1.C2": "〜㈱", "R3.C1": "ｱ"})
    want = "氏名,〜㈱\nｶﾅ,①髙\nｱ,\n"
    assert new.decode(MS) == want
    assert wdec(50221, new) == want.replace("髙", wdec(50221, iso2022ms.encode("髙")))  # 89〜92 区の字だけ、Windows は外字として読む
    j = wenc(50221, "氏名,住所\nｶﾅ,①髙\n")  # Windows が書いたファイル（髙 は ESC $ ( D）に、同じ書き方で足す
    new = _fill(tmp_path, "w.csv", j, {"R1.C2": "〜﨑", "R3.C1": "ｱ"})
    want = "氏名,〜﨑\nｶﾅ,①髙\nｱ,\n"
    assert new.decode(MS) == want and wdec(50221, new) == want, new
