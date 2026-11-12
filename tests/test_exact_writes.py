"""書いた字が、読み直すと違う字になる書き込みを断る（textenc.exact）。

Python のエンコーダは、その文字コードの表にない字の一部を、近い字のバイトにする：Shift_JIS（CP932）の 〜 U+301C は ～ U+FF5E と
同じバイト、EUC-JP の ¥ は \\ のバイト。黙って書くと、ファイルの字が、頼まれた字と違う字になる。とどくは、書かずに、
どの字になるかを伝える（字は置き換えない：エージェントが決める）。その文字コードで書けない字が、規格の表と Windows の表で同じ
所に置かれた字の片方なら、もう一方の字を伝える（対は、2 つの表を全部のバイトで比べて決める：twins）。

- CSV・.txt・JSON の書き込み（Decoded.encode）、Excel 5.0/95・Word 6.0/95・dBASE の書き込み、XFA の XML（文字参照で書く）。
- ISO-2022-KR の印より前に書いた後の、何も変えなくなった印を取る（Shifts.idle）：.txt。CSV は test_stateful_encodings.py。
- ISO-2022-KR の印（ESC $ ) C）のすぐ後ろが ASCII の字になる書き込みは、SI を先に置く（Shifts.lead：Windows の 50225 は、
  印のすぐ後ろを、SI が来るまで韓国語の字として読む）。書く字の途中の 2 つ目の印は書かない（Shifts._lean）。韓国語の字が
  なくなって印だけが残ったファイルも、同じ文字コードのまま（kept_encoding）。
- メール：文字コードの宣言のない（か、us-ascii の）ISO-2022-JP の本文を、印のとおりに読む。制御文字が残る本文は、伝える。
"""
from __future__ import annotations

import codecs
import struct
from types import SimpleNamespace

import pytest

from todoku import api
from todoku.datafile.textenc import EncodingError, decode, exact, twin_hint, twins
from todoku.legacy import mail

# Python の表で、書いて読み直すと違う字になる字（基本多言語面を全部しらべた数：テストの中でも数える）
ONE_WAY = {"cp932": "¢￠ £￡ ¬￢ ‖∥ −－ 〜～", "shift_jis": "¥\\ ‾~", "euc_jp": "¥\\ ‾~", "iso2022_jp_3": "鬝鬜",
           "cp950": "¢￠ £￡ ¥￥ •‧ ‾¯ ∼～ ☉⊙ ♁⊕ ､﹑"}
SAME = ["utf-8", "utf-16", "utf-32", "cp949", "johab", "iso2022_kr", "gbk", "gb2312", "gb18030", "hz", "big5", "big5hkscs",
        "iso2022_jp", "iso2022_jp_1", "iso2022_jp_2", "iso2022_jp_2004", "iso2022_jp_ext", "iso2022_jp_ms", "euc_jis_2004",
        "shift_jis_2004", "cp1252", "latin-1", "koi8_r", "cp437", "cp1251", "cp874", "utf-7"]


def _scan(enc: str) -> tuple[list[tuple[str, str]], list[str]]:
    one, dead = [], []
    for o in range(0x20, 0x10000):
        if 0xD800 <= o < 0xE000:
            continue
        ch = chr(o)
        try:
            b = ch.encode(enc)
        except UnicodeEncodeError:
            continue
        try:
            back = b.decode(enc)
        except UnicodeDecodeError:
            dead.append(ch)
            continue
        if back != ch:
            one.append((ch, back))
    return one, dead


@pytest.mark.parametrize("enc", sorted(ONE_WAY))
def test_one_way_characters_are_refused_with_what_they_become(enc):
    one, dead = _scan(enc)
    assert " ".join(a + b for a, b in one) == ONE_WAY[enc] and not dead
    for ch, back in one:
        with pytest.raises(EncodingError) as e:
            exact(f"a{ch}b", enc)
        said = str(e.value)
        assert f"「{ch}」（U+{ord(ch):04X}）" in said and f"「{back}」（U+{ord(back):04X}）と同じバイトになります" in said
        if enc == "iso2022_jp_3":  # 読み直した字 鬜 U+9B1C そのものは、Python の表では書けない：勧めない
            assert "でよければ" not in said
            continue
        assert f"「{back}」でよければ、その字で書いてください" in said
        assert exact(f"a{back}b", enc) == f"a{back}b".encode(enc)  # 伝えた字は、そのまま書ける


@pytest.mark.parametrize("enc", SAME)
def test_encodings_without_one_way_characters_write_as_before(enc):
    if enc not in ("utf-8", "utf-16", "utf-32", "utf-7", "gb18030"):
        assert _scan(enc) == ([], [])
    for text in ("abc 123", "氏名：山田 太郎", "홍길동", "张三", "Ünïcödé €", "ｶﾅ ¥100"):
        try:
            want = text.encode(enc)
        except UnicodeEncodeError:
            with pytest.raises(EncodingError, match="では書けない字です"):
                exact(text, enc)
            continue
        assert exact(text, enc) == want


def test_bytes_that_cannot_be_read_back_are_refused():
    """EUC-KR のエンコーダは U+3164（ハングルの詰め物）をバイトにするが、そのバイトは EUC-KR として読めない。"""
    assert _scan("euc_kr") == ([], ["ㅤ"])
    with pytest.raises(EncodingError, match="バイトにはなりますが、そのバイトを読み直せません"):
        exact("aㅤ", "euc_kr")
    assert exact("홍길동", "euc_kr") == "홍길동".encode("euc_kr")


def test_twins_come_from_comparing_the_two_tables():
    tw = twins()
    pairs = {frozenset((a, t[0])) for a, ts in tw.items() for t in ts}
    assert pairs == {frozenset(x) for x in (
        "¢￠", "£￡", "¬￢", "‖∥", "−－", "〜～",  # JIS X 0208 の規格の表と CP932
        "¢￠", "£￡", "¥￥", "¯‾", "•‧", "∼～", "☉⊙", "♁⊕", "﹑､", "∕／", "﹨＼",  # Big5 と CP950
        "·・", "—―")}  # GB 2312 と GBK
    for a, ts in tw.items():
        for t, own, other in ts:
            assert a != t and own != other


@pytest.mark.parametrize("enc, bad, good", [
    ("euc_jp", "～", "〜"), ("iso2022_jp", "～", "〜"), ("iso2022_jp_ms", "～", "〜"), ("iso2022_jp", "－", "−"),
    ("iso2022_jp_ms", "∥", "‖"), ("gb2312", "·", "・"), ("big5", "‧", "•"),
])
def test_unencodable_twin_is_told_with_the_character_that_can_be_written(enc, bad, good):
    with pytest.raises(EncodingError) as e:
        exact(f"1{bad}2", enc)
    said = str(e.value)
    assert f"「{bad}」は、このファイルの文字コード" in said and "では書けない字です" in said
    assert f"同じ所に「{good}」（U+{ord(good):04X}）を置きます。「{good}」なら、このファイルに書けます" in said
    assert exact(f"1{good}2", enc).decode(enc) == f"1{good}2"
    assert twin_hint("한", enc) == "" and twin_hint(good, enc) == ""  # 対のない字・書ける字には、何も添えない


# ---------------------------------------------------------------------- 道具から
def _file(tmp_path, name, blob):
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


@pytest.mark.parametrize("name, blob, fid", [
    ("a.csv", "氏名,期間\n山田,\n".encode("cp932"), "R2.C2"),
    ("a.txt", "申請書\n\n期間：\n".encode("cp932"), "P3.c3"),
])
def test_fill_refuses_the_one_way_character_and_writes_the_other(tmp_path, name, blob, fid):
    p = _file(tmp_path, name, blob)
    out = tmp_path / ("o_" + name)
    with pytest.raises(api.ApiError, match="「〜」（U\\+301C）は、このファイルの文字コード（Shift_JIS（CP932））では、「～」（U\\+FF5E）と同じ"):
        api.fill(p, {fid: "4月〜6月"}, str(out))
    assert not out.exists()
    r = api.fill(p, {fid: "4月～6月"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"]
    assert out.read_bytes().decode("cp932") == blob.decode("cp932").rstrip("\n") + "4月～6月\n"


def test_fill_dbf_refuses_the_one_way_character(tmp_path):
    fields = [(b"NAME", b"C", 12, 0)]
    h = bytearray(32)
    h[0] = 0x03
    struct.pack_into("<IHH", h, 4, 1, 32 + 32 + 1, 13)
    h[29] = 0x13  # 言語の印：Shift_JIS
    blob = bytes(h) + b"NAME".ljust(11, b"\0") + b"C" + bytes(4) + bytes([12, 0]) + bytes(14) + b"\x0d" + b" " + b" " * 12 + b"\x1a"
    p = _file(tmp_path, "a.dbf", blob)
    out = tmp_path / "o.dbf"
    with pytest.raises(api.ApiError, match="「〜」（U\\+301C）は.*「～」（U\\+FF5E）と同じバイトになります"):
        api.fill(p, {"A2": "1〜2"}, str(out))
    assert not out.exists()
    r = api.fill(p, {"A2": "1～2"}, str(out))
    assert r["proof_ok"] and r["verified"] and "1～2".encode("cp932") in out.read_bytes()


def test_xfa_xml_writes_one_way_characters_as_character_references():
    from todoku.pdf.xfafill import DataEdit

    def enc(name, s):
        return DataEdit._enc(SimpleNamespace(x=SimpleNamespace(xdp=SimpleNamespace(encoding=name))), s)

    assert enc("UTF-8", "1〜2") == "1〜2".encode("utf-8")
    assert enc("cp932", "1～2 髙") == "1～2 髙".encode("cp932")
    assert enc("cp932", "1〜2 한") == b"1&#12316;2 &#54620;"  # 〜 は、CP932 では ～ のバイトになるので、文字参照
    assert enc("euc-jp", "¥100 〜") == b"&#165;100 " + "〜".encode("euc_jp")
    assert enc("utf-16", "あ〜") == "あ〜".encode("utf-16")


# ---------------------------------------------------------------------- ISO-2022-KR の印
def test_txt_before_a_late_iso2022_kr_header_keeps_one_header(tmp_path):
    blob = b"name:\n\n" + "홍길동\n".encode("iso2022_kr")  # 印は、3 行目の頭
    assert blob == b"name:\n\n\x1b$)C\x0eH+1f5?\x0f\n"
    assert decode(blob).encoding == "iso2022_kr"
    p = _file(tmp_path, "k.txt", blob)
    out = tmp_path / "o.txt"
    r = api.fill(p, {"P1.c5": "김"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"], r
    new = out.read_bytes()
    assert new == b"name:\x1b$)C\x0e1h\x0f\n\n\x0eH+1f5?\x0f\n"  # 印は 1 つ（書いた字の前）。後ろの字のバイトは同じ
    assert new.count(b"\x1b$)C") == 1 and decode(new).text == "name:김\n\n홍길동\n"


KR = b"\x1b$)C"


def _kr(tmp_path, n, blob, values):
    p = _file(tmp_path, f"k{n}.csv", blob)
    out = tmp_path / f"o{n}.csv"
    r = api.fill(p, values, str(out))
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"] and not r["unexpected"], r
    return out.read_bytes()


def test_ascii_right_after_the_iso2022_kr_header_gets_a_shift_in(tmp_path):
    """印のすぐ後ろが ASCII の字になる書き込みは、SI を先に置く。ほかの読み手（Python・iconv・perl）は、SI があっても
    なくても同じ字に読む。印のすぐ後ろが SO（韓国語の字）なら、足さない。"""
    top = "성,b\n".encode("iso2022_kr")
    assert top == KR + b"\x0e<:\x0f,b\n"
    late = b"a,b\n1," + KR + b"\x0eH+\x0f\n"
    rfc = KR + b"abc,\x0eH+\x0f\n"  # RFC 1557 の形：印はファイルの頭で、その後ろが ASCII の字（Windows は、もとから別の字に読む）
    for n, (blob, values, want, text) in enumerate([
            (top, {"R1.C1": "abc"}, KR + b"\x0fabc,b\n", "abc,b\n"),
            (top, {"R1.C1": ""}, KR + b"\x0f,b\n", ",b\n"),
            (top, {"R1.C1": "길"}, KR + b"\x0e1f\x0f,b\n", "길,b\n"),
            (late, {"R2.C2": "x"}, b"a,b\n1," + KR + b"\x0fx\n", "a,b\n1,x\n"),
            (rfc, {"R1.C1": "y"}, KR + b"\x0fy,\x0eH+\x0f\n", "y,홍\n"),
            (rfc, {"R1.C2": "x"}, KR + b"abc,x\n", "abc,x\n"),  # 印のすぐ後ろには書いていない：元のバイトのまま
    ]):
        new = _kr(tmp_path, n, blob, values)
        assert new == want and new.decode("iso2022_kr") == text, (n, new)


def test_a_second_iso2022_kr_header_is_not_written_inside_the_value(tmp_path):
    """ASCII の字の後ろに韓国語の字のある値（「a성」）：Python のエンコーダは、a と 성 の間に印を書く。もう印のある
    ファイルでは、2 つ目の印になるので書かない。印のないファイルでは、書く（Python・Windows が書くのと同じ形）。"""
    top = "성,b\n".encode("iso2022_kr")
    assert "a성".encode("iso2022_kr") == b"a" + KR + b"\x0e<:\x0f"
    assert _kr(tmp_path, 1, top, {"R1.C2": "x길"}) == KR + b"\x0e<:\x0f,x\x0e1f\x0f\n"
    assert _kr(tmp_path, 2, top, {"R1.C1": "a성"}) == KR + b"\x0fa\x0e<:\x0f,b\n"
    late = b"a,b\n1," + KR + b"\x0eH+\x0f\n"
    assert _kr(tmp_path, 3, late, {"R1.C1": "z성"}) == b"z" + KR + b"\x0e<:\x0f,b\n1,\x0eH+\x0f\n"


def test_a_file_left_with_only_the_header_keeps_its_encoding(tmp_path):
    """韓国語の字がなくなって、印（ESC $ ) C）と ASCII の字だけが残ったファイル：Python の表では ISO-2022-JP-2 としても
    読めるが、字は ASCII だけで、元の文字コードで読んで同じ字なので、文字コードは変わっていない（前は、証明が
    「文字コードが変わっています」と止めていた）。"""
    rfc = KR + b"abc,\x0eH+\x0f\n"
    new = _kr(tmp_path, 1, rfc, {"R1.C2": "x"})
    assert decode(new).encoding != "iso2022_kr" and decode(new).text == new.decode("iso2022_kr") == "abc,x\n"
    p = _file(tmp_path, "k.json", b'{"name": "' + KR + b'\x0e1h\x0f"}\n')
    out = tmp_path / "o.json"
    r = api.fill(p, {"/name": "Kim"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"], r
    # JSON は、引用符ごと値を置き換える：後ろの状態を元と同じにするために、元の印を後ろに置く（字は同じ）
    assert out.read_bytes() == b'{"name": "Kim"' + KR + b'\x0e\x0f}\n'
    assert decode(out.read_bytes()).text == '{"name": "Kim"}\n'


# ---------------------------------------------------------------------- メール
ISO_BODY = "①の件、髙橋です。〜\r\n"


@pytest.mark.parametrize("ctype", [b"", b"Content-Type: text/plain\r\n", b"Content-Type: text/plain; charset=us-ascii\r\n"])
def test_mail_body_without_a_charset_is_read_by_its_escape_marks(ctype):
    """宣言がない（か us-ascii の）本文が、ASCII の字と ESC だけ：.txt と同じ決まりで、印のとおりに読み、そのことを伝える。"""
    std = b"Subject: x\r\n" + ctype + b"\r\n" + "山田です。\r\n".encode("iso2022_jp")
    m = mail.read_eml(std)
    assert m.body == "山田です。\r\n"
    assert len(m.notes) == 1 and "印のとおりに、ISO-2022-JP として読みました" in m.notes[0]
    win = b"Subject: x\r\n" + ctype + b"\r\n" + ISO_BODY.encode("iso2022_jp_ms")
    m = mail.read_eml(win)
    assert m.body == ISO_BODY and len(m.notes) == 1 and "CP50221" in m.notes[0]


def test_mail_body_with_a_declared_charset_has_no_note():
    m = mail.read_eml(b"Subject: x\r\nContent-Type: text/plain; charset=iso-2022-jp\r\n\r\n" + "山田です。\r\n".encode("iso2022_jp"))
    assert m.body == "山田です。\r\n" and m.notes == []
    assert mail.read_eml(b"Subject: x\r\n\r\nplain body\r\n").notes == []


def test_mail_body_with_control_characters_is_told():
    """印として読めない ESC（画面の色の印など）・NUL は、字としてそのまま残し、いくつあるかを伝える。"""
    m = mail.read_eml(b"Subject: x\r\n\r\nabc\x1b[0m \x00\r\n")
    assert "\x1b[0m" in m.body
    assert len(m.notes) == 1 and "制御文字（ESC・NUL など、タブ・改行以外）が 2 個あります" in m.notes[0]
