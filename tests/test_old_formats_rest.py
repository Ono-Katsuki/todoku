"""古い形式の残り：.ppt の絵・.msg の圧縮 RTF・.dbf のメモと上書き・Lotus の数式・Access のパスワード。手で作ったファイルで確かめる。"""

from __future__ import annotations

import base64
import random
import struct
import zipfile
import zlib

import jet_maker
import pytest
from test_embeds import make_cfb
from test_more_formats import _dbf

from todoku import api
from todoku.book import kind_of, open_book
from todoku.legacy import lotus, mail
from todoku.legacy.jet import _rc4

PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC")


def _p(tmp_path, name, blob):
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


# ---------------------------------------------------------------- .ppt の絵
def _ppt() -> bytes:
    def rec(ver_inst, typ, body):
        return struct.pack("<HHI", ver_inst, typ, len(body)) + body

    png_rec = rec(0x6E00, 0xF01E, bytes(16) + b"\xff" + PNG)
    jpg = b"\xff\xd8\xff\xe0" + bytes(20)
    jpg_rec = rec(0x46A0, 0xF01D, bytes(16) + b"\xff" + jpg)
    slide = rec(0x000F, 0x03EE, b"")
    doc = rec(0x000F, 0x03E8, rec(0x0001, 0x03E9, bytes(40)))
    edit_at = len(doc) + len(slide)
    pd = rec(0x0, 0x1772, struct.pack("<I", (2 << 20) | 1) + struct.pack("<II", 0, len(doc)))
    pd_at = edit_at
    ue_at = pd_at + len(pd)
    ue = rec(0x0, 0x0FF5, struct.pack("<IIIII", 0, 0, 0, pd_at, 1) + struct.pack("<I", 2) + bytes(4))
    stream = doc + slide + pd + ue
    cu = rec(0x0, 0x0FF6, struct.pack("<IIIHHHBBHH", 0x14, 0xE391C05F, ue_at, 0, 0, 0, 0, 0, 0, 0) + bytes(8))
    return make_cfb({"PowerPoint Document": stream, "Current User": cu, "Pictures": png_rec + jpg_rec})


def test_ppt_pictures_listed_and_viewable(tmp_path):
    blob = _ppt()
    assert kind_of(blob) == "ppt"
    p = _p(tmp_path, "a.ppt", blob)
    k, doc, parts = api.window(p)
    assert k == "ppt"
    ims = doc["images"]
    assert ims[0]["format"] == "png" and len(ims) == 2
    assert "blip1" in "".join(t for _n, t in parts)
    v = api.view_image(p, "blip1")
    assert v["mime"] == "image/png" and base64.b64decode(v["data_base64"]) == PNG
    with pytest.raises(api.ApiError):
        api.view_image(p, "blip99")


# ---------------------------------------------------------------- .msg の圧縮 RTF
def _crc(b: bytes) -> int:
    t = mail._crc_table()
    c = 0
    for x in b:
        c = t[(c ^ x) & 255] ^ (c >> 8)
    return c


def lzfu(raw: bytes, refs: bool = True) -> bytes:
    """MS-OXRTFCP の符号化（辞書の参照つき）。"""
    dic = bytearray(mail._RTF_DICT) + bytearray(4096 - len(mail._RTF_DICT))
    wp = len(mail._RTF_DICT)
    out = bytearray()
    i = 0
    n = len(raw)
    done = False
    while not done:
        ctl = 0
        chunk = bytearray()
        for bit in range(8):
            if i >= n:
                ctl |= 1 << bit
                chunk += struct.pack(">H", wp << 4)
                done = True
                break
            best, off = 0, 0
            if refs:
                for o in range(4096):
                    k = 0
                    if (o + 0) & 4095 == wp:
                        continue
                    while k < 17 and i + k < n and dic[(o + k) & 4095] == raw[i + k] and (o + k) & 4095 != wp:
                        k += 1
                    if k > best:
                        best, off = k, o
                        if k == 17:
                            break
            if best >= 2:
                ctl |= 1 << bit
                chunk += struct.pack(">H", off << 4 | (best - 2))
                for k in range(best):
                    dic[wp] = raw[i + k]
                    wp = (wp + 1) & 4095
                i += best
            else:
                chunk.append(raw[i])
                dic[wp] = raw[i]
                wp = (wp + 1) & 4095
                i += 1
        out.append(ctl)
        out += chunk
    return struct.pack("<IIII", len(out) + 12, len(raw), 0x75465A4C, _crc(bytes(out))) + bytes(out)


RTF = b"{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\pard Hello body\\par second line of the mail\\par}"
HTML_RTF = (b"{\\rtf1\\ansi\\ansicpg932\\fromhtml1 \\deff0{\\fonttbl{\\f0 Arial;}}{\\*\\htmltag64 <html>}{\\*\\htmltag64 <p>}\\htmlrtf \\par\\htmlrtf0 "
            b"HTML body text{\\*\\htmltag72 </p>}\\par}")


@pytest.mark.parametrize("rtf,word", [(RTF, "second line"), (HTML_RTF, "HTML body text")])
def test_compressed_rtf_body(tmp_path, rtf, word):
    assert mail.rtf_decompress(lzfu(rtf)) == rtf
    u16 = lambda s: s.encode("utf-16-le")  # noqa: E731
    blob = make_cfb({"__properties_version1.0": bytes(32), "__substg1.0_0037001F": u16("件名"), "__substg1.0_10090102": lzfu(rtf)})
    p = _p(tmp_path, "a.msg", blob)
    t = "".join(x for _n, x in api.window(p)[2])
    assert word in t and "<p>" not in t and "RTF" in t


def test_compressed_rtf_uncompressed_and_bad_crc():
    raw = b"{\\rtf1 x}"
    mela = struct.pack("<IIII", len(raw) + 12, len(raw), 0x414C454D, 0) + raw
    assert mail.rtf_decompress(mela) == raw
    b = bytearray(lzfu(RTF))
    b[20] ^= 1
    with pytest.raises(mail.MailError, match="CRC"):
        mail.rtf_decompress(bytes(b))
    with pytest.raises(mail.MailError):
        mail.rtf_decompress(b"short")
    with pytest.raises(mail.MailError):
        mail.rtf_decompress(struct.pack("<IIII", 12, 1, 0x12345678, 0))


def test_broken_rtf_in_msg_is_reported(tmp_path):
    blob = make_cfb({"__properties_version1.0": bytes(32), "__substg1.0_10090102": b"\x01\x02"})
    t = "".join(x for _n, x in api.window(_p(tmp_path, "a.msg", blob))[2])
    assert "解けません" in t


# ---------------------------------------------------------------- .dbf
def _memo_rows():
    return [(b" ", "山田".encode("cp932").ljust(12), b"   12.50", b"20260401", b"T", b"         1"),
            (b" ", b"abc".ljust(12), b" " * 8, b" " * 8, b"?", b"         3")]


def _dbt(texts: dict[int, bytes]) -> bytes:
    out = bytearray(512 * 8)
    for n, t in texts.items():
        out[n * 512:n * 512 + len(t) + 2] = t + b"\x1a\x1a"
    return bytes(out)


def _fpt(texts: dict[int, bytes], block: int = 64) -> bytes:
    out = bytearray(block * 10)
    struct.pack_into(">H", out, 6, block)
    for n, t in texts.items():
        out[n * block:n * block + 8 + len(t)] = struct.pack(">II", 1, len(t)) + t
    return bytes(out)


def _with_version(b: bytes, v: int) -> bytes:
    return bytes((v,)) + b[1:]


def test_dbf_memo_from_sibling_dbt_and_fpt(tmp_path):
    d = _with_version(_dbf(rows=_memo_rows()), 0x83)
    (tmp_path / "a.dbf").write_bytes(d)
    p = str(tmp_path / "a.dbf")
    (tmp_path / "A.DBT").write_bytes(_dbt({1: "メモの字".encode("cp932"), 3: b"second"}))
    _k, doc, parts = api.window(p)
    t = "".join(x for _n, x in parts)
    assert "メモの字" in t and "second" in t
    (tmp_path / "A.DBT").unlink()
    (tmp_path / "a.fpt").write_bytes(_fpt({1: "ふぉっくす".encode("cp932"), 3: b"pro"}))
    t = "".join(x for _n, x in api.window(p)[2])
    assert "ふぉっくす" in t and "pro" in t


def test_dbf_memo_dbase4_and_missing_file_note(tmp_path):
    d = _with_version(_dbf(rows=_memo_rows()), 0x8B)
    (tmp_path / "b.dbf").write_bytes(d)
    t = "".join(x for _n, x in api.window(str(tmp_path / "b.dbf"))[2])
    assert "同じ名前のファイル" in t
    dbt = bytearray(1024 * 5)
    struct.pack_into("<H", dbt, 20, 1024)
    body = "四の形".encode("cp932")
    dbt[1024:1024 + 8 + len(body)] = b"\xff\xff\x08\x00" + struct.pack("<I", 8 + len(body)) + body
    (tmp_path / "b.dbt").write_bytes(bytes(dbt))
    t = "".join(x for _n, x in api.window(str(tmp_path / "b.dbf"))[2])
    assert "四の形" in t


def test_dbf_memo_from_zip_member(tmp_path):
    p = tmp_path / "z.zip"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("data/a.dbf", _with_version(_dbf(rows=_memo_rows()), 0x83))
        z.writestr("data/a.dbt", _dbt({1: "書庫の中".encode("cp932"), 3: b"x"}))
    t = "".join(x for _n, x in api.window(f"{p}!/data/a.dbf")[2])
    assert "書庫の中" in t


def test_dbf_same_size_writes(tmp_path):
    blob = _dbf(rows=_memo_rows())
    p = _p(tmp_path, "a.dbf", blob)
    out = tmp_path / "o.dbf"
    r = api.fill(p, {"a!A2": "佐藤", "a!B2": 7.25, "a!C2": "2026-11-03", "a!D2": False}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    new = out.read_bytes()
    assert len(new) == len(blob)
    sh = open_book(new, "o.dbf").sheets[0]
    assert sh.cells[(2, 1)].value == "佐藤" and sh.cells[(2, 2)].value == 7.25 and sh.cells[(2, 3)].value == "2026-11-03" and sh.cells[(2, 4)].value is False
    assert sh.cells[(3, 1)].value == "abc"
    assert new[:32] == blob[:32]  # 見出しは変えない


@pytest.mark.parametrize("key,v,word", [("a!A2", "長" * 7, "幅"), ("a!B2", 1.234, "小数"), ("a!B2", 123456.5, "幅"), ("a!B2", "字", "数の項目"), ("a!C2", "去年", "日付"),
                                        ("a!D2", "はい", "真偽"), ("a!E2", "メモ", "メモ"), ("a!A1", "題", "項目の名前"), ("a!A9", "x", "表の外")])
def test_dbf_refuses_what_does_not_fit(tmp_path, key, v, word):
    p = _p(tmp_path, "a.dbf", _dbf(rows=_memo_rows()))
    out = tmp_path / "o.dbf"
    with pytest.raises(api.ApiError, match=word):
        api.fill(p, {key: v}, str(out))
    assert not out.exists()


def test_dbf_skips_deleted_rows_on_write(tmp_path):
    blob = _dbf()  # 2 件目は削除の印
    p = _p(tmp_path, "a.dbf", blob)
    out = tmp_path / "o.dbf"
    assert api.fill(p, {"a!A3": "第三"}, str(out))["proof_ok"]
    new = out.read_bytes()
    assert b"deleted" in new and open_book(new, "o.dbf").sheets[0].cells[(3, 1)].value == "第三"


def test_dbf_grade(tmp_path):
    g = api.grade(_p(tmp_path, "a.dbf", _dbf()))
    assert g["write"]["grade"] == "B"


# ---------------------------------------------------------------- Lotus の数式
def wk1(formulas: list[tuple[int, int, float, bytes]], cells: bytes = b"") -> bytes:
    def rec(t, b):
        return struct.pack("<HH", t, len(b)) + b

    out = rec(0, struct.pack("<H", 0x0406)) + cells
    for c, r, v, code in formulas:
        out += rec(16, bytes((0xFF,)) + struct.pack("<HHdH", c, r, v, len(code)) + code)
    return out + rec(1, b"")


def _rel(dc: int) -> int:
    return 0x8000 | (dc & 0x3FFF)


def num_rec(c, r, v):
    return struct.pack("<HH", 14, 13) + bytes((0xFF,)) + struct.pack("<HHd", c, r, v)


def test_formula_text_operators_and_refs():
    t = lambda code, c=0, r=0: lotus.formula_text(code, c, r, lambda b: b.decode("cp932"))  # noqa: E731
    ref = lambda c, r: bytes((1,)) + struct.pack("<HH", c, r)  # noqa: E731
    assert t(ref(1, 2) + b"\x03") == ("$B$3", True)
    assert t(bytes((1,)) + struct.pack("<HH", _rel(1), _rel(-1)) + b"\x03", 2, 5) == ("D5", True)
    assert t(ref(0, 0) + bytes((5,)) + struct.pack("<h", 2) + b"\x09\x03") == ("$A$1+2", True)
    # (1+2)*3 と 1+2*3
    one = bytes((5,)) + struct.pack("<h", 1)
    two = bytes((5,)) + struct.pack("<h", 2)
    three = bytes((5,)) + struct.pack("<h", 3)
    assert t(one + two + b"\x09" + three + b"\x0b\x03")[0] == "(1+2)*3"
    assert t(one + two + three + b"\x0b\x09\x03")[0] == "1+2*3"
    assert t(one + two + b"\x0a" + three + b"\x0a\x03")[0] == "1-2-3"
    assert t(one + two + three + b"\x0a\x0a\x03")[0] == "1-(2-3)"
    assert t(two + b"\x08" + two + b"\x0d\x03")[0] == "-2^2"  # 符号が先
    assert t(two + two + b"\x0d\x08\x03")[0] == "-(2^2)"
    assert t(one + b"\x04" + b"\x03")[0] == "(1)"
    assert t(bytes((0,)) + struct.pack("<d", 1.5) + b"\x03")[0] == "1.5"
    assert t(bytes((6,)) + "字".encode("cp932") + b"\0\x03")[0] == '"字"'
    assert t(one + two + b"\x14\x03")[0] == "AND(1,2)" and t(one + b"\x16\x03")[0] == "NOT(1)"
    # 関数：引数の数のバイト
    rng = bytes((2,)) + struct.pack("<HHHH", 0, 0, 0, 2)
    assert t(rng + bytes((0x50, 1)) + b"\x03") == ("SUM($A$1:$A$3)", True)
    assert t(rng + one + bytes((0x51, 2)) + b"\x03") == ("AVERAGE($A$1:$A$3,1)", True)
    assert t(rng + bytes((0x52, 1)) + b"\x03") == ("@COUNT($A$1:$A$3)", False)
    assert t(one + two + three + b"\x3b\x03") == ("IF(1,2,3)", True)
    assert t(one + b"\x22\x03") == ("@INT(1)", False)


@pytest.mark.parametrize("code", [b"", b"\x03", b"\x09\x03", b"\x01\x00", bytes((0x7E,)) + b"\x03", b"\x06abc", bytes((0x50, 5)) + b"\x03",
                                  bytes((5, 1, 0, 5, 2, 0)) + b"\x03", bytes((1,)) + struct.pack("<HH", 0x8000 | 0x2000, 0) + b"\x03"])
def test_formula_bad_codes_keep_saved_value(code):
    blob = wk1([(0, 0, 42.0, code)])
    sh = open_book(blob, "a.wk1").sheets[0]
    c = sh.cells[(1, 1)]
    assert c.value == 42.0 and c.formula == "?" and "未対応" in c.parse_error


def test_wk1_formula_in_window_and_calc(tmp_path):
    sum_code = bytes((2,)) + struct.pack("<HHHH", _rel(0), _rel(-2), _rel(0), _rel(-1)) + bytes((0x50, 1)) + b"\x03"
    int_code = bytes((5,)) + struct.pack("<h", 1) + b"\x22\x03"
    blob = wk1([(0, 2, 30.0, sum_code), (1, 2, 1.0, int_code)], num_rec(0, 0, 10.0) + num_rec(0, 1, 20.0))
    wb = open_book(blob, "a.wk1")
    sh = wb.sheets[0]
    assert sh.cells[(3, 1)].formula == "SUM(A1:A2)" and sh.cells[(3, 1)].parse_error in (None, "")
    assert sh.cells[(3, 2)].formula == "@INT(1)" and "未対応" in sh.cells[(3, 2)].parse_error
    p = _p(tmp_path, "a.wk1", blob)
    doc = api.window(p)[1]
    assert "SUM(A1:A2)" in str(doc) and "@INT(1)" in str(doc)
    from todoku.xlsx.calc import verify

    rep = verify(wb).summary()
    assert rep["mismatch"] == 0 and rep["unsupported"] >= 1


def test_wk3_and_works_keep_next_step(tmp_path):
    p = _p(tmp_path, "a.wk3", b"\x00\x00\x1a\x00\x00\x10\x04\x00" + bytes(30))
    with pytest.raises(api.ApiError, match="公開の資料"):
        api.window(p)
    p = _p(tmp_path, "a.wps", make_cfb({"CONTENTS": bytes(600)}))
    with pytest.raises(api.ApiError, match="公開されていない"):
        api.window(p)
    p = _p(tmp_path, "a.xdw", b"anything" * 20)
    with pytest.raises(api.ApiError, match="公開されていない"):
        api.window(p)


# ---------------------------------------------------------------- Access のパスワード
def _with_password(blob: bytes, pw: str, jet3: bool, date: float = 45000.25) -> bytes:
    size = 126 if jet3 else 128
    head = bytearray(blob[:2048 if jet3 else 4096])
    key = struct.pack("<I", 0x6B39DAC7)
    body = bytearray(_rc4(key, bytes(head[0x18:0x18 + size])))
    at = 0x42 - 0x18
    if jet3:
        raw = pw.encode("cp932").ljust(20, b"\0")
        body[at:at + 20] = raw
    else:
        struct.pack_into("<d", body, 0x72 - 0x18, date)
        mask = int(date).to_bytes(4, "little")
        raw = bytearray(pw.encode("utf-16-le").ljust(40, b"\0"))
        for i in range(40):
            raw[i] ^= mask[i % 4]
        body[at:at + 40] = raw
    head[0x18:0x18 + size] = _rc4(key, bytes(body))
    return bytes(head) + blob[len(head):]


@pytest.mark.parametrize("jet3", [False, True])
def test_access_password_required_and_compared(tmp_path, jet3):
    plain = jet_maker.sample(jet3)
    blob = _with_password(plain, "ひみつ1", jet3)
    p = _p(tmp_path, "a.mdb", blob)
    with pytest.raises(api.PasswordNeeded):
        api.window(p)
    with pytest.raises(api.PasswordNeeded):
        api.grade(p)
    with pytest.raises(api.PasswordWrong):
        api.window(p, password="ちがう")
    k, doc, _parts = api.window(p, password="ひみつ1")
    assert k == "mdb" and doc["sheets"]
    assert api.summary(_p(tmp_path, "b.mdb", plain))["kind"] == "mdb"  # パスワードなし


def test_access_empty_password_jet4_with_date_mask(tmp_path):
    blob = _with_password(jet_maker.sample(False), "", False)
    assert api.summary(_p(tmp_path, "a.mdb", blob))["kind"] == "mdb"


def test_access_password_field_not_text_says_so(tmp_path):
    plain = jet_maker.sample(False)
    head = bytearray(plain[:4096])
    key = struct.pack("<I", 0x6B39DAC7)
    body = bytearray(_rc4(key, bytes(head[0x18:0x18 + 128])))
    body[0x42 - 0x18:0x42 - 0x18 + 40] = bytes([0x01, 0x00]) * 20
    head[0x18:0x18 + 128] = _rc4(key, bytes(body))
    p = _p(tmp_path, "a.mdb", bytes(head) + plain[4096:])
    _k, doc, parts = api.window(p)
    assert "確かめられません" in "".join(x for _n, x in parts)


def test_ace_unreadable_tables_message(tmp_path):
    blob = bytearray(jet_maker.sample(False))
    blob[4:19] = b"Standard ACE DB"
    blob[0x14] = 2
    blob[2 * 4096] = 0x55  # 表の定義のページを壊す（暗号にされたページの代わり）
    p = _p(tmp_path, "a.accdb", bytes(blob))
    with pytest.raises(api.ApiError, match="暗号化"):
        api.window(p)


# ---------------------------------------------------------------- 壊れたファイル
def test_corruption_never_raw(tmp_path):
    r = random.Random(8)
    u16 = lambda s: s.encode("utf-16-le")  # noqa: E731
    bases = {"a.msg": make_cfb({"__properties_version1.0": bytes(32), "__substg1.0_0037001F": u16("件名"), "__substg1.0_10090102": lzfu(RTF)}),
             "a.dbf": _with_version(_dbf(rows=_memo_rows()), 0x83),
             "a.wk1": wk1([(0, 0, 1.0, bytes((5, 1, 0)) + b"\x03")]),
             "a.ppt": _ppt(), "a.mdb": _with_password(jet_maker.sample(False), "x", False)}
    (tmp_path / "a.dbt").write_bytes(_dbt({1: b"memo"}))
    for name, blob in bases.items():
        p = tmp_path / name
        for t in range(25):
            b = bytearray(blob)
            for _ in range(r.randrange(1, 12)):
                b[r.randrange(len(b))] = r.randrange(256)
            p.write_bytes(bytes(b))
            try:
                api.summary(str(p), password="x" if name == "a.mdb" else None)
            except api.ApiError:
                pass
            except (zlib.error, ValueError) as e:
                raise AssertionError(f"{name}: {e!r}") from e
