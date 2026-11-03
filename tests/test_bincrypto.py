"""パスワードで暗号化された古い Excel（.xls）と Word（.doc）を解き、同じパスワードで暗号化し直すテスト。

正解は、Office が作った暗号化のファイルと、その暗号を解いたファイル（corpus/crypto。msoffcrypto-tool のテストのファイル、MIT）。
ないときは飛ばす。Office のファイルがない方式（RC4 の版 1.1、.doc の XOR、40 ビットの RC4 CryptoAPI）と、小さな
ストリーム（ミニストリーム）は、ここで小さな複合ファイルを組み立て、暗号化して解き戻して確かめる（corpus がなくても動く）。
"""

from __future__ import annotations

import re
import struct
from pathlib import Path

import pytest

from todoku import bincrypto
from todoku.bincrypto import (DEFAULT_XLS_PASSWORD, BinKey, CryptoError, WrongPassword, decrypt_binary, encrypt_binary,
                              is_encrypted_binary, is_encrypted_doc, is_encrypted_xls, new_key)
from todoku.cfb import CFB
from todoku.doc.book import DocFile
from todoku.doc.fill import fill as doc_fill
from todoku.docx.window import build as doc_window
from todoku.pdf.crypto import rc4
from todoku.xls.book import XlsWorkbook
from todoku.xlsx.edit import fill

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
CRYPTO = CORPUS / "crypto"
CASES = [("rc4cryptoapi_password.xls", "Password1234_"), ("rc4cryptoapi_password.doc", "Password1234_"),
         ("xor_password_123456789012345.xls", "123456789012345")]
# 暗号化されるストリーム（それ以外は、解いても暗号化し直してもバイト一致）
ENCRYPTED = {"Workbook", "WordDocument", "1Table", "0Table", "Data"}


def pair(name: str) -> tuple[bytes, bytes]:
    enc = CRYPTO / "inputs" / name
    plain = CRYPTO / "outputs" / name.replace(".", "_plain.")
    if not enc.exists() or not plain.exists():
        pytest.skip("corpus/crypto がありません")
    return enc.read_bytes(), plain.read_bytes()


def streams(blob: bytes) -> dict[str, bytes]:
    c = CFB(blob)
    return {p: c.read(p) for p in c.streams()}


# ====================================================================== Office のファイル（正解つき）
@pytest.mark.corpus
def test_xls_rc4_cryptoapi_matches_known_output():
    enc, ref = pair("rc4cryptoapi_password.xls")
    assert is_encrypted_xls(enc) and is_encrypted_binary(enc) and not is_encrypted_doc(enc)
    plain, key = decrypt_binary(enc, "Password1234_")
    assert (key.kind, key.method, key.key_bits, key.label) == ("xls", "cryptoapi", 128, "RC4 CryptoAPI")
    assert not is_encrypted_xls(plain) and len(plain) == len(enc)
    got, want, orig = streams(plain), streams(ref), streams(enc)
    # FilePass は、同じ位置・同じ長さの「種類 0・中身がすべて 0」に置き換わり、ブックのストリームは正解とバイト一致
    assert got == want
    assert got["Workbook"][20:24] == struct.pack("<HH", 0, 200) and orig["Workbook"][20:24] == struct.pack("<HH", 0x2F, 200)
    assert all(got[p] == orig[p] for p in got if p not in ENCRYPTED)


@pytest.mark.corpus
def test_doc_rc4_cryptoapi_matches_known_output():
    enc, ref = pair("rc4cryptoapi_password.doc")
    assert is_encrypted_doc(enc) and not is_encrypted_xls(enc)
    plain, key = decrypt_binary(enc, "Password1234_")
    assert (key.kind, key.method, key.lkey, key.fib_bits) == ("doc", "cryptoapi", 198, 0x0100)
    assert not is_encrypted_doc(plain)
    got, want, orig = streams(plain), streams(ref), streams(enc)
    # WordDocument は正解とバイト一致（fEncrypted と lKey は 0）
    assert got["WordDocument"] == want["WordDocument"]
    assert struct.unpack_from("<HHI", got["WordDocument"], 0x0A) == (0x12F0, 0x00BF, 0)
    # Table の先頭の 198 バイト（EncryptionHeader）は、正解では「解いた」意味のないバイト。こちらは平文の暗号の情報を残す
    assert got["1Table"][198:] == want["1Table"][198:]
    assert got["1Table"][:198] == orig["1Table"][:198] == key.marker
    assert all(got[p] == orig[p] for p in got if p not in ENCRYPTED)


@pytest.mark.corpus
def test_xls_xor_matches_known_output():
    # 正解のファイルは保存し直したもの（ストリームの長さも違う）なので、セルの値で比べる
    enc, ref = pair("xor_password_123456789012345.xls")
    plain, key = decrypt_binary(enc, "123456789012345")
    assert (key.kind, key.method) == ("xls", "xor")
    a, b = XlsWorkbook(plain), XlsWorkbook(ref)
    assert [s.name for s in a.sheets] == [s.name for s in b.sheets] == ["420"]
    va = {k: (c.value, c.formula) for k, c in a.sheet("420").cells.items()}
    vb = {k: (c.value, c.formula) for k, c in b.sheet("420").cells.items()}
    assert len(va) == 10920 and va == vb
    # 印：FilePass（6 バイト：種類 0・XorKey・検証の値）は種類 0 の 6 バイトの 0 に。ほかのストリームはそのまま
    assert streams(plain)["Workbook"][20:30] == struct.pack("<HH", 0, 6) + bytes(6)
    assert all(v == streams(enc)[p] for p, v in streams(plain).items() if p not in ENCRYPTED)


@pytest.mark.corpus
@pytest.mark.parametrize("name, password", CASES)
def test_wrong_or_missing_password(name, password):
    enc, _ref = pair(name)
    with pytest.raises(WrongPassword) as e:
        decrypt_binary(enc, "x" + password)  # XOR は 15 字までしか使わないので、後ろに足すと同じパスワードになる
    assert "違います" in str(e.value) and password not in str(e.value)  # パスワードは出さない
    with pytest.raises(WrongPassword, match="パスワードを教えて"):
        decrypt_binary(enc)


@pytest.mark.corpus
def test_readers_open_the_decrypted_files():
    enc, _ = pair("rc4cryptoapi_password.xls")
    wb = XlsWorkbook(decrypt_binary(enc, "Password1234_")[0])
    cells = wb.sheet("Sheet1").cells
    assert cells[(1, 1)].value == "lorem ipsum" and cells[(1, 2)].value == 3.0 and cells[(1, 2)].formula == "1+2"
    enc, _ = pair("rc4cryptoapi_password.doc")
    doc = DocFile(decrypt_binary(enc, "Password1234_")[0])
    assert doc.text == "lorem ipsum\r"
    assert [p.text for _k, _p, blocks in doc.all_blocks() for p in blocks] == ["lorem ipsum"]


@pytest.mark.corpus
@pytest.mark.parametrize("name, password", CASES)
def test_reencrypting_unedited_file_reproduces_original(name, password):
    enc, _ = pair(name)
    plain, key = decrypt_binary(enc, password)
    again = encrypt_binary(key, plain)
    assert again == enc  # 同じ salt・同じ検証の値・同じ鍵の流れなので、ファイル全体がバイト一致
    plain2, key2 = decrypt_binary(again, password)
    assert plain2 == plain and key2 == key
    assert password not in repr(key) and key.secret.hex() not in repr(key)  # 鍵の repr に中身を出さない
    with pytest.raises(CryptoError, match="すでに暗号化"):
        encrypt_binary(key, enc)  # 二重に暗号化しない


@pytest.mark.corpus
@pytest.mark.parametrize("name, password, sheet", [("rc4cryptoapi_password.xls", "Password1234_", "Sheet1"),
                                                   ("xor_password_123456789012345.xls", "123456789012345", "420")])
def test_xls_edit_survives_reencryption(name, password, sheet):
    enc, _ = pair(name)
    plain, key = decrypt_binary(enc, password)
    res = fill(XlsWorkbook(plain), {f"{sheet}!A1": "とどく", f"{sheet}!C5": 42})
    out = encrypt_binary(key, res.blob)
    assert is_encrypted_xls(out)
    plain2, key2 = decrypt_binary(out, password)
    assert plain2 == res.blob and key2 == key
    cells = XlsWorkbook(plain2).sheet(sheet).cells
    assert cells[(1, 1)].value == "とどく" and cells[(5, 3)].value == 42
    # 暗号の情報（FilePass：salt と検証の値）は元のまま。ブック以外のストリームもそのまま
    a, b = streams(enc), streams(out)
    pa, la = bincrypto._filepass(a["Workbook"])
    pb, lb = bincrypto._filepass(b["Workbook"])
    assert a["Workbook"][pa:pa + 4 + la] == b["Workbook"][pb:pb + 4 + lb]
    assert all(b[p] == v for p, v in a.items() if p != "Workbook")


@pytest.mark.corpus
def test_doc_edit_survives_reencryption():
    enc, _ = pair("rc4cryptoapi_password.doc")
    plain, key = decrypt_binary(enc, "Password1234_")
    doc = DocFile(plain)
    field = doc_window(doc)["fields"][0]["id"]
    res = doc_fill(doc, {field: "とどく"})
    out = encrypt_binary(key, res.blob)
    assert is_encrypted_doc(out)
    assert struct.unpack_from("<HI", streams(out)["WordDocument"], 0x0A)[0] & 0x0100
    plain2, key2 = decrypt_binary(out, "Password1234_")
    assert plain2 == res.blob and key2 == key
    after = DocFile(plain2)
    assert after.text == "lorem ipsumとどく\r"
    assert {f["id"]: f["value"] for f in doc_window(after)["fields"]}[field] == "とどく"
    assert streams(out)["1Table"][:198] == streams(enc)["1Table"][:198]  # 暗号の情報は元のまま


@pytest.mark.corpus
def test_new_key_reproduces_office_headers():
    # 同じ salt・同じ検証の値から作った暗号の情報が、Office の書いたものとバイト一致する
    for name, field in (("rc4cryptoapi_password.xls", 2), ("rc4cryptoapi_password.doc", 0)):
        enc, _ = pair(name)
        _plain, key = decrypt_binary(enc, "Password1234_")
        info = key.marker[field:]
        v = info[12 + struct.unpack_from("<I", info, 8)[0]:]
        salt, verifier = v[4:20], rc4(key._block_key(0), v[20:36])
        made = new_key(key.kind, "cryptoapi", "Password1234_", salt=salt, verifier=verifier, key_bits=128)
        assert made == key
    enc, _ = pair("xor_password_123456789012345.xls")
    assert new_key("xls", "xor", "123456789012345") == decrypt_binary(enc, "123456789012345")[1]


# ====================================================================== 仕様の表と手順
@pytest.mark.corpus
def test_xor_tables_match_the_spec_text():
    spec = CORPUS / "specs" / "MS-OFFCRYPTO.txt"
    if not spec.exists():
        pytest.skip("仕様の文書がありません")
    text = spec.read_text(encoding="utf-8")

    def table(name: str) -> list[int]:
        m = re.search(r"SET " + name + r" TO \((.*?)\)", text, re.S)
        return [int(x, 16) for x in re.findall(r"0x([0-9A-Fa-f]+)", m[1])]

    assert table("PadArray") == list(bincrypto._PAD)
    assert table("InitialCode") == list(bincrypto._INITIAL)
    assert table("XorMatrix") == list(bincrypto._MATRIX) and len(bincrypto._MATRIX) == 105


@pytest.mark.parametrize("n", range(1, 16))
def test_xor_array_methods_1_and_2_agree(n):
    # 2.3.7.2（Method 1：.xls と .doc が使う手順）と 2.3.7.5（Method 2）は、同じ表になる
    for pw in (b"abcdefghijklmno"[:n], bytes(range(0x81, 0x90))[:n]):
        key = bincrypto._xor_key(pw)
        src = list(pw) + list(bincrypto._PAD[:16 - n])
        two = bytes(bincrypto._ror(src[i] ^ (key & 0xFF if i % 2 == 0 else key >> 8)) for i in range(16))
        assert bincrypto._xor_array(pw) == two


def test_xor_password_bytes():
    # 各文字の下位のバイト（0 なら上位のバイト）を 15 字まで。読むときは日本語の ANSI（cp932）も試す
    assert bincrypto._xor_passwords("abc") == [b"abc"]
    assert bincrypto._xor_passwords("x" * 20) == [b"x" * 15]
    assert bincrypto._xor_passwords("あ") == [b"\x42", "あ".encode("cp932")]
    assert bincrypto._xor_passwords("") == []


# ====================================================================== Office のファイルがなくても：組み立てて確かめる
_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_END, _FREE, _FATSECT, _NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


def make_cfb(parts: dict[str, bytes]) -> bytes:
    """小さな複合ファイル（版 3・512 バイトのセクタ）。4096 バイト未満のストリームはミニストリーム（64 バイトのセクタ）に置く。"""
    ss, ms = 512, 64
    names = sorted(parts, key=lambda n: (len(n), n.upper()))
    sectors: list[bytes] = []
    fat: list[int] = []

    def put(data: bytes) -> int:
        if not data:
            return _END
        start, k = len(sectors), -(-len(data) // ss)
        for i in range(k):
            sectors.append(data[i * ss:(i + 1) * ss].ljust(ss, b"\0"))
            fat.append(start + i + 1 if i < k - 1 else _END)
        return start

    mini, minifat, starts = bytearray(), [], {}
    for n in names:
        d = parts[n]
        if len(d) >= 4096:
            starts[n] = put(d)
        else:
            s0, k = len(mini) // ms, -(-len(d) // ms)
            mini += d.ljust(k * ms, b"\0")
            minifat += [s0 + i + 1 for i in range(k - 1)] + [_END]
            starts[n] = s0
    mini_at = put(bytes(mini))
    minifat_at = put(struct.pack(f"<{len(minifat)}I", *minifat)) if minifat else _END

    def entry(name: str, kind: int, start: int, size: int, child: int = _NOSTREAM, right: int = _NOSTREAM) -> bytes:
        raw = (name + "\0").encode("utf-16-le") if name else b""
        return (raw.ljust(64, b"\0") + struct.pack("<HBB3I", len(raw), kind, 1 if name else 0, _NOSTREAM, right, child)
                + bytes(36) + struct.pack("<IQ", start, size))

    ents = [entry("Root Entry", 5, mini_at, len(mini), child=1)]
    for i, n in enumerate(names):
        ents.append(entry(n, 2, starts[n], len(parts[n]), right=i + 2 if i + 1 < len(names) else _NOSTREAM))
    while len(ents) % 4:
        ents.append(entry("", 0, 0, 0))
    dir_at = put(b"".join(ents))
    fat_at = len(sectors)
    fat.append(_FATSECT)
    assert len(fat) <= ss // 4
    sectors.append(struct.pack(f"<{ss // 4}I", *(fat + [_FREE] * (ss // 4 - len(fat)))))
    head = (_MAGIC + bytes(16) + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
            + struct.pack("<9I", 0, 1, dir_at, 0, 4096, minifat_at, -(-len(minifat) * 4 // ss), _END, 0)
            + struct.pack("<109I", fat_at, *([_FREE] * 108)))
    return head + b"".join(sectors)


def rec(t: int, body: bytes = b"") -> bytes:
    return struct.pack("<HH", t, len(body)) + body


def make_workbook(marker_len: int) -> bytes:
    """解いた形の小さな BIFF8 のブック：FilePass の場所に種類 0 の印、シート1つ（A1 に文字、B2 に数）。"""
    def bof(dt: int) -> bytes:
        return rec(0x0809, struct.pack("<HHHHII", 0x0600, dt, 0x0DBB, 0x07CC, 0, 0x0206))

    name = b"Sheet1"
    head = (bof(0x0005) + rec(0x0000, bytes(marker_len)) + rec(0x00E1, b"\xb0\x04") + rec(0x00C1, b"\0\0")
            + rec(0x00E2) + rec(0x0042, b"\xb0\x04"))
    tail = rec(0x00FC, struct.pack("<IIHB", 1, 1, 3, 1) + "とどく".encode("utf-16-le")) + rec(0x000A)
    sheet_at = len(head) + 4 + 8 + len(name) + len(tail)
    boundsheet = rec(0x0085, struct.pack("<IBBBB", sheet_at, 0, 0, len(name), 0) + name)
    sheet = (bof(0x0010) + rec(0x0200, struct.pack("<IIHHH", 0, 2, 0, 2, 0))
             + rec(0x00FD, struct.pack("<HHHI", 0, 0, 15, 0)) + rec(0x0203, struct.pack("<HHHd", 1, 1, 15, 12.5))
             + rec(0x000A))
    return head + boundsheet + tail + sheet


def make_worddoc(header: bytes) -> dict[str, bytes]:
    """解いた形の小さな Word 97 の文書（本文「氏名：」の1段落）。Table の先頭には暗号の情報を平文で置く。"""
    text = "氏名：\r"
    wd = bytearray(4608)
    struct.pack_into("<HHHHHHHI", wd, 0, 0xA5EC, 0x00C1, 0, 0x0411, 0, 0x1200, 0x00BF, 0)  # 1Table を使う・lKey は 0
    struct.pack_into("<H", wd, 0x20, 14)  # csw
    struct.pack_into("<H", wd, 0x3E, 22)  # cslw
    struct.pack_into("<22i", wd, 0x40, len(wd), 0, 0, len(text), *([0] * 18))  # cbMac・ccpText
    struct.pack_into("<H", wd, 0x98, 93)  # cbRgFcLcb
    wd[1024:1024 + 2 * len(text)] = text.encode("utf-16-le")
    clx = b"\x02" + struct.pack("<I", 16) + struct.pack("<II", 0, len(text)) + struct.pack("<HIH", 0, 1024, 0)
    struct.pack_into("<II", wd, 0x9A + 33 * 8, len(header), len(clx))  # fcClx・lcbClx
    # Data（120 バイト）と \x01CompObj は小さいので、ミニストリームに入る（最後のミニセクタは途中まで）
    return {"WordDocument": bytes(wd), "1Table": header + clx, "Data": b"DATA" * 30, "\x01CompObj": bytes(range(114))}


SALT, VERIFIER = bytes(range(16)), bytes(range(100, 116))
METHODS = [("xor", 128), ("rc4", 128), ("cryptoapi", 128), ("cryptoapi", 40), ("cryptoapi", 56)]


@pytest.mark.parametrize("method, bits", METHODS)
def test_xls_round_trip_without_fixtures(method, bits):
    key = new_key("xls", method, "とどく123", salt=SALT, verifier=VERIFIER, key_bits=bits)
    plain = make_cfb({"Workbook": make_workbook(len(key.marker)), "\x05SummaryInformation": b"S" * 200})
    enc = encrypt_binary(key, plain)
    assert is_encrypted_xls(enc) and not is_encrypted_xls(plain) and len(enc) == len(plain)
    a, b = streams(plain), streams(enc)
    assert a["Workbook"] != b["Workbook"] and a["\x05SummaryInformation"] == b["\x05SummaryInformation"]
    got, key2 = decrypt_binary(enc, "とどく123")
    assert got == plain and key2 == key
    with pytest.raises(WrongPassword):
        decrypt_binary(enc, "とどく124")
    cells = XlsWorkbook(got).sheet("Sheet1").cells
    assert cells[(1, 1)].value == "とどく" and cells[(2, 2)].value == 12.5
    # 書き込んで（ブックは大きくなり、ミニストリームから普通のセクタに移る）、暗号化し直しても読める
    res = fill(XlsWorkbook(got), {"Sheet1!A1": "書いた字", "Sheet1!C3": 7})
    again, _ = decrypt_binary(encrypt_binary(key, res.blob), "とどく123")
    assert again == res.blob
    assert {k: c.value for k, c in XlsWorkbook(again).sheet("Sheet1").cells.items()} == \
        {(1, 1): "書いた字", (2, 2): 12.5, (3, 3): 7}


@pytest.mark.parametrize("method, bits", METHODS)
def test_doc_round_trip_without_fixtures(method, bits):
    key = new_key("doc", method, "とどく123", salt=SALT, verifier=VERIFIER, key_bits=bits)
    parts = make_worddoc(key.marker)
    plain = make_cfb(parts)
    assert DocFile(plain).text == "氏名：\r"
    enc = encrypt_binary(key, plain)
    assert is_encrypted_doc(enc) and not is_encrypted_doc(plain) and len(enc) == len(plain)
    a, b = streams(plain), streams(enc)
    assert b["\x01CompObj"] == a["\x01CompObj"]
    for p in ("WordDocument", "1Table", "Data"):
        assert a[p] != b[p]
    flags, _back, lkey = struct.unpack_from("<HHI", b["WordDocument"], 0x0A)
    assert flags & 0x0100 and bool(flags & 0x8000) == (method == "xor") and lkey == key.lkey
    head = bytearray(a["WordDocument"][:68])  # 先頭の 68 バイトは、印（fEncrypted・fObfuscated・lKey）のほかは平文のまま
    struct.pack_into("<H", head, 0x0A, 0x1200 | key.fib_bits)
    struct.pack_into("<I", head, 0x0E, key.lkey)
    assert b["WordDocument"][:68] == bytes(head)
    if method == "xor":  # XOR（Method 2）は 0 のバイトを変えない
        assert all(y == 0 for x, y in zip(a["WordDocument"][68:], b["WordDocument"][68:]) if x == 0)
    else:
        assert b["1Table"][:len(key.marker)] == key.marker  # 暗号の情報は平文
    got, key2 = decrypt_binary(enc, "とどく123")
    assert got == plain and key2 == key
    with pytest.raises(WrongPassword):
        decrypt_binary(enc, "とどく124")
    assert DocFile(got).text == "氏名：\r"


def test_default_password_of_write_protected_workbook():
    # 書き込みパスワードだけで保護したブックは、Excel が決まったパスワードで暗号化する：パスワードなしで開ける
    key = new_key("xls", "cryptoapi", DEFAULT_XLS_PASSWORD, salt=SALT, verifier=VERIFIER)
    enc = encrypt_binary(key, make_cfb({"Workbook": make_workbook(len(key.marker))}))
    plain, key2 = decrypt_binary(enc)
    assert key2.default_password and XlsWorkbook(plain).sheet("Sheet1").cells[(1, 1)].value == "とどく"
    assert encrypt_binary(key2, plain) == enc


def test_errors_are_clear():
    xls = make_cfb({"Workbook": make_workbook(6)})
    with pytest.raises(CryptoError, match="暗号化されていません"):
        decrypt_binary(xls, "x")
    with pytest.raises(CryptoError, match="複合ファイル"):
        decrypt_binary(b"PK\x03\x04 not a cfb", "x")
    assert not is_encrypted_binary(b"") and not is_encrypted_binary(b"%PDF-1.7")
    key = new_key("xls", "xor", "abc")
    with pytest.raises(CryptoError, match="印"):
        encrypt_binary(new_key("xls", "rc4", "abc", salt=SALT, verifier=VERIFIER), xls)  # 印の長さが違う
    with pytest.raises(CryptoError, match="WordDocument"):
        encrypt_binary(new_key("doc", "xor", "abc"), xls)  # 鍵とファイルの種類が違う
    assert isinstance(encrypt_binary(key, xls), bytes)
    with pytest.raises(CryptoError, match="対応していません"):
        bad = bytearray(make_workbook(6))
        bad[24:26] = b"\x05\x00"  # FilePass の種類 5
        bad[20:22] = b"\x2f\x00"
        decrypt_binary(make_cfb({"Workbook": bytes(bad)}), "x")
    assert repr(key) == "BinKey(xls, XOR による難読化)" and isinstance(key, BinKey)
