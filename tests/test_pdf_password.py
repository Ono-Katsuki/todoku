"""開くためのパスワードがかかったPDF（標準のセキュリティハンドラ R2〜R6）を読み書きするテスト。

テストの中で、仕様（PDF 32000-1 7.6.2〜7.6.3.4、ISO 32000-2 7.6.4.4 の Algorithm 2.B・8〜10）の手順どおりに
小さなPDFを暗号化して確かめる（集めたファイルがなくても、R2〜R6 のすべてを確かめられるように）。暗号化の手順は
ここに独立に書き、とどくの復号の実装は使わない（共通なのは、別に既知の値で確かめた RC4 と AES だけ）。
集めた公開のテスト用ファイル（corpus/crypto/pdf、パスワードの分かっているもの）があれば、それでも確かめる。
"""

from __future__ import annotations

import hashlib
import struct
from pathlib import Path

import pytest

from todoku.pdf import aes
from todoku.pdf.crypto import aes_cbc_decrypt, aes_cbc_encrypt, rc4, saslprep
from todoku.pdf.doc import NeedsPassword, PdfDocument, WrongPassword
from todoku.pdf.fill import Update, fill
from todoku.pdf.form import read_fields
from todoku.pdf.objects import Name, PdfError, PdfString, Ref, Stream, decode_text, dumps
from todoku.pdf.proof import prove
from todoku.pdf.window import build

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
FIX = CORPUS / "crypto" / "pdf"

PAD = bytes.fromhex("28BF4E5E4E758A4164004E56FFFA01082E2E00B6D0683E802F0CA9FE6453697A")
FILE_ID = bytes.fromhex("0123456789abcdeffedcba9876543210")
P = -3904  # 印刷・変更・写しを許さない値（上位のビットは 1）


# ====================================================================== テストのための暗号化（仕様の手順のとおり）
def _pad(pw: bytes) -> bytes:
    return (pw + PAD)[:32]


def _hash_2b(pw: bytes, salt: bytes, udata: bytes) -> bytes:
    """ISO 32000-2 Algorithm 2.B。"""
    k = hashlib.sha256(pw + salt + udata).digest()
    rnd = 0
    while True:
        e = aes.cbc_encrypt(k[:16], k[16:32], (pw + k + udata) * 64)
        k = {0: hashlib.sha256, 1: hashlib.sha384, 2: hashlib.sha512}[int.from_bytes(e[:16], "big") % 3](e).digest()
        rnd += 1
        if rnd >= 64 and e[-1] <= rnd - 32:
            return k[:32]


class Enc:
    """仕様の手順で、暗号の辞書と文書の鍵を作り、文字列とストリームを暗号化する。

    rev：2〜6。cfm：V2（RC4）・AESV2（R4）。strf：文字列の暗号フィルター（"Identity" で文字列は暗号化しない）。
    user・owner：パスワードのバイト列（R5・6 は SASLprep の後の UTF-8 を渡す）。"""

    def __init__(self, rev: int, user: bytes, owner: bytes, *, cfm: str = "V2", meta: bool = True, strf: str = "StdCF",
                 length: int = 128):
        self.rev, self.meta = rev, meta
        if rev <= 4:
            self.n = 5 if rev == 2 else length // 8
            self.stm = "aes" if cfm == "AESV2" else "rc4"
            self.str = "identity" if strf == "Identity" else self.stm
            self.O = self._o(owner or user, user)
            self.key = self._key(user)
            self.U = self._u()
            d = {"Filter": Name("Standard"), "V": 1 if rev == 2 else (2 if rev == 3 else 4), "R": rev,
                 "O": PdfString(self.O), "U": PdfString(self.U), "P": P}
            if rev >= 3:
                d["Length"] = length
            if rev == 4:
                d.update({"CF": {"StdCF": {"Type": Name("CryptFilter"), "CFM": Name(cfm), "Length": 16,
                                           "AuthEvent": Name("DocOpen")}},
                          "StmF": Name("StdCF"), "StrF": Name(strf)})
                if not meta:
                    d["EncryptMetadata"] = False
        else:
            self.stm = "aes256"
            self.str = "identity" if strf == "Identity" else "aes256"
            self.key = hashlib.sha256(b"todoku test file key").digest()
            h = (lambda pw, salt, u: hashlib.sha256(pw + salt + u).digest()) if rev == 5 else _hash_2b
            uvs, uks, ovs, oks = b"uvsalt01", b"uksalt02", b"ovsalt03", b"oksalt04"
            # Algorithm 8・9：U と UE、O と OE（O の計算には 48 バイトの U を使う）
            self.U = h(user, uvs, b"") + uvs + uks
            ue = aes.cbc_encrypt(h(user, uks, b""), bytes(16), self.key)
            self.O = h(owner, ovs, self.U) + ovs + oks
            oe = aes.cbc_encrypt(h(owner, oks, self.U), bytes(16), self.key)
            # Algorithm 10：Perms（P を 64 ビットに延ばし、T/F、adb、4 バイトの任意の値）
            perms = struct.pack("<q", P | ~0xFFFFFFFF) + (b"T" if meta else b"F") + b"adb" + b"todo"
            d = {"Filter": Name("Standard"), "V": 5, "R": rev, "Length": 256,
                 "CF": {"StdCF": {"CFM": Name("AESV3"), "Length": 32, "AuthEvent": Name("DocOpen")}},
                 "StmF": Name("StdCF"), "StrF": Name(strf), "O": PdfString(self.O), "U": PdfString(self.U),
                 "OE": PdfString(oe), "UE": PdfString(ue), "P": P,
                 "Perms": PdfString(aes.cbc_encrypt(self.key, bytes(16), perms))}
            if not meta:
                d["EncryptMetadata"] = False
        for k in ("O", "U", "OE", "UE", "Perms"):
            if k in d:
                d[k].hexform = True
        self.dict = d

    # Algorithm 3：O
    def _o(self, owner: bytes, user: bytes) -> bytes:
        d = hashlib.md5(_pad(owner)).digest()
        if self.rev >= 3:
            for _ in range(50):
                d = hashlib.md5(d).digest()
        k = d[:self.n]
        x = rc4(k, _pad(user))
        if self.rev >= 3:
            for i in range(1, 20):
                x = rc4(bytes(b ^ i for b in k), x)
        return x

    # Algorithm 2：文書の鍵
    def _key(self, user: bytes) -> bytes:
        h = hashlib.md5(_pad(user) + self.O + struct.pack("<i", P) + FILE_ID)
        if self.rev >= 4 and not self.meta:
            h.update(b"\xff\xff\xff\xff")
        d = h.digest()
        if self.rev >= 3:
            for _ in range(50):
                d = hashlib.md5(d[:self.n]).digest()
        return d[:self.n]

    # Algorithm 4・5：U
    def _u(self) -> bytes:
        if self.rev == 2:
            return rc4(self.key, PAD)
        x = rc4(self.key, hashlib.md5(PAD + FILE_ID).digest())
        for i in range(1, 20):
            x = rc4(bytes(b ^ i for b in self.key), x)
        return x + b"\x00" * 16

    # Algorithm 1（AES-256 は 1.A）
    def crypt(self, method: str, data: bytes, num: int, gen: int) -> bytes:
        if method == "identity":
            return data
        if method == "aes256":
            key = self.key
        else:
            salt = b"sAlT" if method == "aes" else b""
            key = hashlib.md5(self.key + struct.pack("<I", num)[:3] + struct.pack("<H", gen) + salt).digest()[:min(self.n + 5, 16)]
        if method == "rc4":
            return rc4(key, data)
        iv = hashlib.md5(b"test iv" + struct.pack("<II", num, gen) + data).digest()
        pad = 16 - len(data) % 16
        return iv + aes.cbc_encrypt(key, iv, data + bytes([pad]) * pad)

    def strings(self, o, num: int):
        if isinstance(o, PdfString):
            out = PdfString(self.crypt(self.str, bytes(o), num, 0))
            out.hexform = True
            return out
        if isinstance(o, list):
            return [self.strings(x, num) for x in o]
        if isinstance(o, dict):
            return {k: self.strings(v, num) for k, v in o.items()}
        return o


CONTENT = (b"BT /F1 14 Tf 72 760 Td (Application form) Tj ET\n"
           b"BT /F1 10 Tf 72 700 Td (Name) Tj ET\n0.5 w 140 690 m 400 690 l S\n/X1 Do\n")
XMP = b'<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"></x:xmpmeta><?xpacket end="r"?>'


def make_pdf(enc: Enc | None) -> bytes:
    """1ページの申請書（フォームの欄「name」＝Taro・情報の辞書・XMP・Identity の暗号フィルターの付いた描画）。"""
    t = lambda s: PdfString(s.encode())  # noqa: E731
    objs = {
        1: {"Type": Name("Catalog"), "Pages": Ref(2), "AcroForm": {"Fields": [Ref(6)], "DA": t("/Helv 0 Tf 0 g")},
            "Metadata": Ref(8)},
        2: {"Type": Name("Pages"), "Kids": [Ref(3)], "Count": 1},
        3: {"Type": Name("Page"), "Parent": Ref(2), "MediaBox": [0, 0, 595, 842], "Contents": Ref(4), "Annots": [Ref(6)],
            "Resources": {"Font": {"F1": Ref(5)}, "XObject": {"X1": Ref(9)}}},
        4: Stream({}, CONTENT),
        5: {"Type": Name("Font"), "Subtype": Name("Type1"), "BaseFont": Name("Helvetica"), "Encoding": Name("WinAnsiEncoding")},
        6: {"Type": Name("Annot"), "Subtype": Name("Widget"), "FT": Name("Tx"), "T": t("name"), "TU": t("Name"),
            "V": t("Taro"), "Rect": [150, 690, 400, 712], "P": Ref(3), "F": 4, "DA": t("/Helv 10 Tf 0 g")},
        7: {"Title": t("Encrypted form"), "Author": t("todoku")},
        8: Stream({"Type": Name("Metadata"), "Subtype": Name("XML")}, XMP),
        # 暗号フィルター Identity（7.6.5）：このストリームは、暗号化された文書の中でも暗号化しない
        9: Stream({"Type": Name("XObject"), "Subtype": Name("Form"), "BBox": [0, 0, 595, 842],
                   "Resources": {"Font": {"F1": Ref(5)}}, "Filter": [Name("Crypt")],
                   "DecodeParms": [{"Type": Name("CryptFilterDecodeParms"), "Name": Name("Identity")}]},
                  b"BT /F1 10 Tf 72 640 Td (Identity stream) Tj ET"),
    }
    if enc is not None and enc.rev < 4:
        # 暗号フィルターは V 4 から。R2・3 では、このストリームも普通に暗号化する
        objs[9].dict.pop("Filter")
        objs[9].dict.pop("DecodeParms")
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = {}
    for num, o in objs.items():
        offs[num] = len(out)
        if isinstance(o, Stream):
            data = o.raw
            d = dict(o.dict)
            if enc is not None:
                crypt = num == 9 and "Filter" in d
                meta = num == 8 and not enc.meta and enc.rev >= 4
                data = data if (crypt or meta) else enc.crypt(enc.stm, data, num, 0)
                d = enc.strings(d, num)
            d["Length"] = len(data)
            body = dumps(d) + b"\nstream\n" + data + b"\nendstream"
        else:
            body = dumps(enc.strings(o, num) if enc is not None else o)
        out += f"{num} 0 obj\n".encode() + body + b"\nendobj\n"
    trailer = {"Size": 11 if enc is not None else 10, "Root": Ref(1), "Info": Ref(7)}
    if enc is not None:
        offs[10] = len(out)
        out += b"10 0 obj\n" + dumps(enc.dict) + b"\nendobj\n"
        trailer["Encrypt"] = Ref(10)
    ids = [PdfString(FILE_ID), PdfString(FILE_ID[::-1])]
    for s in ids:
        s.hexform = True
    trailer["ID"] = ids
    xref = len(out)
    out += f"xref\n0 {trailer['Size']}\n0000000000 65535 f \n".encode()
    for num in range(1, trailer["Size"]):
        out += f"{offs[num]:010d} 00000 n \n".encode()
    out += b"trailer\n" + dumps(trailer) + f"\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(out)


USER, OWNER = "user-pass", "owner-pass"
CASES = {
    "R2-RC4-40": dict(rev=2),
    "R3-RC4-128": dict(rev=3),
    "R3-RC4-40": dict(rev=3, length=40),
    "R4-RC4": dict(rev=4),
    "R4-AES128": dict(rev=4, cfm="AESV2"),
    "R4-AES128-平文のメタデータ": dict(rev=4, cfm="AESV2", meta=False),
    "R4-AES128-文字列はIdentity": dict(rev=4, cfm="AESV2", strf="Identity"),
    "R5-AES256": dict(rev=5),
    "R6-AES256": dict(rev=6),
    "R6-AES256-平文のメタデータ": dict(rev=6, meta=False),
}
_BLOBS: dict[str, bytes] = {}


def encrypted(case: str) -> bytes:
    if case not in _BLOBS:
        _BLOBS[case] = make_pdf(Enc(user=USER.encode(), owner=OWNER.encode(), **CASES[case]))
    return _BLOBS[case]


def page_text(doc: PdfDocument) -> list[str]:
    return build(doc)["pages"][0]["_plain"]


# ====================================================================== 部品
def test_aes_known_answers():
    # FIPS 197 附属書 C.1〜C.3
    pt = bytes.fromhex("00112233445566778899aabbccddeeff")
    for n, want in ((16, "69c4e0d86a7b0430d8cdb78070b4c55a"), (24, "dda97ca4864cdfe06eaf70a0ec0d7191"),
                    (32, "8ea2b7ca516745bfeafc49904b496089")):
        k = bytes(range(n))
        assert aes.ecb_encrypt_block(k, pt).hex() == want
        assert aes.ecb_decrypt_block(k, bytes.fromhex(want)) == pt
    # 表を引く実装と、1バイトずつの実装（crypto.py）が同じ結果になる
    data = hashlib.sha512(b"todoku").digest() * 5
    for n in (16, 24, 32):
        k, iv = hashlib.sha256(bytes([n])).digest()[:n], hashlib.md5(bytes([n])).digest()
        c = aes.cbc_encrypt(k, iv, data)
        assert c == aes_cbc_encrypt(k, data, iv)
        assert aes.cbc_decrypt(k, iv, c) == data == aes_cbc_decrypt(k, c, iv, unpad=False)


@pytest.mark.parametrize("given, want", [("I­X", "IX"), ("user", "user"), ("USER", "USER"), ("ª", "a"),
                                         ("Ⅸ", "IX"), ("ｐａｓｓ　ワード", "pass ワード")])
def test_saslprep(given, want):
    # RFC 4013 3章の例（最後は、全角の英字と全角の空白：NFKC と、ASCII でない空白の対応づけ）
    assert saslprep(given) == want


@pytest.mark.parametrize("given", ["\u0007", "ا1"])
def test_saslprep_prohibited(given):
    with pytest.raises(ValueError):
        saslprep(given)


# ====================================================================== 仕様の手順で暗号化した小さなPDF
@pytest.mark.parametrize("case", list(CASES))
def test_open_with_user_and_owner_password(case):
    plain = PdfDocument(make_pdf(None))
    blob = encrypted(case)
    for pw, auth in ((USER, "user"), (OWNER, "owner")):
        doc = PdfDocument(blob, "e.pdf", password=pw)
        assert doc.decryptor.auth == auth
        assert page_text(doc) == page_text(plain) == ["Application form", "Name", "Identity stream"]
        assert {f.name: f.value_text() for f in read_fields(doc)} == {"name": "Taro"}
        assert decode_text(doc.info()["Title"]) == "Encrypted form"
        meta = doc.resolve(doc.root["Metadata"])
        assert doc.decode(meta) == XMP
        if doc.decryptor.r >= 5:
            assert doc.decryptor.perms_ok is True


@pytest.mark.parametrize("case", list(CASES))
def test_missing_or_wrong_password(case):
    blob = encrypted(case)
    for pw in (None, ""):
        with pytest.raises(NeedsPassword) as e:
            PdfDocument(blob, "e.pdf", password=pw)
        assert not isinstance(e.value, WrongPassword)
    with pytest.raises(WrongPassword) as e:
        PdfDocument(blob, "e.pdf", password="not-the-password")
    # パスワードの例外は、PDF の例外の一種（どちらも「パスワードが要る」として受けられる）
    assert isinstance(e.value, NeedsPassword) and isinstance(e.value, PdfError)
    assert "not-the-password" not in str(e.value)


def test_password_is_not_kept():
    secret = "kept-nowhere-7731"
    blob = make_pdf(Enc(6, secret.encode(), b"owner-x"))
    doc = PdfDocument(blob, "e.pdf", password=secret)

    def seen(o, depth=0):
        if depth > 4:
            return False
        if isinstance(o, (str, bytes)):
            return (secret in o) if isinstance(o, str) else (secret.encode() in o)
        if isinstance(o, dict):
            return any(seen(v, depth + 1) for v in o.values())
        if isinstance(o, (list, tuple)):
            return any(seen(v, depth + 1) for v in o)
        return seen(vars(o), depth + 1) if hasattr(o, "__dict__") else False

    assert not seen(vars(doc.decryptor)) and not any(seen(v) for k, v in vars(doc).items() if k != "blob")
    # 違うパスワードの例外：とどくの中の関数の変数（トレースバック）にも、渡したパスワードが残らない
    wrong = "never-in-traceback-42"
    with pytest.raises(WrongPassword) as e:
        PdfDocument(blob, "e.pdf", password=wrong)
    assert e.value.__context__ is None and e.value.__cause__ is None
    tb = e.value.__traceback__.tb_next  # テストの関数そのものの枠は除く
    while tb is not None:
        assert wrong not in repr(tb.tb_frame.f_locals)
        tb = tb.tb_next


def test_non_ascii_passwords():
    # R6：Acrobat は SASLprep で正規化してから UTF-8 にする。全角のまま入力しても開ける
    blob = make_pdf(Enc(6, saslprep("ｐａｓｓ　ワード").encode(), b"o"))
    assert PdfDocument(blob, password="ｐａｓｓ　ワード").decryptor.auth == "user"
    assert PdfDocument(blob, password="pass ワード").decryptor.auth == "user"
    # R4：PDFDocEncoding で表せる字はそのバイト。日本語は、作ったソフトの文字コード（シフト JIS・UTF-8）
    for raw, given in ((b"\xe6\xf8\xe5", "æøå"), ("申請".encode("cp932"), "申請"), ("申請".encode("utf-8"), "申請")):
        blob = make_pdf(Enc(4, raw, b"o", cfm="AESV2"))
        assert PdfDocument(blob, password=given).decryptor.auth == "user"
    # バイト列で渡したときは、そのまま使う
    blob = make_pdf(Enc(4, "申請".encode("cp932"), b"o"))
    assert PdfDocument(blob, password="申請".encode("cp932")).decryptor.auth == "user"
    with pytest.raises(WrongPassword):
        PdfDocument(blob, password="申請".encode("euc_jp"))


def test_unencrypted_ignores_password():
    doc = PdfDocument(make_pdf(None), password="anything")
    assert doc.decryptor is None and page_text(doc)[0] == "Application form"


@pytest.mark.parametrize("case", list(CASES))
def test_fill_and_prove_encrypted(case):
    blob = encrypted(case)
    doc = PdfDocument(blob, "e.pdf", password=USER)
    res = fill(doc, {"name": "山田 太郎"})
    assert res.blob[:len(blob)] == blob
    pr = prove(blob, res.blob, {"name"}, password=USER)
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    for pw in (USER, OWNER):
        after = PdfDocument(res.blob, "e.pdf", password=pw)
        assert {f.name: f.value_text() for f in read_fields(after)} == {"name": "山田 太郎"}
        assert page_text(after) == ["Application form", "Name", "Identity stream"]
    with pytest.raises(NeedsPassword):
        PdfDocument(res.blob, "e.pdf")
    tail = res.blob[len(blob):]
    if CASES[case].get("strf") != "Identity":
        # 書き足した所に、書いた字（UTF-16）も見た目の描画の命令も、平文では出てこない
        assert "山田".encode("utf-16-be") not in tail and b"/Tx BMC" not in tail


def test_prove_needs_the_password():
    # 証明は、パスワードがなければ開けず、違えば断る（黙って通さない）
    blob = encrypted("R4-AES128")
    res = fill(PdfDocument(blob, password=OWNER), {"name": "x"})
    assert prove(blob, res.blob, {"name"}, password=OWNER).ok
    with pytest.raises(NeedsPassword):
        prove(blob, res.blob, {"name"})
    with pytest.raises(WrongPassword):
        prove(blob, res.blob, {"name"}, password="nope")


def test_prove_compares_decrypted_values():
    # 暗号化された文書でも、証明は復号した値で比べる：書くと言っていない所（欄の見出し TU）の変更を見つける
    blob = encrypted("R4-AES128")
    doc = PdfDocument(blob, "e.pdf", password=USER)
    up = Update(doc)
    fd = up.get(Ref(6))
    fd["V"] = PdfString(b"Hanako")
    fd["TU"] = PdfString(b"Changed label")
    new = up.build()
    pr = prove(blob, new, {"name"}, password=USER)
    assert not pr.ok and any("TU" in u for u in pr.unexpected)
    assert {f.name: f.value_text() for f in read_fields(PdfDocument(new, password=USER))} == {"name": "Hanako"}


def test_unsupported_handler_is_not_a_password_error():
    # 公開鍵の暗号など、パスワードの暗号でないもの：パスワードの例外ではなく「対応していない」と断る
    blob = encrypted("R4-RC4").replace(b"/Filter /Standard", b"/Filter /PubSec01", 1)
    with pytest.raises(PdfError) as e:
        PdfDocument(blob, password=USER)
    assert not isinstance(e.value, NeedsPassword)


# ====================================================================== 公開のテスト用ファイル（パスワードが分かっているもの）
needs_fix = pytest.mark.skipif(not FIX.exists(), reason="corpus/crypto/pdf がありません（SOURCES.txt を参照）")

# pypdf の resources/encryption（qpdf で unencrypted.pdf を暗号化したもの。パスワードは pypdf の tests/test_encryption.py）
PYPDF = {
    "r2-empty-password.pdf": (None, "empty"), "r2-user-password.pdf": ("asdfzxcv", "user"),
    "r2-owner-password.pdf": ("asdfzxcv", "empty"), "r3-empty-password.pdf": (None, "empty"),
    "r3-user-password.pdf": ("asdfzxcv", "user"), "r4-user-password.pdf": ("asdfzxcv", "user"),
    "r4-owner-password.pdf": ("asdfzxcv", "empty"), "r4-aes-user-password.pdf": ("asdfzxcv", "user"),
    "r5-empty-password.pdf": (None, "empty"), "r5-user-password.pdf": ("asdfzxcv", "user"),
    "r5-owner-password.pdf": ("asdfzxcv", "empty"), "r6-empty-password.pdf": (None, "empty"),
    "r6-user-password.pdf": ("asdfzxcv", "user"), "r6-owner-password.pdf": ("asdfzxcv", "empty"),
    "r6-both-passwords.pdf": ("bar", "owner"),
}
# pypdf のテストが確かめている、復号した情報の辞書（空の値は除く）
PYPDF_INFO = {"Author": "cheng", "CreationDate": "D:20220414132421+05'24'", "Creator": "WPS Writer",
              "ModDate": "D:20220414132421+05'24'", "SourceModified": "D:20220414132421+05'24'", "Trapped": "/False"}


def _info(doc: PdfDocument) -> dict:
    out = {}
    for k, v in doc.info().items():
        v = decode_text(v) if isinstance(v, bytes) else (f"/{v.name}" if isinstance(v, Name) else v)
        if v:
            out[k] = v
    return out


@pytest.mark.corpus
@needs_fix
@pytest.mark.parametrize("name", list(PYPDF))
def test_pypdf_fixtures(name):
    pw, auth = PYPDF[name]
    plain = build(PdfDocument((FIX / "pypdf" / "unencrypted.pdf").read_bytes()))["pages"]
    blob = (FIX / "pypdf" / name).read_bytes()
    doc = PdfDocument(blob, name, password=pw)
    assert doc.decryptor.auth == auth
    assert _info(doc) == PYPDF_INFO
    # 復号した文書の窓（字・欄・配置）が、暗号化する前のファイルの窓と同じ
    assert build(doc)["pages"] == plain
    assert plain[0]["_plain"] == ["pdf encryption test", "password: asdfzxcv"]
    if auth != "empty":
        with pytest.raises(NeedsPassword):
            PdfDocument(blob, name)
        with pytest.raises(WrongPassword):
            PdfDocument(blob, name, password="asdfzxcV")


@pytest.mark.corpus
@needs_fix
def test_pypdf_aes128_without_key_length():
    # 暗号の辞書に Length がない AES-128：鍵の長さは暗号フィルターから決まる（128 ビット）
    plain = build(PdfDocument((FIX / "pypdf" / "unencrypted.pdf").read_bytes()))["pages"]
    doc = PdfDocument((FIX / "pypdf" / "r4-aes-v2-no-key-length.pdf").read_bytes())
    assert (doc.decryptor.auth, doc.decryptor.length) == ("empty", 16)
    assert build(doc)["pages"] == plain


@pytest.mark.corpus
@needs_fix
@pytest.mark.parametrize("name, pw", [("r2-owner-password.pdf", b"asdfzxcv"), ("r4-owner-password.pdf", b"asdfzxcv"),
                                      ("r5-owner-password.pdf", b"asdfzxcv"), ("r6-owner-password.pdf", b"asdfzxcv"),
                                      ("r3-user-password.pdf", b"asdfzxcv"), ("r4-aes-user-password.pdf", b"asdfzxcv"),
                                      ("r5-user-password.pdf", b""), ("r6-user-password.pdf", b"")])
def test_pypdf_owner_password_algorithm(name, pw):
    # 権限のパスワード（R2〜4 は Algorithm 7、R5・6 は 2.A の c・d）から求めた鍵が、開くためのパスワードの鍵と同じ。
    # qpdf は、権限のパスワードを空にすると、R2〜4 では開くためのパスワードで O を作り（Algorithm 3 の a）、R5・6 では空のまま
    doc = PdfDocument((FIX / "pypdf" / name).read_bytes(), name, password="asdfzxcv")
    assert doc.decryptor._owner(pw) == doc.decryptor.key


@pytest.mark.corpus
@needs_fix
def test_pypdf_both_passwords_and_cleartext_metadata():
    for name in ("r6-both-passwords.pdf", "r6-cleartext-metadata.pdf"):
        blob = (FIX / "pypdf" / name).read_bytes()
        a, b = PdfDocument(blob, password="foo"), PdfDocument(blob, password="bar")
        assert (a.decryptor.auth, b.decryptor.auth) == ("user", "owner")
        assert a.decryptor.key == b.decryptor.key and build(a)["pages"] == build(b)["pages"]
    # EncryptMetadata false：文書のメタデータは暗号化されていない（pypdf のテストと同じく、XMP の始まり）
    assert b.decryptor.meta is False and b.decode(b.resolve(b.root["Metadata"])).startswith(b"<?xpacket")
    crazy = build(PdfDocument((FIX / "pypdf" / "crazyones.pdf").read_bytes()))["pages"]
    assert build(b)["pages"][0]["_plain"] == crazy[0]["_plain"]
    enc256 = PdfDocument((FIX / "pypdf" / "crazyones-encrypted-256.pdf").read_bytes(), password="password")
    assert build(enc256)["pages"] == crazy and crazy[0]["_plain"][0] == "The Crazy Ones"


# pdf.js の test/pdfs（パスワードは test/test_manifest.json）。字は、開いて読んだもの（意味の通る字であることを確かめた）
PDFJS = {
    "issue15893_reduced.pdf": ("test", 3, ["Issue 15893 - password"]),
    "issue6010_1.pdf": ("abc", 6, ["Issue 6010"]),
    "issue6010_2.pdf": ("æøå", 6, ["Issue 6010"]),
    "issue21579.pdf": ("pässwört", 5, ["repro1a: AES-256 R5, password ’passwoert’ with umlauts",
                                       "If you can read this, decryption succeeded."]),
    "saslprep-r6.pdf": ("SªSL\xadprep", 6, ["Hello pdf.js world"]),
}


@pytest.mark.corpus
@needs_fix
@pytest.mark.parametrize("name", list(PDFJS))
def test_pdfjs_fixtures(name):
    pw, rev, text = PDFJS[name]
    blob = (FIX / "pdfjs" / name).read_bytes()
    doc = PdfDocument(blob, name, password=pw)
    assert doc.decryptor.r == rev and doc.decryptor.auth == "user"
    assert build(doc)["pages"][0]["_plain"] == text
    with pytest.raises(NeedsPassword):
        PdfDocument(blob, name)


@pytest.mark.corpus
@needs_fix
def test_pdfjs_saslprep_is_needed():
    # 正規化していない UTF-8 のバイト列そのものでは開けない（SASLprep で「SaSLprep」にしたものだけが合う）
    blob = (FIX / "pdfjs" / "saslprep-r6.pdf").read_bytes()
    with pytest.raises(WrongPassword):
        PdfDocument(blob, password="SªSL\xadprep".encode("utf-8"))
    assert PdfDocument(blob, password=b"SaSLprep").decryptor.auth == "user"


@pytest.mark.corpus
@needs_fix
def test_pdfjs_aes128_scans_and_identity_strings():
    # Acrobat で暗号化した AES-128 のスキャン：メタデータの XML と、JPEG 2000 の画像の署名が、復号して出てくる
    doc = PdfDocument((FIX / "pdfjs" / "issue3371.pdf").read_bytes(), password="ELXRTQWS")
    assert decode_text(doc.info()["Creator"]) == "Kyocera FS-C2126MFP+"
    assert doc.decode(doc.resolve(doc.root["Metadata"])).startswith(b"<?xpacket")
    xo = doc.resolve(doc.resolve(doc.pages()[0]["Resources"])["XObject"])
    assert doc.decode(doc.resolve(xo["Im1"])).startswith(b"\x00\x00\x00\x0cjP  \r\n\x87\n")
    # 文字列は暗号化しない（StrF Identity）・ストリームは AES-128・メタデータは平文
    doc = PdfDocument((FIX / "pdfjs" / "bug1782186.pdf").read_bytes(), password="Hello")
    assert (doc.decryptor.string_method, doc.decryptor.stream_method, doc.decryptor.meta) == ("identity", "aesv2", False)
    assert b"/OC /Layer0 BDC" in doc.page_content(doc.pages()[0])


@pytest.mark.corpus
@needs_fix
@pytest.mark.parametrize("name, pw", [("r2-user-password.pdf", "asdfzxcv"), ("r4-user-password.pdf", "asdfzxcv"),
                                      ("r4-aes-user-password.pdf", "asdfzxcv"), ("r6-user-password.pdf", "asdfzxcv"),
                                      ("r6-both-passwords.pdf", "bar")])
def test_fill_and_prove_pypdf_fixtures(name, pw):
    blob = (FIX / "pypdf" / name).read_bytes()
    spec = {"after": "asdfzxcv", "value": "とどく ABC 123", "where": "below"}
    res = fill(PdfDocument(blob, name, password=pw), {"@1": spec})
    pr = prove(blob, res.blob, {"p1.w1"}, {"p1.w1": spec}, password=pw)
    assert pr.ok and pr.prefix_ok, (pr.unexpected, pr.lint_errors)
    after = PdfDocument(res.blob, name, password=pw)
    assert {f.name: f.value_text() for f in read_fields(after)} == {"p1.w1": "とどく ABC 123"}
    assert "とどく".encode("utf-16-be") not in res.blob[len(blob):] and b"ABC 123" not in res.blob[len(blob):]
    with pytest.raises(NeedsPassword):
        PdfDocument(res.blob, name)
