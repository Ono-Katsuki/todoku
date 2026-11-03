"""パスワードで暗号化された .xlsx・.docx（MS-OFFCRYPTO 2.3.4）を解き、同じパスワードのまま暗号化し直す。

正解：Office が暗号化したファイル（corpus/crypto、パスワード Password1234_）を解くと、Office が保存した元のファイルとバイト単位で一致する。
corpus がないときは、仕様の生成の手順（2.3.4.7・2.3.4.13・2.3.4.14）どおりにテストの中で作った入れ物で確かめる。
作る側は pdf/crypto.py の AES（officecrypto が使う aes.py とは別の実装）を使う。
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import io
import random
import re
import struct
import zipfile
from pathlib import Path

import pytest

from todoku.aes import Aes
from todoku.cfb import CFB, ENDOFCHAIN, FATSECT, FREESECT, MAGIC, NOSTREAM
from todoku.officecrypto import (CryptoError, WrongPassword, decrypt_ooxml, encrypt_ooxml, is_encrypted_ooxml,
                                 verify_integrity)
from todoku.pdf.crypto import aes_cbc_decrypt, aes_cbc_encrypt, aes_ecb_encrypt

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
CRYPTO = CORPUS / "crypto"
PASSWORD = "Password1234_"
FIXTURES = [("example_password.docx", "example.docx", "agile"), ("example_password.xlsx", "example.xlsx", "agile"),
            ("ecma376standard_password.docx", "ecma376standard_password_plain.docx", "standard")]


def fixture(name: str, out: str) -> tuple[bytes, bytes]:
    a, b = CRYPTO / "inputs" / name, CRYPTO / "outputs" / out
    if not (a.exists() and b.exists()):
        pytest.skip("corpus/crypto がありません")
    return a.read_bytes(), b.read_bytes()


# ------------------------------------------------------------------ AES
def test_aes_fips197_known_answers():
    plain = bytes.fromhex("00112233445566778899aabbccddeeff")
    for key, want in [("000102030405060708090a0b0c0d0e0f", "69c4e0d86a7b0430d8cdb78070b4c55a"),
                      ("000102030405060708090a0b0c0d0e0f1011121314151617", "dda97ca4864cdfe06eaf70a0ec0d7191"),
                      ("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f", "8ea2b7ca516745bfeafc49904b496089")]:
        a = Aes(bytes.fromhex(key))
        assert a.encrypt_ecb(plain).hex() == want
        assert a.decrypt_ecb(bytes.fromhex(want)) == plain


def test_aes_cbc_sp800_38a():
    key, iv = bytes.fromhex("2b7e151628aed2a6abf7158809cf4f3c"), bytes.fromhex("000102030405060708090a0b0c0d0e0f")
    plain = bytes.fromhex("6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e51"
                          "30c81c46a35ce411e5fbc1191a0a52eff69f2445df4f9b17ad2b417be66c3710")
    cipher = bytes.fromhex("7649abac8119b246cee98e9b12e9197d5086cb9b507219ee95db113a917678b2"
                           "73bed6b8e3c1743b7116e69e222295163ff1caa1681fac09120eca307586e1a7")
    assert Aes(key).encrypt_cbc(plain, iv) == cipher
    assert Aes(key).decrypt_cbc(cipher, iv) == plain


def test_aes_agrees_with_pdf_crypto():
    rnd = random.Random(7)
    for klen in (16, 24, 32):
        for n in (0, 16, 64, 4096):
            key, iv, data = rnd.randbytes(klen), rnd.randbytes(16), rnd.randbytes(n)
            a = Aes(key)
            assert a.encrypt_cbc(data, iv) == aes_cbc_encrypt(key, data, iv)
            assert a.decrypt_cbc(data, iv) == aes_cbc_decrypt(key, data, iv, unpad=False)
            assert a.encrypt_ecb(data) == aes_ecb_encrypt(key, data)
            assert a.decrypt_ecb(a.encrypt_ecb(data)) == data
    with pytest.raises(ValueError):
        Aes(b"\x00" * 16).encrypt_ecb(b"\x00" * 15)


# ------------------------------------------------------------------ テストの中で作る複合ファイル
def build_cfb(streams: dict[str, bytes]) -> bytes:
    """版3（512 バイトのセクタ）の小さな複合ファイル。ストリームは根の直下、4096 バイト未満はミニストリームに置く。
    並び：[FAT][ディレクトリ][ミニ FAT][ミニストリーム][大きいストリーム]。"""
    ss, ms = 512, 64
    names = sorted(streams, key=lambda n: (len(n), n.upper()))
    mini, minifat, start = bytearray(), [], {}
    for n in names:
        d = streams[n]
        if len(d) < 4096:
            k = -(-len(d) // ms)
            first = len(mini) // ms
            start[n] = first if k else ENDOFCHAIN
            minifat += list(range(first + 1, first + k)) + [ENDOFCHAIN] if k else []
            mini += d.ljust(k * ms, b"\x00")
    big = [n for n in names if len(streams[n]) >= 4096]
    nsec = lambda b: -(-b // ss)  # noqa: E731
    sizes = [nsec(128 * (len(names) + 1)), nsec(4 * len(minifat)), nsec(len(mini))] + [nsec(len(streams[n])) for n in big]
    n_fat = 1
    while n_fat * (ss // 4) < n_fat + sum(sizes):
        n_fat += 1
    fat, pos, firsts = [FATSECT] * n_fat, n_fat, []
    for k in sizes:
        firsts.append(pos if k else ENDOFCHAIN)
        fat += list(range(pos + 1, pos + k)) + [ENDOFCHAIN] if k else []
        pos += k
    fat += [FREESECT] * (n_fat * (ss // 4) - len(fat))
    dir_start, minifat_start, mini_start = firsts[:3]
    start.update(zip(big, firsts[3:]))

    def entry(name: str, kind: int, right: int, child: int, first: int, size: int) -> bytes:
        raw = name.encode("utf-16-le")
        return (raw.ljust(64, b"\x00") + struct.pack("<HBB3I", len(raw) + 2 if name else 0, kind, 1, NOSTREAM, right, child)
                + bytes(36) + struct.pack("<IQ", first, size))

    dirs = entry("Root Entry", 5, NOSTREAM, 1 if names else NOSTREAM, mini_start, len(mini))
    for i, n in enumerate(names):
        dirs += entry(n, 2, i + 2 if i + 1 < len(names) else NOSTREAM, NOSTREAM, start[n], len(streams[n]))
    while len(dirs) % ss:
        dirs += entry("", 0, NOSTREAM, NOSTREAM, 0, 0)
    header = (MAGIC + bytes(16) + struct.pack("<HHHHH6x", 0x3E, 3, 0xFFFE, 9, 6)
              + struct.pack("<9I", 0, n_fat, dir_start, 0, 4096, minifat_start, sizes[1], ENDOFCHAIN, 0)
              + struct.pack("<109I", *(list(range(n_fat)) + [FREESECT] * (109 - n_fat))))
    minifat += [FREESECT] * (sizes[1] * ss // 4 - len(minifat))
    body = struct.pack(f"<{len(fat)}I", *fat) + dirs + struct.pack(f"<{len(minifat)}I", *minifat)
    body += bytes(mini).ljust(sizes[2] * ss, b"\x00")
    for n, k in zip(big, sizes[3:]):
        body += streams[n].ljust(k * ss, b"\x00")
    return header + body


def assert_well_formed(blob: bytes) -> None:
    """複合ファイルのつながりが矛盾しないこと：どのセクタも持ち主は1つ、鎖の長さは大きさどおり、持ち主のないセクタは空き。"""
    c = CFB(blob)
    assert len(blob) % c.ssize == 0
    n = len(blob) // c.ssize - 1
    owner: dict[int, str] = {}

    def claim(sectors: list[int], who: str, size: int | None = None) -> None:
        if size is not None:
            assert len(sectors) == -(-size // c.ssize), who
        for s in sectors:
            assert s < n and s not in owner, (who, s, owner.get(s))
            owner[s] = who

    claim(c.fat_sectors, "FAT")
    assert all(c.fat[s] == FATSECT for s in c.fat_sectors)
    claim(c.chain_difat(), "DIFAT")
    claim(c.chain(c.dir_start), "dir")
    if c.minifat:
        claim(c.chain(c.minifat_start), "minifat")
    root = c.entries[0]
    claim(c.chain(root.start) if root.start != ENDOFCHAIN else [], "ministream", root.size)
    used_mini: set[int] = set()
    for path, e in c.streams().items():
        if e.size >= c.cutoff:
            claim(c.chain(e.start), path, e.size)
        else:
            chain, s = [], e.start
            while s not in (ENDOFCHAIN, FREESECT) and len(chain) <= len(c.minifat):
                chain.append(s)
                s = c.minifat[s]
            assert len(chain) == -(-e.size // c.msize) and not used_mini & set(chain), path
            assert all(m < root.size // c.msize for m in chain), path
            used_mini |= set(chain)
    assert all(c.fat[s] == FREESECT for s in range(len(c.fat)) if s not in owner)


def test_build_cfb_reads_back():
    rnd = random.Random(1)
    streams = {"A": rnd.randbytes(100), "Big": rnd.randbytes(5000), "Empty": b"", "Z": rnd.randbytes(4095)}
    blob = build_cfb(streams)
    assert_well_formed(blob)
    c = CFB(blob)
    assert {p: c.read(p) for p in c.streams()} == streams


def test_cfb_replace_mini_same_size_any_length():
    """ミニストリームの、64 の倍数でない長さのストリームを同じ大きさで差し替えても、ほかのバイトは変わらない。"""
    rnd = random.Random(2)
    streams = {"A": rnd.randbytes(1289), "B": rnd.randbytes(200), "C": rnd.randbytes(9000)}
    blob = build_cfb(streams)
    new = rnd.randbytes(1289)
    out = CFB(blob).replace("A", new)
    assert CFB(out).read("A") == new and CFB(out).read("B") == streams["B"] and CFB(out).read("C") == streams["C"]
    diff = [i for i in range(len(blob)) if blob[i] != out[i]]
    assert len(out) == len(blob) and len(diff) <= 1289
    assert CFB(blob).replace("A", streams["A"]) == blob


def test_cfb_replace_wipe_zeroes_old_data():
    rnd = random.Random(3)
    old = bytes(range(1, 256)) * 40  # 0 を含まない 10200 バイト
    blob = build_cfb({"S": old, "M": b"\x01" * 3000})
    out = CFB(blob).replace("S", b"\x02" * 4200, wipe=True)
    assert_well_formed(out)
    assert CFB(out).read("S") == b"\x02" * 4200 and CFB(out).read("M") == b"\x01" * 3000
    assert old[4200:4300] not in out and old[:100] not in out  # 縮んで使わなくなった元のデータは残らない
    assert old[6000:6100] in CFB(blob).replace("S", b"\x02" * 4200)  # wipe しなければ残る（これまでどおり）
    new_m = rnd.randbytes(5000)
    out = CFB(blob).replace("M", new_m, wipe=True)  # ミニストリームからふつうのセクタへ
    assert_well_formed(out)
    assert CFB(out).read("M") == new_m and CFB(out).read("S") == old
    assert b"\x01" * 64 not in out


# ------------------------------------------------------------------ 仕様の手順どおりに暗号化した入れ物を、テストの中で作る
BK = {"input": "fea7d2763b4b9e79", "hash": "d7aa0f6d3061344e", "key": "146e0be7abacd0d6",
      "hmac_key": "5fb2ad010cb9e1f6", "hmac_value": "a0677f02b22c8433"}


def _pad(b: bytes, rnd: random.Random | None = None) -> bytes:
    k = -len(b) % 16
    return b + (rnd.randbytes(k) if rnd else bytes(k))


def agile_container(password: str, package: bytes, seed: int = 0, hash_name: str = "SHA512", bits: int = 256,
                    spin: int = 100000, hmac_key_len: int | None = None, extra: dict[str, bytes] | None = None) -> bytes:
    """2.3.4.10〜2.3.4.15 の作り方どおりの Agile の入れ物。詰め物は Office と同じく乱数。"""
    rnd = random.Random(seed)
    alg = {"SHA512": "sha512", "SHA1": "sha1", "SHA256": "sha256", "SHA384": "sha384", "SHA-1": "sha1"}[hash_name]
    H = lambda b: hashlib.new(alg, b).digest()  # noqa: E731
    hs = hashlib.new(alg).digest_size
    key_salt, pw_salt, inter, verifier = rnd.randbytes(16), rnd.randbytes(16), rnd.randbytes(bits // 8), rnd.randbytes(16)
    hmac_key = rnd.randbytes(hmac_key_len or hs)
    h = H(pw_salt + password.encode("utf-16-le"))
    for i in range(spin):
        h = H(struct.pack("<I", i) + h)
    pk = lambda bk: (H(h + bytes.fromhex(BK[bk])) + b"\x36" * 64)[:bits // 8]  # noqa: E731
    iv = lambda bk: H(key_salt + bk)[:16]  # noqa: E731
    body = b"".join(aes_cbc_encrypt(inter, _pad(package[i:i + 4096], rnd), iv(struct.pack("<I", i // 4096)))
                    for i in range(0, len(package), 4096))
    stream = struct.pack("<Q", len(package)) + body
    mac = hmac.new(hmac_key, stream, alg).digest()
    b64 = lambda b: base64.b64encode(b).decode()  # noqa: E731
    attrs = (f'saltSize="16" blockSize="16" keyBits="{bits}" hashSize="{hs}" cipherAlgorithm="AES" '
             f'cipherChaining="ChainingModeCBC" hashAlgorithm="{hash_name}"')
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" '
           'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
           f'<keyData {attrs} saltValue="{b64(key_salt)}"/>'
           f'<dataIntegrity encryptedHmacKey="{b64(aes_cbc_encrypt(inter, _pad(hmac_key, rnd), iv(bytes.fromhex(BK["hmac_key"]))))}" '
           f'encryptedHmacValue="{b64(aes_cbc_encrypt(inter, _pad(mac, rnd), iv(bytes.fromhex(BK["hmac_value"]))))}"/>'
           '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
           f'<p:encryptedKey spinCount="{spin}" {attrs} saltValue="{b64(pw_salt)}" '
           f'encryptedVerifierHashInput="{b64(aes_cbc_encrypt(pk("input"), verifier, pw_salt))}" '
           f'encryptedVerifierHashValue="{b64(aes_cbc_encrypt(pk("hash"), _pad(H(verifier)), pw_salt))}" '
           f'encryptedKeyValue="{b64(aes_cbc_encrypt(pk("key"), _pad(inter), pw_salt))}"/>'
           '</keyEncryptor></keyEncryptors></encryption>')
    info = struct.pack("<HHI", 4, 4, 0x40) + xml.encode()
    return build_cfb({"EncryptionInfo": info, "EncryptedPackage": stream, **(extra or {})})


def standard_key(password: str, salt: bytes, bits: int) -> bytes:
    """2.3.4.7。"""
    h = hashlib.sha1(salt + password.encode("utf-16-le")).digest()
    for i in range(50000):
        h = hashlib.sha1(struct.pack("<I", i) + h).digest()
    h = hashlib.sha1(h + struct.pack("<I", 0)).digest()
    x1 = hashlib.sha1(bytes(a ^ b for a, b in zip(h.ljust(64, b"\x00"), b"\x36" * 64))).digest()
    x2 = hashlib.sha1(bytes(a ^ b for a, b in zip(h.ljust(64, b"\x00"), b"\x5c" * 64))).digest()
    return (x1 + x2)[:bits // 8]


def standard_container(password: str, package: bytes, seed: int = 0, bits: int = 128, version: tuple[int, int] = (4, 2),
                       flags: int = 0x24, alg: int | None = None) -> bytes:
    """2.3.4.5〜2.3.4.8 の作り方どおりの Standard の入れ物。"""
    rnd = random.Random(seed)
    salt, verifier = rnd.randbytes(16), rnd.randbytes(16)
    key = standard_key(password, salt, bits)
    alg = {128: 0x660E, 192: 0x660F, 256: 0x6610}[bits] if alg is None else alg
    header = (struct.pack("<8I", flags, 0, alg, 0x8004, bits, 0x18, 0, 0)
              + "Microsoft Enhanced RSA and AES Cryptographic Provider\0".encode("utf-16-le"))
    info = (struct.pack("<HHII", *version, flags, len(header)) + header + struct.pack("<I", 16) + salt
            + aes_ecb_encrypt(key, verifier) + struct.pack("<I", 20) + aes_ecb_encrypt(key, _pad(hashlib.sha1(verifier).digest())))
    stream = struct.pack("<Q", len(package)) + aes_ecb_encrypt(key, _pad(package, rnd))
    return build_cfb({"EncryptionInfo": info, "EncryptedPackage": stream})


def fake_package(n: int, seed: int = 0) -> bytes:
    return (b"PK\x03\x04" + random.Random(seed).randbytes(n))[:n]


def assert_reencrypted(original: bytes, new: bytes, key, package: bytes) -> None:
    """暗号化し直したファイル：同じパスワードで解けて中身に戻り、改ざん検知が合い、ほかのストリームは元のまま、
    EncryptionInfo は encryptedHmacValue の値のほかは元のまま。"""
    assert_well_formed(new)
    back, _ = decrypt_ooxml(new, PASSWORD)
    assert back == package
    assert verify_integrity(new, key)
    a, b = CFB(original), CFB(new)
    assert set(a.streams()) == set(b.streams())
    for p in a.streams():
        if p not in ("EncryptedPackage", "EncryptionInfo"):
            assert a.read(p) == b.read(p), p
    old_info, new_info = a.read("EncryptionInfo"), b.read("EncryptionInfo")
    assert len(old_info) == len(new_info)
    diff = [i for i in range(len(old_info)) if old_info[i] != new_info[i]]
    if key.method == "standard" or not key.has_integrity:
        assert not diff
    elif diff:
        m = re.search(rb'encryptedHmacValue="([^"]*)"', old_info)
        assert m.start(1) <= diff[0] and diff[-1] < m.end(1)


# ------------------------------------------------------------------ 作った入れ物で（corpus なしでも）
@pytest.mark.parametrize("hash_name,bits", [("SHA512", 256), ("SHA1", 128), ("SHA256", 192), ("SHA384", 256), ("SHA-1", 256)])
def test_agile_selfmade(hash_name, bits):
    package = fake_package(9000, 1)
    blob = agile_container(PASSWORD, package, hash_name=hash_name, bits=bits, spin=1000,
                           extra={"\x05SummaryInformation": b"\x05" * 300})
    assert is_encrypted_ooxml(blob)
    got, key = decrypt_ooxml(blob, PASSWORD)
    assert got == package
    assert (key.method, key.cipher, key.spin_count) == ("agile", f"AES-{bits}", 1000)
    assert verify_integrity(blob, key) and len(key.hmac_key) == hashlib.new(key.digest).digest_size
    assert encrypt_ooxml(key, package) == blob  # 中身を変えなければ、元とバイト単位で同じ
    for n in (2, 15, 4080, 4088, 4096, 4097, 9000 - 700, 9000 + 5000):
        changed = fake_package(n, n)
        assert_reencrypted(blob, encrypt_ooxml(key, changed), key, changed)


def test_agile_office_parameters_selfmade():
    """Office の既定（AES-256・SHA512・10万回）。"""
    package = fake_package(20000, 2)
    blob = agile_container(PASSWORD, package)
    got, key = decrypt_ooxml(blob, PASSWORD)
    assert got == package and key.spin_count == 100000 and key.label == "パスワードで暗号化（Agile・AES-256・SHA512）"
    assert encrypt_ooxml(key, package) == blob


def test_agile_hmac_key_of_salt_size():
    """仕様の文（saltSize の長さの HMAC の鍵）どおりに作ったファイルも、合う方の長さを選んで読み書きできる。"""
    package = fake_package(5000, 3)
    blob = agile_container(PASSWORD, package, spin=10, hmac_key_len=16)
    got, key = decrypt_ooxml(blob, PASSWORD)
    assert got == package and len(key.hmac_key) == 16
    changed = fake_package(6000, 4)
    assert_reencrypted(blob, encrypt_ooxml(key, changed), key, changed)


@pytest.mark.parametrize("bits", [128, 192, 256])
def test_standard_selfmade(bits):
    package = fake_package(7000, bits)
    blob = standard_container(PASSWORD, package, bits=bits)
    got, key = decrypt_ooxml(blob, PASSWORD)
    assert got == package and (key.method, key.cipher, key.has_integrity) == ("standard", f"AES-{bits}", False)
    assert verify_integrity(blob, key)
    assert encrypt_ooxml(key, package) == blob
    for n in (10, 4000, 4090, 7000 + 3000):
        changed = fake_package(n, n)
        assert_reencrypted(blob, encrypt_ooxml(key, changed), key, changed)


def test_small_package_stays_in_mini_stream_or_moves_out():
    """4096 バイトより小さい EncryptedPackage：縮むときは元の長さのまま（ミニストリームに残る）、伸びるときはふつうのセクタへ。"""
    package = fake_package(3000, 5)
    blob = agile_container(PASSWORD, package, spin=10)
    _, key = decrypt_ooxml(blob, PASSWORD)
    shrunk = encrypt_ooxml(key, package[:2500])
    assert CFB(shrunk).streams()["EncryptedPackage"].size == key.stream_len == 3016
    assert len(shrunk) == len(blob)
    grown = encrypt_ooxml(key, package + b"x" * 200)
    assert CFB(grown).streams()["EncryptedPackage"].size == 4104
    for new, pkg in ((shrunk, package[:2500]), (grown, package + b"x" * 200)):
        assert_reencrypted(blob, new, key, pkg)
        assert package[:64] not in new  # 中身の平文は入れ物に出てこない


def test_wrong_password_selfmade():
    blob = agile_container(PASSWORD, fake_package(5000), spin=10)
    for pw in ("", "password1234_", "Password1234", "Password1234_ "):
        with pytest.raises(WrongPassword):
            decrypt_ooxml(blob, pw)
    with pytest.raises(WrongPassword):
        decrypt_ooxml(standard_container(PASSWORD, fake_package(5000)), "nope")


def test_japanese_password():
    pw = "申請書のパスワード🔑"
    package = fake_package(5000, 6)
    got, key = decrypt_ooxml(agile_container(pw, package, spin=10), pw)
    assert got == package
    got, key = decrypt_ooxml(standard_container(pw, package), pw)
    assert got == package


def test_tampered_package_is_refused():
    blob = bytearray(agile_container(PASSWORD, fake_package(9000), spin=10))
    _, key = decrypt_ooxml(bytes(blob), PASSWORD)
    c = CFB(bytes(blob))
    e = c.streams()["EncryptedPackage"]
    off = (c.chain(e.start)[3] + 1) * c.ssize + 5
    blob[off] ^= 1
    with pytest.raises(CryptoError, match="改ざん検知"):
        decrypt_ooxml(bytes(blob), PASSWORD)
    assert not verify_integrity(bytes(blob), key)


def test_verify_integrity_rejects_other_keys():
    blob = agile_container(PASSWORD, fake_package(5000), seed=1, spin=10)
    other = agile_container(PASSWORD, fake_package(5000), seed=2, spin=10)
    _, key = decrypt_ooxml(blob, PASSWORD)
    assert verify_integrity(blob, key) and not verify_integrity(other, key)
    assert not verify_integrity(b"not a file", key)


def test_password_is_not_kept_or_shown():
    pw = "Secret-Pw-42"
    blob = agile_container(pw, fake_package(5000), spin=10)
    _, key = decrypt_ooxml(blob, pw)
    assert pw not in repr(key) and pw not in key.label
    assert all(pw.encode(e) not in v for v in vars(key).values() if isinstance(v, bytes) for e in ("utf-16-le", "utf-8"))
    with pytest.raises(WrongPassword) as ei:
        decrypt_ooxml(blob, "Secret-Pw-43")
    assert "Secret-Pw" not in str(ei.value)


# ------------------------------------------------------------------ 対応していない形・暗号化されていないファイル
def test_not_encrypted_inputs():
    zipped = io.BytesIO()
    with zipfile.ZipFile(zipped, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
    for blob, message, encrypted in [(zipped.getvalue(), "暗号化されていない", False), (b"%PDF-1.7\n", "ではありません", False),
                                     (build_cfb({"Workbook": b"\x09\x08" * 100}), "EncryptedPackage がありません", False),
                                     (build_cfb({"EncryptedPackage": b"\x00" * 100, "DRM": b"x"}), "権利管理", True)]:
        assert is_encrypted_ooxml(blob) == encrypted
        with pytest.raises(CryptoError, match=message) as ei:
            decrypt_ooxml(blob, PASSWORD)
        assert not isinstance(ei.value, WrongPassword)
    assert not is_encrypted_ooxml(b"") and not is_encrypted_ooxml(MAGIC + b"\x00" * 10)
    with pytest.raises(CryptoError, match="読めません"):
        decrypt_ooxml(MAGIC + b"\x00" * 10, PASSWORD)


def test_unsupported_methods_are_named():
    package = fake_package(5000)
    with pytest.raises(CryptoError, match="拡張の暗号化"):
        decrypt_ooxml(standard_container(PASSWORD, package, version=(4, 3), flags=0x10, alg=0), PASSWORD)
    with pytest.raises(CryptoError, match="拡張の暗号化"):
        decrypt_ooxml(standard_container(PASSWORD, package, flags=0x10, alg=0), PASSWORD)
    with pytest.raises(CryptoError, match="RC4"):
        decrypt_ooxml(standard_container(PASSWORD, package, flags=0x04, alg=0x6801), PASSWORD)
    with pytest.raises(CryptoError, match="版 1.1"):
        decrypt_ooxml(standard_container(PASSWORD, package, version=(1, 1)), PASSWORD)
    agile = agile_container(PASSWORD, package, spin=10)
    info = CFB(agile).read("EncryptionInfo")
    for old, new, message in [(b'cipherAlgorithm="AES"', b'cipherAlgorithm="DES"', "暗号の方式"),
                              (b'"ChainingModeCBC"', b'"ChainingModeCFB"', "つなぎ方"),
                              (b'hashAlgorithm="SHA512"', b'hashAlgorithm="WHIRL"', "ハッシュの方式"),
                              (b'spinCount="10"', b'spinCount="99999999"', "spinCount"),
                              (b"keyEncryptor/password\">", b"keyEncryptor/certificate\">", "証明書")]:
        assert old in info
        blob = build_cfb({"EncryptionInfo": info.replace(old, new), "EncryptedPackage": CFB(agile).read("EncryptedPackage")})
        with pytest.raises(CryptoError, match=message):
            decrypt_ooxml(blob, PASSWORD)


def test_encrypt_refuses_non_zip():
    _, key = decrypt_ooxml(agile_container(PASSWORD, fake_package(5000), spin=10), PASSWORD)
    with pytest.raises(CryptoError, match="ZIP"):
        encrypt_ooxml(key, b"")
    with pytest.raises(CryptoError, match="ZIP"):
        encrypt_ooxml(key, b"\xd0\xcf\x11\xe0")


def test_decrypted_non_zip_is_refused():
    """Standard には改ざん検知がない：解いた中身が ZIP でなければ、壊れたファイルとして止める。"""
    with pytest.raises(CryptoError, match="ZIP") as ei:
        decrypt_ooxml(standard_container(PASSWORD, b"XX" + bytes(5000)), PASSWORD)
    assert not isinstance(ei.value, WrongPassword)


# ------------------------------------------------------------------ Office が暗号化したファイル（corpus）
@pytest.mark.corpus
@pytest.mark.parametrize("name,out,method", FIXTURES, ids=[f[0] for f in FIXTURES])
def test_office_file_decrypts_to_known_output(name, out, method):
    blob, want = fixture(name, out)
    assert is_encrypted_ooxml(blob)
    got, key = decrypt_ooxml(blob, PASSWORD)
    assert got == want
    assert key.method == method
    with pytest.raises(WrongPassword):
        decrypt_ooxml(blob, "Password1234")


@pytest.mark.corpus
@pytest.mark.parametrize("name,out,method", FIXTURES, ids=[f[0] for f in FIXTURES])
def test_office_file_reencrypts_identically(name, out, method):
    blob, want = fixture(name, out)
    _, key = decrypt_ooxml(blob, PASSWORD)
    assert verify_integrity(blob, key)
    if method == "agile":
        assert key.has_integrity and len(key.hmac_key) == 64  # Office は hashSize（SHA-512 の 64 バイト）の鍵を使う
    assert encrypt_ooxml(key, want) == blob


@pytest.mark.corpus
@pytest.mark.parametrize("name,out,method", FIXTURES, ids=[f[0] for f in FIXTURES])
def test_office_file_written_and_reencrypted(name, out, method):
    """中身の ZIP の1つの部品を書き換えて暗号化し直す：同じパスワードで開け、書き換えたとおりの中身になる。"""
    blob, want = fixture(name, out)
    _, key = decrypt_ooxml(blob, PASSWORD)
    src = zipfile.ZipFile(io.BytesIO(want))
    part = "word/document.xml" if "word/document.xml" in src.namelist() else "xl/sharedStrings.xml"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for info in src.infolist():
            data = src.read(info)
            z.writestr(info, data.replace(b"</", "記入した内容</".encode(), 1) if info.filename == part else data)
    for package in (buf.getvalue(), want + random.Random(1).randbytes(6000), want[:3500]):
        new = encrypt_ooxml(key, package)
        assert_reencrypted(blob, new, key, package)
    back, _ = decrypt_ooxml(encrypt_ooxml(key, buf.getvalue()), PASSWORD)
    assert "記入した内容".encode() in zipfile.ZipFile(io.BytesIO(back)).read(part)


@pytest.mark.corpus
def test_office_file_tamper_detected():
    blob, _ = fixture("example_password.xlsx", "example.xlsx")
    c = CFB(blob)
    e = c.streams()["EncryptedPackage"]
    bad = bytearray(blob)
    bad[(c.chain(e.start)[-1] + 1) * c.ssize] ^= 0x80
    with pytest.raises(CryptoError, match="改ざん検知"):
        decrypt_ooxml(bytes(bad), PASSWORD)
