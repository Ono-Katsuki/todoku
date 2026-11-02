"""PDFの暗号（標準のセキュリティハンドラ）を解く：RC4、AES-128、AES-256。外部のライブラリは使わない。

行政のPDFには「編集の制限」だけをかけたもの（開くパスワードは空）が多い。空のパスワードで開けるものは開く。
"""

from __future__ import annotations

import hashlib
import struct

PAD = bytes.fromhex("28BF4E5E4E758A4164004E56FFFA01082E2E00B6D0683E802F0CA9FE6453697A")


def rc4(key: bytes, data: bytes) -> bytes:
    s = list(range(256))
    j = 0
    for i in range(256):
        j = (j + s[i] + key[i % len(key)]) & 0xFF
        s[i], s[j] = s[j], s[i]
    out = bytearray(len(data))
    i = j = 0
    for k, b in enumerate(data):
        i = (i + 1) & 0xFF
        j = (j + s[i]) & 0xFF
        s[i], s[j] = s[j], s[i]
        out[k] = b ^ s[(s[i] + s[j]) & 0xFF]
    return bytes(out)


# ---------------------------------------------------------------------- AES（復号だけ）
_SBOX = [0] * 256
_INV = [0] * 256


def _init_sbox():
    p = q = 1
    while True:
        p = p ^ ((p << 1) & 0xFF) ^ (0x1B if p & 0x80 else 0)
        q ^= q << 1
        q ^= q << 2
        q ^= q << 4
        q &= 0xFF
        if q & 0x80:
            q ^= 0x09
        x = q ^ ((q << 1) | (q >> 7)) & 0xFF ^ ((q << 2) | (q >> 6)) & 0xFF ^ ((q << 3) | (q >> 5)) & 0xFF ^ ((q << 4) | (q >> 4)) & 0xFF
        x = (x ^ 0x63) & 0xFF
        _SBOX[p] = x
        _INV[x] = p
        if p == 1:
            break
    _SBOX[0] = 0x63
    _INV[0x63] = 0


_init_sbox()


def _xt(a: int) -> int:
    return ((a << 1) ^ 0x1B) & 0xFF if a & 0x80 else a << 1


def _mul(a: int, b: int) -> int:
    r = 0
    while b:
        if b & 1:
            r ^= a
        a = _xt(a)
        b >>= 1
    return r


def _expand(key: bytes) -> list[list[int]]:
    nk = len(key) // 4
    nr = nk + 6
    w = [list(key[4 * i:4 * i + 4]) for i in range(nk)]
    rcon = 1
    for i in range(nk, 4 * (nr + 1)):
        t = list(w[i - 1])
        if i % nk == 0:
            t = [_SBOX[t[1]] ^ rcon, _SBOX[t[2]], _SBOX[t[3]], _SBOX[t[0]]]
            rcon = _xt(rcon)
        elif nk > 6 and i % nk == 4:
            t = [_SBOX[x] for x in t]
        w.append([w[i - nk][k] ^ t[k] for k in range(4)])
    return [sum((w[4 * r + c] for c in range(4)), []) for r in range(nr + 1)]


def _decrypt_block(rk: list[list[int]], block: bytes) -> bytes:
    nr = len(rk) - 1
    s = [block[i] ^ rk[nr][i] for i in range(16)]
    for rnd in range(nr - 1, -1, -1):
        # 逆 ShiftRows
        s = [s[((i // 4 - i % 4) % 4) * 4 + i % 4] for i in range(16)]
        s = [_INV[x] for x in s]
        s = [s[i] ^ rk[rnd][i] for i in range(16)]
        if rnd:
            t = []
            for c in range(4):
                a = s[4 * c:4 * c + 4]
                t += [_mul(a[0], 14) ^ _mul(a[1], 11) ^ _mul(a[2], 13) ^ _mul(a[3], 9),
                      _mul(a[0], 9) ^ _mul(a[1], 14) ^ _mul(a[2], 11) ^ _mul(a[3], 13),
                      _mul(a[0], 13) ^ _mul(a[1], 9) ^ _mul(a[2], 14) ^ _mul(a[3], 11),
                      _mul(a[0], 11) ^ _mul(a[1], 13) ^ _mul(a[2], 9) ^ _mul(a[3], 14)]
            s = t
    return bytes(s)


def _encrypt_block(rk: list[list[int]], block: bytes) -> bytes:
    nr = len(rk) - 1
    s = [block[i] ^ rk[0][i] for i in range(16)]
    for rnd in range(1, nr + 1):
        s = [_SBOX[x] for x in s]
        s = [s[((i // 4 + i % 4) % 4) * 4 + i % 4] for i in range(16)]
        if rnd != nr:
            t = []
            for c in range(4):
                a = s[4 * c:4 * c + 4]
                t += [_mul(a[0], 2) ^ _mul(a[1], 3) ^ a[2] ^ a[3], a[0] ^ _mul(a[1], 2) ^ _mul(a[2], 3) ^ a[3],
                      a[0] ^ a[1] ^ _mul(a[2], 2) ^ _mul(a[3], 3), _mul(a[0], 3) ^ a[1] ^ a[2] ^ _mul(a[3], 2)]
            s = t
        s = [s[i] ^ rk[rnd][i] for i in range(16)]
    return bytes(s)


def aes_cbc_decrypt(key: bytes, data: bytes, iv: bytes | None = None, unpad: bool = True) -> bytes:
    rk = _expand(key)
    if iv is None:
        iv, data = data[:16], data[16:]
    out = bytearray()
    prev = iv
    for i in range(0, len(data) - len(data) % 16, 16):
        blk = data[i:i + 16]
        dec = _decrypt_block(rk, blk)
        out += bytes(a ^ b for a, b in zip(dec, prev))
        prev = blk
    if unpad and out and 1 <= out[-1] <= 16 and out.endswith(bytes([out[-1]]) * out[-1]):
        out = out[:-out[-1]]
    return bytes(out)


def aes_ecb_encrypt(key: bytes, data: bytes) -> bytes:
    rk = _expand(key)
    return b"".join(_encrypt_block(rk, data[i:i + 16]) for i in range(0, len(data), 16))


def aes_cbc_encrypt(key: bytes, data: bytes, iv: bytes) -> bytes:
    rk = _expand(key)
    out = bytearray()
    prev = iv
    for i in range(0, len(data), 16):
        blk = bytes(a ^ b for a, b in zip(data[i:i + 16], prev))
        prev = _encrypt_block(rk, blk)
        out += prev
    return bytes(out)


# ---------------------------------------------------------------------- 標準のセキュリティハンドラ
class Decryptor:
    def __init__(self, enc: dict, file_id: bytes, password: bytes = b""):
        self.v = int(enc.get("V", 0))
        self.r = int(enc.get("R", 2))
        self.length = int(enc.get("Length", 40)) // 8 if self.v != 5 else 32
        self.o = bytes(enc.get("O", b""))
        self.u = bytes(enc.get("U", b""))
        self.p = int(enc.get("P", 0))
        self.meta = enc.get("EncryptMetadata", True) is not False
        self.aes = False
        cf = enc.get("CF") or {}
        stmf = getattr(enc.get("StmF"), "name", None)
        if self.v >= 4 and stmf and stmf in cf:
            cfm = getattr(cf[stmf].get("CFM"), "name", "V2")
            self.aes = cfm in ("AESV2", "AESV3")
            if cfm == "AESV2":
                self.length = 16
        self.identity = stmf == "Identity"
        if self.r >= 5:
            self.key = self._key_r6(password, enc)
        else:
            self.key = self._key(password, file_id)
            if not self._check_user(file_id):
                raise ValueError("開くためのパスワードが必要なPDFです")

    def _key(self, pw: bytes, file_id: bytes) -> bytes:
        h = hashlib.md5((pw + PAD)[:32] + self.o + struct.pack("<i", self.p) + file_id)
        if self.r >= 4 and not self.meta:
            h.update(b"\xff\xff\xff\xff")
        key = h.digest()
        n = 5 if self.r == 2 else self.length
        if self.r >= 3:
            for _ in range(50):
                key = hashlib.md5(key[:n]).digest()
        return key[:n]

    def _check_user(self, file_id: bytes) -> bool:
        if self.r == 2:
            return rc4(self.key, PAD) == self.u[:32]
        h = hashlib.md5(PAD + file_id).digest()
        x = rc4(self.key, h)
        for i in range(1, 20):
            x = rc4(bytes(k ^ i for k in self.key), x)
        return x[:16] == self.u[:16]

    def _key_r6(self, pw: bytes, enc: dict) -> bytes:
        ue = bytes(enc.get("UE", b""))
        vsalt, ksalt = self.u[32:40], self.u[40:48]
        if self._hash6(pw, vsalt, b"") != self.u[:32]:
            raise ValueError("開くためのパスワードが必要なPDFです")
        ik = self._hash6(pw, ksalt, b"")
        return aes_cbc_decrypt(ik, ue, iv=b"\x00" * 16, unpad=False)

    def _hash6(self, pw: bytes, salt: bytes, udata: bytes) -> bytes:
        k = hashlib.sha256(pw + salt + udata).digest()
        if self.r == 5:
            return k
        i = 0
        while True:
            k1 = (pw + k + udata) * 64
            e = aes_cbc_encrypt_noiv(k[:16], k[16:32], k1)
            mod = sum(e[:16]) % 3
            k = (hashlib.sha256, hashlib.sha384, hashlib.sha512)[mod](e).digest()
            i += 1
            if i >= 64 and e[-1] <= i - 32:
                return k[:32]

    def obj_key(self, num: int, gen: int) -> bytes:
        if self.r >= 5:
            return self.key
        k = self.key + struct.pack("<I", num)[:3] + struct.pack("<H", gen)
        if self.aes:
            k += b"sAlT"
        return hashlib.md5(k).digest()[:min(len(self.key) + 5, 16)]

    def decrypt(self, data: bytes, num: int, gen: int) -> bytes:
        if self.identity:
            return data
        key = self.obj_key(num, gen)
        if self.aes:
            return aes_cbc_decrypt(key, data) if len(data) >= 16 else b""
        return rc4(key, data)


def _encrypt(self, data: bytes, num: int, gen: int) -> bytes:
    """書き込み用：新しい文字列・ストリームを、元の文書と同じ鍵で暗号化する。
    AES の初期化ベクトルは、中身から決まる値にする（同じ入力からは同じファイルができるように）。"""
    if self.identity:
        return data
    key = self.obj_key(num, gen)
    if self.aes:
        iv = hashlib.md5(b"todoku-iv" + struct.pack("<II", num, gen) + data).digest()
        pad = 16 - len(data) % 16
        return iv + aes_cbc_encrypt(key, data + bytes([pad]) * pad, iv)
    return rc4(key, data)


Decryptor.encrypt = _encrypt


def aes_cbc_encrypt_noiv(key: bytes, iv: bytes, data: bytes) -> bytes:
    return aes_cbc_encrypt(key, data, iv)
