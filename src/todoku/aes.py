"""AES（FIPS-197）を、表引き（T テーブル）で速く計算する。外部のライブラリは使わない。

1ラウンドの SubBytes・ShiftRows・MixColumns を、256 項目の表 4 つを引いて XOR するだけにまとめる（Rijndael の提案書 5.2.1 の方法）。
復号は「同値な逆暗号」（FIPS-197 5.3.5）：途中のラウンド鍵に InvMixColumns をかけておき、暗号と同じ形で計算する。
状態は4つの32ビットの語（列ごと、上位バイトが0行目）で持つ。

pdf/crypto.py の AES（1バイトずつ計算する）より10倍以上速い。暗号化された Office のファイルは中身全体（数MB）を解くので、こちらを使う。
"""

from __future__ import annotations

import struct


def _tables():
    # GF(2^8)（x^8+x^4+x^3+x+1）の掛け算を、3 を生成元にした指数・対数の表で
    exp, log = [0] * 510, [0] * 256
    x = 1
    for i in range(255):
        exp[i] = exp[i + 255] = x
        log[x] = i
        x ^= (x << 1) ^ (0x11B if x & 0x80 else 0)  # x*3 = x*2 ^ x

    def mul(a: int, b: int) -> int:
        return exp[log[a] + log[b]] if a and b else 0

    sbox, inv = [0] * 256, [0] * 256
    for a in range(256):
        b = exp[255 - log[a]] if a else 0  # 逆数
        s = b
        for k in range(1, 5):
            s ^= ((b << k) | (b >> (8 - k))) & 0xFF
        sbox[a] = s ^ 0x63
        inv[s ^ 0x63] = a
    te = [[0] * 256 for _ in range(4)]
    td = [[0] * 256 for _ in range(4)]
    for a in range(256):
        s, t = sbox[a], inv[a]
        w = (mul(s, 2) << 24) | (s << 16) | (s << 8) | mul(s, 3)  # MixColumns の列 [2,1,1,3]
        v = (mul(t, 14) << 24) | (mul(t, 9) << 16) | (mul(t, 13) << 8) | mul(t, 11)  # InvMixColumns の列 [14,9,13,11]
        for k in range(4):  # 行がずれた表は、語を 8 ビットずつ右に回したもの
            te[k][a] = ((w >> 8 * k) | (w << (32 - 8 * k))) & 0xFFFFFFFF
            td[k][a] = ((v >> 8 * k) | (v << (32 - 8 * k))) & 0xFFFFFFFF
    return sbox, inv, te, td, mul


_SBOX, _INV, _TE, _TD, _MUL = _tables()


def _sub(w: int) -> int:
    return (_SBOX[w >> 24] << 24) | (_SBOX[(w >> 16) & 255] << 16) | (_SBOX[(w >> 8) & 255] << 8) | _SBOX[w & 255]


def _encryptor(keys):
    """1ブロックを暗号化する関数（ShiftRows で、行 r は列を r だけ左から取る）。"""
    T0, T1, T2, T3 = _TE
    S = _SBOX
    (k0, k1, k2, k3), mid, (l0, l1, l2, l3) = keys[0], keys[1:-1], keys[-1]

    def block(s0: int, s1: int, s2: int, s3: int) -> tuple[int, int, int, int]:
        s0 ^= k0
        s1 ^= k1
        s2 ^= k2
        s3 ^= k3
        for r0, r1, r2, r3 in mid:
            t0 = T0[s0 >> 24] ^ T1[(s1 >> 16) & 255] ^ T2[(s2 >> 8) & 255] ^ T3[s3 & 255] ^ r0
            t1 = T0[s1 >> 24] ^ T1[(s2 >> 16) & 255] ^ T2[(s3 >> 8) & 255] ^ T3[s0 & 255] ^ r1
            t2 = T0[s2 >> 24] ^ T1[(s3 >> 16) & 255] ^ T2[(s0 >> 8) & 255] ^ T3[s1 & 255] ^ r2
            s3 = T0[s3 >> 24] ^ T1[(s0 >> 16) & 255] ^ T2[(s1 >> 8) & 255] ^ T3[s2 & 255] ^ r3
            s0, s1, s2 = t0, t1, t2
        return ((S[s0 >> 24] << 24 | S[(s1 >> 16) & 255] << 16 | S[(s2 >> 8) & 255] << 8 | S[s3 & 255]) ^ l0,
                (S[s1 >> 24] << 24 | S[(s2 >> 16) & 255] << 16 | S[(s3 >> 8) & 255] << 8 | S[s0 & 255]) ^ l1,
                (S[s2 >> 24] << 24 | S[(s3 >> 16) & 255] << 16 | S[(s0 >> 8) & 255] << 8 | S[s1 & 255]) ^ l2,
                (S[s3 >> 24] << 24 | S[(s0 >> 16) & 255] << 16 | S[(s1 >> 8) & 255] << 8 | S[s2 & 255]) ^ l3)

    return block


def _decryptor(keys):
    """1ブロックを復号する関数（同値な逆暗号。InvShiftRows で、行 r は列を r だけ右から取る）。"""
    T0, T1, T2, T3 = _TD
    S = _INV
    (k0, k1, k2, k3), mid, (l0, l1, l2, l3) = keys[0], keys[1:-1], keys[-1]

    def block(s0: int, s1: int, s2: int, s3: int) -> tuple[int, int, int, int]:
        s0 ^= k0
        s1 ^= k1
        s2 ^= k2
        s3 ^= k3
        for r0, r1, r2, r3 in mid:
            t0 = T0[s0 >> 24] ^ T1[(s3 >> 16) & 255] ^ T2[(s2 >> 8) & 255] ^ T3[s1 & 255] ^ r0
            t1 = T0[s1 >> 24] ^ T1[(s0 >> 16) & 255] ^ T2[(s3 >> 8) & 255] ^ T3[s2 & 255] ^ r1
            t2 = T0[s2 >> 24] ^ T1[(s1 >> 16) & 255] ^ T2[(s0 >> 8) & 255] ^ T3[s3 & 255] ^ r2
            s3 = T0[s3 >> 24] ^ T1[(s2 >> 16) & 255] ^ T2[(s1 >> 8) & 255] ^ T3[s0 & 255] ^ r3
            s0, s1, s2 = t0, t1, t2
        return ((S[s0 >> 24] << 24 | S[(s3 >> 16) & 255] << 16 | S[(s2 >> 8) & 255] << 8 | S[s1 & 255]) ^ l0,
                (S[s1 >> 24] << 24 | S[(s0 >> 16) & 255] << 16 | S[(s3 >> 8) & 255] << 8 | S[s2 & 255]) ^ l1,
                (S[s2 >> 24] << 24 | S[(s1 >> 16) & 255] << 16 | S[(s0 >> 8) & 255] << 8 | S[s3 & 255]) ^ l2,
                (S[s3 >> 24] << 24 | S[(s2 >> 16) & 255] << 16 | S[(s1 >> 8) & 255] << 8 | S[s0 & 255]) ^ l3)

    return block


class Aes:
    """1つの鍵（16・24・32 バイト）の AES。ブロックは 16 バイト。データの長さは 16 の倍数であること。"""

    def __init__(self, key: bytes):
        nk = len(key) // 4
        if len(key) not in (16, 24, 32):
            raise ValueError("AES の鍵は 16・24・32 バイトです")
        nr = nk + 6
        w = list(struct.unpack(f">{nk}I", key))
        rcon = 1
        for i in range(nk, 4 * (nr + 1)):  # 鍵の拡張（FIPS-197 5.2）
            t = w[i - 1]
            if i % nk == 0:
                t = _sub(((t << 8) | (t >> 24)) & 0xFFFFFFFF) ^ (rcon << 24)
                rcon = _MUL(rcon, 2)
            elif nk > 6 and i % nk == 4:
                t = _sub(t)
            w.append(w[i - nk] ^ t)
        enc = [tuple(w[4 * r:4 * r + 4]) for r in range(nr + 1)]
        T0, T1, T2, T3 = _TD
        dec = [enc[nr]] + [tuple(T0[_SBOX[x >> 24]] ^ T1[_SBOX[(x >> 16) & 255]] ^ T2[_SBOX[(x >> 8) & 255]] ^ T3[_SBOX[x & 255]]
                                 for x in enc[r]) for r in range(nr - 1, 0, -1)] + [enc[0]]
        self._enc = _encryptor(enc)
        self._dec = _decryptor(dec)

    @staticmethod
    def _words(data: bytes) -> tuple[int, ...]:
        if len(data) % 16:
            raise ValueError("AES のデータの長さが 16 の倍数ではありません")
        return struct.unpack(f">{len(data) // 4}I", data)

    def _ecb(self, f, data: bytes) -> bytes:
        w = self._words(data)
        out: list[int] = []
        for i in range(0, len(w), 4):
            out += f(w[i], w[i + 1], w[i + 2], w[i + 3])
        return struct.pack(f">{len(out)}I", *out)

    def encrypt_ecb(self, data: bytes) -> bytes:
        return self._ecb(self._enc, data)

    def decrypt_ecb(self, data: bytes) -> bytes:
        return self._ecb(self._dec, data)

    def encrypt_cbc(self, data: bytes, iv: bytes) -> bytes:
        w = self._words(data)
        c0, c1, c2, c3 = struct.unpack(">4I", iv)
        f = self._enc
        out: list[int] = []
        for i in range(0, len(w), 4):
            c0, c1, c2, c3 = f(w[i] ^ c0, w[i + 1] ^ c1, w[i + 2] ^ c2, w[i + 3] ^ c3)
            out += (c0, c1, c2, c3)
        return struct.pack(f">{len(out)}I", *out)

    def decrypt_cbc(self, data: bytes, iv: bytes) -> bytes:
        raw = self.decrypt_ecb(data)
        if not raw:
            return b""
        # 各ブロックの復号に、1つ前の暗号文（先頭は初期化ベクトル）を XOR する：全体を1つの整数にしてまとめて
        prev = iv[:16] + data[:-16]
        return (int.from_bytes(raw, "big") ^ int.from_bytes(prev, "big")).to_bytes(len(raw), "big")
