"""PDFの暗号（標準のセキュリティハンドラ）を解く：RC4、AES-128、AES-256。外部のライブラリは使わない。

行政のPDFには「編集の制限」だけをかけたもの（開くパスワードは空）が多い。空のパスワードで開けるものは開く。
開くためのパスワードがかかったものは、渡されたパスワード（開くためのパスワードか、権限のパスワード）で開く。
パスワードそのものは、持たない・記録しない・返さない。
"""

from __future__ import annotations

import hashlib
import hmac
import stringprep
import struct
import unicodedata

from . import aes as _aes
from .objects import PDFDOC, PdfError

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
class NeedsPassword(PdfError):
    """開くためのパスワードがかかったPDFで、パスワードが渡されていない。"""

    def __init__(self, msg: str = "開くためのパスワードがかかったPDFです。パスワードを指定してください"):
        super().__init__(msg)


class WrongPassword(NeedsPassword):
    """渡されたパスワードが、開くためのパスワードにも権限のパスワードにも合わない（NeedsPassword の一種）。"""

    def __init__(self, msg: str = "パスワードが違います（このPDFの、開くためのパスワードにも権限のパスワードにも合いません）"):
        super().__init__(msg)


# 暗号フィルターの方式（PDF 32000-1 表25。AESV3 は ISO 32000-2）→ この中での呼び名
_CFM = {"V2": "rc4", "AESV2": "aesv2", "AESV3": "aesv3"}


class Decryptor:
    """標準のセキュリティハンドラ（PDF 32000-1 7.6.3。AES-256 は ISO 32000-2 7.6.4）：文書の鍵を求め、文字列とストリームの暗号を解く・かける。

    パスワードは、まず空で試す（開くためのパスワードのないPDF。7.6.3.1）。開けなければ、渡されたパスワードを
    開くためのパスワード（ユーザー）として、次に権限のパスワード（所有者）として確かめる（Algorithm 6・7、2.A）。
    パスワードそのものは持たない。持つのは文書の鍵と、どれで開いたか（auth：empty・user・owner）だけ。
    """

    def __init__(self, enc: dict, file_id: bytes, password: str | bytes | None = None, resolve=None):
        res = resolve or (lambda o: o)

        def get(k, default=None):
            v = res(enc.get(k))
            return default if v is None else v

        handler = getattr(get("Filter"), "name", "")
        if handler != "Standard":
            raise PdfError(f"この暗号（{handler or '名前なし'}）には対応していません。読めるのはパスワードの暗号（Standard）だけです")
        self.v = int(get("V", 0))
        self.r = int(get("R", 0))
        if self.r not in (2, 3, 4, 5, 6):
            raise PdfError(f"この暗号の版（R {self.r}）には対応していません")
        self.p = int(get("P", 0)) & 0xFFFFFFFF  # 32 ビットの符号なしの値として扱う（7.6.3.2）
        self.o = bytes(get("O", b""))
        self.u = bytes(get("U", b""))
        self.oe = bytes(get("OE", b""))
        self.ue = bytes(get("UE", b""))
        if not self.o or not self.u or (self.r >= 5 and (len(self.o) < 48 or len(self.u) < 48)):
            raise PdfError("暗号の辞書の O・U が壊れています")
        self.meta = get("EncryptMetadata", True) is not False
        self.file_id = bytes(file_id)
        # 暗号フィルター（7.6.5）：名前 → 方式。V 4 以上は、文字列・ストリーム・埋め込みファイルで別の方式にできる
        cf = get("CF", {})
        self._cf = {}
        self._cf_len = {}
        for name, d in (cf.items() if isinstance(cf, dict) else ()):
            d = res(d)
            if isinstance(d, dict):
                self._cf[name] = getattr(res(d.get("CFM")), "name", "None")
                if isinstance(res(d.get("Length")), int):
                    self._cf_len[name] = res(d.get("Length"))
        if self.v >= 4:
            stmf = getattr(get("StmF"), "name", "Identity")  # 省いたときは Identity（表20）
            strf = getattr(get("StrF"), "name", "Identity")
            self.stream_method = self.method(stmf)
            self.string_method = self.method(strf)
            eff = getattr(get("EFF"), "name", None)
            self.embedded_method = self.method(eff) if eff else self.stream_method
        elif self.v in (0, 1, 2):
            self.stream_method = self.string_method = self.embedded_method = "rc4"
        else:
            raise PdfError(f"この暗号の版（V {self.v}）には対応していません")
        self.length = self._key_length(get("Length"), (stmf, strf) if self.v >= 4 else ())
        self.key, self.auth = self._authenticate(password)
        # AES-256 の権限の値（Perms）が、文書の鍵で解いて P と合うか（ISO 32000-2 Algorithm 13）。合わなくても開く
        self.perms_ok = self._check_perms(get("Perms")) if self.r >= 5 else None

    # ------------------------------------------------------------------ 方式と鍵の長さ
    def method(self, name: str) -> str:
        """暗号フィルターの名前 → 方式（identity・rc4・aesv2・aesv3）。"""
        if name == "Identity":
            return "identity"  # 標準の名前（表26）。CF に同じ名前があっても使わない
        cfm = self._cf.get(name)
        if cfm is None:
            raise PdfError(f"暗号フィルター {name} が、暗号の辞書（CF）にありません")
        if cfm not in _CFM:
            raise PdfError(f"この暗号の方式（{cfm}）には対応していません")
        return _CFM[cfm]

    def _key_length(self, length, filters) -> int:
        """文書の鍵のバイト数（Algorithm 2 の n）。"""
        if self.r >= 5:
            return 32
        if self.r == 2 or self.v in (0, 1):
            return 5  # 40 ビット
        if "aesv2" in (self.stream_method, self.string_method, self.embedded_method):
            return 16  # AES-128 の鍵は 128 ビット
        if length is None:
            # V 4 で Length を省いたとき：暗号フィルターの Length（標準のセキュリティハンドラは 8 ビット単位。16 なら 128 ビット）
            for name in filters:
                n = self._cf_len.get(name)
                if isinstance(n, int):
                    length = n * 8 if n <= 16 else n
                    break
        bits = int(length) if isinstance(length, (int, float)) else (128 if self.v >= 4 else 40)
        return min(16, max(5, bits // 8))

    # ------------------------------------------------------------------ パスワード
    def _authenticate(self, password) -> tuple[bytes, str]:
        key = self._user(b"")
        if key is not None:
            return key, "empty"
        if password is None or password in ("", b""):
            raise NeedsPassword()
        cands = _candidates(password, self.r)
        for pw in cands:
            key = self._user(pw)
            if key is not None:
                return key, "user"
        for pw in cands:
            key = self._owner(pw)
            if key is not None:
                return key, "owner"
        raise WrongPassword()

    def _user(self, pw: bytes) -> bytes | None:
        """開くためのパスワードを確かめる（R2〜4 は Algorithm 6、R5・6 は 2.A の e・f）。合えば文書の鍵。"""
        if self.r >= 5:
            if len(self.u) < 48 or self._hash(pw, self.u[32:40], b"") != self.u[:32]:
                return None
            return self._unwrap(self._hash(pw, self.u[40:48], b""), self.ue, "UE")
        return self._user_padded((pw + PAD)[:32])

    def _owner(self, pw: bytes) -> bytes | None:
        """権限のパスワードを確かめる（R2〜4 は Algorithm 7、R5・6 は 2.A の c・d）。合えば文書の鍵。"""
        if self.r >= 5:
            u48 = self.u[:48]
            if len(self.o) < 48 or self._hash(pw, self.o[32:40], u48) != self.o[:32]:
                return None
            return self._unwrap(self._hash(pw, self.o[40:48], u48), self.oe, "OE")
        # Algorithm 3 の a〜d で RC4 の鍵を作り、O を解くと、開くためのパスワード（32 バイトに詰めたもの）が出る
        h = hashlib.md5((pw + PAD)[:32]).digest()
        if self.r >= 3:
            for _ in range(50):
                h = hashlib.md5(h).digest()
        k = h[:self.length]
        x = self.o[:32]
        if self.r == 2:
            x = rc4(k, x)
        else:
            for i in range(19, -1, -1):
                x = rc4(bytes(b ^ i for b in k), x)
        return self._user_padded(x)

    def _user_padded(self, padded: bytes) -> bytes | None:
        """R2〜4：32 バイトに詰めたパスワードから鍵を作り（Algorithm 2）、U と照らす（Algorithm 4・5 の、最後の手順の前まで）。"""
        h = hashlib.md5(padded + self.o[:32] + struct.pack("<I", self.p) + self.file_id)
        if self.r >= 4 and not self.meta:
            h.update(b"\xff\xff\xff\xff")
        key = h.digest()
        n = self.length
        if self.r >= 3:
            for _ in range(50):
                key = hashlib.md5(key[:n]).digest()
        key = key[:n]
        if self.r == 2:
            return key if rc4(key, PAD) == self.u[:32] else None
        x = rc4(key, hashlib.md5(PAD + self.file_id).digest())
        for i in range(1, 20):
            x = rc4(bytes(b ^ i for b in key), x)
        return key if x == self.u[:16] else None  # R3 以上は先頭の 16 バイトだけを比べる

    def _hash(self, pw: bytes, salt: bytes, udata: bytes) -> bytes:
        """R5：SHA-256。R6：ISO 32000-2 Algorithm 2.B（AES-128 と SHA-2 を 64 回以上くり返す）。"""
        k = hashlib.sha256(pw + salt + udata).digest()
        if self.r == 5:
            return k
        rnd = 0
        while True:
            e = _aes.cbc_encrypt(k[:16], k[16:32], (pw + k + udata) * 64)
            # E の先頭 16 バイトを数とみて 3 で割った余り（256 ≡ 1 なので、バイトの和の余りと同じ）で、次のハッシュを選ぶ
            k = (hashlib.sha256, hashlib.sha384, hashlib.sha512)[sum(e[:16]) % 3](e).digest()
            rnd += 1
            if rnd >= 64 and e[-1] <= rnd - 32:
                return k[:32]

    def _unwrap(self, ik: bytes, wrapped: bytes, name: str) -> bytes:
        """UE・OE を解いて、文書の鍵（32 バイト）を取り出す（AES-256、CBC、初期化ベクトル 0、詰め物なし）。"""
        if len(wrapped) < 32:
            raise PdfError(f"暗号の辞書の {name} が壊れています")
        return _aes.cbc_decrypt(ik, bytes(16), wrapped[:32])

    def _check_perms(self, perms) -> bool:
        if not isinstance(perms, bytes) or len(perms) < 16:
            return False
        b = _aes.ecb_decrypt_block(self.key, bytes(perms[:16]))
        return b[9:12] == b"adb" and int.from_bytes(b[:4], "little") == self.p

    # ------------------------------------------------------------------ 暗号を解く・かける（Algorithm 1、AES-256 は 1.A）
    def object_key(self, num: int, gen: int, aes: bool) -> bytes:
        if self.r >= 5:
            return self.key  # AES-256 は、文書の鍵をそのまま使う
        k = self.key + (num & 0xFFFFFF).to_bytes(3, "little") + (gen & 0xFFFF).to_bytes(2, "little")
        if aes:
            k += b"sAlT"
        return hashlib.md5(k).digest()[:min(len(self.key) + 5, 16)]

    def stream_method_for(self, crypt: str | None = None, kind: str = "") -> str:
        """ストリームの方式。crypt：Filter の Crypt が指す暗号フィルター。kind：metadata（文書のメタデータ）・embedded（埋め込みファイル）。"""
        if crypt is not None:
            return self.method(crypt)
        if kind == "metadata" and not self.meta and self.v >= 4:
            return "identity"  # EncryptMetadata が false：文書のメタデータは暗号化しない
        if kind == "embedded":
            return self.embedded_method
        return self.stream_method

    def decrypt_string(self, data: bytes, num: int, gen: int) -> bytes:
        return self._decrypt(self.string_method, data, num, gen)

    def decrypt_stream(self, data: bytes, num: int, gen: int, crypt: str | None = None, kind: str = "") -> bytes:
        return self._decrypt(self.stream_method_for(crypt, kind), data, num, gen)

    def encrypt_string(self, data: bytes, num: int, gen: int) -> bytes:
        return self._encrypt(self.string_method, data, num, gen)

    def encrypt_stream(self, data: bytes, num: int, gen: int, crypt: str | None = None, kind: str = "") -> bytes:
        return self._encrypt(self.stream_method_for(crypt, kind), data, num, gen)

    def _decrypt(self, method: str, data: bytes, num: int, gen: int) -> bytes:
        if method == "identity":
            return data
        if method == "rc4":
            return rc4(self.object_key(num, gen, False), data)
        if len(data) < 16:
            return b""
        out = _aes.cbc_decrypt(self.object_key(num, gen, True), data[:16], data[16:])
        # 詰め物（RFC 2898、7.6.2）を外す。形の合わないときは、そのまま
        if out and 1 <= out[-1] <= 16 and out.endswith(bytes([out[-1]]) * out[-1]):
            out = out[:-out[-1]]
        return out

    def _encrypt(self, method: str, data: bytes, num: int, gen: int) -> bytes:
        """書き込み用：新しい文字列・ストリームを、元の文書と同じ鍵で暗号化する。
        AES の初期化ベクトルは、同じ入力から同じファイルができるよう乱数にせず、文書の鍵を使った HMAC で作る
        （鍵を知らない人が、初期化ベクトルから中身を当てて確かめることはできない）。"""
        if method == "identity":
            return data
        if method == "rc4":
            return rc4(self.object_key(num, gen, False), data)
        iv = hmac.new(self.key, b"todoku-iv" + struct.pack("<QQ", num, gen) + data, hashlib.sha256).digest()[:16]
        pad = 16 - len(data) % 16
        return iv + _aes.cbc_encrypt(self.object_key(num, gen, True), iv, data + bytes([pad]) * pad)


# ---------------------------------------------------------------------- パスワードのバイト列
_PDFDOC_INV = {ch: code for code, ch in PDFDOC.items()}


def _pdfdoc(s: str) -> bytes:
    """PDFDocEncoding（PDF 32000-1 附属書 D）のバイト列。表せない字があれば ValueError（字は伝えない）。"""
    out = bytearray()
    for ch in s:
        code = _PDFDOC_INV.get(ch)
        if code is None:
            if ord(ch) > 0xFF or ord(ch) in PDFDOC:
                raise ValueError("PDFDocEncoding で表せない字があります")
            code = ord(ch)
        out.append(code)
    return bytes(out)


def _candidates(password: str | bytes, r: int) -> list[bytes]:
    """確かめるパスワードのバイト列（重ならないように、確からしい順に）。

    R2〜4：PDFDocEncoding（Algorithm 2 の a）。表せない字（日本語など）は、作ったソフトによって
    シフト JIS（日本語の Windows の文字コード）か UTF-8 なので、それも試す。
    R5・6：SASLprep で正規化した UTF-8 の、先頭 127 バイト（ISO 32000-2 Algorithm 2.A）。
    正規化しないで作るソフトもあるので、そのままの UTF-8 も試す。"""
    if isinstance(password, (bytes, bytearray)):
        b = bytes(password)
        return [b[:127] if r >= 5 else b]
    if r >= 5:
        forms = (lambda s: saslprep(s).encode("utf-8")[:127], lambda s: s.encode("utf-8")[:127])
    else:
        forms = (_pdfdoc, lambda s: s.encode("utf-8"), lambda s: s.encode("cp932"))
    out: list[bytes] = []
    for form in forms:
        try:
            b = form(password)
        except ValueError:  # 表せない字（UnicodeEncodeError も ValueError の一種）
            continue
        if b and b not in out:
            out.append(b)
    return out


_PROHIBITED = (stringprep.in_table_c12, stringprep.in_table_c21_c22, stringprep.in_table_c3, stringprep.in_table_c4,
               stringprep.in_table_c5, stringprep.in_table_c6, stringprep.in_table_c7, stringprep.in_table_c8,
               stringprep.in_table_c9)


def saslprep(s: str) -> str:
    """RFC 4013 の SASLprep（RFC 3454 の stringprep の手順で、表は Unicode 3.2）。使えない字があれば ValueError。

    確かめる側（問い合わせ）として扱うので、未割り当ての字は許す。"""
    # 対応づけ（RFC 4013 2.1）：何にも対応しない字（B.1）は消し、ASCII でない空白（C.1.2）は空白にする
    s = "".join(" " if stringprep.in_table_c12(ch) else ch for ch in s if not stringprep.in_table_b1(ch))
    # 正規化（2.2）：NFKC
    s = unicodedata.ucd_3_2_0.normalize("NFKC", s)
    # 使えない字（2.3）：制御文字・私用の字・非文字・サロゲート・表示を変える字・タグなど
    if any(t(ch) for ch in s for t in _PROHIBITED):
        raise ValueError("SASLprep では使えない字があります")
    # 双方向の字（2.4、RFC 3454 6章）：右から左の字があるなら、左から右の字はなく、最初と最後が右から左の字
    ral = [stringprep.in_table_d1(ch) for ch in s]
    if any(ral) and (any(stringprep.in_table_d2(ch) for ch in s) or not (ral[0] and ral[-1])):
        raise ValueError("SASLprep の双方向の字の決まりに合いません")
    return s


def aes_cbc_encrypt_noiv(key: bytes, iv: bytes, data: bytes) -> bytes:
    return aes_cbc_encrypt(key, data, iv)
