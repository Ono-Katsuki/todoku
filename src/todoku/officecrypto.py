"""パスワードで暗号化された Office のファイル（.xlsx・.xlsm・.docx・.docm）を解き、同じパスワードのまま暗号化し直す。

Microsoft の仕様書 MS-OFFCRYPTO 2.3.4（ECMA-376 Document Encryption）の手順どおり。暗号化したファイルは複合ファイル（CFB）で、
中身の ZIP を暗号化した EncryptedPackage ストリーム（先頭8バイトが中身の大きさ StreamSize）と、鍵の情報の EncryptionInfo ストリームを持つ。
- Agile（EncryptionInfo の版 4.4。鍵の情報は XML。ふつうは AES-256・SHA-512・10万回）：鍵 2.3.4.11、初期化ベクトル 2.3.4.12、
  パスワードの確かめと中間鍵 2.3.4.13、改ざん検知（HMAC）2.3.4.14、4096 バイトの区切りごとの CBC 2.3.4.15
- Standard（版 2.2・3.2・4.2。AES-128/192/256・SHA-1・5万回・中身全体を ECB）：2.3.4.5〜2.3.4.9
それ以外（Extensible・RC4・証明書だけの鍵・権利管理（IRM）など）は、推測で解かずに、対応していないことを伝えて止める。

使い方：
- is_encrypted_ooxml(blob) -> bool：EncryptedPackage を持つ複合ファイル（暗号化・保護された .xlsx・.docx など）か
- decrypt_ooxml(blob, password) -> (中身の ZIP, 鍵)：パスワードが違えば WrongPassword、対応していない方式や壊れたファイルは CryptoError
- encrypt_ooxml(鍵, 中身の ZIP) -> bytes：元と同じパスワード・同じ鍵の情報で暗号化し直したファイル
- verify_integrity(blob, 鍵) -> bool：その鍵で暗号化したファイルで、改ざん検知（Agile の HMAC）が合うか

暗号化し直すときに変えるのは、EncryptedPackage と、Agile の encryptedHmacValue（XML の属性の値）だけ。
塩・中間鍵・ほかのストリーム（DataSpaces など）は1バイトも変えないので、パスワードは元のまま。
中身を変えなければ、元のファイルとバイト単位で一致する（最後のブロックの詰め物も、元の詰め物を使う）。
書き出す前に読み直し、鍵で解くと中身に戻ること・HMAC が合うこと・ほかのストリームが元のままであることを確かめる。

改ざん検知の鍵の長さ：仕様書（2.3.4.14）は keyData の saltSize と書くが、Office は hashSize（SHA-512 なら64バイト）の鍵を使う。
決めつけずに、元のファイルの HMAC が合う方（hashSize、saltSize の順）を使い、どちらも合わなければ止める。
パスワードは鍵を導くのに使うだけで、鍵にもエラーの文にも残さない。
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import re
import struct
import xml.etree.ElementTree as ET
from dataclasses import dataclass

from .aes import Aes
from .cfb import CFB, MAGIC

INFO, PACKAGE = "EncryptionInfo", "EncryptedPackage"
NS = "{http://schemas.microsoft.com/office/2006/encryption}"
PASSWORD_URI = "http://schemas.microsoft.com/office/2006/keyEncryptor/password"
SEGMENT = 4096  # 2.3.4.15：区切りの大きさ
MAX_SPIN = 10_000_000  # 2.3.4.10：spinCount の上限
STANDARD_SPIN = 50_000  # 2.3.4.7

# blockKey（2.3.4.13・2.3.4.14）
BK_VERIFIER_INPUT = bytes.fromhex("fea7d2763b4b9e79")
BK_VERIFIER_HASH = bytes.fromhex("d7aa0f6d3061344e")
BK_KEY = bytes.fromhex("146e0be7abacd0d6")
BK_HMAC_KEY = bytes.fromhex("5fb2ad010cb9e1f6")
BK_HMAC_VALUE = bytes.fromhex("a0677f02b22c8433")

# EncryptionHeader（2.3.1・2.3.2）
F_CRYPTOAPI, F_EXTERNAL, F_AES = 0x04, 0x10, 0x20
AES_BITS = {0x660E: 128, 0x660F: 192, 0x6610: 256}
SHA1_ID = 0x8004

HASHES = {"SHA1": "sha1", "SHA256": "sha256", "SHA384": "sha384", "SHA512": "sha512", "MD5": "md5"}
_HMAC_VALUE = re.compile(rb"""\sencryptedHmacValue\s*=\s*(["'])([^"']*)\1""")


class CryptoError(ValueError):
    pass


class WrongPassword(CryptoError):
    pass


@dataclass(frozen=True, repr=False, eq=False)
class OoxmlKey:
    """暗号化し直すための鍵と、元のファイル。パスワードそのものは持たない。"""

    method: str  # "agile" / "standard"
    cipher: str  # 例 "AES-256"
    hash_name: str  # 例 "SHA512"
    spin_count: int
    container: bytes  # 元のファイル（ほかのストリームをそのまま残すため）
    info: bytes  # 元の EncryptionInfo
    size: int  # 元の中身（ZIP）の大きさ
    stream_len: int  # 元の EncryptedPackage の長さ
    tail: bytes  # 元の中身より後ろの詰め物（復号したもの）
    data_key: bytes  # 中身を暗号化する鍵（Agile は中間鍵、Standard はパスワードから導いた鍵）
    digest: str = "sha1"  # hashlib の名前（Agile：区切りの初期化ベクトルと HMAC）
    hash_size: int = 20
    key_salt: bytes = b""  # keyData の塩（Agile）
    block_size: int = 16
    hmac_key: bytes | None = None  # 改ざん検知の鍵（Agile で dataIntegrity があるとき）
    hmac_span: tuple[int, int] = (0, 0)  # EncryptionInfo の中の encryptedHmacValue の値の位置
    hmac_pad: bytes = b""  # encryptedHmacValue を復号したものの、HMAC より後ろの詰め物

    @property
    def has_integrity(self) -> bool:
        return self.hmac_key is not None

    @property
    def label(self) -> str:
        return f"パスワードで暗号化（{'Agile' if self.method == 'agile' else 'Standard'}・{self.cipher}・{self.hash_name}）"

    def __repr__(self) -> str:
        return f"OoxmlKey({self.method}, {self.cipher}, {self.hash_name}, spin={self.spin_count})"


# ---------------------------------------------------------------------- 公開する関数
def is_encrypted_ooxml(blob: bytes) -> bool:
    """EncryptedPackage を持つ複合ファイル（パスワードや権利管理で保護された .xlsx・.docx など）か。"""
    if blob[:8] != MAGIC:
        return False
    try:
        return CFB(blob).has(PACKAGE)
    except Exception:  # noqa: BLE001  複合ファイルとして読めなければ、暗号化された Office のファイルではない
        return False


def decrypt_ooxml(blob: bytes, password: str) -> tuple[bytes, OoxmlKey]:
    """パスワードで中身の ZIP を取り出す。書き戻しに使う鍵もいっしょに返す。"""
    if not isinstance(password, str):
        raise TypeError("パスワードは文字列で渡してください")
    info, stream = _open(blob)
    if len(info) < 8:
        raise CryptoError("鍵の情報（EncryptionInfo）が短すぎます（壊れています）")
    major, minor = struct.unpack_from("<HH", info)
    pw = password.encode("utf-16-le", "surrogatepass")  # 仕様どおり Unicode（UTF-16）の文字の並び
    if (major, minor) == (4, 4):
        return _agile_decrypt(blob, info, stream, pw)
    if minor == 2 and major in (2, 3, 4):
        return _standard_decrypt(blob, info, stream, pw)
    if minor == 3 and major in (3, 4):
        raise CryptoError("拡張の暗号化（Extensible Encryption：ほかの会社の暗号の部品を使うもの）には対応していません")
    raise CryptoError(f"この暗号化の方式（EncryptionInfo の版 {major}.{minor}）には対応していません")


def encrypt_ooxml(key: OoxmlKey, package: bytes) -> bytes:
    """中身の ZIP を、key と同じパスワード・同じ鍵の情報で暗号化し、元のファイルの EncryptedPackage を差し替えたファイルを返す。"""
    package = bytes(package)
    if package[:2] != b"PK":
        raise CryptoError("暗号化する中身が ZIP（.xlsx・.docx の形）ではありません")
    c = CFB(key.container)
    stream = _encrypt_stream(key, package, c.cutoff)
    info = _with_hmac(key, stream) if key.hmac_key is not None else key.info
    blob = c.replace(PACKAGE, stream, wipe=True)
    if info != key.info:
        blob = CFB(blob).replace(INFO, info)
    _check(key, blob, package, stream, info)
    return blob


def verify_integrity(blob: bytes, key: OoxmlKey) -> bool:
    """blob が key と同じパスワード・同じ鍵の情報で暗号化されていて、改ざん検知が合うか。

    Agile：EncryptionInfo が（encryptedHmacValue の値のほかは）key と同じで、EncryptedPackage 全体（StreamSize を含む）の
    HMAC が合えば True。Standard（と dataIntegrity のない Agile）には改ざん検知のしくみがない。そのときは EncryptionInfo が
    key と同じで、中身の大きさの欄が EncryptedPackage に収まっていれば True（key.has_integrity が False なら HMAC は確かめていない）。"""
    try:
        c = CFB(blob)
        info, stream = c.read(INFO), c.read(PACKAGE)
    except Exception:  # noqa: BLE001  読めないファイルは、合わないと答える
        return False
    a, b = key.hmac_span if key.hmac_key is not None else (len(key.info), len(key.info))
    if len(info) != len(key.info) or info[:a] != key.info[:a] or info[b:] != key.info[b:]:
        return False
    if len(stream) < 8 or 8 + _round_up(struct.unpack_from("<Q", stream)[0], key.block_size) > len(stream):
        return False
    if key.hmac_key is None:
        return True
    try:
        raw = _cbc_decrypt(key.data_key, _b64(info[a:b].decode("ascii")), _iv(key.digest, key.key_salt, BK_HMAC_VALUE, key.block_size))
    except (CryptoError, UnicodeDecodeError):
        return False
    return hmac.compare_digest(hmac.new(key.hmac_key, stream, key.digest).digest(), raw[:key.hash_size])


# ---------------------------------------------------------------------- 共通
def _open(blob: bytes) -> tuple[bytes, bytes]:
    if blob[:2] == b"PK":
        raise CryptoError("暗号化されていないファイルです（ZIP のままの .xlsx・.docx など）")
    if blob[:8] != MAGIC:
        raise CryptoError("暗号化された Office のファイル（.xlsx・.docx などをパスワードで保護したもの）ではありません")
    try:
        c = CFB(blob)
        has_package, has_info = c.has(PACKAGE), c.has(INFO)
        if has_package and has_info:
            return c.read(INFO), c.read(PACKAGE)
    except Exception as e:  # noqa: BLE001
        raise CryptoError("複合ファイルとして読めません（壊れています）") from e
    if not has_package:
        raise CryptoError("暗号化された .xlsx・.docx ではありません（複合ファイルに EncryptedPackage がありません）")
    raise CryptoError("パスワードの鍵の情報（EncryptionInfo）がありません。権利管理（IRM）で保護されたファイルには対応していません")


def _h(name: str, data: bytes) -> bytes:
    return hashlib.new(name, data).digest()


def _spin(name: str, first: bytes, count: int) -> bytes:
    """H0 = H(塩 + パスワード)、Hn = H(繰り返しの番号（4バイト）+ Hn-1) を count 回（2.3.4.7・2.3.4.11）。"""
    new = getattr(hashlib, name)
    h = new(first).digest()
    for i in range(count):
        h = new(i.to_bytes(4, "little") + h).digest()
    return h


def _fit(b: bytes, n: int) -> bytes:
    """n バイトにそろえる：長ければ切り、短ければ 0x36 を足す（2.3.4.11・2.3.4.12）。"""
    return b[:n] + b"\x36" * (n - len(b))


def _iv(name: str, salt: bytes, block_key: bytes, size: int) -> bytes:
    """2.3.4.12：初期化ベクトル = H(塩 + blockKey) を、ブロックの大きさにそろえたもの。"""
    return _fit(_h(name, salt + block_key), size)


def _round_up(n: int, block: int) -> int:
    return -(-n // block) * block


def _b64(value: str | None) -> bytes:
    try:
        return base64.b64decode("".join((value or "").split()), validate=True)
    except ValueError as e:
        raise CryptoError("鍵の情報の値（base64）が読めません（壊れています）") from e


def _cbc_decrypt(key: bytes, data: bytes, iv: bytes) -> bytes:
    if not data or len(data) % 16:
        raise CryptoError("鍵の情報の暗号化された値の長さが、ブロックの倍数ではありません（壊れています）")
    return Aes(key).decrypt_cbc(data, iv)


def _package(plain: bytes, size: int) -> bytes:
    """解いた中身は ZIP のはず（Standard には改ざん検知がないので、壊れていてもここで気づく）。"""
    if plain[:min(size, 2)] != b"PK":
        raise CryptoError("解いた中身が ZIP（.xlsx・.docx の形）になっていません（壊れています）")
    return plain[:size]


def _split(stream: bytes, block: int) -> tuple[int, bytes]:
    """EncryptedPackage を、中身の大きさ（StreamSize）と、解けるブロックの並びに分ける。"""
    if len(stream) < 8:
        raise CryptoError("EncryptedPackage が短すぎます（壊れています）")
    size = struct.unpack_from("<Q", stream)[0]
    body = stream[8:len(stream) - (len(stream) - 8) % block]
    if size > len(body):
        raise CryptoError("暗号化された中身が途中で切れています（壊れています）")
    return size, body


# ---------------------------------------------------------------------- Agile（2.3.4.10〜2.3.4.15）
@dataclass(frozen=True)
class _Params:
    """keyData・encryptedKey の属性。"""

    salt: bytes
    salt_size: int
    block: int
    bits: int
    hash_size: int
    digest: str  # hashlib の名前
    hash_name: str  # 書いてあるとおりの名前


def _params(el: ET.Element, where: str) -> _Params:
    try:
        salt_size, block, bits, hash_size = (int(el.get(k) or "") for k in ("saltSize", "blockSize", "keyBits", "hashSize"))
    except ValueError as e:
        raise CryptoError(f"鍵の情報（{where}）の数が読めません（壊れています）") from e
    cipher, chaining, hname = (el.get(k) or "" for k in ("cipherAlgorithm", "cipherChaining", "hashAlgorithm"))
    if cipher.upper() != "AES" or bits not in (128, 192, 256) or block != 16:
        raise CryptoError(f"暗号の方式（{cipher}・{bits} ビット）には対応していません（AES の 128・192・256 ビットだけ）")
    if chaining != "ChainingModeCBC":
        raise CryptoError(f"暗号のつなぎ方（{chaining}）には対応していません（CBC だけ）")
    digest = HASHES.get(hname.upper().replace("-", ""))
    if digest is None:
        raise CryptoError(f"ハッシュの方式（{hname}）には対応していません（SHA-1・SHA-256・SHA-384・SHA-512・MD5）")
    if hashlib.new(digest).digest_size != hash_size:
        raise CryptoError(f"鍵の情報（{where}）のハッシュの大きさ（{hash_size}）が {hname} と合いません（壊れています）")
    return _Params(_b64(el.get("saltValue")), salt_size, block, bits, hash_size, digest, hname)


def _agile_decrypt(blob: bytes, info: bytes, stream: bytes, pw: bytes) -> tuple[bytes, OoxmlKey]:
    body = info[8:].rstrip(b"\x00\r\n\t ")
    if b"<!DOCTYPE" in body or b"<!ENTITY" in body:
        raise CryptoError("鍵の情報（EncryptionInfo）の XML に、使われないはずの宣言（DOCTYPE）があります")
    try:
        root = ET.fromstring(body)
    except ET.ParseError as e:
        raise CryptoError("鍵の情報（EncryptionInfo）の XML が読めません（壊れています）") from e
    kd_el = root.find(NS + "keyData")
    if root.tag != NS + "encryption" or kd_el is None:
        raise CryptoError("鍵の情報（EncryptionInfo）が Agile の形（encryption・keyData）になっていません")
    kd = _params(kd_el, "keyData")
    pk_el = next((ke.find(f"{{{PASSWORD_URI}}}encryptedKey") for ke in root.iterfind(f"{NS}keyEncryptors/{NS}keyEncryptor")
                  if ke.get("uri") == PASSWORD_URI), None)
    if pk_el is None:
        raise CryptoError("パスワードで開く鍵（PasswordKeyEncryptor）がありません（証明書だけで暗号化されたファイルには対応していません）")
    pk = _params(pk_el, "encryptedKey")
    if pk.digest != kd.digest:
        raise CryptoError("鍵の情報の keyData と encryptedKey で、ハッシュの方式が違います（仕様では同じ）")
    try:
        spin = int(pk_el.get("spinCount") or "")
    except ValueError as e:
        raise CryptoError("鍵の情報（spinCount）が読めません（壊れています）") from e
    if not 0 <= spin <= MAX_SPIN:
        raise CryptoError(f"ハッシュを繰り返す回数（spinCount {spin}）が仕様の範囲（0〜1000万）を外れています")

    # 2.3.4.13 を逆にたどる：パスワードから3つの鍵を導き、確かめの値を解いて比べ、中間鍵を取り出す
    hn = _spin(pk.digest, pk.salt + pw, spin)
    iv = _fit(pk.salt, pk.block)

    def password_key(block_key: bytes) -> bytes:
        return _fit(_h(pk.digest, hn + block_key), pk.bits // 8)  # 2.3.4.11

    verifier = _cbc_decrypt(password_key(BK_VERIFIER_INPUT), _b64(pk_el.get("encryptedVerifierHashInput")), iv)[:pk.salt_size]
    verifier_hash = _cbc_decrypt(password_key(BK_VERIFIER_HASH), _b64(pk_el.get("encryptedVerifierHashValue")), iv)[:pk.hash_size]
    if len(verifier) != pk.salt_size or len(verifier_hash) != pk.hash_size:
        raise CryptoError("パスワードの確かめの値（encryptedVerifierHash）が短すぎます（壊れています）")
    if not hmac.compare_digest(_h(pk.digest, verifier), verifier_hash):
        raise WrongPassword("パスワードが違います")
    data_key = _cbc_decrypt(password_key(BK_KEY), _b64(pk_el.get("encryptedKeyValue")), iv)[:kd.bits // 8]
    if len(data_key) != kd.bits // 8:
        raise CryptoError("中間鍵（encryptedKeyValue）が短すぎます（壊れています）")

    # 2.3.4.14：改ざん検知。鍵の長さは、元のファイルの HMAC が合う方に決める
    hmac_key, span, pad = None, (0, 0), b""
    di = root.find(NS + "dataIntegrity")
    if di is not None:
        raw_key = _cbc_decrypt(data_key, _b64(di.get("encryptedHmacKey")), _iv(kd.digest, kd.salt, BK_HMAC_KEY, kd.block))
        raw_value = _cbc_decrypt(data_key, _b64(di.get("encryptedHmacValue")), _iv(kd.digest, kd.salt, BK_HMAC_VALUE, kd.block))
        want = raw_value[:kd.hash_size]
        for n in dict.fromkeys((kd.hash_size, kd.salt_size)):
            if n <= len(raw_key) and hmac.compare_digest(hmac.new(raw_key[:n], stream, kd.digest).digest(), want):
                hmac_key = raw_key[:n]
                break
        else:
            raise CryptoError("改ざん検知の値（HMAC）が合いません。ファイルが壊れているか、暗号化したあとで書き換えられています")
        pad = raw_value[kd.hash_size:]
        found = list(_HMAC_VALUE.finditer(info))
        if len(found) != 1 or found[0].group(2) != (di.get("encryptedHmacValue") or "").encode("ascii", "replace"):
            raise CryptoError("鍵の情報の中で encryptedHmacValue の場所が1つに決まりません（書き戻せません）")
        span = found[0].span(2)

    size, body = _split(stream, kd.block)
    plain = _agile_crypt(Aes(data_key), kd.salt, kd.digest, kd.block, body, encrypt=False)
    key = OoxmlKey(method="agile", cipher=f"AES-{kd.bits}", hash_name=kd.hash_name, spin_count=spin, container=blob, info=info,
                   size=size, stream_len=len(stream), tail=plain[size:], data_key=data_key, digest=kd.digest,
                   hash_size=kd.hash_size, key_salt=kd.salt, block_size=kd.block, hmac_key=hmac_key, hmac_span=span, hmac_pad=pad)
    return _package(plain, size), key


def _agile_crypt(aes: Aes, salt: bytes, digest: str, block: int, data: bytes, encrypt: bool) -> bytes:
    """2.3.4.15：4096 バイトの区切りごとに、区切りの番号（4バイト）を blockKey にした初期化ベクトルで CBC。"""
    f = aes.encrypt_cbc if encrypt else aes.decrypt_cbc
    return b"".join(f(data[i:i + SEGMENT], _iv(digest, salt, (i // SEGMENT).to_bytes(4, "little"), block))
                    for i in range(0, len(data), SEGMENT))


def _with_hmac(key: OoxmlKey, stream: bytes) -> bytes:
    """2.3.4.14：新しい EncryptedPackage 全体（StreamSize を含む）の HMAC を暗号化し、encryptedHmacValue の値だけを差し替える。"""
    mac = hmac.new(key.hmac_key or b"", stream, key.digest).digest()
    enc = Aes(key.data_key).encrypt_cbc(mac + key.hmac_pad, _iv(key.digest, key.key_salt, BK_HMAC_VALUE, key.block_size))
    value = base64.b64encode(enc)
    a, b = key.hmac_span
    if len(value) != b - a:
        raise CryptoError("encryptedHmacValue の長さが元と変わります（書き戻せません）")
    return key.info[:a] + value + key.info[b:]


# ---------------------------------------------------------------------- Standard（2.3.4.5〜2.3.4.9）
def _standard_decrypt(blob: bytes, info: bytes, stream: bytes, pw: bytes) -> tuple[bytes, OoxmlKey]:
    if len(info) < 12 + 32:
        raise CryptoError("鍵の情報（EncryptionInfo）が短すぎます（壊れています）")
    copy_flags, header_size = struct.unpack_from("<II", info, 4)
    flags, _extra, alg, alg_hash, key_bits = struct.unpack_from("<5I", info, 12)
    if (flags | copy_flags) & F_EXTERNAL:
        raise CryptoError("拡張の暗号化（Extensible Encryption：ほかの会社の暗号の部品を使うもの）には対応していません")
    if flags & F_CRYPTOAPI and not flags & F_AES:
        raise CryptoError("RC4 の暗号化（CryptoAPI）には対応していません（.xlsx・.docx の AES だけ）")
    if not (flags & F_CRYPTOAPI and flags & F_AES):
        raise CryptoError(f"鍵の情報の Flags（0x{flags:02X}）が、Standard の暗号化（AES）の形ではありません")
    bits = AES_BITS.get(alg) if alg else 128  # AlgID が 0 なら Flags で決まる：AES-128
    if bits is None or key_bits not in (0, bits):
        raise CryptoError(f"暗号の方式（AlgID 0x{alg:04X}・{key_bits} ビット）には対応していません（AES の 128・192・256 ビットだけ）")
    if alg_hash not in (0, SHA1_ID):
        raise CryptoError(f"ハッシュの方式（AlgIDHash 0x{alg_hash:04X}）には対応していません（SHA-1 だけ）")
    v = 12 + header_size  # EncryptionVerifier（2.3.3）
    if len(info) < v + 40 + 32:
        raise CryptoError("鍵の情報（EncryptionVerifier）が短すぎます（壊れています）")
    salt_size, = struct.unpack_from("<I", info, v)
    verifier_hash_size, = struct.unpack_from("<I", info, v + 36)
    if salt_size != 16 or verifier_hash_size != 20:
        raise CryptoError("鍵の情報（EncryptionVerifier）の塩・ハッシュの大きさが仕様（16・20）と違います（壊れています）")
    salt = info[v + 4:v + 20]

    # 2.3.4.7：鍵を導く。2.3.4.9：確かめの値を解き、その SHA-1 が、解いた確かめのハッシュと合うか
    h = _h("sha1", _spin("sha1", salt + pw, STANDARD_SPIN) + b"\x00\x00\x00\x00")  # ブロックの番号は 0
    x1 = _h("sha1", bytes(b ^ 0x36 for b in h) + b"\x36" * (64 - len(h)))
    x2 = _h("sha1", bytes(b ^ 0x5C for b in h) + b"\x5c" * (64 - len(h)))
    data_key = (x1 + x2)[:bits // 8]
    aes = Aes(data_key)
    verifier = aes.decrypt_ecb(info[v + 20:v + 36])
    if not hmac.compare_digest(_h("sha1", verifier), aes.decrypt_ecb(info[v + 40:v + 72])[:20]):
        raise WrongPassword("パスワードが違います")

    size, body = _split(stream, 16)
    plain = aes.decrypt_ecb(body)
    key = OoxmlKey(method="standard", cipher=f"AES-{bits}", hash_name="SHA-1", spin_count=STANDARD_SPIN, container=blob,
                   info=info, size=size, stream_len=len(stream), tail=plain[size:], data_key=data_key)
    return _package(plain, size), key


# ---------------------------------------------------------------------- 暗号化し直す
def _encrypt_stream(key: OoxmlKey, package: bytes, cutoff: int) -> bytes:
    """新しい EncryptedPackage：StreamSize（8バイト）+ 暗号文。

    中身の大きさが元と同じなら、元の詰め物を使う（中身も同じなら、元とバイト単位で同じになる）。違えば0で詰める。
    4096 バイト（ミニストリームの境）より小さい EncryptedPackage は、複合ファイルの中で大きさを変えられない。
    そのときは0を暗号化して足し、元より小さければ元の長さに、大きければふつうのセクタに置ける長さにする
    （仕様 2.3.4.4 のとおり、流れは StreamSize より長くてよく、読む側は StreamSize より後ろを捨てる）。"""
    n, block = len(package), key.block_size
    if n == key.size and 8 + n + len(key.tail) == key.stream_len:
        plain = package + key.tail
    else:
        total = _round_up(n, block)
        if 8 + total < cutoff and 8 + total != key.stream_len:
            same = key.stream_len - 8
            total = same if total <= same and key.stream_len < cutoff and same % block == 0 else _round_up(cutoff - 8, block)
        plain = package + bytes(total - n)
    aes = Aes(key.data_key)
    if key.method == "agile":
        body = _agile_crypt(aes, key.key_salt, key.digest, key.block_size, plain, encrypt=True)
    else:
        body = aes.encrypt_ecb(plain)
    return struct.pack("<Q", n) + body


def _check(key: OoxmlKey, blob: bytes, package: bytes, stream: bytes, info: bytes) -> None:
    """書き出す前の確かめ：読み直して、差し替えた2つのストリームが意図どおりで、ほかは元のままで、
    鍵で解くと中身に戻り、改ざん検知が合うこと。1つでも食い違えば書き出さない。"""
    old, new = CFB(key.container), CFB(blob)
    if set(old.streams()) != set(new.streams()):
        raise CryptoError("暗号化し直したファイルの確かめで、ストリームの並びが元と変わりました。書き出しません")
    for path in old.streams():
        want = {PACKAGE.lower(): stream, INFO.lower(): info}.get(path.lower())
        if new.read(path) != (old.read(path) if want is None else want):
            raise CryptoError(f"暗号化し直したファイルの確かめで、ストリーム「{path}」が食い違いました。書き出しません")
    size, body = _split(stream, key.block_size)
    aes = Aes(key.data_key)
    if key.method == "agile":
        plain = _agile_crypt(aes, key.key_salt, key.digest, key.block_size, body, encrypt=False)
    else:
        plain = aes.decrypt_ecb(body)
    if size != len(package) or plain[:size] != package:
        raise CryptoError("暗号化し直したファイルの確かめで、解いた中身が元の中身と食い違いました。書き出しません")
    if not verify_integrity(blob, key):
        raise CryptoError("暗号化し直したファイルの確かめで、改ざん検知の値（HMAC）が合いませんでした。書き出しません")
