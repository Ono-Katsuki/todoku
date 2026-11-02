"""パスワードで暗号化された古い Excel（.xls、BIFF8）と Word（.doc、Word 97〜2003）を、本人から聞いたパスワードで解き、
todoku が書き込んだあとで、元と同じパスワード（同じ salt・同じ検証の値）で暗号化し直す。外部のライブラリは使わない。

    plain, key = decrypt_binary(blob, password)   # 暗号を解いた複合ファイル。XlsWorkbook(plain)・DocFile(plain) でそのまま開ける
    ...（plain を読み、書き込んで new_plain を作る）...
    out = encrypt_binary(key, new_plain)           # 元と同じパスワードで開くファイル。書き込まずに戻すと、元とバイト単位で一致

- is_encrypted_xls(blob)・is_encrypted_doc(blob)・is_encrypted_binary(blob)：パスワードで暗号化されているか
  （複合ファイルでないもの・壊れたものは False）
- decrypt_binary(blob, password=None) -> (暗号を解いたファイル, BinKey)。password が None か空のときは、Excel が自分で使う
  決まったパスワード（DEFAULT_XLS_PASSWORD。書き込みパスワードだけで保護したブックは、これで暗号化されている）を試す。
  合わなければ WrongPassword、対応していない方式・壊れた暗号の情報・暗号化されていないファイルは CryptoError（日本語の理由）
- encrypt_binary(key, plain_blob) -> bytes：暗号を解いた形のファイル（書き込んだあとのものでよい）を、同じ鍵で暗号化し直す
- new_key(kind, method, password, salt=…, verifier=…)：新しいパスワードの鍵（乱数は使わない。salt などは呼ぶ側が決める）

対応する方式（MS-OFFCRYPTO 2.3.5・2.3.6・2.3.7、MS-XLS 2.2.10・2.4.117、MS-DOC 2.2.6・2.5.1）
- .xls：XOR による難読化（FilePass の wEncryptionType=0）、RC4（版 1.1）、RC4 CryptoAPI（版 2.2・3.2・4.2）
- .doc：XOR による難読化（fObfuscated=1）、RC4（版 1.1）、RC4 CryptoAPI（版 2.2・3.2・4.2）
  それ以外（Excel 5.0/95 の BIFF5、AES の指定など）は CryptoError。

暗号化されているところ（ほかのストリームと、下に書いたバイトは、解くときも暗号化し直すときも1バイトも変えない）
- .xls：Workbook・Revision Log・_SX_DB_CUR の中のストリーム（BIFF のレコードの並び）。レコードの頭（種類と長さ）、
  BOF・FilePass・UsrExcl・FileLock・InterfaceHdr・RRDInfo・RRDHead のレコード、BoundSheet8.lbPlyPos は暗号化しない。
  RC4 は 1024 バイトごとに鍵を作り直す（ブロックの番号＝ストリームの中の位置 ÷ 1024）。暗号化しないバイトの分も、
  鍵の流れ（キーストリーム）は進める。XOR は 16 バイトの表を使う。表の何番目から使うかは仕様に書かれていないので、
  Office が作ったファイルで確かめた：レコードの中身の各バイトは、（ストリームの中の位置 ＋ そのレコードの中身の長さ）を
  16 で割った余りの番目を使う。
- .doc：WordDocument（先頭の 68 バイトを除く）・Table（先頭の lKey バイトの EncryptionHeader を除く）・Data。RC4 は
  512 バイトごとに鍵を作り直す。XOR はストリームの中の位置を 16 で割った余りの番目を使い、0 と表の値そのものは変えない。
- RC4 CryptoAPI で「文書のプロパティも暗号化」したファイルの encryption ストリーム（プロパティ・埋め込み・ActiveX の部品の
  入れ物）は、解かずにそのまま残す（todoku の読み取りは使わない）。

解いた形（plain）での、暗号の印の扱い（どちらも、暗号化していないファイルと同じに見える形）
- .xls：FilePass のレコードを、同じ位置・同じ長さの「種類 0x0000・中身がすべて 0」のレコードに置き換える。ほかのレコードの
  位置は変わらないので、BoundSheet8.lbPlyPos・INDEX・EXTSST はそのまま正しい。種類 0 は BIFF8 にないレコードで、todoku の
  読み取りと書き込みは読み飛ばす（そのまま残す）。encrypt_binary は、この印を元の FilePass に戻す。
- .doc：FibBase の fEncrypted・fObfuscated を 0、lKey を 0 にする。Table ストリームの先頭の EncryptionHeader は平文のまま
  残す（FIB のどの表もそこを指さない）。encrypt_binary は、元のビットと lKey を戻す。
  これは msoffcrypto-tool（MIT）の出力と同じ形（違うのは .doc の Table の先頭だけ：あちらはそこも「解いて」意味のない
  バイトにしている）。解いた形は、そのまま人に渡すものではない（渡すと、パスワードのかかっていないファイルになる）。
  必ず encrypt_binary を通し、is_encrypted_*() で確かめてから返す。

鍵（BinKey）はパスワードを持たない（パスワードから作ったハッシュ・表と、元の暗号の情報を持つ）。それでもファイルを解ける
ので、パスワードと同じように扱い、記録に残したり返したりしない。repr にも中身は出さない。
"""

from __future__ import annotations

import hashlib
import struct
from dataclasses import dataclass

from .cfb import CFB
from .pdf.crypto import rc4

# 書き込みパスワードだけで保護したブックを、Excel が暗号化するときのパスワード（MS-OFFCRYPTO 2.4.2.2 の注 <27>）
DEFAULT_XLS_PASSWORD = "VelvetSweatshop"


class CryptoError(ValueError):
    """暗号を解けない・かけ直せない（理由は日本語）。"""


class WrongPassword(CryptoError):
    """パスワードが違う（または、パスワードが要る）。"""


_METHODS = {"xor": "XOR による難読化", "rc4": "RC4", "cryptoapi": "RC4 CryptoAPI"}

# BIFF8 のレコードの種類（MS-XLS 2.3）
_BOF, _EOF, _FILEPASS, _BOUNDSHEET = 0x0809, 0x000A, 0x002F, 0x0085
# 暗号化しないレコード：BOF・FilePass・UsrExcl・FileLock・InterfaceHdr・RRDInfo・RRDHead（MS-XLS 2.2.10）
_XLS_PLAIN = {0x0809, 0x002F, 0x0194, 0x0195, 0x00E1, 0x0196, 0x0138}
_XLS_BLOCK = 1024

# FibBase（MS-DOC 2.5.2）
_F_ENCRYPTED, _F_WHICHTBL, _F_OBFUSCATED = 0x0100, 0x0200, 0x8000
_FIB_PLAIN = 68  # WordDocument の先頭の 68 バイトは暗号化しない
_DOC_BLOCK = 512

# EncryptionHeader.Flags（MS-OFFCRYPTO 2.3.1）
_FCRYPTOAPI, _FDOCPROPS, _FEXTERNAL, _FAES = 0x04, 0x08, 0x10, 0x20


# ====================================================================== 鍵
@dataclass(frozen=True, repr=False)
class BinKey:
    """暗号化し直すための鍵。パスワードそのものは持たない。"""

    kind: str  # "xls" / "doc"
    method: str  # "xor" / "rc4" / "cryptoapi"
    secret: bytes  # XOR：16 バイトの表。RC4：H1 の先頭 5 バイト。RC4 CryptoAPI：H0（SHA-1）
    marker: bytes  # .xls：FilePass の中身。.doc：Table ストリームの先頭の EncryptionHeader（XOR では空）
    key_bits: int = 128  # RC4 CryptoAPI の鍵の長さ（ビット）
    fib_bits: int = 0  # .doc：元の FibBase の fEncrypted・fObfuscated
    lkey: int = 0  # .doc：元の FibBase.lKey
    default_password: bool = False  # 決まったパスワード（DEFAULT_XLS_PASSWORD）で開いた

    def __repr__(self) -> str:
        return f"BinKey({self.kind}, {_METHODS[self.method]})"  # 鍵の中身は出さない

    @property
    def label(self) -> str:
        """方式の名前（「RC4 CryptoAPI」など）。"""
        return _METHODS[self.method]

    def _block_key(self, block: int) -> bytes:
        """RC4 の、ブロックの番号ごとの鍵（RC4：MS-OFFCRYPTO 2.3.6.2、RC4 CryptoAPI：2.3.5.2）。"""
        return _block_key(self.method, self.secret, self.key_bits, block)


def _block_key(method: str, secret: bytes, bits: int, block: int) -> bytes:
    b = struct.pack("<I", block)
    if method == "rc4":
        return hashlib.md5(secret + b).digest()  # Hfinal = MD5(H1 の先頭 5 バイト + ブロックの番号) の 128 ビットすべて
    h = hashlib.sha1(secret + b).digest()  # Hfinal = SHA1(H0 + ブロックの番号)
    if bits == 40:
        return h[:5] + bytes(11)  # 40 ビットのときだけ、残りの 88 ビットを 0 にした 128 ビットの鍵
    return h[: bits // 8]


def _utf16(password: str) -> bytes:
    return password.encode("utf-16-le", "surrogatepass")  # パスワードは Unicode の文字の並びとして使う


def _wrong(default: bool) -> WrongPassword:
    if default:
        return WrongPassword("パスワードで暗号化されています。開くためのパスワードを教えてください")
    return WrongPassword("パスワードが違います（開くためのパスワードと合いません）")


def _rc4_open(info: bytes, password: str, default: bool, what: str) -> tuple[str, bytes, int]:
    """RC4 の暗号の情報（RC4：MS-OFFCRYPTO 2.3.6.1、RC4 CryptoAPI：2.3.5.1）でパスワードを確かめる。
    (方式, 鍵を作る元のハッシュ, 鍵の長さ) を返す。"""
    if len(info) < 4:
        raise CryptoError(f"{what}の暗号の情報が短すぎます（壊れています）")
    major, minor = struct.unpack_from("<HH", info, 0)
    if (major, minor) == (1, 1):
        if len(info) < 52:
            raise CryptoError(f"{what}の暗号の情報が短すぎます（壊れています）")
        salt, ev, evh = info[4:20], info[20:36], info[36:52]
        h0 = hashlib.md5(_utf16(password)).digest()
        h1 = hashlib.md5((h0[:5] + salt) * 16).digest()  # 先頭 5 バイト ＋ salt の 21 バイトを 16 回くり返した 336 バイト
        secret = h1[:5]
        d = rc4(_block_key("rc4", secret, 128, 0), ev + evh)  # 検証の値とそのハッシュは、続けて解く（鍵を作り直さない）
        if hashlib.md5(d[:16]).digest() != d[16:32]:
            raise _wrong(default)
        return "rc4", secret, 128
    if major in (2, 3, 4) and minor == 2:
        if len(info) < 12:
            raise CryptoError(f"{what}の暗号の情報が短すぎます（壊れています）")
        hsize = struct.unpack_from("<I", info, 8)[0]
        if hsize < 32 or 12 + hsize + 60 > len(info):
            raise CryptoError(f"{what}の暗号の情報（EncryptionHeader）の長さが合いません（壊れています）")
        flags, _extra, alg, alg_hash, bits = struct.unpack_from("<IIIII", info, 12)
        if flags & (_FEXTERNAL | _FAES) or not flags & _FCRYPTOAPI or alg not in (0, 0x6801) or alg_hash not in (0, 0x8004):
            raise CryptoError(f"{what}の暗号の方式（AlgID 0x{alg:04X}・AlgIDHash 0x{alg_hash:04X}）には対応していません")
        bits = bits or 40  # 0 は 40 ビット
        if not (40 <= bits <= 128 and bits % 8 == 0):
            raise CryptoError(f"{what}の暗号の鍵の長さ（{bits} ビット）が仕様にありません")
        v = info[12 + hsize:]
        salt_size, salt, ev = struct.unpack_from("<I", v, 0)[0], v[4:20], v[20:36]
        hash_size, evh = struct.unpack_from("<I", v, 36)[0], v[40:60]
        if salt_size != 16 or hash_size != 20:
            raise CryptoError(f"{what}の暗号の検証の情報（EncryptionVerifier）が仕様と違います（壊れています）")
        h0 = hashlib.sha1(salt + _utf16(password)).digest()  # 繰り返さない
        d = rc4(_block_key("cryptoapi", h0, bits, 0), ev + evh)
        if hashlib.sha1(d[:16]).digest() != d[16:36]:
            raise _wrong(default)
        return "cryptoapi", h0, bits
    raise CryptoError(f"{what}の暗号の版（{major}.{minor}）には対応していません")


# ---------------------------------------------------------------------- XOR による難読化（MS-OFFCRYPTO 2.3.7）
_PAD = bytes([0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00])
_INITIAL = (0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C,
            0xA96A, 0x4EC3)
_MATRIX = (0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF,
           0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750,
           0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, 0x6F45, 0xDE8A, 0xAD35, 0x4A4B,
           0x9496, 0x390D, 0x721A, 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, 0x47D3, 0x8FA6, 0x0F6D,
           0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40,
           0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, 0x76B4,
           0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
           0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462,
           0x48C4)
_ROTL5 = bytes(((b << 5) | (b >> 3)) & 0xFF for b in range(256))
_ROTR5 = bytes(((b >> 5) | (b << 3)) & 0xFF for b in range(256))


def _xor_passwords(password: str) -> list[bytes]:
    """XOR に使う1バイトの文字の並び（15 字まで）。書くときの方法（各文字の下位のバイト、それが 0 なら上位のバイト）と、
    読むときにも試す方法（日本語の Windows の ANSI のコードページ cp932）の両方（MS-OFFCRYPTO 2.3.7.4）。"""
    raw = _utf16(password)
    units = struct.unpack(f"<{len(raw) // 2}H", raw)[:15]
    out = [bytes((u & 0xFF) or (u >> 8) for u in units)]
    try:
        ansi = password.encode("cp932")[:15]
        if ansi not in out:
            out.append(ansi)
    except UnicodeEncodeError:
        pass
    return [p for p in out if p]


def _xor_key(pw: bytes) -> int:
    """CreateXorKey_Method1（2.3.7.2）。"""
    key = _INITIAL[len(pw) - 1]
    cur = 0x68
    for ch in reversed(pw):
        for _ in range(7):
            if ch & 0x40:
                key ^= _MATRIX[cur]
            ch = (ch << 1) & 0xFF
            cur -= 1
    return key


def _xor_verifier(pw: bytes) -> int:
    """CreatePasswordVerifier_Method1（2.3.7.1）。"""
    v = 0
    for b in reversed(bytes([len(pw)]) + pw):
        v = (((v >> 14) & 1) | ((v << 1) & 0x7FFF)) ^ b
    return v ^ 0xCE4B


def _ror(b: int) -> int:
    return ((b >> 1) | (b << 7)) & 0xFF


def _xor_array(pw: bytes) -> bytes:
    """CreateXorArray_Method1（2.3.7.2）：16 バイトの表。偶数の番目は鍵の下位のバイト、奇数の番目は上位のバイトと混ぜる。"""
    key = _xor_key(pw)
    hi, lo = key >> 8, key & 0xFF
    arr = bytearray(16)
    i = len(pw)
    if i % 2 == 1:
        arr[i] = _ror(_PAD[0] ^ hi)
        i -= 1
        arr[i] = _ror(pw[-1] ^ lo)
    while i > 0:
        i -= 1
        arr[i] = _ror(pw[i] ^ hi)
        i -= 1
        arr[i] = _ror(pw[i] ^ lo)
    i, pad = 15, 15 - len(pw)
    while pad > 0:
        arr[i] = _ror(_PAD[pad] ^ hi)
        i, pad = i - 1, pad - 1
        arr[i] = _ror(_PAD[pad] ^ lo)
        i, pad = i - 1, pad - 1
    return bytes(arr)


# ====================================================================== ストリームを解く・暗号化する
def _keystream(key: BinKey, n: int, size: int) -> bytes:
    """位置 0 から n バイトぶんの RC4 の鍵の流れ。size バイトごとに、ブロックの番号を1つ増やして鍵を作り直す。"""
    return b"".join(rc4(key._block_key(b), bytes(size)) for b in range(-(-n // size)))[:n]


def _xor(a: bytes, b: bytes) -> bytes:
    n = len(a)
    return (int.from_bytes(a, "little") ^ int.from_bytes(b, "little")).to_bytes(n, "little")


def _biff(data: bytes):
    """BIFF のレコードの並び：(頭の位置, 種類, 中身の長さ)。ストリームの終わりで切れたレコードからあとは、触らない。"""
    pos, n = 0, len(data)
    while pos + 4 <= n:
        t, ln = struct.unpack_from("<HH", data, pos)
        if pos + 4 + ln > n:
            return
        yield pos, t, ln
        pos += 4 + ln


def _xls_spans(data: bytes):
    """暗号化するバイトの範囲：(始まり, 終わり, そのレコードの中身の長さ)。"""
    for pos, t, ln in _biff(data):
        if t in _XLS_PLAIN:
            continue
        a = pos + 4 + (4 if t == _BOUNDSHEET else 0)  # BoundSheet8.lbPlyPos（中身の先頭の 4 バイト）は暗号化しない
        z = pos + 4 + ln
        if a < z:
            yield a, z, ln


def _xls_stream(data: bytes, key: BinKey, encrypt: bool) -> bytes:
    """BIFF のレコードの並びのストリームを解く（encrypt=True なら暗号化する）。"""
    spans = list(_xls_spans(data))
    mask = bytearray(len(data))
    if key.method == "xor":
        # 表を 16 バイトごとにくり返した列：位置 p のバイトには、表の (p + レコードの中身の長さ) % 16 番目
        rep = key.secret * ((len(data) + 0x10000) // 16 + 1)
        for a, z, ln in spans:
            mask[a:z] = rep[a + ln:z + ln]
        out = bytearray(data)
        if encrypt:  # 5 ビット左に回してから、表と XOR（EncryptData_Method1）
            for a, z, _ln in spans:
                out[a:z] = out[a:z].translate(_ROTL5)
            return _xor(bytes(out), bytes(mask))
        out = bytearray(_xor(data, bytes(mask)))  # 表と XOR してから、5 ビット右に回す（DecryptData_Method1）
        for a, z, _ln in spans:
            out[a:z] = out[a:z].translate(_ROTR5)
        return bytes(out)
    ks = _keystream(key, len(data), _XLS_BLOCK)
    for a, z, _ln in spans:
        mask[a:z] = ks[a:z]
    return _xor(data, bytes(mask))


def _doc_stream(data: bytes, key: BinKey, skip: int) -> bytes:
    """Word のストリームを解く・暗号化する（どちらも同じ計算）。先頭の skip バイトはそのまま（鍵の流れは位置 0 から数える）。"""
    if key.method == "xor":
        out = bytearray(data)
        for i in range(16):
            k = key.secret[i]
            # 0 と、表の値そのものは変えない（XOR Data Transformation Method 2、2.3.7.6）
            table = bytes(b if b in (0, k) else b ^ k for b in range(256))
            start = skip + (i - skip) % 16  # skip 以降で、位置 % 16 が i になる最初の位置
            out[start::16] = data[start::16].translate(table)
        return bytes(out)
    ks = bytearray(_keystream(key, len(data), _DOC_BLOCK))
    ks[:skip] = bytes(len(ks[:skip]))
    return _xor(data, bytes(ks))


# ====================================================================== 複合ファイル
def _open(blob: bytes) -> CFB:
    try:
        return CFB(blob)
    except (ValueError, struct.error, IndexError, KeyError) as e:
        raise CryptoError(f"複合ファイル（.xls・.doc）として読めません（{e}）") from None


def _root(cfb: CFB, *names: str) -> str | None:
    """いちばん上の階層のストリームの、本当の名前（大文字・小文字を区別しない）。"""
    low = {n.lower() for n in names}
    return next((p for p in cfb.streams() if "/" not in p and p.lower() in low), None)


def _read(cfb: CFB, path: str) -> bytes:
    try:
        data = cfb.read(path)
    except (ValueError, struct.error, IndexError, KeyError) as e:
        raise CryptoError(f"ストリーム {path} を読めません（{e}）") from None
    if len(data) != cfb.streams()[path].size:
        raise CryptoError(f"ストリーム {path} が途中で切れています（ファイルが壊れています）")
    return data


def _overwrite(out: bytearray, cfb: CFB, path: str, data: bytes) -> None:
    """同じ大きさのストリームの中身を、元のセクタ（小さなストリームはミニストリームのセクタ）の上に書き直す。
    ほかのバイト（FAT・ディレクトリ・セクタの余り）は1バイトも変えない。"""
    e = cfb.streams()[path]
    if len(data) != e.size:
        raise CryptoError(f"ストリーム {path} の大きさが変わりました（{e.size} → {len(data)} バイト）")
    ss = cfb.ssize
    places = []
    try:
        if e.size < cfb.cutoff:
            ms = cfb.msize
            root = cfb.chain(cfb.entries[0].start)
            s = e.start
            for i in range(0, len(data), ms):
                if s >= len(cfb.minifat) or s * ms // ss >= len(root):
                    raise CryptoError(f"ストリーム {path} のつながりが壊れています")
                off = s * ms
                places.append(((root[off // ss] + 1) * ss + off % ss, i, ms))
                s = cfb.minifat[s]
        else:
            secs = cfb.chain(e.start)
            if len(secs) * ss < len(data):
                raise CryptoError(f"ストリーム {path} のつながりが壊れています")
            places = [((sec + 1) * ss, k * ss, ss) for k, sec in enumerate(secs[: -(-len(data) // ss)])]
    except CryptoError:
        raise
    except ValueError as e_:  # セクタのつながりの循環など
        raise CryptoError(f"ストリーム {path} のつながりが壊れています（{e_}）") from None
    for at, i, size in places:
        chunk = data[i:i + size]
        if at + len(chunk) > len(out):
            raise CryptoError(f"ストリーム {path} が、ファイルの終わりを越えています")
        out[at:at + len(chunk)] = chunk


# ====================================================================== 見分ける
def _workbook(cfb: CFB) -> str | None:
    return _root(cfb, "Workbook", "Book")


def _filepass(data: bytes) -> tuple[int, int] | None:
    """ブック全体の部分（最初の EOF まで）の FilePass：(頭の位置, 中身の長さ)。"""
    for k, (pos, t, ln) in enumerate(_biff(data)):
        if k == 0 and t != _BOF:
            return None
        if t == _FILEPASS:
            return pos, ln
        if t == _EOF:
            break
    return None


def is_encrypted_xls(blob: bytes) -> bool:
    """パスワードで暗号化された Excel のブック（.xls）か（FilePass のレコードがある）。"""
    try:
        cfb = CFB(blob)
        name = _workbook(cfb)
        return name is not None and _filepass(cfb.read(name)) is not None
    except Exception:  # noqa: BLE001
        return False


def is_encrypted_doc(blob: bytes) -> bool:
    """パスワードで暗号化された Word の文書（.doc）か（FibBase.fEncrypted が 1）。"""
    try:
        cfb = CFB(blob)
        name = _root(cfb, "WordDocument")
        if name is None:
            return False
        wd = cfb.read(name)
        return len(wd) >= _FIB_PLAIN and struct.unpack_from("<H", wd, 0)[0] == 0xA5EC and \
            bool(struct.unpack_from("<H", wd, 0x0A)[0] & _F_ENCRYPTED)
    except Exception:  # noqa: BLE001
        return False


def is_encrypted_binary(blob: bytes) -> bool:
    """パスワードで暗号化された .xls か .doc か（書き出す前に、暗号化し直したかを確かめるのにも使う）。"""
    return is_encrypted_xls(blob) or is_encrypted_doc(blob)


# ====================================================================== 解く
def decrypt_binary(blob: bytes, password: str | None = None) -> tuple[bytes, BinKey]:
    """暗号を解いた複合ファイルと、暗号化し直すための鍵を返す。暗号化したストリーム以外は、元のバイトのまま。"""
    cfb = _open(blob)
    if _workbook(cfb) is not None:
        return _decrypt_xls(blob, cfb, password)
    if _root(cfb, "WordDocument") is not None:
        return _decrypt_doc(blob, cfb, password)
    raise CryptoError("Excel のブック（.xls）でも、Word の文書（.doc）でもありません")


def _biff_streams(cfb: CFB, name: str) -> list[str]:
    """暗号化される BIFF のストリーム：ブック・変更履歴（Revision Log）・ピボットテーブルのキャッシュ（_SX_DB_CUR の中）。"""
    rest = [p for p in cfb.streams() if p.lower() == "revision log" or p.lower().startswith("_sx_db_cur/")]
    return [name] + sorted(rest)


def _decrypt_xls(blob: bytes, cfb: CFB, password: str | None) -> tuple[bytes, BinKey]:
    name = _workbook(cfb)
    data = _read(cfb, name)
    if len(data) < 8 or struct.unpack_from("<H", data, 0)[0] != _BOF:
        raise CryptoError("Excel のブックとして読めません")
    if struct.unpack_from("<H", data, 4)[0] != 0x0600:
        raise CryptoError("Excel 5.0/95 以前の形式（BIFF5）の暗号には対応していません")
    fp = _filepass(data)
    if fp is None:
        raise CryptoError("このブックは、パスワードで暗号化されていません")
    pos, ln = fp
    info = data[pos + 4:pos + 4 + ln]
    default = not password
    pw = password or DEFAULT_XLS_PASSWORD
    if len(info) < 2:
        raise CryptoError("Excel のブックの暗号の情報（FilePass）が短すぎます（壊れています）")
    kind = struct.unpack_from("<H", info, 0)[0]
    if kind == 0:  # XOR による難読化：XorKey と検証の値（MS-XLS 2.5.343）
        if len(info) < 6:
            raise CryptoError("Excel のブックの暗号の情報（FilePass）が短すぎます（壊れています）")
        xkey, xver = struct.unpack_from("<HH", info, 2)
        cand = next((p for p in _xor_passwords(pw) if _xor_key(p) == xkey and _xor_verifier(p) == xver), None)
        if cand is None:
            raise _wrong(default)
        key = BinKey("xls", "xor", _xor_array(cand), info, default_password=default)
    elif kind == 1:  # RC4 か RC4 CryptoAPI（続く版の番号で分かれる）
        method, secret, bits = _rc4_open(info[2:], pw, default, "Excel のブック")
        key = BinKey("xls", method, secret, info, key_bits=bits, default_password=default)
    else:
        raise CryptoError(f"この暗号の方式（FilePass の種類 {kind}）には対応していません")
    out = bytearray(blob)
    for path in _biff_streams(cfb, name):
        plain = bytearray(_xls_stream(_read(cfb, path), key, encrypt=False))
        if path == name:
            plain[pos:pos + 2] = b"\x00\x00"  # FilePass → 同じ長さの、種類 0・中身がすべて 0 のレコード（位置は変えない）
            plain[pos + 4:pos + 4 + ln] = bytes(ln)
        _overwrite(out, cfb, path, bytes(plain))
    return bytes(out), key


def _doc_parts(cfb: CFB, wd: bytes) -> tuple[str, str, str | None]:
    """(WordDocument, Table, Data) のストリームの名前。Table は FibBase.fWhichTblStm で決まる。"""
    flags = struct.unpack_from("<H", wd, 0x0A)[0]
    want = "1Table" if flags & _F_WHICHTBL else "0Table"
    table = _root(cfb, want)
    if table is None:
        raise CryptoError(f"Word の文書の表のストリーム（{want}）がありません")
    return _root(cfb, "WordDocument"), table, _root(cfb, "Data")


def _check_fib(wd: bytes) -> None:
    if len(wd) < _FIB_PLAIN or struct.unpack_from("<H", wd, 0)[0] != 0xA5EC:
        raise CryptoError("Word 97〜2003 の文書ではありません")


def _decrypt_doc(blob: bytes, cfb: CFB, password: str | None) -> tuple[bytes, BinKey]:
    wd = _read(cfb, _root(cfb, "WordDocument"))
    _check_fib(wd)
    flags = struct.unpack_from("<H", wd, 0x0A)[0]
    lkey = struct.unpack_from("<I", wd, 0x0E)[0]
    if not flags & _F_ENCRYPTED:
        raise CryptoError("この文書は、パスワードで暗号化されていません")
    wd_name, t_name, d_name = _doc_parts(cfb, wd)
    table = _read(cfb, t_name)
    if not password:
        raise _wrong(True)
    bits = flags & (_F_ENCRYPTED | _F_OBFUSCATED)
    if flags & _F_OBFUSCATED:  # XOR：lKey は検証の値（上位 16 ビットが XorKey、下位が検証の値。2.3.7.4）
        cand = next((p for p in _xor_passwords(password) if (_xor_key(p) << 16 | _xor_verifier(p)) == lkey), None)
        if cand is None:
            raise _wrong(False)
        key = BinKey("doc", "xor", _xor_array(cand), b"", fib_bits=bits, lkey=lkey)
    else:  # RC4 か RC4 CryptoAPI：lKey は Table の先頭の EncryptionHeader の長さ
        if lkey < 52 or lkey > len(table):
            raise CryptoError("Word の文書の暗号の情報（EncryptionHeader）の長さが合いません（壊れています）")
        header = table[:lkey]
        method, secret, kbits = _rc4_open(header, password, False, "Word の文書")
        key = BinKey("doc", method, secret, header, key_bits=kbits, fib_bits=bits, lkey=lkey)
    out = bytearray(blob)
    for path, skip in ((wd_name, _FIB_PLAIN), (t_name, len(key.marker)), (d_name, 0)):
        if path is None:
            continue
        plain = bytearray(_doc_stream(wd if path == wd_name else _read(cfb, path), key, skip))
        if path == wd_name:
            # 暗号化していない文書と同じ FIB にする：fEncrypted・fObfuscated を 0、lKey を 0（MS-DOC 2.5.2）
            struct.pack_into("<H", plain, 0x0A, flags & ~(_F_ENCRYPTED | _F_OBFUSCATED))
            struct.pack_into("<I", plain, 0x0E, 0)
        _overwrite(out, cfb, path, bytes(plain))
    return bytes(out), key


# ====================================================================== 暗号化し直す
def encrypt_binary(key: BinKey, plain_blob: bytes) -> bytes:
    """暗号を解いた形のファイル（decrypt_binary が返したもの、またはそれに書き込んだもの）を、鍵と同じパスワードで
    暗号化し直す。暗号の情報（salt・検証の値）は元のまま。暗号化するストリーム以外は1バイトも変えない。"""
    if not isinstance(key, BinKey):
        raise CryptoError("鍵（BinKey）ではありません")
    cfb = _open(plain_blob)
    if key.kind == "xls":
        return _encrypt_xls(key, plain_blob, cfb)
    return _encrypt_doc(key, plain_blob, cfb)


def _placeholder(data: bytes, ln: int) -> int:
    """ブック全体の部分の、FilePass を置き換えた印（種類 0・同じ長さ・中身がすべて 0）の位置。ちょうど1つでなければ断る。"""
    hits = []
    for pos, t, n in _biff(data):
        if t == _EOF:
            break
        if t == 0 and n == ln and not any(data[pos + 4:pos + 4 + n]):
            hits.append(pos)
    if len(hits) != 1:
        raise CryptoError("暗号の印（FilePass）を戻す場所が見つかりません。暗号を解いたときのファイルから作ったものか確かめてください")
    return hits[0]


def _encrypt_xls(key: BinKey, blob: bytes, cfb: CFB) -> bytes:
    name = _workbook(cfb)
    if name is None:
        raise CryptoError("鍵は Excel のブックのものですが、ファイルにブックのストリームがありません")
    data = bytearray(_read(cfb, name))
    if _filepass(data) is not None:
        raise CryptoError("このブックは、すでに暗号化されています（暗号を解いた形のファイルを渡してください）")
    pos = _placeholder(data, len(key.marker))
    data[pos:pos + 2] = struct.pack("<H", _FILEPASS)
    data[pos + 4:pos + 4 + len(key.marker)] = key.marker
    out = bytearray(blob)
    for path in _biff_streams(cfb, name):
        src = bytes(data) if path == name else _read(cfb, path)
        _overwrite(out, cfb, path, _xls_stream(src, key, encrypt=True))
    return bytes(out)


def _encrypt_doc(key: BinKey, blob: bytes, cfb: CFB) -> bytes:
    if _root(cfb, "WordDocument") is None:
        raise CryptoError("鍵は Word の文書のものですが、ファイルに WordDocument のストリームがありません")
    wd = bytearray(_read(cfb, _root(cfb, "WordDocument")))
    _check_fib(wd)
    flags = struct.unpack_from("<H", wd, 0x0A)[0]
    if flags & _F_ENCRYPTED:
        raise CryptoError("この文書は、すでに暗号化されています（暗号を解いた形のファイルを渡してください）")
    wd_name, t_name, d_name = _doc_parts(cfb, wd)
    table = _read(cfb, t_name)
    if table[:len(key.marker)] != key.marker:
        raise CryptoError("Table ストリームの先頭の暗号の情報（EncryptionHeader）が、暗号を解いたときと違います")
    struct.pack_into("<H", wd, 0x0A, flags | key.fib_bits)
    struct.pack_into("<I", wd, 0x0E, key.lkey)
    out = bytearray(blob)
    for path, skip in ((wd_name, _FIB_PLAIN), (t_name, len(key.marker)), (d_name, 0)):
        if path is None:
            continue
        src = bytes(wd) if path == wd_name else (table if path == t_name else _read(cfb, path))
        _overwrite(out, cfb, path, _doc_stream(src, key, skip))
    return bytes(out)


# ====================================================================== 新しい鍵
def new_key(kind: str, method: str, password: str, *, salt: bytes = b"", verifier: bytes = b"",
            key_bits: int = 128) -> BinKey:
    """新しいパスワードの鍵と、暗号の情報（.xls の FilePass の中身、.doc の EncryptionHeader）を作る。
    salt と verifier（RC4 のときに要る。どちらも 16 バイト）は呼ぶ側が決める：同じ入力からは同じ鍵。
    暗号化していないファイルにパスワードをかけるには、ファイルの側に印の場所（.xls は種類 0 の同じ長さのレコード、
    .doc は Table の先頭の lKey バイト）を作る必要があり、それはここでは扱わない。"""
    if kind not in ("xls", "doc") or method not in _METHODS:
        raise CryptoError(f"鍵の種類（{kind}・{method}）が違います")
    if method == "xor":
        pws = _xor_passwords(password)
        if not pws:
            raise CryptoError("パスワードが空です")
        pw = pws[0]
        xkey, xver = _xor_key(pw), _xor_verifier(pw)
        if kind == "xls":
            return BinKey("xls", "xor", _xor_array(pw), struct.pack("<HHH", 0, xkey, xver))
        return BinKey("doc", "xor", _xor_array(pw), b"", fib_bits=_F_ENCRYPTED | _F_OBFUSCATED, lkey=xkey << 16 | xver)
    if len(salt) != 16 or len(verifier) != 16:
        raise CryptoError("salt と検証の値は、どちらも 16 バイトにしてください")
    if method == "rc4":
        h0 = hashlib.md5(_utf16(password)).digest()
        secret = hashlib.md5((h0[:5] + salt) * 16).digest()[:5]
        enc = rc4(_block_key("rc4", secret, 128, 0), verifier + hashlib.md5(verifier).digest())
        header = struct.pack("<HH", 1, 1) + salt + enc
        bits = 128
    else:
        if not (40 <= key_bits <= 128 and key_bits % 8 == 0):
            raise CryptoError(f"鍵の長さ（{key_bits} ビット）が仕様にありません")
        secret = hashlib.sha1(salt + _utf16(password)).digest()
        enc = rc4(_block_key("cryptoapi", secret, key_bits, 0), verifier + hashlib.sha1(verifier).digest())
        # 56 ビットより長い鍵は Enhanced の CSP（Base は 56 ビットまで。MS-OFFCRYPTO 2.3.5.1 の注 <24>）
        csp = "Microsoft Enhanced Cryptographic Provider v1.0" if key_bits > 56 else "Microsoft Base Cryptographic Provider v1.0"
        flags = _FCRYPTOAPI | _FDOCPROPS  # 文書のプロパティは暗号化しない（encryption ストリームを作らない）
        eh = struct.pack("<8I", flags, 0, 0x6801, 0x8004, key_bits, 1, 0, 0) + (csp + "\0").encode("utf-16-le")
        header = (struct.pack("<HHII", 4, 2, flags, len(eh)) + eh + struct.pack("<I", 16) + salt + enc[:16]
                  + struct.pack("<I", 20) + enc[16:36])
        bits = key_bits
    if kind == "xls":
        return BinKey("xls", method, secret, struct.pack("<H", 1) + header, key_bits=bits)
    return BinKey("doc", method, secret, header, key_bits=bits, fib_bits=_F_ENCRYPTED, lkey=len(header))
