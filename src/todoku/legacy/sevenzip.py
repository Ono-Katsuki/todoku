"""7-Zip の書庫（.7z）を読む（読むだけ）。入れ物の形は 7-Zip の 7zFormat.txt のとおり。

- 署名の見出し（32 バイト）→ 終わりの見出し（Header か、それを縮めた・暗号にした EncodedHeader）→ StreamsInfo（パック・フォルダ・サブストリーム）と FilesInfo
- フォルダはコーダーのつながり（binder）。主な出力から逆にたどって解く。解くもの：Copy・LZMA・LZMA2（lzma）・Deflate（zlib）・Deflate64・BZip2（bz2）・
  PPMd（ppmd7）・BCJ・BCJ2・ARM・ARMT・ARM64・PPC・SPARC・IA64・Delta・Swap（branch）・AES-256＋SHA-256（aes）
- 解いた結果は、見出しの CRC で確かめる（合わなければ渡さない）。暗号のフォルダで合わないときは、パスワードの違いとして扱う
- 解いたフォルダは控える（連続圧縮：1 つのフォルダに多くのファイルが入るので、ファイルごとに解き直さない）

読まないもの：ほかの圧縮（RISC-V の変換・Zstandard など。名前を挙げて伝える）。分割した書庫（.7z.001…）は、1 つにつないでから。
"""

from __future__ import annotations

import bz2
import hashlib
import lzma
import struct
import zlib
from dataclasses import dataclass, field

from ..aes import Aes
from . import branch, ppmd7

SIGNATURE = b"7z\xbc\xaf\x27\x1c"
_ID_NAMES = {b"\x00": "Copy", b"\x03": "Delta", b"\x04": "BCJ", b"\x05": "PPC", b"\x06": "IA64", b"\x07": "ARM", b"\x08": "ARMT", b"\x09": "SPARC",
             b"\x0a": "ARM64", b"\x0b": "RISC-V", b"\x21": "LZMA2", b"\x03\x01\x01": "LZMA", b"\x03\x03\x01\x03": "BCJ", b"\x03\x03\x01\x1b": "BCJ2",
             b"\x03\x03\x02\x05": "PPC", b"\x03\x03\x04\x01": "IA64", b"\x03\x03\x05\x01": "ARM", b"\x03\x03\x07\x01": "ARMT", b"\x03\x03\x08\x05": "SPARC",
             b"\x03\x04\x01": "PPMD", b"\x04\x01\x08": "Deflate", b"\x04\x01\x09": "Deflate64", b"\x04\x02\x02": "BZip2", b"\x06\xf1\x07\x01": "AES",
             b"\x02\x03\x02": "Swap2", b"\x02\x03\x04": "Swap4"}
MAX_CYCLES = 24  # 鍵の作り直しの回数の上限（2 の何乗か）。ふつうは 19
MAX_FILES = 1_000_000


class SevenZipError(ValueError):
    pass


class EncryptedError(SevenZipError):
    """パスワードが要る・違う（archives が NeedsPassword・WrongPassword に直す）。"""

    def __init__(self, msg: str, wrong: bool = False):
        super().__init__(msg)
        self.wrong = wrong


@dataclass
class Coder:
    id: bytes
    num_in: int
    num_out: int
    props: bytes


@dataclass
class Folder:
    coders: list[Coder]
    bind: list[tuple[int, int]]  # (入力の番号, 出力の番号)
    packed: list[int]  # パックのストリームがつながる入力の番号
    sizes: list[int] = field(default_factory=list)  # 出力ごとの大きさ
    crc: int | None = None

    @property
    def main_out(self) -> int:
        bound = {o for _i, o in self.bind}
        for i in range(sum(c.num_out for c in self.coders)):
            if i not in bound:
                return i
        raise SevenZipError("フォルダの主な出力がありません")

    @property
    def size(self) -> int:
        return self.sizes[self.main_out]

    @property
    def encrypted(self) -> bool:
        return any(c.id == b"\x06\xf1\x07\x01" for c in self.coders)


@dataclass
class Entry:
    name: str
    size: int
    csize: int
    crc: int | None
    folder: int | None  # 解くフォルダ（空のファイルは None）
    offset: int  # フォルダを解いた中の位置
    is_dir: bool = False
    encrypted: bool = False


class _Reader:
    def __init__(self, b: bytes):
        self.b = b
        self.p = 0

    def byte(self) -> int:
        if self.p >= len(self.b):
            raise SevenZipError("見出しが途中で終わっています")
        v = self.b[self.p]
        self.p += 1
        return v

    def take(self, n: int) -> bytes:
        if n < 0 or self.p + n > len(self.b):
            raise SevenZipError("見出しが途中で終わっています")
        v = self.b[self.p:self.p + n]
        self.p += n
        return v

    def num(self) -> int:
        first = self.byte()
        mask = 0x80
        value = 0
        for i in range(8):
            if not first & mask:
                high = first & (mask - 1)
                value += high << (8 * i)
                return value
            value |= self.byte() << (8 * i)
            mask >>= 1
        return value

    def count(self, what: str, limit: int = MAX_FILES) -> int:
        n = self.num()
        if n > limit:
            raise SevenZipError(f"{what}の数（{n:,}）が大きすぎます（壊れています）")
        return n

    def bits(self, n: int) -> list[bool]:
        out = []
        mask = 0
        cur = 0
        for _ in range(n):
            if mask == 0:
                cur = self.byte()
                mask = 0x80
            out.append(bool(cur & mask))
            mask >>= 1
        return out

    def opt_bits(self, n: int) -> list[bool]:
        return [True] * n if self.byte() else self.bits(n)


def _crcs(r: _Reader, n: int) -> list[int | None]:
    defined = r.opt_bits(n)
    return [struct.unpack("<I", r.take(4))[0] if d else None for d in defined]


def _read_folder(r: _Reader) -> Folder:
    nc = r.count("コーダー", 32)
    if nc == 0:
        raise SevenZipError("コーダーのないフォルダです")
    coders = []
    for _ in range(nc):
        flag = r.byte()
        if flag & 0xC0:
            raise SevenZipError("コーダーの印が読めません")
        cid = r.take(flag & 0xF)
        ni, no = 1, 1
        if flag & 0x10:
            ni, no = r.count("入力", 32), r.count("出力", 32)
        props = r.take(r.count("プロパティ", 1 << 16)) if flag & 0x20 else b""
        coders.append(Coder(cid, ni, no, props))
    total_in = sum(c.num_in for c in coders)
    total_out = sum(c.num_out for c in coders)
    if total_out < 1 or total_in < total_out - 1:
        raise SevenZipError("コーダーのつながりが壊れています")
    bind = [(r.count("つなぎ", 1 << 16), r.count("つなぎ", 1 << 16)) for _ in range(total_out - 1)]
    npacked = total_in - (total_out - 1)
    if npacked == 1:
        bound = {i for i, _o in bind}
        free = [i for i in range(total_in) if i not in bound]
        if len(free) != 1:
            raise SevenZipError("コーダーのつながりが壊れています")
        packed = free
    else:
        packed = [r.count("つなぎ", 1 << 16) for _ in range(npacked)]
    return Folder(coders, bind, packed)


@dataclass
class Streams:
    pack_pos: int = 0
    pack_sizes: list[int] = field(default_factory=list)
    folders: list[Folder] = field(default_factory=list)
    sub_counts: list[int] = field(default_factory=list)
    sub_sizes: list[int] = field(default_factory=list)
    sub_crcs: list[int | None] = field(default_factory=list)


def _skip_to(r: _Reader, want: int) -> None:
    while True:
        t = r.num()
        if t == want:
            return
        if t == 0:
            raise SevenZipError("見出しに必要な項目がありません")
        r.take(r.num())


def _pack_info(r: _Reader, st: Streams) -> None:
    st.pack_pos = r.num()
    n = r.count("パックのストリーム")
    _skip_to(r, 0x09)
    st.pack_sizes = [r.num() for _ in range(n)]
    while True:
        t = r.num()
        if t == 0:
            return
        if t == 0x0A:
            _crcs(r, n)
        else:
            r.take(r.num())


def _unpack_info(r: _Reader, st: Streams) -> None:
    _skip_to(r, 0x0B)
    n = r.count("フォルダ")
    if r.byte() != 0:
        raise SevenZipError("外にある見出しの形（External）は読めません")
    st.folders = [_read_folder(r) for _ in range(n)]
    _skip_to(r, 0x0C)
    for f in st.folders:
        f.sizes = [r.num() for _ in range(sum(c.num_out for c in f.coders))]
    while True:
        t = r.num()
        if t == 0:
            return
        if t == 0x0A:
            for f, c in zip(st.folders, _crcs(r, n)):
                f.crc = c
        else:
            r.take(r.num())


def _sub_info(r: _Reader, st: Streams) -> None:
    st.sub_counts = [1] * len(st.folders)
    t = r.num()
    if t == 0x0D:
        st.sub_counts = [r.count("ファイル") for _ in st.folders]
        t = r.num()
    while t not in (0x09, 0x0A, 0x00):
        r.take(r.num())
        t = r.num()
    sizes = []
    for f, n in zip(st.folders, st.sub_counts):
        if n == 0:
            continue
        left = f.size
        if t == 0x09:
            for _ in range(n - 1):
                v = r.num()
                sizes.append(v)
                left -= v
            if left < 0:
                raise SevenZipError("ファイルの大きさの合計が、フォルダの大きさを超えています")
        elif n != 1:
            raise SevenZipError("ファイルの大きさの表がありません")
        sizes.append(left)
    if t == 0x09:
        t = r.num()
    st.sub_sizes = sizes
    need = sum(n for f, n in zip(st.folders, st.sub_counts) if not (n == 1 and f.crc is not None))
    got: list[int | None] = []
    while t != 0:
        if t == 0x0A:
            got = _crcs(r, need)
        else:
            r.take(r.num())
        t = r.num()
    crcs: list[int | None] = []
    k = 0
    for f, n in zip(st.folders, st.sub_counts):
        if n == 1 and f.crc is not None:
            crcs.append(f.crc)
        else:
            for _ in range(n):
                crcs.append(got[k] if k < len(got) else None)
                k += 1
    st.sub_crcs = crcs


def _streams_info(r: _Reader) -> Streams:
    st = Streams()
    t = r.num()
    if t == 0x06:
        _pack_info(r, st)
        t = r.num()
    if t == 0x07:
        _unpack_info(r, st)
        t = r.num()
    if t == 0x08:
        _sub_info(r, st)
        t = r.num()
    else:
        st.sub_counts = [1] * len(st.folders)
        st.sub_sizes = [f.size for f in st.folders]
        st.sub_crcs = [f.crc for f in st.folders]
    if t != 0:
        raise SevenZipError("見出しの StreamsInfo が壊れています")
    return st


def _files_info(r: _Reader, st: Streams) -> list[Entry]:
    n = r.count("ファイル")
    empty = [False] * n
    empty_file: list[bool] = []
    anti: list[bool] = []
    names: list[str] = []
    while True:
        t = r.num()
        if t == 0:
            break
        size = r.num()
        body = _Reader(r.take(size))
        if t == 0x0E:
            empty = body.bits(n)
        elif t == 0x0F:
            empty_file = body.bits(sum(empty))
        elif t == 0x10:
            anti = body.bits(sum(empty))
        elif t == 0x11:
            if body.byte() != 0:
                raise SevenZipError("外にある名前の形（External）は読めません")
            raw = body.b[body.p:]
            if len(raw) % 2:
                raise SevenZipError("名前の表が壊れています")
            pieces = raw.decode("utf-16-le", "surrogatepass").split("\0")
            names = pieces[:-1] if pieces and pieces[-1] == "" else pieces
    if names and len(names) != n:
        raise SevenZipError("名前の数がファイルの数と合いません")
    out: list[Entry] = []
    sub = 0
    fidx = 0
    left_in_folder = st.sub_counts[0] if st.sub_counts else 0
    offset = 0
    ei = 0
    pack_starts = _pack_starts(st)
    for i in range(n):
        nm = (names[i] if names else f"{i + 1}").replace("\\", "/")
        if empty[i]:
            is_file = empty_file[ei] if ei < len(empty_file) else False
            is_anti = anti[ei] if ei < len(anti) else False
            ei += 1
            if not is_anti:
                out.append(Entry(nm, 0, 0, None, None, 0, is_dir=not is_file))
            continue
        while left_in_folder == 0:
            fidx += 1
            offset = 0
            if fidx >= len(st.folders):
                raise SevenZipError("ファイルの数がフォルダの中身と合いません")
            left_in_folder = st.sub_counts[fidx]
        size = st.sub_sizes[sub]
        csize = sum(st.pack_sizes[pack_starts[fidx] + k] for k in range(len(st.folders[fidx].packed)))
        out.append(Entry(nm, size, csize if st.sub_counts[fidx] == 1 else 0, st.sub_crcs[sub], fidx, offset,
                         encrypted=st.folders[fidx].encrypted))
        offset += size
        sub += 1
        left_in_folder -= 1
    return out


def _pack_starts(st: Streams) -> list[int]:
    out = []
    k = 0
    for f in st.folders:
        out.append(k)
        k += len(f.packed)
    if k > len(st.pack_sizes):
        raise SevenZipError("パックのストリームの数が、フォルダのつながりと合いません")
    return out


# ====================================================================== 解く
_keys: dict[tuple, bytes] = {}


def _derive_key(password: str, salt: bytes, cycles: int) -> bytes:
    pw = password.encode("utf-16-le", "surrogatepass")
    if cycles == 0x3F:
        return (salt + pw + bytes(32))[:32]
    if cycles > MAX_CYCLES:
        raise SevenZipError(f"鍵を作る回数（2 の {cycles} 乗）が大きすぎます（壊れています）")
    key = (hashlib.sha256(pw + salt).digest(), cycles)
    if key in _keys:
        return _keys[key]
    h = hashlib.sha256()
    base = salt + pw
    n = 1 << cycles
    i = 0
    while i < n:  # 2^19 回：数千回ぶんずつまとめて渡す
        chunk = b"".join(base + j.to_bytes(8, "little") for j in range(i, min(i + 4096, n)))
        h.update(chunk)
        i += 4096
    d = h.digest()
    if len(_keys) > 8:
        _keys.clear()
    _keys[key] = d
    return d


def _aes(data: bytes, props: bytes, password: str | None) -> bytes:
    if password is None:
        raise EncryptedError("この 7z の書庫はパスワードで保護されています")
    if not props:
        raise SevenZipError("AES のプロパティがありません")
    b0 = props[0]
    cycles = b0 & 0x3F
    salt_n = iv_n = 0
    p = 1
    if b0 & 0xC0:
        if len(props) < 2:
            raise SevenZipError("AES のプロパティが壊れています")
        salt_n = ((b0 >> 7) & 1) + (props[1] >> 4)
        iv_n = ((b0 >> 6) & 1) + (props[1] & 0xF)
        p = 2
    if p + salt_n + iv_n > len(props) or salt_n > 16 or iv_n > 16:
        raise SevenZipError("AES のプロパティが壊れています")
    salt = props[p:p + salt_n]
    iv = (props[p + salt_n:p + salt_n + iv_n] + bytes(16))[:16]
    usable = len(data) - len(data) % 16
    return Aes(_derive_key(password, salt, cycles)).decrypt_cbc(data[:usable], iv)


def _lzma(data: bytes, props: bytes, size: int, cap: int) -> bytes:
    if len(props) < 5:
        raise SevenZipError("LZMA のプロパティが足りません")
    prop = props[0]
    if prop >= 9 * 5 * 5:
        raise SevenZipError("LZMA のプロパティが壊れています")
    lc = prop % 9
    rest = prop // 9
    lp, pb = rest % 5, rest // 5
    # 辞書はファイルの大きさより大きくなくてよい（宣言だけ大きな書庫で、メモリを取りすぎない）
    dic = max(min(int.from_bytes(props[1:5], "little"), max(size, 4096)), 4096)
    d = lzma.LZMADecompressor(format=lzma.FORMAT_RAW, filters=[{"id": lzma.FILTER_LZMA1, "dict_size": dic, "lc": lc, "lp": lp, "pb": pb}])
    out = d.decompress(data, max_length=cap + 1)
    return out


def _lzma2(data: bytes, props: bytes, size: int, cap: int) -> bytes:
    if not props:
        raise SevenZipError("LZMA2 のプロパティが足りません")
    bits = props[0] & 0x3F
    if bits > 40:
        raise SevenZipError("LZMA2 のプロパティが壊れています")
    dic = 0xFFFFFFFF if bits == 40 else (2 | (bits & 1)) << (bits // 2 + 11)
    dic = max(min(dic, max(size, 4096)), 4096)
    d = lzma.LZMADecompressor(format=lzma.FORMAT_RAW, filters=[{"id": lzma.FILTER_LZMA2, "dict_size": dic}])
    return d.decompress(data, max_length=cap + 1)


def _bzip2(data: bytes, cap: int) -> bytes:
    out = bytearray()
    src = data
    while src:
        d = bz2.BZ2Decompressor()
        piece = d.decompress(src, max_length=cap - len(out) + 1)
        out += piece
        while not d.eof and not d.needs_input:
            if len(out) > cap:
                return bytes(out)
            out += d.decompress(b"", max_length=cap - len(out) + 1)
        if not d.eof:
            break
        src = d.unused_data
        if len(out) > cap:
            break
    return bytes(out)


def _deflate(data: bytes, cap: int) -> bytes:
    d = zlib.decompressobj(-15)
    out = bytearray(d.decompress(data, cap + 1))
    while d.unconsumed_tail and len(out) <= cap:
        out += d.decompress(d.unconsumed_tail, cap - len(out) + 1)
    return bytes(out)


# Deflate64（辞書 64 KB。長さの符号 285 が 16 ビットの追加、距離の符号 30・31 が 14 ビットの追加）
_LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 3]
_LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 16]
_DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153]
_DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14]
_CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]


def _lsb_table(lengths: list[int]) -> tuple[list[int], list[int]]:
    count = [0] * 16
    for x in lengths:
        count[x] += 1
    count[0] = 0
    offs = [0] * 16
    for i in range(1, 15):
        offs[i + 1] = offs[i] + count[i]
    syms = [0] * len(lengths)
    for s, x in enumerate(lengths):
        if x:
            syms[offs[x]] = s
            offs[x] += 1
    return count, syms


# 符号の表の作り方と 1 字ずつ解く形（下の sym）は、zlib の contrib/puff/puff.c（Mark Adler）と同じ形に見える。作った記録で由来を
# 確かめられないので、派生として扱い、zlib の許諾の表示を THIRD_PARTY_NOTICES.md に置く
def deflate64(data: bytes, cap: int) -> bytes:
    out = bytearray()
    pos = 0
    acc = 0
    nbits = 0
    n = len(data)

    def bits(k: int) -> int:
        nonlocal pos, acc, nbits
        while nbits < k:
            if pos >= n:
                raise SevenZipError("Deflate64 のデータが途中で終わっています")
            acc |= data[pos] << nbits
            pos += 1
            nbits += 8
        v = acc & ((1 << k) - 1)
        acc >>= k
        nbits -= k
        return v

    def sym(tab) -> int:
        count, syms = tab
        code = first = index = 0
        for ln in range(1, 16):
            code |= bits(1)
            c = count[ln]
            if code - c < first:
                return syms[index + code - first]
            index += c
            first = (first + c) << 1
            code <<= 1
        raise SevenZipError("Deflate64 の符号が壊れています")

    fixed = None
    while True:
        final = bits(1)
        typ = bits(2)
        if typ == 0:
            acc = nbits = 0 if nbits < 8 else acc
            drop = nbits & 7
            acc >>= drop
            nbits -= drop
            size = bits(16)
            if bits(16) != size ^ 0xFFFF:
                raise SevenZipError("Deflate64 の無圧縮の長さが壊れています")
            while size and nbits:
                out.append(bits(8))
                size -= 1
            out += data[pos:pos + size]
            if pos + size > n:
                raise SevenZipError("Deflate64 のデータが途中で終わっています")
            pos += size
        elif typ == 3:
            raise SevenZipError("Deflate64 のブロックの形が壊れています")
        else:
            if typ == 1:
                if fixed is None:
                    fixed = (_lsb_table([8] * 144 + [9] * 112 + [7] * 24 + [8] * 8), _lsb_table([5] * 32))
                lit, dist = fixed
            else:
                hlit, hdist, hclen = bits(5) + 257, bits(5) + 1, bits(4) + 4
                cl = [0] * 19
                for i in range(hclen):
                    cl[_CL_ORDER[i]] = bits(3)
                cl_t = _lsb_table(cl)
                lengths: list[int] = []
                while len(lengths) < hlit + hdist:
                    s = sym(cl_t)
                    if s < 16:
                        lengths.append(s)
                    elif s == 16:
                        if not lengths:
                            raise SevenZipError("Deflate64 の符号の長さが壊れています")
                        lengths += [lengths[-1]] * (3 + bits(2))
                    elif s == 17:
                        lengths += [0] * (3 + bits(3))
                    else:
                        lengths += [0] * (11 + bits(7))
                lit, dist = _lsb_table(lengths[:hlit]), _lsb_table(lengths[hlit:hlit + hdist])
            while True:
                s = sym(lit)
                if s < 256:
                    out.append(s)
                elif s == 256:
                    break
                else:
                    s -= 257
                    if s >= len(_LEN_BASE):
                        raise SevenZipError("Deflate64 の長さの符号が壊れています")
                    ln = _LEN_BASE[s] + bits(_LEN_EXTRA[s])
                    d = sym(dist)
                    if d >= len(_DIST_BASE):
                        raise SevenZipError("Deflate64 の距離の符号が壊れています")
                    dd = _DIST_BASE[d] + bits(_DIST_EXTRA[d])
                    if dd > len(out):
                        raise SevenZipError("Deflate64 の参照が、解いた範囲の外を指しています")
                    if dd >= ln:
                        st = len(out) - dd
                        out += out[st:st + ln]
                    else:
                        st = len(out) - dd
                        for k in range(ln):
                            out.append(out[st + k])
                    if len(out) > cap:
                        return bytes(out)
        if len(out) > cap:
            return bytes(out)
        if final:
            return bytes(out)


def _apply(c: Coder, inputs: list[bytes], size: int, cap: int, password: str | None) -> bytes:
    cid = c.id
    name = _ID_NAMES.get(cid)
    if cid != b"\x03\x03\x01\x1b" and len(inputs) != 1:
        raise SevenZipError("コーダーの入力の数が合いません")
    data = inputs[0]
    if name == "Copy":
        return data
    if name == "LZMA":
        return _lzma(data, c.props, size, cap)
    if name == "LZMA2":
        return _lzma2(data, c.props, size, cap)
    if name == "PPMD":
        if size > cap:
            raise SevenZipError(f"PPMd の展開の大きさ（{size:,} バイト）が上限を超えます")
        return ppmd7.decode_7z(data, c.props, size)
    if name == "Deflate":
        return _deflate(data, cap)
    if name == "Deflate64":
        return deflate64(data, cap)
    if name == "BZip2":
        return _bzip2(data, cap)
    if name == "AES":
        return _aes(data, c.props, password)
    if name == "Delta":
        return branch.delta(data, (c.props[0] if c.props else 0) + 1)
    if name == "BCJ":
        return branch.x86(data)
    if name == "BCJ2":
        if len(inputs) != 4:
            raise SevenZipError("BCJ2 の入力の数が合いません")
        return branch.bcj2(inputs[0], inputs[1], inputs[2], inputs[3], size)
    if name == "ARM":
        return branch.arm(data)
    if name == "ARMT":
        return branch.armt(data)
    if name == "ARM64":
        return branch.arm64(data)
    if name == "PPC":
        return branch.ppc(data)
    if name == "SPARC":
        return branch.sparc(data)
    if name == "IA64":
        return branch.ia64(data)
    if name == "Swap2":
        return branch.swap(data, 2)
    if name == "Swap4":
        return branch.swap(data, 4)
    raise SevenZipError(f"圧縮の方式 {name or cid.hex()} は、とどくでは解けません。7-Zip で展開したファイルか、ZIP にしたものを開いてください")


def decode_folder(blob: bytes, base: int, st_pack_pos: int, pack_sizes: list[int], f: Folder, password: str | None, what: str = "7z") -> bytes:
    """フォルダを解く。上限は limits（展開したバイトで数える）。"""
    from ..limits import budget

    b = budget()
    cap = b.room(None)
    if f.size > cap:
        raise b.exceeded(cap, None, what)
    packs = []
    pos = base + st_pack_pos
    for sz in pack_sizes:
        if pos + sz > len(blob) or sz < 0:
            raise SevenZipError("圧縮したデータが書庫の外に出ています（途中で切れています）")
        packs.append(blob[pos:pos + sz])
        pos += sz
    in_start, out_start = [], []
    ni = no = 0
    for c in f.coders:
        in_start.append(ni)
        out_start.append(no)
        ni += c.num_in
        no += c.num_out
    bind_in = dict(f.bind)
    done: dict[int, bytes] = {}

    def out_of(o: int, depth: int) -> bytes:
        if depth > 32:
            raise SevenZipError("コーダーのつながりが輪になっています")
        if o in done:
            return done[o]
        ci = next((i for i, c in enumerate(f.coders) if out_start[i] <= o < out_start[i] + c.num_out), None)
        if ci is None:
            raise SevenZipError("コーダーのつながりが壊れています")
        c = f.coders[ci]
        if c.num_out != 1:
            raise SevenZipError("出力の複数あるコーダーは読めません")
        inputs = []
        for k in range(c.num_in):
            idx = in_start[ci] + k
            if idx in bind_in:
                inputs.append(out_of(bind_in[idx], depth + 1))
            elif idx in f.packed:
                inputs.append(packs[f.packed.index(idx)])
            else:
                raise SevenZipError("コーダーのつながりが壊れています")
        size = f.sizes[o]
        data = _apply(c, inputs, size, cap, password)
        if len(data) > cap:
            raise b.exceeded(cap, None, what)
        if len(data) < size:
            raise SevenZipError(f"解いた大きさ（{len(data):,}）が宣言（{size:,}）より小さい（途中で切れているか、壊れています）")
        data = data[:size]
        b.charge(len(data), what)
        done[o] = data
        return data

    try:
        out = out_of(f.main_out, 0)
    except (lzma.LZMAError, zlib.error, OSError, EOFError, ppmd7.PpmdError, ValueError, IndexError) as e:
        if isinstance(e, (EncryptedError,)) or type(e).__name__ == "LimitExceeded":
            raise
        if f.encrypted:
            raise EncryptedError("パスワードが違います（または書庫が壊れています）", wrong=True) from None
        if isinstance(e, SevenZipError):
            raise
        raise SevenZipError(f"解けません（{e}）") from None
    if f.crc is not None and zlib.crc32(out) & 0xFFFFFFFF != f.crc:
        if f.encrypted:
            raise EncryptedError("パスワードが違います（または書庫が壊れています）", wrong=True)
        raise SevenZipError("解いた結果が CRC と合いません（書庫が壊れています）")
    return out


@dataclass
class Archive:
    blob: bytes
    base: int
    st: Streams
    entries: list[Entry]
    header_encrypted: bool = False

    def pack_sizes_of(self, fi: int) -> tuple[int, list[int]]:
        starts = _pack_starts(self.st)
        pos = self.st.pack_pos + sum(self.st.pack_sizes[:starts[fi]])
        return pos, self.st.pack_sizes[starts[fi]:starts[fi] + len(self.st.folders[fi].packed)]

    def folder_bytes(self, fi: int, password: str | None) -> bytes:
        pos, sizes = self.pack_sizes_of(fi)
        return decode_folder(self.blob, self.base, pos, sizes, self.st.folders[fi], password)


def find_start(blob: bytes) -> int | None:
    """書庫の始まり：先頭か、自己解凍の .exe の中の、見出しの和の合う署名。"""
    def ok(p: int) -> bool:
        if p + 32 > len(blob) or blob[p:p + 6] != SIGNATURE:
            return False
        return zlib.crc32(blob[p + 12:p + 32]) & 0xFFFFFFFF == struct.unpack_from("<I", blob, p + 8)[0]

    if ok(0):
        return 0
    if blob[:6] == SIGNATURE:
        return 0  # 署名は合うが和が合わない：開くときに「壊れている」と伝える
    if blob[:2] == b"MZ":
        p = blob.find(SIGNATURE, 0, 1 << 24)
        while p != -1:
            if ok(p):
                return p
            p = blob.find(SIGNATURE, p + 1, 1 << 24)
    return None


def read_archive(blob: bytes, password: str | None) -> Archive:
    base = find_start(blob)
    if base is None:
        raise SevenZipError("7z の書庫の見出しがありません")
    if base + 32 > len(blob):
        raise SevenZipError("7z の見出しが途中で切れています")
    crc0 = struct.unpack_from("<I", blob, base + 8)[0]
    if zlib.crc32(blob[base + 12:base + 32]) & 0xFFFFFFFF != crc0:
        raise SevenZipError("7z の見出しの CRC が合いません（壊れています）")
    off, size, crc = struct.unpack_from("<QQI", blob, base + 12)
    start = base + 32
    if size == 0:
        return Archive(blob, start, Streams(), [])
    if start + off + size > len(blob):
        raise SevenZipError("7z の見出しがファイルの外にあります（途中で切れているか、分割した書庫の一部です。分割した書庫は、1 つにつないでから開いてください）")
    head = blob[start + off:start + off + size]
    if zlib.crc32(head) & 0xFFFFFFFF != crc:
        raise SevenZipError("7z の見出しの CRC が合いません（壊れています）")
    hdr_enc = False
    try:
        for _ in range(8):
            r = _Reader(head)
            t = r.num()
            if t == 0x17:  # EncodedHeader：見出しそのものが縮めてある（暗号のこともある）
                st = _streams_info(r)
                if not st.folders or not st.folders[0].sizes:
                    raise SevenZipError("見出しの StreamsInfo が壊れています")
                f = st.folders[0]
                hdr_enc = hdr_enc or f.encrypted
                head = decode_folder(blob, start, st.pack_pos, st.pack_sizes[:len(f.packed)], f, password, "7z の見出し")
                continue
            if t != 0x01:
                raise SevenZipError("7z の見出しが読めません")
            break
        else:
            raise SevenZipError("7z の見出しの入れ子が深すぎます")
        st = Streams()
        t = r.num()
        if t == 0x02:
            while True:
                pid = r.num()
                if pid == 0:
                    break
                r.take(r.num())
            t = r.num()
        if t == 0x03:
            _streams_info(r)
            t = r.num()
        if t == 0x04:
            st = _streams_info(r)
            t = r.num()
        entries = _files_info(r, st) if t == 0x05 else []
    except EncryptedError as e:
        if hdr_enc and e.wrong:
            raise EncryptedError("パスワードが違います", wrong=True) from None
        raise
    except (IndexError, struct.error, OverflowError, UnicodeDecodeError) as e:
        raise SevenZipError(f"7z の見出しが壊れています（{type(e).__name__}）") from None
    return Archive(blob, start, st, entries, hdr_enc)
