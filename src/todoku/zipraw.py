"""ZIPをバイト単位で忠実に読み書きする。

xlsx / xlsm / pptx は XML を ZIP でまとめたもの。一般的なライブラリで保存し直すと、
圧縮率・タイムスタンプ・エントリ順・余分なフィールドが変わり、元とバイト一致しなくなる。

ここでは ZIP を「ローカルヘッダ・圧縮データ・データ記述子・中央ディレクトリ・終端レコード」
というバイト列の並びとしてそのまま保持する。変更しないエントリは圧縮済みのバイトを
そのままコピーし、変更するエントリだけを圧縮し直して、ヘッダのCRC・サイズ・オフセットを直す。
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field

LOCAL_SIG = b"PK\x03\x04"
CENTRAL_SIG = b"PK\x01\x02"
EOCD_SIG = b"PK\x05\x06"
ZIP64_EOCD_SIG = b"PK\x06\x06"
DD_SIG = b"PK\x07\x08"


class ZipFormatError(ValueError):
    pass


class AmbiguousZip(ZipFormatError):
    """ZIP の中身が二通りに読める（同じ名前・大文字小文字だけ違う名前の部品、中央ディレクトリとローカルヘッダで違う名前、
    重なった部品）。読む道具によって違う中身が出てくるので、読まない・書かない（証明できない）。"""

    def __init__(self, detail: str):
        super().__init__(f"ZIP の中身が二通りに読めます（{detail}）。読む道具によって違う中身になるので、読み書きしません")
        self.detail = detail


@dataclass
class Entry:
    name: str
    # 中央ディレクトリのレコード（固定部 + 名前 + extra + コメント）
    central: bytes
    # ローカルヘッダ（固定部 + 名前 + extra）
    local: bytes
    # 圧縮されたデータ
    data: bytes
    # データ記述子（なければ b""）
    descriptor: bytes
    # このエントリの直後、次のエントリまでの未知のバイト（通常は空）
    gap: bytes = b""
    # 書き換え後の非圧縮データ（Noneなら元のまま）
    _replacement: bytes | None = field(default=None, repr=False)

    @property
    def method(self) -> int:
        return struct.unpack_from("<H", self.central, 10)[0]

    @property
    def flags(self) -> int:
        return struct.unpack_from("<H", self.central, 8)[0]

    @property
    def crc(self) -> int:
        return struct.unpack_from("<I", self.central, 16)[0]

    def read(self) -> bytes:
        """非圧縮のデータを返す。"""
        if self._replacement is not None:
            return self._replacement
        if self.flags & 0x1:
            raise ZipFormatError(f"{self.name}: 暗号化されたエントリは読めません")
        if self.method == 0:
            raw = self.data
        elif self.method == 8:
            # 宣言された大きさは信じず、出てくるバイトを数えながら展開する（上限：limits）
            from .limits import inflate

            raw = inflate(self.data, -15, what=self.name)
        else:
            raise ZipFormatError(f"{self.name}: 未対応の圧縮方式 {self.method}")
        if zlib.crc32(raw) & 0xFFFFFFFF != self.crc:
            raise ZipFormatError(f"{self.name}: CRCが一致しません")
        return raw

    @property
    def modified(self) -> bool:
        return self._replacement is not None


class RawZip:
    """バイト列を保ったまま扱える ZIP。"""

    def __init__(self, blob: bytes):
        self.original = blob
        self.prefix = b""
        self.entries: list[Entry] = []
        self.cd_tail = b""  # 中央ディレクトリの後ろ（EOCDとコメント）
        self._parse(blob)

    # ------------------------------------------------------------------ 読み込み
    def _parse(self, blob: bytes) -> None:
        eocd_pos = blob.rfind(EOCD_SIG, max(0, len(blob) - 65557))
        if eocd_pos < 0:
            raise ZipFormatError("ZIPの終端レコードが見つかりません")
        if blob.rfind(ZIP64_EOCD_SIG, 0, eocd_pos) >= 0 and eocd_pos >= 20 and blob[eocd_pos - 20:eocd_pos - 16] == b"PK\x06\x07":
            raise ZipFormatError("ZIP64 には未対応です")
        (_, disk, cd_disk, n_disk, n_total, cd_size, cd_offset, comment_len) = struct.unpack_from(
            "<4sHHHHIIH", blob, eocd_pos
        )
        if disk != 0 or cd_disk != 0 or n_disk != n_total:
            raise ZipFormatError("分割ZIPには未対応です")
        self.cd_tail = blob[eocd_pos:]
        self._cd_gap = blob[cd_offset + cd_size:eocd_pos]

        # 中央ディレクトリ
        centrals: list[tuple[str, bytes, int, int]] = []
        p = cd_offset
        for _ in range(n_total):
            if blob[p:p + 4] != CENTRAL_SIG:
                raise ZipFormatError("中央ディレクトリが壊れています")
            (flags, method, csize, usize, nlen, xlen, clen, local_off) = (
                struct.unpack_from("<H", blob, p + 8)[0],
                struct.unpack_from("<H", blob, p + 10)[0],
                struct.unpack_from("<I", blob, p + 20)[0],
                struct.unpack_from("<I", blob, p + 24)[0],
                struct.unpack_from("<H", blob, p + 28)[0],
                struct.unpack_from("<H", blob, p + 30)[0],
                struct.unpack_from("<H", blob, p + 32)[0],
                struct.unpack_from("<I", blob, p + 42)[0],
            )
            if 0xFFFFFFFF in (csize, usize, local_off):
                raise ZipFormatError("ZIP64 には未対応です")
            raw_name = blob[p + 46:p + 46 + nlen]
            name = raw_name.decode("utf-8" if flags & 0x800 else "cp437")
            rec_len = 46 + nlen + xlen + clen
            centrals.append((name, blob[p:p + rec_len], local_off, csize))
            p += rec_len

        _no_two_readings(blob, centrals, cd_offset)
        # ローカルヘッダの順に並べる（中央ディレクトリの順とは限らない）
        order = sorted(range(len(centrals)), key=lambda i: centrals[i][2])
        self._cd_order = order  # 中央ディレクトリの順番→エントリ番号
        first_off = centrals[order[0]][2] if centrals else cd_offset
        self.prefix = blob[:first_off]
        entries_by_cd: dict[int, Entry] = {}
        sorted_entries: list[Entry] = []
        for rank, i in enumerate(order):
            name, central, off, csize = centrals[i]
            if blob[off:off + 4] != LOCAL_SIG:
                raise ZipFormatError(f"{name}: ローカルヘッダが見つかりません")
            nlen, xlen = struct.unpack_from("<HH", blob, off + 26)
            hdr_end = off + 30 + nlen + xlen
            data_end = hdr_end + csize
            flags = struct.unpack_from("<H", blob, off + 6)[0]
            desc = b""
            if flags & 0x8:
                if blob[data_end:data_end + 4] == DD_SIG:
                    desc = blob[data_end:data_end + 16]
                else:
                    desc = blob[data_end:data_end + 12]
            end = data_end + len(desc)
            nxt = centrals[order[rank + 1]][2] if rank + 1 < len(order) else cd_offset
            e = Entry(
                name=name,
                central=central,
                local=blob[off:hdr_end],
                data=blob[hdr_end:data_end],
                descriptor=desc,
                gap=blob[end:nxt],
            )
            entries_by_cd[i] = e
            sorted_entries.append(e)
        self.entries = sorted_entries
        self._entries_cd_order = [entries_by_cd[i] for i in range(len(centrals))]

    # ------------------------------------------------------------------ 参照
    def names(self) -> list[str]:
        return [e.name for e in self._entries_cd_order]

    def get(self, name: str) -> Entry:
        # 名前の重なる ZIP は読まない（AmbiguousZip）ので、名前 → エントリの表で引ける（部品が多くても 1 回で）
        idx = getattr(self, "_by_name", None)
        if idx is None or len(idx) != len(self.entries):
            idx = self._by_name = {}
            for e in self.entries:
                idx.setdefault(e.name, e)
        e = idx.get(name)
        if e is None:
            raise KeyError(name)
        return e

    def has(self, name: str) -> bool:
        try:
            self.get(name)
        except KeyError:
            return False
        return True

    def read(self, name: str) -> bytes:
        return self.get(name).read()

    # ------------------------------------------------------------------ 書き換え
    def replace(self, name: str, content: bytes) -> None:
        e = self.get(name)
        if e.read() == content:
            e._replacement = None
            return
        e._replacement = content

    def modified_names(self) -> list[str]:
        return [e.name for e in self.entries if e.modified]

    def add(self, name: str, content: bytes, like: str) -> None:
        """新しいエントリを最後に足す。ヘッダーは既存のエントリ like のもの（時刻・圧縮方式など）を写し、名前と大きさだけを替える。
        ほかのエントリのバイトは変えない（位置は後ろにずれない：最後に足すので）。"""
        if self.has(name):
            raise ZipFormatError(f"{name}: 同じ名前のエントリがあります")
        src = self.get(like)
        raw_name = name.encode("utf-8")
        flags = (struct.unpack_from("<H", src.local, 6)[0] & ~0x8) | 0x800  # データ記述子を使わない・名前は UTF-8
        nlen, xlen = struct.unpack_from("<HH", src.local, 26)
        local = bytearray(src.local[:30]) + raw_name + src.local[30 + nlen:30 + nlen + xlen]
        struct.pack_into("<HH", local, 6, flags, src.method)
        struct.pack_into("<HH", local, 26, len(raw_name), xlen)
        cn, cx, _cc = struct.unpack_from("<HHH", src.central, 28)
        central = bytearray(src.central[:46]) + raw_name + src.central[46 + cn:46 + cn + cx]
        struct.pack_into("<H", central, 8, flags)
        struct.pack_into("<HHH", central, 28, len(raw_name), cx, 0)
        e = Entry(name=name, central=bytes(central), local=bytes(local), data=b"", descriptor=b"")
        e._replacement = content
        self.entries.append(e)
        self._entries_cd_order.append(e)
        self.added = getattr(self, "added", []) + [name]

    def drop_added(self) -> None:
        """add で足したエントリを取り除く（元のエントリだけに戻す）。"""
        gone = set(getattr(self, "added", []))
        self.entries = [e for e in self.entries if e.name not in gone]
        self._by_name = None
        self._entries_cd_order = [e for e in self._entries_cd_order if e.name not in gone]
        self.added = []

    # ------------------------------------------------------------------ 書き出し
    def build(self) -> bytes:
        out = bytearray(self.prefix)
        new_offsets: dict[int, int] = {}
        new_centrals: dict[int, bytes] = {}
        for e in self.entries:
            offset = len(out)
            if e.modified:
                local, data, desc, central = self._rebuild_entry(e, offset)
            else:
                local, data, desc = e.local, e.data, e.descriptor
                central = _set_u32(e.central, 42, offset)
            new_offsets[id(e)] = offset
            new_centrals[id(e)] = central
            out += local + data + desc + e.gap
        cd_start = len(out)
        for e in self._entries_cd_order:
            out += new_centrals[id(e)]
        cd_size = len(out) - cd_start
        out += self._cd_gap
        tail = bytearray(self.cd_tail)
        if getattr(self, "added", None):
            struct.pack_into("<HH", tail, 8, len(self._entries_cd_order), len(self._entries_cd_order))
        struct.pack_into("<I", tail, 12, cd_size)
        struct.pack_into("<I", tail, 16, cd_start)
        out += tail
        return bytes(out)

    def _rebuild_entry(self, e: Entry, offset: int) -> tuple[bytes, bytes, bytes, bytes]:
        raw = e._replacement
        assert raw is not None
        crc = zlib.crc32(raw) & 0xFFFFFFFF
        if e.method == 8:
            # 元の圧縮レベルは分からないので既定のレベルで圧縮する（このエントリだけが変わる）
            c = zlib.compressobj(zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15)
            data = c.compress(raw) + c.flush()
        elif e.method == 0:
            data = raw
        else:
            raise ZipFormatError(f"{e.name}: 未対応の圧縮方式 {e.method}")
        local = bytearray(e.local)
        flags = struct.unpack_from("<H", local, 6)[0]
        if flags & 0x8:
            # データ記述子を使う形式はそのまま保つ
            desc = bytearray(e.descriptor)
            base = 4 if desc[:4] == DD_SIG else 0
            struct.pack_into("<III", desc, base, crc, len(data), len(raw))
            desc = bytes(desc)
        else:
            struct.pack_into("<III", local, 14, crc, len(data), len(raw))
            desc = b""
        central = bytearray(e.central)
        struct.pack_into("<III", central, 16, crc, len(data), len(raw))
        struct.pack_into("<I", central, 42, offset)
        return bytes(local), data, desc, bytes(central)


def _no_two_readings(blob: bytes, centrals: list[tuple[str, bytes, int, int]], cd_offset: int) -> None:
    """中央ディレクトリとローカルヘッダが、ただ一通りの中身を指すか（そうでなければ AmbiguousZip）：
    名前が重ならない（大文字小文字を除いても。OOXML の部品の名前は大文字小文字を区別しない：ECMA-376 第2部 6.2.2.3）、
    ローカルヘッダの名前が中央ディレクトリの名前と同じ、部品のバイトの範囲が重ならず、中央ディレクトリより前にある。"""
    seen: dict[str, str] = {}
    spans = []
    for name, central, off, csize in centrals:
        key = name.casefold()
        if key in seen:
            what = "同じ名前" if seen[key] == name else "大文字・小文字だけが違う名前"
            raise AmbiguousZip(f"{what}の部品が 2 つあります：{seen[key]!r}・{name!r}")
        seen[key] = name
        if blob[off:off + 4] != LOCAL_SIG or off + 30 > len(blob):
            raise ZipFormatError(f"{name}: ローカルヘッダが見つかりません")
        nlen, xlen = struct.unpack_from("<HH", blob, off + 26)
        cn = struct.unpack_from("<H", central, 28)[0]
        if blob[off + 30:off + 30 + nlen] != central[46:46 + cn]:
            raise AmbiguousZip(f"部品 {name!r} の名前が、中央ディレクトリとローカルヘッダで違います")
        end = off + 30 + nlen + xlen + csize
        if struct.unpack_from("<H", blob, off + 6)[0] & 0x8:
            end += 12  # データ記述子（少なくとも 12 バイト）
        spans.append((off, end, name))
    spans.sort()
    for (a0, a1, an), (b0, _b1, bn) in zip(spans, spans[1:]):
        if b0 < a1:
            raise AmbiguousZip(f"部品 {an!r} と {bn!r} のバイトが重なっています")
    if spans and spans[-1][1] > cd_offset:
        raise AmbiguousZip(f"部品 {spans[-1][2]!r} のバイトが、中央ディレクトリに重なっています")


def _set_u32(b: bytes, pos: int, value: int) -> bytes:
    if struct.unpack_from("<I", b, pos)[0] == value:
        return b
    ba = bytearray(b)
    struct.pack_into("<I", ba, pos, value)
    return bytes(ba)
