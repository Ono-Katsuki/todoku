"""複合ファイル（Compound File Binary：古いExcelの .xls、Wordの .doc、マクロの vbaProject.bin）を、
バイト単位で読み書きする（MS-CFB）。

書き換えたいストリームのセクタだけを差し替え、それ以外のバイトは1バイトも変えない。
ストリームが大きくなるときは、ファイルの末尾にセクタを足してつなぐ。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF


class CfbError(ValueError):
    pass


@dataclass
class DirEntry:
    index: int
    name: str
    kind: int  # 1 ストレージ / 2 ストリーム / 5 ルート
    left: int
    right: int
    child: int
    start: int
    size: int
    path: str = ""


class CFB:
    def __init__(self, blob: bytes):
        if blob[:8] != MAGIC:
            raise CfbError("複合ファイルではありません")
        self.blob = blob
        (self.minor, self.major, order, sshift, mshift) = struct.unpack_from("<HHHHH", blob, 0x18)
        if order != 0xFFFE:
            raise CfbError("バイト順が想定外です")
        self.ssize = 1 << sshift
        self.msize = 1 << mshift
        (self.n_dir_sectors, self.n_fat, self.dir_start, _txn, self.cutoff, self.minifat_start, self.n_minifat,
         self.difat_start, self.n_difat) = struct.unpack_from("<IIIIIIIII", blob, 0x28)
        difat = list(struct.unpack_from("<109I", blob, 0x4C))
        sec = self.difat_start
        per = self.ssize // 4 - 1
        guard = 0
        while sec not in (ENDOFCHAIN, FREESECT) and guard < 100000:
            vals = struct.unpack_from(f"<{per + 1}I", blob, self._off(sec))
            difat += vals[:per]
            sec = vals[per]
            guard += 1
        self.fat_sectors = [s for s in difat if s not in (FREESECT, ENDOFCHAIN)][: self.n_fat]
        fat: list[int] = []
        for s in self.fat_sectors:
            fat += struct.unpack_from(f"<{self.ssize // 4}I", blob, self._off(s))
        self.fat = fat
        dir_bytes = self._chain_bytes(self.dir_start)
        self.entries: list[DirEntry] = []
        for i in range(len(dir_bytes) // 128):
            e = dir_bytes[i * 128:(i + 1) * 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16-le", errors="replace")
            kind = e[66]
            left, right, child = struct.unpack_from("<III", e, 68)
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<Q", e, 120)[0]
            if self.major == 3:
                size &= 0xFFFFFFFF
            self.entries.append(DirEntry(i, name, kind, left, right, child, start, size))
        self.minifat: list[int] = []
        if self.n_minifat and self.minifat_start not in (ENDOFCHAIN, FREESECT):
            mb = self._chain_bytes(self.minifat_start)
            self.minifat = list(struct.unpack_from(f"<{len(mb) // 4}I", mb, 0))
        root = self.entries[0] if self.entries else None
        self.ministream = self._chain_bytes(root.start)[: root.size] if root and root.start != ENDOFCHAIN else b""
        self._paths()

    # ------------------------------------------------------------------ 読む
    def _off(self, sector: int) -> int:
        return (sector + 1) * self.ssize

    def chain(self, start: int) -> list[int]:
        out = []
        s = start
        seen = set()
        while s not in (ENDOFCHAIN, FREESECT) and s < len(self.fat):
            if s in seen:
                raise CfbError("セクタのつながりが循環しています")
            seen.add(s)
            out.append(s)
            s = self.fat[s]
        return out

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self.blob[self._off(s):self._off(s) + self.ssize] for s in self.chain(start))

    def _paths(self) -> None:
        def walk(idx: int, prefix: str, depth: int = 0):
            if idx == NOSTREAM or idx >= len(self.entries) or depth > 512:
                return
            e = self.entries[idx]
            walk(e.left, prefix, depth + 1)
            e.path = prefix + e.name
            if e.kind == 1:
                walk(e.child, e.path + "/", depth + 1)
            walk(e.right, prefix, depth + 1)

        if self.entries:
            walk(self.entries[0].child, "")

    def streams(self) -> dict[str, DirEntry]:
        return {e.path: e for e in self.entries if e.kind == 2 and e.path}

    def has(self, path: str) -> bool:
        return self._find(path) is not None

    def _find(self, path: str) -> DirEntry | None:
        low = path.lower()
        for p, e in self.streams().items():
            if p.lower() == low:
                return e
        return None

    def read(self, path: str) -> bytes:
        e = self._find(path)
        if e is None:
            raise KeyError(path)
        if e.size < self.cutoff:
            out = bytearray()
            s = e.start
            while s not in (ENDOFCHAIN, FREESECT) and s < len(self.minifat) and len(out) < e.size:
                out += self.ministream[s * self.msize:(s + 1) * self.msize]
                s = self.minifat[s]
            return bytes(out[: e.size])
        return self._chain_bytes(e.start)[: e.size]

    # ------------------------------------------------------------------ 書く
    def replace(self, path: str, data: bytes, pad_ok: bool = False, wipe: bool = False) -> bytes:
        """1つのストリームの中身を差し替えたファイルを返す。ほかのバイトは変えない。

        pad_ok=True（後ろに0を足してよいストリーム。Excelのブックは最後の EOF の後ろを読まない）のとき、
        小さすぎるストリームは区切りの大きさまで0を足して、ふつうのセクタに置く。
        wipe=True のとき、ストリームが縮んだり移ったりして使わなくなった元のデータ（空いたセクタ・ミニセクタと、
        最後のセクタの新しい終わりから元の終わりまで）を0で消す。暗号化したファイルに古い暗号文を残さないため。"""
        e = self._find(path)
        if e is None:
            raise KeyError(path)
        if len(data) < self.cutoff:
            if e.size < self.cutoff and len(data) == e.size:
                return self._replace_mini(e, data)
            if not pad_ok:
                raise CfbError("小さなストリーム（4096バイト未満）の大きさを変えるには、後ろに0を足せるストリームである必要があります")
            data = data + b"\x00" * (self.cutoff - len(data))
        out = bytearray(self.blob)
        fat = list(self.fat)
        per = self.ssize // 4
        fat_sectors = list(self.fat_sectors)
        n_total = (len(out) - self.ssize + self.ssize - 1) // self.ssize
        if len(out) < (n_total + 1) * self.ssize:
            out += b"\x00" * ((n_total + 1) * self.ssize - len(out))

        def alloc() -> int:
            nonlocal n_total
            idx = n_total
            n_total += 1
            out.extend(b"\x00" * self.ssize)
            while len(fat) < n_total:
                # FAT のセクタも足りない：FAT のセクタを末尾に足す
                fidx = n_total
                n_total += 1
                out.extend(b"\x00" * self.ssize)
                fat.extend([FREESECT] * per)
                fat[fidx] = FATSECT
                fat_sectors.append(fidx)
            if len(fat) <= idx:
                fat.extend([FREESECT] * per)
            return idx

        old_regular = e.size >= self.cutoff
        sectors = self.chain(e.start) if old_regular else []
        old_sectors = list(sectors)
        freed_mini: list[int] = []
        if not old_regular:
            # もとはミニストリームにあった：ミニのセクタを空きにして、ふつうのセクタに移す
            s_ = e.start
            seen = 0
            while s_ not in (ENDOFCHAIN, FREESECT) and s_ < len(self.minifat) and seen < 1 << 20:
                nxt = self.minifat[s_]
                self.minifat[s_] = FREESECT
                freed_mini.append(s_)
                s_ = nxt
                seen += 1
        need = -(-len(data) // self.ssize)
        while len(sectors) < need:
            sectors.append(alloc())
        for s_ in sectors[need:]:
            fat[s_] = FREESECT
        sectors = sectors[:need]
        for a, b in zip(sectors, sectors[1:]):
            fat[a] = b
        fat[sectors[-1]] = ENDOFCHAIN
        for i, s_ in enumerate(sectors):
            chunk = data[i * self.ssize:(i + 1) * self.ssize]
            off = self._off(s_)
            out[off:off + len(chunk)] = chunk
        if wipe:
            for k, s_ in enumerate(old_sectors):
                off = self._off(s_)
                lo = 0 if k >= need else max(len(data) - k * self.ssize, 0)
                hi = self.ssize if k >= need else min(e.size - k * self.ssize, self.ssize)
                if lo < hi:
                    out[off + lo:off + hi] = bytes(hi - lo)
            if freed_mini:
                ms = bytearray(self.ministream)
                for m in freed_mini:
                    sl = slice(m * self.msize, (m + 1) * self.msize)
                    ms[sl] = bytes(len(ms[sl]))
                for k, sec in enumerate(self.chain(self.entries[0].start)):
                    chunk = bytes(ms[k * self.ssize:(k + 1) * self.ssize])
                    out[self._off(sec):self._off(sec) + len(chunk)] = chunk
        # DIFAT（FAT のセクタの一覧）：先頭の 109 個はヘッダー、残りは DIFAT のセクタに書く
        head = fat_sectors[:109]
        rest = fat_sectors[109:]
        difat_secs = self.chain_difat()
        need_difat = -(-len(rest) // (per - 1)) if rest else 0
        while len(difat_secs) < need_difat:
            d = alloc()
            fat[d] = DIFSECT
            difat_secs.append(d)
            head = fat_sectors[:109]
            rest = fat_sectors[109:]
            need_difat = -(-len(rest) // (per - 1)) if rest else 0
        for k, d in enumerate(difat_secs[:need_difat]):
            vals = rest[k * (per - 1):(k + 1) * (per - 1)]
            vals += [FREESECT] * (per - 1 - len(vals))
            nxt = difat_secs[k + 1] if k + 1 < need_difat else ENDOFCHAIN
            struct.pack_into(f"<{per}I", out, self._off(d), *vals, nxt)
        # FAT を書き戻す
        for k, s_ in enumerate(fat_sectors):
            part = fat[k * per:(k + 1) * per]
            part += [FREESECT] * (per - len(part))
            struct.pack_into(f"<{per}I", out, self._off(s_), *part)
        struct.pack_into("<I", out, 0x2C, len(fat_sectors))
        struct.pack_into("<109I", out, 0x4C, *(head + [FREESECT] * (109 - len(head))))
        if need_difat:
            struct.pack_into("<II", out, 0x44, difat_secs[0], need_difat)
        # ミニ FAT（ミニのセクタを空けたとき）
        if not old_regular and self.minifat:
            mf = list(struct.pack(f"<{len(self.minifat)}I", *self.minifat))
            mb = bytes(mf)
            for k, s_ in enumerate(self.chain(self.minifat_start)):
                chunk = mb[k * self.ssize:(k + 1) * self.ssize]
                out[self._off(s_):self._off(s_) + len(chunk)] = chunk
        # ディレクトリ：始まりのセクタと大きさ
        dsecs = self.chain(self.dir_start)
        dper = self.ssize // 128
        sec = dsecs[e.index // dper]
        base = self._off(sec) + (e.index % dper) * 128
        struct.pack_into("<I", out, base + 116, sectors[0])
        struct.pack_into("<Q" if self.major == 4 else "<I", out, base + 120, len(data))
        return bytes(out)

    def chain_difat(self) -> list[int]:
        out = []
        s_ = self.difat_start
        per = self.ssize // 4
        while s_ not in (ENDOFCHAIN, FREESECT) and len(out) < 100000:
            out.append(s_)
            s_ = struct.unpack_from("<I", self.blob, self._off(s_) + (per - 1) * 4)[0]
        return out

    def _replace_mini(self, e: DirEntry, data: bytes) -> bytes:
        ms = bytearray(self.ministream)
        s, i = e.start, 0
        while s not in (ENDOFCHAIN, FREESECT) and i * self.msize < len(data):
            chunk = data[i * self.msize:(i + 1) * self.msize]
            ms[s * self.msize:s * self.msize + len(chunk)] = chunk  # 最後のミニセクタの、データより後ろは元のまま
            s = self.minifat[s]
            i += 1
        root = self.entries[0]
        out = bytearray(self.blob)
        for k, sec in enumerate(self.chain(root.start)):
            chunk = bytes(ms[k * self.ssize:(k + 1) * self.ssize])
            out[self._off(sec):self._off(sec) + len(chunk)] = chunk
        return bytes(out)


class SubStorage:
    """複合ファイルの中の入れ物（「ObjectPool/_1774977835」など）を、それだけの複合ファイルのように見せる（読むだけ）。
    文書に埋め込んだ Excel のブックを、.xls を読む処理でそのまま読むのに使う。"""

    def __init__(self, parent: CFB, prefix: str):
        self.parent = parent
        self.prefix = prefix.rstrip("/")
        self.blob = parent.blob

    def streams(self) -> dict[str, DirEntry]:
        pre = self.prefix + "/"
        return {n[len(pre):]: e for n, e in self.parent.streams().items() if n.startswith(pre)}

    def has(self, path: str) -> bool:
        return self.parent.has(self.prefix + "/" + path)

    def read(self, path: str) -> bytes:
        return self.parent.read(self.prefix + "/" + path)
