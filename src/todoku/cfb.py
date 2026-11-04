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


BROKEN = "壊れた Office の旧形式（複合ファイル）です"


def broken(e: Exception) -> CfbError:
    """複合ファイルの記録の長さ・位置が合わない（途中で切れた・壊れた）ときの、todoku のエラー（struct.error などを包む）。"""
    return CfbError(f"{BROKEN}：記録の長さや位置が合わず、読めません（途中で切れているか、壊れています）")


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
        try:
            self._parse(blob)
        except CfbError:
            raise
        except (struct.error, IndexError, ValueError, OverflowError, MemoryError) as e:
            raise broken(e) from None

    def _parse(self, blob: bytes) -> None:
        self.blob = blob
        (self.minor, self.major, order, sshift, mshift) = struct.unpack_from("<HHHHH", blob, 0x18)
        if order != 0xFFFE:
            raise CfbError("バイト順が想定外です")
        self.ssize = 1 << sshift
        self.msize = 1 << mshift
        (self.n_dir_sectors, self.n_fat, self.dir_start, _txn, self.cutoff, self.minifat_start, self.n_minifat,
         self.difat_start, self.n_difat) = struct.unpack_from("<IIIIIIIII", blob, 0x28)
        if self.cutoff != 4096:  # 決まった値（MS-CFB 2.2）。ほかの値は壊れている（宣言どおりに 2 GB のミニストリームの空きを作らせない）
            raise CfbError(f"{BROKEN}：ミニストリームの境目が {self.cutoff} です（4096 に決まっています）")
        if not 7 <= sshift <= 16 or not 2 <= mshift <= sshift:
            raise CfbError(f"{BROKEN}：セクタの大きさの宣言（2^{sshift}・2^{mshift}）が壊れています")
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

    def mini_chain(self, start: int):
        """ミニ FAT のつながりを順に（使う所まで読めばよいので、少しずつ）：輪になっていれば CfbError（どのミニセクタも 1 度まで）。"""
        s, seen = start, set()
        while s not in (ENDOFCHAIN, FREESECT) and s < len(self.minifat):
            if s in seen:
                raise CfbError("ミニセクタのつながりが循環しています")
            seen.add(s)
            yield s
            s = self.minifat[s]

    def _mini_bytes(self, start: int, size: int) -> bytearray:
        out = bytearray()
        for s in self.mini_chain(start):
            if len(out) >= size:
                break
            piece = self.ministream[s * self.msize:(s + 1) * self.msize]
            if not piece:  # ミニストリームの外を指す：ここまで
                break
            out += piece
        return out[:size]

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self.blob[self._off(s):self._off(s) + self.ssize] for s in self.chain(start))

    def _paths(self) -> None:
        seen: set[int] = set()  # たどった記録：左・右・子が輪になった木（壊れた・悪意のあるファイル）も、1 度ずつだけたどる

        def walk(idx: int, prefix: str, depth: int = 0):
            if idx == NOSTREAM or idx >= len(self.entries) or depth > 512 or idx in seen:
                return
            seen.add(idx)
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
            return bytes(self._mini_bytes(e.start, e.size))
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
            while s_ not in (ENDOFCHAIN, FREESECT) and s_ < len(self.minifat) and seen <= len(self.minifat):
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
        for i, s in enumerate(self.mini_chain(e.start)):
            if i * self.msize >= len(data):
                break
            chunk = data[i * self.msize:(i + 1) * self.msize]
            ms[s * self.msize:s * self.msize + len(chunk)] = chunk  # 最後のミニセクタの、データより後ろは元のまま
        root = self.entries[0]
        out = bytearray(self.blob)
        for k, sec in enumerate(self.chain(root.start)):
            chunk = bytes(ms[k * self.ssize:(k + 1) * self.ssize])
            out[self._off(sec):self._off(sec) + len(chunk)] = chunk
        return bytes(out)


def name_key(name: str) -> tuple[int, str]:
    """ディレクトリの木の並び順（MS-CFB）：名前の長さ（UTF-16 の単位の数）、同じなら大文字にした字の順。"""
    return (len(name.encode("utf-16-le")) // 2, name.upper())


class CfbEditor:
    """複合ファイルの、ストリームの差し替え（大きさを変えてよい）と、ストリームの取り除き。

    CFB.replace と同じく、手を入れるストリームのセクタ・ミニセクタと、それをつなぐ表（FAT・ミニ FAT・ディレクトリ）だけを
    書き換え、ほかのバイトは変えない。使わなくなったセクタ・ミニセクタは 0 で消す（元のデータを残さない）。
    - ふつうのセクタのストリーム：元のセクタを順に使い、足りなければファイルの末尾に足す
    - ミニストリームのストリーム（4096 バイト未満）：元のミニセクタを順に使い、足りなければミニストリームの末尾に足す
      （ミニストリームを入れるルートのセクタ・ミニ FAT のセクタも、足りなければ末尾に足す）
    - 取り除くストリーム：データを消し、ディレクトリの項目を空き（種類 0）にして、その入れ物の子の木を作り直す
      （名前の順の釣り合った二分木。いちばん深い段だけを赤にした赤黒木）
    DIFAT のセクタが要る大きさ（FAT のセクタが 109 を超える）には対応しない（CfbError）。"""

    def __init__(self, blob: bytes):
        c = CFB(blob)
        self.c = c
        ss = c.ssize
        self.out = bytearray(blob)
        if len(self.out) % ss:
            self.out += bytes(ss - len(self.out) % ss)
        self.fat = list(c.fat)
        self.fat_sectors = list(c.fat_sectors)
        if c.chain_difat():
            raise CfbError("DIFAT のセクタを持つ大きな複合ファイルには、大きさを変える書き込みをしません")
        self.minifat = list(c.minifat)
        root = c.entries[0]
        self.root_chain = c.chain(root.start) if root.start not in (ENDOFCHAIN, FREESECT) else []
        self.ministream = bytearray(c.ministream)
        self.minifat_chain = c.chain(c.minifat_start) if c.minifat_start not in (ENDOFCHAIN, FREESECT) else []
        self.dir_chain = c.chain(c.dir_start)
        self.dir = bytearray(c._chain_bytes(c.dir_start))
        self.n_sectors = len(self.out) // ss - 1

    # ------------------------------------------------------------------ セクタ
    def _off(self, s: int) -> int:
        return (s + 1) * self.c.ssize

    def _alloc(self) -> int:
        """ファイルの末尾に、ふつうのセクタを1つ足す（FAT が足りなければ、FAT のセクタも足す）。"""
        ss, per = self.c.ssize, self.c.ssize // 4
        idx = self.n_sectors
        self.n_sectors += 1
        self.out += bytes(ss)
        while len(self.fat) < self.n_sectors:
            if len(self.fat_sectors) >= 109:
                raise CfbError("FAT のセクタが 109 を超えるので、大きさを変えられません")
            f = self.n_sectors
            self.n_sectors += 1
            self.out += bytes(ss)
            self.fat += [FREESECT] * per
            self.fat[f] = FATSECT
            self.fat_sectors.append(f)
        return idx

    def _free(self, secs: list[int]) -> None:
        for s in secs:
            self.fat[s] = FREESECT
            o = self._off(s)
            self.out[o:o + self.c.ssize] = bytes(self.c.ssize)

    def _mini_chain(self, start: int) -> list[int]:
        out, s = [], start
        while s not in (ENDOFCHAIN, FREESECT) and s < len(self.minifat) and len(out) <= len(self.minifat):
            out.append(s)
            s = self.minifat[s]
        return out

    def _free_mini(self, secs: list[int]) -> None:
        ms = self.c.msize
        for m in secs:
            self.minifat[m] = FREESECT
            self.ministream[m * ms:(m + 1) * ms] = bytes(ms)

    # ------------------------------------------------------------------ ディレクトリ
    def _entry(self, path: str) -> DirEntry:
        e = self.c._find(path)
        if e is None:
            raise KeyError(path)
        return e

    def _set_entry(self, idx: int, start: int, size: int) -> None:
        base = idx * 128
        struct.pack_into("<I", self.dir, base + 116, start)
        struct.pack_into("<Q" if self.c.major == 4 else "<I", self.dir, base + 120, size)
        if self.c.major != 4:
            struct.pack_into("<I", self.dir, base + 124, 0)

    # ------------------------------------------------------------------ 操作
    def replace(self, path: str, data: bytes) -> None:
        e = self._entry(path)
        c = self.c
        old_mini = e.size < c.cutoff
        if old_mini:
            old_secs = self._mini_chain(e.start) if e.size else []
        else:
            old_secs = c.chain(e.start)
        if len(data) < c.cutoff:
            # ミニストリームに置く
            ms = c.msize
            if not old_mini:
                self._free(old_secs)
                old_secs = []
            need = -(-len(data) // ms)
            secs = old_secs[:need]
            self._free_mini(old_secs[need:])
            while len(secs) < need:
                m = len(self.ministream) // ms
                self.ministream += bytes(ms)
                self.minifat.append(FREESECT)
                secs.append(m)
            for a, b in zip(secs, secs[1:]):
                self.minifat[a] = b
            if secs:
                self.minifat[secs[-1]] = ENDOFCHAIN
            for k, m in enumerate(secs):
                chunk = data[k * ms:(k + 1) * ms]
                self.ministream[m * ms:(m + 1) * ms] = chunk + bytes(ms - len(chunk))
            self._set_entry(e.index, secs[0] if secs else ENDOFCHAIN, len(data))
            return
        ss = c.ssize
        if old_mini:
            self._free_mini(old_secs)
            old_secs = []
        need = -(-len(data) // ss)
        secs = old_secs[:need]
        self._free(old_secs[need:])
        while len(secs) < need:
            secs.append(self._alloc())
        for a, b in zip(secs, secs[1:]):
            self.fat[a] = b
        self.fat[secs[-1]] = ENDOFCHAIN
        for k, s in enumerate(secs):
            chunk = data[k * ss:(k + 1) * ss]
            o = self._off(s)
            self.out[o:o + ss] = chunk + bytes(ss - len(chunk))
        self._set_entry(e.index, secs[0], len(data))

    def remove(self, path: str) -> None:
        """ストリームを取り除く（データを 0 で消し、ディレクトリの項目を空きにし、入れ物の子の木を作り直す）。"""
        e = self._entry(path)
        c = self.c
        if e.size < c.cutoff:
            self._free_mini(self._mini_chain(e.start) if e.size else [])
        else:
            self._free(c.chain(e.start))
        parent = self._parent_of(e.index)
        kids = [k for k in self._siblings(self._child_of(parent)) if k != e.index]
        kids.sort(key=lambda k: name_key(self._name(k)))
        top = self._rebuild(kids)
        struct.pack_into("<I", self.dir, parent * 128 + 76, top)
        base = e.index * 128
        self.dir[base:base + 128] = bytes(128)
        struct.pack_into("<III", self.dir, base + 68, NOSTREAM, NOSTREAM, NOSTREAM)

    def _name(self, idx: int) -> str:
        base = idx * 128
        n = struct.unpack_from("<H", self.dir, base + 64)[0]
        return bytes(self.dir[base:base + max(0, n - 2)]).decode("utf-16-le", errors="replace")

    def _links(self, idx: int) -> tuple[int, int, int]:
        return struct.unpack_from("<III", self.dir, idx * 128 + 68)

    def _child_of(self, idx: int) -> int:
        return self._links(idx)[2]

    def _siblings(self, top: int) -> list[int]:
        out, stack, seen = [], [top], set()
        while stack:
            k = stack.pop()
            if k == NOSTREAM or k in seen or k * 128 >= len(self.dir):
                continue
            seen.add(k)
            out.append(k)
            left, right, _child = self._links(k)
            stack += [left, right]
        return out

    def _parent_of(self, idx: int) -> int:
        todo, done = [0], set()
        while todo:
            p = todo.pop()
            if p in done:  # 輪になった木
                continue
            done.add(p)
            sib = self._siblings(self._child_of(p))
            if idx in sib:
                return p
            todo += [k for k in sib if self.dir[k * 128 + 66] == 1]
        raise CfbError("取り除くストリームの入れ物が見つかりません")

    def _rebuild(self, kids: list[int]) -> int:
        """名前の順に並べた子 → 釣り合った二分木（根の番号）。いちばん深い段が埋まっていなければ、その段を赤にする。"""
        depth: dict[int, int] = {}

        def build(lo: int, hi: int, d: int) -> int:
            if lo >= hi:
                return NOSTREAM
            mid = (lo + hi) // 2
            k = kids[mid]
            depth[k] = d
            left, right = build(lo, mid, d + 1), build(mid + 1, hi, d + 1)
            struct.pack_into("<II", self.dir, k * 128 + 68, left, right)
            return k

        top = build(0, len(kids), 0)
        if kids:
            deepest = max(depth.values())
            full = len(kids) == (1 << (deepest + 1)) - 1
            for k, d in depth.items():
                self.dir[k * 128 + 67] = 0 if (d == deepest and not full) else 1  # 0 赤・1 黒
        return top

    # ------------------------------------------------------------------ 書き出し
    def blob(self) -> bytes:
        c = self.c
        ss, per = c.ssize, c.ssize // 4
        # ミニストリーム（ルートのセクタ）
        need = -(-len(self.ministream) // ss)
        while len(self.root_chain) < need:
            self.root_chain.append(self._alloc())
        for a, b in zip(self.root_chain, self.root_chain[1:]):
            self.fat[a] = b
        if self.root_chain:
            self.fat[self.root_chain[-1]] = ENDOFCHAIN
        for k, s in enumerate(self.root_chain):
            chunk = bytes(self.ministream[k * ss:(k + 1) * ss])
            o = self._off(s)
            self.out[o:o + len(chunk)] = chunk  # ミニストリームの終わりより後ろ（最後のセクタの残り）は元のまま
        self._set_entry(0, self.root_chain[0] if self.root_chain else ENDOFCHAIN, len(self.ministream))
        # ミニ FAT
        mf = list(self.minifat)
        need = -(-len(mf) * 4 // ss)
        while len(self.minifat_chain) < need:
            self.minifat_chain.append(self._alloc())
        for a, b in zip(self.minifat_chain, self.minifat_chain[1:]):
            self.fat[a] = b
        if self.minifat_chain:
            self.fat[self.minifat_chain[-1]] = ENDOFCHAIN
        mf += [FREESECT] * (len(self.minifat_chain) * per - len(mf))
        for k, s in enumerate(self.minifat_chain):
            struct.pack_into(f"<{per}I", self.out, self._off(s), *mf[k * per:(k + 1) * per])
        # ディレクトリ
        for k, s in enumerate(self.dir_chain):
            o = self._off(s)
            self.out[o:o + ss] = self.dir[k * ss:(k + 1) * ss]
        # FAT と見出し
        for k, s in enumerate(self.fat_sectors):
            part = self.fat[k * per:(k + 1) * per]
            part += [FREESECT] * (per - len(part))
            struct.pack_into(f"<{per}I", self.out, self._off(s), *part)
        struct.pack_into("<I", self.out, 0x2C, len(self.fat_sectors))
        struct.pack_into("<II", self.out, 0x3C, self.minifat_chain[0] if self.minifat_chain else ENDOFCHAIN,
                         len(self.minifat_chain))
        head = self.fat_sectors + [FREESECT] * (109 - len(self.fat_sectors))
        struct.pack_into("<109I", self.out, 0x4C, *head)
        return bytes(self.out)


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

    def contents(self) -> dict[str, bytes]:
        """入れ物の中のすべてのストリーム（入れ物からの道筋 → 中身）。build_cfb で、それだけの複合ファイルにできる。"""
        return {p: self.read(p) for p in self.streams()}


def build_cfb(streams: dict[str, bytes]) -> bytes:
    """ストリーム（「VBA/dir」のように / で入れ物を表す）から、新しい複合ファイル（版3、512 バイトのセクタ）を作る。

    文書に埋め込んだ Excel のブックの入れ物を、それだけの .xls として読み書き・証明するのに使う（利用者のファイルには
    書かない。書き戻すのは、中のストリームだけ）。4096 バイト未満のストリームはミニストリームに置く。
    ディレクトリの木は、名前の順（長さ、大文字にした字）の釣り合った二分木。"""
    SS, MS, CUT = 512, 64, 4096
    children: dict[str, list[str]] = {"": []}  # 入れ物 → 子の名前
    for p in streams:
        parts = p.split("/")
        for k in range(1, len(parts)):
            me = "/".join(parts[:k])
            if me not in children:
                children[me] = []
                children.setdefault("/".join(parts[:k - 1]), []).append(parts[k - 1])
        children.setdefault("/".join(parts[:-1]), []).append(parts[-1])
    order = [""] + sorted((n for n in children if n), key=lambda n: (n.count("/"), n)) + sorted(streams)
    index = {p: i for i, p in enumerate(order)}
    links: dict[int, tuple[int, int]] = {}

    def tree(parent: str) -> int:
        kids = sorted(set(children.get(parent, [])), key=lambda n: (len(n), n.upper()))
        paths = [(parent + "/" + c) if parent else c for c in kids]

        def build(lo: int, hi: int) -> int:
            if lo >= hi:
                return NOSTREAM
            mid = (lo + hi) // 2
            links[index[paths[mid]]] = (build(lo, mid), build(mid + 1, hi))
            return index[paths[mid]]

        return build(0, len(paths))

    child = {p: tree(p) for p in children}
    mini = bytearray()
    minifat: list[int] = []
    start: dict[str, int] = {}
    for p in sorted(streams):
        d = streams[p]
        if 0 < len(d) < CUT:
            first, n = len(mini) // MS, -(-len(d) // MS)
            mini += d + b"\x00" * (n * MS - len(d))
            minifat += list(range(first + 1, first + n)) + [ENDOFCHAIN]
            start[p] = first
    big = [p for p in sorted(streams) if len(streams[p]) >= CUT]
    n_dir = -(-len(order) * 128 // SS)
    n_minifat = -(-len(minifat) * 4 // SS)
    n_mini = -(-len(mini) // SS)
    n_data = n_mini + sum(-(-len(streams[p]) // SS) for p in big)
    n_fat = 1
    while n_fat * (SS // 4) < n_fat + n_dir + n_minifat + n_data:
        n_fat += 1
    if n_fat > 109:
        raise CfbError("作る複合ファイルが大きすぎます")
    fat = [FREESECT] * (n_fat * (SS // 4))
    at = n_fat
    for k in range(n_fat):
        fat[k] = FATSECT

    def chain(n: int) -> int:
        nonlocal at
        if n == 0:
            return ENDOFCHAIN
        first = at
        for k in range(n):
            fat[at + k] = at + k + 1 if k + 1 < n else ENDOFCHAIN
        at += n
        return first

    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    mini_start = chain(n_mini)
    for p in big:
        start[p] = chain(-(-len(streams[p]) // SS))
    out = bytearray(SS * (1 + at))
    out[0:8] = MAGIC
    struct.pack_into("<HHHHH", out, 0x18, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", out, 0x28, 0, n_fat, dir_start, 0, CUT, minifat_start, n_minifat, ENDOFCHAIN, 0)
    struct.pack_into("<109I", out, 0x4C, *(list(range(n_fat)) + [FREESECT] * (109 - n_fat)))
    struct.pack_into(f"<{len(fat)}I", out, SS, *fat)

    def put(sector: int, data: bytes) -> None:
        out[SS * (1 + sector):SS * (1 + sector) + len(data)] = data

    if n_minifat:
        mf = minifat + [FREESECT] * (n_minifat * (SS // 4) - len(minifat))
        put(minifat_start, struct.pack(f"<{len(mf)}I", *mf))
    if n_mini:
        put(mini_start, bytes(mini))
    for p in big:
        put(start[p], streams[p])
    dir_bytes = bytearray(n_dir * SS)
    for i in range(n_dir * SS // 128):
        struct.pack_into("<III", dir_bytes, i * 128 + 68, NOSTREAM, NOSTREAM, NOSTREAM)
    for i, p in enumerate(order):
        raw = (p.rsplit("/", 1)[-1] if p else "Root Entry").encode("utf-16-le")[:62] + b"\x00\x00"
        kind = 5 if p == "" else (2 if p in streams else 1)
        left, right = links.get(i, (NOSTREAM, NOSTREAM))
        if kind == 5:
            st_, size = mini_start, len(mini)
        elif kind == 2:
            st_, size = start.get(p, ENDOFCHAIN), len(streams[p])
        else:
            st_, size = 0, 0
        base = i * 128
        dir_bytes[base:base + len(raw)] = raw
        struct.pack_into("<HBB", dir_bytes, base + 64, len(raw), kind, 1)
        struct.pack_into("<III", dir_bytes, base + 68, left, right, child.get(p, NOSTREAM) if kind != 2 else NOSTREAM)
        struct.pack_into("<IQ", dir_bytes, base + 116, st_, size)
    put(dir_start, bytes(dir_bytes))
    return bytes(out)
