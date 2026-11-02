"""複合ファイル（MS-CFB）の1つのストリームを、大きさを変えて差し替える（小さなストリーム＝ミニストリームも）。

cfb.CFB.replace は、ふつうのセクタ（512・4096 バイト）に置くストリームの大きさを変えられるが、4096 バイト未満の
小さなストリーム（ミニストリームの 64 バイトのセクタに置く）の大きさは変えられない。ActiveX の部品の中身
（"contents"・.xls の "Ctls"）はたいてい小さいので、ここでミニストリームの大きさを変える。

決まり（MS-CFB 2.4〜2.6）：
- 4096 バイト（ヘッダーの Mini Stream Cutoff Size）未満のストリームは、ミニストリームに置く。以上はふつうのセクタに置く
- ミニストリームは、ルートの入れ物（Root Entry）のふつうのセクタのつながり。大きさはルートの入れ物の Stream Size
- ミニセクタのつながりはミニ FAT に、ミニ FAT のセクタのつながりは FAT に。ミニ FAT のセクタの数はヘッダーにある
書き換えないストリームのバイトは動かさない。足りないセクタはファイルの後ろに足し、使わなくなったセクタ・ミニセクタは
空き（FREESECT）にして 0 で消す（古い中身を残さない）。
"""

from __future__ import annotations

import struct

from .cfb import CFB, DIFSECT, ENDOFCHAIN, FATSECT, FREESECT, CfbError


class _Writer:
    def __init__(self, blob: bytes):
        self.c = CFB(blob)
        c = self.c
        self.ss = c.ssize
        self.per = self.ss // 4
        self.out = bytearray(blob)
        n = (len(self.out) - self.ss + self.ss - 1) // self.ss
        if len(self.out) < (n + 1) * self.ss:
            self.out += bytes((n + 1) * self.ss - len(self.out))
        self.fat = list(c.fat)
        self.fat_secs = list(c.fat_sectors)
        self.minifat = list(c.minifat)
        self.minifat_secs = c.chain(c.minifat_start) if c.n_minifat and c.minifat_start not in (ENDOFCHAIN, FREESECT) else []
        root = c.entries[0]
        self.ms = bytearray(c.ministream)
        self.root_secs = c.chain(root.start) if root.start not in (ENDOFCHAIN, FREESECT) else []

    # ------------------------------------------------------------------ ふつうのセクタ
    def _off(self, s: int) -> int:
        return (s + 1) * self.ss

    def _nsec(self) -> int:
        return len(self.out) // self.ss - 1

    def alloc(self) -> int:
        idx = self._nsec()
        self.out += bytes(self.ss)
        self._grow_fat(idx + 1)
        self.fat[idx] = ENDOFCHAIN
        return idx

    def _grow_fat(self, n: int) -> None:
        while len(self.fat) < n:
            f = self._nsec()
            self.out += bytes(self.ss)
            self.fat.extend([FREESECT] * self.per)
            self.fat[f] = FATSECT
            self.fat_secs.append(f)

    def chain_of(self, start: int) -> list[int]:
        out, s = [], start
        while s not in (ENDOFCHAIN, FREESECT) and s < len(self.fat) and len(out) <= len(self.fat):
            out.append(s)
            s = self.fat[s]
        return out

    def put_regular(self, chain: list[int], data: bytes, old_size: int | None = None) -> list[int]:
        """chain（ふつうのセクタ）に data を書く。足りなければ足し、余れば空きにして 0 で消す。"""
        need = -(-len(data) // self.ss)
        chain = list(chain)
        while len(chain) < need:
            chain.append(self.alloc())
        for s in chain[need:]:
            self.fat[s] = FREESECT
            self.out[self._off(s):self._off(s) + self.ss] = bytes(self.ss)
        chain = chain[:need]
        for a, b in zip(chain, chain[1:]):
            self.fat[a] = b
        if chain:
            self.fat[chain[-1]] = ENDOFCHAIN
        for i, s in enumerate(chain):
            part = data[i * self.ss:(i + 1) * self.ss]
            self.out[self._off(s):self._off(s) + self.ss] = part + bytes(self.ss - len(part))
        return chain

    # ------------------------------------------------------------------ ミニセクタ
    def mini_chain(self, start: int) -> list[int]:
        out, s = [], start
        while s not in (ENDOFCHAIN, FREESECT) and s < len(self.minifat) and len(out) <= len(self.minifat):
            out.append(s)
            s = self.minifat[s]
        return out

    def put_mini(self, chain: list[int], data: bytes) -> list[int]:
        ms = self.c.msize
        need = -(-len(data) // ms)
        chain = list(chain)
        for s in chain[need:]:
            self.minifat[s] = FREESECT
            self.ms[s * ms:(s + 1) * ms] = bytes(ms)
        chain = chain[:need]
        used = set(chain)
        while len(chain) < need:
            free = next((i for i, v in enumerate(self.minifat) if v == FREESECT and i not in used
                         and (i + 1) * ms <= len(self.ms)), None)
            if free is None:
                free = len(self.ms) // ms
                self.ms += bytes(ms)
                if free >= len(self.minifat):
                    self.minifat.extend([FREESECT] * (self.ss // 4))
            chain.append(free)
            used.add(free)
        for a, b in zip(chain, chain[1:]):
            self.minifat[a] = b
        if chain:
            self.minifat[chain[-1]] = ENDOFCHAIN
        for i, s in enumerate(chain):
            part = data[i * ms:(i + 1) * ms]
            self.ms[s * ms:(s + 1) * ms] = part + bytes(ms - len(part))
        return chain

    # ------------------------------------------------------------------ 仕上げ
    def _dir_entry_off(self, index: int) -> int:
        dsecs = self.c.chain(self.c.dir_start)
        dper = self.ss // 128
        return self._off(dsecs[index // dper]) + (index % dper) * 128

    def set_entry(self, index: int, start: int, size: int) -> None:
        base = self._dir_entry_off(index)
        struct.pack_into("<I", self.out, base + 116, start)
        if self.c.major == 4:
            struct.pack_into("<Q", self.out, base + 120, size)
        else:
            struct.pack_into("<I", self.out, base + 120, size)
            # 版 3 では、上の4バイトは 0（MS-CFB 2.6.3）
            struct.pack_into("<I", self.out, base + 124, 0)

    def finish(self, mini_changed: bool) -> bytes:
        c = self.c
        if mini_changed:
            # ミニストリーム（ルートの入れ物の中身）：ふつうのセクタのつながりに書き戻す
            root = c.entries[0]
            self.root_secs = self.put_regular(self.root_secs, bytes(self.ms))
            self.set_entry(0, self.root_secs[0] if self.root_secs else ENDOFCHAIN, len(self.ms))
            del root
            # ミニ FAT：セクタのつながりに書き戻す（増えたら足す）
            n_used = max((i + 1 for i, v in enumerate(self.minifat) if v != FREESECT), default=0)
            per = self.ss // 4
            n_secs = max(-(-n_used // per), len(self.minifat_secs) if n_used else 0)
            mf = self.minifat + [FREESECT] * (n_secs * per - len(self.minifat))
            raw = struct.pack(f"<{len(mf)}I", *mf[:n_secs * per]) if n_secs else b""
            self.minifat_secs = self.put_regular(self.minifat_secs, raw)
            struct.pack_into("<II", self.out, 0x3C, self.minifat_secs[0] if self.minifat_secs else ENDOFCHAIN,
                             len(self.minifat_secs))
        # DIFAT（FAT のセクタの一覧）：先頭の 109 個はヘッダー、残りは DIFAT のセクタ
        difat_secs = self.c.chain_difat()
        while True:
            rest = self.fat_secs[109:]
            need = -(-len(rest) // (self.per - 1)) if rest else 0
            if len(difat_secs) >= need:
                break
            d = self.alloc()
            self.fat[d] = DIFSECT
            difat_secs.append(d)
        head = self.fat_secs[:109]
        rest = self.fat_secs[109:]
        need = -(-len(rest) // (self.per - 1)) if rest else 0
        for k, d in enumerate(difat_secs[:need]):
            vals = rest[k * (self.per - 1):(k + 1) * (self.per - 1)]
            vals += [FREESECT] * (self.per - 1 - len(vals))
            nxt = difat_secs[k + 1] if k + 1 < need else ENDOFCHAIN
            struct.pack_into(f"<{self.per}I", self.out, self._off(d), *vals, nxt)
        for k, s in enumerate(self.fat_secs):
            part = self.fat[k * self.per:(k + 1) * self.per]
            part += [FREESECT] * (self.per - len(part))
            struct.pack_into(f"<{self.per}I", self.out, self._off(s), *part)
        struct.pack_into("<I", self.out, 0x2C, len(self.fat_secs))
        struct.pack_into("<109I", self.out, 0x4C, *(head + [FREESECT] * (109 - len(head))))
        if need:
            struct.pack_into("<II", self.out, 0x44, difat_secs[0], need)
        return bytes(self.out)


def replace_stream(blob: bytes, path: str, data: bytes) -> bytes:
    """複合ファイル blob のストリーム path を data に差し替えたファイル。大きさが変わってもよい（小さなストリームも）。
    ほかのストリームのバイトは変えない。"""
    c = CFB(blob)
    e = c._find(path)
    if e is None:
        raise KeyError(path)
    if e.size < c.cutoff and len(data) == e.size:
        return c.replace(path, data)
    if len(data) >= c.cutoff:
        return c.replace(path, data, wipe=True)  # ふつうのセクタへ（小さなストリームからの移し替えも、cfb.py が扱う）
    w = _Writer(blob)
    if e.size >= c.cutoff:
        # ふつうのセクタ → ミニストリーム：元のセクタを空けて 0 で消す
        for s in w.chain_of(e.start):
            w.fat[s] = FREESECT
            w.out[w._off(s):w._off(s) + w.ss] = bytes(w.ss)
        old_chain: list[int] = []
    else:
        old_chain = w.mini_chain(e.start) if e.size else []
    if not data:
        for s in old_chain:
            w.minifat[s] = FREESECT
            w.ms[s * c.msize:(s + 1) * c.msize] = bytes(c.msize)
        w.set_entry(e.index, ENDOFCHAIN, 0)
        return w.finish(True)
    chain = w.put_mini(old_chain, data)
    w.set_entry(e.index, chain[0], len(data))
    return w.finish(True)


def storage_clsid(cfb: CFB, path: str) -> str | None:
    """入れ物（ストレージ）の CLSID（ディレクトリの項目の 80 バイト目から16バイト、MS-CFB 2.6.1）。path が "" ならルート。"""
    from .activex import guid_str

    if path == "":
        idx = 0
    else:
        low = path.lower().rstrip("/")
        idx = next((e.index for e in cfb.entries if e.kind in (1, 5) and e.path.lower() == low), None)
        if idx is None:
            return None
    raw = cfb._chain_bytes(cfb.dir_start)[idx * 128 + 80:idx * 128 + 96]
    if len(raw) < 16 or raw == bytes(16):
        return None
    return guid_str(raw)


def check(blob: bytes) -> list[str]:
    """書き換えた複合ファイルの形の検査：すべてのストリームが読め、つながりが重ならず、大きさに合う数のセクタがあること。"""
    errs = []
    try:
        c = CFB(blob)
    except CfbError as e:
        return [f"複合ファイルとして読めません：{e}"]
    seen: dict[int, str] = {}
    mseen: dict[int, str] = {}
    for p, e in c.streams().items():
        if e.size >= c.cutoff:
            ch = c.chain(e.start)
            if len(ch) * c.ssize < e.size:
                errs.append(f"{p}: セクタが足りません")
            for s in ch:
                if s in seen:
                    errs.append(f"{p}: セクタ {s} が {seen[s]} と重なっています")
                seen[s] = p
        elif e.size:
            s, n = e.start, 0
            while s not in (ENDOFCHAIN, FREESECT) and s < len(c.minifat):
                if s in mseen:
                    errs.append(f"{p}: ミニセクタ {s} が {mseen[s]} と重なっています")
                    break
                mseen[s] = p
                n += 1
                s = c.minifat[s]
            if n * c.msize < e.size:
                errs.append(f"{p}: ミニセクタが足りません")
            if (max(mseen) + 1) * c.msize > len(c.ministream) if mseen else False:
                errs.append(f"{p}: ミニストリームの外を指しています")
        try:
            c.read(p)
        except Exception as ex:  # noqa: BLE001
            errs.append(f"{p}: 読めません（{ex}）")
    return errs
