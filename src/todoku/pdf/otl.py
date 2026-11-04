"""OpenType のレイアウトの表（GDEF・GSUB・GPOS）を読み、字形の並びに当てる。

仕様：OpenType 1.9（Microsoft Typography の「Common Table Formats」「GDEF」「GSUB」「GPOS」）。
当て方（段階・字の目印（mask）・ZWJ と ZWNJ の飛ばし方・記号の付け方）は、HarfBuzz（MIT。LICENSES/LICENSE-harfbuzz.txt） の OpenType の組み立てを Python に写したもの
（同じ書体で同じ字形の番号と位置になることを、tests で hb-shape と照らして確かめる）。

扱うもの：
- GSUB：1 単独・2 複数・4 合字・5 文脈・6 連鎖文脈・7 拡張・8 逆向きの連鎖（3 代替は、選ぶ字形なので使わない）
- GPOS：1 単独・2 対・3 筆記体の接続・4 記号を基の字に・5 記号を合字に・6 記号を記号に・7 文脈・8 連鎖文脈・9 拡張
- GDEF：字形の分類・記号の接続の分類・記号の集まり（なければ、字の一般カテゴリから分類を決める）
使わないもの：装置の表（Device）・可変の書体（FeatureVariations）。表が読めないときは ValueError（組み立てずに断る）。
"""

from __future__ import annotations

import struct

BASE, LIGATURE, MARK, COMPONENT = 1, 2, 3, 4
SUBSTITUTED, LIGATED, MULTIPLIED = 1, 2, 4
ZWNJ, ZWJ = 0x200C, 0x200D

_u16 = struct.Struct(">H").unpack_from
_s16 = struct.Struct(">h").unpack_from


def u16(d: bytes, o: int) -> int:
    return _u16(d, o)[0]


def s16(d: bytes, o: int) -> int:
    return _s16(d, o)[0]


def coverage(d: bytes, off: int) -> dict[int, int]:
    """Coverage：字形 → 番号。"""
    fmt = u16(d, off)
    if fmt == 1:
        n = u16(d, off + 2)
        return {g: i for i, g in enumerate(struct.unpack_from(f">{n}H", d, off + 4))}
    if fmt == 2:
        out = {}
        n = u16(d, off + 2)
        for r in range(n):
            s, e, si = struct.unpack_from(">HHH", d, off + 4 + 6 * r)
            for g in range(s, e + 1):
                out[g] = si + g - s
        return out
    raise ValueError(f"Coverage の形 {fmt} を読めません")


def classdef(d: bytes, off: int) -> dict[int, int]:
    """ClassDef：字形 → 分類（書いていない字形は 0）。"""
    fmt = u16(d, off)
    out = {}
    if fmt == 1:
        start, n = struct.unpack_from(">HH", d, off + 2)
        for i, v in enumerate(struct.unpack_from(f">{n}H", d, off + 6)):
            if v:
                out[start + i] = v
        return out
    if fmt == 2:
        n = u16(d, off + 2)
        for r in range(n):
            s, e, c = struct.unpack_from(">HHH", d, off + 4 + 6 * r)
            for g in range(s, e + 1):
                out[g] = c
        return out
    raise ValueError(f"ClassDef の形 {fmt} を読めません")


# 字形の分類が誤っていると分かっている書体の GDEF（GDEF・GSUB・GPOS の長さ）。HarfBuzz の hb-ot-layout.cc の
# GDEF::is_blocklisted と同じもの（Tahoma の各版・Times New Roman Italic・Himalaya・Cantarell・Padauk・Courier New）。
# この GDEF は使わず、字の一般カテゴリから分類を決める（間隔のある記号を、幅のない記号として扱ってしまうため）。
GDEF_BLOCKLIST = {
    (442, 2874, 42038), (430, 2874, 40662), (442, 2874, 39116), (430, 2874, 39374), (490, 3046, 41638),
    (478, 3046, 41902), (898, 12554, 46470), (910, 12566, 47732), (928, 23298, 59332), (940, 23310, 60732),
    (964, 23836, 60072), (976, 23832, 61456), (994, 24474, 60336), (1006, 24470, 61740), (1006, 24576, 61346),
    (1018, 24572, 62828), (1006, 24576, 61352), (1018, 24572, 62834), (832, 7324, 47162), (844, 7302, 45474),
    (180, 13054, 7254), (192, 12638, 7254), (192, 12690, 7254), (188, 248, 3852), (188, 264, 3426),
    (1058, 47032, 11818), (1046, 47030, 12600), (1058, 71796, 16770), (1046, 71790, 17862), (1046, 71788, 17112),
    (1058, 71794, 17514), (1330, 109904, 57938), (1330, 109904, 58972), (1004, 59092, 14836), (588, 5078, 14418),
}


class GDEF:
    def __init__(self, d: bytes | None):
        self.present = bool(d)
        self.glyph_class: dict[int, int] = {}
        self.mark_class: dict[int, int] = {}
        self.mark_sets: list[set[int]] = []
        if not d:
            return
        _maj, minor, gc, _al, _lc, mac = struct.unpack_from(">HHHHHH", d, 0)
        if gc:
            self.glyph_class = classdef(d, gc)
        if mac:
            self.mark_class = classdef(d, mac)
        if minor >= 2 and len(d) >= 14:
            ms = u16(d, 12)
            if ms:
                n = u16(d, ms + 2)
                for i in range(n):
                    o = struct.unpack_from(">I", d, ms + 4 + 4 * i)[0]
                    self.mark_sets.append(set(coverage(d, ms + o)))


# ====================================================================== 字形の並び
class Info:
    """字形の並びの 1 つ。chars：この字形が表す元の字（合字なら合わせたもの）。"""
    __slots__ = ("gid", "cp", "cluster", "mask", "gclass", "mclass", "lig_id", "lig_comp", "flags", "syl", "cat",
                 "pos", "chars", "xa", "ya", "xo", "yo", "attach", "atype", "feat", "extra", "skind")

    def __init__(self, gid: int, cp: int, cluster: int, chars: str):
        self.gid = gid
        self.cp = cp
        self.cluster = cluster
        self.mask = 0
        self.gclass = BASE
        self.mclass = 0
        self.lig_id = 0
        self.lig_comp = 0
        self.flags = 0
        self.syl = 0
        self.cat = ""
        self.pos = 0
        self.chars = chars
        self.xa = self.ya = self.xo = self.yo = 0
        self.attach = -1
        self.atype = 0  # 1：記号、2：筆記体
        self.feat: set[str] = set()  # この字形を置き換えた機能（init・fina などを確かめる）
        self.extra = None  # 合字の部品の数
        self.skind = ""  # 音節の種類（インド系の文字）

    def copy(self) -> "Info":
        n = Info(self.gid, self.cp, self.cluster, self.chars)
        for k in Info.__slots__:
            setattr(n, k, getattr(self, k))
        n.feat = set(self.feat)
        return n

    @property
    def is_mark(self) -> bool:
        return self.gclass == MARK

    @property
    def ligated(self) -> bool:
        return bool(self.flags & LIGATED)

    @property
    def multiplied(self) -> bool:
        return bool(self.flags & MULTIPLIED)

    def __repr__(self) -> str:
        return f"<{self.gid} {self.chars!r} m={self.mask:x} c={self.gclass} s={self.syl} {self.cat}/{self.pos}>"


class Lookup:
    __slots__ = ("type", "flag", "mfs", "subs")

    def __init__(self, typ: int, flag: int, mfs: int | None, subs: list):
        self.type = typ
        self.flag = flag
        self.mfs = mfs
        self.subs = subs


class Stage:
    """1 つの段階で当てる参照の表：[(参照の番号, mask, auto_zwj, auto_zwnj, 音節ごと, 機能の名前)]。"""

    def __init__(self, items: list, pause=None):
        self.items = items
        self.pause = pause


# ====================================================================== GSUB・GPOS
class Layout:
    def __init__(self, d: bytes | None, gsub: bool):
        self.d = d or b""
        self.gsub = gsub
        self._lookups: dict[int, Lookup] = {}
        self.scripts: dict[str, int] = {}
        self.feature_tags: list[str] = []
        self.feature_lookups: list[list[int]] = []
        if not d:
            return
        _maj, _min, sl, fl, ll = struct.unpack_from(">HHHHH", d, 0)
        self.ll = ll
        n = u16(d, sl)
        for i in range(n):
            tag = d[sl + 2 + 6 * i:sl + 6 + 6 * i].decode("latin-1")
            self.scripts[tag] = sl + u16(d, sl + 6 + 6 * i)
        n = u16(d, fl)
        for i in range(n):
            tag = d[fl + 2 + 6 * i:fl + 6 + 6 * i].decode("latin-1")
            fo = fl + u16(d, fl + 6 + 6 * i)
            cnt = u16(d, fo + 2)
            self.feature_tags.append(tag)
            self.feature_lookups.append(list(struct.unpack_from(f">{cnt}H", d, fo + 4)))

    def script(self, tags) -> str | None:
        for t in tags:
            if t in self.scripts:
                return t
        return None

    def features(self, script_tag: str | None) -> dict[str, list[int]]:
        """その文字（の既定の言語）で使える機能 → 参照の番号。"""
        out: dict[str, list[int]] = {}
        if script_tag is None or script_tag not in self.scripts:
            return out
        so = self.scripts[script_tag]
        dl = u16(self.d, so)
        if not dl:
            return out
        lo = so + dl
        req, cnt = struct.unpack_from(">HH", self.d, lo + 2)
        idx = list(struct.unpack_from(f">{cnt}H", self.d, lo + 6))
        if req != 0xFFFF:
            idx.append(req)
        for i in idx:
            if i < len(self.feature_tags):
                out.setdefault(self.feature_tags[i], [])
                for x in self.feature_lookups[i]:
                    if x not in out[self.feature_tags[i]]:
                        out[self.feature_tags[i]].append(x)
        return out

    # ------------------------------------------------------------------ 参照を読む
    def lookup(self, idx: int) -> Lookup:
        lk = self._lookups.get(idx)
        if lk is None:
            lk = self._lookups[idx] = self._read_lookup(idx)
        return lk

    def _read_lookup(self, idx: int) -> Lookup:
        d = self.d
        n = u16(d, self.ll)
        if idx >= n:
            raise ValueError(f"参照 {idx} がありません")
        off = self.ll + u16(d, self.ll + 2 + 2 * idx)
        typ, flag, cnt = struct.unpack_from(">HHH", d, off)
        offs = [off + x for x in struct.unpack_from(f">{cnt}H", d, off + 6)]
        mfs = u16(d, off + 6 + 2 * cnt) if flag & 0x10 else None
        ext = 7 if self.gsub else 9
        subs = []
        real = typ
        for so in offs:
            t = typ
            if typ == ext:
                _fmt, t, o32 = struct.unpack_from(">HHI", d, so)
                so += o32
            real = t
            subs.append(self._read_sub(t, so))
        return Lookup(real, flag, mfs, subs)

    def _read_sub(self, t: int, o: int):
        d = self.d
        fmt = u16(d, o)
        if self.gsub:
            if t == 1:
                cov = coverage(d, o + u16(d, o + 2))
                if fmt == 1:
                    delta = s16(d, o + 4)
                    return ("single", {g: (g + delta) & 0xFFFF for g in cov})
                n = u16(d, o + 4)
                subs = struct.unpack_from(f">{n}H", d, o + 6)
                return ("single", {g: subs[i] for g, i in cov.items() if i < n})
            if t == 2:
                cov = coverage(d, o + u16(d, o + 2))
                n = u16(d, o + 4)
                seqs = []
                for i in range(n):
                    so = o + u16(d, o + 6 + 2 * i)
                    c = u16(d, so)
                    seqs.append(tuple(struct.unpack_from(f">{c}H", d, so + 2)))
                return ("multiple", {g: seqs[i] for g, i in cov.items() if i < n})
            if t == 3:
                return ("alternate",)
            if t == 4:
                cov = coverage(d, o + u16(d, o + 2))
                n = u16(d, o + 4)
                sets = []
                for i in range(n):
                    so = o + u16(d, o + 6 + 2 * i)
                    c = u16(d, so)
                    ligs = []
                    for j in range(c):
                        lo = so + u16(d, so + 2 + 2 * j)
                        lg, cc = struct.unpack_from(">HH", d, lo)
                        ligs.append((lg, tuple(struct.unpack_from(f">{cc - 1}H", d, lo + 4))))
                    sets.append(ligs)
                return ("ligature", {g: sets[i] for g, i in cov.items() if i < n})
            if t == 5:
                return self._read_context(o)
            if t == 6:
                return self._read_chain(o)
            if t == 8:
                cov = coverage(d, o + u16(d, o + 2))
                p = o + 4
                nb = u16(d, p)
                back = [set(coverage(d, o + x)) for x in struct.unpack_from(f">{nb}H", d, p + 2)]
                p += 2 + 2 * nb
                nl = u16(d, p)
                look = [set(coverage(d, o + x)) for x in struct.unpack_from(f">{nl}H", d, p + 2)]
                p += 2 + 2 * nl
                ns = u16(d, p)
                subs = struct.unpack_from(f">{ns}H", d, p + 2)
                return ("reverse", {g: subs[i] for g, i in cov.items() if i < ns}, back, look)
            raise ValueError(f"GSUB の参照の種類 {t} を読めません")
        # GPOS
        if t == 1:
            cov = coverage(d, o + u16(d, o + 2))
            vf = u16(d, o + 4)
            if fmt == 1:
                v = _value(d, o + 6, vf)
                return ("psingle", {g: v for g in cov})
            n = u16(d, o + 6)
            size = _vsize(vf)
            vals = [_value(d, o + 8 + size * i, vf) for i in range(n)]
            return ("psingle", {g: vals[i] for g, i in cov.items() if i < n})
        if t == 2:
            cov = coverage(d, o + u16(d, o + 2))
            vf1, vf2 = struct.unpack_from(">HH", d, o + 4)
            s1, s2 = _vsize(vf1), _vsize(vf2)
            if fmt == 1:
                n = u16(d, o + 8)
                sets = []
                for i in range(n):
                    so = o + u16(d, o + 10 + 2 * i)
                    c = u16(d, so)
                    m = {}
                    p = so + 2
                    for _ in range(c):
                        g2 = u16(d, p)
                        if g2 not in m:
                            m[g2] = (_value(d, p + 2, vf1), _value(d, p + 2 + s1, vf2))
                        p += 2 + s1 + s2
                    sets.append(m)
                return ("pair1", {g: sets[i] for g, i in cov.items() if i < n}, vf2)
            cd1 = classdef(d, o + u16(d, o + 8))
            cd2 = classdef(d, o + u16(d, o + 10))
            c1, c2 = struct.unpack_from(">HH", d, o + 12)
            rows = []
            p = o + 16
            for _i in range(c1):
                row = []
                for _j in range(c2):
                    row.append((_value(d, p, vf1), _value(d, p + s1, vf2)))
                    p += s1 + s2
                rows.append(row)
            return ("pair2", set(cov), cd1, cd2, rows, vf2)
        if t == 3:
            cov = coverage(d, o + u16(d, o + 2))
            n = u16(d, o + 4)
            recs = []
            for i in range(n):
                en, ex = struct.unpack_from(">HH", d, o + 6 + 4 * i)
                recs.append((_anchor(d, o + en) if en else None, _anchor(d, o + ex) if ex else None))
            return ("cursive", {g: recs[i] for g, i in cov.items() if i < n})
        if t in (4, 5, 6):
            mcov = coverage(d, o + u16(d, o + 2))
            bcov = coverage(d, o + u16(d, o + 4))
            ncls = u16(d, o + 6)
            ma = o + u16(d, o + 8)
            ba = o + u16(d, o + 10)
            marks = {}
            nm = u16(d, ma)
            mrec = [struct.unpack_from(">HH", d, ma + 2 + 4 * i) for i in range(nm)]
            for g, i in mcov.items():
                if i < nm:
                    cls, ao = mrec[i]
                    marks[g] = (cls, _anchor(d, ma + ao))
            if t in (4, 6):
                nb = u16(d, ba)
                bases = {}
                for g, i in bcov.items():
                    if i >= nb:
                        continue
                    row = []
                    for c in range(ncls):
                        ao = u16(d, ba + 2 + 2 * (i * ncls + c))
                        row.append(_anchor(d, ba + ao) if ao else None)
                    bases[g] = row
                return ("markbase" if t == 4 else "markmark", marks, bases)
            nl = u16(d, ba)
            ligs = {}
            for g, i in bcov.items():
                if i >= nl:
                    continue
                lo = ba + u16(d, ba + 2 + 2 * i)
                ncomp = u16(d, lo)
                comps = []
                for k in range(ncomp):
                    row = []
                    for c in range(ncls):
                        ao = u16(d, lo + 2 + 2 * (k * ncls + c))
                        row.append(_anchor(d, lo + ao) if ao else None)
                    comps.append(row)
                ligs[g] = comps
            return ("marklig", marks, ligs)
        if t == 7:
            return self._read_context(o)
        if t == 8:
            return self._read_chain(o)
        raise ValueError(f"GPOS の参照の種類 {t} を読めません")

    def _records(self, p: int, n: int) -> list[tuple[int, int]]:
        return [struct.unpack_from(">HH", self.d, p + 4 * i) for i in range(n)]

    def _read_context(self, o: int):
        """文脈（GSUB 5・GPOS 7）→ ("ctx", 最初の字形の Coverage, 規則を選ぶ関数の材料, 規則)。
        規則：(後ろの並び, 入力の 2 つ目から, 先の並び, 参照の記録)。並びの要素は、照らし方 kind で読む。"""
        d = self.d
        fmt = u16(d, o)
        if fmt == 1:
            cov = coverage(d, o + u16(d, o + 2))
            n = u16(d, o + 4)
            sets = []
            for i in range(n):
                so = u16(d, o + 6 + 2 * i)
                rules = []
                if so:
                    so += o
                    for j in range(u16(d, so)):
                        ro = so + u16(d, so + 2 + 2 * j)
                        gc, sc = struct.unpack_from(">HH", d, ro)
                        inp = tuple(struct.unpack_from(f">{gc - 1}H", d, ro + 4))
                        rules.append(((), inp, (), self._records(ro + 4 + 2 * (gc - 1), sc)))
                sets.append(rules)
            return ("ctx", "g", cov, None, sets, None, None)
        if fmt == 2:
            cov = coverage(d, o + u16(d, o + 2))
            cd = classdef(d, o + u16(d, o + 4))
            n = u16(d, o + 6)
            sets = []
            for i in range(n):
                so = u16(d, o + 8 + 2 * i)
                rules = []
                if so:
                    so += o
                    for j in range(u16(d, so)):
                        ro = so + u16(d, so + 2 + 2 * j)
                        gc, sc = struct.unpack_from(">HH", d, ro)
                        inp = tuple(struct.unpack_from(f">{gc - 1}H", d, ro + 4))
                        rules.append(((), inp, (), self._records(ro + 4 + 2 * (gc - 1), sc)))
                sets.append(rules)
            return ("ctx", "c", cov, cd, sets, None, None)
        if fmt == 3:
            gc, sc = struct.unpack_from(">HH", d, o + 2)
            covs = [set(coverage(d, o + x)) for x in struct.unpack_from(f">{gc}H", d, o + 6)]
            rule = ((), tuple(covs[1:]), (), self._records(o + 6 + 2 * gc, sc))
            return ("ctx", "v", covs[0], None, [[rule]], None, None)
        raise ValueError(f"文脈の形 {fmt} を読めません")

    def _read_chain(self, o: int):
        d = self.d
        fmt = u16(d, o)

        def rule_at(ro):
            nb = u16(d, ro)
            back = tuple(struct.unpack_from(f">{nb}H", d, ro + 2))
            p = ro + 2 + 2 * nb
            ni = u16(d, p)
            inp = tuple(struct.unpack_from(f">{ni - 1}H", d, p + 2))
            p += 2 + 2 * (ni - 1)
            nl = u16(d, p)
            look = tuple(struct.unpack_from(f">{nl}H", d, p + 2))
            p += 2 + 2 * nl
            ns = u16(d, p)
            return (back, inp, look, self._records(p + 2, ns))

        if fmt in (1, 2):
            cov = coverage(d, o + u16(d, o + 2))
            if fmt == 1:
                n = u16(d, o + 4)
                base = o + 6
                cds = None
            else:
                cds = (classdef(d, o + u16(d, o + 4)), classdef(d, o + u16(d, o + 6)), classdef(d, o + u16(d, o + 8)))
                n = u16(d, o + 10)
                base = o + 12
            sets = []
            for i in range(n):
                so = u16(d, base + 2 * i)
                rules = []
                if so:
                    so += o
                    for j in range(u16(d, so)):
                        rules.append(rule_at(so + u16(d, so + 2 + 2 * j)))
                sets.append(rules)
            if fmt == 1:
                return ("ctx", "g", cov, None, sets, None, None)
            return ("ctx", "c", cov, cds[1], sets, cds[0], cds[2])
        if fmt == 3:
            p = o + 2
            nb = u16(d, p)
            back = tuple(set(coverage(d, o + x)) for x in struct.unpack_from(f">{nb}H", d, p + 2))
            p += 2 + 2 * nb
            ni = u16(d, p)
            inp = [set(coverage(d, o + x)) for x in struct.unpack_from(f">{ni}H", d, p + 2)]
            p += 2 + 2 * ni
            nl = u16(d, p)
            look = tuple(set(coverage(d, o + x)) for x in struct.unpack_from(f">{nl}H", d, p + 2))
            p += 2 + 2 * nl
            ns = u16(d, p)
            rule = (back, tuple(inp[1:]), look, self._records(p + 2, ns))
            return ("ctx", "v", inp[0], None, [[rule]], None, None)
        raise ValueError(f"連鎖文脈の形 {fmt} を読めません")


def _vsize(vf: int) -> int:
    return 2 * bin(vf & 0xFF).count("1")


def _value(d: bytes, p: int, vf: int) -> tuple[int, int, int, int]:
    v = [0, 0, 0, 0]
    for bit in range(8):
        if vf & (1 << bit):
            if bit < 4:
                v[bit] = s16(d, p)
            p += 2
    return (v[0], v[1], v[2], v[3])


def _anchor(d: bytes, o: int) -> tuple[int, int]:
    return (s16(d, o + 2), s16(d, o + 4))


# ====================================================================== 当てる
class Shaper:
    """1 つの書体（字形の数・GDEF・GSUB・GPOS）に、段階ごとに参照を当てる。"""

    def __init__(self, gdef: GDEF, gsub: Layout, gpos: Layout, num_glyphs: int):
        self.gdef = gdef
        self.gsub = gsub
        self.gpos = gpos
        self.num = num_glyphs
        self._lig_id = 0
        self.rtl = False

    # ------------------------------------------------------------------ 字形の分類
    def set_class(self, info: Info, synth_mark: bool | None = None) -> None:
        g = self.gdef
        if g.present and g.glyph_class:
            info.gclass = g.glyph_class.get(info.gid, 0) or BASE
            if info.gclass not in (BASE, LIGATURE, MARK, COMPONENT):
                info.gclass = BASE
        elif synth_mark is not None:
            info.gclass = MARK if synth_mark else BASE
        info.mclass = g.mark_class.get(info.gid, 0) if info.gclass == MARK else 0

    def _ignored(self, info: Info, lk: Lookup, flag: int | None = None) -> bool:
        f = lk.flag if flag is None else flag
        c = info.gclass
        if (f & 0x2 and c == BASE) or (f & 0x4 and c == LIGATURE) or (f & 0x8 and c == MARK):
            return True
        if c == MARK:
            if f & 0x10:
                sets = self.gdef.mark_sets
                return lk.mfs is None or lk.mfs >= len(sets) or info.gid not in sets[lk.mfs]
            mat = (f & 0xFF00) >> 8
            if mat and info.mclass != mat:
                return True
        return False

    def _seek(self, buf: list[Info], start: int, step: int, lk: Lookup, ctx: dict, context: bool, pred,
              flag: int | None = None) -> int | None:
        """start から step の向きに、飛ばす字形を飛ばして、pred に合う字形の位置（合わなければ None）。
        HarfBuzz の skipping_iterator と同じ：参照の旗で無視する字形は飛ばす。ZWJ・ZWNJ は、合えば合わせ、
        合わなければ、飛ばしてよいときだけ飛ばす。"""
        k = start
        gpos = not self.gsub_mode
        ign_zwnj = gpos or (context and ctx["auto_zwnj"])
        ign_zwj = context or ctx["auto_zwj"]
        mask = -1 if context else ctx["mask"]
        syl = ctx["syl"]
        while 0 <= k < len(buf):
            inf = buf[k]
            if self._ignored(inf, lk, flag):
                k += step
                continue
            ok = (inf.mask & mask) and (not syl or inf.syl == syl) and (pred is None or pred(inf))
            if ok:
                return k
            if (inf.cp == ZWNJ and ign_zwnj) or (inf.cp == ZWJ and ign_zwj):
                k += step
                continue
            return None
        return None

    # ------------------------------------------------------------------ 段階
    def run(self, buf: list[Info], stages: list[Stage], gsub: bool) -> list[Info]:
        self.gsub_mode = gsub
        table = self.gsub if gsub else self.gpos
        for st in stages:
            if st.pause is not None and gsub:
                buf = st.pause(buf) or buf
            for idx, mask, azwj, azwnj, per_syl, tag in st.items:
                lk = table.lookup(idx)
                ctx = {"mask": mask, "auto_zwj": azwj, "auto_zwnj": azwnj, "per_syl": per_syl, "syl": 0, "tag": tag}
                if gsub and lk.type == 8:
                    for i in range(len(buf) - 1, -1, -1):
                        inf = buf[i]
                        if inf.mask & mask and not self._ignored(inf, lk):
                            ctx["syl"] = inf.syl if per_syl else 0
                            self._reverse(buf, i, lk, ctx)
                    continue
                i = 0
                while i < len(buf):
                    inf = buf[i]
                    if inf.mask & mask and not self._ignored(inf, lk):
                        ctx["syl"] = inf.syl if per_syl else 0
                        before = len(buf)
                        nxt = self.apply_at(buf, i, lk, ctx)
                        if nxt is not None and (nxt > i or len(buf) < before):
                            i = nxt
                            continue
                    i += 1
        return buf

    def apply_at(self, buf: list[Info], i: int, lk: Lookup, ctx: dict) -> int | None:
        """参照 lk を位置 i に当てる（どれかの部分の表が当たれば、次に見る位置）。"""
        for sub in lk.subs:
            r = self._apply_sub(buf, i, sub, lk, ctx)
            if r is not None:
                return r
        return None

    def _apply_sub(self, buf, i, sub, lk, ctx):
        kind = sub[0]
        inf = buf[i]
        g = inf.gid
        if kind == "single":
            ng = sub[1].get(g)
            if ng is None:
                return None
            self._replace(inf, ng, ctx)
            return i + 1
        if kind == "multiple":
            seq = sub[1].get(g)
            if seq is None:
                return None
            if not seq:
                del buf[i]
                return i
            if len(seq) == 1:
                self._replace(inf, seq[0], ctx)
                return i + 1
            out = []
            for k, ng in enumerate(seq):
                n = inf.copy()
                n.gid = ng
                n.flags = inf.flags | SUBSTITUTED | MULTIPLIED
                n.lig_id = 0
                n.lig_comp = k
                n.chars = inf.chars if k == 0 else ""
                n.feat.add(ctx["tag"])
                self.set_class(n)
                out.append(n)
            buf[i:i + 1] = out
            return i + len(seq)
        if kind == "ligature":
            sets = sub[1].get(g)
            if sets is None:
                return None
            for lg, comps in sets:
                pos = self._match_input(buf, i, comps, lambda inf, m: inf.gid == m, lk, ctx)
                if pos is None:
                    continue
                return self._ligate(buf, pos, lg, ctx)
            return None
        if kind == "alternate":
            return None
        if kind == "ctx":
            return self._apply_ctx(buf, i, sub, lk, ctx)
        # GPOS
        if kind == "psingle":
            v = sub[1].get(g)
            if v is None:
                return None
            self._add(inf, v)
            return i + 1
        if kind == "pair1":
            m = sub[1].get(g)
            if m is None:
                return None
            j = self._seek(buf, i + 1, 1, lk, ctx, False, None)
            if j is None or buf[j].gid not in m:
                return None
            v1, v2 = m[buf[j].gid]
            self._add(inf, v1)
            self._add(buf[j], v2)
            return j + 1 if sub[2] else j
        if kind == "pair2":
            _k, cov, cd1, cd2, rows, vf2 = sub
            if g not in cov:
                return None
            j = self._seek(buf, i + 1, 1, lk, ctx, False, None)
            if j is None:
                return None
            c1, c2 = cd1.get(g, 0), cd2.get(buf[j].gid, 0)
            if c1 >= len(rows) or c2 >= len(rows[c1]):
                return None
            v1, v2 = rows[c1][c2]
            self._add(inf, v1)
            self._add(buf[j], v2)
            return j + 1 if vf2 else j
        if kind == "cursive":
            rec = sub[1].get(g)
            if rec is None or rec[0] is None:
                return None
            j = self._seek(buf, i - 1, -1, lk, ctx, False, None)
            if j is None:
                return None
            prev = sub[1].get(buf[j].gid)
            if prev is None or prev[1] is None:
                return None
            self._cursive(buf, j, i, prev[1], rec[0], lk)
            return i + 1
        if kind == "markbase":
            mk = sub[1].get(g)
            if mk is None:
                return None
            j = self._prev_base(buf, i, lk, ctx)
            if j is None or buf[j].gid not in sub[2]:
                return None
            anc = sub[2][buf[j].gid][mk[0]] if mk[0] < len(sub[2][buf[j].gid]) else None
            if anc is None:
                return None
            self._attach(buf, i, j, mk[1], anc)
            return i + 1
        if kind == "marklig":
            mk = sub[1].get(g)
            if mk is None:
                return None
            j = self._prev_base(buf, i, lk, ctx)
            if j is None or buf[j].gid not in sub[2]:
                return None
            comps = sub[2][buf[j].gid]
            if not comps:
                return None
            lig, mark = buf[j], inf
            if lig.lig_id and lig.lig_id == mark.lig_id and mark.lig_comp > 0:
                ci = min(len(comps), mark.lig_comp) - 1
            else:
                ci = len(comps) - 1
            row = comps[ci]
            anc = row[mk[0]] if mk[0] < len(row) else None
            if anc is None:
                return None
            self._attach(buf, i, j, mk[1], anc)
            return i + 1
        if kind == "markmark":
            mk = sub[1].get(g)
            if mk is None:
                return None
            j = self._seek(buf, i - 1, -1, lk, ctx, False, None, flag=lk.flag & ~0xE)
            if j is None or not buf[j].is_mark:
                return None
            a, b = inf, buf[j]
            good = False
            if a.lig_id == b.lig_id:
                good = a.lig_id == 0 or a.lig_comp == b.lig_comp
            elif (a.lig_id > 0 and not a.lig_comp) or (b.lig_id > 0 and not b.lig_comp):
                good = True
            if not good or b.gid not in sub[2]:
                return None
            row = sub[2][b.gid]
            anc = row[mk[0]] if mk[0] < len(row) else None
            if anc is None:
                return None
            self._attach(buf, i, j, mk[1], anc)
            return i + 1
        raise ValueError(f"参照 {kind} を当てられません")

    # ------------------------------------------------------------------ GSUB の部品
    def _replace(self, inf: Info, ng: int, ctx: dict) -> None:
        inf.gid = ng
        inf.flags |= SUBSTITUTED
        inf.feat.add(ctx["tag"])
        self.set_class(inf)

    def _match_input(self, buf, i, seq, eq, lk, ctx) -> list[int] | None:
        pos = [i]
        first = buf[i]
        k = i
        for m in seq:
            k = self._seek(buf, k + 1, 1, lk, ctx, False, lambda inf, m=m: eq(inf, m))
            if k is None:
                return None
            this = buf[k]
            if first.lig_id and first.lig_comp:
                if first.lig_id != this.lig_id or first.lig_comp != this.lig_comp:
                    return None
            elif this.lig_id and this.lig_comp and this.lig_id != first.lig_id:
                return None
            pos.append(k)
        return pos

    @staticmethod
    def _ncomp(inf: Info) -> int:
        return inf.extra if inf.gclass == LIGATURE and isinstance(inf.extra, int) else 1

    def _ligate(self, buf: list[Info], pos: list[int], lg: int, ctx: dict) -> int:
        """HarfBuzz の ligate_input と同じ：合字の部品の間・後ろの記号に、どの部品に付くか（lig_id・lig_comp）を覚えさせる。"""
        first = buf[pos[0]]
        comps = [buf[p] for p in pos]
        is_base_lig = first.gclass == BASE and all(c.is_mark for c in comps[1:])
        is_mark_lig = all(c.is_mark for c in comps)
        is_lig = not is_base_lig and not is_mark_lig
        total = sum(self._ncomp(c) for c in comps)
        lig_id = 0
        if is_lig:
            self._lig_id = (self._lig_id % 7) + 1
            lig_id = self._lig_id
        last_lig_id, last_num = first.lig_id, self._ncomp(first)
        so_far = last_num
        first.gid = lg
        first.flags = (first.flags | SUBSTITUTED | LIGATED) & ~MULTIPLIED
        first.chars = "".join(c.chars for c in comps)
        first.cluster = min(c.cluster for c in comps)
        first.feat.add(ctx["tag"])
        self.set_class(first)
        if not (self.gdef.present and self.gdef.glyph_class) and is_lig:
            first.gclass = LIGATURE
        if is_lig:
            first.lig_id = lig_id
            first.lig_comp = 0
            first.extra = total
        for a, b in zip(pos, pos[1:]):
            if not is_mark_lig:
                for k in range(a + 1, b):
                    if is_lig:
                        mk = buf[k]
                        this = mk.lig_comp or last_num
                        mk.lig_id = lig_id
                        mk.lig_comp = so_far - last_num + min(this, last_num)
            last_lig_id, last_num = buf[b].lig_id, self._ncomp(buf[b])
            so_far += last_num
        if not is_mark_lig and last_lig_id:
            for k in range(pos[-1] + 1, len(buf)):
                mk = buf[k]
                if mk.lig_id != last_lig_id or not mk.lig_comp:
                    break
                mk.lig_id = lig_id
                mk.lig_comp = so_far - last_num + min(mk.lig_comp, last_num)
        for p in reversed(pos[1:]):
            del buf[p]
        return pos[-1] + 1 - (len(pos) - 1)

    def _apply_ctx(self, buf, i, sub, lk, ctx):
        _k, kind, cov, cd, sets, bcd, lcd = sub
        g = buf[i].gid
        if g not in cov:
            return None
        if kind == "g":
            rules = sets[cov[g]] if cov[g] < len(sets) else []
            eq_in = eq_b = eq_l = (lambda inf, m: inf.gid == m)
        elif kind == "c":
            c = cd.get(g, 0)
            rules = sets[c] if c < len(sets) else []
            eq_in = (lambda inf, m: cd.get(inf.gid, 0) == m)
            bc = bcd if bcd is not None else cd
            lc = lcd if lcd is not None else cd
            eq_b = (lambda inf, m: bc.get(inf.gid, 0) == m)
            eq_l = (lambda inf, m: lc.get(inf.gid, 0) == m)
        else:
            rules = sets[0]
            eq_in = eq_b = eq_l = (lambda inf, m: inf.gid in m)
        for back, inp, look, recs in rules:
            pos = self._match_input(buf, i, inp, eq_in, lk, ctx)
            if pos is None:
                continue
            k = pos[-1]
            ok = True
            for m in look:
                k = self._seek(buf, k + 1, 1, lk, ctx, True, lambda inf, m=m: eq_l(inf, m))
                if k is None:
                    ok = False
                    break
            if not ok:
                continue
            k = i
            for m in back:
                k = self._seek(buf, k - 1, -1, lk, ctx, True, lambda inf, m=m: eq_b(inf, m))
                if k is None:
                    ok = False
                    break
            if not ok:
                continue
            return self._apply_records(buf, pos, recs, ctx)
        return None

    def _apply_records(self, buf, pos, recs, ctx) -> int:
        table = self.gsub if self.gsub_mode else self.gpos
        n0 = len(buf)
        end = pos[-1] + 1
        pos = list(pos)
        for seq, li in recs:
            if seq >= len(pos):
                continue
            p = pos[seq]
            if p >= len(buf):
                continue
            before = len(buf)
            self.apply_at(buf, p, table.lookup(li), ctx)
            delta = len(buf) - before
            if delta == 0:
                continue
            if delta > 0:
                pos = pos[:seq + 1] + [p + 1 + k for k in range(delta)] + [x + delta for x in pos[seq + 1:]]
            else:
                delta = max(delta, (seq + 1) - len(pos))
                pos = pos[:seq + 1] + [x + delta for x in pos[seq + 1 - delta:]]
        return end + (len(buf) - n0)

    def _reverse(self, buf, i, lk, ctx) -> None:
        for sub in lk.subs:
            _k, m, back, look = sub
            g = buf[i].gid
            if g not in m:
                continue
            k = i
            ok = True
            for cv in look:
                k = self._seek(buf, k + 1, 1, lk, ctx, True, lambda inf, cv=cv: inf.gid in cv)
                if k is None:
                    ok = False
                    break
            if not ok:
                continue
            k = i
            for cv in back:
                k = self._seek(buf, k - 1, -1, lk, ctx, True, lambda inf, cv=cv: inf.gid in cv)
                if k is None:
                    ok = False
                    break
            if not ok:
                continue
            self._replace(buf[i], m[g], ctx)
            return

    # ------------------------------------------------------------------ GPOS の部品
    @staticmethod
    def _add(inf: Info, v) -> None:
        inf.xo += v[0]
        inf.yo += v[1]
        inf.xa += v[2]
        inf.ya += v[3]

    def _prev_base(self, buf, i, lk, ctx) -> int | None:
        """記号の前の、記号でない字形（複数の置き換えの 2 つ目以降は飛ばす：HarfBuzz の MarkBasePos と同じ）。"""
        k = i - 1
        while True:
            k = self._seek(buf, k, -1, lk, ctx, False, None, flag=0x8)
            if k is None:
                return None
            inf = buf[k]
            if (not inf.multiplied or inf.lig_comp == 0 or k == 0 or buf[k - 1].is_mark or not buf[k - 1].multiplied
                    or inf.lig_id != buf[k - 1].lig_id or inf.lig_comp != buf[k - 1].lig_comp + 1):
                return k
            k -= 1

    @staticmethod
    def _attach(buf, i, j, mark_anchor, base_anchor) -> None:
        m = buf[i]
        m.xo = base_anchor[0] - mark_anchor[0]
        m.yo = base_anchor[1] - mark_anchor[1]
        m.attach = j
        m.atype = 1

    def _cursive(self, buf, i, j, exit_a, entry_a, lk) -> None:
        """i：前の字形（出口）、j：今の字形（入口）。HarfBuzz の CursivePosFormat1 と同じ。"""
        pi, pj = buf[i], buf[j]
        if not self.rtl:
            pi.xa = exit_a[0] + pi.xo
            d = entry_a[0] + pj.xo
            pj.xa -= d
            pj.xo -= d
        else:
            d = exit_a[0] + pi.xo
            pi.xa -= d
            pi.xo -= d
            pj.xa = entry_a[0] + pj.xo
        child, parent = i, j
        y = entry_a[1] - exit_a[1]
        if not lk.flag & 0x1:
            child, parent = parent, child
            y = -y
        buf[child].atype = 2
        buf[child].attach = parent
        buf[child].yo = y

    # ------------------------------------------------------------------ 確かめるための当てはめ
    def would_substitute(self, lookups: list[int], gids: list[int]) -> bool:
        """機能の参照のどれかが、この字形の並びだけ（前後の文脈なし）を置き換えるか（HarfBuzz の would_apply）。"""
        for li in lookups:
            lk = self.gsub.lookup(li)
            for sub in lk.subs:
                kind = sub[0]
                g = gids[0]
                if kind in ("single", "multiple", "alternate") and len(gids) == 1:
                    if kind != "alternate" and g in sub[1]:
                        return True
                elif kind == "ligature":
                    for _lg, comps in sub[1].get(g, []):
                        if list(comps) == list(gids[1:]):
                            return True
                elif kind == "ctx":
                    _k, kd, cov, cd, sets, _b, _l = sub
                    if g not in cov:
                        continue
                    if kd == "g":
                        rules = sets[cov[g]] if cov[g] < len(sets) else []
                        eq = (lambda x, m: x == m)
                    elif kd == "c":
                        c = cd.get(g, 0)
                        rules = sets[c] if c < len(sets) else []
                        eq = (lambda x, m, cd=cd: cd.get(x, 0) == m)
                    else:
                        rules = sets[0]
                        eq = (lambda x, m: x in m)
                    for back, inp, look, _recs in rules:
                        if back or look or len(inp) != len(gids) - 1:
                            continue
                        if all(eq(x, m) for x, m in zip(gids[1:], inp)):
                            return True
        return False


def positions(buf: list[Info], rtl: bool, zero_marks: bool) -> tuple[list[tuple[Info, float, float]], int]:
    """GPOS の後：字形を見た目の順（左から右）に並べ、字形ごとの原点（字形の単位）と、並びの幅。"""
    if zero_marks:
        for inf in buf:
            if inf.is_mark:
                inf.xa = 0
                inf.ya = 0
    order = list(range(len(buf)))
    if rtl:
        order.reverse()
    pen = 0
    pen_at = {}
    for k in order:
        pen_at[k] = pen
        pen += buf[k].xa
    width = pen
    absx: dict[int, float] = {}
    absy: dict[int, float] = {}

    def resolve(k, depth=0):
        if k in absx:
            return
        inf = buf[k]
        if inf.atype == 1 and 0 <= inf.attach < len(buf) and depth < 32:
            resolve(inf.attach, depth + 1)
            absx[k] = absx[inf.attach] + inf.xo
            absy[k] = absy[inf.attach] + inf.yo
        elif inf.atype == 2 and 0 <= inf.attach < len(buf) and depth < 32:
            resolve(inf.attach, depth + 1)
            absx[k] = pen_at[k] + inf.xo
            absy[k] = absy[inf.attach] + inf.yo
        else:
            absx[k] = pen_at[k] + inf.xo
            absy[k] = inf.yo

    for k in range(len(buf)):
        resolve(k)
    return [(buf[k], absx[k], absy[k]) for k in order], width
