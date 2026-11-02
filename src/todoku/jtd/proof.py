"""一太郎の文書の、字数を変える書き込み・圧縮した文書への書き込みの証明。

書き込み処理（edit.py）の内部の値を使わずに、元のファイルと書き込み後のファイルを読み直して比べる：
1. ストリーム：書き込み後にあるのは「元のストリーム − 取り除いた行・ページの割り付けの控え（LineMark・PageMark・PaperMark）」
   だけ。直すと決めた構造（本文・ストリームの一覧・字を打つ位置・表の位置の表・字数の同じ上書きをしたストリーム）のほかは、
   1 バイトも同じ
2. 本文の Ssmg の入れ物：見出しはブロックの数のほか同じ。書いたセグメントのほかのブロック・目録・後ろのバイトは同じ。
   書いたセグメントは、元のブロックの番号を同じ順に使い、足したブロックは元のブロックの後ろの番号
3. 本文の字の並び：元の単位の並びに、書き込み（置き換え・足す・減らす）だけを当てたものと一致
4. 文字の書式の区切り：書式を変えるバイト列はすべて同じ。数は、書いた所を含む区切りだけが、足した・減らした字の数だけ違う
5. 表の位置の表・字を打つ位置：ずらした位置のほかは同じで、ずらした位置は「元の位置より前で足した・減らした字の数」だけ違う
6. ストリームの一覧：見出しは項目の数のほか同じ。項目は、取り除いたストリームの項目がないことと、大きさの変わったストリームの
   大きさが今の大きさであることのほか同じ
7. 複合ファイルの入れ物：見出し（FAT・ミニ FAT の数と位置のほか）・ディレクトリの項目（名前・種類・日時・CLSID）は同じ。
   変わったセクタは、表（FAT・ミニ FAT・ディレクトリ）と、中身の変わったストリームのセクタ・ミニセクタと、空いて 0 で
   消したセクタだけ。取り除いたストリームの元のデータは残っていない
8. 読み直し：書き込み後の文書を読み直すと、段落の数は同じで、どの段落の文字も「元の文字に、書いた所だけを当てたもの」
圧縮した文書は、外側の入れ物も同じように比べ、JSCompDocument を解いて（CRC-16 と長さで確かめて）中の文書を比べる。
"""

from __future__ import annotations

import struct

from ..cfb import CFB, ENDOFCHAIN, FREESECT, NOSTREAM, name_key
from ..docx.document import iter_paras
from .book import SSMG, TEXT, JtdError, JtdFile, para_units, ssmg

LAYOUT = ("LineMark", "PageMark", "PaperMark")


def _d(name: str) -> str:
    return name.replace("\x04", "").replace("\x05", "")


# ====================================================================== 位置の対応（証明のための、別の書き方）
def _shift(ops, p: int, unit: bool = False) -> int | None:
    """元の単位の位置 p → 書き込み後の位置（足した字は、足した所より後ろの位置を後ろへ。減らした範囲の中は None）。
    unit=True は「その単位」を指す値（足した所にちょうど当たる値も、その単位とともに後ろへ）。"""
    out = p
    for op in ops:
        keep = min(len(op.text), op.b - op.a)
        if len(op.text) > op.b - op.a:
            at = op.b if op.b > op.a else op.a
            if p > at or (unit and p == at):
                out += len(op.text) - (op.b - op.a)
        elif len(op.text) < op.b - op.a:
            lo, hi = op.a + keep, op.b
            if lo < p < hi:
                return None
            if p >= hi:
                out -= hi - lo
    return out


def _apply_units(units, ops) -> list[int]:
    out: list[int] = []
    i = 0
    for op in sorted(ops, key=lambda o: o.a):
        out += units[i:op.a]
        out += [ord(c) for c in op.text]
        i = op.b
    return out + list(units[i:])


def _runs(rest: bytes):
    out, pos = [], 0
    while pos < len(rest):
        if rest[pos] == 0:
            out.append(("n", struct.unpack_from(">I", rest, pos + 1)[0]))
            pos += 5
        elif rest[pos] == 0xFE:
            st = pos
            pos += 1
            while rest[pos] != 0xFF:
                pos += 2 + rest[pos + 1]
            pos += 2
            out.append(("a", rest[st:pos]))
        else:
            raise JtdError("文字の書式の区切りの形が違います")
    return out


def _ssmg_raw(data: bytes):
    count, bsize, nblocks = struct.unpack_from(">III", data, 8)
    pos = 20 + nblocks * bsize
    dirs = []
    for _ in range(count):
        sid, z, ln, n, n2 = struct.unpack_from(">IIIII", data, pos)
        dirs.append((sid, z, ln, n, n2, list(struct.unpack_from(f">{n}I", data, pos + 20))))
        pos += 20 + 4 * n
    return count, bsize, nblocks, dirs, data[pos:]


# ====================================================================== 入れ物（複合ファイル）
def check_container(x: bytes, y: bytes, changed: set[str], removed: set[str]) -> list[str]:
    """複合ファイルの入れ物の比較。changed：中身の変わってよいストリーム、removed：取り除いたストリーム。"""
    bad: list[str] = []
    ca, cb = CFB(x), CFB(y)
    ss, ms = ca.ssize, ca.msize
    if (cb.ssize, cb.msize, cb.cutoff, cb.major) != (ca.ssize, ca.msize, ca.cutoff, ca.major):
        return ["複合ファイルのセクタの大きさ・版が変わりました"]
    # 見出し：FAT の数（0x2C）・ミニ FAT の位置と数（0x3C〜0x43）・FAT の位置の一覧（0x4C〜）のほかは同じ
    for lo, hi in ((0, 0x2C), (0x30, 0x3C), (0x44, 0x4C)):
        if x[lo:hi] != y[lo:hi]:
            bad.append(f"複合ファイルの見出し（{lo:#x}〜{hi:#x}）が変わりました")
    # ディレクトリ
    if len(cb.entries) != len(ca.entries):
        bad.append("ディレクトリの項目の数が変わりました")
        return bad
    rem_idx = {ca._find(n).index for n in removed}
    chg_idx = {ca._find(n).index for n in changed if ca._find(n) is not None}
    da = ca._chain_bytes(ca.dir_start)
    db = cb._chain_bytes(cb.dir_start)
    parents = set()
    for i in rem_idx:
        for p in range(len(ca.entries)):
            todo = [ca.entries[p].child] if ca.entries[p].kind in (1, 5) else []
            seen = set()
            while todo:
                k = todo.pop()
                if k == NOSTREAM or k in seen or k >= len(ca.entries):
                    continue
                seen.add(k)
                if k == i:
                    parents.add(p)
                todo += [ca.entries[k].left, ca.entries[k].right]
    rebuilt = set()
    for p in parents:
        todo, seen = [cb.entries[p].child], set()
        while todo:
            k = todo.pop()
            if k == NOSTREAM or k in seen or k >= len(cb.entries):
                continue
            seen.add(k)
            rebuilt.add(k)
            todo += [cb.entries[k].left, cb.entries[k].right]
    for i in range(len(ca.entries)):
        ea, eb = da[i * 128:(i + 1) * 128], db[i * 128:(i + 1) * 128]
        if i in rem_idx:
            if eb[:68] != bytes(68) or struct.unpack_from("<III", eb, 68) != (NOSTREAM,) * 3 or eb[80:] != bytes(48):
                bad.append(f"取り除いたストリーム {_d(ca.entries[i].name)!r} のディレクトリの項目が、空きになっていません")
            continue
        if ea[:67] != eb[:67] or ea[80:116] != eb[80:116]:
            bad.append(f"ディレクトリの項目 {_d(ca.entries[i].name)!r} の名前・種類・日時などが変わりました")
        if ea[67] != eb[67] and i not in rebuilt:
            bad.append(f"ディレクトリの項目 {_d(ca.entries[i].name)!r} の色が変わりました")
        if ea[68:76] != eb[68:76] and i not in rebuilt:
            bad.append(f"ディレクトリの項目 {_d(ca.entries[i].name)!r} の木のつながりが変わりました")
        if ea[76:80] != eb[76:80] and i not in parents:
            bad.append(f"ディレクトリの項目 {_d(ca.entries[i].name)!r} の子のつながりが変わりました")
        if ea[116:] != eb[116:] and i not in chg_idx and i != 0:
            bad.append(f"ストリーム {_d(ca.entries[i].name)!r} の位置・大きさが変わりました")
    # 作り直した木：名前の順の二分探索木で、赤黒木の決まり（赤の子は黒・どの道も黒の数が同じ）を満たす
    for p in parents:
        order: list[int] = []

        def walk(k: int, depth: int) -> int:
            if k == NOSTREAM:
                return 1
            if depth > 64:
                raise ValueError
            e = cb.entries[k]
            lb = walk(e.left, depth + 1)
            order.append(k)
            rb = walk(e.right, depth + 1)
            red = db[k * 128 + 67] == 0
            for c in (e.left, e.right):
                if red and c != NOSTREAM and db[c * 128 + 67] == 0:
                    raise ValueError
            if lb != rb or lb < 0:
                raise ValueError
            return lb + (0 if red else 1)

        try:
            walk(cb.entries[p].child, 0)
            keys = [name_key(cb.entries[k].name) for k in order]
            if keys != sorted(keys) or len(set(keys)) != len(keys):
                bad.append("作り直したディレクトリの木が、名前の順になっていません")
        except (ValueError, RecursionError):
            bad.append("作り直したディレクトリの木が、赤黒木の決まりを満たしません")
    # セクタ：変わったセクタは、表と、中身の変わったストリームのものと、0 で消した空きのセクタだけ
    own: set[int] = set(cb.fat_sectors) | set(cb.chain(cb.dir_start))
    if cb.minifat_start not in (ENDOFCHAIN, FREESECT):
        own |= set(cb.chain(cb.minifat_start))
    root_b = cb.chain(cb.entries[0].start) if cb.entries[0].start not in (ENDOFCHAIN, FREESECT) else []
    for n in changed:
        e = cb._find(n)
        if e is not None and e.size >= cb.cutoff:
            own |= set(cb.chain(e.start))
    na, nb = len(x) // ss - 1, len(y) // ss - 1
    if nb < na:
        bad.append("複合ファイルが短くなりました")
    mini_ok: set[int] = set()
    for n in changed:
        e = cb._find(n)
        if e is not None and e.size < cb.cutoff:
            s, k = e.start, 0
            while s not in (ENDOFCHAIN, FREESECT) and s < len(cb.minifat) and k <= len(cb.minifat):
                mini_ok.add(s)
                s = cb.minifat[s]
                k += 1
    for s in range(nb):
        sa_ = x[(s + 1) * ss:(s + 2) * ss] if s < na else None
        sb_ = y[(s + 1) * ss:(s + 2) * ss]
        if sa_ == sb_ or s in own:
            continue
        if s in root_b:
            # ミニストリームのセクタ：変わったミニセクタは、中身の変わったストリームのものか、0 で消した空きのもの
            base = root_b.index(s) * (ss // ms)
            for j in range(ss // ms):
                m = base + j
                xa = sa_[j * ms:(j + 1) * ms] if sa_ is not None else None
                xb = sb_[j * ms:(j + 1) * ms]
                if xa == xb or m in mini_ok:
                    continue
                if m < len(cb.minifat) and cb.minifat[m] == FREESECT and xb == bytes(ms):
                    continue
                if m >= len(cb.minifat):
                    continue  # ミニストリームの終わりより後ろ
                bad.append(f"ミニセクタ {m} が、書き込みと関係なく変わりました")
                break
            continue
        if s < len(cb.fat) and cb.fat[s] == FREESECT and sb_ == bytes(ss):
            continue
        bad.append(f"セクタ {s} が、書き込みと関係なく変わりました")
    # 取り除いた・中身の変わったストリームの元のデータが、空きのセクタに残っていない
    for n in removed | changed:
        e = ca._find(n)
        if e is None:
            continue
        if e.size >= ca.cutoff:
            for s in ca.chain(e.start):
                if s < len(cb.fat) and cb.fat[s] == FREESECT and y[(s + 1) * ss:(s + 2) * ss] != bytes(ss):
                    bad.append(f"{_d(n)!r} の元のデータが、空きのセクタ {s} に残っています")
                    break
    for n in removed:
        if cb.has(n):
            bad.append(f"取り除いたはずのストリーム {_d(n)!r} が残っています")
    return bad


# ====================================================================== 本文の比較
def check_text(pr, a: JtdFile, b: JtdFile, ops, edits) -> list[str]:
    """本文の Ssmg・字の並び・文字の書式の区切りの比較。"""
    bad: list[str] = []
    x, y = a.cfb.read("DocumentText"), b.cfb.read("DocumentText")
    cx, bx, nx, dx, tx = _ssmg_raw(x)
    cy, by, ny, dy, ty = _ssmg_raw(y)
    if (cx, bx) != (cy, by) or x[:16] != y[:16] or tx != ty:
        bad.append("本文の Ssmg の見出し（セグメントの数・ブロックの大きさ）か、目録の後ろが変わりました")
        return bad
    k0 = ops[0].k
    sid = a.texts[k0].sid
    used_other = set()
    for ea, eb in zip(dx, dy):
        if ea[0] != eb[0]:
            bad.append("本文のセグメントの順が変わりました")
            return bad
        if ea[0] == sid:
            if eb[5][:len(ea[5])] != ea[5] or any(b_ < nx for b_ in eb[5][len(ea[5]):]) or len(set(eb[5])) != len(eb[5]):
                bad.append("書いたセグメントのブロックの番号が、元の並びのままでないか、足したブロックが元のブロックと重なります")
            if ny != nx + len(eb[5]) - len(ea[5]):
                bad.append(f"本文のブロックの数（{nx} → {ny}）が、足したブロックの数と合いません")
        else:
            if ea != eb:
                bad.append(f"本文のセグメント {ea[0]} の目録が変わりました")
            used_other |= set(ea[5])
    for blk in used_other:
        if x[20 + blk * bx:20 + (blk + 1) * bx] != y[20 + blk * bx:20 + (blk + 1) * bx]:
            bad.append(f"本文の、書いていないセグメントのブロック {blk} が変わりました")
            break
    sa_ = {s.sid: s.data for s in ssmg(x, "DocumentText")}
    sb_ = {s.sid: s.data for s in ssmg(y, "DocumentText")}
    da, db = sa_[sid], sb_[sid]
    if da[:8] != TEXT or db[:8] != TEXT:
        bad.append("書いたセグメントが字の並びではありません")
        return bad
    n1 = struct.unpack_from(">I", da, 8)[0]
    n2 = struct.unpack_from(">I", db, 8)[0]
    ua = struct.unpack_from(f">{n1}H", da, 12)
    ub = list(struct.unpack_from(f">{n2}H", db, 12))
    mine = [op for op in ops if op.k == k0] + [_Op(k, i, i + 1, ch) for (k, i), ch in edits.items() if k == k0]
    want = _apply_units(ua, mine)
    if ub != want:
        first = next((i for i, (p, q) in enumerate(zip(ub, want)) if p != q), min(len(ub), len(want)))
        bad.append(f"本文の字の並びが、元の並びに書き込みを当てたものと違います（単位 {first} から）")
    ra, rb = _runs(da[12 + 2 * n1:]), _runs(db[12 + 2 * n2:])
    if [k for k, _ in ra] != [k for k, _ in rb] or any(p != q for (k, p), (_k, q) in zip(ra, rb) if k == "a"):
        bad.append("文字の書式（書式を変えるバイト列）か、区切りの並びが変わりました")
        return bad
    # 区切りの数：足した字は「足した所の前の字（前の書式を引き継ぐとき）か、足した所」を含む区切りに、減らした字は
    # その字を含んでいた区切りから
    bounds, pos = [], 0
    for j, (k, v) in enumerate(ra):
        if k == "n":
            bounds.append((j, pos, pos + v))
            pos += v

    def at(i: int):
        return next((j for j, lo, hi in bounds if lo <= i < hi), None)

    exp = {j: v for j, (k, v) in enumerate(ra) if k == "n"}
    for op in ops:
        d = len(op.text) - (op.b - op.a)
        if d > 0:
            q = op.b if op.b > op.a else op.a
            j = at(q - 1 if op.inherit == "prev" and q else q)
            if op.inherit == "both" and at(q - 1 if q else 0) != at(q):
                bad.append(f"{op.key}: 書式の切り替わる所に字を足しています")
            if j is not None:
                exp[j] += d
        elif d < 0:
            for i in range(op.a + len(op.text), op.b):
                j = at(i)
                if j is not None:
                    exp[j] -= 1
    got = {j: v for j, (k, v) in enumerate(rb) if k == "n"}
    if got != exp:
        bad.append("文字の書式の区切りの数が、書いた所の区切りを足した・減らした字の数だけ変えたものと違います")
    old = {j: v for j, (k, v) in enumerate(ra) if k == "n"}
    pr.structures.append(f"本文の字の並び（{n1:,} → {n2:,} 単位）・書いた所を含む文字の書式の区切りの数"
                         f"（{sum(1 for j in exp if exp[j] != old[j])} か所）・Ssmg の目録（ブロック {nx} → {ny} 個）")
    return bad


class _Op:
    def __init__(self, k, a, b, text):
        self.k, self.a, self.b, self.text, self.key, self.inherit = k, a, b, text, "", "both"


def _tcnt_vals(data: bytes) -> list[tuple[int, int]]:
    out = []
    for s in ssmg(data, "DocumentTextPositionTables"):
        d = s.data
        cnt = struct.unpack_from(">I", d, 10)[0]
        for k in range(cnt):
            for off in (16 + 29 * k, 20 + 29 * k):
                out.append((s.offset(off), struct.unpack_from(">I", d, off)[0]))
    return out


def check_positions(pr, a: JtdFile, b: JtdFile, ops) -> list[str]:
    """表の位置の表・字を打つ位置：ずらした位置のほかは同じ。"""
    bad: list[str] = []
    grow = [op for op in ops if len(op.text) != op.b - op.a]
    if a.cfb.has("DocumentTextPositionTables"):
        x, y = a.cfb.read("DocumentTextPositionTables"), b.cfb.read("DocumentTextPositionTables")
        if len(x) != len(y):
            bad.append("表の位置の表の大きさが変わりました")
        else:
            vals = _tcnt_vals(x)
            allowed = {o + j for o, _v in vals for j in range(4)}
            if any(x[i] != y[i] for i in range(len(x)) if i not in allowed):
                bad.append("表の位置の表の、位置のほかのバイトが変わりました")
            moved = 0
            for o, v in vals:
                w = _shift(grow, v, unit=True)
                got = struct.unpack_from(">I", y, o)[0]
                if w is None or got != w:
                    bad.append(f"表の位置の表の位置 {v} が、{got} になっています（思っていたのは {w}）")
                    break
                moved += got != v
            pr.structures.append(f"表の位置の表（表の範囲の位置 {len(vals)} 個のうち {moved} 個をずらした）")
    if a.cfb.has("DocumentEditStyles"):
        x, y = a.cfb.read("DocumentEditStyles"), b.cfb.read("DocumentEditStyles")
        at = None
        pos = 4
        while pos + 6 <= len(x):
            t, ln = struct.unpack_from(">HI", x, pos)
            if t == 0x2000:
                at = pos + 14
            pos += 6 + ln
        if len(x) != len(y) or any(x[i] != y[i] for i in range(len(x)) if at is None or not at <= i < at + 4):
            bad.append("字を打つ位置（DocumentEditStyles）の、位置のほかのバイトが変わりました")
        elif at is not None:
            v, got = struct.unpack_from(">I", x, at)[0], struct.unpack_from(">I", y, at)[0]
            if v != 0xFFFFFFFF:
                w = _shift(grow, v)
                if w is None:
                    ok = got in {_shift(grow, op.a + len(op.text)) for op in grow if len(op.text) < op.b - op.a}
                else:
                    ok = got == w
                if not ok:
                    bad.append(f"字を打つ位置 {v} が、{got} になっています（思っていたのは {w}）")
                elif got != v:
                    pr.structures.append(f"字を打つ位置（{v} → {got}）")
    return bad


def check_seginfo(pr, x: bytes, y: bytes, cb: CFB, removed: set[str], changed: set[str], label: str = "") -> list[str]:
    """ストリームの一覧：見出しは項目の数のほか同じ、項目は取り除いた項目がないことと大きさのほか同じ。"""
    bad: list[str] = []
    try:
        esize, count = struct.unpack_from("<HH", x, 16)
        total = (len(x) - 48) // 96
        lead = total - count
        if x[:18] != y[:18] or x[20:48 + 96 * lead] != y[20:48 + 96 * lead]:
            return ["ストリームの一覧の見出しが、項目の数のほかに変わりました"]
        slots = [x[48 + 96 * k:48 + 96 * (k + 1)] for k in range(lead, total)]
        names = [s[:64].decode("utf-16-le", errors="replace").split("\x00")[0] for s in slots]
        want = bytearray(x[:48 + 96 * lead])
        n = 0
        resized = []
        for s, name in zip(slots, names):
            if name in removed:
                continue
            item = bytearray(s)
            if name in changed and cb.has(name):
                struct.pack_into("<I", item, 68, cb._find(name).size)
                if item != s:
                    resized.append(name)
            want += item
            n += 1
        struct.pack_into("<H", want, 18, n)
        if bytes(want) != y:
            bad.append("ストリームの一覧が、取り除いた項目と大きさのほかに変わりました")
        gone = [nm for nm in names if nm in removed]
        parts = ([f"項目 {len(gone)} 個（{'・'.join(map(_d, gone))}）を除いた"] if gone else []) \
            + ([f"{'・'.join(map(_d, resized))} の大きさを今の大きさに直した"] if resized else [])
        if parts:
            pr.structures.append(f"{label}ストリームの一覧（" + "・".join(parts) + "）")
    except (struct.error, UnicodeDecodeError) as e:
        bad.append(f"ストリームの一覧を比べられません：{e}")
    return bad


def expected_text(a: JtdFile, p, ops, edits) -> str | None:
    """元の段落の文字に、書き込みだけを当てた文字（書いた所がなければ None）。"""
    us = para_units(a, p)
    touched = False
    out = []
    jend = getattr(p, "jend", None)
    for u, ch in zip(us, p.text):
        if u is None:
            out.append(ch)
            continue
        k, i = u
        for op in ops:
            if op.k == k and op.a == op.b == i and op.inherit == "both":
                out.append(op.text)  # 段落の頭に足す
                touched = True
        repl = None
        dropped = False
        for op in ops:
            if op.k == k and op.a <= i < op.b:
                j = i - op.a
                touched = True
                if j < len(op.text):
                    repl = op.text[j]
                else:
                    dropped = True
        if (k, i) in edits:
            repl = edits[(k, i)]
            touched = True
        if not dropped:
            out.append(repl if repl is not None else ch)
        for op in ops:
            if op.k != k:
                continue
            if op.b > op.a and i == op.b - 1 and len(op.text) > op.b - op.a:
                out.append(op.text[op.b - op.a:])
            elif op.a == op.b and op.inherit == "prev" and i == op.a - 1:
                out.append(op.text)
                touched = True
    if not p.text and jend is not None:
        for op in ops:
            if op.k == jend.k and op.a == op.b == jend.s:
                out.append(op.text)
                touched = True
    return "".join(out) if touched else None


# ====================================================================== 全体
def prove_general(pr, original: bytes, new: bytes, edits: dict, ops: list) -> None:
    a = JtdFile(original)
    try:
        b = JtdFile(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません：{e}")
        return
    if a.compressed != b.compressed:
        pr.unexpected.append("圧縮して保存した文書かどうかが変わりました")
        return
    grow = [op for op in ops if len(op.text) != op.b - op.a]
    removed = {n for n in LAYOUT if a.cfb.has(n)} if grow else set()
    edited_streams = {a.texts[k].stream for (k, _i) in edits} | {a.texts[op.k].stream for op in ops}
    changed = set(edited_streams)
    if grow:
        changed |= {"\x04JSRV_SegmentInformation", "DocumentEditStyles", "DocumentTextPositionTables"}
    # 圧縮した文書：外側の入れ物と、JSCompDocument
    if a.compressed:
        from .lha import jscomp_parts, unpack_jscomp

        oa, ob = CFB(original), CFB(new)
        if set(oa.streams()) != set(ob.streams()):
            pr.unexpected.append("圧縮した文書の外側のストリームの構成が変わりました")
            return
        for n in oa.streams():
            if n in ("JSCompDocument", "\x04JSRV_SegmentInformation"):
                continue
            pr.streams_total += 1
            if oa.read(n) == ob.read(n):
                pr.streams_identical += 1
            else:
                pr.unexpected.append(f"圧縮した文書の外側のストリーム {_d(n)!r} が変わりました")
        try:
            sa, sb = oa.read("JSCompDocument"), ob.read("JSCompDocument")
            pa_, pb_ = jscomp_parts(sa), jscomp_parts(sb)
            if pa_["prefix"] != pb_["prefix"] or pa_["tail"] != pb_["tail"]:
                pr.unexpected.append("JSCompDocument の前置き・後ろが変わりました")
            ha, hb = bytearray(pa_["header"]), bytearray(pb_["header"])
            for h, pt_ in ((ha, pa_), (hb, pb_)):
                h[1] = 0
                h[7:15] = bytes(8)
                h[pt_["crc_at"]:pt_["crc_at"] + 2] = b"\0\0"
            if ha != hb:
                pr.unexpected.append("LHA の見出しが、長さ・CRC・チェックサムのほかに変わりました")
            if sum(pb_["header"][2:2 + pb_["hsize"]]) & 0xFF != pb_["checksum"]:
                pr.unexpected.append("LHA の見出しのチェックサムが合いません")
            inner = unpack_jscomp(sb)  # 解いて、見出しの CRC-16 と長さで確かめる
            if inner != b.cfb.blob:
                pr.unexpected.append("JSCompDocument を解いた中身が、読み直した文書と違います")
            pr.structures.append(f"JSCompDocument（-lh5- で縮め直し：{pa_['packed']:,} → {pb_['packed']:,} バイト。"
                                 f"解き直して CRC-16 {pb_['crc']:04X} と長さ {pb_['original']:,} バイトが合うことを確かめた）")
        except JtdError as e:
            pr.unexpected.append(f"縮め直した文書を解けません：{e}")
            return
        bad = check_seginfo(pr, oa.read("\x04JSRV_SegmentInformation"), ob.read("\x04JSRV_SegmentInformation"), ob,
                            set(), {"JSCompDocument"}, "外側の")
        pr.unexpected += ["外側の" + m for m in bad]
        pr.unexpected += ["外側の" + m for m in check_container(original, new, {"JSCompDocument", "\x04JSRV_SegmentInformation"}, set())]
    xa, xb = a.cfb.blob, b.cfb.blob
    ca, cb = a.cfb, b.cfb
    sa, sb = set(ca.streams()), set(cb.streams())
    if sb != sa - removed:
        pr.unexpected.append(f"ストリームの構成が、取り除いた割り付けの控えのほかに変わりました：{sorted(map(_d, sb ^ (sa - removed)))}")
        return
    for n in sorted(sa - removed):
        pr.streams_total += 1
        x, y = ca.read(n), cb.read(n)
        if n not in changed:
            if x == y:
                pr.streams_identical += 1
            else:
                pr.unexpected.append(f"ストリーム {_d(n)!r} が変わりました（変えてよいストリームではありません）")
        elif x == y:
            pr.streams_identical += 1
    if removed:
        pr.structures.append("行・ページの割り付けの控え（" + "・".join(sorted(removed)) + "）を取り除いた")
    # 字数の同じ上書きをした、本文でない字のストリーム：変わったのは、その字の 2 バイトずつだけ
    k_body = ops[0].k if ops else None
    for name in edited_streams - {a.texts[k_body].stream if k_body is not None else ""}:
        x, y = ca.read(name), cb.read(name)
        want = {}
        for (k, i), ch in edits.items():
            if a.texts[k].stream == name:
                off = a.texts[k].unit_offset(i)
                want[off], want[off + 1] = ord(ch).to_bytes(2, "big")
        if len(x) != len(y) or any(x[i] != y[i] and i not in want for i in range(len(x))) or any(y[i] != v for i, v in want.items()):
            pr.unexpected.append(f"ストリーム {_d(name)!r} が、上書きした字のほかに変わりました")
    if ops:
        pr.unexpected += check_text(pr, a, b, ops, edits)
    if grow:
        pr.unexpected += check_positions(pr, a, b, ops)
        pr.unexpected += check_seginfo(pr, ca.read("\x04JSRV_SegmentInformation"), cb.read("\x04JSRV_SegmentInformation"),
                                       cb, removed, changed)
    pr.unexpected += check_container(xa, xb, {n for n in changed if ca.has(n)}, removed)
    # 字の並び：本文のほかの文字のセグメントは、上書きした字のほか同じ
    if len(a.texts) != len(b.texts):
        pr.lint_errors.append("書き込み後の文書の、文字のセグメントの数が変わりました")
        return
    for x, y in zip(a.texts, b.texts):
        if k_body is not None and x.k == k_body:
            continue
        if len(x.units) != len(y.units) or any(u != v and ((x.k, i) not in edits or ord(edits[(x.k, i)]) != v)
                                               for i, (u, v) in enumerate(zip(x.units, y.units))):
            pr.lint_errors.append(f"{_d(x.stream)}#{x.sid} の字が、上書きした字のほかに変わりました")
    # 読み直し
    pa = [p for _k, _p, bl in a.all_blocks() for p in iter_paras(bl)]
    pb = [p for _k, _p, bl in b.all_blocks() for p in iter_paras(bl)]
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.readback.append(f"段落の数が変わりました（{len(pa)} → {len(pb)}）")
        return
    for n, (p, q) in enumerate(zip(pa, pb), 1):
        exp = expected_text(a, p, ops, edits)
        if exp is not None:
            pr.paragraphs_changed += 1
        else:
            exp = p.text
        if q.text != exp:
            pr.readback.append(f"段落 {n}：読み直すと「{q.text[:40]}」で、思っていた「{exp[:40]}」と違います")
    pr.chars_written = len(edits) + sum(len(op.text) for op in ops)
    pr.chars_delta = sum(len(op.text) - (op.b - op.a) for op in ops)
