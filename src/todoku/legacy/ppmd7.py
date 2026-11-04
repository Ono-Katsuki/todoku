"""PPMd の variant H（Dmitry Shkarin の PPMd var.H。7-Zip の PPMD）を Python で解く。

モデルの形（12 バイトのユニットの割り当て・文脈・状態・SEE・二値文脈）は、公開された LZMA SDK の Ppmd7.c の説明のとおり。
割り当て方そのものがモデルの動きを決める（メモリが尽きた所でモデルをやり直す）ので、ユニットの割り当てもそのとおりに動かす。
範囲符号は 7z の RangeDec7z（LZMA SDK の Ppmd7Dec.c と同じ）。LZMA SDK は public domain。

文脈（12 バイト）：[0:2] 状態の数・[2:4] 頻度の和・[4:8] 状態の表の位置・[8:12] 親の文脈の位置。状態が 1 つなら [2:8] がその状態。
状態（6 バイト）：[0] 字・[1] 頻度・[2:6] 次の文脈の位置。
"""

from __future__ import annotations

MAX_ORDER = 64
UNIT = 12
MAX_FREQ = 124
PERIOD_BITS = 7
INT_BITS = 7
BIN_SCALE = 1 << (INT_BITS + PERIOD_BITS)
NUM_INDEXES = 38
MAX_MEM = 1 << 30  # 持てるモデルの大きさ（これより大きな宣言は、メモリを守るために解かない）
_INIT_BIN_ESC = (0x3CDD, 0x1F3F, 0x59BF, 0x48F3, 0x64A1, 0x5ABC, 0x6632, 0x6051)
_EXP_ESCAPE = (25, 14, 9, 7, 5, 5, 4, 4, 4, 3, 3, 3, 2, 2, 2, 2)


class PpmdError(ValueError):
    pass


def _tables():
    i2u = [0] * NUM_INDEXES
    u2i = [0] * 128
    k = 0
    for i in range(NUM_INDEXES):
        step = 4 if i >= 12 else (i >> 2) + 1
        for _ in range(step):
            u2i[k] = i
            k += 1
        i2u[i] = k
    ns2bs = [0] * 256
    ns2bs[0], ns2bs[1] = 0 << 1, 1 << 1
    for i in range(2, 11):
        ns2bs[i] = 2 << 1
    for i in range(11, 256):
        ns2bs[i] = 3 << 1
    ns2i = [0] * 256
    for i in range(3):
        ns2i[i] = i
    i = 3
    m, k = 3, 1
    while i < 256:
        ns2i[i] = m
        k -= 1
        if k == 0:
            m += 1
            k = m - 2
        i += 1
    hb2 = [0] * 0x40 + [8] * 0xC0
    return i2u, u2i, ns2bs, ns2i, hb2


_I2U, _U2I, _NS2BS, _NS2I, _HB2 = _tables()


class RangeDec7z:
    """7z の範囲符号（LZMA と同じ、桁の繰り上がりのある符号の復号）。先頭の 1 バイトは 0。"""

    def __init__(self, data, pos: int = 0):
        self.d = data
        self.p = pos
        self.range = 0xFFFFFFFF
        if pos + 5 > len(data) or data[pos] != 0:
            raise PpmdError("PPMd の範囲符号の始まりが壊れています")
        self.code = int.from_bytes(data[pos + 1:pos + 5], "big")
        self.p = pos + 5
        if self.code >= 0xFFFFFFFF:
            raise PpmdError("PPMd の範囲符号の始まりが壊れています")

    def threshold(self, total: int) -> int:
        self.range //= total
        return self.code // self.range

    def _norm(self) -> None:
        r = self.range
        if r < 0x1000000:
            d, p = self.d, self.p
            n = len(d)
            self.code = ((self.code << 8) | (d[p] if p < n else 0)) & 0xFFFFFFFF
            p += 1
            r <<= 8
            if r < 0x1000000:
                self.code = ((self.code << 8) | (d[p] if p < n else 0)) & 0xFFFFFFFF
                p += 1
                r <<= 8
            self.range = r
            self.p = p

    def decode(self, start: int, size: int) -> None:
        self.code -= start * self.range
        self.range *= size
        self._norm()

    def bit(self, size0: int, total: int) -> int:
        bound = (self.range // total) * size0
        if self.code < bound:
            self.range = bound
            self._norm()
            return 0
        self.code -= bound
        self.range -= bound
        self._norm()
        return 1


class Ppmd7:
    def __init__(self, mem_size: int, order: int):
        if not 2 <= order <= MAX_ORDER:
            raise PpmdError(f"PPMd の次数（{order}）が範囲の外です")
        if not (1 << 11) <= mem_size <= MAX_MEM:
            raise PpmdError(f"PPMd のモデルの大きさ（{mem_size:,} バイト）は、とどくで持てる大きさ（{MAX_MEM:,} バイト）の外です")
        self.size = mem_size
        self.align = 4 - (mem_size & 3)
        total = self.align + mem_size + UNIT
        if total > 1 << 24:  # 大きなモデルは、触った所だけメモリを使う形で持つ
            import mmap

            self.m = mmap.mmap(-1, total)
        else:
            self.m = bytearray(total)
        self.order = order
        self.free = [0] * NUM_INDEXES
        self.bin = [0] * (128 * 64)
        self.see_sum = [0] * (25 * 16 + 1)
        self.see_shift = [0] * (25 * 16 + 1)
        self.see_count = [0] * (25 * 16 + 1)
        self.see_shift[400] = PERIOD_BITS
        self.see_count[400] = 64
        self.init_esc = 0
        self.hi_bits = 0
        self.rc = None
        self.restart()

    # ------------------------------------------------------------ 割り当て
    def _insert(self, node: int, idx: int) -> None:
        self.m[node:node + 4] = self.free[idx].to_bytes(4, "little")
        self.free[idx] = node

    def _remove(self, idx: int) -> int:
        node = self.free[idx]
        m = self.m
        self.free[idx] = m[node] | m[node + 1] << 8 | m[node + 2] << 16 | m[node + 3] << 24
        return node

    def _split(self, ptr: int, old: int, new: int) -> None:
        nu = _I2U[old] - _I2U[new]
        ptr += _I2U[new] * UNIT
        i = _U2I[nu - 1]
        if _I2U[i] != nu:
            i -= 1
            k = _I2U[i]
            self._insert(ptr + k * UNIT, nu - k - 1)
        self._insert(ptr, i)

    def _glue(self) -> None:
        m = self.m
        u32 = lambda o: m[o] | m[o + 1] << 8 | m[o + 2] << 16 | m[o + 3] << 24  # noqa: E731
        head = self.align + self.size
        n = head
        self.glue_count = 255
        for i in range(NUM_INDEXES):
            nu = _I2U[i]
            nxt = self.free[i]
            self.free[i] = 0
            while nxt:
                node = nxt
                nxt = u32(node)
                m[node + 4:node + 8] = n.to_bytes(4, "little")  # Next
                m[n + 8:n + 12] = node.to_bytes(4, "little")  # Prev
                n = node
                m[node:node + 4] = bytes((0, 0, nu & 255, nu >> 8))  # Stamp 0・NU
        m[head:head + 2] = b"\x01\x00"
        m[head + 4:head + 8] = n.to_bytes(4, "little")
        m[n + 8:n + 12] = head.to_bytes(4, "little")
        if self.lo_unit != self.hi_unit:
            m[self.lo_unit:self.lo_unit + 2] = b"\x01\x00"
        # つなぐ
        while n != head:
            nu = m[n + 2] | m[n + 3] << 8
            while True:
                node2 = n + nu * UNIT
                nu += m[node2 + 2] | m[node2 + 3] << 8
                if (m[node2] | m[node2 + 1]) != 0 or nu >= 0x10000:
                    break
                prev, nxt = u32(node2 + 8), u32(node2 + 4)
                m[prev + 4:prev + 8] = nxt.to_bytes(4, "little")
                m[nxt + 8:nxt + 12] = prev.to_bytes(4, "little")
                m[n + 2] = nu & 255
                m[n + 3] = nu >> 8
            n = u32(n + 4)
        # 空きの表に戻す
        n = u32(head + 4)
        while n != head:
            node = n
            nxt = u32(n + 4)
            nu = m[node + 2] | m[node + 3] << 8
            while nu > 128:
                self._insert(node, NUM_INDEXES - 1)
                nu -= 128
                node += 128 * UNIT
            i = _U2I[nu - 1]
            if _I2U[i] != nu:
                i -= 1
                k = _I2U[i]
                self._insert(node + k * UNIT, nu - k - 1)
            self._insert(node, i)
            n = nxt

    def _alloc_rare(self, idx: int) -> int:
        if self.glue_count == 0:
            self._glue()
            if self.free[idx]:
                return self._remove(idx)
        i = idx
        while True:
            i += 1
            if i == NUM_INDEXES:
                num = _I2U[idx] * UNIT
                self.glue_count -= 1
                if self.units_start - self.text > num:
                    self.units_start -= num
                    return self.units_start
                return 0
            if self.free[i]:
                break
        ret = self._remove(i)
        self._split(ret, i, idx)
        return ret

    def _alloc(self, idx: int) -> int:
        if self.free[idx]:
            return self._remove(idx)
        num = _I2U[idx] * UNIT
        if num <= self.hi_unit - self.lo_unit:
            ret = self.lo_unit
            self.lo_unit += num
            return ret
        return self._alloc_rare(idx)

    def _shrink(self, old: int, old_nu: int, new_nu: int) -> int:
        i0, i1 = _U2I[old_nu - 1], _U2I[new_nu - 1]
        if i0 == i1:
            return old
        if self.free[i1]:
            ptr = self._remove(i1)
            n = new_nu * UNIT
            self.m[ptr:ptr + n] = self.m[old:old + n]
            self._insert(old, i0)
            return ptr
        self._split(old, i0, i1)
        return old

    # ------------------------------------------------------------ モデル
    def restart(self) -> None:
        m = self.m
        self.free = [0] * NUM_INDEXES
        self.text = self.align
        self.hi_unit = self.text + self.size
        self.lo_unit = self.units_start = self.hi_unit - self.size // 8 // UNIT * 7 * UNIT
        self.glue_count = 0
        self.order_fall = self.order
        self.run_length = self.init_rl = -min(self.order, 12) - 1
        self.prev_success = 0
        self.hi_unit -= UNIT
        ctx = self.hi_unit
        self.min_context = self.max_context = ctx
        m[ctx + 8:ctx + 12] = b"\0\0\0\0"
        m[ctx:ctx + 2] = b"\x00\x01"  # 256
        m[ctx + 2:ctx + 4] = (257).to_bytes(2, "little")
        s = self.lo_unit
        self.found = s
        self.lo_unit += 128 * UNIT
        m[ctx + 4:ctx + 8] = s.to_bytes(4, "little")
        for i in range(256):
            m[s:s + 6] = bytes((i, 1, 0, 0, 0, 0))
            s += 6
        b = self.bin
        for i in range(128):
            for k in range(8):
                val = BIN_SCALE - _INIT_BIN_ESC[k] // (i + 2)
                for j in range(0, 64, 8):
                    b[i * 64 + k + j] = val
        for i in range(25):
            for k in range(16):
                self.see_shift[i * 16 + k] = PERIOD_BITS - 4
                self.see_sum[i * 16 + k] = (5 * i + 10) << (PERIOD_BITS - 4)
                self.see_count[i * 16 + k] = 4

    def _create_successors(self, skip: bool) -> int:
        m = self.m
        c = self.min_context
        fs = self.found
        up_branch = m[fs + 2] | m[fs + 3] << 8 | m[fs + 4] << 16 | m[fs + 5] << 24
        fsym = m[fs]
        ps = [] if skip else [fs]
        while True:
            suf = m[c + 8] | m[c + 9] << 8 | m[c + 10] << 16 | m[c + 11] << 24
            if not suf:
                break
            c = suf
            if m[c] | m[c + 1] << 8 != 1:
                s = m[c + 4] | m[c + 5] << 8 | m[c + 6] << 16 | m[c + 7] << 24
                while m[s] != fsym:
                    s += 6
            else:
                s = c + 2
            succ = m[s + 2] | m[s + 3] << 8 | m[s + 4] << 16 | m[s + 5] << 24
            if succ != up_branch:
                c = succ
                if not ps:
                    return c
                break
            ps.append(s)
        up_sym = m[up_branch]
        up_succ = up_branch + 1
        if m[c] | m[c + 1] << 8 == 1:
            up_freq = m[c + 3]
        else:
            s = m[c + 4] | m[c + 5] << 8 | m[c + 6] << 16 | m[c + 7] << 24
            while m[s] != up_sym:
                s += 6
            cf = m[s + 1] - 1
            s0 = (m[c + 2] | m[c + 3] << 8) - (m[c] | m[c + 1] << 8) - cf
            up_freq = (1 + (int(5 * cf > s0) if 2 * cf <= s0 else (2 * cf + 3 * s0 - 1) // (2 * s0))) & 255
        succ_b = up_succ.to_bytes(4, "little")
        one = bytes((1, 0, up_sym, up_freq)) + succ_b + c.to_bytes(4, "little")
        for p in reversed(ps):
            if self.hi_unit != self.lo_unit:
                self.hi_unit -= UNIT
                c1 = self.hi_unit
            elif self.free[0]:
                c1 = self._remove(0)
            else:
                c1 = self._alloc_rare(0)
                if not c1:
                    return 0
            m[c1:c1 + 12] = one[:8] + c.to_bytes(4, "little")
            m[p + 2:p + 6] = c1.to_bytes(4, "little")
            c = c1
        return c

    def _update_model(self) -> None:
        m = self.m
        fs = self.found
        fsym, ffreq = m[fs], m[fs + 1]
        f_succ = m[fs + 2] | m[fs + 3] << 8 | m[fs + 4] << 16 | m[fs + 5] << 24
        mc = self.min_context
        if ffreq < MAX_FREQ // 4:
            c = m[mc + 8] | m[mc + 9] << 8 | m[mc + 10] << 16 | m[mc + 11] << 24
            if c:
                if m[c] | m[c + 1] << 8 == 1:
                    if m[c + 3] < 32:
                        m[c + 3] += 1
                else:
                    s = m[c + 4] | m[c + 5] << 8 | m[c + 6] << 16 | m[c + 7] << 24
                    if m[s] != fsym:
                        s += 6
                        while m[s] != fsym:
                            s += 6
                        if m[s + 1] >= m[s - 5]:
                            a, b = bytes(m[s:s + 6]), bytes(m[s - 6:s])
                            m[s - 6:s] = a
                            m[s:s + 6] = b
                            s -= 6
                    if m[s + 1] < MAX_FREQ - 9:
                        m[s + 1] += 2
                        sf = (m[c + 2] | m[c + 3] << 8) + 2
                        m[c + 2] = sf & 255
                        m[c + 3] = (sf >> 8) & 255
        if self.order_fall == 0:
            c = self._create_successors(True)
            self.min_context = self.max_context = c
            if not c:
                self.restart()
                return
            m[fs + 2:fs + 6] = c.to_bytes(4, "little")
            return
        m[self.text] = fsym
        self.text += 1
        successor = self.text
        if self.text >= self.units_start:
            self.restart()
            return
        if f_succ:
            if f_succ <= successor:
                cs = self._create_successors(False)
                if not cs:
                    self.restart()
                    return
                f_succ = cs
            self.order_fall -= 1
            if self.order_fall == 0:
                successor = f_succ
                if self.max_context != self.min_context:
                    self.text -= 1
        else:
            m[fs + 2:fs + 6] = successor.to_bytes(4, "little")
            f_succ = self.min_context
        mc = self.min_context
        ns = m[mc] | m[mc + 1] << 8
        ffreq = m[fs + 1]
        s0 = (m[mc + 2] | m[mc + 3] << 8) - ns - (ffreq - 1)
        c = self.max_context
        succ_b = successor.to_bytes(4, "little")
        while c != mc:
            ns1 = m[c] | m[c + 1] << 8
            if ns1 != 1:
                if ns1 & 1 == 0:
                    old_nu = ns1 >> 1
                    i = _U2I[old_nu - 1]
                    if i != _U2I[old_nu]:
                        ptr = self._alloc(i + 1)
                        if not ptr:
                            self.restart()
                            return
                        old = m[c + 4] | m[c + 5] << 8 | m[c + 6] << 16 | m[c + 7] << 24
                        n = old_nu * UNIT
                        m[ptr:ptr + n] = m[old:old + n]
                        self._insert(old, i)
                        m[c + 4:c + 8] = ptr.to_bytes(4, "little")
                sf = m[c + 2] | m[c + 3] << 8
                sf = sf + int(2 * ns1 < ns) + 2 * int(4 * ns1 <= ns and sf <= 8 * ns1)
            else:
                s = self._alloc(0)
                if not s:
                    self.restart()
                    return
                m[s:s + 6] = m[c + 2:c + 8]
                m[c + 4:c + 8] = s.to_bytes(4, "little")
                fr = m[s + 1]
                fr = fr << 1 if fr < MAX_FREQ // 4 - 1 else MAX_FREQ - 4
                m[s + 1] = fr
                sf = fr + self.init_esc + int(ns > 3)
            cf = 2 * ffreq * (sf + 6)
            total = s0 + sf
            if cf < 6 * total:
                cf = 1 + int(cf > total) + int(cf >= 4 * total)
                sf += 3
            else:
                cf = 4 + int(cf >= 9 * total) + int(cf >= 12 * total) + int(cf >= 15 * total)
                sf += cf
            m[c + 2] = sf & 255
            m[c + 3] = (sf >> 8) & 255
            s = (m[c + 4] | m[c + 5] << 8 | m[c + 6] << 16 | m[c + 7] << 24) + ns1 * 6
            m[s] = fsym
            m[s + 1] = cf
            m[s + 2:s + 6] = succ_b
            ns1 += 1
            m[c] = ns1 & 255
            m[c + 1] = ns1 >> 8
            c = m[c + 8] | m[c + 9] << 8 | m[c + 10] << 16 | m[c + 11] << 24
        self.max_context = self.min_context = f_succ

    def _rescale(self) -> None:
        m = self.m
        mc = self.min_context
        stats = m[mc + 4] | m[mc + 5] << 8 | m[mc + 6] << 16 | m[mc + 7] << 24
        s = self.found
        if s != stats:
            tmp = bytes(m[s:s + 6])
            m[stats + 6:s + 6] = m[stats:s]
            m[stats:stats + 6] = tmp
            s = stats
        sum_f = m[mc + 2] | m[mc + 3] << 8
        esc = sum_f - m[s + 1]
        adder = 1 if self.order_fall else 0
        fr = (m[s + 1] + 4 + adder) >> 1
        m[s + 1] = fr
        sum_freq = fr
        ns = m[mc] | m[mc + 1] << 8
        i = ns - 1
        while True:
            s += 6
            esc -= m[s + 1]
            fr = (m[s + 1] + adder) >> 1
            m[s + 1] = fr
            sum_freq += fr
            if fr > m[s - 5]:
                tmp = bytes(m[s:s + 6])
                s1 = s
                while True:
                    m[s1:s1 + 6] = m[s1 - 6:s1]
                    s1 -= 6
                    if s1 == stats or fr <= m[s1 - 5]:
                        break
                m[s1:s1 + 6] = tmp
            i -= 1
            if not i:
                break
        if m[s + 1] == 0:
            num_stats = ns
            while True:
                i += 1
                s -= 6
                if m[s + 1] != 0:
                    break
            esc += i
            ns -= i
            m[mc] = ns & 255
            m[mc + 1] = ns >> 8
            if ns == 1:
                sym, fr = m[stats], m[stats + 1]
                succ = bytes(m[stats + 2:stats + 6])
                while True:
                    fr -= fr >> 1
                    esc >>= 1
                    if esc <= 1:
                        break
                self._insert(stats, _U2I[((num_stats + 1) >> 1) - 1])
                self.found = mc + 2
                m[mc + 2] = sym
                m[mc + 3] = fr
                m[mc + 4:mc + 8] = succ
                return
            n0 = (num_stats + 1) >> 1
            n1 = (ns + 1) >> 1
            if n0 != n1:
                stats = self._shrink(stats, n0, n1)
                m[mc + 4:mc + 8] = stats.to_bytes(4, "little")
        sf = (sum_freq + esc - (esc >> 1)) & 0xFFFF
        m[mc + 2] = sf & 255
        m[mc + 3] = sf >> 8
        self.found = m[mc + 4] | m[mc + 5] << 8 | m[mc + 6] << 16 | m[mc + 7] << 24

    def _next_context(self) -> None:
        m = self.m
        fs = self.found
        c = m[fs + 2] | m[fs + 3] << 8 | m[fs + 4] << 16 | m[fs + 5] << 24
        if self.order_fall == 0 and c > self.text:
            self.min_context = self.max_context = c
        else:
            self._update_model()

    def _add_sum(self, mc: int, n: int) -> None:
        m = self.m
        v = (m[mc + 2] | m[mc + 3] << 8) + n
        m[mc + 2] = v & 255
        m[mc + 3] = (v >> 8) & 255

    def _update1(self) -> None:
        m = self.m
        s = self.found
        fr = m[s + 1] + 4
        m[s + 1] = fr & 255
        self._add_sum(self.min_context, 4)
        if fr > m[s - 5]:
            a, b = bytes(m[s:s + 6]), bytes(m[s - 6:s])
            m[s - 6:s] = a
            m[s:s + 6] = b
            s -= 6
            self.found = s
            if fr > MAX_FREQ:
                self._rescale()
        self._next_context()

    def _update1_0(self) -> None:
        m = self.m
        s = self.found
        mc = self.min_context
        fr = m[s + 1]
        ps = int(2 * fr > (m[mc + 2] | m[mc + 3] << 8))
        self.prev_success = ps
        self.run_length += ps
        self._add_sum(mc, 4)
        fr += 4
        m[s + 1] = fr & 255
        if fr > MAX_FREQ:
            self._rescale()
        self._next_context()

    def _update2(self) -> None:
        m = self.m
        s = self.found
        self._add_sum(self.min_context, 4)
        fr = m[s + 1] + 4
        m[s + 1] = fr & 255
        if fr > MAX_FREQ:
            self._rescale()
        self.run_length = self.init_rl
        self._update_model()

    def _make_esc(self, num_masked: int) -> tuple[int, int]:
        m = self.m
        mc = self.min_context
        ns = m[mc] | m[mc + 1] << 8
        if ns == 256:
            return 400, 1
        non = ns - num_masked
        suf = m[mc + 8] | m[mc + 9] << 8 | m[mc + 10] << 16 | m[mc + 11] << 24
        see = _NS2I[non - 1] * 16 + int(non < (m[suf] | m[suf + 1] << 8) - ns) + 2 * int((m[mc + 2] | m[mc + 3] << 8) < 11 * ns) \
            + 4 * int(num_masked > non) + self.hi_bits
        ssum = self.see_sum[see]
        r = ssum >> self.see_shift[see]
        self.see_sum[see] = (ssum - r) & 0xFFFF
        return see, r + (r == 0)

    # ------------------------------------------------------------ 1 字
    def decode_symbol(self) -> int:
        """字（0〜255）。-1 はデータの終わりの印、-2 はデータの誤り。"""
        m = self.m
        rc = self.rc
        mc = self.min_context
        ns = m[mc] | m[mc + 1] << 8
        if ns != 1:
            s = m[mc + 4] | m[mc + 5] << 8 | m[mc + 6] << 16 | m[mc + 7] << 24
            sf = m[mc + 2] | m[mc + 3] << 8
            count = rc.threshold(sf)
            hi = m[s + 1]
            if count < hi:
                rc.decode(0, hi)
                self.found = s
                sym = m[s]
                self._update1_0()
                return sym
            self.prev_success = 0
            for _ in range(ns - 1):
                s += 6
                f = m[s + 1]
                hi += f
                if hi > count:
                    rc.decode(hi - f, f)
                    self.found = s
                    sym = m[s]
                    self._update1()
                    return sym
            if count >= sf:
                return -2
            self.hi_bits = _HB2[m[self.found]]
            rc.decode(hi, sf - hi)
            masked = bytearray(256)
            for _ in range(ns):
                masked[m[s]] = 1
                s -= 6
        else:
            suf = m[mc + 8] | m[mc + 9] << 8 | m[mc + 10] << 16 | m[mc + 11] << 24
            sym1 = m[mc + 2]
            self.hi_bits = _HB2[m[self.found]]
            idx = (m[mc + 3] - 1) * 64 + self.prev_success + _NS2BS[(m[suf] | m[suf + 1] << 8) - 1] + self.hi_bits \
                + 2 * _HB2[sym1] + (0x20 if self.run_length < 0 else 0)
            b = self.bin
            prob = b[idx]
            if rc.bit(prob, BIN_SCALE) == 0:
                b[idx] = prob + (1 << INT_BITS) - ((prob + (1 << (PERIOD_BITS - 2))) >> PERIOD_BITS)
                self.found = mc + 2
                fr = m[mc + 3]
                if fr < 128:
                    m[mc + 3] = fr + 1
                self.prev_success = 1
                self.run_length += 1
                self._next_context()
                return sym1
            prob = prob - ((prob + (1 << (PERIOD_BITS - 2))) >> PERIOD_BITS)
            b[idx] = prob
            self.init_esc = _EXP_ESCAPE[prob >> 10]
            masked = bytearray(256)
            masked[sym1] = 1
            self.prev_success = 0
        while True:
            num_masked = m[mc] | m[mc + 1] << 8
            while True:
                self.order_fall += 1
                suf = m[mc + 8] | m[mc + 9] << 8 | m[mc + 10] << 16 | m[mc + 11] << 24
                if not suf:
                    self.min_context = mc
                    return -1
                mc = suf
                ns = m[mc] | m[mc + 1] << 8
                if ns != num_masked:
                    break
            self.min_context = mc
            hi = 0
            s = m[mc + 4] | m[mc + 5] << 8 | m[mc + 6] << 16 | m[mc + 7] << 24
            ps = []
            num = ns - num_masked
            while len(ps) < num:
                if not masked[m[s]]:
                    hi += m[s + 1]
                    ps.append(s)
                s += 6
            see, freq = self._make_esc(num_masked)
            freq += hi
            count = rc.threshold(freq)
            if count < hi:
                hi = 0
                for s in ps:
                    f = m[s + 1]
                    hi += f
                    if hi > count:
                        break
                rc.decode(hi - f, f)
                sh = self.see_shift[see]
                if sh < PERIOD_BITS:
                    self.see_count[see] -= 1
                    if self.see_count[see] == 0:
                        self.see_sum[see] = (self.see_sum[see] << 1) & 0xFFFF
                        self.see_count[see] = (3 << sh) & 255
                        self.see_shift[see] = sh + 1
                self.found = s
                sym = m[s]
                self._update2()
                return sym
            if count >= freq:
                return -2
            rc.decode(hi, freq - hi)
            self.see_sum[see] = (self.see_sum[see] + freq) & 0xFFFF
            for s in ps:
                masked[m[s]] = 1


def decode_7z(data, props: bytes, size: int) -> bytes:
    """7z の PPMD（プロパティ：次数 1 バイト・モデルの大きさ 4 バイト）を size バイトまで解く。"""
    if len(props) < 5:
        raise PpmdError("PPMd のプロパティが足りません")
    model = Ppmd7(int.from_bytes(props[1:5], "little"), props[0])
    rc = RangeDec7z(data)
    model.rc = rc
    out = bytearray()
    dec = model.decode_symbol
    try:
        for _ in range(size):
            c = dec()
            if c < 0:
                raise PpmdError("PPMd のデータが途中で終わっているか、壊れています")
            out.append(c)
    except (IndexError, ZeroDivisionError, OverflowError, ValueError) as e:
        if isinstance(e, PpmdError):
            raise
        raise PpmdError(f"PPMd のデータが壊れています（{type(e).__name__}）") from None
    return bytes(out)
