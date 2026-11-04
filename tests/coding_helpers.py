"""試験で符号の流れを手で作る小さな部品：上の桁から詰めるビット列、ハフマン符号、LZMA の範囲符号（1 ビットの確率つき）。"""

from __future__ import annotations

import heapq


class BitWriter:
    """上の桁から順に詰める。"""

    def __init__(self):
        self.out = bytearray()
        self.acc = 0
        self.n = 0

    def add(self, value: int, width: int) -> None:
        for i in range(width - 1, -1, -1):
            self.acc = (self.acc << 1) | ((value >> i) & 1)
            self.n += 1
            if self.n == 8:
                self.out.append(self.acc)
                self.acc = self.n = 0

    def done(self) -> bytes:
        if self.n:
            self.add(0, 8 - self.n)
        return bytes(self.out)


def huffman_lengths(freq: list[int], maxlen: int = 15) -> list[int]:
    """符号の長さ。maxlen を超えるときは、頻度をならして作り直す。使う字が 1 つなら長さ 1。"""
    freq = list(freq)
    while True:
        heap: list = [(f, s, (s,)) for s, f in enumerate(freq) if f]
        out = [0] * len(freq)
        if len(heap) == 1:
            out[heap[0][1]] = 1
            return out
        heapq.heapify(heap)
        order = len(freq)
        while len(heap) > 1:
            fa, _a, sa = heapq.heappop(heap)
            fb, _b, sb = heapq.heappop(heap)
            for s in sa + sb:
                out[s] += 1
            heapq.heappush(heap, (fa + fb, order, sa + sb))
            order += 1
        if max(out, default=0) <= maxlen:
            return out
        freq = [(f + 1) // 2 if f else 0 for f in freq]


def canonical_codes(lengths: list[int]) -> list[tuple[int, int]]:
    """字ごとの (符号, 長さ)。短い符号から、同じ長さは字の順。"""
    table = [(0, 0)] * len(lengths)
    code = 0
    for width in range(1, max(lengths, default=0) + 1):
        for sym, ln in enumerate(lengths):
            if ln == width:
                table[sym] = (code, width)
                code += 1
        code <<= 1
    return table


class LzmaRangeEncoder:
    """LZMA の範囲符号（11 ビットの確率、移動 5。LZMA SDK の説明のとおり）。"""

    def __init__(self):
        self.low = 0
        self.range = 0xFFFFFFFF
        self.cache = 0
        self.pending = 1
        self.out = bytearray()

    def _shift_low(self) -> None:
        if self.low < 0xFF000000 or self.low > 0xFFFFFFFF:
            carry = self.low >> 32
            byte = self.cache
            for _ in range(self.pending):
                self.out.append((byte + carry) & 0xFF)
                byte = 0xFF
            self.pending = 0
            self.cache = (self.low >> 24) & 0xFF
        self.pending += 1
        self.low = (self.low & 0x00FFFFFF) << 8

    def encode_bit(self, probs: list[int], index: int, bit: int) -> None:
        p = probs[index]
        bound = (self.range >> 11) * p
        if bit:
            self.low += bound
            self.range -= bound
            probs[index] = p - (p >> 5)
        else:
            self.range = bound
            probs[index] = p + ((2048 - p) >> 5)
        while self.range < 1 << 24:
            self.range = (self.range << 8) & 0xFFFFFFFF
            self._shift_low()

    def finish(self) -> bytes:
        for _ in range(5):
            self._shift_low()
        return bytes(self.out)
