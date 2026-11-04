"""試験で符号の流れを手で作る小さな部品：LZMA の範囲符号（1 ビットの確率つき）。"""

from __future__ import annotations


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
