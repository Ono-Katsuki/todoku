"""大きな書類を分けて読む：共通の部品（範囲の指定・続きの印 cursor・1 回分の結果・バイトを写しながら書き換える）。

決まり：読む量を減らしてよいが、読めない所を作らない。
- 窓の全体が limits.CONTEXT_BUDGET（字）に入らない書類は、open_document では概要だけを返し、エージェントが要る所を範囲で頼む
- どの結果も、どこまで読んだか（covered）・残りがいくらか（remaining）・続きの頼み方（next。そのまま read_window に渡せる）を示す
- 続きをたどり終えると、全部を読んだことになる（tests/test_lazy.py で、形式ごとに確かめる）
"""

from __future__ import annotations

import hashlib
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from .. import limits

_RANGE = re.compile(r"\s*(\d+)?\s*(?:[-‐−–—~〜～:]\s*(\d+)?)?\s*")


class RangeError(ValueError):
    pass


def parse_range(spec, what: str = "範囲") -> tuple[int | None, int | None]:
    """「5」「5-10」「5〜」「〜10」・5・[5, 10] → (始め, 終わり)。1 から数える。指定がなければ (None, None)。"""
    if spec is None or spec == "":
        return None, None
    if isinstance(spec, bool):
        raise RangeError(f"{what} の指定「{spec}」は、1 から数えた番号か「5-10」の形で書いてください")
    if isinstance(spec, int):
        return spec, spec
    if isinstance(spec, (list, tuple)) and len(spec) == 2:
        return int(spec[0]), int(spec[1])
    m = _RANGE.fullmatch(str(spec))
    if not m or (m[1] is None and m[2] is None):
        raise RangeError(f"{what} の指定「{spec}」は、1 から数えた番号か「5-10」の形で書いてください")
    if m[0].strip() == (m[1] or ""):  # 番号 1 つ
        return int(m[1]), int(m[1])
    a = int(m[1]) if m[1] else None
    b = int(m[2]) if m[2] else None
    if a is not None and b is not None and b < a:
        raise RangeError(f"{what} の指定「{spec}」は、始めが終わりより後ろです")
    if a is not None and a < 1:
        raise RangeError(f"{what} は 1 から数えます（「{spec}」）")
    return a, b


@dataclass
class Page:
    """1 回分の窓。"""

    text: str  # 見せる Markdown（見出しの注意も含む）
    covered: str  # ここで読んだ範囲（言葉で）
    next: dict | None = None  # 続きの頼み方（read_window の引数）。None なら、頼んだ範囲を最後まで読んだ
    remaining: str = ""  # 残り（言葉で）
    body: list[str] = field(default_factory=list)  # 中身の行（見出し・注意を除く。全部をつなぐと全体の読み取りになる）

    def footer(self) -> str:
        if self.next is None:
            return f"（ここまでで {self.covered} を読みました。頼んだ範囲はこれで終わりです）"
        args = "・".join(f"{k}: {v!r}" if not isinstance(v, str) else f"{k}: \"{v}\"" for k, v in self.next.items())
        return f"（ここまでで {self.covered} を読みました。残り {self.remaining}。続きは read_window に {args} を渡してください）"

    def to_dict(self) -> dict:
        return {"covered": self.covered, "remaining": self.remaining, "next": self.next, "done": self.next is None}


def clamp_chars(max_chars) -> int:
    """1 回の大きさ（字）。指定がなければ limits.PAGE_CHARS。小さすぎる指定でも進めるように、下限を置く。"""
    try:
        n = int(max_chars) if max_chars else limits.PAGE_CHARS
    except (TypeError, ValueError):
        n = limits.PAGE_CHARS
    return max(200, n)


CUT_MARK = "〔この行はここで分けました〕"


def cut_note(a: int, b: int, total: int) -> str:
    return f"〔長い行なので分けて出しています：{a + 1:,}〜{b:,} 字目／全 {total:,} 字〕"


def comma(n: int) -> str:
    return f"{n:,}"


class ByteRanges:
    """書き換えなかった範囲：元のファイルと書いたファイルの位置と、元のファイルのハッシュ。"""

    def __init__(self):
        self.items: list[tuple[int, int, int, str]] = []  # (元の始め, 元の終わり, 書いた先の始め, sha256)
        self.moves: list[tuple[int, int]] = []  # (元の位置, 書いた先の位置)：編集の直後

    def to_new(self, pos: int) -> int:
        """書き換えていない元の位置 pos → 書いたファイルの位置。"""
        import bisect

        i = bisect.bisect_right(self.items, pos, key=lambda t: t[0]) - 1
        a, _b, d, _h = self.items[i]
        return d + pos - a

    @property
    def total(self) -> int:
        return sum(b - a for a, b, _d, _h in self.items)


def _fsync_close(f) -> None:
    f.flush()
    os.fsync(f.fileno())
    f.close()


def stream_patch(src, edits: list[tuple[int, int, bytes]], dst: Path) -> tuple[ByteRanges, str, str]:
    """元のファイル（Source）を dst に写しながら、edits（元の始め, 元の終わり, 置き換えるバイト）だけを置き換える。
    メモリは一定（編集の数に比例）。(書き換えなかった範囲, 元の sha256, 書いたファイルの sha256)。"""
    edits = sorted(edits, key=lambda e: (e[0], e[1]))
    ranges = ByteRanges()
    h_old = hashlib.sha256()
    h_new = hashlib.sha256()
    at = 0
    total = src.size
    with open(dst, "wb") as f:
        out = 0

        def copy(a: int, b: int) -> None:
            nonlocal out
            h = hashlib.sha256()
            d0 = out
            for c in src.chunks(a, b):
                h.update(c)
                h_old.update(c)
                h_new.update(c)
                f.write(c)
                out += len(c)
                limits.progress("書き出しています", a, total)
            ranges.items.append((a, b, d0, h.hexdigest()))

        for x, y, new in edits:
            if x < at:
                raise ValueError("同じ所に2つの書き込みが重なりました")
            copy(at, x)
            for c in src.chunks(x, y):
                h_old.update(c)
            h_new.update(new)
            f.write(new)
            out += len(new)
            at = y
        copy(at, total)
        f.flush()
        os.fsync(f.fileno())
    return ranges, h_old.hexdigest(), h_new.hexdigest()


def verify_ranges(dst: Path, ranges: ByteRanges) -> list[str]:
    """書いたファイルを読み直して、書き換えなかった範囲のハッシュが元のファイルと同じことを確かめる。"""
    bad = []
    fd = os.open(dst, os.O_RDONLY)
    try:
        for a, b, d, want in ranges.items:
            h = hashlib.sha256()
            pos, end = d, d + (b - a)
            while pos < end:
                c = os.pread(fd, min(1 << 20, end - pos), pos)
                if not c:
                    break
                h.update(c)
                pos += len(c)
            if h.hexdigest() != want:
                bad.append(f"書き換えていない {a:,}〜{b:,} バイト目のハッシュが、元のファイルと違います")
    finally:
        os.close(fd)
    return bad


def temp_beside(dst: Path) -> Path:
    import tempfile

    dst.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=f".{dst.stem}.", suffix=f".todoku-tmp{dst.suffix}", dir=dst.parent)
    os.close(fd)
    return Path(name)


class LazyDoc:
    """分けて読める書類（形式ごとに作る）。large：窓の全体が予算に入らない（見積もり est_chars は字の数）。"""

    kind = ""
    est_chars = 0
    streams = False  # 丸ごとメモリに読まずに扱える

    @property
    def large(self) -> bool:
        return self.est_chars > limits.CONTEXT_BUDGET

    def overview(self) -> dict:
        raise NotImplementedError

    def overview_text(self) -> str:
        raise NotImplementedError

    def read(self, rng: dict, max_chars: int) -> Page:
        raise NotImplementedError

    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        raise NotImplementedError

    def fill(self, values: dict, out_path: str, force: bool, overwrite: bool, said: dict | None) -> dict:
        raise NotImplementedError

    def grade(self) -> dict | None:
        return None

    # 検査（audit_document）：区切りごとに進む。state は JSON にできる辞書（途中から続けるための控え）
    def audit_begin(self) -> dict:
        return {"step": 0, "checks": [], "done": False}

    def audit_step(self, state: dict) -> dict:
        raise NotImplementedError
