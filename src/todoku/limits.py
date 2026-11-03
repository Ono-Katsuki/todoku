"""上限（予算）：壊れた・悪意のある書類から、PC を守る。

役所の書類は、どこから来たか分からないファイルとして開く。小さなファイルが展開すると何 GB にもなる（圧縮の爆弾）、
参照が輪になっていて読み終わらない、といった書類で、メモリや時間を使い切らないよう、次の上限を置く。
上限は、ふつうの申請書類（数 MB・数百ページまで）が決して当たらない大きさにしてあり、環境変数で変えられる。
上限に当たったときは LimitExceeded（ValueError の一種）を投げ、エージェントには「大きすぎるか、壊れている」と伝える。

1 回の道具の呼び出し（窓を開く・書き込む・検査する）ごとに Budget を 1 つ作り（call_budget）、その呼び出しの中の
展開（ZIP の部品・PDF のストリーム・一太郎の圧縮）と、PDF の描画の命令を、同じ Budget で数える。
呼び出しの外（ライブラリとして使うとき）は、1 回の展開の上限だけを守る（合計は数えない）。

環境変数（大きさは MB）：
- TODOKU_MAX_FILE_MB（200）            開くファイルの大きさ
- TODOKU_MAX_INFLATE_PART_MB（256）    1 つの部品・ストリームを展開した大きさ（宣言された大きさではなく、実際に出てきたバイト）
- TODOKU_MAX_INFLATE_TOTAL_MB（2048）  1 回の呼び出しで展開するバイトの合計
- TODOKU_MAX_PDF_OPS（200000000）     1 回の呼び出しで読む、PDF の描画の字句（数・名前・命令）の数の合計
- TODOKU_MAX_PAGE_OPS（5000000）       1 ページ（とその部品）で読む、PDF の描画の字句の数
- TODOKU_MAX_PIXELS（100000000）       1 つの画像を画素に開く大きさ（幅 × 高さ）
- TODOKU_TIME_LIMIT（120）             1 回の呼び出しの秒数（MCP サーバー：別のプロセスで動かし、超えたら止める）
- TODOKU_MEMORY_LIMIT_MB（2048）       その別のプロセスのメモリ
- TODOKU_WINDOW_CACHE_MB（256）        窓の控え（同じ書類を開き直さない）の大きさの合計
"""

from __future__ import annotations

import contextlib
import contextvars
import os
import zlib

MB = 1024 * 1024


def _env(name: str, default: float, scale: int = 1) -> int:
    try:
        v = float(os.environ.get(name, "") or default)
    except ValueError:
        v = default
    return int(v * scale) if v > 0 else int(default * scale)


MAX_FILE = _env("TODOKU_MAX_FILE_MB", 200, MB)
MAX_INFLATE_PART = _env("TODOKU_MAX_INFLATE_PART_MB", 256, MB)
MAX_INFLATE_TOTAL = _env("TODOKU_MAX_INFLATE_TOTAL_MB", 2048, MB)
MAX_PDF_OPS = _env("TODOKU_MAX_PDF_OPS", 200_000_000)
MAX_PAGE_OPS = _env("TODOKU_MAX_PAGE_OPS", 5_000_000)
MAX_PIXELS = _env("TODOKU_MAX_PIXELS", 100_000_000)
TIME_LIMIT = _env("TODOKU_TIME_LIMIT", 120)
MEMORY_LIMIT = _env("TODOKU_MEMORY_LIMIT_MB", 2048, MB)
WINDOW_CACHE = _env("TODOKU_WINDOW_CACHE_MB", 256, MB)
# マクロ（VBA）：1 つの文字列の長さ・1 つの配列の要素の数・1 回の実行で作る大きな値（1 MB 以上の文字列）の合計
MAX_VBA_STRING = _env("TODOKU_MAX_VBA_STRING", 16 * MB)
MAX_VBA_ARRAY = _env("TODOKU_MAX_VBA_ARRAY", 10_000_000)
MAX_VBA_BYTES = _env("TODOKU_MAX_VBA_BYTES_MB", 1024, MB)
MAX_VBA_RANGE = _env("TODOKU_MAX_VBA_RANGE", 2_000_000)  # マクロが 1 度に書く範囲のマスの数（1 列まるごと 1,048,576 は書ける）

STOPPED = "この書類は大きすぎるか、壊れているため止めました"


class LimitExceeded(ValueError):
    """上限に当たった（書類が大きすぎるか、壊れている）。"""

    def __init__(self, what: str, total: bool = False):
        super().__init__(f"{STOPPED}：{what}")
        self.what = what
        self.total = total  # 1 回の呼び出しの合計に当たった（1 つの部品の上限ではなく）


def _mb(n: int) -> str:
    return f"{n / MB:,.0f} MB" if n >= MB else f"{n:,} バイト"


class Budget:
    """1 回の呼び出しの予算。total=None なら合計は数えない。"""

    def __init__(self, part: int | None = None, total: int | None = None, ops: int | None = None):
        self.part = MAX_INFLATE_PART if part is None else part
        self.total = total
        self.ops_total = ops if ops is not None else float("inf")
        self.inflated = 0
        self.ops = 0

    def room(self, part: int | None = None) -> int:
        """これから 1 つ展開してよい大きさ。"""
        cap = self.part if part is None else min(part, self.part)
        if self.total is not None:
            cap = min(cap, max(0, self.total - self.inflated))
        return cap

    def charge(self, n: int, what: str = "") -> None:
        self.inflated += n
        if self.total is not None and self.inflated > self.total:
            raise LimitExceeded(f"展開したデータの合計が {_mb(self.total)} を超えます{what and '（' + what + '）'}", True)

    def exceeded(self, cap: int, part: int | None, what: str = "") -> LimitExceeded:
        head = f"{what}：" if what else ""
        if self.total is not None and cap < min(self.part, part if part is not None else self.part):
            return LimitExceeded(f"{head}展開したデータの合計が {_mb(self.total)} を超えます", True)
        return LimitExceeded(f"{head}展開すると {_mb(cap)} を超えます")

    def ops_exceeded(self) -> LimitExceeded:
        return LimitExceeded(f"PDF の描画の字句（数・名前・命令）が、1 回の呼び出しで {self.ops_total:,} 個を超えます"
                             "（部品が部品を何度も描くなど）", True)


_CURRENT: contextvars.ContextVar[Budget | None] = contextvars.ContextVar("todoku_budget", default=None)
_LOOSE = Budget()  # 呼び出しの外：1 回の展開の上限だけ


def budget() -> Budget:
    return _CURRENT.get() or _LOOSE


@contextlib.contextmanager
def call_budget(b: Budget | None = None):
    """1 回の呼び出しの予算を置く（すでに置いてあれば、それを使い続ける：呼び出しの中で api を呼び直しても 1 つ）。"""
    if _CURRENT.get() is not None and b is None:
        yield _CURRENT.get()
        return
    token = _CURRENT.set(b or Budget(total=MAX_INFLATE_TOTAL, ops=MAX_PDF_OPS))
    try:
        yield _CURRENT.get()
    finally:
        _CURRENT.reset(token)


def inflate(data: bytes, wbits: int = 15, what: str = "", part: int | None = None, tolerant: bool = False) -> bytes:
    """zlib・deflate を、出てくるバイトを数えながら展開する（宣言された大きさは信じない）。
    part：この展開だけの上限（予算の 1 つの展開の上限より小さくするとき）。上限を超えたら LimitExceeded。
    tolerant：壊れたデータを、読めるところまで返す（PDF）。そうでなければ zlib.error（途中で終わるデータも）。"""
    b = budget()
    cap = b.room(part)
    d = zlib.decompressobj(wbits)
    out = bytearray()
    src = data
    try:
        while True:
            piece = d.decompress(src, min(1 << 22, cap - len(out) + 1))
            out += piece
            if len(out) > cap:
                raise b.exceeded(cap, part, what)
            src = d.unconsumed_tail
            if d.eof or (not src and not piece):
                break
        if not d.eof and not tolerant:
            raise zlib.error("圧縮したデータが途中で終わっています")
    except zlib.error:
        if not tolerant:
            raise
    b.charge(len(out), what)
    return bytes(out)


def check_output(n: int, what: str = "", part: int | None = None) -> None:
    """展開して n バイトになるか（展開の前に、出てくる大きさが分かるとき）。上限を超えるなら LimitExceeded。"""
    b = budget()
    cap = b.room(part)
    if n > cap:
        raise b.exceeded(cap, part, what)
