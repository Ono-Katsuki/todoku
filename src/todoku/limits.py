"""上限（予算）：壊れた・悪意のある書類から、PC を守る。

役所の書類は、どこから来たか分からないファイルとして開く。小さなファイルが展開すると何 GB にもなる（圧縮の爆弾）、
参照が輪になっていて読み終わらない、といった書類で、メモリや時間を使い切らないよう、次の上限を置く。
上限は「壊れた・悪意のある」ファイルを止めるためのもので、大きいだけの本物のファイルを止めるためのものではない。
大きな書類は読む量を分ける（lazy/：概要を返し、エージェントが要る所をページ・行の範囲で頼む）。読めない所は作らない。
上限に当たったときは LimitExceeded（ValueError の一種）を投げ、エージェントには、頼み直せる小さな範囲を伝える（isolate.smaller_request）。

1 回の道具の呼び出し（窓を開く・書き込む・検査する）ごとに Budget を 1 つ作り（call_budget）、その呼び出しの中の
展開（ZIP の部品・PDF のストリーム・一太郎の圧縮）と、PDF の描画の命令を、同じ Budget で数える。
呼び出しの外（ライブラリとして使うとき）は、1 回の展開の上限だけを守る（合計は数えない）。

環境変数（大きさは MB）：
- TODOKU_MAX_FILE_MB（メモリの上限の半分）  メモリに丸ごと読むファイルの大きさ。これより大きなファイルは断らずに、写像（lazy/filebytes）で読む
                                        （分けて読める形式は、もともとディスクから少しずつ読む）
- TODOKU_MAX_INFLATE_PART_MB（256）    1 つの部品・ストリームを展開した大きさの基本の上限。実際の上限は、これと「圧縮したバイト × 比率」の大きい方
                                        （宣言された大きさではなく、実際に出てきたバイトを数える）
- TODOKU_MAX_INFLATE_TOTAL_MB（768 と、メモリの上限の半分の小さい方）  1 回の呼び出しで展開するバイトの合計（同じく、これと
                                        「ファイルの大きさ × 比率」の大きい方）。メモリの上限より先に、理由の言える形で止まるように、小さく置く
- TODOKU_MAX_INFLATE_RATIO（200）      圧縮の比率の上限。ふつうの XML は 10〜100 倍で、deflate の理屈の上限は 1032 倍。爆弾はここで止まる
- TODOKU_CONTEXT_BUDGET_CHARS（400000） 窓の全体を一度に返してよい大きさ（字）。およそ 20 万トークンの会話に入らなくなる大きさ。
                                        これを超える書類は、概要だけを返し、ページ・行の範囲で読む（lazy/）
- TODOKU_LAZY_MIN_KB（4096）            分けて読む道（lazy）に入る、中身のバイトの下限。これより小さな Word・一太郎・.xls などは窓を作ってから分ける
- TODOKU_PAGE_CHARS（60000）           範囲で読むときの 1 回の大きさ（字）。続きは cursor で読める（読み減らすだけで、読めない所は作らない）
- TODOKU_MAX_PDF_OPS（200000000）     1 回の呼び出しで読む、PDF の描画の字句（数・名前・命令）の数の合計
- TODOKU_MAX_PAGE_OPS（5000000）       1 ページ（とその部品）で読む、PDF の描画の字句の数
- TODOKU_MAX_PIXELS（100000000）       1 つの画像を画素に開く大きさ（幅 × 高さ）
- TODOKU_MAX_ITEMS（メモリの上限 ÷ 512 バイト）  1 回の呼び出しで、宣言された数から作る項目の合計（字形の表・/W の組・XFA のくり返し・
                                        セル・XML の要素・配列の要素など。budget().count）。実際の上限は、これと「ファイルの大きさ × 4」の大きい方
- TODOKU_TIME_LIMIT（120）             1 回の呼び出しの秒数（MCP サーバー：別のプロセスで動かし、超えたら止める）
- TODOKU_MEMORY_LIMIT_MB（物理メモリの 1/4。2048 以上）  その別のプロセスのメモリ
- TODOKU_WINDOW_CACHE_MB（256）        窓の控え（同じ書類を開き直さない）の大きさの合計
- TODOKU_MAX_XFA_BLANK（10000）         XFA：データのないくり返し（occur の min・initial）を作る数（1 つの書類）。超える分は作らず、窓に事実として書く
- TODOKU_MAX_DECLARED_CELLS（262144）  表計算：ファイルの宣言（入力規則・数式の参照・配列数式の範囲・繰り返し）から、セルの要素のない所を 1 マスずつの欄として
                                        一度に作る数（1 ブック）。超える範囲は「範囲の所」（xlsx/places.py）として同じ事実を持つ（読めない所は作らない）
- TODOKU_MAX_ARRAY_CELLS（4194304）     表計算：数式の範囲を計算のために一度にリストに開く数。超える範囲は、頼まれたマスだけ読む（formula/lazyarr.py）
- TODOKU_CALC_SECONDS（TODOKU_TIME_LIMIT の 1/4）  表計算：1 回の呼び出しで、式で答えられない大きな配列の集計を読み進める秒数。
                                        終わらなければ、進んだ所と続け方を伝える（長い仕事：jobs.py）
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



def _physical_memory() -> int:
    """この機械の物理メモリ（バイト）。分からなければ 0。"""
    try:
        return os.sysconf("SC_PAGE_SIZE") * os.sysconf("SC_PHYS_PAGES")
    except (AttributeError, ValueError, OSError):
        return 0


MAX_INFLATE_PART = _env("TODOKU_MAX_INFLATE_PART_MB", 256, MB)
MAX_PDF_OPS = _env("TODOKU_MAX_PDF_OPS", 200_000_000)
MAX_PAGE_OPS = _env("TODOKU_MAX_PAGE_OPS", 5_000_000)
MAX_PIXELS = _env("TODOKU_MAX_PIXELS", 100_000_000)
TIME_LIMIT = _env("TODOKU_TIME_LIMIT", 120)
MEMORY_LIMIT = _env("TODOKU_MEMORY_LIMIT_MB", max(2048, _physical_memory() / 4 / MB), MB)
MAX_FILE = _env("TODOKU_MAX_FILE_MB", MEMORY_LIMIT / 2 / MB, MB)
# 展開の合計は、メモリの上限より十分に小さく置く（同じ大きさだと、展開の予算で理由を言う前に、働き手が殺される。展開したものをつなぐと
# 一時に 2 倍を使う）。大きなファイルは allow_file で広がるので、本物の書類には当たらない
MAX_INFLATE_TOTAL = _env("TODOKU_MAX_INFLATE_TOTAL_MB", min(768, MEMORY_LIMIT / 2 / MB), MB)
# 1 回の呼び出しで作る項目の合計：Python の小さな値は 1 つ数十〜数百バイトなので、メモリの上限より先に、理由の言える形で止まる大きさ
MAX_ITEMS = _env("TODOKU_MAX_ITEMS", MEMORY_LIMIT / 512)
ITEMS_PER_FILE_BYTE = 4  # 本物の書類の項目は、ファイルのバイトに比例する（圧縮した XML でも、1 バイトから数個まで）
MAX_INFLATE_RATIO = _env("TODOKU_MAX_INFLATE_RATIO", 200)
CONTEXT_BUDGET = _env("TODOKU_CONTEXT_BUDGET_CHARS", 400_000)
PAGE_CHARS = _env("TODOKU_PAGE_CHARS", 60_000)
LAZY_MIN_BYTES = _env("TODOKU_LAZY_MIN_KB", 4096, 1024)
WINDOW_CACHE = _env("TODOKU_WINDOW_CACHE_MB", 256, MB)
# マクロ（VBA）：1 つの文字列の長さ・1 つの配列の要素の数・1 回の実行で作る大きな値（1 MB 以上の文字列）の合計
MAX_VBA_STRING = _env("TODOKU_MAX_VBA_STRING", 16 * MB)
MAX_VBA_ARRAY = _env("TODOKU_MAX_VBA_ARRAY", 10_000_000)
MAX_VBA_BYTES = _env("TODOKU_MAX_VBA_BYTES_MB", 1024, MB)
MAX_VBA_RANGE = _env("TODOKU_MAX_VBA_RANGE", 2_000_000)  # マクロが 1 度に書く範囲のマスの数（1 列まるごと 1,048,576 は書ける）
MAX_XFA_BLANK = _env("TODOKU_MAX_XFA_BLANK", 10_000)
MAX_DECLARED_CELLS = _env("TODOKU_MAX_DECLARED_CELLS", 262_144)
MAX_ARRAY_CELLS = _env("TODOKU_MAX_ARRAY_CELLS", 4_194_304)
CALC_SECONDS = _env("TODOKU_CALC_SECONDS", TIME_LIMIT / 4)

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

    def __init__(self, part: int | None = None, total: int | None = None, ops: int | None = None, items: int | None = None):
        self.part = MAX_INFLATE_PART if part is None else part
        self.total = total
        self.ops_total = ops if ops is not None else float("inf")
        self.inflated = 0
        self.ops = 0
        # items=None（長い仕事・呼び出しの外）：合計は数えず、1 度に頼む数だけを上限と比べる
        self.items_total = items
        self.items_cap = MAX_ITEMS if items is None else items
        self.items = 0
        self.calc_left: float = CALC_SECONDS  # 大きな配列の集計に使ってよい秒数の残り（長い仕事は無限）

    def allow_file(self, size: int) -> None:
        """大きなファイル：合計の上限を、ファイルの大きさに見合うだけ広げる（小さなファイルの爆弾は、いままでの上限で止まる）。"""
        if self.total is not None:
            self.total = max(self.total, size * MAX_INFLATE_RATIO)
        self.items_cap = max(self.items_cap, size * ITEMS_PER_FILE_BYTE)
        if self.items_total is not None:
            self.items_total = self.items_cap

    def count(self, n: int = 1, what: str = "項目") -> None:
        """これから n 個の項目を作る（宣言された数から広げる展開はすべて、作る前にここを通す：1 つずつの上限の内でも、合わせて止まる）。
        上限を超えたら LimitExceeded（読む範囲を小さく頼み直せば読める：smaller_request）。"""
        self.items += n
        if n > self.items_cap or (self.items_total is not None and self.items > self.items_total):
            raise LimitExceeded(f"{what}：1 回の呼び出しで作る項目が {self.items_cap:,} 個を超えます（大きな数が宣言された壊れた書類か、"
                                "一度に読む範囲が広すぎます。ページ・行の範囲を小さくして頼み直してください。"
                                "TODOKU_MAX_ITEMS で変えられます）", True)

    def count_part(self, part: str, n: int, what: str = "") -> None:
        """部品 part の中の項目の数（XML の要素・CSV のマス）。同じ部品を読み直しても（検査で書いては読み直す）、いちばん大きな数だけを 1 度数える。"""
        seen = self.__dict__.setdefault("parts_seen", {})
        had = seen.get(part, 0)
        if n > had:
            seen[part] = n
            self.count(n - had, what or f"{part} の要素")

    def alloc(self, n: int, what: str = "") -> None:
        """これから 1 度に n バイトを取る（C の中の 1 回の確保は途中で止められないので、取る前に確かめる）。"""
        if n > MEMORY_LIMIT // 2:
            raise LimitExceeded(f"{what or 'データ'}：1 度に {_mb(n)} を取ることになります（メモリの上限 {_mb(MEMORY_LIMIT)} の半分まで）。"
                                "範囲や解像度を小さくして頼み直してください")

    def room(self, part: int | None = None, compressed: int = 0) -> int:
        """これから 1 つ展開してよい大きさ。compressed：展開するデータの圧縮したバイト数。大きなデータには、比率の分だけ広げる
        （part を指定した展開は、その指定のとおり）。"""
        cap = self.part if part is None else min(part, self.part)
        if part is None and compressed:
            cap = max(cap, compressed * MAX_INFLATE_RATIO)
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
    token = _CURRENT.set(b or Budget(total=MAX_INFLATE_TOTAL, ops=MAX_PDF_OPS, items=MAX_ITEMS))
    try:
        yield _CURRENT.get()
    finally:
        _CURRENT.reset(token)


def inflate(data: bytes, wbits: int = 15, what: str = "", part: int | None = None, tolerant: bool = False) -> bytes:
    """zlib・deflate を、出てくるバイトを数えながら展開する（宣言された大きさは信じない）。
    part：この展開だけの上限（予算の 1 つの展開の上限より小さくするとき）。上限を超えたら LimitExceeded。
    tolerant：壊れたデータを、読めるところまで返す（PDF）。そうでなければ zlib.error（途中で終わるデータも）。"""
    b = budget()
    cap = b.room(part, len(data))
    d = zlib.decompressobj(wbits)
    pieces: list[bytes] = []  # 少しずつ伸ばさない（伸ばすたびに写すと、大きなデータで一時に何倍も使う）
    got = 0
    src = data
    try:
        while True:
            piece = d.decompress(src, min(1 << 22, cap - got + 1))
            pieces.append(piece)
            got += len(piece)
            if got > cap:
                raise b.exceeded(cap, part, what)
            src = d.unconsumed_tail
            if d.eof or (not src and not piece):
                break
        if not d.eof and not tolerant:
            raise zlib.error("圧縮したデータが途中で終わっています")
    except zlib.error:
        if not tolerant:
            raise
    b.charge(got, what)
    out = b"".join(pieces)
    if got >= 1 << 26:
        pieces.clear()
        release_memory()
    return out


class StreamInflate:
    """展開を、出てくるバイトを数えながら少しずつ進める（丸ごとメモリに置かない。圧縮の爆弾は比率の上限で止まる）。"""

    def __init__(self, what: str = "", wbits: int = -15, compressed: int = 0):
        self.what = what
        self.d = zlib.decompressobj(wbits)
        self.out = 0
        self.compressed = compressed
        self.cap = max(MAX_INFLATE_PART, compressed * MAX_INFLATE_RATIO)

    def feed(self, data: bytes, max_out: int = 0) -> bytes:
        piece = self.d.decompress(data, max_out) if max_out else self.d.decompress(data)
        self.out += len(piece)
        if self.out > self.cap:
            raise LimitExceeded(f"{self.what or '部品'}：圧縮した {_mb(self.compressed)} が {_mb(self.cap)} を超えて展開されます"
                                f"（比率 {MAX_INFLATE_RATIO} 倍まで。TODOKU_MAX_INFLATE_RATIO で変えられます）。圧縮の爆弾か、壊れたファイルです")
        return piece


# ---------------------------------------------------------------------- 進み具合
_PROGRESS = {"report": None}


def set_progress_reporter(fn) -> None:
    """働き手（isolate）が、進み具合の伝え先を置く。"""
    _PROGRESS["report"] = fn


def progress(text: str, done: int = 0, total: int = 0) -> None:
    """長い処理の進み具合を伝える（止まったときに、どこまで進んだかを言えるように。止める判断には使わない）。"""
    fn = _PROGRESS["report"]
    if fn is not None:
        fn(text, done, total)


def release_memory() -> None:
    """使い終えた大きなメモリを OS に返す（macOS の malloc は、大きな塊を自分の控えに残し、常駐に数えられたままになる）。"""
    global _RELEASE
    if _RELEASE is None:
        import ctypes
        import sys

        _RELEASE = lambda: None  # noqa: E731
        try:
            if sys.platform == "darwin":
                fn = ctypes.CDLL("libSystem.dylib").malloc_zone_pressure_relief
                fn.argtypes = [ctypes.c_void_p, ctypes.c_size_t]
                fn.restype = ctypes.c_size_t
                _RELEASE = lambda: fn(None, 0)
            elif sys.platform.startswith("linux"):
                trim = ctypes.CDLL(None).malloc_trim
                trim.argtypes = [ctypes.c_size_t]
                _RELEASE = lambda: trim(0)
        except (OSError, AttributeError):
            pass
    _RELEASE()


_RELEASE = None


class CalcClock:
    """大きな配列の集計に使ってよい時間：1 回の呼び出しの予算で数える（呼び出しの外は、集計ごとに持ち時間。formula/lazyarr.py）。"""

    def current(self) -> Budget:
        b = budget()
        return Budget() if b is _LOOSE else b

    def left(self, b: Budget) -> float:
        return max(0.0, b.calc_left)

    def spend(self, b: Budget, seconds: float) -> None:
        b.calc_left -= seconds

    def progress(self, text: str, done: int, total: int) -> None:
        progress(text, done, total)


def check_output(n: int, what: str = "", part: int | None = None) -> None:
    """展開して n バイトになるか（展開の前に、出てくる大きさが分かるとき）。上限を超えるなら LimitExceeded。"""
    b = budget()
    cap = b.room(part)
    if n > cap:
        raise b.exceeded(cap, part, what)
