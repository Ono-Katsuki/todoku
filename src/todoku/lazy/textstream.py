"""文字だけのファイルを、丸ごとメモリに置かずに読む：文字コードの判定（datafile.textenc と同じ答え）・行（記録）の位置の索引。

- ファイルは os.pread で少しずつ読む（mmap にしない：触ったページが常駐メモリに数えられ、メモリの見張りに当たるため）
- 文字コードは textenc.decode と同じ順に決める。字の並びはファイルの全部を見る（増分のデコーダ。メモリは一定）
- 索引：STRIDE 個ごとの記録の始まりのバイト位置だけを array に持つ（行数に比例するが、1 行あたり 1 バイト未満）
- 区切りのバイト（改行・「,」・タブ・「;」・引用符）は、CP932・EUC-JP・Big5・GBK・UHC の 2 バイト目には現れない（どれも 0x40 以上か
  0xA1 以上）ので、バイトのまま探せる。UTF-16・UTF-32 は、字の単位（2・4 バイト）にそろった所だけを見る
- 切り替えの印で字の読み方が変わる文字コード（ISO-2022・HZ・UTF-7）は、同じバイトが、前の印によって別の字になるので、
  バイトのままでは探せない：字を UTF-8 に直した写し（recode）の上で、同じ読み手を使う
"""

from __future__ import annotations

import codecs
import functools
import hashlib
import os
from .pread import O_BINARY, pread
import re
from array import array

from ..datafile.textenc import (
    BOMS,
    ISO2022,
    MULTI,
    NAME_JA,
    WIDE,
    Decoded,
    EncodingError,
    _HIGH,
    _clip,
    breaks_before,
    canonical,
    clean,
    controls,
    declared,
    differing,
    first_offset,
    high_bytes,
    single_readings,
    snippet,
    stateful,
)

CHUNK = 1 << 20
STRIDE = 64  # 索引に載せる間隔（記録の数）。ページを読むときに余分に読むのは、高々この数の記録


class StreamDecoded(Decoded):
    """Decoded と同じ問い合わせに答える（text は空。字はファイルの範囲ごとにデコードする）。"""

    has_nul = False
    first_sample = ""

    @property
    def binary(self) -> bool:
        return self.basis in ("none", "ascii") and self.has_nul

    @staticmethod
    def sample_of(text, cand) -> str:  # 窓の文字コードで読んだ、字の違う最初の行（見本）
        return cand.get("own_sample", "")


class MemorySource:
    """メモリの中のバイト列を、Source と同じ形で読む（文書に埋め込んだブック）。ファイルに書き出さない。"""

    def __init__(self, data: bytes):
        self.path = None
        self.data = data
        self.size = len(data)
        self.key = (len(data), 0)

    def read(self, a: int, b: int) -> bytes:
        return self.data[a:b]

    def chunks(self, a: int = 0, b: int | None = None, size: int = CHUNK):
        b = self.size if b is None else min(b, self.size)
        while a < b:
            yield self.data[a:min(a + size, b)]
            a += size

    def close(self) -> None:
        pass


class Source:
    """開いたファイル（読むだけ）。"""

    def __init__(self, path: str):
        self.path = path
        self.fd = os.open(path, os.O_RDONLY | O_BINARY)
        st = os.fstat(self.fd)
        self.size = st.st_size
        self.key = (st.st_size, st.st_mtime_ns)

    def read(self, a: int, b: int) -> bytes:
        out = []
        while a < b:
            piece = pread(self.fd, min(b - a, 1 << 26), a)
            if not piece:
                break
            out.append(piece)
            a += len(piece)
        return out[0] if len(out) == 1 else b"".join(out)

    def chunks(self, a: int = 0, b: int | None = None, size: int = CHUNK):
        b = self.size if b is None else b
        while a < b:
            piece = pread(self.fd, min(size, b - a), a)
            if not piece:
                return
            yield piece
            a += len(piece)

    def sha256(self) -> str:
        h = hashlib.sha256()
        for c in self.chunks():
            h.update(c)
        return h.hexdigest()

    def close(self) -> None:
        if self.fd >= 0:
            os.close(self.fd)
            self.fd = -1

    def __del__(self):
        try:
            self.close()
        except (OSError, AttributeError, TypeError):  # インタプリタの終わりでは、os がもうないことがある
            pass


# ---------------------------------------------------------------------- 文字コード
def _strict_stream(src: Source, enc: str, start: int = 0):
    """(読めたか, きれいか, 制御文字の数, 字のハッシュ, 行の区切りがあるか)。読めなければ (False, …)。"""
    dec = codecs.getincrementaldecoder(enc)("strict")
    h = hashlib.sha1()
    is_clean, ctl, nl = True, 0, False
    try:
        for c in src.chunks(start):
            t = dec.decode(c)
            if is_clean and not t.isascii() and not clean(t):
                is_clean = False
            ctl += controls(t) if not _no_control(t) else 0
            nl = nl or "\n" in t or "\r" in t
            h.update(t.encode("utf-8", "surrogatepass"))
        t = dec.decode(b"", True)
        h.update(t.encode("utf-8", "surrogatepass"))
    except UnicodeDecodeError:
        return False, False, 0, b"", False
    return True, is_clean, ctl, h.digest(), nl


_CTL = re.compile(r"[\x00-\x08\x0e-\x1f\x7f]")


def _no_control(t: str) -> bool:
    return _CTL.search(t) is None


def _lines_of(src: Source, enc: str, start: int):
    """デコードした行を順に（文字コードの候補の見本を作るだけ）。"""
    dec = codecs.getincrementaldecoder(enc)("strict")
    carry = ""
    for c in src.chunks(start):
        carry += dec.decode(c)
        lines = carry.splitlines(True)
        carry = lines.pop() if lines and not lines[-1].endswith(("\n", "\r")) else ""
        yield from lines
    if carry:
        yield carry


def _first_diff(src: Source, e1: str, e2: str) -> tuple[int, str, str]:
    for k, (x, y) in enumerate(zip(_lines_of(src, e1, 0), _lines_of(src, e2, 0))):
        if x != y:
            return k, _clip(x.rstrip("\r\n")), _clip(y.rstrip("\r\n"))
    return 0, "", ""


def _legacy(src: Source) -> StreamDecoded | None:
    """textenc._legacy と同じ答え。多バイトは増分のデコーダでファイルの全部を読み、1バイトは、ファイルにあるバイトの種類を
    1 回だけ集めて、文字コードごとの表で見る（文字コードの数だけ読み直さない）。"""
    multi: list[tuple[str, bytes, int]] = []
    for enc in MULTI:
        ok, is_clean, ctl, digest, _nl = _strict_stream(src, enc)
        if ok and is_clean and all(digest != d for _e, d, _c in multi):
            multi.append((enc, digest, ctl))
    high = high_bytes(src.chunks())
    single = single_readings(high)
    if not multi and not single:
        return None
    lines: dict[int, int] = {}

    def cand(e2: str, off: int, own: str) -> dict:
        if off not in lines:
            lines[off] = breaks_before(src.chunks(), off)
        return {"encoding": e2, "label": NAME_JA.get(e2, e2), "line": lines[off],
                "sample": snippet(src.read, src.size, off, e2), "own_sample": snippet(src.read, src.size, off, own)}

    if multi:
        enc, _d, ctl = multi[0]
        cands = []
        for e2, _d2, _c2 in multi[1:]:
            k, mine, theirs = _first_diff(src, enc, e2)
            cands.append({"encoding": e2, "label": NAME_JA.get(e2, e2), "line": k, "sample": theirs, "own_sample": mine})
        off = first_offset(src.chunks(), _HIGH)
        cands += [cand(e2, off, enc) for e2, _key in single]
        basis = "multibyte"
    else:
        enc, key = single[0]
        ctl = _strict_stream(src, enc)[2]
        cands = [cand(e2, first_offset(src.chunks(), differing(high, key, k2)), enc) for e2, k2 in single[1:]]
        basis = "singlebyte"
    return StreamDecoded(enc, b"", "", False, basis if not cands else "order", cands, ctl)


def detect(src: Source, encoding: str | None = None) -> StreamDecoded:
    """textenc.decode と同じ順に決める（字の全部を見るが、メモリは一定）。UnicodeDecodeError・EncodingError も同じ。"""
    if encoding:
        return _given(src, encoding)
    head = src.read(0, 4)
    for bom, enc in BOMS:
        if head.startswith(bom):
            ok, _cl, ctl, _d, _nl = _strict_stream(src, enc, len(bom))
            if not ok and enc == "utf-32-le":
                continue
            if not ok:
                raise UnicodeDecodeError(enc, head, len(bom), src.size, "BOM の文字コードとして読めません")
            d = StreamDecoded(enc, bom, "", False, "bom", [], ctl)
            return d
    has_nul, ascii_only, ctl, esc = False, True, 0, 0
    for c in src.chunks():
        has_nul = has_nul or b"\x00" in c
        if ascii_only and not c.isascii():
            ascii_only = False
        esc += c.count(b"\x1b") if ascii_only else 0
    if has_nul:
        best = None
        got = []
        for enc in WIDE:
            ok, is_clean, c_, digest, nl = _strict_stream(src, enc)
            if ok and is_clean:
                got.append((enc, digest, c_, nl))
        if got and got[0][2] == 0 and got[0][3]:
            cands = []
            first = got[0]
            for e2, d2, _c, _n in got[1:]:
                if d2 != first[1]:
                    k, mine, theirs = _first_diff(src, first[0], e2)
                    cands.append({"encoding": e2, "label": NAME_JA.get(e2, e2), "line": k, "sample": theirs, "own_sample": mine})
            best = StreamDecoded(first[0], b"", "", False, "nul" if not cands else "order", cands, 0)
        if best is not None:
            best.has_nul = True
            return best
    if ascii_only:
        got3 = _escaped(src, esc) if esc and not has_nul else None
        if got3 is not None:
            return got3
        for c in src.chunks():
            ctl += controls(c.decode("ascii")) if _CTL.search(c.decode("ascii")) else 0
        d = StreamDecoded("utf-8", b"", "", True, "ascii", [], ctl)
        d.has_nul = has_nul
        return d
    ok, _cl, ctl, _d, _nl = _strict_stream(src, "utf-8")
    if ok:
        return StreamDecoded("utf-8", b"", "", False, "utf-8", [], ctl)
    if not has_nul:
        decl = declared(src.read(0, 4096))
        e = canonical(decl[0]) if decl else None
        if e and e not in WIDE:
            ok, is_clean, ctl, _d, _nl = _strict_stream(src, e)
            if ok and is_clean:
                return StreamDecoded(e, b"", "", False, "declared", [], ctl, decl[1])
        got2 = _legacy(src)
        if got2 is not None:
            return got2
    ctl = 0
    for c in src.chunks():
        t = c.decode("latin-1")
        ctl += controls(t) if _CTL.search(t) else 0
    d = StreamDecoded("latin-1", b"", "", False, "none", [], ctl)
    d.has_nul = has_nul
    return d


def _escaped(src: Source, esc: int) -> StreamDecoded | None:
    """ASCII のバイトだけで ESC のあるファイルを、ISO/IEC 2022 の切り替えの印のとおりに読む（textenc の決め方の 3'）。"""
    got = []
    for enc in ISO2022:
        dec = codecs.getincrementaldecoder(enc)("strict")
        h = hashlib.sha1()
        ok, ctl, left = True, 0, 0
        try:
            for c in src.chunks():
                t = dec.decode(c)
                if "\x0e" in t or "\x0f" in t or (not t.isascii() and not clean(t)):
                    ok = False
                    break
                left += t.count("\x1b")
                ctl += controls(t) if not _no_control(t) else 0
                h.update(t.encode("utf-8", "surrogatepass"))
            if ok:
                t = dec.decode(b"", True)
                h.update(t.encode("utf-8", "surrogatepass"))
        except UnicodeDecodeError:
            continue
        if ok and left < esc and all(h.digest() != d for _e, d, _c in got):
            got.append((enc, h.digest(), ctl))
    if not got:
        return None
    first = got[0]
    cands = []
    for e2, _d, _c in got[1:]:
        k, mine, theirs = _first_diff(src, first[0], e2)
        cands.append({"encoding": e2, "label": NAME_JA.get(e2, e2), "line": k, "sample": theirs, "own_sample": mine})
    return StreamDecoded(first[0], b"", "", False, "escape" if not cands else "order", cands, first[2])


RECODE_IN_MEMORY = 64 << 20  # これより小さいファイルの写しは、メモリに置く（大きいものは一時ファイル）


class RecodedSource(Source):
    """切り替えの印のある文字コードのファイルを、UTF-8 に直した写し（一時ファイル。閉じると消える）。"""

    def __init__(self, src, enc: str, start: int):
        import tempfile

        fd, tmp = tempfile.mkstemp(prefix="todoku-", suffix=".utf8")
        self.path = None
        self.fd = -1
        self._tmp = tmp
        try:
            with os.fdopen(fd, "wb") as f:
                for piece in _recoded(src, enc, start):
                    f.write(piece)
            self.fd = os.open(tmp, os.O_RDONLY | O_BINARY)
            self.size = os.fstat(self.fd).st_size
        except BaseException:
            self.close()
            raise
        self.key = src.key
        try:  # 開いたまま消せる OS では、ここで消す（止まっても残らない）。消せない OS では、閉じるときに消す
            os.unlink(tmp)
            self._tmp = None
        except OSError:
            pass

    def close(self) -> None:
        super().close()
        if self._tmp:
            try:
                os.unlink(self._tmp)
                self._tmp = None
            except OSError:
                pass


class SpilledSource(Source):
    """書庫から取り出した大きな中身を、一時ファイルに置いて読む（メモリに置いたままにしない。閉じると消える）。"""

    def __init__(self, data: bytes):
        import tempfile

        fd, tmp = tempfile.mkstemp(prefix="todoku-", suffix=".member")
        self.path = None
        self.fd = -1
        self._tmp = tmp
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(data)
            self.fd = os.open(tmp, os.O_RDONLY | O_BINARY)
            self.size = os.fstat(self.fd).st_size
        except BaseException:
            self.close()
            raise
        self.key = (len(data), 0)
        try:  # 開いたまま消せる OS では、ここで消す（止まっても残らない）。消せない OS では、閉じるときに消す
            os.unlink(tmp)
            self._tmp = None
        except OSError:
            pass

    def close(self) -> None:
        super().close()
        if self._tmp:
            try:
                os.unlink(self._tmp)
                self._tmp = None
            except OSError:
                pass


def member_source(data: bytes):
    """書庫から取り出した中身の読み手：小さければメモリ、大きければ一時ファイル（切り替えの印のある文字コードの写しと同じ境）。"""
    return MemorySource(data) if len(data) <= RECODE_IN_MEMORY else SpilledSource(data)


def _recoded(src, enc: str, start: int):
    dec = codecs.getincrementaldecoder(enc)("strict")
    try:
        for c in src.chunks(start):
            t = dec.decode(c)
            if t:
                yield t.encode("utf-8")
        t = dec.decode(b"", True)
        if t:
            yield t.encode("utf-8")
    except UnicodeError:
        raise EncodingError(f"このファイルの字を、{NAME_JA.get(enc, enc)} から UTF-8 に直せませんでした") from None


def recode(src, dec: StreamDecoded):
    """切り替えの印のある文字コードのファイル → 字を UTF-8 に直した写しと、その写しの読み方。字は同じ（厳密に読んで、厳密に書く）。
    写しの上では、行・区切り・引用符をバイトのまま探せる。"""
    if src.size <= RECODE_IN_MEMORY:
        out = MemorySource(b"".join(_recoded(src, dec.encoding, len(dec.bom))))
        out.key = src.key
    else:
        out = RecodedSource(src, dec.encoding, len(dec.bom))
    return out, StreamDecoded("utf-8", b"", "", False, "given", [], dec.controls)


def _given(src: Source, encoding: str) -> StreamDecoded:
    enc = canonical(encoding)
    if enc is None:
        raise EncodingError(f"文字コード「{encoding}」を知りません（例：UTF-8・Shift_JIS・EUC-JP・GBK・Big5・EUC-KR・Windows-1252）")
    head = src.read(0, 4)
    if enc in ("utf-16", "utf-32"):
        hit = next(((b, e) for b, e in BOMS if e.startswith(enc) and head.startswith(b)), None)
        if hit is None:
            raise EncodingError(f"{encoding}：BOM がないので、向きが決まりません。{enc}-le か {enc}-be を指定してください")
        enc = hit[1]
    bom = next((b for b, e in BOMS if e == enc and head.startswith(b)), b"")
    try:
        ok, _cl, ctl, _d, _nl = _strict_stream(src, enc, len(bom))
    except LookupError:
        raise EncodingError(f"文字コード「{encoding}」を知りません") from None
    if not ok:
        raise EncodingError(f"指定された文字コード（{NAME_JA.get(enc, enc)}）では、このファイルを厳密に読めません")
    return StreamDecoded(enc, bom, "", False, "given", [], ctl)


# ---------------------------------------------------------------------- 記録の位置
def unit_of(enc: str) -> int:
    return 4 if enc.startswith("utf-32") else 2 if enc.startswith("utf-16") else 1


class Window:
    """ファイルの上を動く読み出しの枠（search は枠をまたいで探す）。"""

    def __init__(self, src: Source, start: int, unit: int):
        self.src = src
        self.base = start
        self.buf = src.read(start, start + CHUNK)
        self.unit = unit
        self.origin = start  # 字の単位をそろえる基準

    def search(self, pat: re.Pattern, pos: int):
        """pos 以降で pat に合う所：(ファイルの中の始め, 終わり, 合った字)。なければ None。"""
        src = self.src
        overlap = 8
        while True:
            if not (self.base <= pos <= self.base + len(self.buf)):
                self.base = pos
                self.buf = src.read(pos, pos + CHUNK)
            at = pos - self.base
            while True:
                m = pat.search(self.buf, at)
                if m is None:
                    near = None
                    break
                if self.unit > 1 and (m.start() + self.base - self.origin) % self.unit:
                    at = m.start() + 1
                    continue
                # 枠の終わりにかかる合い方（CR の次の LF が枠の外など）は、枠を動かして読み直す
                if m.end() >= len(self.buf) - overlap and self.base + len(self.buf) < src.size:
                    near = self.base + m.start()
                    break
                return self.base + m.start(), self.base + m.end(), m.group()
            end = self.base + len(self.buf)
            if end >= src.size:
                return None
            pos = near if near is not None else max(pos, end - overlap - (end - overlap - self.origin) % self.unit)
            self.base = pos
            self.buf = src.read(pos, pos + CHUNK)


class TextIndex:
    """記録（CSV の行・文字のファイルの行）の位置の索引。"""

    def __init__(self, end_of_data: int = 0):
        self.count = 0
        self.starts = array("Q")
        self.end_of_data = end_of_data
        self.terms = {"\r\n": 0, "\n": 0, "\r": 0}
        self.longest = 0  # いちばん長い記録のバイト数
        self.ends_with_newline = False

    def add(self, pos: int) -> None:
        if self.count % STRIDE == 0:
            self.starts.append(pos)
        self.count += 1

    def bucket(self, first: int, last: int) -> tuple[int, int, int]:
        """記録 first〜last（1 から）を含むファイルの範囲：(始めのバイト, 終わりのバイト, 範囲の頭の記録の番号)。"""
        i = (first - 1) // STRIDE
        j = (last - 1) // STRIDE + 1
        a = self.starts[i]
        b = self.starts[j] if j < len(self.starts) else self.end_of_data
        return a, b, i * STRIDE + 1

    @property
    def newline(self) -> str:
        return max(("\r\n", "\n", "\r"), key=lambda t: (self.terms[t], t == "\r\n")) if any(self.terms.values()) else ""

    @property
    def mixed_newlines(self) -> bool:
        return sum(1 for v in self.terms.values() if v) > 1


def tokens(enc: str, *chars: str) -> list[bytes]:
    return [c.encode(enc) for c in chars]


_NL = re.compile(rb"\r\n|\n|\r")


def index_lines(src: Source, start: int, enc: str) -> TextIndex:
    """行（CRLF・LF・CR で区切る）の索引。最後の改行の後ろに字がなければ、そこに行は作らない。"""
    from ..limits import progress

    size = src.size
    idx = TextIndex(size)
    if size <= start:
        return idx
    idx.add(start)
    unit = unit_of(enc)
    rec = start
    if unit == 1:
        pos = start
        carry_cr = False
        for c in src.chunks(start):
            c0 = 0
            if carry_cr:  # 前の枠の終わりが CR：次のバイトが LF なら CRLF
                carry_cr = False
                if c[:1] == b"\n":
                    idx.terms["\r\n"] += 1
                    c0 = 1
                else:
                    idx.terms["\r"] += 1
                rec = pos + c0
                if rec < size:
                    idx.add(rec)
            for m in _NL.finditer(c, c0):
                e = m.end()
                t = m.group()
                n = pos + m.start() - rec
                if n > idx.longest:
                    idx.longest = n
                if e == len(c) and t == b"\r" and pos + e < size:
                    carry_cr = True
                    break
                idx.terms[t.decode()] += 1
                rec = pos + e
                if rec < size:
                    idx.add(rec)
            pos += len(c)
            progress("行の位置を調べています", pos, size)
        if carry_cr:
            idx.terms["\r"] += 1
            rec = size
    else:
        win = Window(src, start, unit)
        pat = re.compile(b"|".join(re.escape(x) for x in tokens(enc, "\r\n", "\n", "\r")))
        pos = start
        while True:
            hit = win.search(pat, pos)
            if hit is None:
                break
            a, b, t = hit
            idx.terms[t.decode(enc)] += 1
            idx.longest = max(idx.longest, a - pos)
            pos = rec = b
            if b < size:
                idx.add(b)
    idx.longest = max(idx.longest, size - rec)
    idx.ends_with_newline = rec == size
    return idx


class CsvStats:
    """ファイル全体で 1 度だけ調べる、CSV の書き方の事実。"""

    def __init__(self):
        self.width = 0
        self.quote_all = True  # 字のあるマスが、すべて引用符で囲まれている（字のあるマスがなければ False にする）
        self.any_filled = False
        self.any_empty = False
        self.all_empty_quoted = True
        self.empty_lines = 0
        self.first_empty_lines: list[int] = []
        self.problem_rows = 0
        self.first_problems: list[str] = []
        self.nul_rows: list[int] = []
        self.quoted_any = False

    @property
    def quote_empty(self) -> bool:
        return self.quote_all and self.any_empty and self.all_empty_quoted

    def problem(self, text: str) -> None:
        self.problem_rows += 1
        if len(self.first_problems) < 100:
            self.first_problems.append(text)


def head_text(src: Source, start: int, enc: str) -> str:
    """区切りを数える、頭の 200,000 字。"""
    dec = codecs.getincrementaldecoder(enc)("replace")
    return dec.decode(src.read(start, start + 800_000))[:200_000]


def sniff_delimiter(src: Source, start: int, enc: str, kind: str) -> tuple[str, list, bool]:
    """datafile.csvfile.delimiters と同じ規則・同じ答え（頭の 200,000 字・200 行）。"""
    from ..datafile.csvfile import delimiters

    return delimiters(head_text(src, start, enc), kind)


@functools.lru_cache(maxsize=None)
def second_byte_leads(enc: str, delim: str) -> frozenset:
    """区切りのバイトが、この文字コードの 2 バイトの字の 2 バイト目にもなるとき、1 バイト目になるバイトの集まり
    （Shift_JIS の「ポ」は 0x83 0x7C で、2 バイト目が「|」と同じ）。ならなければ空。
    「,」とタブは、Python の 2 バイトの文字コードのどれでも 2 バイト目にならない。「;」は Johab でだけ 2 バイト目になる。
    切り替えの印で字の読み方が変わる文字コード（ISO-2022・HZ・UTF-7）は、バイトの並びから区切りを探せないので、
    エージェントが選んだ「,」・タブ・「;」のほかの区切りは断る。"""
    from ..datafile.csvfile import SEPS, DelimiterError, delim_ja

    name = codecs.lookup(enc).name
    if name in ("ascii", "utf-8", "utf-8-sig") or name.startswith(("utf-16", "utf-32")):
        return frozenset()
    if name.startswith("iso2022") or name in ("hz", "utf-7"):
        if delim in SEPS:
            return frozenset()
        raise DelimiterError(f"区切り {delim_ja(delim)} は、このファイルの文字コード（{enc}）では、範囲で分けて読むときに選べません"
                             "（切り替えの印で字の読み方が変わるので、バイトの並びから区切りを決められません）")

    def one(b: bytes) -> bool:
        try:
            return len(b.decode(enc)) == 1
        except UnicodeDecodeError:
            return False

    d = ord(delim)
    if not any(one(bytes([lead, d])) for lead in range(0x80, 0x100)):
        return frozenset()
    # 1 バイトでは字にならず、次のバイトと合わせて字になるバイト
    return frozenset(lead for lead in range(0x80, 0x100)
                     if not one(bytes([lead])) and any(one(bytes([lead, t])) for t in range(0x30, 0x100)))


def _second_byte(src: Source, lo: int, at: int, leads: frozenset) -> bool:
    """at のバイトが、2 バイトの字の 2 バイト目か（lo は字の境目）。すぐ前に続く「1 バイト目になるバイト」の数が奇数なら
    2 バイト目：1 バイト目にならないバイトは字の終わりなので、その次から 2 バイトずつ組になる。"""
    n = 0
    j = at
    while j > lo:
        blk = src.read(max(lo, j - 256), j)
        k = len(blk)
        while k and blk[k - 1] in leads:
            k -= 1
        n += len(blk) - k
        if k or not blk:
            break
        j -= len(blk)
    return n & 1 == 1


def index_csv(src: Source, start: int, enc: str, delim: str) -> tuple[TextIndex, CsvStats]:
    """RFC 4180 の読み方（datafile.csvfile.parse と同じ）で記録の境目を決め、書き方の事実を集める。メモリは一定（引用符の中が
    どれだけ長くても、枠を動かしながら閉じる引用符を探す）。"""
    from ..limits import progress

    size = src.size
    idx = TextIndex(size)
    st = CsvStats()
    if size <= start:
        st.quote_all = False
        return idx, st
    unit = unit_of(enc)
    toks = tokens(enc, delim, '"', "\r", "\n")
    d_b, q_b, cr_b, lf_b = toks
    field_end = re.compile(b"|".join(re.escape(x) for x in (d_b, cr_b, lf_b)))
    quote = re.compile(re.escape(q_b))
    win = Window(src, start, unit)
    leads = second_byte_leads(enc, delim)  # エージェントが選んだ区切りが、字の 2 バイト目にもなる文字コード

    def end_of_field(j: int):
        """j（字の境目）から先の、マスの終わり（区切り・改行）。字の 2 バイト目の、区切りと同じバイトは飛ばす。"""
        at = j
        while True:
            hit = win.search(field_end, at)
            if hit is None or not leads or hit[2] != d_b or not _second_byte(src, j, hit[0], leads):
                return hit
            at = hit[1]
    pos = start
    nrec = 0
    tick = 0

    def note_field(filled: bool, quoted: bool) -> None:
        if filled:
            st.any_filled = True
            if not quoted:
                st.quote_all = False
        else:
            st.any_empty = True
            if not quoted:
                st.all_empty_quoted = False

    while pos < size:
        nrec += 1
        idx.add(pos)
        rec0 = pos
        nf = 0
        tick += 1
        if tick & 0xFFFF == 0:
            progress("行の位置を調べています", pos, size)
        if unit == 1 and not leads:
            # 速い道：引用符のない行（行の中の区切りを数えるだけ）
            buf_at = pos - win.base
            buf = win.buf
            if not (0 <= buf_at < len(buf)):
                win.base = pos
                win.buf = buf = src.read(pos, pos + CHUNK)
                buf_at = 0
            m = _NL.search(buf, buf_at)
            if m is not None and m.end() < len(buf):
                q = buf.find(q_b, buf_at, m.start())
                if q < 0:
                    n = buf.count(d_b, buf_at, m.start()) + 1
                    if n > st.width:
                        st.width = n
                    if m.start() == buf_at:
                        st.empty_lines += 1
                        if len(st.first_empty_lines) < 20:
                            st.first_empty_lines.append(nrec)
                        st.any_empty = True
                        st.all_empty_quoted = False
                    elif st.quote_all or st.all_empty_quoted:
                        # 字のあるマスがあれば、引用符で囲んでいないマス：quote_all は崩れる
                        if m.start() - buf_at > n - 1:
                            st.any_filled = True
                            st.quote_all = False
                        if buf.find(d_b + d_b, buf_at, m.start()) >= 0 or buf[buf_at:buf_at + len(d_b)] == d_b or \
                                buf[m.start() - len(d_b):m.start()] == d_b:
                            st.any_empty = True
                            st.all_empty_quoted = False
                    t = m.group()
                    idx.terms[t.decode()] += 1
                    k = m.start() - buf_at
                    if k > idx.longest:
                        idx.longest = k
                    pos = win.base + m.end()
                    idx.ends_with_newline = pos == size
                    continue
        # 一般の道：マスを順にたどる
        while True:
            nf += 1
            head = src.read(pos, pos + unit) if not (win.base <= pos < win.base + len(win.buf) - unit) else \
                win.buf[pos - win.base:pos - win.base + unit]
            if head == q_b:
                st.quoted_any = True
                j = pos + unit
                closed = False
                value_any = False
                while True:
                    hit = win.search(quote, j)
                    if hit is None:
                        break
                    if hit[0] > j:
                        value_any = True
                    nxt = src.read(hit[1], hit[1] + unit)
                    if nxt == q_b:
                        value_any = True
                        j = hit[1] + unit
                        continue
                    j = hit[1]
                    closed = True
                    break
                if not closed:
                    st.problem(f"{nrec} 行目：閉じていない引用符があります（ファイルの終わりまでを1つのマスとして読みました）")
                    note_field(True, True)
                    pos = size
                    t = None
                    break
                hit = end_of_field(j)
                m_end = hit[0] if hit else size
                if m_end > j:
                    st.problem(f"{nrec} 行目：引用符を閉じたあとに字があります")
                    value_any = True
                note_field(value_any, True)
            else:
                hit = end_of_field(pos)
                m_end = hit[0] if hit else size
                note_field(m_end > pos, False)
            if hit is None:
                pos = size
                t = None
                break
            if hit[2] == d_b:
                pos = hit[1]
                if pos >= size:  # 区切りで終わる：空のマスが 1 つ続く
                    nf += 1
                    note_field(False, False)
                    t = None
                    break
                continue
            # 改行
            if hit[2] == cr_b and src.read(hit[1], hit[1] + unit) == lf_b:
                idx.terms["\r\n"] += 1
                pos = hit[1] + unit
            else:
                idx.terms[hit[2].decode(enc)] += 1
                pos = hit[1]
            t = 1
            break
        if nf > st.width:
            st.width = nf
        if nf == 1 and pos - rec0 <= 2 * unit and t == 1:
            st.empty_lines += 1
            if len(st.first_empty_lines) < 20:
                st.first_empty_lines.append(nrec)
        idx.longest = max(idx.longest, pos - rec0)
        idx.ends_with_newline = t == 1 and pos == size
    if not st.any_filled:
        st.quote_all = False
    return idx, st


def span_hash(src: Source, a: int, b: int) -> str:
    h = hashlib.sha256()
    for c in src.chunks(a, b):
        h.update(c)
    return h.hexdigest()
