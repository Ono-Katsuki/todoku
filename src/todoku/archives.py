"""中にファイルが入っている形式の入口：ZIP のほか、tar・gzip・LZH・7z と、メール（.eml・.msg。添付ファイルが中のファイル）。

中のファイルの指定は ZIP と同じ「a.tar!/中の名前」。
- 読む：tar（ustar・GNU・pax。標準の tarfile で位置を得る）・.tar.gz・.tgz（tar として扱う）・.gz（中のファイルは 1 つ）・
  LZH（legacy/lzh.py）・7z（legacy/sevenzip.py）・メール（legacy/mail.py）。RAR は解かない（legacy/convert_first.py：別の道具で展開する手順を返す）。
  7z のパスワードは use_password で置く（呼び出しごとに api が置き直す。結果にも記録にも残さない）。
- 書く：ZIP・tar・gz は、中のファイルに書いて入れ直す。
  - tar：書いた項目の見出しの大きさと和、データと埋め草だけが変わる。ほかの項目のバイトはそのまま（unchanged_except で確かめる）
  - gz：見出し（時刻・名前など）はそのまま。データは縮め直すので、圧縮したバイトは変わるが、解いた中身が書いたものとバイト一致すること
    （api が取り出し直して確かめる）。複数の部分をつないだ gz には書かない（どの部分かを証明できない）
  - LZH・7z とメールには書かない
"""

from __future__ import annotations

import contextvars
import hashlib
import io
import struct
import tarfile
import time
import zlib
from dataclasses import dataclass

from . import ziparc
from .limits import LimitExceeded
from .unlock import NeedsPassword, WrongPassword
from .ziparc import MAX_MEMBER, ArchiveError

KINDS = frozenset({"zip", "tar", "gz", "bz2", "xz", "lzh", "7z", "eml", "msg"})
PASSWORDED = frozenset({"7z"})
WRITABLE = frozenset({"zip", "tar", "gz"})
MAIL = frozenset({"eml", "msg"})
KIND_JA = {"tar": "tar（ファイルをまとめたもの）", "gz": "gzip（縮めた 1 つのファイル）", "bz2": "bzip2（縮めた 1 つのファイル。読むだけ）",
           "xz": "xz（縮めた 1 つのファイル。読むだけ）", "lzh": "LZH（.lzh・.lha。読むだけ）", "7z": "7-Zip の書庫（.7z。読むだけ）",
           "eml": "メール（.eml。読むだけ）", "msg": "Outlook のメール（.msg。読むだけ）"}
NO_WRITE = {"bz2": "bzip2 には書きません（とどくは縮め直しません）。取り出したファイルを別の名前で書き出します。path に中のファイルを指定して、その形式で読んでください",
            "xz": "xz には書きません（とどくは縮め直しません）。取り出したファイルを別の名前で書き出します。path に中のファイルを指定して、その形式で読んでください",
            "lzh": "LZH の書庫には書きません（とどくは LZH を縮め直しません）。取り出したファイルを別の名前で書き出します。path に中のファイルを指定して、その形式で読んでください",
            "7z": "7z の書庫には書きません（とどくは 7z を縮め直しません）。取り出したファイルを別の名前で書き出します。path に中のファイルを指定して、その形式で読んでください",
            "eml": "メールには書きません", "msg": "メールには書きません"}


class _ArcNeeds(NeedsPassword, ArchiveError):
    """書庫を開くのにパスワードが要る（api がパスワードを聞く形に直す。窓の一覧では、そのファイルの理由になる）。"""


class _ArcWrong(WrongPassword, ArchiveError):
    """書庫のパスワードが違う。"""


@dataclass
class AMember:
    name: str
    size: int
    csize: int
    encrypted: bool = False
    is_dir: bool = False
    named: str = "UTF-8"
    why: str = ""
    raw: object = None


# ====================================================================== 見分ける
def tar_header_ok(b: bytes) -> bool:
    """tar の見出しの和（POSIX.1：和の欄を空白として、512 バイトの符号なしの和）か。"""
    if len(b) < 512 or b[:512] == bytes(512):
        return False
    try:
        chk = int(b[148:156].split(b"\0")[0].strip() or b"-1", 8)
    except ValueError:
        return False
    return chk == sum(b[:148]) + 8 * 32 + sum(b[156:512])


def is_gzip(b: bytes) -> bool:
    return b[:3] == b"\x1f\x8b\x08"


def _peek_gz(b: bytes, n: int) -> bytes:
    try:
        d = zlib.decompressobj(31)
        return d.decompress(b, n)
    except zlib.error:
        return b""


def _packed_kind(b: bytes) -> str | None:
    """bzip2（"BZh" と 1〜9、ブロックの印）・xz（6 バイトの印）か。"""
    if b[:3] == b"BZh" and b[3:4] in tuple(b"%d" % i for i in range(1, 10)) and b[4:10] in (b"1AY&SY", b"\x17rE8P\x90"):
        return "bz2"
    return "xz" if b[:6] == b"\xfd7zXZ\0" else None


def _new_decompressor(k: str):
    import bz2
    import lzma

    return bz2.BZ2Decompressor() if k == "bz2" else lzma.LZMADecompressor(format=lzma.FORMAT_XZ)


def _peek_packed(b: bytes, k: str, n: int) -> bytes:
    import lzma

    try:
        return _new_decompressor(k).decompress(b, n)
    except (OSError, EOFError, lzma.LZMAError, ValueError):
        return b""


def unpack(b: bytes, k: str, what: str = "") -> bytes:
    """bzip2・xz を解く（つないだストリームも）。大きさは展開の予算と MAX_MEMBER まで。読むだけ。"""
    import lzma

    from .limits import budget

    what = what or KIND_JA.get(k, k)
    bud = budget()
    cap = min(bud.room(None), MAX_MEMBER)
    out = bytearray()
    src = b
    try:
        while True:
            d = _new_decompressor(k)
            out += d.decompress(src, min(1 << 22, cap - len(out) + 1))
            while not d.eof and len(out) <= cap:
                if d.needs_input:
                    raise ArchiveError(f"{what}：データが途中で終わっています")
                out += d.decompress(b"", min(1 << 22, cap - len(out) + 1))
            if len(out) > cap:
                raise bud.exceeded(cap, None, what)
            src = d.unused_data.lstrip(b"\0") if k == "xz" else d.unused_data
            if not src:
                break
            if _packed_kind(src) != k:
                raise ArchiveError(f"{what}：終わりの後ろに、読めないバイトが {len(src):,} バイトあります")
    except (ArchiveError, LimitExceeded):
        raise
    except (OSError, EOFError, lzma.LZMAError, ValueError) as e:
        raise ArchiveError(f"{what}：解けません（{e}）") from None
    bud.charge(len(out), what)
    return bytes(out)


def sniff(blob: bytes, filename: str = "") -> str | None:
    """中身で決まる入れ物の形式（なければ None）。"""
    if is_gzip(blob):
        return "tar" if tar_header_ok(_peek_gz(blob, 512)) else "gz"
    packed = _packed_kind(blob)
    if packed:
        return "tar" if tar_header_ok(_peek_packed(blob, packed, 512)) else packed
    if tar_header_ok(blob[:512]):
        return "tar"
    if blob[:2] == b"MZ" and not filename.lower().endswith((".lzh", ".lha", ".rar", ".7z", ".exe")):
        return None
    from .legacy.lzh import start_of

    if start_of(blob[:1 << 20] if blob[:2] == b"MZ" else blob[:4096]) is not None:
        return "lzh"
    if blob[:2] == b"MZ":  # 自己解凍の .exe の中の 7z（実行せずに、中の書庫を読む）
        from .legacy import sevenzip

        if sevenzip.find_start(blob) is not None:
            return "7z"
    return None


# ====================================================================== gz
def _gz_header(b: bytes) -> tuple[int, str | None]:
    """(見出しの大きさ, FNAME の名前)。RFC 1952。"""
    if not is_gzip(b) or len(b) < 10:
        raise ArchiveError("gzip の見出しが壊れています")
    flg = b[3]
    p = 10
    name = None
    try:
        if flg & 4:
            xlen = struct.unpack_from("<H", b, p)[0]
            p += 2 + xlen
        if flg & 8:
            e = b.index(0, p)
            raw = b[p:e]
            name = _decode_name(raw)
            p = e + 1
        if flg & 16:
            p = b.index(0, p) + 1
        if flg & 2:
            p += 2
    except (ValueError, struct.error):
        raise ArchiveError("gzip の見出しが壊れています") from None
    return p, name


def _decode_name(raw: bytes) -> str:
    for enc in ("utf-8", "cp932", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")


def gunzip(b: bytes, what: str = "gzip") -> tuple[bytes, bool]:
    """(解いた中身, 複数の部分をつないだ gz か)。CRC と大きさを確かめる。上限は limits（展開したバイトで数える）。"""
    out = bytearray()
    pos = 0
    n = 0
    try:
        while pos < len(b):
            if b[pos:] .strip(b"\0") == b"":
                break  # 終わりの埋め草
            hs, _nm = _gz_header(b[pos:])
            d = zlib.decompressobj(-15)
            src = b[pos + hs:]
            part = _inflate_stream(d, src, what, len(out))
            tail = d.unused_data
            if len(tail) < 8:
                raise ArchiveError(f"{what}：データが途中で終わっています")
            crc, isize = struct.unpack_from("<II", tail, 0)
            if zlib.crc32(part) & 0xFFFFFFFF != crc or len(part) & 0xFFFFFFFF != isize:
                raise ArchiveError(f"{what}：解いた結果が CRC と合いません（壊れています）")
            out += part
            pos = len(b) - len(tail) + 8
            n += 1
    except zlib.error as e:
        raise ArchiveError(f"{what}：解けません（{e}）") from None
    return bytes(out), n > 1


def _inflate_stream(d, src: bytes, what: str, done: int) -> bytes:
    from .limits import budget

    b = budget()
    cap = min(b.room(None), MAX_MEMBER - done)
    out = bytearray()
    piece = d.decompress(src, min(1 << 22, cap + 1))
    out += piece
    while d.unconsumed_tail and len(out) <= cap:
        out += d.decompress(d.unconsumed_tail, min(1 << 22, cap - len(out) + 1))
    if len(out) > cap:
        raise b.exceeded(cap, None, what)
    if not d.eof:
        raise ArchiveError(f"{what}：データが途中で終わっています")
    b.charge(len(out), what)
    return bytes(out)


def gzip_again(original: bytes, content: bytes) -> bytes:
    hs, _n = _gz_header(original)
    c = zlib.compressobj(9, zlib.DEFLATED, -15)
    body = c.compress(content) + c.flush()
    return original[:hs] + body + struct.pack("<II", zlib.crc32(content) & 0xFFFFFFFF, len(content) & 0xFFFFFFFF)


# ====================================================================== tar
def _tar_plain(blob: bytes) -> tuple[bytes, bool]:
    if is_gzip(blob):
        data, _multi = gunzip(blob, "tar.gz")
        return data, True
    packed = _packed_kind(blob)
    if packed:
        return unpack(blob, packed, f"tar.{packed}"), True
    return blob, False


def _tar_infos(data: bytes) -> list[tarfile.TarInfo]:
    try:
        with tarfile.open(fileobj=io.BytesIO(data), mode="r:", encoding="utf-8", errors="surrogateescape") as t:
            return t.getmembers()
    except tarfile.TarError as e:
        raise ArchiveError(f"tar として読めません（{e}）") from None


def _tar_name(ti: tarfile.TarInfo) -> tuple[str, str]:
    s = ti.name
    try:
        s.encode("utf-8")
        return s, "ASCII" if s.isascii() else "UTF-8"
    except UnicodeEncodeError:
        raw = s.encode("utf-8", "surrogateescape")
        try:
            return raw.decode("cp932"), "Shift_JIS（推定）"
        except UnicodeDecodeError:
            return raw.decode("latin-1"), "Latin-1（どの文字コードでも読めません）"


def _tar_members(data: bytes) -> list[AMember]:
    out = []
    for ti in _tar_infos(data):
        if not ti.isreg():
            continue
        nm, how = _tar_name(ti)
        out.append(AMember(nm, ti.size, ti.size, named=how, raw=ti))
    return out


def tar_replace(data: bytes, ti: tarfile.TarInfo, content: bytes) -> bytes:
    if ti.sparse is not None or "size" in (ti.pax_headers or {}) or ti.offset_data - ti.offset != 512:
        raise ArchiveError("大きなファイルの拡張見出し・穴のある項目には書けません（どの大きさが正しいかを証明できません）")
    h = bytearray(data[ti.offset:ti.offset + 512])
    if len(content) >= 8 ** 11:
        raise ArchiveError("大きすぎて tar の見出しに書けません")
    h[124:136] = b"%011o\0" % len(content)
    h[148:156] = b" " * 8
    h[148:156] = b"%06o\0 " % sum(h)
    old_end = ti.offset_data + ((ti.size + 511) // 512) * 512
    pad = (512 - len(content) % 512) % 512
    return data[:ti.offset] + bytes(h) + content + bytes(pad) + data[old_end:]


# ====================================================================== 7z
_PW: contextvars.ContextVar[str | None] = contextvars.ContextVar("todoku_archive_password", default=None)
_held: dict = {}  # いちばん新しく開いた書庫：{"blob": 中身, "key": ハッシュ, "arc": 読んだ書庫, "folders": {番号: 解いたバイト}}
_HOLD_MAX = 1 << 30  # 解いたフォルダを控える大きさの合計
SLOW_SECONDS = 45  # 窓・段階で、中のファイルの形式を確かめるのに使う時間（呼び出しの時間の上限の前に、残りを path で指してもらう）


SIBLING: contextvars.ContextVar = contextvars.ContextVar("todoku_sibling", default=None)  # 隣のファイルを読む関数（dBase のメモ）


def use_sibling(reader) -> None:
    """reader(語幹, 拡張子の並び) → (バイト, 名前) か None。ファイルの隣か、書庫の同じフォルダから読む。"""
    SIBLING.set(reader)


def current_password() -> str | None:
    return _PW.get()


def use_password(password: str | None) -> None:
    """この呼び出しのパスワード。パスワードのある書庫の控えは、呼び出しをまたいで残さない。"""
    if password is not None or _PW.get() is not None:
        _held.clear()
    _PW.set(password)


def _need_password(e) -> Exception:
    return _ArcWrong(f"{e}") if getattr(e, "wrong", False) else _ArcNeeds(f"{e}")


def _opened(blob: bytes, k: str):
    if _held.get("blob") is blob:
        return _held["arc"]
    key = (k, hashlib.sha1(blob).digest())
    if _held.get("key") == key:
        _held["blob"] = blob
        return _held["arc"]
    arc = _parse(blob, k)
    _held.clear()
    _held.update({"blob": blob, "key": key, "arc": arc, "folders": {}})
    return arc


def _parse(blob: bytes, k: str):
    from .legacy import sevenzip

    try:
        return sevenzip.read_archive(blob, _PW.get())
    except sevenzip.EncryptedError as e:
        raise _need_password(e) from None
    except sevenzip.SevenZipError as e:
        raise ArchiveError(str(e)) from None


def _packed_members(blob: bytes, k: str) -> list[AMember]:
    arc = _opened(blob, k)
    return [AMember(e.name, e.size, e.csize, encrypted=e.encrypted, named="書庫に書かれた名前（UTF-16）", raw=e) for e in arc.entries if not e.is_dir]


def _packed_extract(blob: bytes, m: AMember, k: str) -> bytes:
    from .legacy import sevenzip

    arc = _opened(blob, k)
    e = m.raw
    if e.folder is None:
        return b""
    folders = _held["folders"]
    try:
        data = folders.get(e.folder)
        if data is None:
            data = arc.folder_bytes(e.folder, _PW.get())
            if sum(len(v) for v in folders.values()) + len(data) <= _HOLD_MAX:
                folders[e.folder] = data
    except sevenzip.EncryptedError as ex:
        raise _need_password(ex) from None
    except sevenzip.SevenZipError as ex:
        raise ArchiveError(str(ex)) from None
    out = data[e.offset:e.offset + e.size]
    if e.crc is not None and zlib.crc32(out) & 0xFFFFFFFF != e.crc:
        if e.encrypted:
            raise _ArcWrong("パスワードが違います（または書庫が壊れています）")
        raise ArchiveError(f"「{e.name}」を解いた結果が CRC と合いません（書庫が壊れています）")
    return out


def preflight(blob: bytes, k: str) -> None:
    """開く前に、パスワードを確かめる：見出しの暗号（要る・違う）と、パスワードを渡されたときは、いちばん小さな暗号のファイルを解いてみる（違う）。"""
    if k == "mdb":
        from .legacy.jet import check_password

        err = check_password(blob, _PW.get())
        if err == "need":
            raise _ArcNeeds("このデータベースはパスワードで保護されています")
        if err == "wrong":
            raise _ArcWrong("パスワードが違います")
        return
    if k not in PASSWORDED:
        return
    ms = _packed_members(blob, k)
    if _PW.get() is None:
        return
    locked = [m for m in ms if m.encrypted]
    if locked:
        _packed_extract(blob, min(locked, key=lambda m: m.size), k)


# ====================================================================== 入口
def _mail(blob: bytes, k: str):
    from .legacy.mail import read_eml, read_msg

    return read_eml(blob) if k == "eml" else read_msg(blob)


def _name_of_gz(blob: bytes, filename: str) -> str:
    _h, nm = _gz_header(blob)
    if nm:
        return nm.replace("\\", "/").rsplit("/", 1)[-1] or "中身"
    base = filename.rsplit("/", 1)[-1]
    return base[:-3] if base.lower().endswith(".gz") and len(base) > 3 else (base + ".out" if base else "中身")


def files(blob: bytes, k: str, filename: str = "") -> list:
    if k == "zip":
        return ziparc.files(blob)
    if k == "tar":
        data, _gz = _tar_plain(blob)
        return _tar_members(data)
    if k in ("bz2", "xz"):
        base = filename.rsplit("/", 1)[-1]
        ext = f".{k}"
        nm = base[:-len(ext)] if base.lower().endswith(ext) and len(base) > len(ext) else (base + ".out" if base else "中身")
        return [AMember(nm, len(unpack(blob, k)), len(blob), named="ファイルの名前から")]
    if k == "gz":
        nm = _name_of_gz(blob, filename)
        isize = struct.unpack_from("<I", blob, len(blob) - 4)[0] if len(blob) >= 18 else 0
        return [AMember(nm, isize, len(blob), named="gzip の FNAME か、ファイルの名前から")]
    if k in PASSWORDED:
        return _packed_members(blob, k)
    if k == "lzh":
        from .legacy.lzh import LzhError, entries

        try:
            es = entries(blob)
        except LzhError as e:
            raise ArchiveError(str(e)) from None
        return [AMember(e.name, e.size, e.csize, named=e.named, raw=e) for e in es if not e.is_dir]
    if k in MAIL:
        from .legacy.mail import MailError

        try:
            m = _mail(blob, k)
        except MailError as e:
            raise ArchiveError(str(e)) from None
        return [AMember(n, len(b), len(b), named="メールの添付の名前", raw=b) for n, b in m.attachments]
    raise ArchiveError(f"この形式（{k}）は中のファイルを持ちません")


def _find(ms: list, name: str, k: str):
    for m in ms:
        if m.name == name:
            return m
    raise ArchiveError(f"{KIND_JA.get(k, k)}の中に「{name}」がありません。中のファイル：{'、'.join(m.name for m in ms[:30])}")


def extract(blob: bytes, name: str, k: str, filename: str = "") -> bytes:
    if k == "zip":
        return ziparc.extract(blob, name)
    if k == "tar":
        data, _gz = _tar_plain(blob)
        ti = _find(_tar_members(data), name, k).raw
        return data[ti.offset_data:ti.offset_data + ti.size]
    if k == "gz":
        _find(files(blob, k, filename), name, k)
        return gunzip(blob)[0]
    if k in ("bz2", "xz"):
        _find(files(blob, k, filename), name, k)
        return unpack(blob, k)
    m = _find(files(blob, k, filename), name, k)
    if k in PASSWORDED:
        return _packed_extract(blob, m, k)
    if k == "lzh":
        from .legacy.lzh import LzhError, extract as lx

        try:
            return lx(blob, m.raw)
        except LzhError as e:
            raise ArchiveError(str(e)) from None
    return m.raw


def replace(blob: bytes, name: str, content: bytes, k: str, filename: str = "") -> bytes:
    if k == "zip":
        return ziparc.replace(blob, name, content)
    if k == "tar":
        if _packed_kind(blob):
            raise ArchiveError(f"tar.{_packed_kind(blob)} には書きません（とどくは {_packed_kind(blob)} に縮め直しません）。"
                               "中のファイルを別の名前で書き出してください（.tar.gz・.tar なら、入れ直して書けます）")
        data, wrapped = _tar_plain(blob)
        if wrapped and gunzip(blob)[1]:
            raise ArchiveError("複数の部分をつないだ gz には書きません")
        new = tar_replace(data, _find(_tar_members(data), name, k).raw, content)
        return gzip_again(blob, new) if wrapped else new
    if k == "gz":
        if gunzip(blob)[1]:
            raise ArchiveError("複数の部分をつないだ gz には書きません（どの部分に書いたかを証明できません）")
        return gzip_again(blob, content)
    raise ArchiveError(NO_WRITE.get(k, "この形式には書けません"))


def unchanged_except(old: bytes, new: bytes, name: str, k: str, filename: str = "") -> list[str]:
    if k == "zip":
        return ziparc.unchanged_except(old, new, name)
    if k == "gz":
        ho, _n = _gz_header(old)
        hn, _n2 = _gz_header(new)
        return [] if old[:ho] == new[:hn] else ["gzip の見出しが変わっています"]
    if k == "tar":
        a, wa = _tar_plain(old)
        b, wb = _tar_plain(new)
        problems = []
        if wa:
            ho, _n = _gz_header(old)
            hn, _n2 = _gz_header(new)
            if old[:ho] != new[:hn]:
                problems.append("gzip の見出しが変わっています")
        ma, mb = _tar_infos(a), _tar_infos(b)
        if [x.name for x in ma] != [x.name for x in mb]:
            return problems + ["tar の中のファイルの一覧（順）が変わっています"]
        tgt = _find(_tar_members(a), name, k).raw.offset
        ends = [x.offset for x in ma[1:]] + [None]
        endsb = [x.offset for x in mb[1:]] + [None]
        for x, y, ex, ey in zip(ma, mb, ends, endsb):
            if x.offset == tgt:
                h1, h2 = bytearray(a[x.offset:x.offset + 512]), bytearray(b[y.offset:y.offset + 512])
                h1[124:136] = h2[124:136] = bytes(12)
                h1[148:156] = h2[148:156] = bytes(8)
                if h1 != h2:
                    problems.append(f"「{x.name}」の見出しの、大きさと和のほかが変わっています")
                continue
            ax = a[x.offset:ex] if ex is not None else a[x.offset:]
            by = b[y.offset:ey] if ey is not None else b[y.offset:]
            if ax != by:
                problems.append(f"「{x.name}」のバイトが変わっています")
        last_a = ma[-1]
        last_b = mb[-1]
        tail_a = a[last_a.offset_data + ((last_a.size + 511) // 512) * 512:]
        tail_b = b[last_b.offset_data + ((last_b.size + 511) // 512) * 512:]
        if tail_a != tail_b:
            problems.append("tar の終わりの埋め草が変わっています")
        return problems
    return [NO_WRITE.get(k, "この形式には書けません")]


def window(blob: bytes, name: str, kind_ja: dict, k: str) -> tuple[dict, list[tuple[str, str]]]:
    if k == "zip":
        return ziparc.window(blob, name, kind_ja)
    from .book import kind_of

    head_parts: list[tuple[str, str]] = []
    body_lines: list[str] = []
    notes: list[str] = []
    if k in MAIL:
        from .legacy.mail import MailError, lines_of

        try:
            m = _mail(blob, k)
        except MailError as e:
            raise ArchiveError(str(e)) from None
        body_lines = lines_of(m)
        notes = list(m.notes)
    ms = files(blob, k, name)
    rows = []
    spent = False
    started = time.monotonic()
    late = False
    for mem in ms:
        kind, why = "", ""
        if k in PASSWORDED and not late and time.monotonic() - started > SLOW_SECONDS:
            late = True  # Python で解く書庫：呼び出しの時間の上限に当たる前に、残りは形式を確かめない
        if spent:
            why = "1 回に展開できる大きさの合計を超えるので、形式を確かめていません（path で指せば開けます）"
        elif late:
            why = "解くのに時間がかかるので、形式を確かめていません（path で指せば開けます）"
        else:
            try:
                kind = kind_of(extract(blob, mem.name, k, name), ziparc.base_name(mem.name))
            except ArchiveError as e:
                why = str(e)
            except LimitExceeded as e:
                why = str(e)
                spent = e.total
            except Exception as e:  # noqa: BLE001
                why = f"形式を判別できません（{type(e).__name__}）"
        ok = kind in kind_ja
        rows.append({"name": mem.name, "size": mem.size, "kind": kind if ok else "", "kind_ja": kind_ja.get(kind, ""),
                     "readable": ok, "why": why or ("" if ok else "todoku が扱えない形式です"), "named": mem.named})
    ja = KIND_JA.get(k, k)
    lines = [f"# {ja}：{name}", ""]
    if k in MAIL:
        lines += [f"> {x}" for x in notes] + ([""] if notes else [])
        lines += body_lines + ["", "## 添付ファイル", ""]
        lines.append(f"添付 {len(rows)} 個。" + (f"path に「{name}{ziparc.SEP}名前」と指定して、ほかの書類と同じに開けます。" if rows else ""))
    else:
        how = "書き込むと、そのファイルだけを入れ替えて書き出します" if k in WRITABLE else "この書庫には書き込みません（読むだけ）"
        lines.append(f"中のファイル {len(rows)} 個。中のファイルは、path に「{name}{ziparc.SEP}中の名前」と指定して、ほかの書類と同じに開けます（{how}）。")
    lines.append("")
    locked = sum(1 for m in ms if getattr(m, "encrypted", False))
    if locked and k in PASSWORDED:
        lines += [f"{locked} 個のファイルはパスワードで保護されています。本人にパスワードを聞いて、password に入れて開いてください。", ""]
    guessed = [r for r in rows if "推定" in r["named"]]
    if guessed:
        lines += [f"名前の文字コードがファイルに書かれていないので、{len(guessed)} 個の名前は、バイトの並びから読み方を推定しました"
                  f"（{'・'.join(sorted({r['named'] for r in guessed}))}）。", ""]
    if rows:
        lines += ["| 中の名前 | 形式 | 大きさ（展開後） |", "|---|---|---|"]
        for r in rows:
            kind = r["kind_ja"] or f"扱えない（{r['why']}）"
            lines.append(f"| {r['name'].replace('|', '｜')} | {kind} | {r['size']:,} バイト |")
    doc = {"kind": k, "name": name, "members": rows, "fields": []}
    if k == "lzh":
        from .legacy.convert_first import needs_tool
        from .legacy.lzh import needs_other_tool

        if any(needs_other_tool(m.raw) for m in ms):
            doc["needs_tool"] = needs_tool("lzh", name)
    if k in MAIL:
        doc["body_lines"] = body_lines
        doc["notes"] = notes
    return doc, head_parts + [("一覧", "\n".join(lines) + "\n")]


def mail_body(blob: bytes, k: str) -> str:
    """検査用：本文の字。"""
    m = _mail(blob, k)
    return m.body + "".join(v for _k, v in m.headers)

