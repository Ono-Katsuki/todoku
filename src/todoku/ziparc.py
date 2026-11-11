"""ZIP（ファイルをまとめたもの）：中のファイルを「a.zip!/中の名前」で指して、ほかの形式と同じに読む・書く。

- 窓：中のファイルの一覧（名前・形式・大きさ）と、開くときの指定（path に「a.zip!/名前」）。
- 中のファイル：取り出して、その形式の窓・入力欄・書き込み・検査・段階をそのまま使う（ZIP の中の ZIP も、続けて「!/」で指す）。
- 書き込み：中のファイルに書き、ZIP は書いたファイルの部品だけを入れ替える（ほかの部品は、圧縮したバイトのまま写す：zipraw）。
- 名前の文字コード：UTF-8 の印（APPNOTE 4.4.4）か、Info-ZIP の Unicode の名前（0x7075）があれば、それ。どちらもない名前は、
  ZIP の決まりでは CP437 だが、実際には作った機械の文字コード（Shift_JIS・EUC-KR・GBK・CP850・CP866 など）で書かれていて、
  ファイルからは決まらない。namecode の決め方で読む：ASCII・UTF-8 の形の名前はそのとおり。ほかは、ZIP の中の名前をまとめて、
  読める文字コードをすべて試し、1 つだけなら、それ。2 つ以上なら、決まった順のいちばん前で示して、ほかを候補として窓に並べる
  （エージェントが encoding で選ぶか、正しい読み方の名前で「a.zip!/名前」と指す：名前のバイトをきれいに読める読み方なら、どれでも指せる）。
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass
from pathlib import PurePosixPath

from .zipraw import AmbiguousZip, RawZip, ZipFormatError

SEP = "!/"
MAX_MEMBER = 512 * 1024 * 1024  # 取り出す1つのファイルの大きさの上限（展開後）。これより大きいものは取り出さない
MAX_RATIO = 1000  # 圧縮率の上限（これより大きいものは、展開して大きくなる ZIP として取り出さない）


class ArchiveError(ValueError):
    pass


@dataclass
class Member:
    name: str  # 表示し、path で指す名前
    raw: str  # zipraw の名前（CP437 か UTF-8 で読んだもの）
    size: int  # 展開後の大きさ（中央ディレクトリの値）
    csize: int
    encrypted: bool
    is_dir: bool
    named: str  # 名前の読み方（UTF-8 の印・Unicode の名前・バイトの形・決められない など）
    rawname: bytes | None = None  # 文字コードの印のない名前のバイト（印のある名前は None）


def _unicode_extra(central: bytes) -> str | None:
    """Info-ZIP の Unicode の名前（extra 0x7075：版 1・元の名前の CRC・UTF-8 の名前）。"""
    nlen, xlen = struct.unpack_from("<HH", central, 28)
    raw_name = central[46:46 + nlen]
    x = central[46 + nlen:46 + nlen + xlen]
    i = 0
    while i + 4 <= len(x):
        tag, size = struct.unpack_from("<HH", x, i)
        body = x[i + 4:i + 4 + size]
        if tag == 0x7075 and len(body) >= 5 and body[0] == 1:
            if struct.unpack_from("<I", body, 1)[0] == zlib.crc32(raw_name) & 0xFFFFFFFF:
                try:
                    return body[5:].decode("utf-8")
                except UnicodeDecodeError:
                    return None
        i += 4 + size
    return None


def _labelled_name(e) -> tuple[str, str] | None:
    """文字コードの印のある名前：(名前, 読み方)。印がなければ None（名前のバイトは e.name.encode("cp437")）。"""
    flags = struct.unpack_from("<H", e.central, 8)[0]
    if flags & 0x800:
        return e.name, "UTF-8（ZIP の印）"
    u = _unicode_extra(e.central)
    if u is not None:
        return u, "UTF-8（Unicode の名前）"
    return None


def _open(blob: bytes) -> RawZip:
    try:
        return RawZip(blob)
    except AmbiguousZip as e:
        raise ArchiveError(str(e)) from None
    except (ZipFormatError, struct.error, IndexError) as e:
        raise ArchiveError(f"ZIP として読めません（{e}）") from None


def listing(blob: bytes, z: RawZip | None = None, encoding: str | None = None):
    """(中の項目, 印のない名前の読み方 namecode.Reading)。encoding：エージェントが選んだ、印のない名前の文字コード。"""
    from . import namecode

    z = z or _open(blob)
    known = [_labelled_name(e) for e in z._entries_cd_order]
    raws = [e.name.encode("cp437") if lab is None else None for e, lab in zip(z._entries_cd_order, known)]
    try:
        reading = namecode.decide([r for r in raws if r is not None], encoding, fallback="cp437")
    except namecode.NameCodeError as e:
        raise ArchiveError(str(e)) from None
    out = []
    seen: set[str] = set()
    for e, lab, raw in zip(z._entries_cd_order, known, raws):
        name, how = lab if lab is not None else reading.text(raw)
        if name in seen:  # 読み方を変えて同じ名前になったもの：ZIP の中の名前のままで指す
            name, how = e.name, "CP437（ZIP の決まり）"
        seen.add(name)
        csize, usize = struct.unpack_from("<II", e.central, 20)
        out.append(Member(name, e.name, usize, csize, bool(e.flags & 0x1), name.endswith("/"), how,
                          raw if raw is not None and not raw.isascii() else None))
    return out, reading


def members(blob: bytes, z: RawZip | None = None) -> list[Member]:
    return listing(blob, z)[0]


def files(blob: bytes, z: RawZip | None = None) -> list[Member]:
    """ファイル（フォルダの印を除く）。"""
    return [m for m in members(blob, z) if not m.is_dir]


def find_by_reading(ms: list, name: str, raw_of, fix=None):
    """name が、印のない名前のバイトの、きれいに読める読み方のどれかと同じ項目（1 つだけのとき）。なければ None。
    エージェントが、窓の候補から正しい読み方を選んで、その名前で指したときのため。"""
    from . import namecode

    hits = []
    for m in ms:
        raw = raw_of(m)
        if raw is not None and any((fix(t) if fix else t) == name for t in namecode.readings(raw)):
            hits.append(m)
    if len(hits) > 1:
        raise ArchiveError(f"「{name}」という読み方は、中の {len(hits)} 個のファイルの名前に当たります"
                           f"（{'、'.join(m.name for m in hits[:5])}）。窓の一覧の名前で指してください")
    return hits[0] if hits else None


def _find(blob: bytes, name: str, ms: list[Member] | None = None) -> Member:
    ms = members(blob) if ms is None else ms
    for m in ms:
        if m.name == name and not m.is_dir:
            return m
    hit = find_by_reading([m for m in ms if not m.is_dir], name, lambda m: m.rawname)
    if hit is not None:
        return hit
    names = "、".join([m.name for m in ms if not m.is_dir][:30])
    raise ArchiveError(f"ZIP の中に「{name}」がありません。中のファイル：{names}")


def _limit(m: Member) -> int:
    """取り出してよい大きさ：MAX_MEMBER と、圧縮したバイトの MAX_RATIO 倍（小さな部品は 1 MB まで）の小さい方。
    中央ディレクトリの「展開後の大きさ」は偽れるので、展開して実際に出てきたバイトで比べる。"""
    return min(MAX_MEMBER, max(m.csize * MAX_RATIO, 1 << 20)) if m.csize else MAX_MEMBER


_OTHER_METHODS = (9, 12, 14)  # ZIP の APPNOTE の圧縮の番号：9 Deflate64・12 BZIP2・14 LZMA


def _other_method(method: int, data: bytes, cap: int, name: str) -> bytes:
    """Deflate・無圧縮のほかの圧縮（読むだけ。書くときは、書いた項目を Deflate で入れ直す）。"""
    import bz2
    import lzma
    import struct

    from .legacy.sevenzip import SevenZipError, deflate64

    try:
        if method == 9:
            return deflate64(data, cap)
        if method == 12:
            d = bz2.BZ2Decompressor()
            raw = d.decompress(data, cap + 1)
        else:
            # APPNOTE 5.8.8：版（2 バイト）・性質の大きさ（2 バイト）・LZMA の性質（5 バイト）の後が LZMA の流れ（終わりの印つきのことがある）
            _ver, psize = struct.unpack_from("<HH", data, 0)
            props = data[4:4 + psize]
            filt = lzma._decode_filter_properties(lzma.FILTER_LZMA1, props)
            d = lzma.LZMADecompressor(format=lzma.FORMAT_RAW, filters=[filt])
            raw = d.decompress(data[4 + psize:], cap + 1)
        if len(raw) > cap:
            raise ArchiveError(f"「{name}」は展開すると上限（{cap:,} バイト）を超えるので、取り出しません")
        return raw
    except ArchiveError:
        raise
    except (SevenZipError, OSError, EOFError, lzma.LZMAError, struct.error, ValueError) as x:
        raise ArchiveError(f"「{name}」を取り出せません（{x}）") from None


def _read(z: RawZip, m: Member) -> bytes:
    from .limits import LimitExceeded, inflate

    name = m.name
    if m.encrypted:
        raise ArchiveError(f"「{name}」は ZIP のパスワードで暗号化されていて、取り出せません")
    if m.size > MAX_MEMBER:
        raise ArchiveError(f"「{name}」は展開すると {m.size:,} バイトで、取り出せる大きさ（{MAX_MEMBER:,} バイト）を超えます")
    if m.csize and m.size / m.csize > MAX_RATIO:
        raise ArchiveError(f"「{name}」は圧縮率が {m.size / m.csize:,.0f} 倍で、展開して大きくなる ZIP として取り出しません")
    e = z.get(m.raw)
    try:
        if e.method == 8:
            cap = _limit(m)
            try:
                raw = inflate(e.data, -15, what=f"「{name}」", part=cap)
            except LimitExceeded as x:
                if cap < MAX_MEMBER and not x.total:
                    raise ArchiveError(f"「{name}」は展開すると圧縮したバイトの {MAX_RATIO:,} 倍を超え、展開して大きくなる ZIP として"
                                       "取り出しません") from None
                raise
            if zlib.crc32(raw) & 0xFFFFFFFF != e.crc:
                raise ZipFormatError(f"{name}: CRCが一致しません")
            return raw
        if e.method in _OTHER_METHODS:
            raw = _other_method(e.method, e.data, _limit(m), name)
            if zlib.crc32(raw) & 0xFFFFFFFF != e.crc:
                raise ZipFormatError(f"{name}: CRCが一致しません")
            return raw
        return e.read()
    except (ZipFormatError, zlib.error) as x:
        raise ArchiveError(f"「{name}」を取り出せません（{x}）") from None


def extract(blob: bytes, name: str) -> bytes:
    z = _open(blob)
    return _read(z, _find(blob, name, members(blob, z)))


def replace(blob: bytes, name: str, content: bytes) -> bytes:
    """中のファイル name を content にした ZIP（ほかの部品は、圧縮したバイトのまま）。"""
    z = _open(blob)
    m = _find(blob, name, members(blob, z))
    z.replace(m.raw, content)
    return z.build()


def unchanged_except(old: bytes, new: bytes, name: str) -> list[str]:
    """new が old と、name の部品のほかは同じか（同じなら空。違えば、その説明）。"""
    a, b = _open(old), _open(new)
    m = _find(old, name, members(old, a))
    problems = []
    if a.names() != b.names():
        problems.append("ZIP の中のファイルの一覧（順）が変わっています")
    by_name = {e.name: e for e in b.entries}
    for ea in a.entries:
        if ea.name == m.raw:
            continue
        eb = by_name.get(ea.name)
        if eb is None:
            problems.append(f"「{ea.name}」がなくなっています")
            continue
        if (ea.local, ea.data, ea.descriptor) != (eb.local, eb.data, eb.descriptor) or ea.central[:42] != eb.central[:42] \
                or ea.central[46:] != eb.central[46:]:
            problems.append(f"「{ea.name}」のバイトが変わっています")
    return problems


def split(path: str) -> tuple[str, list[str]] | None:
    """「a.zip!/b.csv」→（a.zip, [b.csv]）。「!/」がなければ None。"""
    if SEP not in path:
        return None
    outer, *inner = path.split(SEP)
    if not inner or any(not x for x in inner):
        return None
    return outer, inner


def base_name(member: str) -> str:
    return PurePosixPath(member).name


CHOOSE_NAME = ("選んだら、その文字コードの名前を encoding で渡すと、一覧をその読み方で出し直します。"
               "中のファイルは、選んだ読み方の名前で「{name}!/名前」と指せます（名前のバイトをきれいに読める読み方なら、どれでも指せます）。")


def window(blob: bytes, name: str, kind_ja: dict, encoding: str | None = None) -> tuple[dict, list[tuple[str, str]]]:
    """窓：中のファイルの一覧（名前・形式・大きさ・名前の読み方）と、開くときの指定。"""
    from .book import kind_of

    from .limits import LimitExceeded

    z = _open(blob)  # 中央ディレクトリは 1 度だけ読む
    ms, reading = listing(blob, z, encoding)
    rows = []
    spent = False  # 1 回の呼び出しで展開できる合計を使い切った：残りの形式は確かめない
    for m in ms:
        if m.is_dir:
            continue
        kind, why = "", ""
        if m.encrypted:
            why = "ZIP のパスワードで暗号化されています（取り出せません）"
        elif spent:
            why = "1 回に展開できる大きさの合計を超えるので、形式を確かめていません（path で指せば開けます）"
        else:
            try:
                kind = kind_of(_read(z, m), base_name(m.name))
            except ArchiveError as e:
                why = str(e)
            except LimitExceeded as e:
                why = str(e)
                spent = e.total
            except Exception as e:  # noqa: BLE001
                why = f"形式を判別できません（{type(e).__name__}）"
        ok = kind in kind_ja
        rows.append({"name": m.name, "size": m.size, "kind": kind if ok else "", "kind_ja": kind_ja.get(kind, ""),
                     "readable": ok, "why": why or ("" if ok else "todoku が扱えない形式です"), "named": m.named})
    lines = [f"# ZIP：{name}", "",
             f"中のファイル {len(rows)} 個。中のファイルは、path に「{name}{SEP}中の名前」と指定して、ほかの書類と同じに開けます"
             "（書き込むと、ZIP はそのファイルの部品だけを入れ替えて書き出します）。", ""]
    notes = reading.note_lines("名前", CHOOSE_NAME.format(name=name))
    if notes:
        lines += notes + [""]
    lines += ["| 中の名前 | 形式 | 大きさ（展開後） |", "|---|---|---|"]
    for r in rows:
        kind = r["kind_ja"] or f"扱えない（{r['why']}）"
        lines.append(f"| {r['name'].replace('|', '｜')} | {kind} | {r['size']:,} バイト |")
    doc = {"kind": "zip", "name": name, "members": rows}
    if reading.info():
        doc["name_encoding"] = reading.info()
    return doc, [("一覧", "\n".join(lines))]
