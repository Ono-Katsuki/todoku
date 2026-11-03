"""ZIP（ファイルをまとめたもの）：中のファイルを「a.zip!/中の名前」で指して、ほかの形式と同じに読む・書く。

- 窓：中のファイルの一覧（名前・形式・大きさ）と、開くときの指定（path に「a.zip!/名前」）。
- 中のファイル：取り出して、その形式の窓・入力欄・書き込み・検査・段階をそのまま使う（ZIP の中の ZIP も、続けて「!/」で指す）。
- 書き込み：中のファイルに書き、ZIP は書いたファイルの部品だけを入れ替える（ほかの部品は、圧縮したバイトのまま写す：zipraw）。
- 名前の文字コード：ZIP の決まり（APPNOTE 4.4.4）では、UTF-8 の印がなければ CP437。Info-ZIP の Unicode の名前（0x7075）が
  あればそれを使う。どちらもなく、名前のバイトが UTF-8 として正しければ UTF-8、Shift_JIS（CP932）として正しければ CP932 と読み、
  読み方を一覧に書く（ファイルには書かれていないので、読み方は推定）。
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass
from pathlib import PurePosixPath

from .zipraw import RawZip, ZipFormatError

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
    named: str  # 名前の読み方（utf-8 の印・Unicode の名前・推定）


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


def _display_name(e) -> tuple[str, str]:
    flags = struct.unpack_from("<H", e.central, 8)[0]
    if flags & 0x800:
        return e.name, "UTF-8（ZIP の印）"
    u = _unicode_extra(e.central)
    if u is not None:
        return u, "UTF-8（Unicode の名前）"
    raw = e.name.encode("cp437")
    if raw.isascii():
        return raw.decode("ascii"), "ASCII"
    for enc, how in (("utf-8", "UTF-8（推定）"), ("cp932", "Shift_JIS（推定）")):
        try:
            return raw.decode(enc), how
        except UnicodeDecodeError:
            continue
    return e.name, "CP437（ZIP の決まり）"


def members(blob: bytes) -> list[Member]:
    try:
        z = RawZip(blob)
    except (ZipFormatError, struct.error, IndexError) as e:
        raise ArchiveError(f"ZIP として読めません（{e}）") from None
    out = []
    seen: set[str] = set()
    for e in z._entries_cd_order:
        name, how = _display_name(e)
        if name in seen:  # 読み方を変えて同じ名前になったもの：ZIP の中の名前のままで指す
            name, how = e.name, "CP437（ZIP の決まり）"
        seen.add(name)
        csize, usize = struct.unpack_from("<II", e.central, 20)
        out.append(Member(name, e.name, usize, csize, bool(e.flags & 0x1), name.endswith("/"), how))
    return out


def files(blob: bytes) -> list[Member]:
    """ファイル（フォルダの印を除く）。"""
    return [m for m in members(blob) if not m.is_dir]


def _find(blob: bytes, name: str) -> Member:
    for m in members(blob):
        if m.name == name and not m.is_dir:
            return m
    names = "、".join(m.name for m in files(blob)[:30])
    raise ArchiveError(f"ZIP の中に「{name}」がありません。中のファイル：{names}")


def extract(blob: bytes, name: str) -> bytes:
    m = _find(blob, name)
    if m.encrypted:
        raise ArchiveError(f"「{name}」は ZIP のパスワードで暗号化されていて、取り出せません")
    if m.size > MAX_MEMBER:
        raise ArchiveError(f"「{name}」は展開すると {m.size:,} バイトで、取り出せる大きさ（{MAX_MEMBER:,} バイト）を超えます")
    if m.csize and m.size / m.csize > MAX_RATIO:
        raise ArchiveError(f"「{name}」は圧縮率が {m.size / m.csize:,.0f} 倍で、展開して大きくなる ZIP として取り出しません")
    try:
        return RawZip(blob).read(m.raw)
    except (ZipFormatError, zlib.error) as e:
        raise ArchiveError(f"「{name}」を取り出せません（{e}）") from None


def replace(blob: bytes, name: str, content: bytes) -> bytes:
    """中のファイル name を content にした ZIP（ほかの部品は、圧縮したバイトのまま）。"""
    m = _find(blob, name)
    z = RawZip(blob)
    z.replace(m.raw, content)
    return z.build()


def unchanged_except(old: bytes, new: bytes, name: str) -> list[str]:
    """new が old と、name の部品のほかは同じか（同じなら空。違えば、その説明）。"""
    a, b = RawZip(old), RawZip(new)
    m = _find(old, name)
    problems = []
    if a.names() != b.names():
        problems.append("ZIP の中のファイルの一覧（順）が変わっています")
    for ea in a.entries:
        if ea.name == m.raw:
            continue
        try:
            eb = b.get(ea.name)
        except KeyError:
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


def window(blob: bytes, name: str, kind_ja: dict) -> tuple[dict, list[tuple[str, str]]]:
    """窓：中のファイルの一覧（名前・形式・大きさ・名前の読み方）と、開くときの指定。"""
    from .book import kind_of

    rows = []
    for m in files(blob):
        kind, why = "", ""
        if m.encrypted:
            why = "ZIP のパスワードで暗号化されています（取り出せません）"
        else:
            try:
                kind = kind_of(extract(blob, m.name), base_name(m.name))
            except ArchiveError as e:
                why = str(e)
            except Exception as e:  # noqa: BLE001
                why = f"形式を判別できません（{type(e).__name__}）"
        ok = kind in kind_ja
        rows.append({"name": m.name, "size": m.size, "kind": kind if ok else "", "kind_ja": kind_ja.get(kind, ""),
                     "readable": ok, "why": why or ("" if ok else "todoku が扱えない形式です"), "named": m.named})
    lines = [f"# ZIP：{name}", "",
             f"中のファイル {len(rows)} 個。中のファイルは、path に「{name}{SEP}中の名前」と指定して、ほかの書類と同じに開けます"
             "（書き込むと、ZIP はそのファイルの部品だけを入れ替えて書き出します）。", ""]
    guessed = [r for r in rows if "推定" in r["named"]]
    if guessed:
        lines += [f"名前の文字コードがファイルに書かれていないので、{len(guessed)} 個の名前は、バイトの並びから読み方を推定しました"
                  f"（{'・'.join(sorted({r['named'] for r in guessed}))}）。", ""]
    lines += ["| 中の名前 | 形式 | 大きさ（展開後） |", "|---|---|---|"]
    for r in rows:
        kind = r["kind_ja"] or f"扱えない（{r['why']}）"
        lines.append(f"| {r['name'].replace('|', '｜')} | {kind} | {r['size']:,} バイト |")
    doc = {"kind": "zip", "name": name, "members": rows}
    return doc, [("一覧", "\n".join(lines))]
