"""窓を作らずに、ファイルの字を順に読む（read_window の raw）：窓を作るのが上限に当たる大きな書類でも、中身に届くための最後の道。

- ZIP の書類（Word・PowerPoint・OpenDocument・HWPX・Excel）：XML の部品を順に展開して、タグを除いた字を出す（メモリは一定）
- 文字のファイル（JSON・RTF・HTML・1 つの XML の OpenDocument）：字のまま（文字コードは textstream で決める）
入力欄は見つけない（窓ではない）。続きは cursor。どの部品のどこを読んでいるかを示す。
"""

from __future__ import annotations

import codecs
import re
import zipfile

from .base import Page, RangeError, comma
from .textstream import Source, detect

_TAG = re.compile(r"<[^>]*>")
_BLOCK_END = re.compile(r"</(?:\w+:)?(?:p|tr|row|si|text-p|h)\s*>")


def _xml_names(z: zipfile.ZipFile) -> list[str]:
    return [i.filename for i in z.infolist() if i.filename.lower().endswith((".xml", ".rels")) is True and not i.filename.endswith(".rels")
            and not i.filename.startswith(("_rels/", "[Content_Types]"))]


def _plain(xml_piece: str) -> str:
    s = _BLOCK_END.sub("\n", xml_piece)
    s = _TAG.sub("", s)
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def read_raw(path: str, cursor: str | None, max_chars: int, encoding: str | None = None) -> Page:
    with open(path, "rb") as f:
        head = f.read(2)
    return _read_zip(path, cursor, max_chars) if head == b"PK" else _read_text(path, cursor, max_chars, encoding)


def _read_zip(path: str, cursor: str | None, max_chars: int) -> Page:
    from .. import limits
    from .textstream import CHUNK  # noqa: F401
    from .xlsxdoc import Member

    part, skip = 0, 0
    if cursor:
        m = re.fullmatch(r"raw:(\d+):(\d+)", str(cursor))
        if not m:
            raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません")
        part, skip = int(m[1]), int(m[2])
    src = Source(path)
    try:
        with zipfile.ZipFile(path) as z:
            names = _xml_names(z)
            out: list[str] = []
            size = 0
            start = (part, skip)
            while part < len(names):
                m_ = Member(src, z.getinfo(names[part]))
                dec = codecs.getincrementaldecoder("utf-8")("replace")
                seen = 0  # この部品で、ここまでに出した字
                carry = ""
                piece_all: list[str] = []
                done = False
                for c in m_.chunks():
                    carry += dec.decode(c)
                    k = carry.rfind(">")
                    if k < 0:
                        continue
                    text = _plain(carry[:k + 1])
                    carry = carry[k + 1:]
                    if seen + len(text) <= skip:
                        seen += len(text)
                        continue
                    text = text[max(0, skip - seen):] if skip > seen else text
                    seen = max(seen, skip)
                    room = max_chars - size
                    if len(text) > room:
                        piece_all.append(text[:room])
                        seen += room
                        size += room
                        done = True
                        break
                    piece_all.append(text)
                    size += len(text)
                    seen += len(text)
                    limits.progress(f"{names[part]} の字を読んでいます")
                got = "".join(piece_all)
                if got.strip():
                    out.append((f"〔部品 {names[part]}〕\n" if not skip else "") + got)
                if done:
                    return Page("".join(out), f"部品 {start[0] + 1}〜{part + 1}／全 {len(names)} 部品（窓ではなく字のまま）",
                                {"raw": True, "cursor": f"raw:{part}:{seen}"}, f"部品 {comma(len(names) - part)} つの残り", out)
                part += 1
                skip = 0
            return Page("".join(out), f"部品 {start[0] + 1}〜{len(names)}／全 {len(names)} 部品（窓ではなく字のまま）", None, "", out)
    finally:
        src.close()


def _read_text(path: str, cursor: str | None, max_chars: int, encoding: str | None) -> Page:
    src = Source(path)
    try:
        dec_info = detect(src, encoding)
        at = len(dec_info.bom)
        chars = 0
        if cursor:
            m = re.fullmatch(r"raw:(\d+):(\d+)", str(cursor))
            if not m:
                raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません")
            at, chars = int(m[1]), int(m[2])
        dec = codecs.getincrementaldecoder(dec_info.encoding)("strict")
        out: list[str] = []
        size = 0
        pos = at
        for c in src.chunks(at, size=1 << 16):
            # 字の境目で止められるよう、少しずつ。読みすぎた分は捨てて、次は続きの位置から読み直す
            t = dec.decode(c)
            pos += len(c)
            if size + len(t) >= max_chars:
                keep = max_chars - size
                out.append(t[:keep])
                size += keep
                rest = t[keep:]
                pending = dec.getstate()[0]
                back = len(rest.encode(dec_info.encoding)) + len(pending)
                return Page("".join(out), f"{chars + 1:,}〜{chars + size:,} 字（窓ではなく字のまま）", {"raw": True, "cursor": f"raw:{pos - back}:{chars + size}"},
                            f"{comma(src.size - (pos - back))} バイト", ["".join(out)])
            out.append(t)
            size += len(t)
        out.append(dec.decode(b"", True))
        text = "".join(out)
        return Page(text, f"{chars + 1:,}〜{chars + len(text):,} 字（窓ではなく字のまま）", None, "", [text])
    finally:
        src.close()
