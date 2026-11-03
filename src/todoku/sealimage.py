"""本人の印影・署名の画像（PNG・JPEG）を読む：書類に置くための、大きさ・色・透け（アルファ）。

本人が話すことも手を動かすこともできないとき、紙に押印・自署はできない。本人がはっきり同意したときだけ、
本人の印影（はんこを押した跡）や署名の画像を、書類の決めた所に置く。ここは、その画像を外部のライブラリなしに開く部分。

- PNG（ISO/IEC 15948）：濃さ・RGB・パレット・濃さ＋アルファ・RGBA。ビットの深さ 1〜8・16（16 は上位のバイト）。
  インターレースの PNG は開かない（はっきり断る）。tRNS（透けの指定）もアルファにする。
- JPEG（ITU-T T.81）：開かずにそのまま埋める（PDF の DCTDecode）。大きさと成分の数は SOF から読む。
  ベースライン・拡張の順次（SOF0・SOF1）の、濃さ（1成分）か YCbCr（3成分）だけ。プログレッシブ・CMYK は断る
  （置いた後、view_page で描いて確かめられないため）。
"""

from __future__ import annotations

import hashlib
import struct
import zlib
from dataclasses import dataclass
from pathlib import Path

MAX_PIXELS = 4_000_000  # 画素の数の上限（開く時間のため。印影・署名の画像は、ふつうこれよりずっと小さい）

# 値の書き方の名前：{"seal": 画像のパス} は印影、{"signature": 画像のパス} は署名
KINDS = {"seal": "印影", "signature": "署名"}

CONSENT_MESSAGE = ("本人の{what}の画像を置くには、本人の同意が要ります。どの書類の、どこに、どの画像を置くかを本人に伝えて確かめてから、"
                   "\"consent\": true を付けてもう一度呼んでください（同意は、本人がその場で示したものだけです）")


class SealError(ValueError):
    pass


@dataclass
class SealImage:
    kind: str  # png / jpeg
    data: bytes  # ファイルのバイト列
    w: int
    h: int
    channels: int  # 1（濃さ）か 3（RGB・YCbCr）
    samples: bytes | None = None  # PNG：開いた色の値（8 ビット、1行ずつ、w*h*channels バイト）
    alpha: bytes | None = None  # PNG：透け（w*h バイト。255＝透けない）。全部 255 なら None

    @property
    def opaque(self) -> bool:
        return self.alpha is None

    @property
    def sha256(self) -> str:
        return hashlib.sha256(self.data).hexdigest()

    @property
    def ext(self) -> str:
        return "jpeg" if self.kind == "jpeg" else "png"

    @property
    def mime(self) -> str:
        return "image/jpeg" if self.kind == "jpeg" else "image/png"


def seal_key(spec: dict) -> str | None:
    """値の指定が画像を置くもの（seal・signature）なら、その名前。どちらでもなければ None。両方あれば断る。"""
    if not isinstance(spec, dict):
        return None
    keys = [k for k in KINDS if k in spec]
    if len(keys) > 1:
        raise SealError("seal（印影）と signature（署名）は、どちらか1つだけを指定してください")
    return keys[0] if keys else None


def check_consent(spec: dict, key: str) -> None:
    """本人の同意（"consent": true）がなければ断る。真偽の true だけを同意とみなす（"yes" などの字は受け付けない）。"""
    if spec.get("consent") is not True:
        raise SealError(CONSENT_MESSAGE.format(what=KINDS[key]))


def load(spec: dict, key: str) -> SealImage:
    """指定の画像を読む（同意を確かめてから）。"""
    check_consent(spec, key)
    path = spec.get(key)
    if not isinstance(path, str) or not path.strip():
        raise SealError(f"{key} には、本人の{KINDS[key]}の画像（PNG か JPEG）のファイルのパスを指定してください")
    p = Path(path).expanduser()
    if not p.is_file():
        raise SealError(f"{KINDS[key]}の画像が見つかりません：{path}")
    return decode(p.read_bytes())


def decode(data: bytes) -> SealImage:
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return _png(data)
    if data[:2] == b"\xff\xd8":
        return _jpeg(data)
    raise SealError("画像は PNG か JPEG にしてください（このファイルは、どちらでもありません）")


# ---------------------------------------------------------------------- JPEG
def _jpeg(data: bytes) -> SealImage:
    pos = 2
    while pos + 4 <= len(data):
        if data[pos] != 0xFF:
            raise SealError("JPEG の区切り（マーカー）が読めません")
        m = data[pos + 1]
        if m == 0xFF:
            pos += 1
            continue
        if m in (0xD8, 0x01) or 0xD0 <= m <= 0xD7:
            pos += 2
            continue
        if m == 0xD9 or m == 0xDA:
            break
        ln = struct.unpack_from(">H", data, pos + 2)[0]
        seg = data[pos + 4:pos + 2 + ln]
        if m in (0xC0, 0xC1):
            if len(seg) < 6:
                raise SealError("JPEG のフレームの情報が短すぎます")
            p, h, w, nf = seg[0], *struct.unpack_from(">HH", seg, 1), seg[5]
            if p != 8:
                raise SealError(f"JPEG：{p} ビットの精度の画像は置けません（8 ビットの JPEG か PNG にしてください）")
            if nf not in (1, 3):
                raise SealError(f"JPEG：成分が {nf} つの画像（CMYK など）は置けません（RGB か濃さの JPEG、または PNG にしてください）")
            if not (w and h):
                raise SealError("JPEG：画像の大きさが 0 です")
            return SealImage("jpeg", data, w, h, nf)
        if 0xC2 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
            what = "プログレッシブ" if m in (0xC2, 0xC6, 0xCA) else f"SOF{m - 0xC0} の符号化"
            raise SealError(f"JPEG：{what}の画像は置けません（置いた後に描いて確かめられないため）。"
                            "ベースラインの JPEG か PNG にしてください")
        pos += 2 + ln
    raise SealError("JPEG：画像の大きさ（SOF）が見つかりません")


# ---------------------------------------------------------------------- PNG
_CHANNELS = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}


def _png(data: bytes) -> SealImage:
    pos = 8
    ihdr = None
    plte = b""
    trns = None
    idat = bytearray()
    while pos + 8 <= len(data):
        ln, typ = struct.unpack_from(">I4s", data, pos)
        body = data[pos + 8:pos + 8 + ln]
        if len(body) != ln or pos + 12 + ln > len(data):
            raise SealError("PNG：途中で切れています")
        crc = struct.unpack_from(">I", data, pos + 8 + ln)[0]
        if zlib.crc32(typ + body) & 0xFFFFFFFF != crc:
            raise SealError(f"PNG：{typ.decode('latin-1')} の部分の検査の値（CRC）が合いません（壊れています）")
        if typ == b"IHDR":
            ihdr = struct.unpack(">IIBBBBB", body)
        elif typ == b"PLTE":
            plte = body
        elif typ == b"tRNS":
            trns = body
        elif typ == b"IDAT":
            idat += body
        elif typ == b"IEND":
            break
        pos += 12 + ln
    if ihdr is None:
        raise SealError("PNG：画像の情報（IHDR）がありません")
    w, h, depth, ctype, comp, filt, inter = ihdr
    if ctype not in _CHANNELS or comp != 0 or filt != 0:
        raise SealError("PNG：この色の種類・圧縮の方式は、PNG の仕様にありません")
    if inter != 0:
        raise SealError("PNG：インターレース（飛び越し）の PNG は置けません。インターレースなしの PNG にしてください")
    if depth not in (1, 2, 4, 8, 16) or (ctype in (2, 4, 6) and depth < 8) or (ctype == 3 and depth == 16):
        raise SealError(f"PNG：この色の種類で {depth} ビットの深さは、PNG の仕様にありません")
    if not (w and h):
        raise SealError("PNG：画像の大きさが 0 です")
    if w * h > MAX_PIXELS:
        raise SealError(f"画像が大きすぎます（{w}×{h} 画素）。{MAX_PIXELS:,} 画素以下に縮めた画像にしてください")
    try:
        raw = zlib.decompress(bytes(idat))
    except zlib.error as e:
        raise SealError(f"PNG：画像のデータを開けません（{e}）") from None
    nc = _CHANNELS[ctype]
    bits = nc * depth
    stride = (w * bits + 7) // 8
    bpp = max(1, bits // 8)
    if len(raw) < (stride + 1) * h:
        raise SealError("PNG：画像のデータが足りません")
    rows = _unfilter(raw, stride, h, bpp)
    # 1画素ずつの値（8 ビット）に
    if depth == 16:
        vals = [r[0::2] for r in rows]
    elif depth == 8:
        vals = rows
    else:
        maxv = (1 << depth) - 1
        per = 8 // depth
        vals = []
        for r in rows:
            out = bytearray(w)
            for x in range(w):
                v = (r[x // per] >> (8 - depth * (x % per + 1))) & maxv
                out[x] = v if ctype == 3 else v * 255 // maxv
            vals.append(bytes(out))
    samples = bytearray()
    alpha = bytearray()
    if ctype == 3:
        if not plte or len(plte) % 3:
            raise SealError("PNG：パレットの画像に、パレット（PLTE）がありません")
        n = len(plte) // 3
        ta = (trns or b"") + b"\xff" * (n - len(trns or b""))
        for r in vals:
            for v in r[:w]:
                if v >= n:
                    raise SealError("PNG：パレットにない色の番号があります")
                samples += plte[v * 3:v * 3 + 3]
                alpha.append(ta[v])
        channels = 3
    elif ctype in (0, 2):
        channels = nc
        key = None
        if trns is not None and ctype == 0 and len(trns) >= 2:
            key = bytes([_key8(trns[0:2], depth)])
        elif trns is not None and ctype == 2 and len(trns) >= 6:
            key = bytes(_key8(trns[i:i + 2], depth) for i in (0, 2, 4))
        for r in vals:
            samples += r[:w * nc]
            if key is not None:
                for x in range(w):
                    alpha.append(0 if r[x * nc:(x + 1) * nc] == key else 255)
    else:  # 4：濃さ＋アルファ、6：RGBA
        channels = nc - 1
        for r in vals:
            r = r[:w * nc]
            if channels == 1:
                samples += r[0::2]
                alpha += r[1::2]
            else:
                px = bytearray(w * 3)
                px[0::3], px[1::3], px[2::3] = r[0::4], r[1::4], r[2::4]
                samples += px
                alpha += r[3::4]
    a = bytes(alpha) if alpha and alpha.count(255) != len(alpha) else None
    return SealImage("png", data, w, h, channels, bytes(samples), a)


def _key8(two: bytes, depth: int) -> int:
    """tRNS の色（2 バイト）を、開いた値（8 ビット）と同じ物差しに。"""
    v = struct.unpack(">H", two)[0]
    if depth == 16:
        return v >> 8
    if depth == 8:
        return v & 0xFF
    maxv = (1 << depth) - 1
    return (v & maxv) * 255 // maxv


def _unfilter(raw: bytes, stride: int, h: int, bpp: int) -> list[bytes]:
    """PNG の行ごとのフィルター（0 なし・1 Sub・2 Up・3 Average・4 Paeth）を戻す。"""
    rows = []
    prev = bytearray(stride)
    pos = 0
    for _ in range(h):
        ft = raw[pos]
        line = bytearray(raw[pos + 1:pos + 1 + stride])
        pos += 1 + stride
        if ft == 1:
            for i in range(bpp, stride):
                line[i] = (line[i] + line[i - bpp]) & 0xFF
        elif ft == 2:
            line = bytearray((a + b) & 0xFF for a, b in zip(line, prev))
        elif ft == 3:
            for i in range(stride):
                left = line[i - bpp] if i >= bpp else 0
                line[i] = (line[i] + ((left + prev[i]) >> 1)) & 0xFF
        elif ft == 4:
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                b = prev[i]
                c = prev[i - bpp] if i >= bpp else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 0xFF
        elif ft != 0:
            raise SealError(f"PNG：行のフィルターの種類（{ft}）が、PNG の仕様にありません")
        rows.append(bytes(line))
        prev = line
    return rows


def caution(what: str) -> str:
    """本人に伝えること：画像の印影・署名が受け付けられないことがある（決めるのは様式の決まり。todoku は言葉で決めない）。"""
    return (f"置いたのは本人の{what}の画像です。実印の押印や自署（手書きの署名）が求められる所では、画像では受け付けられないことがあります。"
            "様式の注意書き（押印・署名の決まり）を確かめ、分からなければ提出先に確かめてください")
