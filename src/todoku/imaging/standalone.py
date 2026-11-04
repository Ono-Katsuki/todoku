"""画像そのもののファイルを、書類として開く（kind "image"）：窓・ページ（フレーム）の画像・段階・検査・PDF に包む。

エージェントは自分の目で画像を読むので、とどくはページ（フレーム）の画素をそのまま渡す。中身を字にはしない（推測しない）。
書き込みはできない：画像の中の字の位置に書くことを証明する方法がないので、to_pdf で PDF に包み、PDF の枠（mm）に書いて、証明はその PDF で行う。
"""

from __future__ import annotations

import base64
import struct
import time
import zlib

from .formats import FORMAT_JA, Frame, FrameInfo, ImageError, decode, info

KIND_JA = "画像のファイル"
DEFAULT_SIDE = 1600  # ページ全体を見るときの長い辺の画素
MAX_SIDE = 4000
MAX_PIXELS = 12_000_000
A4 = (210.0, 297.0)
DECODE_SECONDS = 60.0  # 開くのにかける時間の上限（JPEG 2000・JBIG2 はそこまでに開けた所を返す。働き手の上限 TODOKU_TIME_LIMIT より短く）


def _frames(blob: bytes) -> tuple[str, list[FrameInfo]]:
    return info(blob)


def _frame_line(i: int, f: FrameInfo) -> str:
    if f.dpi:
        return f"{i} 枚目：{f.w}×{f.h} 画素（{f.dpi[0]:.0f} dpi・紙の上で約 {f.w / f.dpi[0] * 25.4:.0f}×{f.h / f.dpi[1] * 25.4:.0f} mm）"
    return f"{i} 枚目：{f.w}×{f.h} 画素（解像度の記録なし）"


def build(blob: bytes, name: str) -> dict:
    fmt, frames = _frames(blob)
    return {"file": name, "kind": "image", "format": fmt, "format_ja": FORMAT_JA.get(fmt, fmt),
            "pages": [{"number": i, "width_px": f.w, "height_px": f.h, "dpi": list(f.dpi) if f.dpi else None,
                       "size_mm": [round(f.w / f.dpi[0] * 25.4, 1), round(f.h / f.dpi[1] * 25.4, 1)] if f.dpi else None}
                      for i, f in enumerate(frames, 1)],
            "fields": [], "statement": "中身は画像の中にあり、字の情報はありません" if fmt != "svg" else "ベクターの絵です。text の要素の字は窓にあります",
            "texts": _svg_texts(blob) if fmt == "svg" else []}


def _svg_texts(blob: bytes) -> list[str]:
    from .svg import texts

    return texts(blob)[:500]


def markdown_parts(win: dict) -> list[tuple[str, str]]:
    n = len(win["pages"])
    fr = [FrameInfo(p["width_px"], p["height_px"], tuple(p["dpi"]) if p["dpi"] else None) for p in win["pages"]]
    lines = [f"# {win['file']}（{win['format_ja']}の画像）", "",
             f"このファイルは画像です（{n} ページ）。**中身は画像の中にあり、字の情報はありません**。"
             "view_page（page＝何枚目か）で画像を見て、エージェントの目で読みます。読めない所は推測で埋めず、本人に伝えてください。", ""]
    if win["format"] == "svg":
        lines[2] = (f"このファイルはベクターの絵（SVG）です（{n} ページ）。view_page で絵として見ます（小さな字は region で拡大すると、なめらかなまま細かく描きます）。"
                    "読めない所は推測で埋めず、本人に伝えてください。")
        if win.get("texts"):
            lines += ["絵の中の字（text の要素。図の中の位置は view_page で確かめてください）：", ""] + [f"- {t}" for t in win["texts"]] + [""]
    lines += [f"- {_frame_line(i, f)}" for i, f in enumerate(fr, 1)]
    lines += ["", "画像の上には直接書けません。書く必要があれば、to_pdf で PDF に包み、できた PDF に fill_document で枠（mm）を指定して書きます"
              "（書いた証明は、その PDF について行います）。"]
    return [("画像", "\n".join(lines))]


def summary(doc: dict) -> dict:
    return {"format": doc["format"], "pages": len(doc["pages"])}


# ====================================================================== ページ（フレーム）の画像
def view_frame(blob: bytes, page: int, region: list | None = None, marks: list | None = None, dpi: float | None = None,
               deadline: float | None = None) -> dict:
    """view_page と同じ形の結果（text・png_base64・width・height・not_drawn）。解像度が分かれば mm の目盛り、なければ画素の目盛り。"""
    from ..pdf.pageimage import with_ruler
    from ..pdf.render.image import area_down

    fmt, frames = _frames(blob)
    if not 1 <= page <= len(frames):
        raise ValueError(f"page は 1〜{len(frames)} です")
    fi = frames[page - 1]
    fw, fh = fi.w, fi.h
    if not fw or not fh:
        m = decode(blob, page - 1)
        fi = FrameInfo(m.w, m.h, fi.dpi)
        fw, fh = fi.w, fi.h
    has_dpi = bool(fi.dpi)
    unit_px = (fi.dpi[0] / 25.4) if has_dpi else 1.0  # 目盛りの 1 単位が、元の画素でいくつか
    unit_py = (fi.dpi[1] / 25.4) if has_dpi else 1.0
    pw, ph = fw / unit_px, fh / unit_py  # 目盛りの単位でのページの大きさ
    if region is not None:
        if not (isinstance(region, list) and len(region) == 4):
            raise ValueError("region は [左, 上, 右, 下] で指定してください")
        x0, y0, x1, y1 = (float(v) for v in region)
        x0, x1 = sorted((x0, x1))
        y0, y1 = sorted((y0, y1))
        x0, y0, x1, y1 = max(0.0, x0), max(0.0, y0), min(pw, x1), min(ph, y1)
        if x1 - x0 < 0.5 or y1 - y0 < 0.5:
            raise ValueError(f"範囲が小さすぎるか、画像（{pw:.0f}×{ph:.0f}）の外です")
    else:
        x0, y0, x1, y1 = 0.0, 0.0, pw, ph
    # 出す画素の倍率（出す画素 / 元の画素）
    cw, ch = (x1 - x0) * unit_px, (y1 - y0) * unit_py  # 範囲の元の画素
    if dpi:
        if has_dpi:
            k = dpi / fi.dpi[0]
        else:
            raise ValueError("この画像には解像度の記録がないので、dpi では指定できません（省くか、region で拡大してください）")
    else:
        k = min(DEFAULT_SIDE / max(cw, ch), 4.0) if region is not None else min(1.0, DEFAULT_SIDE / max(cw, ch))
    k = min(k, MAX_SIDE / max(cw, ch), (MAX_PIXELS / (cw * ch)) ** 0.5)
    # 開く：画面に要る画素の数（全体の画像の大きさで）
    target = int(fw * fh * min(1.0, k * k))
    if deadline is None:
        deadline = time.monotonic() + DECODE_SECONDS
    want = None
    if region is not None:  # 範囲だけを開く（分けて開ける形式。予算は開く分に掛かるので、どの部分でも元の解像度まで開ける）
        want = (int(x0 * unit_px), int(y0 * unit_py), min(fw, int(x1 * unit_px + 0.999)), min(fh, int(y1 * unit_py + 0.999)))
    fr = decode(blob, page - 1, target, deadline, want)
    wx0, wy0, wx1, wy1 = fr.window or (0, 0, fw, fh)
    scale_x, scale_y = fr.w / (wx1 - wx0), fr.h / (wy1 - wy0)  # 開いた画像 / 元の画素
    sx0, sy0 = max(0, int((x0 * unit_px - wx0) * scale_x)), max(0, int((y0 * unit_py - wy0) * scale_y))
    sx1, sy1 = min(fr.w, int((x1 * unit_px - wx0) * scale_x + 0.999)), min(fr.h, int((y1 * unit_py - wy0) * scale_y + 0.999))
    crop_w, crop_h = max(1, sx1 - sx0), max(1, sy1 - sy0)
    rows = [fr.rgb[(y * fr.w + sx0) * 3:(y * fr.w + sx0 + crop_w) * 3] for y in range(sy0, sy0 + crop_h)]
    crop = b"".join(rows)
    W = max(1, round(cw * k))
    s = W / (x1 - x0)  # 目盛りの 1 単位の出す画素（縦横の解像度が違う画像は、紙の上の形にして見せる：目盛りが縦横とも合う）
    H = max(1, round((y1 - y0) * s))
    if H > MAX_SIDE:
        s *= MAX_SIDE / H
        W, H = max(1, round((x1 - x0) * s)), MAX_SIDE
    out = crop
    cw2, ch2 = crop_w, crop_h
    if W < cw2 or H < ch2:  # 縮める向きは面積の平均で（細い線を落とさない）
        cw2, ch2 = min(W, cw2), min(H, ch2)
        out = area_down(out, crop_w, crop_h, 3, cw2, ch2)
    if (cw2, ch2) != (W, H):
        out = _nearest(out, cw2, ch2, W, H)
    r = {"rgb": out, "w": W, "h": H, "scale": s, "origin_mm": (x0, y0)}
    ticks = None if has_dpi else _pixel_ticks(s)
    png, w, h, notes = with_ruler(r, marks, ticks=ticks)
    if has_dpi:
        unit = "mm"
        head = (f"{page} 枚目（{FORMAT_JA.get(fmt, fmt)}・{pw:.0f}×{ph:.0f} mm・{fi.dpi[0]:.0f} dpi）の、左 {x0:.0f}〜{x1:.0f} mm・上 {y0:.0f}〜{y1:.0f} mm の範囲"
                f"（1 mm＝{s:.1f} 画素）。上と左の帯の数は、画像の左上からの mm です。")
    else:
        unit = "画素"
        head = (f"{page} 枚目（{FORMAT_JA.get(fmt, fmt)}・{fw}×{fh} 画素。解像度の記録がないので、紙の上の mm は分かりません）の、"
                f"左 {x0:.0f}〜{x1:.0f}・上 {y0:.0f}〜{y1:.0f} 画素の範囲。上と左の帯の数は、画像の左上からの画素です。")
    inside = "中身は絵の中にあります（text の要素の字は窓にもあります）" if fmt == "svg" else "中身は画像の中にあります（字の情報はありません）"
    text = (head + f"{inside}。小さな字は region（[左, 上, 右, 下]、単位は{unit}）で拡大して読みます。"
            "読めない所は推測で埋めず、本人に伝えてください。書くには to_pdf で PDF に包みます。")
    if len(frames) > 1:
        text += f" 全部で {len(frames)} 枚です。"
    if scale_x > 0.999:
        text += f" この範囲は元の解像度（{fw}×{fh} 画素のまま）で開きました。"
    else:
        text += (f" この範囲は元の {fw}×{fh} 画素の約 {1 / scale_x:.1f} 分の 1 の解像度で開きました。"
                 "細かく見るには region で範囲を狭めてください（どの部分でも、狭めれば元の解像度まで開けます）。")
    if fr.notes:
        text += " " + " ".join(fr.notes)
    if fr.problems:
        text += " 描けなかった所：" + "／".join(fr.problems[:5])
    if notes:
        text += " " + " ".join(notes)
    return {"text": text, "png_base64": base64.b64encode(png).decode("ascii"), "width": w, "height": h, "not_drawn": list(fr.problems)[:8]}


def _pixel_ticks(s: float) -> tuple[int, int]:
    """画素の目盛り：(短い線の間隔, 数を書く間隔)。s：1 画素の出す画素。"""
    for major in (100, 200, 500, 1000, 2000, 5000):
        if major * s >= 50:
            return major // 5 if major % 5 == 0 else major // 2, major
    return 2000, 10000


def _nearest(px: bytes, w: int, h: int, W: int, H: int) -> bytes:
    cols = [min(w - 1, x * w // W) for x in range(W)]
    rows = []
    cache: dict = {}
    for y in range(H):
        sy = min(h - 1, y * h // H)
        r = cache.get(sy)
        if r is None:
            base = sy * w * 3
            line = bytearray(W * 3)
            src = px[base:base + w * 3]
            line[0::3] = bytes([src[c * 3] for c in cols])
            line[1::3] = bytes([src[c * 3 + 1] for c in cols])
            line[2::3] = bytes([src[c * 3 + 2] for c in cols])
            r = bytes(line)
            cache = {sy: r}
        rows.append(r)
    return b"".join(rows)


# ====================================================================== 段階・検査
def grade(blob: bytes, name: str):
    from ..grade import Reason, _final, _locked

    try:
        fmt, frames = _frames(blob)
    except (ImageError, struct.error, IndexError) as e:
        return _locked("開けない", f"画像のファイルとして読めません（{e}）", "開けないので、書けません")
    ja = FORMAT_JA.get(fmt, fmt)
    where = [f"{i} 枚目" for i in range(1, len(frames) + 1)]
    reasons = [Reason("E", "画像のファイル", f"{ja} の画像（{len(frames)} ページ）で、字の情報がありません。エージェントがページの画像（view_page）を見て読みます。"
                      "読めない所は推測で埋めません", where)]
    write = [Reason("D", "画像のファイルには直接書けない", "画像の画素に書いたことを証明する方法がないので、このファイルには書けません。書く必要があれば to_pdf で PDF に包み、"
                    "できた PDF に枠（mm）で書きます（証明はその PDF について行います）", where)]
    return _final(reasons, write)


def audit(blob: bytes, name: str):
    """検査：すべてのフレームの見出しを読め、開けること。"""
    from ..xlsx.audit import Audit, Check

    _fmt, frames = _frames(blob)
    bad = []
    for i, _f in enumerate(frames):
        try:
            if _fmt == "heic":  # 全部を開くと何分もかかる：はじめのタイルが開けることを確かめる（見出しは全部読んである）
                fr = decode(blob, i, 250_000, time.monotonic() + DECODE_SECONDS, (0, 0, 1, 1))
                bad += [f"{i + 1} 枚目：{p}" for p in fr.problems if not fr.incomplete]
                continue
            fr = decode(blob, i, 250_000)
            bad += [f"{i + 1} 枚目：{p}" for p in fr.problems]
        except Exception as e:  # noqa: BLE001
            bad.append(f"{i + 1} 枚目：{type(e).__name__}：{e}")
    checks = [Check("画像が開ける", not bad, f"{len(frames)} 枚の画像を開きました" if not bad else f"開けない所があります（{len(bad)} か所）", bad[:20]),
              Check("字の情報", True, "画像のファイルなので、字の情報はありません。エージェントが画像を見て読みます", [], agent=True)]
    return Audit(name, checks)


# ====================================================================== PDF に包む
def to_pdf(blob: bytes) -> tuple[bytes, list[dict], list[str]]:
    """画像を PDF に包む：(PDF, ページの一覧, 知らせ)。解像度があれば紙の上の大きさ、なければ A4 に収まる大きさ（仮に、と伝える）。
    JPEG はそのまま入れる（画質を変えない。回す必要のあるものは画素で）。ほかは Flate の RGB か濃さ。"""
    fmt, frames = _frames(blob)
    objs: list[bytes] = [b"", b""]
    kids: list[str] = []
    listing: list[dict] = []
    notes: list[str] = []
    for i, fi in enumerate(frames):
        fw, fh = fi.w, fi.h
        fr = None
        passthrough = None
        if fmt == "jpeg" and _jpeg_upright(blob):
            passthrough = blob
        else:
            fr = decode(blob, i, None, time.monotonic() + DECODE_SECONDS)
            if fr.incomplete:  # 途中の画素を PDF にしない（PDF の見た目が本物と違ってしまう）
                raise ImageError(f"まだ全部を開けていないので、PDF に包めません（{'／'.join(fr.problems[:1])}）")
            fw, fh = fr.w, fr.h
            notes += [f"{i + 1} 枚目：{p}" for p in fr.problems]
        dpi = fi.dpi
        assumed = False
        if dpi:
            mw, mh = fw / dpi[0] * 25.4, fh / dpi[1] * 25.4
        else:
            k = min(A4[0] / fw, A4[1] / fh) if fw <= fh else min(A4[1] / fw, A4[0] / fh)
            mw, mh = fw * k, fh * k
            assumed = True
        pw, ph = mw * 72 / 25.4, mh * 72 / 25.4
        img_no = len(objs) + 3
        if passthrough is not None:
            ncomp = _jpeg_components(blob)
            cs = {1: "/DeviceGray", 3: "/DeviceRGB", 4: "/DeviceCMYK"}.get(ncomp, "/DeviceRGB")
            # Adobe の CMYK の JPEG は反転して保存される（Decode で戻す）
            dec = "/Decode[1 0 1 0 1 0 1 0]" if ncomp == 4 and b"Adobe" in blob[:4096] else ""
            head = f"<</Type/XObject/Subtype/Image/Width {fw}/Height {fh}/BitsPerComponent 8/ColorSpace{cs}{dec}/Filter/DCTDecode/Length {len(passthrough)}>>".encode()
            img = head + b"stream\n" + passthrough + b"\nendstream"
        else:
            gray = fr.rgb[0::3] == fr.rgb[1::3] and fr.rgb[1::3] == fr.rgb[2::3]
            raw = zlib.compress(fr.rgb[0::3] if gray else fr.rgb, 6)
            cs = "/DeviceGray" if gray else "/DeviceRGB"
            head = f"<</Type/XObject/Subtype/Image/Width {fw}/Height {fh}/BitsPerComponent 8/ColorSpace{cs}/Filter/FlateDecode/Length {len(raw)}>>".encode()
            img = head + b"stream\n" + raw + b"\nendstream"
        cont = f"q {pw:.4f} 0 0 {ph:.4f} 0 0 cm /Im0 Do Q".encode()
        cobj = b"<</Length %d>>stream\n" % len(cont) + cont + b"\nendstream"
        page_no = len(objs) + 1
        objs.append(f"<</Type/Page/Parent 2 0 R/MediaBox[0 0 {pw:.4f} {ph:.4f}]/Contents {page_no + 1} 0 R/Resources<</XObject<</Im0 {img_no} 0 R>>>>>>".encode())
        objs.append(cobj)
        objs.append(img)
        kids.append(f"{page_no} 0 R")
        listing.append({"page": i + 1, "size_mm": [round(mw, 1), round(mh, 1)], "dpi_assumed": assumed})
        if assumed:
            notes.append(f"{i + 1} 枚目：解像度の記録がないので、A4 に収まる大きさ（{mw:.0f}×{mh:.0f} mm）で包みました（紙の上の本当の大きさではありません）")
    objs[0] = b"<</Type/Catalog/Pages 2 0 R>>"
    objs[1] = f"<</Type/Pages/Kids[{' '.join(kids)}]/Count {len(kids)}>>".encode()
    out = bytearray(b"%PDF-1.5\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for n, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % n + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        out += b"%010d 00000 n \n" % o
    out += b"trailer<</Size %d/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)
    return bytes(out), listing, notes


def _jpeg_upright(blob: bytes) -> bool:
    """そのまま PDF に入れられる JPEG か：向きが正しく、PDF の DCTDecode が扱える符号化（算術符号・12 ビット・可逆は画素で入れる）。"""
    from .formats import _jpeg_meta, jpeg_plain

    return _jpeg_meta(blob)[3] == 1 and jpeg_plain(blob)


def _jpeg_components(blob: bytes) -> int:
    from .formats import _jpeg_segments

    for m, seg in _jpeg_segments(blob):
        if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
            return seg[5]
    return 3


__all__ = ["KIND_JA", "Frame", "audit", "build", "grade", "markdown_parts", "summary", "to_pdf", "view_frame"]
