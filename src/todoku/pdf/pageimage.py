"""PDF のページを、エージェントが見る画像にする（スキャンした紙のページ）。

ページに置かれた画像を、置いた行列（CTM）どおりに、見た目の向き（/Rotate を反映）でページの大きさの画像に描く。
端に、ページの左上からの mm の目盛りを付ける：エージェントは画像を見て、書く所を mm の枠で指定する。
書く予定の枠（marks）を赤で描き込める：書く前に、枠の位置をエージェントが目で確かめるため。
字や線（ベクトル）は描かない（画像だけのページのための道具。字のあるページは窓で読む）。
"""

from __future__ import annotations

import base64
import struct
import zlib

from .content import page_content
from .layout import View
from .raster import Gray, RasterError, image_gray

MM = 72 / 25.4  # 1 mm のポイント


def _inv(m):
    a, b, c, d, e, f = m
    det = a * d - b * c
    if abs(det) < 1e-12:
        return None
    return (d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det)


def render(doc, page_no: int, region_mm: list | None = None, max_side: int = 1600) -> dict:
    """ページ（の範囲）の画像。region_mm：[左, 上, 右, 下]（ページの左上からの mm）。
    返す：{rgb: 画素（RGB）, w, h, scale（1 mm の画素）, origin_mm（画像の左上の mm）, skipped（描けなかった画像）}。"""
    pages = doc.pages()
    if not 1 <= page_no <= len(pages):
        raise ValueError(f"page は 1〜{len(pages)} です")
    page = pages[page_no - 1]
    view = View(doc, page)
    pw, ph = view.width / MM, view.height / MM  # mm
    x0, y0, x1, y1 = region_mm or [0, 0, pw, ph]
    x0, y0, x1, y1 = max(0.0, min(x0, x1)), max(0.0, min(y0, y1)), min(pw, max(x0, x1)), min(ph, max(y0, y1))
    if x1 - x0 < 1 or y1 - y0 < 1:
        raise ValueError(f"範囲が小さすぎるか、ページ（{pw:.0f}×{ph:.0f} mm）の外です")
    s = max_side / max(x1 - x0, y1 - y0)  # 1 mm の画素
    W, H = max(1, int((x1 - x0) * s)), max(1, int((y1 - y0) * s))
    canvas = bytearray(b"\xff" * (W * H))
    pc = page_content(doc, page)
    skipped = []
    for pic in pc.pictures:
        if pic.ctm is None or pic.image is None:
            continue
        inv = _inv(pic.ctm)
        if inv is None:
            continue
        try:
            g = image_gray(doc, pic.image, max_pixels=max(4_000_000, W * H * 2))
        except (RasterError, ValueError, IndexError, KeyError, zlib.error) as e:
            skipped.append(f"{pic.name}：{e}")
            continue
        _draw(canvas, W, H, s, x0, y0, view, inv, g)
    rgb = bytearray(W * H * 3)
    rgb[0::3] = canvas
    rgb[1::3] = canvas
    rgb[2::3] = canvas
    return {"rgb": rgb, "w": W, "h": H, "scale": s, "origin_mm": (x0, y0), "page_mm": (pw, ph), "skipped": skipped,
            "view": view}


def _draw(canvas: bytearray, W: int, H: int, s: float, x0: float, y0: float, view: View, inv, g: Gray) -> None:
    """見た目の画素 (X, Y) → mm → 見た目の pt → ページの座標 → 画像の単位の四角 → 画像の画素。1行ずつ、一次の式で。"""
    a, b, c, d, e, f = inv

    def img_xy(X: float, Y: float) -> tuple[float, float]:
        vx, vy = (x0 + X / s) * MM, (y0 + Y / s) * MM
        px, py = view.back(vx, vy)
        u, v = a * px + c * py + e, b * px + d * py + f
        return u * g.w, (1 - v) * g.h

    gw, gh, src = g.w, g.h, g.px
    for Y in range(H):
        ix0, iy0 = img_xy(0.5, Y + 0.5)
        ix1, iy1 = img_xy(1.5, Y + 0.5)
        dx, dy = ix1 - ix0, iy1 - iy0
        row = canvas[Y * W:(Y + 1) * W]
        changed = False
        for X in range(W):
            ix, iy = int(ix0 + X * dx), int(iy0 + X * dy)
            if 0 <= ix < gw and 0 <= iy < gh:
                v = src[iy * gw + ix]
                if v < row[X]:  # 重ねて描くとき、暗い方（白い紙の上のインク）
                    row[X] = v
                    changed = True
        if changed:
            canvas[Y * W:(Y + 1) * W] = row


# ---------------------------------------------------------------------- 目盛りと枠
_DIGITS = {
    "0": ["111", "101", "101", "101", "111"], "1": ["010", "110", "010", "010", "111"], "2": ["111", "001", "111", "100", "111"],
    "3": ["111", "001", "111", "001", "111"], "4": ["101", "101", "111", "001", "001"], "5": ["111", "100", "111", "001", "111"],
    "6": ["111", "100", "111", "101", "111"], "7": ["111", "001", "001", "001", "001"], "8": ["111", "101", "111", "101", "111"],
    "9": ["111", "101", "111", "001", "111"], "-": ["000", "000", "111", "000", "000"]}


def _text(img: bytearray, W: int, H: int, x: int, y: int, t: str, color=(0, 0, 0), k: int = 2) -> None:
    for ch in t:
        pat = _DIGITS.get(ch)
        if pat:
            for r, line in enumerate(pat):
                for cc, bit in enumerate(line):
                    if bit == "1":
                        _box(img, W, H, x + cc * k, y + r * k, x + cc * k + k, y + r * k + k, color)
        x += 4 * k


def _box(img, W, H, x0, y0, x1, y1, color) -> None:
    x0, y0, x1, y1 = max(0, x0), max(0, y0), min(W, x1), min(H, y1)
    if x1 <= x0 or y1 <= y0:
        return
    line = bytes(color) * (x1 - x0)
    for y in range(y0, y1):
        img[(y * W + x0) * 3:(y * W + x1) * 3] = line


def with_ruler(r: dict, marks: list | None = None, written: list | None = None) -> tuple[bytes, int, int, list[str]]:
    """目盛り（10 mm ごとに数、5 mm ごとに短い線）を上と左の帯に付け、marks（{box:[左,上,右,下] mm, label}）を赤で描く。"""
    M = 26  # 帯の幅（画素）
    W, H, s = r["w"], r["h"], r["scale"]
    ox, oy = r["origin_mm"]
    TW, TH = W + M, H + M
    img = bytearray(b"\xff" * (TW * TH * 3))
    for y in range(H):
        img[((y + M) * TW + M) * 3:((y + M) * TW + M + W) * 3] = r["rgb"][y * W * 3:(y + 1) * W * 3]
    gray = (150, 150, 150)
    step = 10 if s >= 3 else 20 if s >= 1.5 else 50
    first = int(ox // 5 + 1) * 5
    for mm in range(first, int(ox + W / s) + 1, 5):
        X = M + int((mm - ox) * s)
        long_ = mm % step == 0
        _box(img, TW, TH, X, M - (10 if long_ else 5), X + 1, M, gray if not long_ else (0, 0, 0))
        if long_:
            _text(img, TW, TH, X + 2, 2, str(mm))
    first = int(oy // 5 + 1) * 5
    for mm in range(first, int(oy + H / s) + 1, 5):
        Y = M + int((mm - oy) * s)
        long_ = mm % step == 0
        _box(img, TW, TH, M - (10 if long_ else 5), Y, M, Y + 1, gray if not long_ else (0, 0, 0))
        if long_:
            _text(img, TW, TH, 1, Y + 2, str(mm))
    notes = []
    blue = (0, 70, 220)
    for i, m in enumerate(written or [], 1):
        bx = m["box"]
        X0, Y0 = M + int((bx[0] - ox) * s), M + int((bx[1] - oy) * s)
        X1, Y1 = M + int((bx[2] - ox) * s), M + int((bx[3] - oy) * s)
        _box(img, TW, TH, X0, Y0, X1 + 1, Y0 + 1, blue)
        _box(img, TW, TH, X0, Y1, X1 + 1, Y1 + 1, blue)
        _box(img, TW, TH, X0, Y0, X0 + 1, Y1 + 1, blue)
        _box(img, TW, TH, X1, Y0, X1 + 1, Y1 + 1, blue)
        _text(img, TW, TH, X1 + 3, Y0, str(i), blue)
    red = (220, 0, 0)
    for i, m in enumerate(marks or [], 1):
        bx = m.get("box") if isinstance(m, dict) else None
        if not (isinstance(bx, list) and len(bx) == 4):
            notes.append(f"{i} 番目の枠：box は [左, 上, 右, 下]（mm）で指定してください")
            continue
        X0, Y0 = M + int((bx[0] - ox) * s), M + int((bx[1] - oy) * s)
        X1, Y1 = M + int((bx[2] - ox) * s), M + int((bx[3] - oy) * s)
        for t in range(2):
            _box(img, TW, TH, X0 - t, Y0 - t, X1 + t + 1, Y0 - t + 1, red)
            _box(img, TW, TH, X0 - t, Y1 + t, X1 + t + 1, Y1 + t + 1, red)
            _box(img, TW, TH, X0 - t, Y0 - t, X0 - t + 1, Y1 + t + 1, red)
            _box(img, TW, TH, X1 + t, Y0 - t, X1 + t + 1, Y1 + t + 1, red)
        lab = str(m.get("label") or i)
        _text(img, TW, TH, X0 + 2, max(M, Y0 - 13), "".join(ch for ch in lab if ch in _DIGITS) or str(i), red)
    return _png_rgb(img, TW, TH), TW, TH, notes


def _png_rgb(img: bytes, w: int, h: int) -> bytes:
    def chunk(t: bytes, b: bytes) -> bytes:
        return struct.pack(">I", len(b)) + t + b + struct.pack(">I", zlib.crc32(t + b) & 0xFFFFFFFF)

    raw = b"".join(b"\0" + bytes(img[y * w * 3:(y + 1) * w * 3]) for y in range(h))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 6)) + chunk(b"IEND", b""))


def view_page(doc, page_no: int, region_mm: list | None = None, marks: list | None = None, max_side: int = 1600,
              written: list | None = None) -> dict:
    """エージェントに見せるページの画像（PNG・base64）と、読み方の説明。written：値の入った欄（青い枠と番号で描く）。"""
    r = render(doc, page_no, region_mm, max_side)
    png, w, h, notes = with_ruler(r, marks, written)
    ox, oy = r["origin_mm"]
    pw, ph = r["page_mm"]
    text = (f"{page_no} ページ（{pw:.0f}×{ph:.0f} mm）の、左 {ox:.0f}〜{ox + r['w'] / r['scale']:.0f} mm・"
            f"上 {oy:.0f}〜{oy + r['h'] / r['scale']:.0f} mm の範囲（1 mm＝{r['scale']:.1f} 画素）。"
            "上と左の帯の数は、ページの左上からの mm です。書く所は、fill_document に "
            '{"@1": {"page": ページ, "box": [左, 上, 右, 下]（mm）, "value": 書く字}} で指定します。'
            "小さな字は region（[左, 上, 右, 下] mm）で拡大して読み、書く前に marks（[{\"box\": [...]}]）で枠を描いて、"
            "位置を目で確かめてください。この画像には、ページに置かれた画像だけを描いています（字や線のデータは窓で読みます）。")
    if r["skipped"]:
        text += " 描けなかった画像：" + "／".join(r["skipped"][:5])
    if written:
        text += (" 青い枠は、値の入った入力欄です（この画像には字を描いていないので、値は次のとおり）："
                 + "／".join(f"{i}＝{m['id']}「{m['value']}」" for i, m in enumerate(written, 1)))
    if notes:
        text += " " + " ".join(notes)
    return {"text": text, "png_base64": base64.b64encode(png).decode("ascii"), "width": w, "height": h}


_CAN = {"FlateDecode", "Fl", "ASCII85Decode", "A85", "ASCIIHexDecode", "AHx", "LZWDecode", "LZW", "RunLengthDecode", "RL",
        "DCTDecode", "DCT", "CCITTFaxDecode", "CCF", "Crypt"}


def unreadable_images(doc, page: dict) -> list[str]:
    """ページの画像のうち、todoku が開けない圧縮（JBIG2・JPEG 2000 など）のもの（画像の名前：圧縮）。"""
    from .objects import Stream

    out = []
    for pic in page_content(doc, page).pictures:
        im = pic.image
        d = im.dict if isinstance(im, Stream) else (im[0] if isinstance(im, (tuple, list)) and im else {})
        f = doc.resolve(d.get("Filter", d.get("F")))
        names = [getattr(doc.resolve(x), "name", "") for x in (f if isinstance(f, list) else [f] if f is not None else [])]
        bad = [n for n in names if n not in _CAN]
        if bad:
            out.append(f"{pic.name}：{'・'.join(bad)}")
    return out
