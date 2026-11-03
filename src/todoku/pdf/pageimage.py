"""PDF のページを、エージェントが見る画像にする（スキャンした紙のページ）。

ページに置かれた画像を、置いた行列（CTM）どおりに、見た目の向き（/Rotate を反映）でページの大きさの画像に描く。
端に、ページの左上からの mm の目盛りを付ける：エージェントは画像を見て、書く所を mm の枠で指定する。
書く予定の枠（marks）を赤で描き込める：書く前に、枠の位置をエージェントが目で確かめるため。
ページの字や線（ベクトル）は描かない（画像だけのページのための道具。字のあるページは窓で読む）。
欄・注釈の見た目（/AP）の字（todoku が書いた欄の字など）は、見た目のフォームを仕様どおりの行列で置いた位置・大きさ・向きで、
字ごとの四角（字の送りの幅 × 字の高さ）と、その字の基線（字の下の線）として描く：書いた字が欄に正しい向き・大きさで
収まっているかを、エージェントが目で確かめられる（字の形そのものは描かない。字の値は説明の文に添える）。
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
    pw, ph = view.width / view.mm, view.height / view.mm  # mm（UserUnit のあるページも、紙の上の mm）
    x0, y0, x1, y1 = region_mm or [0, 0, pw, ph]
    x0, y0, x1, y1 = max(0.0, min(x0, x1)), max(0.0, min(y0, y1)), min(pw, max(x0, x1)), min(ph, max(y0, y1))
    if x1 - x0 < 1 or y1 - y0 < 1:
        raise ValueError(f"範囲が小さすぎるか、ページ（{pw:.0f}×{ph:.0f} mm）の外です")
    s = max_side / max(x1 - x0, y1 - y0)  # 1 mm の画素
    W, H = max(1, int((x1 - x0) * s)), max(1, int((y1 - y0) * s))
    canvas = bytearray(b"\xff" * (W * H))
    pc = page_content(doc, page)
    skipped = []
    stamps = []
    annot_failed: list[str] = []
    cells: list = []
    pics = list(pc.pictures) + _annot_pictures(doc, page, stamps, annot_failed, cells)
    for pic in pics:
        if pic.ctm is None or pic.image is None:
            continue
        inv = _inv(pic.ctm)
        if inv is None:
            continue
        try:
            g = _with_smask(doc, pic.image, image_gray(doc, pic.image, max_pixels=max(4_000_000, W * H * 2)))
        except (RasterError, ValueError, IndexError, KeyError, zlib.error) as e:
            skipped.append(f"{pic.name}：{e}")
            continue
        _draw(canvas, W, H, s, x0, y0, view, inv, g)
    for gl in cells:
        _draw_glyph(canvas, W, H, s, x0, y0, view, gl)
    rgb = bytearray(W * H * 3)
    rgb[0::3] = canvas
    rgb[1::3] = canvas
    rgb[2::3] = canvas
    return {"rgb": rgb, "w": W, "h": H, "scale": s, "origin_mm": (x0, y0), "page_mm": (pw, ph), "skipped": skipped,
            "view": view, "stamps": stamps, "annot_failed": annot_failed, "glyphs": len(cells), "user_unit": view.unit}


def _annot_pictures(doc, page: dict, stamps: list, failed: list | None = None, glyphs: list | None = None) -> list:
    """注釈の見た目（/AP の N）に描かれた画像：置いた印影・署名の画像（スタンプの注釈）など。
    見た目のフォームを、仕様（ISO 32000-1 12.5.5）どおり、BBox を Matrix で写した四角を Rect に合わせる行列で描く。
    隠した注釈（F の Hidden）と、印刷しない注釈は描かない。stamps には、画像のある注釈の名前と位置（mm）を入れる。
    failed には、見た目を描けなかった注釈（名前と理由）を入れる（黙って描かないままにしない）。
    glyphs には、見た目に描かれた字（content.Glyph：ページの座標。同じ行列で置いたもの）を入れる。"""
    from .content import Interpreter, apply, mul
    from .objects import PdfString, Stream, decode_text

    out = []
    view = View(doc, page)
    for a in doc.resolve(page.get("Annots")) or []:
        ad = doc.resolve(a)
        if not isinstance(ad, dict) or getattr(ad.get("Subtype"), "name", "") in ("Popup", "Link"):
            continue
        flags = int(doc.resolve(ad.get("F")) or 0)
        if flags & 2:
            continue
        apd = doc.resolve(ad.get("AP"))
        n = doc.resolve(apd.get("N")) if isinstance(apd, dict) else None
        if isinstance(n, dict) and not isinstance(n, Stream):
            st = doc.resolve(ad.get("AS"))
            n = doc.resolve(n.get(getattr(st, "name", "")))
        if not isinstance(n, Stream):
            continue
        try:
            rect = [float(doc.resolve(v)) for v in doc.resolve(ad.get("Rect"))]
            bbox = [float(doc.resolve(v)) for v in doc.resolve(n.dict.get("BBox"))]
            m = n.dict.get("Matrix")
            m = tuple(float(doc.resolve(v)) for v in doc.resolve(m)) if m is not None else (1, 0, 0, 1, 0, 0)
        except (TypeError, ValueError) as e:
            if failed is not None:
                failed.append(f"{getattr(ad.get('Subtype'), 'name', '注釈')}：位置か大きさを読めない（{type(e).__name__}）")
            continue
        if len(rect) != 4 or len(bbox) != 4 or len(m) != 6:
            if failed is not None:
                failed.append(f"{getattr(ad.get('Subtype'), 'name', '注釈')}：位置か大きさの数が足りない")
            continue
        pts = [apply(m, x, y) for x in (bbox[0], bbox[2]) for y in (bbox[1], bbox[3])]
        bx0, by0 = min(p[0] for p in pts), min(p[1] for p in pts)
        bx1, by1 = max(p[0] for p in pts), max(p[1] for p in pts)
        if bx1 - bx0 < 1e-6 or by1 - by0 < 1e-6:
            continue
        rx0, ry0, rx1, ry1 = min(rect[0], rect[2]), min(rect[1], rect[3]), max(rect[0], rect[2]), max(rect[1], rect[3])
        sx, sy = (rx1 - rx0) / (bx1 - bx0), (ry1 - ry0) / (by1 - by0)
        fit = (sx, 0, 0, sy, rx0 - bx0 * sx, ry0 - by0 * sy)
        try:
            pc = Interpreter(doc).run_stream(doc.decode(n), doc.resolve(n.dict.get("Resources")) or {}, mul(m, fit))
        except Exception as e:  # noqa: BLE001
            if failed is not None:
                vx0, vy0, vx1, vy1 = view.rect(rx0, ry0, rx1, ry1)
                failed.append(f"{getattr(ad.get('Subtype'), 'name', '注釈')}（左 {vx0 / view.mm:.0f}・上 {vy0 / view.mm:.0f}・右 {vx1 / view.mm:.0f}・"
                              f"下 {vy1 / view.mm:.0f} mm）：{type(e).__name__}")
            continue
        if glyphs is not None:
            glyphs += [g for g in pc.glyphs if not g.invisible and g.text.strip()]
        if pc.pictures:
            out += pc.pictures
            nm = ad.get("NM")
            vx0, vy0, vx1, vy1 = view.rect(rx0, ry0, rx1, ry1)
            stamps.append({"id": decode_text(nm) if isinstance(nm, (bytes, PdfString)) else "",
                           "kind": getattr(ad.get("Subtype"), "name", ""),
                           "box": [vx0 / view.mm, vy0 / view.mm, vx1 / view.mm, vy1 / view.mm]})
    return out


def _with_smask(doc, image, g: Gray) -> Gray:
    """画像の透け（SMask）：白い紙の上に重ねた濃さにする（透ける所は白）。"""
    from .objects import Stream

    if not isinstance(image, Stream):
        return g
    sm = doc.resolve(image.dict.get("SMask"))
    if not isinstance(sm, Stream):
        return g
    a = image_gray(doc, sm, max_pixels=max(4_000_000, g.w * g.h))
    px = bytearray(g.px)
    for y in range(g.h):
        ay = min(a.h - 1, y * a.h // g.h)
        for x in range(g.w):
            al = a.px[ay * a.w + min(a.w - 1, x * a.w // g.w)]
            if al != 255:
                i = y * g.w + x
                px[i] = 255 - (255 - px[i]) * al // 255
    return Gray(g.w, g.h, bytes(px))


def _draw(canvas: bytearray, W: int, H: int, s: float, x0: float, y0: float, view: View, inv, g: Gray) -> None:
    """見た目の画素 (X, Y) → mm → 見た目の pt → ページの座標 → 画像の単位の四角 → 画像の画素。1行ずつ、一次の式で。"""
    a, b, c, d, e, f = inv

    def img_xy(X: float, Y: float) -> tuple[float, float]:
        vx, vy = (x0 + X / s) * view.mm, (y0 + Y / s) * view.mm
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


def _draw_glyph(canvas: bytearray, W: int, H: int, s: float, x0: float, y0: float, view: View, gl) -> None:
    """見た目の字 1 字：字の四角（薄い灰色、縁は濃い灰色）と、基線（原点から送りの向きへの黒い線）。座標はページの座標。"""
    def px(x: float, y: float) -> tuple[float, float]:
        vx, vy = view.pt(x, y)
        return (vx / view.mm - x0) * s, (vy / view.mm - y0) * s

    a, b = px(gl.x0, gl.y0), px(gl.x1, gl.y1)
    X0, X1 = int(min(a[0], b[0])), int(max(a[0], b[0]))
    Y0, Y1 = int(min(a[1], b[1])), int(max(a[1], b[1]))
    X0, Y0, X1, Y1 = max(0, X0), max(0, Y0), min(W - 1, X1), min(H - 1, Y1)
    if X1 < X0 or Y1 < Y0:
        return
    for Y in range(Y0, Y1 + 1):
        row = Y * W
        for X in range(X0, X1 + 1):
            v = 90 if X in (X0, X1) or Y in (Y0, Y1) else 205
            if v < canvas[row + X]:
                canvas[row + X] = v
    # 基線：字の原点から、字の送りの向きに、字の送りの幅だけ（向きが分かるように）
    adv = gl.adv or ((gl.dx ** 2 + gl.dy ** 2) ** 0.5)
    ln = (gl.dx ** 2 + gl.dy ** 2) ** 0.5
    if ln <= 0 or adv <= 0:
        return
    ux, uy = gl.dx / ln * adv, gl.dy / ln * adv
    p0, p1 = px(gl.ox, gl.oy), px(gl.ox + ux, gl.oy + uy)
    steps = max(1, int(max(abs(p1[0] - p0[0]), abs(p1[1] - p0[1])) * 2))
    for i in range(steps + 1):
        t = i / steps
        cx, cy = p0[0] + (p1[0] - p0[0]) * t, p0[1] + (p1[1] - p0[1]) * t
        for ox in (0, 1):
            for oy in (0, 1):
                X, Y = int(cx) - ox, int(cy) - oy
                if 0 <= X < W and 0 <= Y < H:
                    canvas[Y * W + X] = 0


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
            "位置を目で確かめてください。この画像には、ページに置かれた画像だけを描いています（ページの字や線のデータは窓で読みます）。")
    if r.get("user_unit", 1.0) != 1.0:
        text += (f" このページは、座標の 1 単位が {r['user_unit']:g}/72 インチ（UserUnit {r['user_unit']:g}）です。mm は紙の上の"
                 "長さに直してあります。")
    if r.get("glyphs"):
        text += (" 欄・注釈の見た目（書いた欄の字など）の字は、見た目のとおりの位置・大きさ・向きで、字ごとの灰色の四角として"
                 "描いています（字の形そのものは描きません）。各字の黒い線は基線（字の下の線）で、字が正しい向きなら四角の下側に"
                 "左から右へ並びます。四角が欄からはみ出す・つぶれる・横倒しなら、書いた見た目が正しくありません。")
    if r["skipped"]:
        text += " 描けなかった画像：" + "／".join(r["skipped"][:5])
    if r.get("annot_failed"):
        text += (" 見た目を描けなかった注釈があります（この画像には描いていないので、その所に見える印影・字などは、この画像からは"
                 "分かりません）：" + "／".join(r["annot_failed"][:5]))
    if r.get("stamps"):
        text += (" このページに置かれた画像（印影・署名などの注釈）も描いています："
                 + "／".join(f"{m['id'] or m['kind']}（左 {m['box'][0]:.0f}・上 {m['box'][1]:.0f}・右 {m['box'][2]:.0f}・"
                            f"下 {m['box'][3]:.0f} mm）" for m in r["stamps"]))
    if written:
        text += (" 青い枠は、値の入った入力欄です（字の形は描いていないので、値は次のとおり）："
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
