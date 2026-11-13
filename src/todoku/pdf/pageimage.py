"""PDF のページを、エージェントが見る画像にする（見る道具（ビューア）が描くとおりに）。

ページの描画の命令を、仕様（ISO 32000-1 8〜9 章）どおりに実行して描く（render/paint.py）：線・塗り・切り抜き・色・画像・
フォームの部品・字（埋め込んだ書体の字の形。埋め込みのない書体は、このパソコンの書体で代わりに描く）・注釈の見た目（/AP、
todoku が書いた欄の字も）。見た目の向き（/Rotate）・CropBox・UserUnit を反映する。
端に、ページの左上からの mm の目盛りを付ける：エージェントは画像を見て、書く所を mm の枠で指定する。
書く予定の枠（marks）を赤で描き込める：書く前に、枠の位置をエージェントが目で確かめるため。
描けなかったもの（字の形の分からない字は灰色の四角で示す・開けない画像・近い形で描いたもの）は、説明の文で伝える。
"""

from __future__ import annotations

import base64
import math
import struct
import zlib

from .content import page_content
from .layout import View

MM = 72 / 25.4  # 1 mm のポイント
DEFAULT_DPI = 100
MAX_SIDE = 4000  # 画像の 1 辺の画素の上限
MAX_PIXELS = 12_000_000


def _base_matrix(view: View, s: float, x0: float, y0: float):
    """ページの座標 → 画像の画素（見た目の向き・切り取り・1 mm＝s 画素）。"""
    o = view.pt(0.0, 0.0)
    ex = view.pt(1.0, 0.0)
    ey = view.pt(0.0, 1.0)
    k = s / view.mm
    return ((ex[0] - o[0]) * k, (ex[1] - o[1]) * k, (ey[0] - o[0]) * k, (ey[1] - o[1]) * k,
            o[0] * k - x0 * s, o[1] * k - y0 * s)


def render(doc, page_no: int, region_mm: list | None = None, max_side: int | None = None, dpi: float | None = None,
           time_limit: float | None = None, probe_of=None) -> dict:
    """ページ（の範囲）の画像。region_mm：[左, 上, 右, 下]（ページの左上からの mm）。
    大きさ：dpi を指定すればその解像度。なければ、max_side を指定すれば長い辺をその画素に、どちらもなければ、ページ全体は
    100 dpi・範囲は長い辺 1600 画素（拡大して読む）。
    返す：{rgb: 画素（RGB）, w, h, scale（1 mm の画素）, origin_mm（画像の左上の mm）, skipped（描けなかった画像）, ...}。"""
    from ..limits import TIME_LIMIT
    from .render.paint import appearance_notes, render_page

    if time_limit is None:
        time_limit = min(50.0, TIME_LIMIT * 0.4)  # 描いた所までを返す余裕を残す
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
    if dpi:
        s = float(dpi) / 25.4
    elif max_side:
        s = max_side / max(x1 - x0, y1 - y0)
    elif region_mm:
        s = 1600 / max(x1 - x0, y1 - y0)
    else:
        s = DEFAULT_DPI / 25.4
    s = min(s, MAX_SIDE / max(x1 - x0, y1 - y0), (MAX_PIXELS / ((x1 - x0) * (y1 - y0))) ** 0.5)  # 1 mm の画素
    W, H = max(1, math.ceil((x1 - x0) * s - 1e-6)), max(1, math.ceil((y1 - y0) * s - 1e-6))
    cv, rr = render_page(doc, page, W, H, _base_matrix(view, s, x0, y0), time_limit, probe_of)
    stamps: list = []
    annot_failed: list[str] = []
    cells: list = []
    _annot_pictures(doc, page, stamps, annot_failed, cells)
    for f in getattr(rr, "annot_failed", []):
        if not any(f.split("：")[0] in x for x in annot_failed):
            annot_failed.append(f)
    skipped = [f"{nm}（左 {x0 + b[0] / s:.0f}・上 {y0 + b[1] / s:.0f}・右 {x0 + b[2] / s:.0f}・下 {y0 + b[3] / s:.0f} mm）：{why}"
               for nm, why, b in rr.image_failed]
    return {"rgb": cv.buf, "w": W, "h": H, "scale": s, "origin_mm": (x0, y0), "page_mm": (pw, ph),
            "skipped": skipped, "view": view, "stamps": stamps, "annot_failed": annot_failed, "glyphs": len(cells),
            "user_unit": view.unit, "dpi": s * 25.4, "stats": dict(rr.stats), "substituted": dict(rr.subst),
            "missing_fonts": sorted(rr.missing_fonts), "approx": dict(rr.approx), "stopped": rr.stopped,
            "problems": list(rr.problems), "no_appearance": appearance_notes(doc, page), "image_notes": list(rr.image_notes),
            "box_places": [(round(x0 + px / s), round(y0 + py / s)) for px, py in rr.box_places], "box_inks": rr.box_inks}


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


def with_ruler(r: dict, marks: list | None = None, written: list | None = None, ticks: tuple | None = None) -> tuple[bytes, int, int, list[str]]:
    """目盛り（10 mm ごとに数、5 mm ごとに短い線）を上と左の帯に付け、marks（{box:[左,上,右,下] mm, label}）を赤で描く。
    ticks：(短い線の間隔, 数の間隔)。解像度の分からない画像は、単位を画素にして渡す。"""
    M = 26  # 帯の幅（画素）
    W, H, s = r["w"], r["h"], r["scale"]
    ox, oy = r["origin_mm"]
    TW, TH = W + M, H + M
    img = bytearray(b"\xff" * (TW * TH * 3))
    for y in range(H):
        img[((y + M) * TW + M) * 3:((y + M) * TW + M + W) * 3] = r["rgb"][y * W * 3:(y + 1) * W * 3]
    gray = (150, 150, 150)
    step = 10 if s >= 3 else 20 if s >= 1.5 else 50
    tick = 5
    if ticks:
        tick, step = ticks
    # 目盛りの数は画素の数で決まる（ページの宣言の大きさで増やさない）：線が 4 画素、数が 16 画素より詰まるなら間隔を 10 倍に
    while tick * s < 4:
        tick *= 10
    while step * s < 16 or step % tick:
        step *= 10
    first = int(ox // tick + 1) * tick
    for mm in range(first, int(ox + W / s) + 1, tick):
        X = M + int((mm - ox) * s)
        long_ = mm % step == 0
        _box(img, TW, TH, X, M - (10 if long_ else 5), X + 1, M, gray if not long_ else (0, 0, 0))
        if long_:
            _text(img, TW, TH, X + 2, 2, str(mm))
    first = int(oy // tick + 1) * tick
    for mm in range(first, int(oy + H / s) + 1, tick):
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


def view_page(doc, page_no: int, region_mm: list | None = None, marks: list | None = None, max_side: int | None = None,
              written: list | None = None, dpi: float | None = None) -> dict:
    """エージェントに見せるページの画像（PNG・base64）と、読み方の説明。written：値の入った欄（青い枠と番号で描く）。
    dpi：解像度（なければ、ページ全体は 100 dpi、範囲は長い辺 1600 画素）。"""
    r = render(doc, page_no, region_mm, max_side, dpi)
    png, w, h, notes = with_ruler(r, marks, written)
    ox, oy = r["origin_mm"]
    pw, ph = r["page_mm"]
    st = r.get("stats", {})
    text = (f"{page_no} ページ（{pw:.0f}×{ph:.0f} mm）の、左 {ox:.0f}〜{ox + r['w'] / r['scale']:.0f} mm・"
            f"上 {oy:.0f}〜{oy + r['h'] / r['scale']:.0f} mm の範囲（1 mm＝{r['scale']:.1f} 画素・{r['dpi']:.0f} dpi）。"
            "上と左の帯の数は、ページの左上からの mm です。書く所は、fill_document に "
            '{"@1": {"page": ページ, "box": [左, 上, 右, 下]（mm）, "value": 書く字}} で指定します。'
            "小さな字は region（[左, 上, 右, 下] mm）で拡大して読み、書く前に marks（[{\"box\": [...]}]）で枠を描いて、"
            "位置を目で確かめてください。この画像は、ページを見る道具と同じように描いたものです（字・線・画像・欄と注釈の見た目。"
            "書いた欄の字も、その書体の字の形で描いています）。")
    if r.get("user_unit", 1.0) != 1.0:
        text += (f" このページは、座標の 1 単位が {r['user_unit']:g}/72 インチ（UserUnit {r['user_unit']:g}）です。mm は紙の上の"
                 "長さに直してあります。")
    if r.get("substituted"):
        subs = r["substituted"]
        text += (" 埋め込まれていない書体は、このパソコンの書体で代わりに描いています（字の形・幅が、作った人の画面と少し違う"
                 "ことがあります）：" + "／".join(f"{k} → {v}" for k, v in list(subs.items())[:6]))
    if st.get("glyphs_fallback"):
        text += (f" 書体に字の形がない字 {st['glyphs_fallback']} 字は、字（ToUnicode など）が分かるので、このパソコンの別の書体で描いています"
                 "（字は同じで、形が元と違います）。")
    if st.get("glyphs_box"):
        miss = r.get("missing_fonts") or []
        text += (f" 字の形が分からない字 {st['glyphs_box']} 字は、字の位置に灰色の四角で示しています（字の形も字も分かりません。この画像でも読めません。"
                 "その所の中身は推測で埋めず、本人に伝えてください）" + (f"。代わりの書体が見つからない書体：{'・'.join(miss[:5])}" if miss else "")
                 + "。位置（左・上 mm）：" + "／".join(f"{x}・{y}" for x, y in _places(r)) + "。")
    if r.get("approx"):
        text += (" 仕様どおりには描けず、近い形で描いたもの：" + "・".join(f"{k} {v} か所" for k, v in r["approx"].items()) + "。")
    if r.get("stopped"):
        text += " 描くのに時間がかかりすぎるため、途中で止めました（描いていない所があります）。"
    if r.get("image_notes"):
        text += " 画像の開き方：" + "／".join(r["image_notes"][:5]) + "（細かな字が読めないときは、region で範囲を狭めて拡大すると、予算の中で細かな解像度で開きます）。"
    if r["skipped"]:
        text += " 描けなかった画像：" + "／".join(r["skipped"][:5])
    if r.get("annot_failed"):
        text += (" 見た目を描けなかった注釈があります（この画像には描いていないので、その所に見える印影・字などは、この画像からは"
                 "分かりません）：" + "／".join(r["annot_failed"][:5]))
    if r.get("problems"):
        text += " 描けなかった所：" + "／".join(r["problems"][:5])
    if r.get("no_appearance"):
        text += " " + " ".join(r["no_appearance"])
    if r.get("stamps"):
        text += (" このページに置かれた画像（印影・署名などの注釈）も描いています："
                 + "／".join(f"{m['id'] or m['kind']}（左 {m['box'][0]:.0f}・上 {m['box'][1]:.0f}・右 {m['box'][2]:.0f}・"
                            f"下 {m['box'][3]:.0f} mm）" for m in r["stamps"]))
    if written:
        text += (" 青い枠は、値の入った入力欄です（値は次のとおり）："
                 + "／".join(f"{i}＝{m['id']}「{m['value']}」" for i, m in enumerate(written, 1)))
    if notes:
        text += " " + " ".join(notes)
    return {"text": text, "png_base64": base64.b64encode(png).decode("ascii"), "width": w, "height": h,
            "not_drawn": not_drawn(r)}


def _places(r: dict) -> list:
    """字の形が分からない字の位置（同じ所を除く）。"""
    seen: list = []
    for p in r.get("box_places") or []:
        if p not in seen:
            seen.append(p)
    return seen[:12]


def not_drawn(r: dict) -> list[str]:
    """このページで、そのとおりには描けなかったものの短い一覧（なければ空）。"""
    st = r.get("stats", {})
    out = [f"埋め込みのない書体 {k}（{v} で代わりに描いた）" for k, v in list(r.get("substituted", {}).items())[:6]]
    if st.get("glyphs_fallback"):
        out.append(f"書体に字の形がない字 {st['glyphs_fallback']} 字（ToUnicode の字を、別の書体の形で描いた）")
    if st.get("glyphs_box"):
        out.append(f"字の形が分からない字 {st['glyphs_box']} 字（灰色の四角で示した。位置（左・上 mm）：" + "／".join(f"{x}・{y}" for x, y in _places(r)) + "）")
    out += [f"{k} {v} か所（近い形で描いたか、飛ばした）" for k, v in r.get("approx", {}).items()]
    out += [f"画像 {s}（灰色の枠と対角線で示した）" for s in r.get("skipped", [])[:8]]
    out += [f"注釈の見た目 {s}" for s in r.get("annot_failed", [])[:8]]
    out += list(r.get("no_appearance", [])) + list(r.get("problems", []))[:8]
    if r.get("stopped"):
        out.append("時間がかかりすぎるので途中で止めた（描いていない所がある）")
    return out


_CAN = {"FlateDecode", "Fl", "ASCII85Decode", "A85", "ASCIIHexDecode", "AHx", "LZWDecode", "LZW", "RunLengthDecode", "RL",
        "DCTDecode", "DCT", "CCITTFaxDecode", "CCF", "Crypt", "JBIG2Decode", "JPXDecode"}


def unreadable_images(doc, page: dict) -> list[str]:
    """ページの画像のうち、PDF の規格にない名前の圧縮のもの（画像の名前：圧縮）。規格の圧縮はどれも開ける。"""
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
