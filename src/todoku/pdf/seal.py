"""PDF：本人の印影・署名の画像を、本人が同意したときだけ、決めた所に置く（スタンプの注釈）。

置き方（fill の値）：
- {"@1": {"seal": 画像のパス, "consent": true, "page": ページ, "box": [左, 上, 右, 下]（mm）}}：エージェントが決めた枠の中に
  （スキャンした紙のページでも。view_page で見て決める）
- {"@1": {"seal": 画像のパス, "consent": true, "after": "印"}}：その字の右（"where": "below" なら下）の空いている所に。
  {"cover": "印"} なら、その字の上に（白い下地は付けない。画像が透けないときは、その字が隠れる）
- "size"（mm）：画像の長い辺の長さ（字を目印にしたときだけ。省くと、目印の所の高さ・幅に収める）
- 署名の画像は "seal" の代わりに "signature"
画像は縦横の比を保って、枠（目印の所）の中に置く（枠と字の目印は中央、右・下は左にそろえる）。
透けない画像（アルファのない PNG・JPEG）がページの字を隠すときは、置かずに断る（字は消さない）。

書き方：ページの注釈の一覧に、スタンプの注釈（/Subtype /Stamp）を1つ足す。見た目（/AP）は、画像を枠いっぱいに描くだけの
フォームの XObject。画像は XObject（JPEG はファイルのバイト列のまま DCTDecode、PNG は開いた画素を Flate で縮め、
アルファは SMask）。元のページの描画は1バイトも変えない（増分の保存）。
証明（proof.py）と読み戻しは、ここの verify で、足した注釈・見た目・画像が、元の文書と画像から決め直したものと同じかを確かめる。
"""

from __future__ import annotations

import copy
import zlib

from ..sealimage import KINDS, SealError, SealImage, caution
from .anchor import MM, AnchorError, _place, resolve_write
from .content import page_content
from .layout import PageLayout, View
from .objects import Name, PdfString, Ref, Stream, decode_text, text_string

ANNOT_FLAGS = 4 | 128  # 印刷する（Print）・動かせない（Locked）
XNAME = "TdkSeal"  # 見た目の中の、画像の資源の名前
ANNOT_KEYS = {"Type", "Subtype", "Rect", "NM", "Contents", "F", "P", "AP"}


def contents_text(key: str) -> str:
    """注釈の説明（代わりの字）：画面の読み上げにも、何の画像かが分かるように。"""
    return f"本人の{KINDS[key]}の画像（本人の同意を得て置いたもの）"


def stamp_names(doc) -> set[str]:
    """文書にすでにある、スタンプの注釈の名前（/NM）：番号が重ならないように。"""
    out = set()
    for pg in doc.pages():
        for a in doc.resolve(pg.get("Annots")) or []:
            ad = doc.resolve(a)
            if isinstance(ad, dict) and isinstance(ad.get("NM"), (bytes, PdfString)):
                out.add(decode_text(ad["NM"]))
    return out


# ---------------------------------------------------------------------- 置く所を決める
def resolve(doc, spec: dict, key: str, img: SealImage, used: dict | None, layouts: dict | None,
            taken_names: set | None) -> dict:
    """置く所（見た目の座標の矩形、pt）。書き込みと証明で、同じ決め方を使う。
    返す：{id, page, rect（画像の矩形）, slot（目印の所・枠）, mode, anchor（目印の字）, under（画像の下にあるページの字）, label}。"""
    layouts = layouts if layouts is not None else {}
    pages = doc.pages()
    size = spec.get("size")
    if size is not None and not (isinstance(size, (int, float)) and not isinstance(size, bool) and size > 0):
        raise AnchorError("size は、画像の長い辺の長さ（mm、正の数）で指定してください")
    if spec.get("box") is not None:
        n, box = spec.get("page"), spec.get("box")
        if not (isinstance(n, int) and 1 <= n <= len(pages)):
            raise AnchorError(f"枠で置くときは page（1〜{len(pages)}）を指定してください")
        if not (isinstance(box, list) and len(box) == 4 and all(isinstance(v, (int, float)) for v in box)):
            raise AnchorError("box は [左, 上, 右, 下]（ページの左上からの mm）で指定してください")
        if size is not None:
            raise AnchorError("枠（box）で置くときは、size は使いません（画像は枠の中に収めます）")
        view = View(doc, pages[n - 1])
        slot = tuple(v * view.mm for v in box)  # mm → ページの座標（UserUnit も反映）
        if not (0 <= slot[0] < slot[2] <= view.width + 0.01 and 0 <= slot[1] < slot[3] <= view.height + 0.01):
            raise AnchorError(f"box がページ（{view.width / view.mm:.0f}×{view.height / view.mm:.0f} mm）の外か、左＜右・上＜下になっていません")
        mode, anchor = "box", ""
    else:
        if spec.get("after") is None and spec.get("cover") is None:
            raise AnchorError(f'{key} の画像を置く所は、{{"page": …, "box": [左, 上, 右, 下]（mm）}} か、字の目印'
                              '（{"after": "印"}・{"cover": "印"}）で指定してください')
        sub = {k: spec[k] for k in ("after", "cover", "where", "page", "occurrence") if k in spec}
        sub["value"] = ""
        # 目印の所を決めるだけ（番号と重なりは、画像の矩形で下で決める）
        sd = resolve_write(doc, sub, copy.deepcopy(used) if used else {}, layouts, taken_names)
        n = sd["page"]
        slot = tuple(sd["rect"])
        view = View(doc, pages[n - 1])
        mode = "cover" if spec.get("cover") is not None else spec.get("where", "right")
        anchor = spec.get("cover") if mode == "cover" else spec.get("after")
    x0, y0, x1, y1 = slot
    if size is None:
        s = min((x1 - x0) / img.w, (y1 - y0) / img.h)
    else:
        s = size * view.mm / max(img.w, img.h)  # 長い辺が size mm（UserUnit のあるページも、紙の上で size mm）
    w, h = img.w * s, img.h * s
    if mode in ("box", "cover"):
        left, top = (x0 + x1 - w) / 2, (y0 + y1 - h) / 2
    elif mode == "right":
        left, top = x0, (y0 + y1 - h) / 2
    else:
        left, top = x0, y0
    rect = (left, top, left + w, top + h)
    if rect[0] < -0.01 or rect[1] < -0.01 or rect[2] > view.width + 0.01 or rect[3] > view.height + 0.01:
        raise AnchorError(f"画像（{w / view.mm:.0f}×{h / view.mm:.0f} mm）が、ページの外にはみ出します（size を小さくしてください）")
    if n not in layouts:
        layouts[n] = PageLayout(doc, pages[n - 1], page_content(doc, pages[n - 1]), n)
    lay = layouts[n]

    def inside(c, r):
        return r[0] < (c.x0 + c.x1) / 2 < r[2] and r[1] < (c.y0 + c.y1) / 2 < r[3]

    under = [c for c in lay.chars if not c.text.isspace() and inside(c, rect) and not (mode == "cover" and inside(c, slot))]
    if under and img.opaque:
        raise AnchorError(f"画像は透けない（アルファのない PNG・JPEG）ので、置くとページの字（「{''.join(c.text for c in under[:8])}」）"
                          "が隠れます。透ける PNG にするか、置く所・size を変えてください")
    what = KINDS[key]
    if mode == "box":
        label = (f"{n}ページの枠（左 {slot[0] / view.mm:.0f}・上 {slot[1] / view.mm:.0f}・右 {slot[2] / view.mm:.0f}・"
                 f"下 {slot[3] / view.mm:.0f} mm）の中")
    else:
        label = f"{n}ページの「{anchor}」の" + {"right": "右", "below": "下", "cover": "上"}[mode]
    out = _place(n, rect, used, taken_names, {"id": None, "page": n, "kind": "seal"}, "seal", anchor, 0.0,
                 prefix="s", label=label)
    out.update({"slot": [round(v, 2) for v in slot], "mode": mode, "anchor": anchor, "what": what, "mm": view.mm,
                "under": "".join(c.text for c in under)})
    return out


# ---------------------------------------------------------------------- 書く
def image_objects(img: SealImage) -> tuple[Stream, Stream | None]:
    """画像の XObject（と、透けがあれば SMask）。"""
    cs = Name("DeviceGray") if img.channels == 1 else Name("DeviceRGB")
    d = {"Type": Name("XObject"), "Subtype": Name("Image"), "Width": img.w, "Height": img.h, "ColorSpace": cs,
         "BitsPerComponent": 8}
    if img.kind == "jpeg":
        d["Filter"] = Name("DCTDecode")
        return Stream(d, img.data), None
    d["Filter"] = Name("FlateDecode")
    smask = None
    if img.alpha is not None:
        smask = Stream({"Type": Name("XObject"), "Subtype": Name("Image"), "Width": img.w, "Height": img.h,
                        "ColorSpace": Name("DeviceGray"), "BitsPerComponent": 8, "Filter": Name("FlateDecode")},
                       zlib.compress(img.alpha, 9))
    return Stream(d, zlib.compress(img.samples or b"", 9)), smask


def page_rect(view: View, rect) -> list[float]:
    x0, y0, x1, y1 = view.back_rect(*rect)
    return [round(x0, 3), round(y0, 3), round(x1, 3), round(y1, 3)]


def appearance(view: View, prect: list[float]) -> tuple[dict, bytes]:
    """見た目のフォームの辞書（資源は除く）と、描く命令：画像を、見た目の向きで枠いっぱいに描くだけ。"""
    from .fill import _rot_matrix

    w, h = prect[2] - prect[0], prect[3] - prect[1]
    vw, vh = (h, w) if view.rot in (90, 270) else (w, h)
    d = {"Type": Name("XObject"), "Subtype": Name("Form"), "BBox": [0, 0, round(vw, 3), round(vh, 3)]}
    m = _rot_matrix(view.rot, w, h)
    if m is not None:
        d["Matrix"] = [round(v, 3) for v in m]
    ops = f"q\n{vw:.3f} 0 0 {vh:.3f} 0 0 cm\n/{XNAME} Do\nQ\n".encode()
    return d, ops


def write(filler, page_no: int, sd: dict, key: str, img: SealImage, view: View) -> Ref:
    """スタンプの注釈を足す（filler は fill.Filler）。"""
    up = filler.up
    page_ref = filler.pages[page_no - 1].get("_ref")
    if page_ref is None:
        raise SealError(f"{sd['id']}: ページのオブジェクトが見つかりません")
    im, smask = image_objects(img)
    if smask is not None:
        im.dict["SMask"] = up.new(smask)
    im_ref = up.new(im)
    prect = page_rect(view, sd["rect"])
    d, ops = appearance(view, prect)
    d["Resources"] = {"XObject": {XNAME: im_ref}}
    ap = up.new(Stream(d, ops))
    annot = {"Type": Name("Annot"), "Subtype": Name("Stamp"), "Rect": prect, "NM": text_string(sd["id"]),
             "Contents": text_string(contents_text(key)), "F": ANNOT_FLAGS, "P": page_ref, "AP": {"N": ap}}
    ref = up.new(annot)
    filler._attach(page_no, ref, field=False)
    return ref


# ---------------------------------------------------------------------- 確かめる（証明・読み戻し）
def find_stamp(doc, page_no: int, sid: str):
    """ページの注釈の一覧から、名前（/NM）が sid のスタンプの注釈：(オブジェクトの番号, 辞書)。なければ (None, None)。"""
    pages = doc.pages()
    if not 1 <= page_no <= len(pages):
        return None, None
    for a in doc.resolve(pages[page_no - 1].get("Annots")) or []:
        ad = doc.resolve(a)
        if isinstance(ad, dict) and getattr(ad.get("Subtype"), "name", "") == "Stamp" \
                and isinstance(ad.get("NM"), (bytes, PdfString)) and decode_text(ad["NM"]) == sid:
            return (a.num if isinstance(a, Ref) else None), ad
    return None, None


def verify(doc, page_no: int, sd: dict, key: str, img: SealImage) -> tuple[list[str], set[int]]:
    """文書の中のスタンプの注釈が、決め直した置き方（sd）と画像（img）のとおりかを確かめる。
    返す：(違い, 確かめたオブジェクトの番号：注釈・見た目・画像・SMask)。"""
    sid = sd["id"]
    probs: list[str] = []
    nums: set[int] = set()
    num, ad = find_stamp(doc, page_no, sid)
    if ad is None:
        return [f"{sid}：{page_no} ページに、置いたはずの{KINDS[key]}の画像（スタンプの注釈）がありません"], nums
    if num is None:
        return [f"{sid}：注釈が、別のオブジェクトになっていません"], nums
    nums.add(num)
    pg = doc.pages()[page_no - 1]
    view = View(doc, pg)
    want = page_rect(view, sd["rect"])
    got = [float(doc.resolve(v)) for v in (doc.resolve(ad.get("Rect")) or [])]
    if len(got) != 4 or any(abs(a - b) > 0.01 for a, b in zip(want, got)):
        probs.append(f"{sid}：画像の位置 {got} が、決め直した位置 {want} と違います")
    extra = set(ad) - ANNOT_KEYS
    if extra:
        probs.append(f"{sid}：注釈に、思っていない項目 {sorted(extra)} があります")
    if getattr(ad.get("Type"), "name", "") != "Annot" or doc.resolve(ad.get("F")) != ANNOT_FLAGS:
        probs.append(f"{sid}：注釈の種類か印（F）が違います")
    if decode_text(doc.resolve(ad.get("Contents")) or b"") != contents_text(key):
        probs.append(f"{sid}：注釈の説明が違います")
    p = ad.get("P")
    if not (isinstance(p, Ref) and p.num == getattr(pg.get("_ref"), "num", None)):
        probs.append(f"{sid}：注釈の属するページ（P）が違います")
    apd = doc.resolve(ad.get("AP"))
    if not (isinstance(apd, dict) and set(apd) == {"N"} and isinstance(apd["N"], Ref)):
        return probs + [f"{sid}：見た目（AP）が、画像を描くフォーム1つではありません"], nums
    ap = doc.resolve(apd["N"])
    nums.add(apd["N"].num)
    if not isinstance(ap, Stream):
        return probs + [f"{sid}：見た目がストリームではありません"], nums
    d_want, ops_want = appearance(view, got if len(got) == 4 else want)
    apdict = {k: doc.resolve(v) for k, v in ap.dict.items() if k != "Length"}
    res = apdict.pop("Resources", None)
    if _plain(apdict) != _plain(d_want):
        probs.append(f"{sid}：見た目のフォームの設定（BBox・Matrix）が、決め直したものと違います")
    if doc.decode(ap) != ops_want or ap.dict.get("Filter") is not None:
        probs.append(f"{sid}：見た目の描き方が、画像を枠いっぱいに描くだけになっていません")
    xo = doc.resolve(res.get("XObject")) if isinstance(res, dict) else None
    if not (isinstance(res, dict) and set(res) == {"XObject"} and isinstance(xo, dict) and set(xo) == {XNAME}
            and isinstance(xo[XNAME], Ref)):
        return probs + [f"{sid}：見た目の資源が、画像1つではありません"], nums
    nums.add(xo[XNAME].num)
    probs += _check_image(doc, doc.resolve(xo[XNAME]), img, sid, nums)
    return probs, nums


def _plain(o):
    if isinstance(o, dict):
        return {k: _plain(v) for k, v in o.items()}
    if isinstance(o, list):
        return [round(float(v), 3) if isinstance(v, (int, float)) else _plain(v) for v in o]
    if isinstance(o, Name):
        return ("n", o.name)
    return o


def _check_image(doc, im, img: SealImage, sid: str, nums: set[int]) -> list[str]:
    """画像の XObject が、本人の画像と同じか：JPEG はファイルのバイト列そのもの、PNG は開いた画素とアルファ。"""
    if not isinstance(im, Stream):
        return [f"{sid}：画像の XObject がありません"]
    want, smask_want = image_objects(img)
    d = {k: doc.resolve(v) for k, v in im.dict.items() if k not in ("Length", "SMask")}
    probs = []
    if _plain(d) != _plain(want.dict):
        probs.append(f"{sid}：画像の設定（大きさ・色・圧縮）が、本人の画像と違います")
    data = doc.decode(im)  # 暗号を解き、Flate を開く（DCT は開かずにそのまま）
    if img.kind == "jpeg":
        if data != img.data:
            probs.append(f"{sid}：埋めた画像のバイト列が、本人の画像のファイルと一致しません")
    elif data != img.samples:
        probs.append(f"{sid}：埋めた画像の画素が、本人の画像を開いた画素と一致しません")
    sm = im.dict.get("SMask")
    if smask_want is None:
        if sm is not None:
            probs.append(f"{sid}：透けのない画像に、透け（SMask）が付いています")
        return probs
    if not isinstance(sm, Ref) or not isinstance(doc.resolve(sm), Stream):
        return probs + [f"{sid}：画像の透け（SMask）がありません"]
    nums.add(sm.num)
    s = doc.resolve(sm)
    if _plain({k: doc.resolve(v) for k, v in s.dict.items() if k != "Length"}) != _plain(smask_want.dict):
        probs.append(f"{sid}：透け（SMask）の設定が違います")
    if doc.decode(s) != img.alpha:
        probs.append(f"{sid}：透け（SMask）が、本人の画像のアルファと一致しません")
    return probs


def told(sd: dict, key: str) -> str:
    """本人に伝える文：何を、どこに置いたか。"""
    mm = sd.get("mm") or MM  # 1 mm のページの座標の長さ（UserUnit も反映）
    w = (sd["rect"][2] - sd["rect"][0]) / mm
    h = (sd["rect"][3] - sd["rect"][1]) / mm
    under = f"。画像の下に、ページの字「{sd['under']}」があります（画像は透けるので、字は見えます）" if sd.get("under") else ""
    return f"{sd['label']}に、本人の{KINDS[key]}の画像を置きました（横 {w:.0f} mm・縦 {h:.0f} mm）{under}"


__all__ = ["resolve", "write", "verify", "told", "caution", "stamp_names", "find_stamp", "SealError"]
