"""PDF：エージェントがこの書類を「全て正確に読み書きできるか」を、決定的な処理で確かめる。

読む: ページの描画の命令から、窓とは別の読み方（並べずに、文字の命令だけを順に）で全部の字を拾い、窓に1字残らず出ているかを確かめる。
      文字に直せない字・画像だけのページ・注釈の文字も確かめる。
書く: 窓が示したすべての入力欄に試しの値を一度に書き込み、元のバイト列がそのまま残っていること、
      変わったのが書いた欄とその見た目だけであること、修復なしに開けること、書いた値が窓に読み戻せることを確かめる。
"""

from __future__ import annotations

import re
import math
import unicodedata
from collections import Counter

from .. import chars
from ..xlsx.audit import Audit, Check
from .content import tokens
from .doc import PdfDocument
from .fill import FillError, fill
from .fonts import load_font
from .form import read_fields
from .objects import Name, PdfString, Stream, decode_text
from .proof import prove
from .window import build, markdown_parts

BOXES = set(chars.BOXES)  # 窓で ☐・☑ にして示す箱の字（chars.py：決まる箱と、箇条の印かもしれない箱）


def raw_chars(doc: PdfDocument, page: dict) -> list[str]:
    """ページの文字の命令を、並べずに順に読む（窓とは別の読み方）。ActualText の範囲は、その文字にする。"""
    from ..limits import LimitExceeded

    from .pagedraw import verified_refs

    out: list[str] = []
    drawing: set = set()
    written = verified_refs(doc, page)  # とどくが書いた値（欄の値として読む）：窓と同じく、印刷された字として数えない

    def run(data: bytes, res: dict, depth: int):
        if depth > 12:
            return
        fonts = doc.resolve(res.get("Font")) or {}
        xobjs = doc.resolve(res.get("XObject")) or {}
        props = doc.resolve(res.get("Properties")) or {}
        font = None
        stack: list = []
        nest: list[list] = []  # 配列・辞書の入れ子（["[", 中身] / ["<<", 中身]）
        mc: list = []  # (ActualText, 文字の数)
        saved: list = []  # q / Q で保存・復元する書体
        egs = doc.resolve(res.get("ExtGState")) or {}
        for kind, val in tokens(data):
            if kind in ("[", "<<"):
                nest.append([kind, []])
                continue
            if kind in ("]", ">>"):
                if not nest:
                    continue
                k0, items = nest.pop()
                if k0 == "<<":
                    obj = {items[i].name: items[i + 1] for i in range(0, len(items) - 1, 2) if isinstance(items[i], Name)}
                else:
                    obj = items
                (nest[-1][1] if nest else stack).append(obj)
                continue
            if kind != "o":
                (nest[-1][1] if nest else stack).append(val)
                continue
            op = val
            strings = []
            if op == "q":
                saved.append(font)
            elif op == "Q":
                if saved:
                    font = saved.pop()
            elif op == "Tf" and len(stack) >= 2 and isinstance(stack[-2], Name):
                ref = fonts.get(stack[-2].name) if isinstance(fonts, dict) else None
                font = load_font(doc, ref, stack[-2].name) if ref is not None else None
            elif op == "gs" and stack and isinstance(stack[-1], Name) and isinstance(egs, dict):
                g = doc.resolve(egs.get(stack[-1].name))
                fnt = doc.resolve(g.get("Font")) if isinstance(g, dict) else None
                if isinstance(fnt, list) and len(fnt) == 2:
                    font = load_font(doc, fnt[0], "gs")
            elif op in ("Tj", "'") and stack and isinstance(stack[-1], (bytes, PdfString)):
                strings = [stack[-1]]
            elif op == '"' and stack and isinstance(stack[-1], (bytes, PdfString)):
                strings = [stack[-1]]
            elif op == "TJ" and stack and isinstance(stack[-1], list):
                strings = [x for x in stack[-1] if isinstance(x, (bytes, PdfString))]
            elif op == "Do" and stack and isinstance(stack[-1], Name):
                xo = doc.resolve(xobjs.get(stack[-1].name)) if isinstance(xobjs, dict) else None
                if isinstance(xo, Stream) and getattr(xo.dict.get("Subtype"), "name", "") == "Form" \
                        and xo.num not in written and (xo.num or id(xo)) not in drawing:  # 自分を描く部品は、描いている途中にもう一度描かない
                    drawing.add(xo.num or id(xo))
                    try:
                        run(doc.decode(xo), doc.resolve(xo.dict.get("Resources")) or res, depth + 1)
                    except LimitExceeded:
                        raise
                    except Exception:  # noqa: BLE001
                        pass
                    finally:
                        drawing.discard(xo.num or id(xo))
            elif op in ("BDC", "BMC"):
                p = stack[-1] if op == "BDC" and stack else None
                if isinstance(p, Name):
                    p = doc.resolve(props.get(p.name)) if isinstance(props, dict) else None
                at = p.get("ActualText") if isinstance(p, dict) else None
                mc.append((decode_text(at) if isinstance(at, (bytes, PdfString)) else None, len(out)))
            elif op == "EMC" and mc:
                at, start = mc.pop()
                if at is not None and len(out) > start:
                    del out[start:]
                    out.append(at)
            if font is not None:
                for s in strings:
                    for _c, t, *_ in font.decode(bytes(s)):
                        if t:
                            out.append(t)
            stack.clear()

    run(doc.page_content(page), doc.resolve(page.get("Resources")) or {}, 0)
    return out


def _count(texts) -> Counter:
    """字を数える（正準分解して数える：窓で分けて描いた母音を1字にした「ো」は、元の「ে」「া」と同じに数える）。"""
    c = Counter()
    for t in texts:
        for ch in unicodedata.normalize("NFD", t):
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


def overprints(glyphs) -> Counter:
    """重ね書き（同じ字を同じ所に重ねて描いたもの）の、余分な字の数（字ごと）。
    これは「窓と食い違わないか」の照らし合わせ（整合の検査）で、独立の検査ではない：決まりは窓と同じ pagetext.is_overprint
    （同じ字で、送りの向きにそった2つの字の送りの幅の重なりが、狭い方の幅の半分より多く、直角のずれが字の大きさの 1/4 より
    小さく、大きさがほぼ同じもの）をそのまま使い（決まりを 2 か所に書かない）、窓とは別の手順で数え直す：窓の位置の換算（view）を
    通さずにページの座標で、字の幅と大きさから決まる届く範囲の字（重ね書きになりうる字はすべて）と比べる。決まりそのものの正しさは、ここでは確かめない。
    決まりとは別の量（描いた字の箱の重なり）で見て、箱が重ならない字を1字にした所は、窓の読む段階の注意（「重ね書きとして1字に
    した字」C）で知らせる。"""
    from .pagetext import glyph_frame, is_overprint

    out: Counter = Counter()
    kept = _KeptFrames()
    for g in glyphs:
        if not g.text or g.size < 0.5:
            continue
        fr = glyph_frame(g)
        if any(is_overprint(f0, fr, s0, g.size) for f0, s0 in kept.near(g.text, fr, g.size)):
            out[g.text] += 1
        else:
            kept.add(g.text, fr, g.size)
    return out


def redrawn(hidden, glyphs) -> tuple[Counter, list]:
    """塗りの下に隠れた字（hidden）のうち、見える字（glyphs の、見えない描き方でない字）として同じ所に描き直してある字の数
    （字ごと）と、残りの字。overprints と同じく、窓（pagetext.drop_redrawn）と同じ決まり is_overprint を、別の手順で数え直す。"""
    from .pagetext import glyph_frame, is_overprint

    texts = {g.text for g in hidden if g.text and g.size >= 0.5}
    kept = _KeptFrames()
    for g in glyphs:
        if g.text in texts and not g.invisible and g.size >= 0.5:
            fr = glyph_frame(g)
            if math.isfinite(fr[0] + fr[1] + fr[4]):
                kept.add(g.text, fr, g.size)
    out: Counter = Counter()
    rest = []
    for g in hidden:
        if g.text and g.size >= 0.5:
            fr = glyph_frame(g)
            if math.isfinite(fr[0] + fr[1] + fr[4]) and any(is_overprint(fr, f0, g.size, s0) for f0, s0 in kept.near(g.text, fr, g.size)):
                out[g.text] += 1
                continue
        rest.append(g)
    return out, rest


class _KeptFrames:
    """残した字を、字ごとに、幅と大きさの桁（2 の冪）で分け、桁ごとの大きさの升目に置く。重ね書きの決まり（is_overprint）では、
    2 つの字の原点の離れは「広い方の送りの幅 + 大きさの 1/4」までなので、届く升目だけを見れば、重ね書きになりうる字をすべて見る
    （同じ字が同じ行に何万と並ぶページでも、字ごとに全部の字と比べない）。"""

    def __init__(self):
        self.by_text: dict[str, dict[tuple[int, int], dict[tuple[int, int], list]]] = {}

    @staticmethod
    def _class(fr, size) -> tuple[int, int]:
        return (math.frexp(abs(fr[4]))[1], math.frexp(size)[1])

    @staticmethod
    def _reach(wc: int, sc: int) -> float:
        return 2.0 ** wc + 0.5 * 2.0 ** sc  # この桁の字の、届く長さの上限（大きさは、2 つの字の大きさが近いので 2 倍まで見る）

    def add(self, text: str, fr, size: float) -> None:
        wc, sc = self._class(fr, size)
        cell = self._reach(wc, sc)
        key = (math.floor(fr[0] / cell), math.floor(fr[1] / cell)) if math.isfinite(fr[0] + fr[1]) else (0, 0)
        self.by_text.setdefault(text, {}).setdefault((wc, sc), {}).setdefault(key, []).append((fr, size))

    def near(self, text: str, fr, size: float):
        for (wc, sc), grid in self.by_text.get(text, {}).items():
            cell = self._reach(wc, sc)
            reach = max(cell, abs(fr[4]) + 0.5 * size)  # この字の幅のほうが広いこともある
            if not math.isfinite(fr[0] + fr[1] + reach):
                for items in grid.values():
                    yield from items
                continue
            r = math.ceil(reach / cell)
            cx, cy = math.floor(fr[0] / cell), math.floor(fr[1] / cell)
            if (2 * r + 1) ** 2 > len(grid):
                for items in grid.values():
                    yield from items
                continue
            for x in range(cx - r, cx + r + 1):
                for y in range(cy - r, cy + r + 1):
                    yield from grid.get((x, y), ())


def check_reading(doc: PdfDocument, win: dict) -> list[Check]:
    return [check_glyphs(doc, win)] + _check_unmapped_and_images(doc, win)


def check_glyphs(doc: PdfDocument, win: dict) -> Check:
    """読み落とし：ページの描画の命令を順に読んだ字（raw_chars）を、窓に実際に出た字（入力欄の印に置き換えた字は数えない）と
    比べる。窓が1字にした字は、理由ごとに数えて差し引く：重ね書き（overprints が、窓と同じ決まりを窓とは別の手順で数え直す。
    整合の検査）と、0.5pt に満たない字・ページの何倍もある字。窓が、数え直した重ね書きより多くの字を「重ね書き」として落としたら、
    読み落とし。数え直せないページがあれば、通ったことにしない（「確かめられない」）。"""
    from .content import PageContent, page_content

    missing, extra = [], []
    dup_total = 0
    unsure = []
    for pg, page in zip(win["pages"], doc.pages()):
        raw = _count(raw_chars(doc, page))
        shown = _count(pg["_plain"]) - _count(pg.get("_added") or [])  # 窓で足した字（ルビのかっこ）は除く
        tiny = _count(t for k, t in pg["_dropped"] if k == "size")
        merged = _count(t for k, t in pg["_dropped"] if k == "dup")
        try:
            # 画面に見えないレイヤーの字は、窓がレイヤーごとに分けて読む（重ね書きも、その中で 1 字にする）ので、同じ分け方で数え直す
            pc = page_content(doc, page)
            real = Counter()
            for c in [pc] + [c for _st, c in pc.off] + ([PageContent(glyphs=pc.clipped)] if pc.clipped else []):
                real += _count(ch for t, n in overprints(c.glyphs).items() for ch in [t] * n)
            if pc.covered:
                # 塗りの下に隠れた字：見える字として同じ所に描き直してある字は、重ね書きとして数え、残りの中の重ね書きを数える
                again, rest = redrawn(pc.covered, pc.glyphs)
                real += _count(ch for t, n in again.items() for ch in [t] * n)
                real += _count(ch for t, n in overprints(rest).items() for ch in [t] * n)
        except Exception as e:  # noqa: BLE001  数え直せないものは、窓の数を正しいとしない
            unsure.append(f"ページ {pg['number']}: 重ね書きを数え直せませんでした（{type(e).__name__}: {str(e)[:60]}）")
            continue
        dup_total += sum(real.values())
        for ch, n in (merged - real).items():
            missing.append(f"ページ {pg['number']}: 「{ch}」が {n} 字、重ね書きとして1字にされましたが、同じ所に重ねた字ではありません")
        for ch, n in (real - merged).items():
            extra.append(f"ページ {pg['number']}: 「{ch}」が {n} 字、同じ所に重ねて描いた字なのに、別の字として窓に出ています")
        want = raw - tiny - real
        for ch, n in (want - shown).items():
            missing.append(f"ページ {pg['number']}: 「{ch}」が {n} 字、窓に出ていません")
        for ch, n in (shown - raw).items():
            extra.append(f"ページ {pg['number']}: 「{ch}」が {n} 字、元にないのに窓に出ています")
    total = sum(p["chars"] for p in win["pages"])
    if unsure:
        return Check("読み落とし（確かめられない）", False,
                     f"{len(unsure)} ページは、窓の読み落としを確かめられませんでした（重ね書きを数え直せない）", unsure + missing + extra)
    out = [Check("読み落とし（ページの文字）", not missing and not extra,
                 f"ページの文字 {total:,} 字を、窓とは別の読み方（文字の命令を順に）で拾い、すべて窓に出ていることを確かめました"
                 + (f"（同じ字を同じ所に重ねて描いた太字などの {dup_total:,} 字は、1字として読んでいます。重ね書きかは、窓と同じ"
                    "決まり（字の位置と送りの幅）を別の手順で数え直して照らしたもので、決まりそのものは確かめていません）" if dup_total else ""),
                 missing + extra)]
    return out[0]


def _check_unmapped_and_images(doc: PdfDocument, win: dict) -> list[Check]:
    out = []
    unmapped = sum(p["unmapped"] for p in win["pages"])
    out.append(Check("文字に直せない字", unmapped == 0,
                     "すべての字を文字（Unicode）に直せました" if not unmapped else
                     f"{unmapped} 字は、フォントに文字の対応がなく、文字に直せませんでした（窓では �）",
                     [f"ページ {p['number']}: {p['unmapped']} 字" for p in win["pages"] if p["unmapped"]]))
    flags = [(p["number"], fl) for p in win["pages"] for fl in p.get("read_flags") or []]
    hard = [f"{n} ページ：{fl['title']}（{'、'.join(fl['where'][:3])}）" for n, fl in flags if fl["grade"] in "CDEF"]
    soft = [f"{n} ページ：{fl['title']}（{'、'.join(fl['where'][:3])}）" for n, fl in flags if fl["grade"] not in "CDEF"]
    out.append(Check("窓の字を信用できない所", not hard,
                     "窓の字をそのまま信用できない所はありません" if not flags else
                     f"窓の字をそのまま信用できない所が {len(flags)} 種類あります（窓のページの頭に、位置と確かめ方を書いています）",
                     hard + soft, agent=bool(soft) and not hard))
    img = [p["number"] for p in win["pages"] if p.get("image_only")]
    from .pageimage import unreadable_images

    pages = doc.pages()
    # 大きな PDF の検査は 1 ページずつ（doc はそのページだけ、番号は文書の中の番号）：窓のページと、同じ順で対応させる
    by_number = dict(zip((p["number"] for p in win["pages"]), pages))
    bad = {n: unreadable_images(doc, by_number[n]) for n in img}
    bad = {n: v for n, v in bad.items() if v}
    out.append(Check("画像だけのページ", not bad,
                     "すべてのページに文字の情報があります" if not img else
                     f"{len(img)} ページは画像だけで、文字の情報がありません。" +
                     (f"{len(img) - len(bad)} ページは、エージェントがページの画像（view_page）を見て読み、枠（mm）で書きます。" if len(bad) < len(img) else "") +
                     (f"{len(bad)} ページは画像の圧縮を開けません" if bad else ""),
                     [f"ページ {n}" + (f"（開けない：{'・'.join(bad[n])}）" if n in bad else "") for n in img], agent=bool(img) and not bad))
    return out


def _sample(f: dict):
    if f.get("source") == "xfa":
        return f.get("_sample")  # XFA の欄：欄の決まり（書式・桁・選択肢）に合う値
    k = f["kind"]
    if k == "checkbox":
        return not bool(f.get("value"))
    if k == "radio":
        opts = f.get("options") or []
        cur = f.get("value")
        for o in reversed(opts):
            if o != cur:
                return o
        return opts[0] if opts else None
    if k in ("combo", "list"):
        opts = f.get("options") or []
        cur = f.get("value")
        for o in opts:
            if o not in (cur, "", " "):
                return o
        return opts[0] if opts else None
    if k == "comb":
        return ("1234567890" * 4)[: f["boxes"]]
    fmt = f.get("format") or ""
    if fmt.startswith("日付"):
        return "2026-04-01"
    if fmt.startswith(("数", "百分率")):
        return "1"
    if fmt.startswith("郵便番号（5"):
        return "12345"
    if fmt.startswith("郵便番号（9"):
        return "123456789"
    if fmt.startswith("電話"):
        return "0312345678"
    if fmt.startswith("社会保障"):
        return "123456789"
    ml = f.get("maxlen")
    if f.get("comb") and ml:
        return ("1234567890" * 4)[:ml]
    if ml and ml < 3:
        return "1" * ml
    # 欄の大きさに入る値（小さな欄は1字。書く値が入りきらないときは、書き込みが「入りきりません」と断る）
    rect = f.get("rect", (0, 0, 100, 20))
    from .fill import spot_fits

    if f.get("source") == "layout":
        cur = f.get("value")
        if f.get("filled") and isinstance(cur, str) and cur.strip():
            # 書いた後の書類の所：もう欄があるので、書き直しはその欄の決まり（1 行）で。その値を書き直す（入ることが分かっている値）
            return cur
        # 配置から見つけた欄：書き込み（add_widget）と同じ決め方で入る値
        for cand in (("12", "1") if k == "inline" else ("テスト", "12", "1")):
            if spot_fits(cand, rect, f.get("size"), k, f):
                return cand
        return "1"
    cur = f.get("value")
    if isinstance(cur, str) and cur.strip() and not ml:
        # いま値のある欄（書いた後の書類の欄など）：その値を書き直す（その欄に入ることが分かっている値）
        return cur
    x0, y0, x1, y1 = rect
    from .fill import FillError, layout_text

    for cand in (("12", "1") if k == "inline" else ("テスト", "12", "1")):
        try:
            # 窓の矩形は 0.1pt に丸めてあるので、0.5pt 小さい枠で確かめる（書き込みは、欄の本当の大きさで決める）
            layout_text(cand, x1 - x0 - 0.5, y1 - y0 - 0.5, min(float(f.get("size") or 10.5), 10.5), multiline=(y1 - y0) > 25)
            return cand
        except FillError:
            continue
    return "1"


def _all_fields(win: dict) -> list[dict]:
    """ページの入力欄と、ページに置かれていない XFA の入力欄（動的なフォームの欄など）。"""
    out = [f for pg in win["pages"] for f in pg["fields"]]
    seen = {f["id"] for f in out}
    out += [f for f in (win.get("xfa") or {}).get("fields", []) if f["id"] not in seen]
    return out


_NUM_OR_OP = re.compile(rb"([-+]?(?:\d+\.?\d*|\.\d+))|([A-Za-z*]+)")


def raw_circles(doc: PdfDocument, page: dict) -> dict[str, tuple[str, tuple[float, float, float, float]]]:
    """ページの描画の命令をそのまま読み、印 TdkCircle（とどくが描いた○）の中の曲線の点の外枠を、ページの座標
    （ページの初めの座標。○は Q で初めの座標に戻してから描く）で返す：id → (形 "ellipse"・"border", (x0, y0, x1, y1))。
    楕円は 4 つのベジエ曲線で描くので、制御点の外枠が楕円の外枠。窓の座標の換算（View）も circle.drawn も通さない。"""
    data = doc.page_content(page)
    out = {}
    for m in re.finditer(rb"/TdkCircle\s*\(([^)]*)\)(.*?)>>\s*BDC(.*?)EMC", data, re.S):
        nums: list[float] = []
        pts: list[tuple[float, float]] = []
        for t in _NUM_OR_OP.finditer(m[3]):
            if t[1] is not None:
                nums.append(float(t[1]))
                continue
            if t[2] == b"m" and len(nums) >= 2:
                pts.append((nums[-2], nums[-1]))
            elif t[2] == b"c" and len(nums) >= 6:
                pts += [(nums[-6], nums[-5]), (nums[-4], nums[-3]), (nums[-2], nums[-1])]
            nums = []
        if len(pts) < 13:
            continue
        xs, ys = [x for x, _y in pts], [y for _x, y in pts]
        out[m[1].decode("latin-1")] = ("border" if b"/TdkStyle /border" in m[2] else "ellipse",
                                       (min(xs), min(ys), max(xs), max(ys)))
    return out


def raw_circled_text(glyphs, kind: str, box) -> str:
    """ページの字（描画の命令の順、ページの座標）のうち、箱の中心が○の内側にある字（同じ字を同じ所に重ねたものは 1 字）。
    窓の決め方（circle.encloses：字の箱がすっかり内側）とは別の量（箱の中心）で数える。"""
    x0, y0, x1, y1 = box
    cx, cy, a, b = (x0 + x1) / 2, (y0 + y1) / 2, max((x1 - x0) / 2, 1e-6), max((y1 - y0) / 2, 1e-6)
    seen = set()
    out = []
    for g in glyphs:
        if g.invisible or not g.text.strip():
            continue
        gx, gy = g.cx, g.cy
        ok = (x0 <= gx <= x1 and y0 <= gy <= y1) if kind == "border" else ((gx - cx) / a) ** 2 + ((gy - cy) / b) ** 2 <= 1.0
        key = (g.text, round(g.ox, 1), round(g.oy, 1))
        if ok and key not in seen:
            seen.add(key)
            out.append(g.text)
    return "".join(out)


def check_circles(doc: PdfDocument) -> Check | None:
    """書いてある○（とどくが描いた、印 TdkCircle のもの）：試しに書き直さず（選んだものを変えない）、いまの○が囲んでいる字を、
    書くときと同じ決め方（circle.encloses）で読み、窓がその字を「［p1.c1：○］」として出しているか（いまの値を読めるか）を確かめる。
    これは窓と書くときが同じ決め方で同じ配置を見ていることの照らし合わせ（整合の検査）で、決め方そのものの正しさは確かめない。
    そのため別に、描画の命令から直接、○の外枠（ページの座標）と字の箱の中心で数え直し（raw_circles・raw_circled_text：窓の
    座標の換算・行の組み立て・encloses を通さない。字の位置の読み取り content.py は共通）、窓の囲んだ字と同じかを見る。違えば、
    どちらが正しいかは画像で確かめる所として示す（NG にはしない）。
    ○の線がほかの字の面にかかっているもの（前の版で書いた○など）は、値は読めるので NG にはせず、位置を添えて
    エージェントが画像で確かめる所として示す。○のない書類は None。"""
    from . import circle
    from .content import PageContent, page_content
    from .layout import PageLayout

    bad, touching, differ, n_all = [], [], [], 0
    for n, pg in enumerate(doc.pages(), 1):
        shapes = circle.drawn(doc, n)
        if not shapes:
            continue
        pc = page_content(doc, pg)
        lay = PageLayout(doc, pg, pc, n)
        shown = {ident: ln.text[a_:b_] for ln, a_, b_, ident in getattr(lay, "circled", []) if ident}
        raw = raw_circles(doc, pg)
        for ident, kind, g, lw in shapes:
            n_all += 1
            inner, touch = circle.circled_text(lay, kind, g, lw, True)  # 字の墨を切るものだけ（縁に触れるだけは知らせて描く）
            if not inner or ident not in shown:
                bad.append(f"{ident}：囲んでいる字「{''.join(c.text for c in inner)}」・窓に囲んだ字として出ていない"
                           "（2 つの行にかかる・囲んだ字の間に囲まれていない字がある）")
                continue
            if touch:
                touching.append(f"{ident}「{shown[ident]}」：○の線がほかの字「{''.join(c.text for c in touch[:6])}」にかかる"
                                f"（{n} ページ。view_page で確かめる）")
            want = sorted(re.sub(r"\s", "", shown[ident]))
            if ident not in raw:
                differ.append(f"{ident}「{shown[ident]}」：描画の命令から○の形を読み直せない（{n} ページ。view_page で確かめる）")
            else:
                got = raw_circled_text(pc.glyphs, *raw[ident])
                if sorted(got) != want:
                    differ.append(f"{ident}「{shown[ident]}」：描画の命令から数え直すと、○の内側に中心がある字は「{got}」"
                                  f"（{n} ページ。view_page で確かめる）")
    if not n_all:
        return None
    detail = (f"書いてある○ {n_all} 個を、試しに書き直さずに読みました（いまの値）。窓と書くときが同じ決め方（circle.encloses）で"
              "同じ配置を見ているかの照らし合わせ（整合の検査）で、決め方そのものの正しさは確かめません。") + (
        "どれも、囲んだ字を窓に出しています" if not bad else f"{len(bad)} 個は、囲んだ字を窓に出せません")
    if not bad:
        detail += ("。別に、描画の命令から直接（○の外枠と字の箱の中心で）数え直し、" +
                   ("どれも窓の囲んだ字と同じでした" if not differ else
                    f"{len(differ)} 個は窓の囲んだ字と違いました（どちらが正しいかは画像で確かめる）"))
    if touching:
        detail += f"。{len(touching)} 個は○の線がほかの字にかかります（値は読めます。どれを囲んだかは画像で確かめる）"
    return Check("書いてある○", not bad, detail, (bad + differ + touching)[:30], agent=bool(touching or differ) and not bad)


def check_writing(doc: PdfDocument, win: dict, password: str | None = None) -> list[Check]:
    targets: dict[str, object] = {}
    skipped, not_writable = [], []
    for f in _all_fields(win):
        if f["id"] in targets:
            continue
        if f.get("readonly"):
            not_writable.append(f"{f['id']}（{f.get('why_not') or '読み取り専用'}）")
            continue
        if f["kind"] in ("signature", "button"):
            not_writable.append(f"{f['id']}（{'電子署名' if f['kind'] == 'signature' else '押すだけのボタン'}）")
            continue
        if any(o in targets for o in f.get("shared_with") or []):
            continue  # 同じデータを共有する欄（XFA の global）は、1つだけ書く
        v = _sample(f)
        if v is None:
            skipped.append(f"{f['id']}: 試しの値を決められませんでした")
            continue
        targets[f["id"]] = v
    n_fields = len(_all_fields(win))
    passes = _passes(targets, {f["id"]: f for f in _all_fields(win)})
    na = (f"（読み取り専用・電子署名・ボタン・データに結びつかない欄など {len(not_writable)} 欄は、書けない欄なので除外）"
          if not_writable else "")
    many = (f"同じ所の別の読み方（重なる候補）は両方には書けないので、{len(passes)} 回に分けて書き込み、" if len(passes) > 1 else "")
    unexpected, wrong, heads = [], [], []
    formatted = 0
    ok_all = True
    for part in passes:
        try:
            # 試しの値をすべての欄に書けるか（欄の作りの確かめ）：書いた字の重なり（placement）は、本当に書くときに確かめる
            res = fill(doc, part, placement=False)
        except FillError as e:
            return [Check("全入力欄への書き込み", False, f"{len(part)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
        anchored = {p_.id: p_.anchor for p_ in res.patches if p_.anchor}
        pr = prove(doc.blob, res.blob, set(part), anchored or None, password=password)
        ok_all = ok_all and pr.ok
        unexpected += [*pr.unexpected, *pr.lint_errors]
        heads.append(pr.headline())
        after = PdfDocument(res.blob, doc.filename, password=password)
        win2 = build(after)
        got = {f["id"]: f.get("value") for f in _all_fields(win2)}
        written = {p_.id: p_.value for p_ in res.patches}
        for k, v in part.items():
            g = got.get(k)
            w = written.get(k, v)
            if not isinstance(v, bool) and str(w) != str(v):
                formatted += 1  # 欄の書式（日付・数など）を通した値
            v = w if not isinstance(v, bool) else v
            if isinstance(v, bool):
                ok = bool(g) == v
            else:
                ok = str(g) == str(v)
            if not ok:
                wrong.append(f"{k}: 書いた値 {v!r} → 窓では {g!r}")
    head = heads[0] if len(set(heads)) == 1 else " / ".join(dict.fromkeys(heads))
    out = [Check("全入力欄への書き込み", ok_all and not skipped,
                 f"入力欄 {n_fields:,} か所のうち {len(targets):,} に試しの値を{many}書き込み{na}。{head}",
                 [*unexpected, *skipped])]
    note = f"（うち {formatted} 欄は、欄の書式どおりの形にして書いた値）" if formatted else ""
    out.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている{note}", wrong))
    return out


def _passes(targets: dict, by_id: dict) -> list[dict]:
    """書き込みの試しを、同じ所の別の読み方（fill.conflicting）が同じ回に入らないように分ける（たいていは1回）。"""
    from .fill import conflicting

    clash: dict[str, set] = {}
    groups: dict[str, list[dict]] = {}
    for k in targets:
        f = by_id.get(k)
        alt = f.get("alternative_to") if f is not None and f.get("source") == "layout" else None
        if alt:
            clash.setdefault(k, set()).add(alt)
            clash.setdefault(alt, set()).add(k)
            groups.setdefault(alt, []).append(f)
    for g in groups.values():
        for i, a in enumerate(g):
            for b in g[i + 1:]:
                if conflicting(a, b):
                    clash[a["id"]].add(b["id"])
                    clash[b["id"]].add(a["id"])
    passes: list[dict] = []
    for k, v in targets.items():
        for part in passes:
            if not clash.get(k, set()) & part.keys():
                part[k] = v
                break
        else:
            passes.append({k: v})
    return passes or [{}]


def anchored_check(blob: bytes, filename: str, password: str | None = None) -> Check:
    """見出しの字を目印にした書き込み（窓に欄として示していない所に、エージェントが書く）：ページの中に1つしかない字の右の
    空いている所に試しの字を書き、証明（目印から決め直した位置と、足した欄の位置が同じ）と読み戻しを確かめる。"""
    import re

    from .anchor import AnchorError, _lines, resolve_write
    from .fill import FillError, fill
    from .form import read_fields
    from .proof import prove

    doc = PdfDocument(blob, filename, password=password)
    names = {f.name for f in read_fields(doc)}
    layouts: dict = {}
    tried = 0
    for n in range(1, len(doc.pages()) + 1):
        try:
            resolve_write(doc, {"after": "\x00", "value": "", "page": n}, layouts=layouts)
        except AnchorError:
            pass  # ページの配置を読んでおく（見つからないのは当然）
        lay = layouts.get(n)
        if lay is None:
            continue
        texts = [ln.text for _c, ln in _lines(lay)]
        for _cell, ln in _lines(lay):
            core = ln.text.rstrip(" \u3000")
            m = re.search(r"\S{3,}$", core)
            if not m or sum(t.count(m[0]) for t in texts) != 1:
                continue
            spec = {"after": m[0], "value": "（試し）", "page": n}
            try:
                sd = resolve_write(doc, spec, layouts=layouts, taken_names=names)
            except AnchorError:
                continue
            r = sd["rect"]
            if r[2] - r[0] < 5 * sd["size"] or r[3] - r[1] < 0.9 * sd["size"]:
                continue  # 試しの字「（試し）」（4字）が、ふつうの大きさで入る空きを選ぶ
            tried += 1
            try:
                res = fill(PdfDocument(blob, filename, password=password), {"@1": spec}, placement=False)
            except FillError as e:
                return Check("字を目印にした書き込み", False, f"{n} ページの「{m[0]}」の右に書けませんでした", [str(e)])
            pr = prove(blob, res.blob, {sd["id"]}, {sd["id"]: spec}, password=password)
            got = {f["id"]: f.get("value") for pg in build(PdfDocument(res.blob, filename, password=password))["pages"]
                   for f in pg["fields"]}
            ok = pr.ok and got.get(sd["id"]) == "（試し）"
            return Check("字を目印にした書き込み", ok,
                         f"窓に欄として示していない所への書き込み（{n} ページの「{m[0]}」の右に「（試し）」）を {sd['id']} として書き、"
                         f"{pr.headline()}。読み戻し：{'窓に出ている' if got.get(sd['id']) == '（試し）' else '窓に出ていない'}",
                         [*pr.unexpected, *pr.lint_errors])
            if tried >= 5:
                break
    return Check("字を目印にした書き込み", True, "右に空きのある、ページの中に1つしかない字がないので、試していません")


def run(blob: bytes, filename: str = "", password: str | None = None) -> Audit:
    """password：開くのにパスワードが要る PDF のとき（本人に聞いたもの。結果には出さない）。"""
    doc = PdfDocument(blob, filename, password=password)
    win = build(doc)
    au = Audit(filename)
    au.checks += check_reading(doc, win)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    total = sum(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {total:,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                        if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(PdfDocument(blob, filename, password=password), win, password)
    cc = check_circles(doc)
    if cc is not None:
        au.checks.append(cc)
    from ..xlsx.audit import agent_places_check

    ap = agent_places_check([f for p_ in win["pages"] for f in p_["fields"]])
    if ap is not None:
        au.checks.append(ap)
    if (win.get("xfa") or {}).get("readable"):
        au.checks.append(Check("字を目印にした書き込み", True, "XFA のフォームでは、字を目印にした書き込みはしません"
                               "（XFA の欄でない所に書いても Adobe では表示されないことがあるので、XFA の欄の id を指定して書きます）"))
    elif not any(p_.get("image_only") for p_ in win["pages"]) or any(p_.get("fields") for p_ in win["pages"]):
        au.checks.append(anchored_check(blob, filename, password))
    fm = win["form"]
    scripts = _scripts(doc)
    xw = win.get("xfa")
    if xw and xw.get("readable"):
        au.checks += check_xfa(doc, win)
        auto = [s_ for s_ in xw["scripts"] if s_["automatic"]]
        au.checks.append(Check("フォームの仕掛け", True,
                               f"XFA のフォームを読みました（{'動的' if xw['kind'] == 'dynamic' else '静的'}なフォーム）。"
                               + (f"XFA のスクリプト {len(xw['scripts'])} 種類（自動で動くもの {len(auto)}）は実行しません。"
                                  "ソースを窓に出したので、値や欄の出し入れを変えるものかを、エージェントが読んで判断します"
                                  if xw["scripts"] else "スクリプトはありません"),
                               [" / ".join(s_["where"][:2]) for s_ in auto[:20]], agent=bool(auto)))
        return au
    au.checks.append(Check("フォームの仕掛け", not fm["xfa"] and not scripts,
                           "XFA（Adobe の独自のフォーム）も、独自のスクリプトもありません（Acrobat の標準の書式・計算は、同じ規則で行います）"
                           if not fm["xfa"] and not scripts else
                           ("XFA を含みます（Adobe Acrobat 以外では、PDF の欄の値が使われます）。" if fm["xfa"] else "")
                           + (f"独自のスクリプトが {len(scripts)} 個あります（実行しません。その欄は、書いた後に確かめてください）"
                              if scripts else ""), scripts))
    return au


def _norm_ws(t: str) -> str:
    return " ".join(t.split())


def check_xfa(doc: PdfDocument, win: dict) -> list[Check]:
    """XFA：template と datasets の字（文字の節・属性の値）を、窓とは別の読み手（xml.etree）で拾い、すべて窓に出ているかを確かめる。
    静的なフォームでは、PDF の欄の名前が、すべて XFA の欄の名前（id）になっているかも確かめる。"""
    import xml.etree.ElementTree as ET

    from .xfa import Xdp, repair_xml

    xw = win["xfa"]
    xdp = Xdp(doc)
    fields = xw["fields"]
    shown = list(xw["_plain"])
    for f in fields:
        for k in ("label", "tooltip", "value", "display"):
            if isinstance(f.get(k), str):
                shown.append(f[k])
        if isinstance(f.get("default"), str):
            shown.append(f["default"])
        for k in ("options", "option_labels", "export_values", "rules", "value", "assist"):
            if isinstance(f.get(k), list):
                shown += [str(v) for v in f[k]]
        for k in ("on_value", "off_value"):
            if f.get(k):
                shown.append(f[k])
    window_text = _norm_ws("\n".join(shown))
    atoms = {_norm_ws(t) for t in shown}
    scripts = {_norm_ws(s_) for s_ in xw["_scripts"]}
    lists_text = _norm_ws("\n".join(xw["lists"]))
    missing: list[str] = []
    images = 0
    total = 0

    def ok_text(t: str, where: str) -> bool:
        return t in atoms or t in window_text or (where == "lists" and t in lists_text)

    def walk(e, in_script: bool, kind: str):
        nonlocal images, total
        tag = e.tag.rsplit("}", 1)[-1] if isinstance(e.tag, str) else ""
        if tag == "image":
            images += 1
            return
        script = in_script or tag == "script"
        texts = [e.text or ""] + [c.tail or "" for c in e]
        if kind == "data" or kind == "lists":
            texts += [v for k_, v in e.attrib.items() if "xmlns" not in k_ and "xfa-data" not in k_ and "XMLSchema" not in k_]
        for t in texts:
            t = _norm_ws(t)
            if not t:
                continue
            total += 1
            if script and kind == "template":
                if _norm_ws(e.text or "") not in scripts:
                    missing.append(f"スクリプト（{tag}）：{t[:60]}")
                break
            if not ok_text(t, kind):
                missing.append(f"{'template' if kind == 'template' else 'datasets'} の {tag}：「{t[:60]}」")
        for c in e:
            if isinstance(c.tag, str):
                walk(c, script, kind)

    def et(el):
        data = xdp.combined[el.s:el.e]
        fixed, _ = repair_xml(data)
        try:
            return ET.fromstring(fixed)
        except ET.ParseError:
            return ET.fromstring(b'<w xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/">' + fixed + b"</w>")[0]

    t_el = xdp.parts.get("template")
    if t_el is not None:
        walk(et(t_el), False, "template")
    ds = xdp.parts.get("datasets")
    if ds is not None:
        for c in et(ds):
            if not isinstance(c.tag, str) or "data-description" in c.tag:
                continue
            walk(c, False, "data" if c.tag.endswith("}data") else "lists")
    out = [Check("読み落とし（XFA の字）", not missing,
                 f"template と datasets の字 {total:,} 個（文字の節と、データの属性の値）を、窓とは別の読み手で拾い、"
                 + ("すべて窓に出ていることを確かめました" if not missing else f"{len(missing)} 個が窓に出ていません")
                 + f"（画像 {images} 個は字ではないので除く）", missing[:30])]
    if xw["kind"] == "static":
        from .form import read_fields

        ids = {f["id"] for f in fields}
        names = [f.name for f in read_fields(doc)]
        # ページの共通部分（pageArea）は、ページごとに同じものが PDF の欄になる（Page1[1] など）：0 番目に置き換えて照らす
        base = [re.sub(r"(#pageSet\[\d+\]\.[^.]+)\[\d+\]", r"\1[0]", n) for n in names]
        bad = [n for n, b in zip(names, base) if n not in ids and b not in ids]
        out.append(Check("XFA の名前と PDF の欄", not bad,
                         f"PDF の欄 {len(names)} 個の名前が、すべて XFA の欄の id（名前の式）と同じです" if not bad else
                         f"PDF の欄 {len(bad)} 個の名前が、XFA の欄の id と合いません", bad[:30]))
    return out


def _scripts(doc: PdfDocument) -> list[str]:
    """Acrobat の標準の関数（日付・数・計算など）でない、独自のスクリプト。"""
    from . import acrojs

    out = []
    for f in read_fields(doc):
        node = doc.resolve(f.ref) if f.ref is not None else None
        for k, js in acrojs.actions(doc, node).items():
            if not acrojs.known(js):
                out.append(f"{f.name}（{ {'C': '計算', 'F': '表示の書式', 'K': '入力の制限', 'V': '値の確認'}[k] }: {js[:60]}）")
    return out
