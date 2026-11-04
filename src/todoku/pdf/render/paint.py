"""ページの描画の命令を順に実行して、画素に描く（ISO 32000-1 8〜9 章・11・12.5.5）。

描くもの：図形（塗り・線・切り抜き）、色（8.6）、画像（8.9）、フォームの部品（8.10、BBox で切り抜く）、
字（9.4 の字の位置の計算と、字の形：埋め込んだ書体・代わりの書体・型 3 の字の手続き）、字の描き方（Tr：塗り・線・見えない・
切り抜き）、網かけ（8.7.4.5 の型 2 軸・型 3 放射。sh と模様の塗り）、タイルの模様（8.7.3）、透け（ExtGState の CA・ca）、
見える／見えない内容（8.11 Optional Content：既定の設定で OFF の内容は描かない）、注釈の見た目（12.5.5、/AP の N）。
線の調整（10.6.5）：水平・垂直の細い線と、塗りの細い四角は、縁を画素の境に合わせる（見る道具と同じ。罫線が薄くならない）。
描けないもの（型 1・4〜7 の網かけ、透けのマスク SMask、重ね方 BM）は、近い形で描いて数を伝える（黙って捨てない）。読めない命令・
見つからない書体と部品・開けない画像・見た目（/AP）のない欄も数えて伝える（pageimage.not_drawn）。
"""

from __future__ import annotations

import math
import time

from ..content import tokens
from ..fonts import load_font
from ..objects import Name, PdfString, Stream
from .canvas import (Canvas, Clip, _rect_of, fill_rows, flatten, inv, mul, paint, paint_fn, rect_rows, scale_of, scan,
                     stroke_polys, transform_polys)
from .colors import DCMYK, DRGB, GRAY, PatternCS, colorspace, function
from .fontmap import _decoded_name, glyphs_for
from .image import draw_image, image_rgba

IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


class _Stop(Exception):
    pass


class GS:
    __slots__ = ("ctm", "clip", "fcs", "fc", "fpat", "scs", "sc", "spat", "lw", "cap", "join", "miter", "dash", "phase",
                 "ca", "CA", "font", "pf", "fs", "tc", "tw", "th", "tl", "tr", "rise", "base", "nocolor", "smask")

    def __init__(self, ctm, clip):
        self.ctm = ctm
        self.clip = clip
        self.fcs, self.fc, self.fpat = GRAY, (0, 0, 0), None
        self.scs, self.sc, self.spat = GRAY, (0, 0, 0), None
        self.lw, self.cap, self.join, self.miter = 1.0, 0, 0, 10.0
        self.dash, self.phase = None, 0.0
        self.ca = self.CA = 1.0
        self.font = None
        self.pf = None
        self.fs, self.tc, self.tw, self.th, self.tl, self.tr, self.rise = 0.0, 0.0, 0.0, 1.0, 0.0, 0, 0.0
        self.base = ctm  # 模様の座標の基（ページ・フォームの部品の座標）
        self.nocolor = False  # 型 3 の d1：色の命令を使わない
        self.smask = False

    def copy(self):
        g = GS.__new__(GS)
        for k in GS.__slots__:
            setattr(g, k, getattr(self, k))
        return g


def _rgb255(c) -> tuple[int, int, int]:
    return (int(c[0] * 255 + 0.5), int(c[1] * 255 + 0.5), int(c[2] * 255 + 0.5))


class Renderer:
    def __init__(self, doc, cv: Canvas, deadline: float | None = None):
        self.doc = doc
        self.cv = cv
        self.deadline = deadline
        self.stopped = False
        self.ops = 0
        self.gcache: dict = {}
        self.drawing: set = set()
        self.stats = {"glyphs": 0, "glyphs_embedded": 0, "glyphs_system": 0, "glyphs_type3": 0, "glyphs_box": 0,
                      "images": 0, "paths": 0}
        self.subst: dict[str, str] = {}  # PDF の書体の名前 → 代わりに描いた書体
        self.missing_fonts: set[str] = set()
        self.box_places: list = []  # 字の形が分からない字の位置（画素。はじめの数十個）
        self.problems: list[str] = []
        self.approx: dict[str, int] = {}
        self.image_failed: list[str] = []
        self.image_notes: list[str] = []  # 画像の開き方の知らせ（JPEG 2000 の解像度など）
        self._oc_off = self._oc_config()

    # ------------------------------------------------------------------ 見える／見えない内容
    def _oc_config(self) -> set:
        doc = self.doc
        try:
            ocp = doc.resolve(doc.root.get("OCProperties"))
        except Exception:  # noqa: BLE001
            return set()
        if not isinstance(ocp, dict):
            return set()
        d = doc.resolve(ocp.get("D")) or {}
        off: set = set()
        allg = [x for x in (doc.resolve(ocp.get("OCGs")) or [])]
        key = lambda r: getattr(r, "num", id(r))  # noqa: E731
        if getattr(doc.resolve(d.get("BaseState")), "name", "ON") == "OFF":
            on = {key(x) for x in (doc.resolve(d.get("ON")) or [])}
            off = {key(x) for x in allg if key(x) not in on}
        for x in doc.resolve(d.get("OFF")) or []:
            off.add(key(x))
        return off

    def hidden(self, oc) -> bool:
        """内容の印（OCG・OCMD）が、既定の設定で見えないか。"""
        if oc is None or not self._oc_off:
            return False
        doc = self.doc
        key = getattr(oc, "num", None)
        d = doc.resolve(oc)
        if not isinstance(d, dict):
            return False
        t = getattr(d.get("Type"), "name", "")
        if t == "OCMD":
            gs = d.get("OCGs")
            gs = doc.resolve(gs)
            lst = gs if isinstance(gs, list) else [d.get("OCGs")] if gs is not None else []
            states = [getattr(g, "num", id(g)) not in self._oc_off for g in lst]
            if not states:
                return False
            p = getattr(d.get("P"), "name", "AnyOn")
            vis = {"AllOn": all(states), "AnyOn": any(states), "AnyOff": not all(states), "AllOff": not any(states)}.get(p, any(states))
            return not vis
        return key in self._oc_off

    # ------------------------------------------------------------------
    def note(self, what: str) -> None:
        self.approx[what] = self.approx.get(what, 0) + 1

    def run(self, data: bytes, res: dict, gs: GS, depth: int = 0) -> None:
        try:
            self._run(data, res, gs, depth)
        except _Stop:
            if depth:
                raise

    def _run(self, data: bytes, res: dict, gs: GS, depth: int) -> None:
        if depth > 16:
            self.problems.append("部品が深く重なりすぎています（その先は描きません）")
            return
        from ...limits import MAX_PAGE_OPS, LimitExceeded

        doc = self.doc
        res = res if isinstance(res, dict) else {}
        stack: list = []
        arrs: list = []
        dicts: list = []
        gstack: list = []
        path: list = []
        cur = (0.0, 0.0)
        start = (0.0, 0.0)
        pending_clip = None
        tm = tlm = IDENT
        text_clip: list | None = None
        hide: list = []  # 内容の印ごとに、見えない（True）か
        inline = None

        def push(v):
            if arrs:
                arrs[-1].append(v)
            elif dicts:
                dicts[-1].append(v)
            else:
                stack.append(v)

        def nums(k):
            return [float(x) for x in stack[-k:]]

        for kind, val in tokens(data):
            if kind != "o":
                if kind == "[":
                    arrs.append([])
                elif kind == "]":
                    if arrs:
                        push(arrs.pop())
                elif kind == "<<":
                    dicts.append([])
                elif kind == ">>":
                    if dicts:
                        items = dicts.pop()
                        push({items[i].name: items[i + 1] for i in range(0, len(items) - 1, 2) if isinstance(items[i], Name)})
                elif kind == "I":
                    d = {}
                    for i in range(0, len(stack) - 1, 2):
                        if isinstance(stack[i], Name):
                            d[stack[i].name] = stack[i + 1]
                    inline = (d, val)
                    stack.clear()
                else:
                    push(val)
                continue
            op = val
            self.ops += 1
            if self.ops > MAX_PAGE_OPS:
                raise LimitExceeded(f"1 ページの描画の字句が {MAX_PAGE_OPS:,} 個を超えます")
            if self.ops % 2000 == 0 and self.deadline is not None and time.monotonic() > self.deadline:
                self.stopped = True
                raise _Stop()
            invisible = bool(hide) and any(hide)
            try:
                # ---------------------------------------------------- 図形
                if op == "m":
                    x, y = nums(2)
                    path.append((0, x, y))
                    cur = start = (x, y)
                elif op == "l":
                    x, y = nums(2)
                    if not path:
                        path.append((0, *cur))
                    path.append((1, x, y))
                    cur = (x, y)
                elif op == "c":
                    v = nums(6)
                    if not path:
                        path.append((0, *cur))
                    path.append((2, *v))
                    cur = (v[4], v[5])
                elif op == "v":
                    v = nums(4)
                    if not path:
                        path.append((0, *cur))
                    path.append((2, cur[0], cur[1], *v))
                    cur = (v[2], v[3])
                elif op == "y":
                    v = nums(4)
                    if not path:
                        path.append((0, *cur))
                    path.append((2, v[0], v[1], v[2], v[3], v[2], v[3]))
                    cur = (v[2], v[3])
                elif op == "h":
                    if path:
                        path.append((4,))
                        cur = start
                elif op == "re":
                    x, y, w, h = nums(4)
                    path += [(0, x, y), (1, x + w, y), (1, x + w, y + h), (1, x, y + h), (4,)]
                    cur = start = (x, y)
                elif op in ("S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n"):
                    if op in ("s", "b", "b*") and path:
                        path.append((4,))
                    if path and not invisible and op != "n":
                        fill = op in ("f", "F", "f*", "B", "B*", "b", "b*")
                        stroke = op in ("S", "s", "B", "B*", "b", "b*")
                        eo = op in ("f*", "B*", "b*")
                        if fill:
                            self.fill_path(gs, path, eo)
                        if stroke:
                            self.stroke_path(gs, path)
                    if pending_clip is not None and path:
                        self.clip_path(gs, path, pending_clip)
                    pending_clip = None
                    path = []
                elif op in ("W", "W*"):
                    pending_clip = op == "W*"
                # ---------------------------------------------------- 状態
                elif op == "q":
                    gstack.append(gs.copy())
                    if len(gstack) > 4000:
                        raise ValueError("q が多すぎます")
                elif op == "Q":
                    if gstack:
                        gs = gstack.pop()
                elif op == "cm":
                    gs.ctm = mul(tuple(nums(6)), gs.ctm)
                elif op == "w":
                    gs.lw = float(stack[-1])
                elif op == "J":
                    gs.cap = int(stack[-1])
                elif op == "j":
                    gs.join = int(stack[-1])
                elif op == "M":
                    gs.miter = float(stack[-1])
                elif op == "d":
                    arr = stack[-2] if len(stack) >= 2 else []
                    gs.dash = [abs(float(x)) for x in arr] if isinstance(arr, list) and arr else None
                    gs.phase = float(stack[-1])
                elif op == "gs":
                    self.ext_gstate(gs, res, stack[-1])
                # ---------------------------------------------------- 色
                elif op in ("g", "G", "rg", "RG", "k", "K"):
                    if not gs.nocolor:
                        cs = GRAY if op in ("g", "G") else DRGB if op in ("rg", "RG") else DCMYK
                        c = _rgb255(cs.rgb(nums(cs.n)))
                        if op.islower():
                            gs.fcs, gs.fc, gs.fpat = cs, c, None
                        else:
                            gs.scs, gs.sc, gs.spat = cs, c, None
                elif op in ("cs", "CS"):
                    if not gs.nocolor:
                        cs = colorspace(doc, stack[-1], res)
                        c = (0, 0, 0) if isinstance(cs, PatternCS) else _rgb255(cs.rgb(cs.initial()))
                        if op == "cs":
                            gs.fcs, gs.fc, gs.fpat = cs, c, None
                        else:
                            gs.scs, gs.sc, gs.spat = cs, c, None
                elif op in ("sc", "scn", "SC", "SCN"):
                    if not gs.nocolor:
                        fill = op in ("sc", "scn")
                        cs = gs.fcs if fill else gs.scs
                        vals = [float(x) for x in stack if isinstance(x, (int, float))]
                        pat = stack[-1] if stack and isinstance(stack[-1], Name) else None
                        if isinstance(cs, PatternCS):
                            c = _rgb255(cs.rgb(vals)) if cs.under and vals else (0, 0, 0)
                            p = (pat, res, vals) if pat is not None else None
                        else:
                            c = _rgb255(cs.rgb(vals[-cs.n:] if cs.n else vals))
                            p = None
                        if fill:
                            gs.fc, gs.fpat = c, p
                        else:
                            gs.sc, gs.spat = c, p
                # ---------------------------------------------------- 字
                elif op == "BT":
                    tm = tlm = IDENT
                    text_clip = [] if gs.tr >= 4 else None
                elif op == "ET":
                    if text_clip is not None and gs.tr >= 4:
                        gs.clip = gs.clip.with_rows(scan(text_clip, False, gs.clip.rect)) if text_clip else Clip(0, 0, 0, 0, {})
                    text_clip = None
                elif op == "Tf":
                    gs.fs = float(stack[-1])
                    nm = stack[-2]
                    gs.font, gs.pf = self.font(res, nm.name if isinstance(nm, Name) else "")
                elif op == "Tc":
                    gs.tc = float(stack[-1])
                elif op == "Tw":
                    gs.tw = float(stack[-1])
                elif op == "Tz":
                    gs.th = float(stack[-1]) / 100.0
                elif op == "TL":
                    gs.tl = float(stack[-1])
                elif op == "Tr":
                    gs.tr = int(stack[-1])
                    if gs.tr >= 4 and text_clip is None:
                        text_clip = []
                elif op == "Ts":
                    gs.rise = float(stack[-1])
                elif op == "Td":
                    tx, ty = nums(2)
                    tlm = tm = mul((1, 0, 0, 1, tx, ty), tlm)
                elif op == "TD":
                    tx, ty = nums(2)
                    gs.tl = -ty
                    tlm = tm = mul((1, 0, 0, 1, tx, ty), tlm)
                elif op == "Tm":
                    tlm = tm = tuple(nums(6))
                elif op == "T*":
                    tlm = tm = mul((1, 0, 0, 1, 0, -gs.tl), tlm)
                elif op in ("Tj", "'", '"', "TJ"):
                    if op == "'":
                        tlm = tm = mul((1, 0, 0, 1, 0, -gs.tl), tlm)
                    elif op == '"':
                        gs.tw, gs.tc = float(stack[-3]), float(stack[-2])
                        tlm = tm = mul((1, 0, 0, 1, 0, -gs.tl), tlm)
                    arg = stack[-1] if stack else None
                    items = arg if op == "TJ" and isinstance(arg, list) else [arg]
                    for it in items:
                        if isinstance(it, (bytes, PdfString)):
                            tm = self.show(gs, res, it, tm, invisible, text_clip, depth)
                        elif isinstance(it, (int, float)) and gs.pf is not None:
                            adj = -float(it) / 1000.0 * gs.fs
                            tm = mul((1, 0, 0, 1, 0, adj) if gs.pf.vertical else (1, 0, 0, 1, adj * gs.th, 0), tm)
                elif op in ("d0", "d1"):
                    if op == "d1":
                        gs.nocolor = True
                # ---------------------------------------------------- 部品・画像・網かけ
                elif op == "Do":
                    if not invisible and stack and isinstance(stack[-1], Name):
                        self.xobject(gs, res, stack[-1].name, depth)
                elif op == "EI":
                    if not invisible and inline is not None:
                        self.image(gs, inline, res, "埋め込みの画像")
                    inline = None
                elif op == "sh":
                    if not invisible and stack and isinstance(stack[-1], Name):
                        sh = doc.resolve((doc.resolve(res.get("Shading")) or {}).get(stack[-1].name))
                        if sh is not None:
                            self.shade(gs, sh, None, res)
                # ---------------------------------------------------- 内容の印
                elif op in ("BDC", "BMC"):
                    hid = False
                    if op == "BDC" and len(stack) >= 2 and getattr(stack[-2], "name", "") == "OC":
                        props = stack[-1]
                        oc = (doc.resolve(res.get("Properties")) or {}).get(props.name) if isinstance(props, Name) else None
                        hid = self.hidden(oc)
                    hide.append(hid)
                elif op == "EMC":
                    if hide:
                        hide.pop()
            except _Stop:
                raise
            except LimitExceeded:
                raise
            except (IndexError, ValueError, TypeError, ZeroDivisionError, KeyError, AttributeError, OverflowError):
                self.note(f"読めない描画の命令（{op}）")  # その命令だけ飛ばして続ける（黙って捨てない）
            stack.clear()

    # ------------------------------------------------------------------ 状態
    def ext_gstate(self, gs: GS, res: dict, nm) -> None:
        doc = self.doc
        if not isinstance(nm, Name):
            return
        eg = doc.resolve((doc.resolve(res.get("ExtGState")) or {}).get(nm.name))
        if not isinstance(eg, dict):
            return
        for k, v in eg.items():
            v = doc.resolve(v)
            if k == "LW":
                gs.lw = float(v)
            elif k == "LC":
                gs.cap = int(v)
            elif k == "LJ":
                gs.join = int(v)
            elif k == "ML":
                gs.miter = float(v)
            elif k == "D" and isinstance(v, list) and len(v) == 2:
                arr = doc.resolve(v[0]) or []
                gs.dash = [abs(float(doc.resolve(x))) for x in arr] or None
                gs.phase = float(doc.resolve(v[1]))
            elif k == "CA":
                gs.CA = max(0.0, min(1.0, float(v)))
            elif k == "ca":
                gs.ca = max(0.0, min(1.0, float(v)))
            elif k == "Font" and isinstance(v, list) and len(v) == 2:
                pf = load_font(doc, v[0], "gs")
                fd = doc.resolve(v[0])
                gs.pf = pf
                gs.font = glyphs_for(doc, fd, pf) if pf is not None and isinstance(fd, dict) else None
                gs.fs = float(doc.resolve(v[1]))
            elif k == "SMask":
                if not (isinstance(v, Name) and v.name == "None"):
                    self.note("透けのマスク（SMask）")
            elif k == "BM":
                bm = v[0] if isinstance(v, list) and v else v
                if getattr(bm, "name", "Normal") not in ("Normal", "Compatible"):
                    self.note("重ね方（BM）")

    def font(self, res: dict, name: str):
        doc = self.doc
        fonts = doc.resolve(res.get("Font")) or {}
        ref = fonts.get(name) if isinstance(fonts, dict) else None
        if ref is None:
            return None, None
        pf = load_font(doc, ref, name)
        fd = doc.resolve(ref)
        if pf is None or not isinstance(fd, dict):
            return None, None
        g = glyphs_for(doc, fd, pf)
        if g.kind == "system":
            self.subst[_decoded_name(pf.base or name)] = g.subst_name
        elif g.kind == "none":
            self.missing_fonts.add(_decoded_name(pf.base or name))
        return g, pf

    # ------------------------------------------------------------------ 塗り・線・切り抜き
    def _paint_rows(self, gs: GS, rows, fill: bool) -> None:
        pat = gs.fpat if fill else gs.spat
        alpha = gs.ca if fill else gs.CA
        if pat is not None:
            self.pattern_rows(gs, rows, pat, alpha, fill)
            return
        paint(self.cv, rows, gs.fc if fill else gs.sc, alpha, gs.clip)

    def fill_path(self, gs: GS, path: list, eo: bool) -> None:
        if gs.clip.empty:
            return
        self.stats["paths"] += 1
        subs = flatten(path, gs.ctm)
        polys = [_snap_thin_rect(p) for p, _c in subs if len(p) >= 2]
        if not polys:
            return
        rows = fill_rows(polys, eo, gs.clip.rect, separate_rects=gs.ca >= 0.999 and gs.fpat is None and not eo)
        self._paint_rows(gs, rows, True)

    def stroke_path(self, gs: GS, path: list) -> None:
        if gs.clip.empty:
            return
        self.stats["paths"] += 1
        m = gs.ctm
        sc = scale_of(m)
        wdev = gs.lw * sc
        aligned = _axis_uniform(m)
        if wdev < 1.0 or aligned:
            # 1 画素より細い線（太さ 0 を含む）は、画面の座標で 1 画素の線（8.4.3.2）。軸にそろった行列の線は、画面の座標で作って
            # 画素にそろえる（線の調整 10.6.5）
            subs = flatten(path, m)
            if gs.dash:
                subs = [(p, c) for p, c in _dash_user(path, m, gs)]
            w = max(1.0, wdev)
            polys = _adjusted_strokes(subs, w, gs.cap) if aligned else None
            if polys is None:
                polys = stroke_polys(subs, w, gs.cap, gs.join, gs.miter, None, 0.0, 1.0)
        else:
            subs = flatten(path, m, keep=True)
            polys = stroke_polys(subs, gs.lw, gs.cap, gs.join, gs.miter, gs.dash, gs.phase, sc)
            polys = transform_polys(polys, m)
        if not polys:
            return
        rows = fill_rows(polys, False, gs.clip.rect, separate_rects=gs.CA >= 0.999 and gs.spat is None)
        self._paint_rows(gs, rows, False)

    def clip_path(self, gs: GS, path: list, eo: bool) -> None:
        subs = flatten(path, gs.ctm)
        polys = [p for p, _c in subs if len(p) >= 2]
        if not polys:
            gs.clip = Clip(0, 0, 0, 0, {})
            return
        if len(polys) == 1:
            r = _rect_of(polys[0])
            if r is not None:
                gs.clip = gs.clip.with_rect(*r)
                return
        gs.clip = gs.clip.with_rows(scan(polys, eo, gs.clip.rect))

    # ------------------------------------------------------------------ 字
    def show(self, gs: GS, res: dict, s: bytes, tm, invisible: bool, text_clip, depth: int):
        pf, gl = gs.pf, gs.font
        if pf is None:
            if not invisible and gs.tr not in (3, 7):
                self.note("書体が見つからず描いていない字の並び")
            return tm
        fs, th = gs.fs, gs.th
        mode = gs.tr
        for code, text, w, is_space, vpos, _vw in pf.decode(s):
            trm = mul((fs * th, 0, 0, fs, 0, gs.rise), mul(tm, gs.ctm))
            if pf.vertical:
                vx, vy = vpos
                gm = mul((1, 0, 0, 1, -vx, -vy), trm)
                ty = w * fs + gs.tc + (gs.tw if is_space else 0)
                tm = mul((1, 0, 0, 1, 0, ty), tm)
            else:
                gm = trm
                tx = (w * fs + gs.tc + (gs.tw if is_space else 0)) * th
                tm = mul((1, 0, 0, 1, tx, 0), tm)
            if invisible or (mode in (3, 7) and text_clip is None) or gl is None:
                continue
            self.stats["glyphs"] += 1
            self.glyph(gs, res, gl, pf, code, text, gm, mode, text_clip, w, depth)
        return tm

    def glyph(self, gs: GS, res: dict, gl, pf, code: bytes, text: str, m, mode: int, text_clip, w: float, depth: int):
        r = gl.lookup(code, text)
        if r is None and text and text != "�" and text.strip():
            # 字（ToUnicode など）は分かるが、字の形がない：このパソコンのどれかの書体で描く（字が読めるように）
            from .fontmap import any_face_glyph

            r = any_face_glyph(text)
            if r is not None:
                self.stats["glyphs_fallback"] = self.stats.get("glyphs_fallback", 0) + 1
        if r is None:
            # 字の形が分からない字（どの書体にも字がない・埋め込みでない書体で字が決まらない）：灰色の四角で示して数え、位置を伝える。
            # 埋め込んだ書体で字の決まらない符号は、見る道具でも .notdef（ふつうは何も描かない）
            unknown = text == "�" and gl.kind in ("system", "none")
            if (text and text.strip() and text != "�") or unknown or (gl.kind == "none" and code not in (b" ",)):
                if mode not in (3, 7):
                    self.stats["glyphs_box"] += 1
                    if len(self.box_places) < 40:
                        self.box_places.append((m[4], m[5]))
                    self._box(gs, pf, m, w)
            return
        key, ops, gmat = r
        if key == "type3":
            self.stats["glyphs_type3"] += 1
            if mode in (3, 7):
                return
            fm = pf.matrix if hasattr(pf, "matrix") else (0.001, 0, 0, 0.001, 0, 0)
            g2 = gs.copy()
            g2.ctm = mul(fm, m)
            g2.base = g2.ctm
            fres = self.doc.resolve(gl.fd.get("Resources")) or res
            k = ("t3", id(ops))
            if k in self.drawing:
                return
            self.drawing.add(k)
            try:
                self._run(self.doc.decode(ops), fres, g2, depth + 1)
            finally:
                self.drawing.discard(k)
            return
        self.stats["glyphs_system" if gl.kind == "system" else "glyphs_embedded"] += 1
        if not ops:
            return
        M = mul(gmat, m)
        if text_clip is not None and mode >= 4:
            text_clip += [p for p, _c in flatten(ops, M) if len(p) >= 2]
        if mode in (3, 7) or gs.clip.empty:
            return
        if mode in (0, 2, 4, 6):
            sz = math.sqrt(abs(M[0] * M[3] - M[1] * M[2])) * 1000
            if sz < 400 and gs.fpat is None:
                ix, iy = math.floor(M[4]), math.floor(M[5])
                qx, qy = int((M[4] - ix) * 4) / 4, int((M[5] - iy) * 4) / 4
                ck = (id(gl), key, round(M[0], 4), round(M[1], 4), round(M[2], 4), round(M[3], 4), qx, qy)
                rows = self.gcache.get(ck)
                if rows is None:
                    rows = scan([p for p, _c in flatten(ops, (M[0], M[1], M[2], M[3], qx, qy)) if len(p) >= 2])
                    if len(self.gcache) < 30000:
                        self.gcache[ck] = rows
                paint(self.cv, rows, gs.fc, gs.ca, gs.clip, ix, iy)
            else:
                rows = scan([p for p, _c in flatten(ops, M) if len(p) >= 2], False, gs.clip.rect)
                self._paint_rows(gs, rows, True)
        if mode in (1, 2, 5, 6):
            subs = flatten(ops, M)
            wdev = max(1.0, gs.lw * scale_of(gs.ctm))
            polys = stroke_polys(subs, wdev, gs.cap, gs.join, gs.miter, None, 0.0, 1.0)
            self._paint_rows(gs, fill_rows(polys, False, gs.clip.rect), False)

    def _box(self, gs: GS, pf, m, w: float) -> None:
        """字の形が分からない字：字の四角（送りの幅 × 字の高さ）を灰色の枠で示す。"""
        x1 = w if w > 0 else 0.5
        d, a = getattr(pf, "descent", -0.15), getattr(pf, "ascent", 0.85)
        quad = [(0.0, d), (x1, d), (x1, a), (0.0, a)]
        pts = [(m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]) for x, y in quad]
        polys = stroke_polys([(pts, True)], 1.0, 0, 0, 10.0, None, 0.0, 1.0)
        paint(self.cv, fill_rows(polys, False, gs.clip.rect), (150, 150, 150), 1.0, gs.clip)

    # ------------------------------------------------------------------ 部品・画像
    def xobject(self, gs: GS, res: dict, name: str, depth: int) -> None:
        doc = self.doc
        xo = doc.resolve((doc.resolve(res.get("XObject")) or {}).get(name))
        if not isinstance(xo, Stream):
            self.note(f"見つからない部品（{name}）")
            return
        if self.hidden(xo.dict.get("OC")):
            return
        st = getattr(xo.dict.get("Subtype"), "name", "")
        if st == "Image":
            self.image(gs, xo, res, name)
        elif st == "Form":
            key = xo.num or id(xo)
            if key in self.drawing:
                self.problems.append(f"フォームの部品 {name} が自分自身を描いています")
                return
            self.form(gs, xo, res, depth)
        else:
            self.note(f"描けない部品（{st or name}）")

    def form(self, gs: GS, xo: Stream, res: dict, depth: int, matrix=None) -> None:
        doc = self.doc
        key = xo.num or id(xo)
        m = xo.dict.get("Matrix")
        m = tuple(float(doc.resolve(v)) for v in doc.resolve(m)) if isinstance(doc.resolve(m), list) and len(doc.resolve(m)) == 6 else IDENT
        g2 = gs.copy()
        g2.ctm = mul(m, matrix if matrix is not None else gs.ctm)
        g2.base = g2.ctm
        bb = doc.resolve(xo.dict.get("BBox"))
        if isinstance(bb, list) and len(bb) == 4:
            x0, y0, x1, y1 = (float(doc.resolve(v)) for v in bb)
            self.clip_path(g2, [(0, x0, y0), (1, x1, y0), (1, x1, y1), (1, x0, y1), (4,)], False)
        if g2.clip.empty:
            return
        sub_res = doc.resolve(xo.dict.get("Resources")) or res
        self.drawing.add(key)
        try:
            self._run(doc.decode(xo), sub_res, g2, depth + 1)
        except (_Stop, MemoryError, RecursionError):
            raise
        except Exception as e:  # noqa: BLE001
            from ...limits import LimitExceeded

            if isinstance(e, LimitExceeded):
                raise
            self.problems.append(f"フォームの部品を描けません（{type(e).__name__}）")
        finally:
            self.drawing.discard(key)

    def image(self, gs: GS, img, res: dict, name: str) -> None:
        if gs.clip.empty:
            return
        m = gs.ctm
        # 画面での大きさ（画素）：大きすぎる画像は、この大きさの数倍に間引いて開く
        dev = abs(m[0] * m[3] - m[1] * m[2])
        # JPEG は逆 DCT で 1/2・1/4・1/8 に縮めて開ける（縮めても面積の平均）：画面の 2 倍まで。ほかは開いてから箱の平均で縮める
        try:
            im = image_rgba(self.doc, img, 16_000_000, res, int(max(250_000, dev * 2)), self.deadline)
        except Exception as e:  # noqa: BLE001
            from ...limits import LimitExceeded

            if isinstance(e, LimitExceeded):
                raise
            self.image_failed.append((name, str(e), _unit_square_box(m)))
            self._undrawn_image(gs, m)
            return
        self.stats["images"] += 1
        self.image_notes += [f"{name}：{n}" for n in im.notes if f"{name}：{n}" not in self.image_notes]
        self.problems += [f"{name}：{n}" for n in im.problems if f"{name}：{n}" not in self.problems]
        fill = gs.fc
        if im.stencil and gs.fpat is not None:
            fill = (128, 128, 128)
            self.note("模様で塗る型抜きの画像")
        draw_image(self.cv, im, m, gs.clip, gs.ca, fill)

    def _undrawn_image(self, gs: GS, m) -> None:
        """開けない画像：その四角に灰色の枠と対角線を描く（白いままにしない：そこに中身があると見て分かるように）。"""
        pts = [(m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]) for x, y in ((0, 0), (1, 0), (1, 1), (0, 1))]
        subs = [(pts, True), ([pts[0], pts[2]], False), ([pts[1], pts[3]], False)]
        polys = stroke_polys(subs, 1.0, 0, 0, 10.0, None, 0.0, 1.0)
        if polys:
            paint(self.cv, fill_rows(polys, False, gs.clip.rect), (150, 150, 150), 1.0, gs.clip)

    # ------------------------------------------------------------------ 網かけ・模様
    def shade(self, gs: GS, sh, rows, res: dict, matrix=None, alpha: float | None = None) -> None:
        """網かけ（8.7.4.5）。rows：塗る所（None なら切り抜きの全体）。matrix：網かけの座標 → 画面。"""
        doc = self.doc
        d = sh.dict if isinstance(sh, Stream) else sh
        if not isinstance(d, dict):
            return
        t = int(doc.resolve(d.get("ShadingType")) or 0)
        cs = colorspace(doc, d.get("ColorSpace"), res)
        m = matrix if matrix is not None else gs.ctm
        clip = gs.clip
        if rows is None:
            bb = doc.resolve(d.get("BBox"))
            if isinstance(bb, list) and len(bb) == 4:
                x0, y0, x1, y1 = (float(doc.resolve(v)) for v in bb)
                g2 = gs.copy()
                g2.ctm = m
                self.clip_path(g2, [(0, x0, y0), (1, x1, y0), (1, x1, y1), (1, x0, y1), (4,)], False)
                clip = g2.clip
            rows = rect_rows(clip.x0, clip.y0, clip.x1, clip.y1, *clip.rect)
        a = gs.ca if alpha is None else alpha
        im = inv(m)
        if im is None:
            return
        try:
            fn = function(doc, d.get("Function")) if d.get("Function") is not None else None
        except Exception:  # noqa: BLE001
            fn = None
        if t in (2, 3) and fn is not None:
            coords = [float(doc.resolve(v)) for v in doc.resolve(d.get("Coords")) or []]
            dom = [float(doc.resolve(v)) for v in (doc.resolve(d.get("Domain")) or [0, 1])]
            ext = [bool(doc.resolve(v)) for v in (doc.resolve(d.get("Extend")) or [False, False])]
            t0, t1 = dom[0], dom[1]
            lut = []
            for i in range(257):
                tt = t0 + (t1 - t0) * i / 256
                try:
                    lut.append(_rgb255(cs.rgb(fn([tt]))))
                except Exception:  # noqa: BLE001
                    lut.append((128, 128, 128))
            ia, ib, ic, id_, ie, if_ = im
            if t == 2 and len(coords) >= 4:
                xa, ya, xb, yb = coords[:4]
                dx, dy = xb - xa, yb - ya
                den = dx * dx + dy * dy or 1e-12

                def at(X, Y):
                    x, y = ia * X + ic * Y + ie, ib * X + id_ * Y + if_
                    s = ((x - xa) * dx + (y - ya) * dy) / den
                    if s < 0:
                        if not ext[0]:
                            return None
                        s = 0.0
                    elif s > 1:
                        if not ext[1]:
                            return None
                        s = 1.0
                    return lut[int(s * 256)]
            elif t == 3 and len(coords) >= 6:
                x0, y0, r0, x1, y1, r1 = coords[:6]
                cdx, cdy, dr = x1 - x0, y1 - y0, r1 - r0
                A = cdx * cdx + cdy * cdy - dr * dr

                def at(X, Y):
                    x, y = ia * X + ic * Y + ie, ib * X + id_ * Y + if_
                    px, py = x - x0, y - y0
                    B = px * cdx + py * cdy + r0 * dr
                    C = px * px + py * py - r0 * r0
                    if abs(A) < 1e-12:
                        if abs(B) < 1e-12:
                            return None
                        cands = [C / (2 * B)]
                    else:
                        disc = B * B - A * C
                        if disc < 0:
                            return None
                        sq = math.sqrt(disc)
                        cands = sorted([(B + sq) / A, (B - sq) / A], reverse=True)
                    for s in cands:
                        if r0 + s * dr < 0:
                            continue
                        if s < 0:
                            if not ext[0]:
                                continue
                            s = 0.0
                        elif s > 1:
                            if not ext[1]:
                                continue
                            s = 1.0
                        return lut[int(s * 256)]
                    return None
            else:
                self.note("網かけ")
                return
            paint_fn(self.cv, rows, at, a, clip)
            return
        # 描けない型：関数の中ほどの色か、Background で塗る
        self.note(f"網かけ（型 {t}）")
        col = None
        bg = doc.resolve(d.get("Background"))
        if isinstance(bg, list):
            col = _rgb255(cs.rgb([float(doc.resolve(v)) for v in bg]))
        elif fn is not None:
            try:
                col = _rgb255(cs.rgb(fn([0.5] if t != 1 else [0.5, 0.5])))
            except Exception:  # noqa: BLE001
                col = None
        if col is not None:
            paint(self.cv, rows, col, a, clip)

    def pattern_rows(self, gs: GS, rows, pat, alpha: float, fill: bool) -> None:
        doc = self.doc
        nm, res, vals = pat
        p = doc.resolve((doc.resolve(res.get("Pattern")) or {}).get(nm.name))
        d = p.dict if isinstance(p, Stream) else p
        if not isinstance(d, dict):
            self.note(f"見つからない模様（{nm.name}）")
            return
        pm = d.get("Matrix")
        pm = tuple(float(doc.resolve(v)) for v in doc.resolve(pm)) if isinstance(doc.resolve(pm), list) else IDENT
        m = mul(pm, gs.base)
        pt = int(doc.resolve(d.get("PatternType")) or 0)
        if pt == 2:
            sh = doc.resolve(d.get("Shading"))
            if sh is not None:
                self.shade(gs, sh, rows, res, m, alpha)
            return
        if pt == 1 and isinstance(p, Stream):
            self.tiling(gs, p, d, rows, m, alpha, fill, vals)

    def tiling(self, gs: GS, p: Stream, d: dict, rows, m, alpha: float, fill: bool, vals) -> None:
        doc = self.doc
        bb = [float(doc.resolve(v)) for v in (doc.resolve(d.get("BBox")) or [0, 0, 1, 1])]
        xs = float(doc.resolve(d.get("XStep")) or (bb[2] - bb[0]) or 1)
        ys = float(doc.resolve(d.get("YStep")) or (bb[3] - bb[1]) or 1)
        paint_type = int(doc.resolve(d.get("PaintType")) or 1)
        if not rows:
            return
        clip = gs.clip.with_rows(rows)
        if clip.empty:
            return
        im = inv(m)
        if im is None or not xs or not ys:
            return
        corners = [(clip.x0, clip.y0), (clip.x1, clip.y0), (clip.x0, clip.y1), (clip.x1, clip.y1)]
        uv = [(im[0] * x + im[2] * y + im[4], im[1] * x + im[3] * y + im[5]) for x, y in corners]
        i0 = math.floor((min(u for u, _ in uv) - bb[2]) / abs(xs))
        i1 = math.ceil((max(u for u, _ in uv) - bb[0]) / abs(xs))
        j0 = math.floor((min(v for _, v in uv) - bb[3]) / abs(ys))
        j1 = math.ceil((max(v for _, v in uv) - bb[1]) / abs(ys))
        count = (i1 - i0 + 1) * (j1 - j0 + 1)
        cell_px = scale_of(m) * max(abs(xs), abs(ys))
        if count > 4000 or cell_px < 1.5:
            # 細かすぎる模様：同じ色の網とみなし、半分の濃さの色で塗る
            self.note("細かい模様")
            c = gs.fc if fill else gs.sc
            if paint_type == 1:
                c = (128, 128, 128)
            paint(self.cv, rows, c, alpha * 0.5, gs.clip)
            return
        data = doc.decode(p)
        pres = doc.resolve(d.get("Resources")) or {}
        key = ("pat", p.num or id(p))
        if key in self.drawing:
            return
        self.drawing.add(key)
        try:
            for j in range(j0, j1 + 1):
                for i in range(i0, i1 + 1):
                    g2 = GS(mul((1, 0, 0, 1, i * xs, j * ys), m), clip)
                    g2.base = g2.ctm
                    g2.ca = g2.CA = alpha
                    if paint_type == 2:
                        g2.fc = g2.sc = gs.fc if fill else gs.sc
                        g2.nocolor = True
                    self.clip_path(g2, [(0, bb[0], bb[1]), (1, bb[2], bb[1]), (1, bb[2], bb[3]), (1, bb[0], bb[3]), (4,)], False)
                    if g2.clip.empty:
                        continue
                    self._run(data, pres, g2, 8)
        finally:
            self.drawing.discard(key)


def _axis_uniform(m) -> bool:
    """行列が、軸にそろって（回転は 90 度ごと）縦横の拡大が同じか。"""
    a, b, c, d = (abs(v) for v in m[:4])
    sx, sy = max(a, b), max(c, d)
    if sx <= 0 or sy <= 0 or abs(sx - sy) > 1e-3 * max(sx, sy):
        return False
    return (b < 1e-6 * sx and c < 1e-6 * sy) or (a < 1e-6 * sx and d < 1e-6 * sy)


def _snap_span(a: float, b: float) -> tuple[float, float]:
    """縁を画素の境に合わせる（最低 1 画素）。"""
    a, b = float(round(a)), float(round(b))
    return (a, b) if b > a else (a, a + 1.0)


def _adjusted_strokes(subs, w: float, cap: int) -> list | None:
    """線の調整（ISO 32000-1 10.6.5、ページを見る道具がふつう行う）：水平・垂直の線だけの図形を、縁を画素の境に合わせた四角に
    する（細い罫線が 2 画素にまたがって薄くならない。ずれは半画素まで）。斜めや曲線のある図形は None（そのまま線にする）。
    角は四角どうしの重なりで埋まる（線の角の形は、この太さでは見分けられない）。丸い端は四角で描く。"""
    if w >= 3.0:  # 太い線は角・端の形が見えるので、ふつうの線で（縁はなめらか）
        return None
    eps = 1e-3
    hw = w / 2
    out: list = []
    for pts, closed in subs:
        ring = pts + [pts[0]] if closed and len(pts) > 1 else pts
        if len(ring) == 1 or all(abs(p[0] - ring[0][0]) <= eps and abs(p[1] - ring[0][1]) <= eps for p in ring):
            if cap != 0:  # 長さのない線：端の形だけ描く（8.4.3.3）
                x0, x1 = _snap_span(ring[0][0] - hw, ring[0][0] + hw)
                y0, y1 = _snap_span(ring[0][1] - hw, ring[0][1] + hw)
                out.append([(x0, y0), (x1, y0), (x1, y1), (x0, y1)])
            continue
        n = len(ring) - 1
        for i in range(n):
            (px, py), (qx, qy) = ring[i], ring[i + 1]
            dx, dy = abs(qx - px), abs(qy - py)
            if dx > eps and dy > eps:
                return None
            if dx <= eps and dy <= eps:
                continue
            # 端の伸ばし：角（前後に線が続く所）と、四角・丸の端は、太さの半分
            before = hw if (closed or i > 0 or cap != 0) else 0.0
            after = hw if (closed or i < n - 1 or cap != 0) else 0.0
            if dy <= eps:
                lo, hi = (px - before, qx + after) if px <= qx else (qx - after, px + before)
                x0, x1 = _snap_span(lo, hi)
                y0, y1 = _snap_span(py - hw, py + hw)
            else:
                lo, hi = (py - before, qy + after) if py <= qy else (qy - after, py + before)
                y0, y1 = _snap_span(lo, hi)
                x0, x1 = _snap_span(px - hw, px + hw)
            out.append([(x0, y0), (x1, y0), (x1, y1), (x0, y1)])
    return out


def _snap_thin_rect(p: list) -> list:
    """塗りの細い四角（罫線を四角で塗る書類が多い）：細いほうの辺を画素に合わせる（線の調整 10.6.5。2 画素未満のときだけ）。"""
    r = _rect_of(p)
    if r is None:
        return p
    x0, y0, x1, y1 = r
    thin_x, thin_y = 0 < x1 - x0 < 2.0, 0 < y1 - y0 < 2.0
    if not (thin_x or thin_y):
        return p
    if thin_x:
        x0, x1 = _snap_span(x0, x1)
    if thin_y:
        y0, y1 = _snap_span(y0, y1)
    return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]


def _unit_square_box(m) -> tuple[float, float, float, float]:
    """画像の単位の四角の、画素での外枠。"""
    xs = [m[4], m[0] + m[4], m[2] + m[4], m[0] + m[2] + m[4]]
    ys = [m[5], m[1] + m[5], m[3] + m[5], m[1] + m[3] + m[5]]
    return min(xs), min(ys), max(xs), max(ys)


def _dash_user(path, m, gs):
    """細い線の破線：破線の長さは利用者の座標で（8.4.3.6）、線そのものは画面の座標で作る。"""
    from .canvas import _dash

    subs = flatten(path, m, keep=True)
    dashed = _dash(subs, [max(0.0, float(v)) for v in gs.dash], gs.phase) if gs.dash and any(v > 0 for v in gs.dash) else subs
    a, b, c, d, e, f = m
    return [([(a * x + c * y + e, b * x + d * y + f) for x, y in pts], cl) for pts, cl in dashed]


# ====================================================================== 注釈の見た目
def annotation_forms(doc, page: dict):
    """描く注釈の見た目：(注釈の辞書, 見た目のストリーム, 見た目の行列 → ページ) の並び（12.5.5）。
    隠した注釈（F の Hidden）・画面に出さない注釈（NoView）・Popup は描かない。"""
    out = []
    for a in doc.resolve(page.get("Annots")) or []:
        ad = doc.resolve(a)
        if not isinstance(ad, dict):
            continue
        sub = getattr(ad.get("Subtype"), "name", "")
        if sub == "Popup":
            continue
        flags = int(doc.resolve(ad.get("F")) or 0)
        if flags & 2 or flags & 32:
            continue
        apd = doc.resolve(ad.get("AP"))
        n = doc.resolve(apd.get("N")) if isinstance(apd, dict) else None
        if isinstance(n, dict) and not isinstance(n, Stream):
            st = doc.resolve(ad.get("AS"))
            n = doc.resolve(n.get(getattr(st, "name", ""))) if st is not None else None
        if not isinstance(n, Stream):
            continue
        try:
            rect = [float(doc.resolve(v)) for v in doc.resolve(ad.get("Rect"))]
            bbox = [float(doc.resolve(v)) for v in doc.resolve(n.dict.get("BBox"))]
            mm = doc.resolve(n.dict.get("Matrix"))
            mm = tuple(float(doc.resolve(v)) for v in mm) if isinstance(mm, list) and len(mm) == 6 else IDENT
        except (TypeError, ValueError):
            continue
        if len(rect) != 4 or len(bbox) != 4:
            continue
        pts = [(mm[0] * x + mm[2] * y + mm[4], mm[1] * x + mm[3] * y + mm[5]) for x in (bbox[0], bbox[2]) for y in (bbox[1], bbox[3])]
        bx0, by0 = min(p[0] for p in pts), min(p[1] for p in pts)
        bx1, by1 = max(p[0] for p in pts), max(p[1] for p in pts)
        if bx1 - bx0 < 1e-6 or by1 - by0 < 1e-6:
            continue
        rx0, ry0, rx1, ry1 = min(rect[0], rect[2]), min(rect[1], rect[3]), max(rect[0], rect[2]), max(rect[1], rect[3])
        sx, sy = (rx1 - rx0) / (bx1 - bx0), (ry1 - ry0) / (by1 - by0)
        out.append((ad, n, (sx, 0.0, 0.0, sy, rx0 - bx0 * sx, ry0 - by0 * sy)))
    return out


def _field_value(doc, ad: dict):
    """欄の値（V。無ければ親の欄へさかのぼる：12.7.3.1 の継承）。"""
    d, seen = ad, 0
    while isinstance(d, dict) and seen < 16:
        v = doc.resolve(d.get("V"))
        if v is not None:
            return v
        d, seen = doc.resolve(d.get("Parent")), seen + 1
    return None


def appearance_notes(doc, page: dict) -> list[str]:
    """/AP の見た目で描くので、そのとおりには分からないもの（文にして返す）：
    - 見た目（/AP）のない欄と注釈：見る道具によっては値や辞書から見た目を作るが、ここでは描かない（12.5.5）
    - NeedAppearances（12.7.2）：見る道具が欄の見た目を値から作り直すので、位置・大きさが /AP と違って見えることがある"""
    missing: dict[str, int] = {}
    n_widget = 0
    for a in doc.resolve(page.get("Annots")) or []:
        ad = doc.resolve(a)
        if not isinstance(ad, dict):
            continue
        sub = getattr(ad.get("Subtype"), "name", "")
        if sub in ("Popup", "Link") or int(doc.resolve(ad.get("F")) or 0) & 34:
            continue
        if sub == "Widget":
            n_widget += 1
        apd = doc.resolve(ad.get("AP"))
        n = doc.resolve(apd.get("N")) if isinstance(apd, dict) else None
        if n is not None:
            continue
        v = _field_value(doc, ad) if sub == "Widget" else True
        if v in (None, b"", "", False) or getattr(v, "name", "") == "Off":
            continue
        k = "値の入った欄" if sub == "Widget" else f"注釈（{sub or '種類なし'}）"
        missing[k] = missing.get(k, 0) + 1
    out = [f"見た目（/AP）のない{k} {n} 個は描いていません。" for k, n in missing.items()]
    af = doc.resolve(doc.root.get("AcroForm"))
    if n_widget and isinstance(af, dict) and doc.resolve(af.get("NeedAppearances")) is True:
        out.append("この書類は NeedAppearances が true なので、見る道具が欄の見た目を値から作り直します（この画像は /AP のままで、"
                   "位置・大きさが違って見えることがあります）。")
    return out


def render_page(doc, page: dict, W: int, H: int, base, time_limit: float = 60.0) -> tuple[Canvas, Renderer]:
    """ページを W × H の画素に描く。base：ページの座標 → 画素の行列（回転・切り取り・拡大を含む）。"""
    cv = Canvas(W, H)
    r = Renderer(doc, cv, time.monotonic() + time_limit if time_limit else None)
    res = doc.resolve(page.get("Resources")) or {}
    gs = GS(base, Clip(0, 0, W, H))
    try:
        r.run(doc.page_content(page), res, gs, 0)
    except _Stop:
        pass
    r.annot_failed = []
    if not r.stopped:
        for ad, n, fit in annotation_forms(doc, page):
            if r.hidden(ad.get("OC")):
                continue
            g = GS(base, Clip(0, 0, W, H))
            try:
                r.form(g, n, res, 0, matrix=mul(fit, base))
            except _Stop:
                break
            except Exception as e:  # noqa: BLE001
                from ...limits import LimitExceeded

                if isinstance(e, LimitExceeded):
                    raise
                r.annot_failed.append(f"{getattr(ad.get('Subtype'), 'name', '注釈')}：{type(e).__name__}")
    return cv, r
