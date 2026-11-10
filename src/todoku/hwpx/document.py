"""HWPX（한글の文書。OWPML・KS X 6101）を読み、Word（.docx）と同じ形（段落・表・マス・入力欄）にする。

本文の節（Contents/sectionN.xml の hs:sec）の段落（hp:p）は、字の並び（hp:run の hp:t）と、その間の部品（hp:ctrl・表・図形・
フォームの部品）でできている。読むもの：
- 字（hp:t の中の字・hp:tab・hp:lineBreak・hp:fwSpace・hp:nbSpace・hp:hyphen）。変更の記録の消した字（hp:deleteBegin〜deleteEnd）は
  非表示の字にする
- 表（hp:tbl・hp:tr・hp:tc）：マスの番地（hp:cellAddr）と結合（hp:cellSpan）。結合で隠れたマスは、HWPX には書かれないので、
  縦の結合の下のマスを「続き」として作る。マスの保護（tc の protect="1"）
- 入力の欄：누름틀（hp:fieldBegin type="CLICK_HERE" 〜 hp:fieldEnd。値は2つの印のあいだの字、案内の字は Direction の引数）と
  フォームの部品（hp:checkBtn・hp:radioBtn・hp:comboBox・hp:listBox・hp:edit。hp:btn は押すボタン）
- 図形の中の字（hp:rect などの hp:drawText）・表と図形の見出し（hp:caption）・ヘッダー/フッター（hp:header・hp:footer）・
  脚注/文末脚注（hp:footNote・hp:endNote）・メモ（MEMO の欄の中身）・隠れた説明（hp:hiddenComment）・덧말（hp:dutmal）・
  글자 겹침（hp:compose）・数式（hp:equation の hp:script）・絵（hp:pic。説明は hp:shapeComment）
- 段落の番号（header.xml の paraPr の heading と、numbering・bullet の書式）
1字ごとに、どの要素のどのバイトから来たかを覚える（書き込みを小さな差分にするため）。
"""

from __future__ import annotations

from dataclasses import dataclass

from ..docx.document import Cell, Control, Para, Seg, Table, iter_paras
from ..limits import budget
from ..odf.text import Src
from ..odf.xml import El, Txt
from . import numbers
from .package import HwpxError, HwpxPackage, L, kid, kids, script_has_code, top

MARK_PREFIX = "_todoku_"  # 書き込んだ所のしおり（hp:bookmark の名前。「_todoku_」＋欄の id で始まり、「_todoku_end_」＋id で終わる）
END_PREFIX = "_todoku_end_"
SHAPES = {"rect", "ellipse", "arc", "polygon", "curve", "connectLine", "line", "textart", "container", "pic", "ole", "equation",
          "video", "chart"}
FORMS = {"btn", "checkBtn", "radioBtn", "comboBox", "listBox", "edit", "scrollBar"}
# 値を持たない、字の欄（ページ番号・日付など）：見えている字を、書き換えない字として読む
DISPLAY_FIELDS = {"DATE", "DOC_DATE", "PATH", "SUMMERY", "SUMMARY", "USER_INFO", "FORMULA", "CROSSREF", "MAILMERGE",
                  "PRIVATE_INFO_SECURITY", "METADATA", "CITATION", "BIBLIOGRAPHY", "METATAG", "PROOFREADING_MARKS", "TABLE_OF_CONTENTS"}


@dataclass
class _Open:
    """開いている欄（fieldBegin を読んで、fieldEnd をまだ読んでいない）。"""

    id: str
    ctl: Control | None  # 누름틀・字の欄。リンク・しおり・メモは None（中の字は本文の字）


class _ParaState:
    def __init__(self, p: Para):
        self.p = p
        self.pos = 0  # 見える字の数（しおりの位置）
        self.deleted = False  # 変更の記録で消した字の中
        self.fields: list[_Open] = []
        self.cur = None  # (Txt, field) 字の片をつなげるため


class HwpxDocument:
    kind = "hwpx"

    def __init__(self, blob: bytes, filename: str = "", partial: dict | None = None):
        """partial：本文の一部だけの小さなパッケージを読むとき（lazy）：{"section": 読む節, "master": マスターページを読むか, 前までの数え}。"""
        self.blob = blob
        self.filename = filename
        self.pkg = HwpxPackage(blob, filename)
        p_ = partial or {}
        self.controls: list[Control] = []
        self.pictures: list[dict] = [{"n": i + 1, "stub": True} for i in range(p_.get("pictures", 0))]
        self.notes: list[tuple[str, list]] = []
        self.comments: list = []
        self.headers: list[tuple[str, str, list]] = []
        self.marks_raw: list[tuple[str, str, Para, int]] = []
        self.radio_groups: dict[str, Control] = {}
        self.unread: list[str] = []
        self.protection = ""
        self.protection_enforced = False
        self.vba = None
        self.scripts = [(n, s) for n, s in self.pkg.scripts() if script_has_code(s)]
        self.has_macros = False  # Excel・Word のマクロ（VBA）はない。한글のマクロ（JScript）は scripts で示す
        self._header()
        self.num_counters: dict[str, list[int]] = {k: list(v) for k, v in p_.get("num_counters", {}).items()}
        self.outline_id = p_.get("outline_id", "")
        self.body: list = []
        self.parts = list(self.pkg.sections)
        self.part = self.parts[0]
        for part in self.parts:
            if partial is not None and part != p_.get("section"):
                continue
            root = self.pkg.tree(part)
            sec = top(root)
            if L(sec) != "sec":
                raise HwpxError(f"{part} が本文の節（hs:sec）ではありません")
            self.body += self._blocks(sec, part, None, False)
        self.stories: list[tuple[str, str, list]] = []
        for kind, part, blocks in self.headers:
            self.stories.append((kind, part, blocks))
        for kind, blocks in self.notes:
            self.stories.append((kind, self.part, blocks))
        if self.comments:
            self.stories.append(("comments", self.part, self.comments))
        for name in [n for n in self.pkg.names if n.startswith("Contents/masterpage") and n.endswith(".xml") and (partial is None or p_.get("master"))]:
            mp = top(self.pkg.tree(name))
            blocks = []
            for sl in mp.iter():
                if L(sl) == "subList" and (sl.parent is mp):
                    blocks += self._blocks(sl, name, None, True)
            if not blocks:
                blocks = self._blocks(mp, name, None, True)
            if blocks:
                self.stories.append(("master", name, blocks))

    # ------------------------------------------------------------------ 入口（Word と同じ問い合わせ）
    def all_blocks(self):
        yield "body", self.part, self.body
        for kind, part, blocks in self.stories:
            yield kind, part, blocks

    def xml(self, part: str) -> bytes:
        return self.pkg.xml(part)

    def marks(self) -> dict[str, tuple[Para, int, int]]:
        starts: dict[str, tuple[Para, int]] = {}
        out: dict[str, tuple[Para, int, int]] = {}
        for name, which, p, pos in self.marks_raw:
            if which == "end":
                key = name[len(END_PREFIX):]
                if key in starts and starts[key][0] is p:
                    out[key] = (p, starts[key][1], pos)
            elif which == "start":
                starts[name[len(MARK_PREFIX):]] = (p, pos)
        return out

    # ------------------------------------------------------------------ header.xml：文字の形・段落の形・番号
    def _header(self) -> None:
        self.under: dict[str, bool] = {}
        self.looks: dict[str, tuple] = {}  # 文字の形 → (取り消し線か, 字の色 "RRGGBB")：docx/fontfacts.py
        self.para_heading: dict[str, tuple[str, str, int]] = {}
        self.numberings: dict[str, dict[int, tuple[int, str, str]]] = {}
        self.bullets: dict[str, str] = {}
        if not self.pkg.has("Contents/header.xml"):
            return
        root = self.pkg.tree("Contents/header.xml")
        data = self.pkg.xml("Contents/header.xml")
        for el in top(root).iter():
            loc = L(el)
            if loc == "charPr":
                u = kid(el, "underline")
                self.under[el.attrs.get("id", "")] = u is not None and u.attrs.get("type", "NONE") == "BOTTOM"
                so = kid(el, "strikeout")
                c = (el.attrs.get("textColor") or "#000000").lstrip("#").upper()
                self.looks[el.attrs.get("id", "")] = (
                    so is not None and (so.attrs.get("shape") or "NONE").upper() not in ("NONE", ""),
                    c if len(c) == 6 and all(x in "0123456789ABCDEF" for x in c) else "000000")
            elif loc == "paraPr":
                h = kid(el, "heading")
                if h is not None:
                    try:
                        lv = int(h.attrs.get("level", "0") or 0)
                    except ValueError:
                        lv = 0
                    self.para_heading[el.attrs.get("id", "")] = (h.attrs.get("type", "NONE"), h.attrs.get("idRef", ""), lv)
            elif loc == "numbering":
                levels = {}
                for ph in kids(el, "paraHead"):
                    try:
                        lv = int(ph.attrs.get("level", "1") or 1)
                        st = int(ph.attrs.get("start", "1") or 1)
                    except ValueError:
                        continue
                    text = "".join(k.value(data) for k in ph.kids if isinstance(k, Txt))
                    levels[lv] = (st, ph.attrs.get("numFormat", "DIGIT"), text)
                self.numberings[el.attrs.get("id", "")] = levels
            elif loc == "bullet":
                self.bullets[el.attrs.get("id", "")] = el.attrs.get("char", "") or "・"

    def _label(self, el: El) -> str:
        h = self.para_heading.get(el.attrs.get("paraPrIDRef", ""))
        if h is None or h[0] == "NONE":
            return ""
        typ, ref, lv0 = h
        if typ == "BULLET":
            return self.bullets.get(ref, "・")
        num_id = self.outline_id if typ == "OUTLINE" else ref
        levels = self.numberings.get(num_id)
        lvl = lv0 + 1
        if not levels or lvl not in levels:
            return ""
        cnt = self.num_counters.setdefault(num_id, [0] * 11)
        if lvl >= len(cnt):
            return ""
        cnt[lvl] = cnt[lvl] + 1 if cnt[lvl] else levels[lvl][0]
        for k in range(lvl + 1, len(cnt)):
            cnt[k] = 0
        for k in range(1, lvl):
            if not cnt[k] and k in levels:
                cnt[k] = levels[k][0]
        fmts = {k: v[1] for k, v in levels.items()}
        return numbers.label(levels[lvl][2], cnt, fmts)

    # ------------------------------------------------------------------ ブロック
    def _blocks(self, node: El, part: str, cell: Cell | None, protected: bool) -> list:
        out: list = []
        for ch in node.els():
            loc = L(ch)
            if loc == "p":
                out += self._para(ch, part, cell, protected)
            elif loc == "subList":
                out += self._blocks(ch, part, cell, protected)
        return out

    def _sublist(self, el: El | None, part: str, cell: Cell | None, protected: bool) -> list:
        if el is None:
            return []
        sl = kid(el, "subList") if L(el) != "subList" else el
        return self._blocks(sl, part, cell, protected) if sl is not None else []

    def _table(self, el: El, part: str, protected: bool) -> list:
        t = Table(el, part)
        t.protected = protected
        below: dict[int, list[Cell]] = {}
        trs = kids(el, "tr")
        for r, tr in enumerate(trs):
            row: list[Cell] = []
            for tc in kids(tr, "tc"):
                addr = kid(tc, "cellAddr")
                span = kid(tc, "cellSpan")
                try:
                    col = int(addr.attrs.get("colAddr", "0")) if addr is not None else (row[-1].col + row[-1].span if row else 0)
                    cs = max(1, int(span.attrs.get("colSpan", "1"))) if span is not None else 1
                    rs = max(1, int(span.attrs.get("rowSpan", "1"))) if span is not None else 1
                except ValueError:
                    raise HwpxError("表のマスの番地・結合の数が読めません") from None
                c = Cell(tc, r, col, cs, "restart" if rs > 1 else "", table=t)
                c.protected = protected or tc.attrs.get("protect") == "1"
                c.blocks = self._sublist(tc, part, c, c.protected)
                row.append(c)
                last = min(r + rs, len(trs))  # 結合の宣言は表の行の数までに切る（宣言の数を信じない）
                if last > r + 1:
                    budget().count(last - r - 1, "表の縦の結合")
                for rr in range(r + 1, last):
                    cc = Cell(tc, rr, col, cs, "continue", table=t)
                    cc.protected = c.protected
                    below.setdefault(rr, []).append(cc)
            row += below.pop(r, [])
            row.sort(key=lambda x: x.col)
            t.rows.append(row)
        out: list = [t]
        cap = kid(el, "caption")
        if cap is not None:
            out += self._sublist(cap, part, None, protected)
        return out

    # ------------------------------------------------------------------ 段落
    def _para(self, el: El, part: str, cell: Cell | None, protected: bool) -> list:
        p = Para(el, part, [], cell=cell)
        p.el = el
        p.protected = protected
        p.label = self._label(el)
        st = _ParaState(p)
        after: list = []  # 段落の中の表（Word の形では、段落の後ろのブロック）
        data = self.pkg.xml(part)
        for run in kids(el, "run"):
            under = self.under.get(run.attrs.get("charPrIDRef", ""), False)
            n0, look = len(p.segs), self.looks.get(run.attrs.get("charPrIDRef", ""))
            for ch in run.els():
                loc = L(ch)
                if loc == "t":
                    self._t(ch, st, part, data, under, run)
                elif loc == "ctrl":
                    for c in ch.els():
                        self._ctrl(c, ch, st, part, data, run)
                elif loc == "tbl":
                    st.cur = None
                    after += self._table(ch, part, protected)
                elif loc in SHAPES:
                    st.cur = None
                    self._shape(ch, st, part, data, run)
                elif loc in FORMS:
                    st.cur = None
                    self._form(ch, st, part, data, run)
                elif loc == "dutmal":
                    st.cur = None
                    main = kid(ch, "mainText")
                    sub = kid(ch, "subText")
                    mt = _text(main, data) if main is not None else ""
                    stt = _text(sub, data) if sub is not None else ""
                    src = Src("el", part, ch.s, ch.e, el=ch)
                    self._fixed(st, mt, src, run)
                    if stt:
                        self._fixed(st, f"（{stt}）", src, run)
                elif loc == "compose":
                    st.cur = None
                    self._fixed(st, ch.attrs.get("composeText", ""), Src("el", part, ch.s, ch.e, el=ch), run)
                elif loc == "secPr":
                    self.outline_id = ch.attrs.get("outlineShapeIDRef", "") or self.outline_id
                elif loc in ("markpenBegin", "markpenEnd", "titleMark"):
                    continue
                elif loc == "" or loc in ("unknownObj",):
                    continue
            if look is not None:
                for sg in p.segs[n0:]:
                    if sg.kind == "t" and sg.look is None:
                        sg.look = look
        # 閉じていない欄（段落の終わりまで）
        for f in st.fields:
            if f.ctl is not None:
                self._close_ctl(f.ctl, None, st, part)
        st.fields = []
        return [p] + after

    def _fixed(self, st: _ParaState, text: str, src: Src, run: El | None) -> None:
        """窓のために作った字・図形などの字（書き換えない字）：字の欄と同じに扱い、空欄を探さない。"""
        fc = Control("field", "hwpx-fixed", "", src.el, src.part, st.p)
        self._add(st, Seg(text, "other", src, run, False, False, fc))

    def _add(self, st: _ParaState, sg: Seg) -> None:
        if st.deleted:
            sg.hidden = True
        f = next((x.ctl for x in reversed(st.fields) if x.ctl is not None), None)
        if f is not None and sg.field is None:
            sg.field = f
            f.result_segs.append(sg)
        st.p.segs.append(sg)
        if not sg.hidden:
            st.pos += len(sg.text)

    def _t(self, t: El, st: _ParaState, part: str, data: bytes, under: bool, run: El) -> None:
        for k in t.kids:
            if isinstance(k, Txt):
                text, offs = k.decode(data)
                if not text:
                    continue
                f = next((x.ctl for x in reversed(st.fields) if x.ctl is not None), None)
                src = Src("txt", part, offs[0], offs[-1], list(offs[:-1]), list(offs[1:]), el=t)
                sg = Seg(text, "t", src, run, False, under, None)
                if f is not None:
                    sg.field = f
                    f.result_segs.append(sg)
                if st.deleted:
                    sg.hidden = True
                st.p.segs.append(sg)
                if not sg.hidden:
                    st.pos += len(text)
                st.cur = (k, f)
                continue
            loc = L(k)
            st.cur = None
            src = Src("elx", part, k.s, k.e, el=k)
            if loc == "tab":
                self._add(st, Seg("\t", "tab", src, run, False, under))
            elif loc == "lineBreak":
                self._add(st, Seg("\n", "br", src, run))
            elif loc == "fwSpace":
                self._add(st, Seg("　", "t", src, run, False, under))
            elif loc == "nbSpace":
                self._add(st, Seg(" ", "t", src, run, False, under))
            elif loc == "hyphen":
                self._add(st, Seg("-", "other", src, run))
            elif loc == "deleteBegin":
                st.deleted = True
            elif loc == "deleteEnd":
                st.deleted = False
            # insertBegin/insertEnd・markpenBegin/End・titleMark：印だけ

    # ------------------------------------------------------------------ 部品（hp:ctrl）
    def _ctrl(self, c: El, ctrl: El, st: _ParaState, part: str, data: bytes, run: El) -> None:
        st.cur = None
        loc = L(c)
        if loc == "fieldBegin":
            typ = (c.attrs.get("type") or "").upper()
            fid = c.attrs.get("id", "")
            if typ == "CLICK_HERE":
                direction, help_ = _click_params(c, data)
                ctl = Control("text", "hwpx-clickhere", c.attrs.get("name", "") or "", c, part, st.p, help="")
                ctl.direction = direction
                ctl.help_text = help_
                ctl.begin_ctrl = ctrl
                ctl.begin_run = run
                ctl.dirty = c.attrs.get("dirty", "0") == "1"
                if c.attrs.get("editable", "1") == "0":
                    ctl.lock = "contentLocked"
                self.controls.append(ctl)
                st.fields.append(_Open(fid, ctl))
            elif typ == "MEMO":
                sl = kid(c, "subList")
                if sl is not None:
                    self.comments += self._blocks(sl, part, None, True)
                st.fields.append(_Open(fid, None))
            elif typ in DISPLAY_FIELDS:
                fc = Control("field", "hwpx-field", typ, c, part, st.p)
                st.fields.append(_Open(fid, fc))
            else:  # HYPERLINK・BOOKMARK など：中の字は本文の字
                st.fields.append(_Open(fid, None))
            return
        if loc == "fieldEnd":
            ref = c.attrs.get("beginIDRef", "")
            k = next((i for i in range(len(st.fields) - 1, -1, -1) if st.fields[i].id == ref), len(st.fields) - 1 if st.fields else None)
            if k is None:
                return
            f = st.fields.pop(k)
            if f.ctl is not None and f.ctl.kind != "field":
                self._close_ctl(f.ctl, ctrl, st, part)
            return
        if loc == "bookmark":
            name = c.attrs.get("name", "") or ""
            if name.startswith(END_PREFIX):
                self.marks_raw.append((name, "end", st.p, st.pos))
            elif name.startswith(MARK_PREFIX):
                self.marks_raw.append((name, "start", st.p, st.pos))
            return
        if loc in ("header", "footer"):
            blocks = self._sublist(c, part, None, False)
            if blocks:
                self.headers.append((loc, part, blocks))
            return
        if loc in ("footNote", "endNote"):
            kind = "footnotes" if loc == "footNote" else "endnotes"
            n = sum(1 for k_, _b in self.notes if k_ == kind) + 1
            self._add(st, Seg(f"※{n}", "note", Src("el", part, c.s, c.e, el=c), run))
            self.notes.append((kind, self._sublist(c, part, None, False)))
            return
        if loc == "hiddenComment":
            self.comments += self._sublist(c, part, None, True)
            return
        # colPr・pageNum・autoNum・newNum・pageNumCtrl・pageHiding・indexmark：字ではない（ページの番号などは表示のための字）

    def _close_ctl(self, ctl: Control, end_ctrl: El | None, st: _ParaState, part: str) -> None:
        ctl.end_ctrl = end_ctrl
        content = "".join(s.text for s in ctl.result_segs if not s.hidden)
        if not ctl.result_segs:
            at = end_ctrl.s if end_ctrl is not None else ctl.begin_ctrl.e
            sg = Seg("", "field", Src("ctl", part, at, at, el=end_ctrl if end_ctrl is not None else ctl.begin_ctrl), None, False, False, ctl)
            st.p.segs.append(sg)
            ctl.result_segs = [sg]
        # 値のない 누름틀：한글は案内の字（Direction）を表示する（印刷しない）。手を入れていない（dirty="0"）欄の字が案内の字と
        # 同じときも、한글は案内の字として扱う
        ctl.placeholder = content == "" or (not ctl.dirty and bool(ctl.direction) and content.rstrip() == ctl.direction)
        ctl.value = None if ctl.placeholder else content

    # ------------------------------------------------------------------ 図形・絵
    def _shape(self, el: El, st: _ParaState, part: str, data: bytes, run: El) -> None:
        loc = L(el)
        p = st.p
        src = Src("el", part, el.s, el.e, el=el)
        dt = kid(el, "drawText")
        if dt is not None:
            p.boxes += self._sublist(dt, part, None, getattr(p, "protected", False))
        cap = kid(el, "caption")
        if cap is not None:
            p.boxes += self._sublist(cap, part, None, getattr(p, "protected", False))
        if loc == "container":
            for ch in el.els():
                if L(ch) in SHAPES:
                    self._shape(ch, st, part, data, run)
            return
        sc = kid(el, "shapeComment")
        alt = _text(sc, data).strip() if sc is not None else ""
        if loc == "pic":
            self.pictures.append({"n": len(self.pictures) + 1, "alt": alt, "decorative": False, "para": p})
            self._fixed(st, f"〔図: {alt}〕" if alt else f"〔図{len(self.pictures)}（説明なし）〕", src, run)
        elif loc == "equation":
            sc_ = kid(el, "script")
            script = _text(sc_, data).strip() if sc_ is not None else ""
            self._fixed(st, f"〔数式: {script}〕", src, run)
        elif loc == "textart":
            t = el.attrs.get("text", "")
            if t:
                self._fixed(st, t, src, run)
        elif loc == "ole":
            self._fixed(st, f"〔埋め込み: {alt or 'オブジェクト'}〕", src, run)
            self.unread.append(f"埋め込み（OLE）の中身は読めません（{alt or '説明なし'}）")
        elif loc == "chart":
            self._fixed(st, "〔グラフ〕", src, run)
        elif loc == "video":
            self._fixed(st, f"〔動画: {alt}〕" if alt else "〔動画〕", src, run)

    # ------------------------------------------------------------------ フォームの部品（양식 개체）
    def _form(self, el: El, st: _ParaState, part: str, data: bytes, run: El) -> None:
        loc = L(el)
        p = st.p
        src = Src("ctl", part, el.s, el.e, el=el)
        name = el.attrs.get("name", "") or ""
        lock = "contentLocked" if el.attrs.get("enabled", "1") == "0" else ""
        if loc == "btn":
            cap = el.attrs.get("caption", "")
            self._fixed(st, f"〔ボタン: {cap}〕" if cap else "〔ボタン〕", src, run)
            return
        if loc == "scrollBar":
            return
        if loc == "checkBtn":
            on = el.attrs.get("value", "UNCHECKED") == "CHECKED"
            ctl = Control("checkbox", "hwpx-form", name, el, part, p, checked=on, help=el.attrs.get("caption", ""), lock=lock)
            ctl.form_kind = loc
            self.controls.append(ctl)
            sg = Seg("☑" if on else "☐", "field", src, run, False, False, ctl)
            ctl.result_segs = [sg]
            self._add(st, sg)
            if ctl.help:
                self._fixed(st, ctl.help, src, run)
            return
        if loc == "radioBtn":
            gname = el.attrs.get("radioGroupName", "") or name
            g = self.radio_groups.get(gname)
            cap = el.attrs.get("caption", "") or el.attrs.get("name", "")
            if g is None:
                g = Control("dropdown", "hwpx-radio", gname, el, part, p, lock=lock)
                g.form_kind = "radio"
                g.radios = []
                g.radio_paras = []
                self.radio_groups[gname] = g
                self.controls.append(g)
            g.radios.append(el)
            g.radio_paras.append(p)
            g.options.append(cap)
            if el.attrs.get("value") == "CHECKED":
                g.value = cap
            g.result_segs = [Seg(g.value or "", "field", None, None, False, False, g)]
            self._fixed(st, ("◉" if el.attrs.get("value") == "CHECKED" else "○") + cap, src, run)
            return
        if loc in ("comboBox", "listBox"):
            items = kids(el, "listItem")
            opts = [it.attrs.get("displayText", "") or it.attrs.get("value", "") for it in items]
            vals = [it.attrs.get("value", "") for it in items]
            sel = el.attrs.get("selectedValue", "")
            cur = None
            if sel:
                cur = opts[vals.index(sel)] if sel in vals else (sel if sel in opts else sel)
            kind = "combo" if loc == "comboBox" and el.attrs.get("editEnable", "0") == "1" else "dropdown"
            ctl = Control(kind, "hwpx-form", name, el, part, p, value=cur, options=[o for o in opts if o != ""] if any(opts) else [],
                          lock=lock)
            ctl.form_kind = loc
            ctl.item_values = vals
            ctl.item_labels = opts
            self.controls.append(ctl)
            sg = Seg(cur or "", "field", src, run, False, False, ctl)
            ctl.result_segs = [sg]
            self._add(st, sg)
            return
        if loc == "edit":
            te = kid(el, "text")
            cur = _text(te, data) if te is not None else ""
            ro = el.attrs.get("readOnly", "0") == "1"
            ctl = Control("text", "hwpx-form", name, el, part, p, value=cur, lock="contentLocked" if ro else lock)
            ctl.form_kind = loc
            try:
                ml = int(el.attrs.get("maxLength", "0") or 0)
            except ValueError:
                ml = 0
            if 0 < ml < 0x7FFFFFFF:
                ctl.maxlen = ml
            ctl.num_only = el.attrs.get("numOnly", "0") == "1"
            self.controls.append(ctl)
            sg = Seg(cur, "field", src, run, False, False, ctl)
            ctl.result_segs = [sg]
            self._add(st, sg)
            return


def _text(el: El, data: bytes) -> str:
    """要素の中の字をそのままつなぐ（hp:t の中の tab・lineBreak も字にする）。"""
    out = []
    stack = [el]
    while stack:
        n = stack.pop()
        if isinstance(n, El):
            loc = L(n)
            if loc == "tab":
                out.append("\t")
            elif loc == "lineBreak":
                out.append("\n")
            stack.extend(reversed(n.kids))
        else:
            out.append(n.value(data))
    return "".join(out)


def _click_params(begin: El, data: bytes) -> tuple[str, str]:
    """누름틀の引数（hp:parameters の stringParam）：案内の字（Direction）と説明（HelpState）。
    stringParam がなければ、Command の「Direction:wstring:N:」の後ろの N 字。"""
    direction = help_ = ""
    command = ""
    params = kid(begin, "parameters")
    if params is not None:
        for prm in params.els():
            nm = prm.attrs.get("name", "")
            if nm == "Direction":
                direction = _text(prm, data)
            elif nm == "HelpState":
                help_ = _text(prm, data)
            elif nm == "Command":
                command = _text(prm, data)
    if not direction and command:
        import re

        m = re.search(r"Direction:wstring:(\d+):", command)
        if m:
            n = int(m[1])
            direction = command[m.end():m.end() + n]
    return direction.strip(), help_.strip()


def protected_paras(doc: HwpxDocument) -> set[int]:
    out = set()
    for _k, _part, blocks in doc.all_blocks():
        for p in iter_paras(blocks):
            if getattr(p, "protected", False):
                out.add(id(p))
    return out
