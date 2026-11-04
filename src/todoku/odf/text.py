"""OpenDocument の文書（.odt・.ott・.fodt）を読む（ODF 1.3 第3部 5〜9・11・13 章）。

本文・表・リスト（段落番号）・節・目次の中身・枠（テキストボックス・図形の字）・ヘッダー/フッター・脚注/文末脚注・コメント・
フィールド（日付・ページ番号など、入力の欄（text:placeholder・text:text-input・text:drop-down））・
フォームの部品（office:forms：文字・チェック・リスト・オプションボタンなど）を読み、Word（.docx）と同じ形
（段落・表・マス・入力欄）にする。窓・入力欄の推定・字を目印にした書き込みは、Word と同じ処理を使う。

段落の字は、ODF の空白の決まり（6.1.2：段落の中の空白の連なりは1つ、段落の頭と終わりの空白は除く。text:s・text:tab・
text:line-break は字のまま）で作り、1字ごとに、どの要素のどのバイトから来たかを覚える（書き込みを小さな差分にするため）。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..docx.document import Cell, Control, Para, Seg, Table, iter_paras, number_text
from ..limits import MAX_DECLARED_CELLS
from .package import OdfError, OdfPackage, text_of
from .styles import Styles
from .xml import NS, El, Txt

MAX_REPEATED = MAX_DECLARED_CELLS // 4  # 表の繰り返しの宣言（table:number-…-repeated）から広げるマスの数（文書ごと）

MARK_PREFIX = "_todoku_"  # 書き込んだ所のしおり（名前は「_todoku_」＋欄の id そのもの）

_T = "{" + NS["text"] + "}"
_DRAW = "{" + NS["draw"] + "}"
_TABLE = "{" + NS["table"] + "}"
_OFFICE = "{" + NS["office"] + "}"
_FORM = "{" + NS["form"] + "}"
_FIELDNS = "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"
_LOEXT = "{" + NS["loext"] + "}"
# LibreOffice のコンテンツコントロール（loext:content-control、Word の w:sdt を LibreOffice が読み書きする形）の種類
CC_KINDS = (("checkbox", "checkbox"), ("dropdown", "dropdown"), ("combobox", "combo"), ("date", "date"), ("picture", "picture"),
            ("plain-text", "text"))

# 印だけの要素（空白の決まりでは無いものとみなす。6.1.2 の 2a）
MARKS = {_T + x for x in ("bookmark", "bookmark-start", "bookmark-end", "reference-mark", "reference-mark-start",
                          "reference-mark-end", "toc-mark", "toc-mark-start", "toc-mark-end", "user-index-mark",
                          "user-index-mark-start", "user-index-mark-end", "alphabetical-index-mark",
                          "alphabetical-index-mark-start", "alphabetical-index-mark-end", "change", "change-start",
                          "change-end", "soft-page-break", "hidden-paragraph")} | {_OFFICE + "annotation-end"}
# 中の字をそのまま段落の字として読む、入れ物の要素（書式・リンク・メタデータ）
INLINE = {_T + "span", _T + "a", _T + "meta", _T + "meta-field", _DRAW + "a"}
SHAPES = {_DRAW + x for x in ("frame", "custom-shape", "rect", "ellipse", "circle", "line", "polyline", "polygon", "path",
                             "regular-polygon", "connector", "caption", "measure", "g")}
NUM_FORMAT = {"1": "decimal", "a": "lowerLetter", "A": "upperLetter", "i": "lowerRoman", "I": "upperRoman",
              "１, ２, ３, ...": "decimalFullWidth", "１, ２, ３": "decimalFullWidth", "①, ②, ③, ...": "decimalEnclosedCircle",
              "①, ②, ③": "decimalEnclosedCircle", "一, 二, 三, ...": "japaneseCounting", "一, 二, 三": "japaneseCounting",
              "ア, イ, ウ, ...": "aiueoFullWidth", "ア, イ, ウ": "aiueoFullWidth", "イ, ロ, ハ, ...": "irohaFullWidth",
              "イ, ロ, ハ": "irohaFullWidth"}


@dataclass
class Src:
    """段落の字の1片が、XML のどこから来たか（Seg.node に入れる。並べるために .s を持つ）。"""

    kind: str  # txt（文字）/ s（text:s）/ tab / br / field（フィールドの表示）/ el（図の説明・脚注の印など）/ ctl（部品）/ text（入力の欄の中の字）
    part: str
    s: int
    e: int
    starts: list[int] | None = None  # txt：字ごとの始まりのバイト
    ends: list[int] | None = None  # txt：字ごとの終わりのバイト
    el: El | None = None
    count: int = 0  # s：空白の数


@dataclass
class _Ctx:
    underline: bool = False
    hidden: bool = False
    field: Control | None = None
    run: El | None = None


class OdtDocument:
    kind = "odt"

    def __init__(self, blob: bytes, filename: str = "", partial: dict | None = None):
        """partial：本文の一部だけの小さなパッケージを読むとき（lazy）：前までの数え（list_counters・outline_counter・pictures）と、マスターページを読むか（master）。"""
        self.blob = blob
        self.filename = filename
        self.pkg = OdfPackage(blob, filename)
        p_ = partial or {}
        if self.pkg.kind not in ("odt", "", "odp", "odg"):
            raise OdfError(f"OpenDocument の文書ではありません（{self.pkg.mimetype}）")
        body = self.pkg.body()
        # 発表資料（.odp）・図形描画（.odg）：本文はページ（draw:page）の並びで、字は図形・枠の中（ODF 1.3 第 3 部 10.2）
        self.office_text = next((x for x in map(body.find, ("office:text", "office:presentation", "office:drawing")) if x is not None), None)
        if self.office_text is None:
            raise OdfError("文書の本文（office:text）がありません")
        self.page_count = 0
        self.styles = Styles(self.pkg)
        self.unread: list[str] = []  # 読まなかった所
        self.repeats: list[str] = []  # 同じ中身の繰り返し（1 つぶんだけ窓に出している所。中身はぜんぶ読んである）
        self._repeated = 0  # 繰り返しを広げたマスの数
        self.controls: list[Control] = []
        self.pictures: list[dict] = [{"n": i + 1, "stub": True} for i in range(p_.get("pictures", 0))]  # 絵（draw:image）：説明（svg:title・svg:desc）と飾りの印。段階とは別の注意に使う
        self.form_by_id: dict[str, Control] = {}
        self.radio_groups: dict[str, Control] = {}
        self.notes: list[tuple[str, list]] = []  # (footnotes / endnotes, ブロック)
        self.comments: list = []
        self.marks_raw: list[tuple[str, str, Para, int]] = []  # (名前, start/end/point, 段落, 字の位置)
        self.list_counters: dict[str, list[int]] = {k: list(v) for k, v in p_.get("list_counters", {}).items()}
        self.outline_counter = list(p_.get("outline_counter", [0] * 11))
        self.vba = None
        self.has_macros = self.pkg.has_macros
        settings = self.pkg.settings()
        self.protection = "forms" if settings.get("ProtectForm") == "true" else ""
        self.protection_enforced = bool(self.protection)
        self.part = self.pkg.part_of("content.xml")
        self._forms(self.office_text)
        self.body = self._blocks(self.office_text, self.part, None, False)
        self.stories: list[tuple[str, str, list]] = []
        if partial is None or p_.get("master"):
            self._master_pages()
        for kind, blocks in self.notes:
            self.stories.append((kind, self.part, blocks))
        if self.comments:
            self.stories.append(("comments", self.part, self.comments))
        self.macros_info = None

    # ------------------------------------------------------------------ 入口（Word と同じ問い合わせ）
    def all_blocks(self):
        yield "body", self.part, self.body
        for kind, part, blocks in self.stories:
            yield kind, part, blocks

    def xml(self, part: str) -> bytes:
        return self.pkg.xml(part)

    def marks(self) -> dict[str, tuple[Para, int, int]]:
        """書き込んだ所のしおり（_todoku_＋id）：id → (段落, 始まり, 終わり)。"""
        starts: dict[str, tuple[Para, int]] = {}
        out: dict[str, tuple[Para, int, int]] = {}
        for name, which, p, pos in self.marks_raw:
            if not name.startswith(MARK_PREFIX):
                continue
            key = name[len(MARK_PREFIX):]
            if which == "start":
                starts[key] = (p, pos)
            elif which == "end" and key in starts and starts[key][0] is p:
                out[key] = (p, starts[key][1], pos)
            elif which == "point":
                out[key] = (p, pos, pos)
        return out

    # ------------------------------------------------------------------ ヘッダー・フッター
    def _master_pages(self) -> None:
        part = self.pkg.part_of("styles.xml")
        ms = self.pkg.section(part, "master-styles") if (self.pkg.flat or self.pkg.has("styles.xml")) else None
        if ms is None:
            return
        for mp in ms.findall("style:handout-master") + ms.findall("style:master-page"):
            if self.pkg.kind in ("odp", "odg"):
                # スライドマスターの枠の字（プレースホルダーの見本も、印刷される字も。どちらかは決めずに、マスターの字として示す）
                blocks = self._blocks(mp, part, None, False)
                if blocks:
                    self.stories.append(("master", part, blocks))
                continue
            for el in mp.els():
                loc = el.local
                if loc.startswith("header") or loc.startswith("footer"):
                    blocks = self._blocks(el, part, None, False)
                    if blocks:
                        self.stories.append(("header" if loc.startswith("header") else "footer", part, blocks))

    # ------------------------------------------------------------------ ブロック
    def _blocks(self, node: El, part: str, cell: Cell | None, protected: bool, list_ctx: tuple | None = None) -> list:
        out: list = []
        for ch in node.els():
            t = ch.tag
            if t in (_T + "p", _T + "h"):
                out.append(self._para(ch, part, cell, protected))
            elif t == _TABLE + "table":
                out.append(self._table(ch, part, protected or ch.get("table:protected") == "true"))
            elif t == _T + "list":
                out += self._list(ch, part, cell, protected, list_ctx)
            elif t == _T + "numbered-paragraph":
                out += self._blocks(ch, part, cell, protected, list_ctx)
            elif t == _T + "section":
                out += self._blocks(ch, part, cell, protected or ch.get("text:protected") == "true", list_ctx)
            elif t.startswith(_T) and ch.find("text:index-body") is not None:
                out += self._blocks(ch.find("text:index-body"), part, cell, protected, list_ctx)
            elif t in (_T + "index-body", _T + "index-title"):
                out += self._blocks(ch, part, cell, protected, list_ctx)
            elif t == _DRAW + "page":
                out += self._draw_page(ch, part, cell, protected)
            elif t.endswith("}notes") and t.startswith("{urn:oasis:names:tc:opendocument:xmlns:presentation"):
                got = self._blocks(ch, part, cell, protected, list_ctx)
                first = next((b for b in got if isinstance(b, Para)), None)
                if first is not None:
                    first.label = "（ノート）" + (first.label or "")
                out += got
            elif t in SHAPES:
                # ページや段落に固定した枠・図形（段落の外）：中身を、字のない段落の枠として読む
                p = Para(ch, part, [], cell=cell)
                p.el = ch
                p.protected = protected
                self._shape(ch, p, part, _Ctx())
                if p.boxes or p.segs:
                    out.append(p)
            elif t in (_T + "tracked-changes", _OFFICE + "forms", _T + "variable-decls", _T + "sequence-decls",
                       _T + "user-field-decls", _T + "dde-connection-decls", _TABLE + "calculation-settings",
                       _TABLE + "content-validations", _TABLE + "label-ranges", _TABLE + "named-expressions",
                       _TABLE + "database-ranges", _TABLE + "data-pilot-tables", _T + "alphabetical-index-auto-mark-file",
                       _T + "soft-page-break", _OFFICE + "event-listeners"):
                continue
        return out

    def _draw_page(self, el: El, part: str, cell: Cell | None, protected: bool) -> list:
        """発表資料・図形描画の 1 ページ：字のない段落に、何ページ目かと名前を添える（窓の見出し。ファイルの字ではない）。"""
        self.page_count += 1
        name = el.get("draw:name") or ""
        got = self._blocks(el, part, cell, protected)
        head = Para(el, part, [], cell=cell)
        head.el = el
        head.protected = protected
        head.label = f"〔スライド {self.page_count}" + (f"：{name}" if name else "") + "〕"
        return [head] + got

    def _list(self, el: El, part: str, cell: Cell | None, protected: bool, list_ctx: tuple | None) -> list:
        """リスト：項目ごとに番号（リストのスタイルの書式）を数え、最初の段落の頭に付ける。"""
        style = el.get("text:style-name") or (list_ctx[0] if list_ctx else None)
        level = (list_ctx[1] + 1) if list_ctx else 1
        key = style or "#"
        cnt = self.list_counters.setdefault(key, [0] * 11)
        if level == 1 and el.get("text:continue-numbering") != "true" and not el.get("text:continue-list"):
            cnt[:] = [0] * 11
        out: list = []
        for item in el.els():
            if item.tag not in (_T + "list-item", _T + "list-header"):
                continue
            label = ""
            if item.tag == _T + "list-item":
                sv = item.get("text:start-value")
                if sv and sv.isdigit():
                    cnt[level] = int(sv)
                else:
                    cnt[level] += 1
                for k in range(level + 1, 11):
                    cnt[k] = 0
                label = self._list_label(part, style, level, cnt)
            blocks = self._blocks(item, part, cell, protected, (style, level))
            first = next((b for b in blocks if isinstance(b, Para)), None)
            if first is not None and label and item.els() and item.els()[0].tag in (_T + "p", _T + "h"):
                first.label = label
            out += blocks
        return out

    def _list_label(self, part: str, style: str | None, level: int, cnt: list[int]) -> str:
        ls = self.styles.list_style(part, style)
        if ls is None:
            return ""
        lv = next((x for x in ls.els() if x.get("text:level") == str(level)), None)
        if lv is None:
            return ""
        if lv.local == "list-level-style-bullet":
            return lv.get("text:bullet-char", "・") or "・"
        if lv.local == "list-level-style-image":
            return "・"
        fmt = lv.get("style:num-format", "")
        if not fmt:
            return (lv.get("style:num-prefix", "") or "") + (lv.get("style:num-suffix", "") or "")
        shown = int(lv.get("text:display-levels", "1") or 1)
        start = int(lv.get("text:start-value", "1") or 1)
        nums = []
        for lvl in range(max(1, level - shown + 1), level + 1):
            n = cnt[lvl] + start - 1 if lvl == level else cnt[lvl] + start - 1
            lvx = next((x for x in ls.els() if x.get("text:level") == str(lvl)), lv)
            f = lvx.get("style:num-format", fmt) or fmt
            nums.append(number_text(max(n, 1), NUM_FORMAT.get(f, "decimal")))
        return (lv.get("style:num-prefix", "") or "") + ".".join(nums) + (lv.get("style:num-suffix", "") or "")

    def _table(self, el: El, part: str, protected: bool) -> Table:
        t = Table(el, part)
        t.protected = protected
        rows: list[El] = []

        def collect(n: El) -> None:
            for ch in n.els():
                if ch.tag == _TABLE + "table-row":
                    rep = min(int(ch.get("table:number-rows-repeated", "1") or 1), 50)
                    if rep > 1:
                        # 繰り返しを広げてよいのは、文書で MAX_REPEATED マスまで（小さなファイルが何百万ものマスにならないように）
                        cells = sum(min(int(tc.get("table:number-columns-repeated", "1") or 1), 200) for tc in ch.els()
                                    if tc.tag in (_TABLE + "table-cell", _TABLE + "covered-table-cell"))
                        if self._repeated + rep * cells > MAX_REPEATED:
                            self.repeats.append(f"同じ中身を {rep} 回繰り返す表の行は、窓に 1 行だけ出しています")
                            rep = 1
                        else:
                            self._repeated += rep * cells
                    rows.extend([ch] * rep)
                elif ch.tag in (_TABLE + "table-header-rows", _TABLE + "table-rows", _TABLE + "table-row-group"):
                    collect(ch)

        collect(el)
        below: dict[int, tuple[int, int]] = {}  # 列 → (縦に結合したマスの終わりの行, 横の幅)
        for r, tr in enumerate(rows):
            row: list[Cell] = []
            col = 0
            skip = 0
            for tc in tr.els():
                if tc.tag not in (_TABLE + "table-cell", _TABLE + "covered-table-cell"):
                    continue
                rep = min(int(tc.get("table:number-columns-repeated", "1") or 1), 200)
                if rep > 1:
                    if self._repeated + rep > MAX_REPEATED:
                        self.repeats.append(f"同じ中身を {rep} 回繰り返す表のマスは、窓に 1 つだけ出しています")
                        rep = 1
                    else:
                        self._repeated += rep
                for k in range(rep):
                    if tc.tag == _TABLE + "covered-table-cell":
                        if skip > 0:
                            skip -= 1
                            col += 1
                            continue
                        span = below.get(col, (r, 1))[1] if below.get(col, (-1, 1))[0] >= r else 1
                        c = Cell(tc, r, col, span, "continue", table=t)
                        c.blocks = self._blocks(tc, part, c, protected)
                        c.repeated = rep > 1
                        row.append(c)
                        skip = span - 1
                        col += 1
                        continue
                    span = int(tc.get("table:number-columns-spanned", "1") or 1)
                    rspan = int(tc.get("table:number-rows-spanned", "1") or 1)
                    shade = self._shade(part, tc.get("table:style-name"))
                    c = Cell(tc, r, col, span, "restart" if rspan > 1 else "", table=t, shade=shade)
                    c.repeated = rep > 1
                    c.blocks = self._blocks(tc, part, c, protected)
                    if rspan > 1:
                        below[col] = (r + rspan - 1, span)
                    row.append(c)
                    skip = span - 1
                    col += 1
            t.rows.append(row)
        # 繰り返しの行（同じ要素を複数の行として読んだもの）：書き込みでは分けないので、印を付ける
        counts: dict[int, int] = {}
        for tr in rows:
            counts[id(tr)] = counts.get(id(tr), 0) + 1
        for tr, row in zip(rows, t.rows):
            if counts[id(tr)] > 1:
                for c in row:
                    c.repeated = True
        return t

    def _shade(self, part: str, style: str | None) -> str:
        v = self.styles.prop(part, "table-cell", style, "style:table-cell-properties", "fo:background-color")
        if not v or v == "transparent":
            return ""
        return v.lstrip("#").upper()

    # ------------------------------------------------------------------ 段落
    def _para(self, el: El, part: str, cell: Cell | None, protected: bool) -> Para:
        p = Para(el, part, [], style=el.get("text:style-name", ""), cell=cell)
        p.el = el
        p.protected = protected
        p.heading = el.tag == _T + "h"
        style = el.get("text:style-name")
        ctx = _Ctx(underline=self.styles.underline(part, "paragraph", style), hidden=self.styles.hidden(part, "paragraph", style),
                   run=el)
        atoms: list = []
        self._inline(el, p, part, ctx, atoms)
        self._collapse(p, part, atoms)
        for sg in p.segs:  # コンテンツコントロールの中の字（空白の決まりを通した後の片）
            c = sg.field
            if c is not None and c.source == "odf-cc" and c.para is p:
                if not getattr(c, "_segs_done", False):
                    c.result_segs, c._segs_done = [], True
                c.result_segs.append(sg)
        if p.heading and not p.label:
            p.label = self._outline_label(el)
        return p

    def _outline_label(self, el: El) -> str:
        lvl = el.get("text:outline-level")
        os_ = self.styles.outline
        if not lvl or not lvl.isdigit() or os_ is None or el.get("text:is-list-header") == "true":
            return ""
        level = int(lvl)
        lv = next((x for x in os_.els() if x.get("text:level") == str(level)), None)
        if lv is None or not lv.get("style:num-format"):
            return ""
        self.outline_counter[level] += 1
        for k in range(level + 1, 11):
            self.outline_counter[k] = 0
        shown = int(lv.get("text:display-levels", "1") or 1)
        nums = [number_text(max(self.outline_counter[x], 1), NUM_FORMAT.get(lv.get("style:num-format", "1"), "decimal"))
                for x in range(max(1, level - shown + 1), level + 1)]
        return (lv.get("style:num-prefix", "") or "") + ".".join(nums) + (lv.get("style:num-suffix", "") or "")

    def _inline(self, node: El, p: Para, part: str, ctx: _Ctx, atoms: list) -> None:
        data = self.pkg.xml(part)
        for ch in node.kids:
            if isinstance(ch, Txt):
                text, offs = ch.decode(data)
                for k, c in enumerate(text):
                    atoms.append(("ch", ch, c, offs[k], offs[k + 1], ctx))
                continue
            t = ch.tag
            if t == _T + "s":
                atoms.append(("s", ch, int(ch.get("text:c", "1") or 1), ctx))
            elif t == _T + "tab":
                atoms.append(("tab", ch, ctx))
            elif t == _T + "line-break":
                atoms.append(("br", ch, ctx))
            elif t in INLINE:
                if t == _T + "a" and ch.get("xlink:href"):
                    self.__dict__.setdefault("links", []).append((text_of(ch, data), ch.get("xlink:href")))
                st = ch.get("text:style-name")
                c2 = ctx
                if st:
                    u = self.styles.prop(part, "text", st, "style:text-properties", "style:text-underline-style")
                    h = self.styles.prop(part, "text", st, "style:text-properties", "text:display")
                    c2 = _Ctx(ctx.underline if u is None else u != "none", ctx.hidden if h is None else h == "none", ctx.field, ch)
                self._inline(ch, p, part, c2, atoms)
            elif t in MARKS:
                if t == _T + "hidden-paragraph" and ch.get("text:is-hidden") == "true":
                    atoms.append(("hide-para",))
                    continue
                name = ch.get("text:name")
                if t in (_T + "bookmark-start", _T + "bookmark-end", _T + "bookmark") and name:
                    atoms.append(("mark", name, {"bookmark-start": "start", "bookmark-end": "end"}.get(ch.local, "point"), ch))
                else:
                    atoms.append(("mark", None, None, ch))
            elif t == _T + "ruby":
                base = ch.find("text:ruby-base")
                if base is not None:
                    self._inline(base, p, part, ctx, atoms)
                rt = ch.find("text:ruby-text")
                if rt is not None:
                    rtt = text_of(rt, data)
                    if rtt:
                        atoms.append(("seg", [Seg("（", "other", Src("el", part, rt.s, rt.e, el=rt), ctx.run, ctx.hidden),
                                              Seg(rtt, "ruby", Src("el", part, rt.s, rt.e, el=rt), ctx.run, ctx.hidden),
                                              Seg("）", "other", Src("el", part, rt.s, rt.e, el=rt), ctx.run, ctx.hidden)]))
            elif t == _T + "note":
                cit = ch.find("text:note-citation")
                mark = (cit.get("text:label") or text_of(cit, data)) if cit is not None else ""
                atoms.append(("seg", [Seg(f"※{mark}", "note", Src("el", part, ch.s, ch.e, el=ch), ctx.run, ctx.hidden)]))
                nb = ch.find("text:note-body")
                if nb is not None:
                    kind = "endnotes" if ch.get("text:note-class") == "endnote" else "footnotes"
                    self.notes.append((kind, self._blocks(nb, part, None, getattr(p, "protected", False))))
            elif t == _OFFICE + "annotation":
                blocks = [b for b in self._blocks(ch, part, None, True)]
                self.comments += blocks
                atoms.append(("sep",))
            elif t in SHAPES:
                self._shape(ch, p, part, ctx, atoms)
            elif t == _DRAW + "control":
                ctl = self.form_by_id.get(ch.get("draw:control", ""))
                if ctl is None:
                    atoms.append(("sep",))
                    continue
                atoms.append(("seg", self._control_segs(ctl, ch, p, part, ctx)))
            elif t in (_T + "placeholder", _T + "text-input", _T + "drop-down"):
                atoms.append(("seg", self._input_field(ch, p, part, ctx)))
            elif t == _LOEXT + "content-control":
                self._content_control(ch, p, part, ctx, atoms)
            elif t.startswith("{" + _FIELDNS + "}"):
                self._fieldmark(ch, p, part, ctx, atoms)
            elif t.startswith(_T) or t.startswith(_OFFICE):
                # フィールド（日付・ページ番号・変数など）：見えている字を、書き換えない字として読む
                shown = text_of(ch, data)
                if t == _T + "hidden-text" and ch.get("text:is-hidden") == "true":
                    continue
                if shown:
                    fc = Control("field", "odf-field", "", ch, part, p)
                    atoms.append(("seg", [Seg(shown, "t", Src("field", part, ch.s, ch.e, el=ch), ctx.run, ctx.hidden,
                                              ctx.underline, fc)]))
                else:
                    atoms.append(("sep",))
            else:
                # ほかの名前空間の要素（拡張）：中身を読む
                if ch.kids:
                    self._inline(ch, p, part, ctx, atoms)

    def _shape(self, el: El, p: Para, part: str, ctx: _Ctx, atoms: list | None = None) -> None:
        """図形・枠：テキストボックスと図形の中の字は、段落の枠（boxes）として読む。図の説明・埋め込みは字の印にする。"""
        data = self.pkg.xml(part)
        segs = []
        for ch in el.els():
            if ch.tag == _DRAW + "text-box":
                p.boxes += self._blocks(ch, part, None, getattr(p, "protected", False))
            elif ch.tag in (_T + "p", _T + "h", _T + "list", _TABLE + "table"):
                p.boxes += self._blocks(el, part, None, getattr(p, "protected", False))
                break
            elif ch.tag in SHAPES:
                self._shape(ch, p, part, ctx, atoms)
            elif ch.tag == _DRAW + "image":
                pass
            elif ch.tag in (_DRAW + "object", _DRAW + "object-ole"):
                name = (ch.get("xlink:href") or "").strip("./")
                segs.append(Seg(f"〔埋め込み: {name or 'オブジェクト'}〕", "other", Src("el", part, ch.s, ch.e, el=ch), ctx.run, ctx.hidden))
        title = el.find("svg:title")
        desc = el.find("svg:desc")
        alt = " ".join(x for x in (text_of(title, data).strip() if title is not None else "",
                                   text_of(desc, data).strip() if desc is not None else "") if x)
        if el.find("draw:image") is not None:
            # 飾りの印：LibreOffice の拡張（loext:decorative="true"）
            deco = (el.get("loext:decorative") or "").lower() == "true"
            self.pictures.append({"n": len(self.pictures) + 1, "alt": alt, "decorative": deco, "para": p})
            mark = f"〔図: {alt}〕" if alt else "" if deco else f"〔図{len(self.pictures)}（説明なし）〕"
            if mark:
                segs.insert(0, Seg(mark, "other", Src("el", part, el.s, el.e, el=el), ctx.run, ctx.hidden))
        if atoms is not None:
            atoms.append(("seg", segs) if segs else ("sep",))
        else:
            p.segs += segs

    # ------------------------------------------------------------------ 空白の決まり（6.1.2）と、字の片
    def _collapse(self, p: Para, part: str, atoms: list) -> None:
        kept: list = []
        prev_space = False
        at_start = True
        hide_para = any(a[0] == "hide-para" for a in atoms)
        for a in atoms:
            k = a[0]
            if k == "ch":
                c = a[2]
                if c in "\t\r\n":
                    c = " "
                if c == " ":
                    if at_start or prev_space:
                        continue
                    prev_space = True
                else:
                    prev_space = False
                at_start = False
                kept.append(("ch", a[1], c, a[3], a[4], a[5]))
            elif k in ("mark", "hide-para"):
                kept.append(a)
            elif k in ("fm-start", "fm-end"):
                prev_space = False
                kept.append(a)
            else:
                prev_space = False
                at_start = False
                kept.append(a)
        for i in range(len(kept) - 1, -1, -1):  # 段落の終わりの空白
            if kept[i][0] in ("mark", "hide-para", "fm-start", "fm-end"):
                continue
            if kept[i][0] == "ch" and kept[i][2] == " ":
                del kept[i]
            break
        segs: list[Seg] = []
        pos = 0  # 見える字の数（しおりの位置）
        cur = None
        fm: Control | None = None  # いま中にいる、Word の旧来の文字の欄（field:fieldmark-start〜end）
        fm_n0 = 0
        for a in kept:
            k = a[0]
            if k == "fm-start":
                fm, fm_n0, cur = a[1], len(segs), None
                continue
            if k == "fm-end":
                if fm is not None:
                    fm.fm_end = a[1]
                    fm.result_segs = segs[fm_n0:]
                    for sg in fm.result_segs:
                        sg.field = fm
                    if not fm.result_segs:
                        sg = Seg("", "field", Src("ctl", part, a[1].s, a[1].s, el=a[1]), None, False, False, fm)
                        segs.append(sg)
                        fm.result_segs = [sg]
                fm, cur = None, None
                continue
            if k == "ch":
                _k, node, c, s, e, ctx = a
                hidden = ctx.hidden or hide_para
                if cur is not None and cur[0] is node and cur[1] is ctx:
                    sg = segs[-1]
                    sg.text += c
                    sg.node.starts.append(s)
                    sg.node.ends.append(e)
                    sg.node.e = e
                else:
                    sg = Seg(c, "t", Src("txt", part, s, e, [s], [e], el=node.parent), ctx.run, hidden, ctx.underline, ctx.field)
                    segs.append(sg)
                    cur = (node, ctx)
                if not hidden:
                    pos += 1
                continue
            cur = None
            if k == "s":
                _k, el, n, ctx = a
                segs.append(Seg(" " * n, "t", Src("s", part, el.s, el.e, el=el, count=n), ctx.run, ctx.hidden or hide_para,
                                ctx.underline, ctx.field))
                pos += 0 if (ctx.hidden or hide_para) else n
            elif k in ("tab", "br"):
                _k, el, ctx = a
                segs.append(Seg("\t" if k == "tab" else "\n", k, Src(k, part, el.s, el.e, el=el), ctx.run, ctx.hidden or hide_para,
                                ctx.underline if k == "tab" else False, ctx.field))
                pos += 0 if (ctx.hidden or hide_para) else 1
            elif k == "seg":
                for sg in a[1]:
                    if hide_para:
                        sg.hidden = True
                    segs.append(sg)
                    if not sg.hidden:
                        pos += len(sg.text)
            elif k == "mark" and a[1]:
                self.marks_raw.append((a[1], a[2], p, pos))
        p.segs = segs

    # ------------------------------------------------------------------ 入力の欄（フィールド）
    def _input_field(self, el: El, p: Para, part: str, ctx: _Ctx) -> list[Seg]:
        data = self.pkg.xml(part)
        shown = text_of(el, data)
        if el.local == "placeholder":
            desc = el.get("text:description") or ""
            name = desc or shown.strip("<>＜＞ ") or "placeholder"
            ctl = Control("text", "odf-placeholder", name, el, part, p, placeholder=True, help=desc)
        elif el.local == "text-input":
            desc = el.get("text:description") or ""
            ctl = Control("text", "odf-input", desc or "入力", el, part, p, value=shown, help=desc)
        else:
            opts = [x.get("text:value", "") for x in el.findall("text:label")]
            ctl = Control("dropdown", "odf-dropdown", el.get("text:name") or "選択", el, part, p, value=shown, options=opts)
        self._add_control(ctl)
        sg = Seg(shown, "t", Src("text", part, el.s, el.e, el=el), ctx.run, ctx.hidden, ctx.underline, ctl)
        ctl.result_segs = [sg]
        return [sg]

    def _content_control(self, el: El, p: Para, part: str, ctx: _Ctx, atoms: list) -> None:
        """LibreOffice のコンテンツコントロール（loext:content-control）：文字・チェック・ドロップダウン・コンボ・日付・画像。
        名前は loext:alias（なければ loext:tag）、選択肢は loext:list-item の loext:display-text、日付の書式は loext:date-format、
        プレースホルダーを表示しているかは loext:showing-place-holder。中の字は、書式（text:span）のまま段落の字として読む。"""
        def g(k: str) -> str:
            return el.get("loext:" + k) or ""

        kind = next((kk for attr, kk in CC_KINDS if g(attr) == "true"), "richtext")
        opts = [li.get("loext:display-text") or li.get("loext:value") or "" for li in el.findall("loext:list-item")]
        ctl = Control(kind, "odf-cc", g("alias") or g("tag"), el, part, p, options=opts, checked=g("checked") == "true",
                      date_format=g("date-format"), placeholder=g("showing-place-holder") == "true",
                      lock="contentLocked" if g("lock") in ("contentLocked", "sdtContentLocked") else "")
        ctl.values = [li.get("loext:value") or "" for li in el.findall("loext:list-item")]
        self._add_control(ctl)
        n0 = len(atoms)
        self._inline(el, p, part, _Ctx(ctx.underline, ctx.hidden, ctl, ctx.run), atoms)
        if not any(a[0] in ("ch", "s", "tab", "br", "seg") for a in atoms[n0:]):
            # 中が空：書く所は、要素の終わりのタグの前
            at = el.es if el.es != el.e else el.e
            atoms.append(("seg", [Seg("", "field", Src("ctl", part, at, at, el=el), ctx.run, ctx.hidden, False, ctl)]))

    def _fieldmark(self, el: El, p: Para, part: str, ctx: _Ctx, atoms: list) -> None:
        """Word の旧来のフォームの欄を、LibreOffice が書いた形（field:fieldmark）：チェック・ドロップダウンは1つの印、
        文字の欄は field:fieldmark-start と field:fieldmark-end のあいだの字。"""
        typ = el.get("{" + _FIELDNS + "}type", "") or ""
        # 欄の名前：LibreOffice は text:name に書く（古い形は field:name）
        name = el.get("text:name", "") or el.get("{" + _FIELDNS + "}name", "") or ""
        params = {}
        for prm in el.els():
            if prm.local == "param":
                params.setdefault(prm.get("{" + _FIELDNS + "}name", ""), []).append(prm.get("{" + _FIELDNS + "}value", ""))
        if el.local == "fieldmark" and typ.endswith("FORMCHECKBOX"):
            on = (params.get("Checkbox_Checked") or ["false"])[0] == "true"
            ctl = Control("checkbox", "odf-fieldmark", name, el, part, p, checked=on)
            self._add_control(ctl)
            sg = Seg("☑" if on else "☐", "field", Src("ctl", part, el.s, el.e, el=el), ctx.run, ctx.hidden, False, ctl)
            ctl.result_segs = [sg]
            atoms.append(("seg", [sg]))
        elif el.local == "fieldmark" and typ.endswith("FORMDROPDOWN"):
            opts = params.get("Dropdown_ListEntry") or []
            sel = (params.get("Dropdown_Selected") or [""])[0]
            cur = opts[int(sel)] if sel.isdigit() and int(sel) < len(opts) else ""
            ctl = Control("dropdown", "odf-fieldmark", name, el, part, p, value=cur, options=opts)
            self._add_control(ctl)
            sg = Seg(cur, "field", Src("ctl", part, el.s, el.e, el=el), ctx.run, ctx.hidden, False, ctl)
            ctl.result_segs = [sg]
            atoms.append(("seg", [sg]))
        elif el.local == "fieldmark-start" and typ.endswith("FORMTEXT"):
            ctl = Control("text", "odf-fieldmark", name, el, part, p)
            self._add_control(ctl)
            ctl.fm_start = el
            ctl.fm_end = None
            atoms.append(("fm-start", ctl))
        elif el.local == "fieldmark-end":
            atoms.append(("fm-end", el))
        else:
            atoms.append(("sep",))

    # ------------------------------------------------------------------ フォームの部品（office:forms）
    def _forms(self, office_text: El) -> None:
        forms = office_text.find("office:forms")
        if forms is None:
            return
        data = self.pkg.xml(self.part)
        for f in forms.iter("form:form"):
            for el in f.els():
                if not el.tag.startswith(_FORM) or el.local in ("form", "properties", "event-listeners", "connection-resource"):
                    continue
                self._form_control(el, data)

    def _form_control(self, el: El, data: bytes) -> None:
        loc = el.local
        name = el.get("form:name", "") or ""
        cid = el.get("xml:id") or el.get("form:id") or ""
        help_ = el.get("form:title", "") or ""
        locked = "contentLocked" if el.get("form:readonly") == "true" or el.get("form:disabled") == "true" else ""
        if loc in ("text", "textarea", "formatted-text", "number", "date", "time", "password", "combobox"):
            cur = el.get("form:current-value")
            if cur is None:
                cur = el.get("form:value", "")
            if loc == "textarea" and not cur and el.find("text:p") is not None:
                cur = "\n".join(text_of(x, data) for x in el.findall("text:p"))
            kind = {"combobox": "combo", "date": "date"}.get(loc, "text")
            ctl = Control(kind, "odf-form", name, el, self.part, None, value=cur, help=help_, lock=locked)
            ml = el.get("form:max-length")
            if ml and ml.isdigit() and int(ml) > 0:
                ctl.maxlen = int(ml)
            if loc == "combobox":
                ctl.options = [x.get("form:label", "") for x in el.findall("form:item")]
            ctl.form_kind = loc
        elif loc == "checkbox":
            st = el.get("form:current-state") or el.get("form:state") or "unchecked"
            ctl = Control("checkbox", "odf-form", name, el, self.part, None, checked=st == "checked", help=el.get("form:label", "") or help_,
                          lock=locked)
            ctl.form_kind = loc
        elif loc == "listbox":
            opts = el.findall("form:option")
            labels = [o.get("form:label", "") or o.get("form:value", "") for o in opts]
            sel = [o for o in opts if (o.get("form:current-selected") or o.get("form:selected")) == "true"]
            cur = (sel[0].get("form:label", "") or sel[0].get("form:value", "")) if sel else None
            ctl = Control("dropdown", "odf-form", name, el, self.part, None, value=cur, options=labels, help=help_, lock=locked)
            ctl.form_kind = loc
            ctl.multiple = el.get("form:multiple") == "true"
        elif loc == "radio":
            g = self.radio_groups.get(name)
            label = el.get("form:label", "") or el.get("form:value", "")
            if g is None:
                g = Control("dropdown", "odf-radio", name, el, self.part, None, help=help_, lock=locked)
                g.form_kind = "radio"
                g.radios = []
                self.radio_groups[name] = g
                self._add_control(g)
            g.radios.append(el)
            g.options.append(label)
            if (el.get("form:current-selected") or el.get("form:selected")) == "true":
                g.value = label
            g.result_segs = [Seg(g.value or "", "field", None, None, False, False, g)]
            if cid:
                self.form_by_id[cid] = g
            return
        elif loc == "fixed-text":
            ctl = Control("label", "odf-label", name, el, self.part, None, value=el.get("form:label", ""))
            ctl.form_kind = loc
            if cid:
                self.form_by_id[cid] = ctl
            return
        else:
            ctl = Control("other", "odf-other", name, el, self.part, None, value=el.get("form:label", ""))
            ctl.form_kind = loc
            if cid:
                self.form_by_id[cid] = ctl
            return
        self._add_control(ctl)
        if cid:
            self.form_by_id[cid] = ctl

    def _add_control(self, ctl: Control) -> None:
        self.controls.append(ctl)

    def _control_segs(self, ctl: Control, anchor: El, p: Para, part: str, ctx: _Ctx) -> list[Seg]:
        src = Src("ctl", part, anchor.s, anchor.e, el=anchor)
        if ctl.source == "odf-label" or ctl.source == "odf-other":
            return [Seg(ctl.value or "", "other", src, ctx.run, ctx.hidden)] if ctl.value else []
        if ctl.source == "odf-radio":
            # オプションボタン：それぞれの位置に、選んだかどうかと見出しを出す（欄は組の名前で1つ）
            idx = next((k for k, r in enumerate(ctl.radios) if (r.get("xml:id") or r.get("form:id")) == anchor.get("draw:control")), 0)
            label = ctl.options[idx] if idx < len(ctl.options) else ""
            mark = "◉" if ctl.value is not None and label == ctl.value else "○"
            if ctl.para is None:
                ctl.para = p
            return [Seg(mark + label, "other", src, ctx.run, ctx.hidden)]
        if ctl.para is None:
            ctl.para = p
        if ctl.kind == "checkbox":
            sg = Seg("☑" if ctl.checked else "☐", "field", src, ctx.run, ctx.hidden, False, ctl)
            ctl.result_segs = [sg]
            out = [sg]
            if ctl.help:
                out.append(Seg(ctl.help, "other", src, ctx.run, ctx.hidden))
            return out
        sg = Seg(ctl.value or "", "field", src, ctx.run, ctx.hidden, False, ctl)
        ctl.result_segs = [sg]
        return [sg]


def protected_paras(doc: OdtDocument) -> set[int]:
    """保護された節・表の中の段落（id）。"""
    out = set()
    for _k, _part, blocks in doc.all_blocks():
        for p in iter_paras(blocks):
            if getattr(p, "protected", False):
                out.add(id(p))
    return out
