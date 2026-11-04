"""Word に書き込む：書いた欄の部分だけを、元の XML のバイト列への小さな差分として変える。

- 表の空のマス：最初の段落に、段落記号の書式を引き継いだ文字を足す
- 段落の中の空欄：空白を値に置き換える（その部分だけ文字の要素を分ける）。下線を引いた空白は、下線の長さを保つ。
  align（left・center・right）を指定したときは、空きの幅を保って、そろえ方のとおりに前後へ半角の空白を足す（inline_padding）
- □：■（☐ は ☑）にする
- 旧来のフォームの欄・コンテンツコントロール：Word と同じ形で値と状態を書く
書いた所は見えないしおり（_todoku_…）で囲み、書いた後も同じ欄の名前で読み戻せるようにする。
埋め込んだ Excel の表のセル（E1!Sheet1!C4）は todoku.embed で書く（ブックの部品と、表示・印刷に使う表の絵の部品）。
文書の保護（フォームだけ・読み取り専用）は守る（編集できる範囲の外には書かない）。
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from .document import WINGDINGS, DocxDocument, Para
from .window import CHECKED, MARK_PREFIX, Spots, box_mark
from .xmlpos import Node, escape, prefix_of


class FillError(ValueError):
    pass


@dataclass
class Patch:
    id: str
    kind: str
    value: object
    part: str = ""
    anchor: dict | None = None  # 字を目印にした書き込み（{"old": …, "new": …}）。証明で、同じ決め方で位置を確かめる
    key: str = ""  # 指定したときの名前（「@1」など）
    align: dict | None = None  # 空きの中のそろえ方の指定（{"align", "ambiguous_width"}）と、足した空白（"pad": [前, 後ろ] の字数）


# 書いた字に引き継がない文字の書式（w:rPr の子）：消したように見える（取り消し線・非表示）・目立つ（字の色・蛍光ペン・網かけ）もの。
# 空のマスの段落記号・空白の字の要素に残った書式（見本の字を消した跡など）が、本人の名前に付かないように。値が「なし」のものは残す
_DROP_RPR = {"strike": "取り消し線", "dstrike": "二重取り消し線", "vanish": "非表示", "specVanish": "非表示",
             "webHidden": "非表示（Web）", "color": "字の色", "highlight": "蛍光ペン", "shd": "網かけ"}
_RPR_CHILD = re.compile(rb"<(?:[\w.-]+:)?(strike|dstrike|vanish|specVanish|webHidden|color|highlight|shd)\b([^>]*?)(?:/>|>.*?</(?:[\w.-]+:)?\1>)",
                        re.S)


def _rpr_on(name: str, attrs: bytes) -> bool:
    """rPr の子の要素が、見た目を変える値か（「なし」・自動・黒でないか）。"""
    val = re.search(rb'(?:\w+:)?val="([^"]*)"', attrs)
    v = val[1].decode().lower() if val else None
    if name in ("strike", "dstrike", "vanish", "specVanish", "webHidden"):
        return v not in ("false", "0", "off")
    if name == "color":
        return v not in ("auto", "000000", None) or b"themeColor" in attrs
    if name == "highlight":
        return v not in ("none", None)
    if name == "shd":
        fill = re.search(rb'(?:\w+:)?fill="([^"]*)"', attrs)
        f = fill[1].decode().lower() if fill else "auto"
        return f not in ("auto", "ffffff", "") or v not in ("clear", "nil", None)
    return False


def clean_rpr(rpr: bytes) -> tuple[bytes, list[str]]:
    """書いた字に付ける書式（rPr のバイト列）：_DROP_RPR の、見た目を変える値の要素を除いたものと、除いた書式の名前。
    子の要素のなくなった rPr は b""（段落・スタイルの既定の書式）。証明（docx/proof.py）も同じ関数で、認める書式を作る。"""
    if not rpr:
        return rpr, []
    names: list[str] = []

    def drop(m: re.Match) -> bytes:
        name = m[1].decode()
        if _rpr_on(name, m[2]):
            names.append(_DROP_RPR[name])
            return b""
        return m[0]

    out = _RPR_CHILD.sub(drop, rpr)
    if re.fullmatch(rb"<(?:[\w.-]+:)?rPr\b[^>]*>\s*</(?:[\w.-]+:)?rPr>", out):
        out = b""
    return out, list(dict.fromkeys(names))


def written_formatting(blob: bytes, name: str = "") -> list[dict]:
    """書いた後の文書を読み直し、書いた所（しおり _todoku_…）の字の書式に、消したように・目立って見えるもの（_DROP_RPR の
    「なし」でない値）があれば [{id, formatting}]（読み戻しで伝える）。"""
    from ..zipraw import RawZip

    z = RawZip(blob)
    out = []
    start = re.compile(rb'<(?:[\w.-]+:)?bookmarkStart\b[^>]*?\bw?:?id="(\d+)"[^>]*?\bw?:?name="' + re.escape(MARK_PREFIX.encode())
                       + rb'([^"]*)"[^>]*/>')
    for part in z.names():
        if not (part.startswith("word/") and part.endswith(".xml")):
            continue
        xml = z.read(part)
        if MARK_PREFIX.encode() not in xml:
            continue
        for m in start.finditer(xml):
            end = re.search(rb'<(?:[\w.-]+:)?bookmarkEnd\b[^>]*?\bid="' + m[1] + rb'"', xml[m.end():])
            body = xml[m.end():m.end() + end.start()] if end else b""
            found = []
            for rp in re.finditer(rb"<(?:[\w.-]+:)?rPr\b[^>]*>.*?</(?:[\w.-]+:)?rPr>", body, re.S):
                for c in _RPR_CHILD.finditer(rp[0]):
                    if _rpr_on(c[1].decode(), c[2]):
                        found.append(_DROP_RPR[c[1].decode()])
            if found:
                out.append({"id": m[2].decode(), "formatting": list(dict.fromkeys(found))})
    return out


def written_of(patches) -> dict:
    """証明に渡す、書いた欄 → 値（字を目印にした書き込みは目印も、そろえ方を指定した書き込みはその指定も渡す）。"""
    def one(p):
        if p.anchor:
            return {"@": p.anchor, "value": p.value}
        if getattr(p, "align", None):  # ほかの形式の Patch には、そろえ方の指定がない
            return {"value": p.value, **{k: v for k, v in p.align.items() if k != "pad"}}
        return p.value
    return {p.id: one(p) for p in patches}


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class Editor:
    def __init__(self, xml: bytes):
        self.xml = xml
        self.edits: list[tuple[int, int, bytes]] = []
        self.pre = prefix_of(xml)

    def put(self, s: int, e: int, b: bytes) -> None:
        for a, z, _ in self.edits:
            if s < z and a < e or (s == e == a):
                if not (s == e and (s <= a or s >= z)):
                    raise FillError("同じ所に2つの書き込みが重なりました")
        self.edits.append((s, e, b))

    def build(self) -> bytes:
        out = bytearray()
        at = 0
        for s, e, b in sorted(self.edits, key=lambda x: (x[0], x[1])):
            out += self.xml[at:s]
            out += b
            at = e
        out += self.xml[at:]
        return bytes(out)


def _t(pre: str, text: str) -> bytes:
    """文字の要素。改行は改行の要素に、空白が端にあるときは xml:space=preserve を付ける。"""
    out = bytearray()
    for i, line in enumerate(text.split("\n")):
        if i:
            out += f"<{pre}:br/>".encode()
        if line:
            out += f'<{pre}:t xml:space="preserve">'.encode() + escape(line) + f"</{pre}:t>".encode()
    return bytes(out)


def inline_padding(p: Para, a: int, b: int, value: str, opts: dict | None = None, anchored: bool = False) -> tuple[str, str]:
    """段落の中の空欄 [a, b) に value を書くときに、前と後ろに足す半角の空白（書き込みと証明が同じ決まりで使う）。
    そろえ方の指定（opts の align・ambiguous_width）があれば、空きの幅を保って、そのとおりに足す（textwidth.gap_padding）。
    なければ、空欄の字がすべて下線つきのときだけ（字を目印にした書き込みは除く）、下線の長さを保つよう後ろに足す（左にそろえる）。
    曖昧な幅（A）の字は、指定がなければ、段落と値の並びに全角の字があるときだけ全角として数える（UAX #11）。"""
    from ..textwidth import AMBIGUOUS, gap_padding

    if not value:
        return "", ""
    opts = opts or {}
    if not opts:
        pos = p.positions()
        blank = [p.segs[pos[i][0]] for i in range(a, b) if i < len(pos)]
        if anchored or not (blank and all(sg.underline for sg in blank)):
            return "", ""
    text = p.text
    return gap_padding(text[a:b], value, opts.get("align", "left"), context=text,
                       ambiguous_wide=AMBIGUOUS.get(opts.get("ambiguous_width")))


def _is_text_value(v) -> str:
    if isinstance(v, dict):
        raise FillError(f"この所には、字（または数・真偽）で書いてください（受け取った指定 {sorted(v)} は、この所では使えません）")
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return "" if v is None else str(v)


def _as_bool(v) -> bool:
    """チェックの欄の値：true / false と、箱の字（☐・☑ など）だけ（inputcheck.checkbox_value。言葉の意味では決めない）。"""
    from ..inputcheck import GLYPH_OFF, GLYPH_ON, InputError, checkbox_value

    try:
        return checkbox_value(v, GLYPH_ON, GLYPH_OFF)
    except InputError as e:
        raise FillError(str(e)) from None


# 要素の w:id（しおり・変更履歴・コメントの範囲・脚注など、番号の要素すべて。数だけの番号）
_ANNOT_ID = re.compile(rb'\s(?:[\w.-]+:)?id="(\d+)"')
# 変更履歴の入れ物：挿入（ins・moveTo）と削除（del・moveFrom）
REVISION_WRAPPERS = ("ins", "del", "moveTo", "moveFrom")


def revision_wrappers(node: Node) -> list[Node]:
    """node（字の要素など）を囲む、段落の中の変更履歴の入れ物（内側から順に）。"""
    out = []
    x = node.parent
    while x is not None and x.local not in ("p", "body", "tc", "txbxContent", "#root"):
        if x.local in REVISION_WRAPPERS:
            out.append(x)
        x = x.parent
    return out


def revision_counts(doc) -> dict[str, int]:
    """文書（本文・ヘッダー・脚注など）に残っている、承諾も却下もされていない変更履歴の数：{挿入, 削除, 書式}。"""
    n = {"挿入": 0, "削除": 0, "書式": 0}
    for part in doc.parts:
        xml = doc.xml(part)
        n["挿入"] += len(re.findall(rb"<(?:[\w.-]+:)?(?:ins|moveTo)\b[^>]*>", xml))
        n["削除"] += len(re.findall(rb"<(?:[\w.-]+:)?(?:del|moveFrom)\b[^>]*>", xml))
        n["書式"] += len(re.findall(rb"<(?:[\w.-]+:)?(?:rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)\b", xml))
    return n


class Filler:
    def __init__(self, doc: DocxDocument, force: bool = False):
        self.doc = doc
        self.force = force
        self.sp = Spots(doc)
        self.editors: dict[str, Editor] = {}
        self.patches: list[Patch] = []
        self.warnings: list[str] = []
        # しおりの番号：しおりだけでなく、変更履歴（w:ins・w:del・書式の変更）・コメントの範囲などの w:id とも重ねない
        # （Word はこれらの番号を同じ数の並びとして扱い、重なると開くときに修復を求めることがある）
        self.next_bm = 1 + max([int(m) for part in doc.parts for m in _ANNOT_ID.findall(doc.xml(part))] + [0])
        self.run_ops: dict[tuple[str, int], tuple[Node, list]] = {}
        self.bm_ids: dict[int, str] = {}
        self.bound: dict[tuple[str, int, int], tuple[str, str, object]] = {}  # 結び付けた XML の値の所 → (欄, 書く値, 所)
        self._perm_ranges()

    def ed(self, part: str) -> Editor:
        if part not in self.editors:
            self.editors[part] = Editor(self.doc.xml(part))
        return self.editors[part]

    # ------------------------------------------------------------------ 保護
    def _perm_ranges(self) -> None:
        self.ranges: dict[str, list[tuple[int, int]]] = {}
        for part in self.doc.parts:
            xml = self.doc.xml(part)
            starts = {}
            out = []
            for m in re.finditer(rb'<(?:\w+:)?perm(Start|End)\b[^>]*?\bw?:?id="([^"]+)"[^>]*>', xml):
                if m[1] == b"Start":
                    starts[m[2]] = m.start()
                elif m[2] in starts:
                    out.append((starts.pop(m[2]), m.end()))
            self.ranges[part] = out

    def check_editable(self, item: dict, part: str, pos: int) -> None:
        doc = self.doc
        if self.force or not doc.protection_enforced or not doc.protection:
            return
        if doc.protection == "forms":
            if item.get("control") is None:
                raise FillError(f"{item['id']}: この文書は「フォームの欄だけ書ける」ように保護されています（ここは欄ではありません）")
            return
        if doc.protection == "readOnly":
            if any(a <= pos <= b for a, b in self.ranges.get(part, [])):
                return
            raise FillError(f"{item['id']}: この文書は読み取り専用で保護されていて、ここは編集できる範囲の外です")
        raise FillError(f"{item['id']}: この文書は保護されていて（{doc.protection}）、書き込めません")

    # ------------------------------------------------------------------ しおり
    def bookmark(self, name: str) -> tuple[bytes, bytes]:
        pre = "w"
        bid = self.next_bm
        self.next_bm += 1
        bm = MARK_PREFIX + name.replace(".", "_").replace("#", "_")
        self.bm_ids[bid] = name  # 警告では、しおりの名前（T1_R2_C2）でなく欄の名前（T1.R2.C2）で伝える
        return (f'<{pre}:bookmarkStart {pre}:id="{bid}" {pre}:name="{bm}"/>'.encode(),
                f'<{pre}:bookmarkEnd {pre}:id="{bid}"/>'.encode())

    # ------------------------------------------------------------------ 段落の文字を置き換える
    def replace_span(self, item: dict, p: Para, a: int, b: int, value: str, pad: tuple[str, str] = ("", "")) -> None:
        """段落の文字 [a, b) を value にする予定を立てる（同じ文字の要素の中の複数の空欄は、finish でまとめて書く）。
        pad：value の前と後ろに足す空白（しおりの外。inline_padding）。"""
        part = p.part
        pos = p.positions()
        if a == b:
            # 挿入（「市」の後ろに書き足すなど）：直前の字の後ろに入れる。段落の頭（「年　月」の年の前）は、初めの字の前に入れる
            if not pos:
                self.fill_cell(item, value)  # 字の要素のない段落：段落の印の書式で、字の要素を作って書く
                return
            i, off = pos[a - 1] if a > 0 else pos[0]
            at = off + 1 if a > 0 else off
            sg = p.segs[i]
            if sg.kind in ("tab", "sym") and a > 0 and sg.run is not None:
                nxt = p.segs[pos[a][0]] if a < len(pos) else None
                if nxt is not None and nxt.kind == "t" and nxt.node is not None and nxt.run is not None:
                    i, at, sg = pos[a][0], pos[a][1], nxt  # 直前がタブ・記号：次の字の前に入れる
                elif nxt is None:
                    self.append_run(item, p, sg.run, value)  # 段落の終わりのタブ・記号の後ろ：同じ書式の字の要素を足す
                    return
            if sg.kind in ("note", "other") and a == len(pos) and sg.run is not None:
                # 段落の終わりの脚注の印・図の印（「人口※4」）の後ろ：本文の字の書式（印の上付きでなく）で、字の要素を足す
                body = next((x.run for x in reversed(p.segs[:i]) if x.kind == "t" and x.run is not None), None)
                self.append_run(item, p, sg.run, value, style_run=body or False)
                return
            if sg.kind != "t" or sg.node is None or sg.run is None:
                raise FillError(f"{item['id']}: 書き足す位置が、書き換えられない部分にあります")
            self.check_editable(item, part, sg.node.s)
            bm = self.bookmark(item["id"])
            self.run_ops.setdefault((part, sg.run.s), (sg.run, []))[1].append((sg.node, at, at, value, bm, ("", "")))
            return
        if a >= len(pos) or b > len(pos):
            raise FillError(f"{item['id']}: 空欄の位置が見つかりません")
        segs_idx = []
        for k in range(a, b):
            i, _off = pos[k]
            if i not in segs_idx:
                segs_idx.append(i)
        for i in segs_idx:
            sg = p.segs[i]
            if sg.kind not in ("t", "sym") or sg.node is None or sg.run is None:
                raise FillError(f"{item['id']}: 空欄が、書き換えられない部分（{sg.kind}）にかかっています")
        self.check_editable(item, part, p.segs[segs_idx[0]].node.s)
        bm = self.bookmark(item["id"])
        for n, i in enumerate(segs_idx):
            sg = p.segs[i]
            ks = [off for k in range(a, b) for si, off in [pos[k]] if si == i]
            k0, k1 = min(ks), max(ks) + 1
            self.run_ops.setdefault((part, sg.run.s), (sg.run, []))[1].append(
                (sg.node, k0, k1, value if n == 0 else None, bm if n == 0 else None, pad if n == 0 else ("", "")))

    def finish_runs(self) -> None:
        """文字の要素を分けて、値としおりを入れる（1つの文字の要素ごとに1回で書く）。"""
        for (part, _s), (run, ops) in self.run_ops.items():
            ed = self.ed(part)
            pre = ed.pre
            xml = ed.xml
            # 変更履歴の中の字に書く：書いた字を変更履歴の中に入れない（ほかの人の挿入として承諾・却下されないように）。
            # 挿入（w:ins）の中の字は、その挿入をそこで分けて、書いた字を挿入の外に置く。削除の中・入れ子の中には書かない
            wraps = revision_wrappers(run)
            split = None
            if wraps:
                ids = [self.bm_ids.get(int(m_[1]), "") for op in ops if op[4] for m_ in [re.search(rb'\bw?:?id="(\d+)"', op[4][0])] if m_]
                who = "・".join(str(i) for i in ids if i) or "書いた所"
                w0 = wraps[0]
                if len(wraps) > 1 or w0.local != "ins" or run.parent is not w0:
                    kinds = "・".join(dict.fromkeys(w.local for w in wraps))
                    raise FillError(f"{who}: 書く所が、変更履歴（{kinds}）の中にあります。変更履歴の中には書きません"
                                    "（Word で変更履歴を承諾するか元に戻してから、もう一度書いてください）")
                tag = xml[w0.s:w0.se]
                close = xml[w0.es:w0.e]

                def reopen(tag=tag):
                    nid = self.next_bm  # 分けた後ろの挿入には、新しい番号（書き込みごと）
                    self.next_bm += 1
                    return re.sub(rb'(\s(?:[\w.-]+:)?id=")\d+(")', lambda m_: m_[1] + str(nid).encode() + m_[2], tag, count=1)

                split = (close, reopen, w0.get("author", ""))
                self.revision_split = getattr(self, "revision_split", []) + [(who, split[2])]
            rpr = run.find("rPr")
            rpr_b = xml[rpr.s:rpr.e] if rpr is not None else b""
            rpr_v, dropped = clean_rpr(rpr_b)  # 書いた値の字の書式（取り消し線・字の色などは引き継がない）
            r_open = xml[run.s:run.se] if run.se != run.e else f"<{pre}:r>".encode()
            r_close = f"</{pre}:r>".encode()
            out = bytearray()
            buf = bytearray()
            if dropped and any(op[3] is not None for op in ops):
                self.note_dropped([self.bm_ids.get(int(m_[1]), "") for op in ops if op[3] is not None and op[4]
                                   for m_ in [re.search(rb'\bw?:?id="(\d+)"', op[4][0])] if m_], dropped)

            def flush():
                nonlocal buf
                if buf:
                    out.extend(r_open + rpr_b + bytes(buf) + r_close)
                    buf = bytearray()

            by_node: dict[int, list] = {}
            for node, k0, k1, value, bm, pad in ops:
                by_node.setdefault(id(node), []).append((k0, k1, value, bm, pad))
            for kid in run.kids:
                if kid.local == "rPr":
                    continue
                todo = sorted(by_node.get(id(kid), []), key=lambda x: x[0])
                if not todo:
                    buf += xml[kid.s:kid.e]
                    continue
                if kid.local == "sym":
                    k0, k1, value, bm, _pad = todo[0]
                    sym = xml[kid.s:kid.e]
                    if value in ("■", "☑", "☒"):
                        sym = re.sub(rb'(char=")[0-9A-Fa-f]+(")', rb"\g<1>F0FE\g<2>", sym)
                    elif value in ("□", "☐"):
                        # チェックを外す：Wingdings などの記号の字（私用の領域）は Wingdings の空の箱、ふつうの字はその字に
                        m_ = re.search(rb'char="([0-9A-Fa-f]+)"', sym)
                        code = int(m_[1], 16) if m_ else 0
                        pua = code >= 0xE000 or code < 0x100
                        was = WINGDINGS.get(code if code >= 0xF000 else code + 0xF000) if pua else chr(code)
                        if was in CHECKED:  # 印のない箱のチェックを外す（変えない）ときは、元の字のまま
                            to = b"F0A8" if pua else f"{ord(value):04X}".encode()
                            sym = re.sub(rb'(char=")[0-9A-Fa-f]+(")', rb"\g<1>" + to + rb"\g<2>", sym)
                    flush()
                    out.extend((split[0] if split else b"") + (bm[0] if bm else b"") + r_open + rpr_b + sym + r_close
                               + (bm[1] if bm else b"") + (split[1]() if split else b""))
                    continue
                text = kid.text
                cur = 0
                for k0, k1, value, bm, pad in todo:
                    buf += _t(pre, text[cur:k0])
                    if value is not None:
                        buf += _t(pre, pad[0])  # 前に足す空白（右・中央にそろえるとき）：元の字の要素の書式のまま、しおりの外
                        flush()
                        out.extend((split[0] if split else b"") + (bm[0] if bm else b"") + r_open + rpr_v + _t(pre, value)
                                   + r_close + (bm[1] if bm else b"") + (split[1]() if split else b""))
                        buf += _t(pre, pad[1])
                    cur = k1
                buf += _t(pre, text[cur:])
            flush()
            ed.put(run.s, run.e, bytes(out))

    # ------------------------------------------------------------------ 箇条書きの印の □
    def fill_bullet(self, item: dict, on: bool) -> None:
        """箇条書きの印の □ を、この段落だけ ■（☑）にする（todoku.docx.bullet）。チェックを外す（False）のは、白紙の様式では
        何も変えない。"""
        from .bullet import PART, BulletError, level_of, checked_lvl

        p: Para = item["para"]
        bl = item.get("bullet")
        if item.get("filled") or bl is None:
            raise FillError(f"{item['id']}: すでに書き込んだ欄です（書き直すには、元の様式に書いてください）")
        if not on:
            return
        self.check_editable(item, p.part, p.node.s)
        if PART not in self.doc.parts:
            raise FillError(f"{item['id']}: 段落番号の定義（{PART}）がありません")
        try:
            checked_lvl(self.doc.xml(PART), level_of(self.doc.xml(PART), bl["num"], bl["ilvl"])[1])
        except BulletError as e:
            raise FillError(f"{item['id']}: {e}") from None
        self.bullets = getattr(self, "bullets", [])
        self.bullets.append(item)

    def finish_bullets(self) -> None:
        """足す w:num をまとめて numbering.xml に入れ、段落の w:numId を、足した w:num にする。段落の頭にしおりを置く。"""
        from .bullet import PART, insert_at, new_num, next_ids

        items = getattr(self, "bullets", [])
        if not items:
            return
        xml = self.doc.xml(PART)
        ed = self.ed(PART)
        ids = next_ids(xml, len(items))
        body = b"".join(new_num(xml, ed.pre, it["bullet"]["num"], it["bullet"]["ilvl"], nid) for it, nid in zip(items, ids))
        at = insert_at(xml)
        ed.put(at, at, body)
        for it, nid in zip(items, ids):
            p: Para = it["para"]
            ped = self.ed(p.part)
            pre = ped.pre
            pxml = ped.xml
            node = p.node
            ppr = node.find("pPr")
            npr = ppr.find("numPr") if ppr is not None else None
            bm = self.bookmark(it["id"])
            if npr is not None and npr.find("numId") is not None:
                ni = npr.find("numId")
                ped.put(ni.s, ni.e, re.sub(rb'(\b(?:\w+:)?val=")[^"]*(")', lambda m: m[1] + nid.encode() + m[2],
                                         pxml[ni.s:ni.e], count=1))
            else:
                add = f'<{pre}:numPr><{pre}:ilvl {pre}:val="{it["bullet"]["ilvl"]}"/><{pre}:numId {pre}:val="{nid}"/></{pre}:numPr>'.encode()
                if ppr is None:
                    raise FillError(f"{it['id']}: 段落の書式（w:pPr）がありません")
                # CT_PPrBase の並び：pStyle・keepNext・keepLines・pageBreakBefore・framePr・widowControl の後ろに numPr
                before = [k for k in ppr.kids if k.local in ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
                                                             "widowControl")]
                if npr is not None:
                    ped.put(npr.s, npr.e, add)
                elif before:
                    ped.put(before[-1].e, before[-1].e, add)
                elif ppr.se == ppr.e:
                    ped.put(ppr.s, ppr.e, pxml[ppr.s:ppr.e][:-2].rstrip() + b">" + add + f"</{pre}:pPr>".encode())
                    ppr = None
                else:
                    ped.put(ppr.se, ppr.se, add)
            at = ppr.e if ppr is not None else node.find("pPr").e
            ped.put(at, at, bm[0] + bm[1])

    def note_dropped(self, ids: list, names: list[str]) -> None:
        """引き継がなかった書式を、警告で伝える（書いた字の書式を、黙って決めない）。"""
        shown = "・".join(str(i) for i in ids if i) or "書いた所"
        self.warnings.append(f"{shown}: 書く所に残っていた書式（{'・'.join(names)}）は、書いた字に引き継がず、"
                             "段落・スタイルの既定の書式（字の大きさ・書体はそのまま）で書きました")

    # ------------------------------------------------------------------ 表の空のマス
    def fill_cell(self, item: dict, value: str) -> None:
        p: Para = item["para"]
        part = p.part
        ed = self.ed(part)
        pre = ed.pre
        xml = ed.xml
        node = p.node
        self.check_editable(item, part, node.s)
        ppr = node.find("pPr")
        rpr_b = b""
        if ppr is not None:
            mark = ppr.find("rPr")
            if mark is not None:
                inner = b"".join(xml[x.s:x.e] for x in mark.kids if x.local not in ("ins", "del", "moveFrom", "moveTo"))
                rpr_b = f"<{pre}:rPr>".encode() + inner + f"</{pre}:rPr>".encode() if inner else b""
        rpr_b, dropped = clean_rpr(rpr_b)  # 空のマスに残った取り消し線・字の色などは引き継がない
        if dropped:
            self.note_dropped([item["id"]], dropped)
        start_bm, end_bm = self.bookmark(item["id"])
        run = start_bm + f"<{pre}:r>".encode() + rpr_b + _t(pre, value) + f"</{pre}:r>".encode() + end_bm
        if node.se == node.e:
            open_tag = xml[node.s:node.e]
            open_tag = open_tag[:-2].rstrip() + b">"
            ed.put(node.s, node.e, open_tag + run + f"</{pre}:p>".encode())
        else:
            at = ppr.e if ppr is not None else node.se
            ed.put(at, at, run)

    def append_run(self, item: dict, p: Para, after_run, value: str, style_run=None) -> None:
        """字の要素 after_run のすぐ後ろに、同じ書式（rPr）の字の要素を足して書く（タブで終わる段落の後ろなど）。
        style_run：書式を写す字の要素（after_run と違うとき。False なら書式を写さない）。"""
        part = p.part
        ed = self.ed(part)
        pre = ed.pre
        xml = ed.xml
        self.check_editable(item, part, after_run.s)
        if (part, after_run.s) in self.run_ops:
            raise FillError(f"{item['id']}: 同じ字の要素の中に、ほかの書き込みがあります")
        src = after_run if style_run is None else style_run
        rpr = src.find("rPr") if src is not False else None
        rpr_b = xml[rpr.s:rpr.e] if rpr is not None else b""
        rpr_b, dropped = clean_rpr(rpr_b)
        at = after_run.e
        wraps = revision_wrappers(after_run)
        if wraps:
            # 変更履歴の中の字の後ろ：書いた字は、変更履歴の入れ物の外（すぐ後ろ）に置く。入れ物の中で後ろに字があれば書かない
            outer = wraps[-1]
            if re.search(rb"<(?:[\w.-]+:)?(?:r|t|tab|sym)\b", xml[after_run.e:outer.es]):
                raise FillError(f"{item['id']}: 書く所が、変更履歴（{outer.local}）の中にあります。変更履歴の中には書きません"
                                "（Word で変更履歴を承諾するか元に戻してから、もう一度書いてください）")
            at = outer.e
        if dropped:
            self.note_dropped([item["id"]], dropped)
        start_bm, end_bm = self.bookmark(item["id"])
        ed.put(at, at,
               start_bm + f"<{pre}:r>".encode() + rpr_b + _t(pre, value) + f"</{pre}:r>".encode() + end_bm)

    # ------------------------------------------------------------------ フォームの欄・コンテンツコントロール
    def fill_control(self, item: dict, value) -> None:
        c = item["control"]
        part = c.part
        ed = self.ed(part)
        pre = ed.pre
        xml = ed.xml
        inner = [sg.node.s for sg in c.result_segs if sg.node is not None]
        self.check_editable(item, part, inner[0] if inner else c.node.s)
        wraps = revision_wrappers(c.node) + [w for sg in c.result_segs if sg.node is not None for w in revision_wrappers(sg.node)]
        if wraps:
            raise FillError(f"{item['id']}: この欄は、変更履歴（{'・'.join(dict.fromkeys(w.local for w in wraps))}）の中にあります。"
                            "変更履歴の中には書きません（Word で変更履歴を承諾するか元に戻してから、もう一度書いてください）")
        if c.lock in ("contentLocked", "sdtContentLocked") and not self.force:
            raise FillError(f"{item['id']}: 中身を変えられないように固定された欄です")
        if c.source == "formfield":
            ff = c.node.find("ffData")
            if c.kind == "checkbox":
                on = _as_bool(value)
                cb = ff.find("checkBox") if ff is not None else None
                if cb is None:
                    raise FillError(f"{item['id']}: チェックの欄の定義がありません")
                chk = cb.find("checked")
                v = "1" if on else "0"
                if chk is not None:
                    ed.put(chk.s, chk.e, f'<{pre}:checked {pre}:val="{v}"/>'.encode())
                else:
                    at = cb.es if cb.es != cb.e else None
                    if at is None:
                        ed.put(cb.s, cb.e, xml[cb.s:cb.e][:-2].rstrip() + b">" + f'<{pre}:checked {pre}:val="{v}"/></{pre}:checkBox>'.encode())
                    else:
                        ed.put(at, at, f'<{pre}:checked {pre}:val="{v}"/>'.encode())
                self.patches.append(Patch(item["id"], "control", on, part))
                return
            if c.kind == "dropdown":
                text = _is_text_value(value)
                if text not in c.options:
                    raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
                dd = ff.find("ddList")
                res = dd.find("result")
                i = c.options.index(text)
                if res is not None:
                    ed.put(res.s, res.e, f'<{pre}:result {pre}:val="{i}"/>'.encode())
                else:
                    ed.put(dd.se, dd.se, f'<{pre}:result {pre}:val="{i}"/>'.encode())
                self.patches.append(Patch(item["id"], "control", text, part))
                return
            text = _is_text_value(value)
            if c.kind == "date" and c.date_format:
                day = _parse_any_date(text)
                text = format_word_date(day, c.date_format, c.date_lid, c.date_calendar)
                if c.date_lid_from == "文書の設定 w:themeFontLang" and not _same_without_lang(day, c, text):
                    self.warnings.append(f"{item['id']}: 日付の欄の字に言語の指定がないので、文書の設定（w:themeFontLang）の言語"
                                         f" {c.date_lid} で「{text}」と書きました。Word で開く PC の言語で見え方が変わることがあります")
            if c.maxlen and len(text) > c.maxlen:
                raise FillError(f"{item['id']}: 「{text}」は {len(text)} 字です（この欄は {c.maxlen} 字まで）")
            self._replace_result(item, c, text)
            self.patches.append(Patch(item["id"], "input", text, part))
            return
        # コンテンツコントロール
        pr = c.node.find("sdtPr")
        if c.kind == "checkbox":
            on = _as_bool(value)
            box = next((k for k in pr.kids if k.local == "checkbox"), None)
            chk = box.find("checked") if box is not None else None
            if chk is not None:
                tag = xml[chk.s:chk.e]
                tag = re.sub(rb'(:val=")[^"]*(")', rb"\g<1>" + (b"1" if on else b"0") + rb"\g<2>", tag)
                ed.put(chk.s, chk.e, tag)
            state = box.find("checkedState" if on else "uncheckedState") if box is not None else None
            ch = chr(int(state.get("val", "2612" if on else "2610"), 16)) if state is not None else ("☒" if on else "☐")
            segs = [s for s in c.result_segs if s.kind == "t" and s.node is not None]
            if segs:
                n0 = segs[0].node
                ed.put(n0.se, n0.es, escape(ch))
            self.bind(item, c, value)
            self.patches.append(Patch(item["id"], "control", on, part))
            return
        text = _is_text_value(value)
        if any(ch_ in text for ch_ in LINE_BREAKS) and (c.kind in ("date", "dropdown", "combo")
                                                        or (c.kind == "text" and not c.multiline)):
            raise FillError(f"{item['id']}: この欄は1行の欄なので（"
                            + ("w:text の w:multiLine が「はい」でない" if c.kind == "text" else "選択肢・日付の欄")
                            + "）、改行を書けません。改行を除くか、1行にして書いてください")
        if c.kind == "dropdown" and text not in c.options:
            raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
        if c.kind == "date":
            d = _parse_any_date(text)
            text = format_word_date(d, c.date_format or "yyyy/MM/dd", c.date_lid, c.date_calendar)
            dnode = next((k for k in pr.kids if k.local == "date"), None)
            if dnode is not None:
                tag = xml[dnode.s:dnode.se]
                full = d.strftime("%Y-%m-%dT00:00:00Z").encode()
                if re.search(rb"fullDate=\"[^\"]*\"", tag):
                    tag = re.sub(rb'(fullDate=")[^"]*(")', rb"\g<1>" + full + rb"\g<2>", tag)
                else:
                    tag = tag[:-1].rstrip(b"/") + f' {pre}:fullDate="'.encode() + full + b'"' + (b"/>" if dnode.se == dnode.e else b">")
                ed.put(dnode.s, dnode.se, tag)
        self.bind(item, c, value)
        ph = pr.find("showingPlcHdr")
        if ph is not None:
            ed.put(ph.s, ph.e, b"")
        self._replace_result(item, c, text)
        self.patches.append(Patch(item["id"], "input", text, part))

    def bind(self, item: dict, c, value) -> None:
        """結び付けた XML の値（w:dataBinding）も書く予定を立てる（Word は開くとき、その値で欄を表示し直すため）。"""
        got = bound_store(self.doc, c, value, item["id"])
        if got is None:
            return
        t, stored = got
        key = (t.part, t.s, t.e)
        prev = self.bound.get(key)
        if prev is not None and prev[1] != stored:
            raise FillError(f"{item['id']}: {prev[0]} と同じ XML のデータに結び付いた欄に、違う値を書こうとしています"
                            f"（Word では、結び付いた欄はいつも同じ値になります）")
        self.bound[key] = (item["id"], stored, t)

    def finish_bound(self) -> dict[str, bytes]:
        """結び付けた XML の部品の、書いた値の所だけを変えたバイト列。同じデータに結び付いたほかの欄は、警告で伝える。"""
        from .databind import edit, locate

        out: dict[str, bytes] = {}
        by_part: dict[str, list] = {}
        for (part, _s, _e), (fid, stored, t) in self.bound.items():
            by_part.setdefault(part, []).append((t, stored))
        for part, todo in by_part.items():
            xml = self.doc.zip.read(part)
            for t, stored in sorted(todo, key=lambda x: -x[0].s):
                xml = edit(xml, t, stored)
            out[part] = xml
        if self.bound:
            written = {fid for fid, _v, _t in self.bound.values()}
            for name, it in self.sp.items.items():
                c = it.get("control")
                if c is None or name in written or c.source != "sdt":
                    continue
                try:
                    t = locate(self.doc, c)
                except Exception:  # noqa: BLE001
                    continue
                if t is not None and (t.part, t.s, t.e) in self.bound:
                    other = self.bound[(t.part, t.s, t.e)][0]
                    self.warnings.append(f"{name}: {other} と同じ XML のデータに結び付いた欄なので、Word で開くと {other} と同じ値になります")
        return out

    def _replace_result(self, item: dict, c, text: str) -> None:
        ed = self.ed(c.part)
        pre = ed.pre
        xml = ed.xml
        ts = [s for s in c.result_segs if s.kind == "t" and s.node is not None]
        if ts:
            first = ts[0]
            n0 = first.node
            # 改行は改行の要素（w:br）に（w:t の中の改行の字は、Word では空白に見える）
            ed.put(n0.s, n0.e, _t(pre, text) if text else f'<{pre}:t xml:space="preserve"></{pre}:t>'.encode())
            for s in ts[1:]:
                ed.put(s.node.s, s.node.e, f'<{pre}:t xml:space="preserve"></{pre}:t>'.encode())
            # 「ここをクリックして入力」の書式（プレースホルダーの文字のスタイル）は外す（Word が値を入れるときと同じ）
            for s in ts:
                st = placeholder_style(self.doc, s.run) if c.source == "sdt" else None
                if st is not None:
                    ed.put(st.s, st.e, b"")
            return
        # 結果の文字がない：結果の終わりの前に、文字を足す
        if c.source == "formfield" and c.end is not None:
            end_run = c.end.parent
            ed.put(end_run.s, end_run.s, f"<{pre}:r>".encode() + _t(pre, text) + f"</{pre}:r>".encode())
            return
        content = c.node.find("sdtContent")
        if content is not None:
            inner_p = next((k for k in content.kids if k.local == "p"), None)
            if inner_p is not None:
                at = inner_p.find("pPr").e if inner_p.find("pPr") is not None else inner_p.se
                ed.put(at, at, f"<{pre}:r>".encode() + _t(pre, text) + f"</{pre}:r>".encode())
            else:
                ed.put(content.se, content.se, f"<{pre}:r>".encode() + _t(pre, text) + f"</{pre}:r>".encode())
            return
        raise FillError(f"{item['id']}: 値を書く場所がありません")


def revision_warnings(f: "Filler") -> None:
    """変更履歴：承諾も却下もされていない変更履歴があること・変更履歴の記録がオンでも、書いた字は変更履歴にしていないことを伝える。"""
    doc = f.doc
    n = revision_counts(doc)
    if any(n.values()):
        what = "・".join(f"{k} {v}" for k, v in n.items() if v)
        f.warnings.insert(0, f"この文書には、まだ承諾も却下もされていない変更履歴（{what}）があります。Word で開くと変更履歴として"
                             "表示され、承諾・却下で字が変わります。書いた字は、変更履歴の外に、ふつうの字として書きました。"
                             "提出する前に、変更履歴をどうするか本人に確かめてください")
    for who, author in getattr(f, "revision_split", []):
        f.warnings.append(f"{who}: 書く所が、ほかの人（{author or '名前なし'}）の変更履歴の挿入の中だったので、その挿入を書く所で分け、"
                          "書いた字を挿入の外に置きました（挿入を却下すると、挿入の字は消え、書いた字は残ります）")
    if getattr(doc, "track_revisions", False):
        f.warnings.append("この文書は「変更履歴の記録」がオンになっていますが、とどくの書き込みは変更履歴として記録していません"
                          "（ふつうの字として書きました。Word の変更履歴には、書いた所は出ません）")


def _apply_edits(xml: bytes, s0: int, e0: int, edits: list) -> bytes:
    """xml[s0:e0] に、その中の書き換え（start, end, bytes）を当てたバイト列。"""
    out, at = bytearray(), s0
    for s, e, b in sorted(edits, key=lambda x: (x[0], x[1])):
        out += xml[at:s] + b
        at = e
    out += xml[at:e0]
    return bytes(out)


def _paras_under(node: Node, skip: tuple = ("Fallback",)) -> list[Node]:
    """node の中の段落（w:p）を文書の順に。skip の要素（入れ子の代わりの中身）の中は数えない。"""
    out = []
    stack = list(reversed(node.kids))
    while stack:
        n = stack.pop()
        if n.local in skip:
            continue
        if n.local == "p":
            out.append(n)
        stack.extend(reversed(n.kids))
    return out


def fallback_twins(root: Node, xml: bytes, edits: list) -> tuple[list, list]:
    """書き換え（start, end, bytes）のある mc:Choice の、mc:Fallback の扱い：(写す [(Fallback の段落, Choice の段落)],
    取り除く [Fallback])。写せるのは、Choice と Fallback の段落の数が同じで、書き換えのある段落（いちばん内側）の
    Fallback の同じ順の段落のバイト列が、書く前の Choice の段落と同じとき。証明（docx/proof.py）も同じ関数で決める。"""
    copy, drop = [], []
    for ac in root.iter("AlternateContent"):
        ch, fb = ac.find("Choice"), ac.find("Fallback")
        if ch is None or fb is None:
            continue
        hit = [e for e in edits if ch.s <= e[0] and e[1] <= ch.e]
        if not hit:
            continue
        cps, fps = _paras_under(ch), _paras_under(fb)
        pairs = []
        ok = len(cps) == len(fps)
        if ok:
            for k, cp in enumerate(cps):
                inner = [e for e in hit if cp.s <= e[0] and e[1] <= cp.e]
                if not inner or any(q is not cp and cp.s <= q.s and q.e <= cp.e and
                                    any(q.s <= e[0] and e[1] <= q.e for e in inner) for q in cps):
                    continue
                fp = fps[k]
                if xml[cp.s:cp.e] != xml[fp.s:fp.e]:
                    ok = False
                    break
                pairs.append((fp, cp))
            ok = ok and all(any(cp.s <= e[0] and e[1] <= cp.e for cp in cps) for e in hit)
        if ok and pairs:
            copy += pairs
        else:
            drop.append(fb)
    return copy, drop


def mirror_fallbacks(f: "Filler") -> None:
    """図形のテキストボックスは、Word が表示する mc:Choice（wps）と、古い Word・ほかのソフト用の mc:Fallback（VML）の2つに
    同じ中身を持つ。Choice のテキストボックスの段落に書いたら、Fallback の同じ順の段落にも同じ書き込みを写す（写した字には
    しおりを付けない。しおりの名前は文書で1つ）。写せないときは Fallback を取り除き、そのことを警告で伝える
    （書いていない字の古い中身が、ほかのソフトで見えないように）。"""
    for part, ed in list(f.editors.items()):
        if not ed.edits:
            continue
        root = f.doc.parts[part][1]
        xml = ed.xml
        edits = list(ed.edits)
        copy, drop = fallback_twins(root, xml, edits)
        for fp, cp in copy:
            new = _apply_edits(xml, cp.s, cp.e, [e for e in edits if cp.s <= e[0] and e[1] <= cp.e])
            new = re.sub(rb'<(?:[\w.-]+:)?bookmark(?:Start|End)\b[^>]*?\bid="(\d+)"[^>]*/>',
                         lambda m_: b"" if int(m_[1]) in f.bm_ids else m_[0], new)
            ed.put(fp.s, fp.e, new)
        for fb in drop:
            ed.put(fb.s, fb.e, b"")
        if drop:
            f.warnings.append("テキストボックスに書きました。このテキストボックスの、古い Word・ほかのソフト用の代わりの中身"
                              "（mc:Fallback の VML）は、作りが表示用の中身（mc:Choice）と違い、同じ書き込みを写せないので、"
                              "取り除きました（いまの Word は表示用の中身を使います。古い Word では、図形が表示されないことがあります）")


# 行の区切りの字（str.splitlines が区切る字）：1行の欄には書かない
LINE_BREAKS = ("\n", "\r", "\x0b", "\x0c", "\x1c", "\x1d", "\x1e", "\x85", "\u2028", "\u2029")


def bound_store(doc, c, value, fid: str = "") -> "tuple | None":
    """コンテンツコントロールが XML のデータに結び付いていれば (値の所, 書く値)。結び付いていなければ None。
    たどれない・書く値の形が決まらないときは、書かずに FillError（書いた値が、Word で開くと消えるため）。"""
    from .databind import BindError, binding_of, locate, stored_value

    b = binding_of(c)
    if b is None:
        return None
    try:
        t = locate(doc, c)
        disp = ""
        if c.kind == "checkbox":
            shown = _as_bool(value)
        elif c.kind == "date":
            shown = _parse_any_date(_is_text_value(value))
            disp = format_word_date(shown, c.date_format or "yyyy/MM/dd", c.date_lid, c.date_calendar)
        else:
            shown = _is_text_value(value)
        return t, stored_value(c, shown, b, t.value, disp)
    except BindError as e:
        raise FillError(f"{fid}: この欄は文書の中の XML のデータ（w:dataBinding「{b.xpath}」）に結び付いていて、Word は開くときに"
                        f"その値で表示し直します。その値を書けないので、書きませんでした（{e}）") from None


def placeholder_style(doc, run) -> "Node | None":
    """字の要素（w:r）の、プレースホルダーの文字のスタイル（w:rStyle）。Word は、コンテンツコントロールのプレースホルダーを
    「Placeholder Text」のスタイルで表示し、値を入れるとそのスタイルを外す。スタイルの名前（styles.xml の w:name）か、
    名前の分からないときは id で見分ける。書き込みと証明で同じものを使う。"""
    rpr = run.find("rPr") if run is not None else None
    st = rpr.find("rStyle") if rpr is not None else None
    if st is None:
        return None
    sid = st.get("val") or ""
    name = (getattr(doc, "style_names", {}) or {}).get(sid, sid)
    return st if name.replace(" ", "").lower() == "placeholdertext" or sid.lower() == "placeholdertext" else None


# ====================================================================== 日付
def _parse_any_date(text: str) -> dt.date:
    t = str(text).strip()
    try:
        from ..xlsx.edit import parse_date_text

        d = parse_date_text(t)
        if d is not None:
            return d.date() if isinstance(d, dt.datetime) else d
    except Exception:  # noqa: BLE001
        pass
    m = re.fullmatch(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?", t)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))
    raise FillError(f"「{t}」を日付として読めません")


def date_lang(lid: str | None) -> str | None:
    """w:lid（ECMA-376 第1部 17.18.47 ST_Lang：言語の名前（ja-JP）か、16進の LCID（0411））→ locale_fmt の言語の名前。
    読めない・知らない言語は、名前を出さないように「tag:…」「lcid:…」の形（name_table は Unsupported を投げる）。"""
    from ..formula.locale_fmt import _lang_of_lcid, _lang_of_tag
    from ..formula.values import Unsupported

    t = (lid or "").strip()
    if not t:
        return None
    try:
        if re.fullmatch(r"[0-9A-Fa-f]{4}", t):
            return _lang_of_lcid(int(t, 16))
        return _lang_of_tag(t)[0]
    except Unsupported:
        return f"tag:{t}"


# w:calendar（ECMA-376 第1部 17.18.7 ST_CalendarType）のうち、年月日の数がグレゴリオ暦と同じもの。
# gregorianXlit*・gregorianMeFrench・gregorianArabic は、月・曜日の名前の書き方（音訳など）が違うので、名前は出さない
_GREG_NUM = {"", "gregorian", "gregorianUs", "gregorianXlitEnglish", "gregorianXlitFrench", "gregorianMeFrench",
             "gregorianArabic", "none"}
_GREG_NAMES = {"", "gregorian", "gregorianUs", "none"}


def _same_without_lang(d: dt.date, c, text: str) -> bool:
    """言語を決めなくても同じ字になるか（数だけの書式）。言語で変わる書式（名前・元号）なら False。"""
    try:
        return format_word_date(d, c.date_format, None, c.date_calendar) == text
    except FillError:
        return False


def format_word_date(d: dt.date, fmt: str, lid: str | None = None, calendar: str | None = None) -> str:
    """Word の日付の書式（w:dateFormat：yyyy/MM/dd、ggge年M月d日、dddd d MMMM yyyy など）を、言語（w:lid）と
    暦（w:calendar）どおりに字にする。名前（月・曜日）は formula/locale_fmt の表、和暦は formula/dates の元号の表、
    ヒジュラ暦は locale_fmt の表で決まるヒジュラ暦（Windows の HijriCalendar）。タイの暦は西暦 + 543、韓国の暦（檀紀）は
    西暦 + 2333、台湾（民国）の年は西暦 − 1911。確かめていない組み合わせ（ヘブライ暦・サカ暦・民国の元号の名前など）は、
    推測した字を書かずに FillError を投げる。日付の欄は時刻を 0 時として書く（fullDate の T00:00:00Z と同じ）。"""
    from ..formula.dates import to_wareki
    from ..formula.locale_fmt import hijri_tabular, name_table
    from ..formula.values import Unsupported

    cal = calendar or ""
    lang = date_lang(lid)
    try:
        return _word_date(d, fmt, lang, cal, to_wareki, hijri_tabular, name_table)
    except Unsupported as e:
        where = f"言語 {lid or '指定なし'}・暦 {calendar or 'gregorian'}・書式「{fmt}」"
        raise FillError(f"この日付の欄の見える字（{where}）を、Word と同じに作れません（{e}）。"
                        "日付の欄でない所に字で書くか、書式を確かめてください") from None


def _word_date(d, fmt, lang, cal, to_wareki, hijri_tabular, name_table) -> str:
    from ..formula.values import Unsupported

    # 書式を、字（'…' で囲んだ所・記号でない字）と記号の並びに分ける
    toks: list[tuple[str, str]] = []
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == "'":
            j = fmt.find("'", i + 1)
            j = len(fmt) if j < 0 else j
            toks.append(("lit", fmt[i + 1:j] if j > i + 1 else ("'" if fmt[i + 1:i + 2] == "'" else "")))
            i = j + 1
            continue
        if fmt[i:i + 5].upper() == "AM/PM":
            toks.append(("code", "AM/PM"))
            i += 5
            continue
        if ch in "yMdgeaHhms":
            j = i
            while j < len(fmt) and fmt[j] == ch:
                j += 1
            toks.append(("code", fmt[i:j]))
            i = j
            continue
        if ch.isascii() and ch.isalpha():
            raise Unsupported(f"日付の書式の記号「{ch}」")
        toks.append(("lit", ch))
        i += 1
    has_day = any(k == "code" and t[0] == "d" and len(t) <= 2 for k, t in toks)

    def ymd(part: int) -> int:
        """暦の年（0）・月（1）・日（2）。"""
        if cal in _GREG_NUM:
            return (d.year, d.month, d.day)[part]
        if cal == "hijri":
            return hijri_tabular(d.year, d.month, d.day)[part]
        if cal in ("thai", "korea", "japan", "taiwan"):
            if part:
                return (d.year, d.month, d.day)[part]  # 月と日はグレゴリオ暦と同じ
            if cal == "thai":
                return d.year + 543
            if cal == "korea":
                return d.year + 2333
            raise Unsupported(f"暦 {cal} の年の記号 y（元号の年は e で書きます）")
        raise Unsupported(f"暦 {cal}")

    def names(kind: str) -> list[str]:
        # 和暦・民国・タイ・韓国の暦は、月と曜日がグレゴリオ暦と同じ（年の数え方だけが違う）
        if cal not in _GREG_NAMES | {"japan", "taiwan", "thai", "korea"}:
            raise Unsupported(f"暦 {cal} の月・曜日の名前")
        lg = "en" if cal == "gregorianUs" else lang  # 米国の英語のグレゴリオ暦：名前は英語
        if lg is None:
            raise Unsupported("言語（w:lid）の指定のない、月・曜日の名前")
        return name_table(lg, kind, has_day)

    def era() -> tuple[str, int, str]:
        if cal == "japan" or (cal == "" and lang == "ja"):
            name, _abbr, n = to_wareki(d)
            if not name:
                raise Unsupported("明治より前の日付の元号")
            return name, n, "jp"
        if cal == "taiwan":
            if d.year <= 1911:
                raise Unsupported("中華民国の暦より前の日付")
            return "", d.year - 1911, "roc"
        raise Unsupported(f"暦 {cal or 'gregorian'}・言語 {lang or '指定なし'} の元号の記号 g・e")

    sun0 = (d.weekday() + 1) % 7
    out = []
    for idx, (k, t) in enumerate(toks):
        if k == "lit":
            out.append(t)
            continue
        c, n = t[0], len(t)
        if c == "y":
            y = ymd(0)
            if n == 2:
                out.append(f"{y % 100:02d}")
            elif n == 4:
                out.append(f"{y:04d}")
            else:
                raise Unsupported(f"年の記号 {t}")
        elif c == "M":
            if n <= 2:
                m = ymd(1)
                out.append(f"{m:02d}" if n == 2 else str(m))
            elif n in (3, 4):
                out.append(names("mmm" if n == 3 else "mmmm")[d.month - 1])
            else:
                raise Unsupported(f"月の記号 {t}")
        elif c == "d":
            if n <= 2:
                dd = ymd(2)
                out.append(f"{dd:02d}" if n == 2 else str(dd))
            elif n in (3, 4):
                out.append(names("ddd" if n == 3 else "dddd")[sun0])
            else:
                raise Unsupported(f"曜日の記号 {t}")
        elif c == "a":
            if n not in (3, 4):
                raise Unsupported(f"曜日の記号 {t}")
            out.append(names("aaa" if n == 3 else "aaaa")[sun0])
        elif c == "g":
            name, _n, kind = era()
            if kind != "jp" or n > 3:
                raise Unsupported(f"元号の名前の記号 {t}（暦 {cal}）")
            out.append(name if n == 3 else name[:1])
        elif c == "e":
            _name, ey, kind = era()
            nxt = next((x for kk, x in toks[idx + 1:] if kk == "lit"), "")
            if n == 1 and ey == 1 and kind == "jp" and nxt.startswith("年"):
                out.append("元")  # 元号の最初の年は「元年」（日本語の Word の和暦の表示）
            elif n <= 2:
                out.append(f"{ey:02d}" if n == 2 else str(ey))
            else:
                raise Unsupported(f"元号の年の記号 {t}")
        elif c in "Hh":
            h = 0 if c == "H" else 12  # 日付の欄の時刻は 0 時（12 時間の表示では 12）
            out.append(f"{h:02d}" if n == 2 else str(h))
        elif c in "ms":
            out.append("00" if n == 2 else "0")
        elif t == "AM/PM":
            if lang != "en":
                raise Unsupported("英語でない言語の午前・午後の記号 AM/PM")
            out.append("AM")
        else:
            raise Unsupported(f"日付の書式の記号 {t}")
    return "".join(out)


# ====================================================================== 入口
def fill(doc: DocxDocument, values: dict, *, force: bool = False) -> FillResult:
    from ..embed import EmbedError, is_embed_key
    from .window import WriteError, is_bullet, resolve_write

    f = Filler(doc, force)
    problems = []
    used: dict = {}
    embed_parts: dict[str, bytes] = {}
    embed_values = {k: v for k, v in values.items() if is_embed_key(k)}
    if embed_values:
        # 埋め込んだ Excel の表（E1!Sheet1!C4）：ブックの部品と、Word が表示・印刷に使う表の絵の部品を書き直す
        if doc.protection_enforced and doc.protection and not force:
            problems.append(f"{next(iter(embed_values))}: この文書は保護されていて（{doc.protection}）、埋め込みの表には書けません")
        else:
            from ..embed import write as write_embeds

            try:
                ew = write_embeds(doc, embed_values)
                embed_parts = ew.parts
                f.patches += [Patch(k, "input", v, "embed") for k, v in ew.written]
                f.warnings += ew.warnings
            except EmbedError as e:
                problems.append(str(e))
    for key, value in values.items():
        if key in embed_values:
            continue
        item = f.sp.items.get(key)
        try:
            if key.startswith("@"):
                # 規則で示していない所：エージェントが字を目印にして指した所（{"old": 元の字, "new": 書き換えた後の字}）
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    item = resolve_write(f.sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                p = item["para"]
                a, b = item["span"]
                f.replace_span(item, p, a, b, item["value"])
                f.patches.append(Patch(item["id"], "input", item["value"], p.part, anchor=dict(value), key=key))
                continue
            if item is None:
                raise FillError(f"{key}: そのような欄はありません{f.sp.items.hint(key)}")
            if item.get("control") is not None and item["control"].source == "activex":
                # ActiveX の部品：部品の中身（activeX*.bin・性質の袋）と控えの絵を書き換える（axfile）
                from ..activex import AxError
                from ..axfile import word_fill

                f.check_editable(item, item["control"].part, item["control"].node.s)
                try:
                    for fid, shown in word_fill(f, item, value, force):
                        f.patches.append(Patch(fid, "control", shown, item["control"].part))
                except (AxError, ValueError) as e:
                    raise FillError(f"{key}: {e}") from None
                continue
            if item.get("control") is not None:
                f.fill_control(item, value)
                continue
            if key in f.sp.marks:
                raise FillError(f"{key}: すでに書き込んだ欄です（書き直すには、元の様式に書いてください）")
            kind = item["kind"]
            if kind == "cell":
                text = _is_text_value(value)
                p = item["para"]
                if p.text and not p.text.strip(" \u3000\u2002\u2003"):
                    f.replace_span(item, p, 0, len(p.text), text)
                else:
                    f.fill_cell(item, text)
                f.patches.append(Patch(key, "input", text, item["para"].part))
            elif kind == "checkbox" and is_bullet(item):
                on = _as_bool(value)
                f.fill_bullet(item, on)
                f.patches.append(Patch(key, "control", on, item["para"].part))
            elif kind == "checkbox":
                on = _as_bool(value)
                p = item["para"]
                a, b = item["span"]
                orig = p.text[a:b]
                mark = box_mark(orig, on)
                f.replace_span(item, p, a, b, mark)
                f.patches.append(Patch(key, "control", on, p.part))
            elif kind == "inline":
                from ..textwidth import align_options

                try:
                    opts, value = align_options(key, value)
                except ValueError as e:
                    raise FillError(str(e)) from None
                text = _is_text_value(value)
                p = item["para"]
                a, b = item["span"]
                # 下線を引いた空欄（空欄の字がすべて下線つき）は、幅を保つ（下線の長さを変えない）。そろえ方の指定があれば、
                # 空きの幅を保って、そのとおりに前後へ空白を足す（inline_padding。証明も同じ決まりで確かめる）
                pad = inline_padding(p, a, b, text, opts)
                f.replace_span(item, p, a, b, text, pad=pad)
                f.patches.append(Patch(key, "input", text, p.part,
                                       align={**opts, "pad": [len(pad[0]), len(pad[1])]} if opts else None))
                lead, trail = item.get("kept_edges") or (False, False)
                if text and (lead or trail):
                    where = "・".join(x for x, on in (("前の見出し", lead), ("後ろの字", trail)) if on)
                    f.warnings.append(f"{key}: 書いた字が{where}にくっつかないよう、空白の端の 1 字を残して、空白の中に書きました")
            else:
                raise FillError(f"{key}: この種類の欄（{kind}）には書けません")
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    f.finish_bullets()
    f.finish_runs()
    bound_parts = f.finish_bound()
    mirror_fallbacks(f)
    if any(ed.edits for ed in f.editors.values()) or embed_parts or bound_parts:
        revision_warnings(f)
    z = doc.zip
    from ..zipraw import RawZip

    z2 = RawZip(doc.blob)
    for part, ed in f.editors.items():
        if ed.edits:
            z2.replace(part, ed.build())
    for part, data in embed_parts.items():
        z2.replace(part, data)
    for part, data in bound_parts.items():
        z2.replace(part, data)  # 結び付けた XML の部品（customXml/itemN.xml）：書いた値の所だけ
    for part, data in getattr(f, "ax_parts", {}).items():
        z2.replace(part, data)  # ActiveX の部品の中身と控えの絵
    for part, (data, like) in getattr(f, "ax_added", {}).items():
        z2.add(part, data, like)  # 共有していた控えの絵から分けた、その部品だけの絵
    blob = z2.build() if f.editors or embed_parts or bound_parts or getattr(f, "ax_parts", None) else doc.blob
    del z
    return FillResult(blob, f.patches, f.warnings)
