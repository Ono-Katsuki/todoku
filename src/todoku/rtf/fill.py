"""RTF に書き込む：書いた所の字だけを、元のバイト列への小さな差分として変える（ほかのバイトはそのまま）。

- 表の空のマス・空の段落：段落の終わりの印（\\par・\\cell など）の前に字を足す（その所の書式のまま）
- 段落の中の空欄：空欄の字（\\'hh・\\uN と代わりの字・文字）を値に置き換える。字のあいだの制御語（書式）は残す
- □：■（☐ は ☑）にする
- 旧来のフォームの欄：FORMTEXT は結果（\\fldrslt）の字を、FORMCHECKBOX・FORMDROPDOWN は \\ffres の値を変える
書く字は、ASCII はそのまま（\\・{・} はエスケープ）、ほかは \\uN と、その所の \\uc の数だけの代わりの字（文書の文字コードで
書ける字は \\'hh、書けなければ ?）。区切りのない制御語のすぐ後ろに書くときは、区切りの空白を足す（字が制御語とつながらないように）。
書いた所は見えないしおり（{\\*\\bkmkstart _todoku_…}）で囲み、書いた後も同じ id で読める（フォームの欄は、欄の名前のまま）。
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field

from ..docx.document import form_field_notes, form_field_refusal, iter_paras
from ..docx.fill import SYMBOL_FONT, FillError, _as_bool, _is_text_value, inline_padding, tab_note, written_of  # 証明に渡す、書いた所 → 値（Word と同じ）
from ..docx.proof import _expected_text
from ..docx.window import Spots, WriteError, box_mark, resolve_write
from .document import MARK_PREFIX, RtfDocument, color_table
from .lexer import control_word_ends_at

B = b"\x5c"  # 逆斜線


@dataclass
class Patch:
    id: str
    kind: str
    value: object
    part: str = "rtf"
    anchor: dict | None = None
    key: str = ""


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def rtf_text(value: str, uc: int, cp: str) -> bytes:
    """書く字の RTF：ASCII はそのまま、ほかは Unicode（\\uN）と、uc の数だけの代わりの字。"""
    out = bytearray()
    for ch in value:
        o = ord(ch)
        if ch == "\\":
            out += B + B
        elif ch == "{":
            out += B + b"{"
        elif ch == "}":
            out += B + b"}"
        elif ch == "\t":
            out += B + b"tab "
        elif ch in "\r\n":
            if ch == "\n":
                out += B + b"line "
        elif 0x20 <= o < 0x7F:
            out += bytes([o])
        else:
            units = [o] if o < 0x10000 else [0xD800 + ((o - 0x10000) >> 10), 0xDC00 + ((o - 0x10000) & 0x3FF)]
            fb = b""
            if uc > 0:
                try:
                    enc = ch.encode(cp) if not cp.startswith("symbol") else b""
                except (UnicodeEncodeError, LookupError):
                    enc = b""
                if len(units) == 1 and len(enc) == uc:
                    fb = b"".join(B + b"'" + f"{x:02x}".encode() for x in enc)
                else:
                    fb = b"?" * uc
            for k, unit in enumerate(units):
                n = unit if unit < 32768 else unit - 65536
                out += B + b"u" + str(n).encode() + (fb if (uc > 0 and len(units) == 1) else (b"?" * uc if uc > 0 else b" "))
    return bytes(out)


# ====================================================================== 書いた字の書式
LOOK_NAMES = {"strike": "取り消し線", "striked": "二重取り消し線", "v": "非表示", "cf": "字の色", "highlight": "蛍光ペン",
              "cb": "背景の色", "chcbpat": "背景の色（網かけ）", "revised": "変更履歴の挿入の印", "deleted": "変更履歴の削除の印",
              "f": SYMBOL_FONT}


def bad_looks(look: dict | None, colors: list) -> list[str]:
    """消したように・目立って見える書式の制御語（LOOK_NAMES）。"""
    look = look or {}
    out = []
    if look.get("f") is not None:
        out.append("f")  # 記号の書体：引き継ぐと、書いた字が、その書体の別の字（絵・ギリシア文字）になる
    for w in ("strike", "striked", "v", "revised", "deleted"):
        if look.get(w, 0):
            out.append(w)

    def color(i):
        return colors[i] if isinstance(i, int) and 0 <= i < len(colors) else None

    c = color(look.get("cf", 0))
    if c is not None and c != (0, 0, 0):
        out.append("cf")
    for w in ("highlight", "cb", "chcbpat"):
        i = look.get(w, 0)
        c = color(i)
        if i and (c is not None and c != (255, 255, 255)):
            out.append(w)
    return out


class RtfWriter:
    def __init__(self, doc: RtfDocument, force: bool = False):
        self.doc = doc
        self.blob = doc.blob
        self.colors = color_table(doc.blob)
        self.force = force
        self.sp = Spots(doc)
        self.edits: list[tuple[int, int, bytes]] = []
        self.patches: list[Patch] = []
        self.warnings: list[str] = []
        self.marks_used = set(doc.marks())

    def put(self, s: int, e: int, b: bytes) -> None:
        for a, z, _x in self.edits:
            if (s < z and a < e) or (s == e == a == z) or (s == e and a < s < z):
                raise FillError("同じ所に2つの書き込みが重なりました")
        # 区切りのない制御語のすぐ後ろ：区切りの空白を足す（足す字・残る字が、制御語とつながらないように）
        nxt = self.blob[e:e + 1]
        if control_word_ends_at(self.blob, s) and (b or (nxt.isalnum() or nxt in (b" ", b"-"))):
            b = b" " + b
        self.edits.append((s, e, b))

    def bookmark(self, fid: str) -> tuple[bytes, bytes]:
        name = (MARK_PREFIX + fid.replace(".", "_").replace("#", "_")).encode("ascii", "replace")
        return (b"{" + B + b"*" + B + b"bkmkstart " + name + b"}", b"{" + B + b"*" + B + b"bkmkend " + name + b"}")

    def check(self, item: dict) -> None:
        if self.force:
            return
        d = self.doc
        if d.protection_enforced and (d.protection != "forms" or item.get("control") is None):
            raise FillError(f"{item['id']}: この文書は保護されていて（{d.protection}）、ここには書けません")

    def enc(self, value: str, uc: int) -> bytes:
        return rtf_text(value, uc, self.doc.ansicpg)

    def clean(self, look: dict | None, body: bytes, fid: str) -> bytes:
        """書く所の字の書式に、取り消し線・非表示・字の色・蛍光ペン・背景の色・記号の書体があれば、それを打ち消す制御語の
        グループで値を囲む（字の大きさは、その所のまま。書体は、記号の書体のときだけ、その前に選ばれていた書体に戻す）。
        打ち消した書式は警告で伝える。戻す書体のない記号の書体は、そのまま（読み戻しで、書いた字と違えば断る）。"""
        bad = bad_looks(look, self.colors)
        font = None
        if "f" in bad:
            font = look.get("nf") if look.get("nf") is not None else self.doc.deff
            if font is None or font not in self.doc.fonts or self.doc.symbol_font(font) is not None:
                bad.remove("f")
        if not bad or not body:
            return body
        words = b""
        for w in bad:
            if w == "f":
                words += B + b"f" + str(font).encode()
            elif w in ("strike", "striked", "v", "highlight", "revised", "deleted"):
                words += B + w.encode() + b"0"
            else:
                want = (0, 0, 0) if w == "cf" else (255, 255, 255)
                idx = 0 if not self.colors or self.colors[0] is None else (
                    self.colors.index(want) if want in self.colors else None)
                if idx is None:
                    raise FillError(f"{fid}: 書く所の字の{LOOK_NAMES[w]}を打ち消す色が、色の表（\\colortbl）にありません")
                words += B + w.encode() + str(idx).encode()
        names = list(dict.fromkeys(f"{SYMBOL_FONT}（{self.doc.symbol_font(look['f'])}）" if w == "f" else LOOK_NAMES[w] for w in bad))
        kept = "字の大きさ・書体はそのまま" if font is None else f"字の大きさはそのまま。書体は、記号の書体の前の「{self.doc.fonts[font][1]}」"
        self.warnings.append(f"{fid}: 書く所に残っていた書式（{'・'.join(names)}）は、書いた字に引き継がず、打ち消す書式で"
                             f"書きました（{kept}）")
        return b"{" + words + b" " + body + b"}"

    # ------------------------------------------------------------------ 段落の字
    def replace_span(self, item: dict, p, a: int, b: int, value: str, keep_width: bool = False, before_tab: bool = False) -> None:
        self.check(item)
        pos = p.positions()
        bm = self.bookmark(item["id"])
        if a == b and before_tab:
            # タブで始まる空欄：タブの前に、タブの所の書式（下線など）で差し込む。タブは残す
            src = p.segs[pos[a][0]].node if a < len(pos) else None
            if src is None or src.kind != "tab":
                raise FillError(f"{item['id']}: 書き足す位置が、書き換えられない部分にあります")
            self.put(src.s, src.s, bm[0] + self.clean(src.look, self.enc(value, 1), item["id"]) + bm[1])
            return
        if a == b:
            if not pos:
                self.para_end(item, p, value)
                return
            i, off = pos[a - 1] if a > 0 else pos[0]
            src = p.segs[i].node
            if src is None:
                raise FillError(f"{item['id']}: 書き足す位置が、書き換えられない部分にあります")
            if a > 0 and p.segs[i].kind == "note":
                # 脚注の印のすぐ後ろ：印は、上付きの字のグループの中にある（{\super \chftn {\footnote …}}）。その中に書くと、
                # 書いた字が上付きになる。段落の終わりなら段落の印の前に、そうでなければ、次の字の前に（次の字の書式で）書く
                if a == len(pos):
                    self.para_end(item, p, value)
                    return
                nxt = p.segs[pos[a][0]].node
                if nxt is not None:
                    at = nxt.starts[pos[a][1]] if nxt.kind == "txt" else nxt.s
                    self.put(at, at, bm[0] + self.clean(nxt.look, self.enc(value, getattr(nxt, "uc", 1)), item["id"]) + bm[1])
                    return
            if src.kind == "txt":
                at = src.ends[off] if a > 0 else src.starts[off]
            else:
                at = src.e if a > 0 else src.s
            self.put(at, at, bm[0] + self.clean(src.look, self.enc(value, getattr(src, "uc", 1)), item["id"]) + bm[1])
            return
        if b > len(pos):
            raise FillError(f"{item['id']}: 空欄の位置が見つかりません")
        order: list[int] = []
        ranges: dict[int, list[int]] = {}
        for k in range(a, b):
            i, off = pos[k]
            if i not in ranges:
                order.append(i)
                ranges[i] = []
            ranges[i].append(off)
        pad = inline_padding(p, a, b, value)[1] if keep_width else ""  # 下線の長さを保つ（Word と同じ決まり）
        first = True
        for i in order:
            sg = p.segs[i]
            src = sg.node
            if src is not None and src.kind == "tab":
                # 空欄の中のタブ：取り除く（空欄の初めなら、そこに値を書く）
                self.put(src.s, src.e, (bm[0] + self.clean(src.look, self.enc(value, 1), item["id"]) + bm[1] + self.enc(pad, 1))
                         if first else b"")
                first = False
                continue
            if src is None or src.kind != "txt" or (sg.field is not None and sg.field.kind != "field"):
                raise FillError(f"{item['id']}: 空欄が、書き換えられない部分（{src.kind if src else '?'}）にかかっています")
            if sg.field is not None:
                raise FillError(f"{item['id']}: 空欄が、フィールド（日付・ページ番号など）の結果にかかっています")
            k0, k1 = min(ranges[i]), max(ranges[i]) + 1
            new = (bm[0] + self.clean(src.look, self.enc(value, src.uc), item["id"]) + bm[1] + self.enc(pad, src.uc)) if first else b""
            self.put(src.starts[k0], src.ends[k1 - 1], new)
            first = False

    def para_end(self, item: dict, p, value: str) -> None:
        """段落の終わり（\\par・\\cell など）の前に字を足す。"""
        self.check(item)
        bm = self.bookmark(item["id"])
        at = p.mark[0] if p.mark else p.end
        self.put(at, at, bm[0] + self.clean(getattr(p, "mark_look", None), self.enc(value, getattr(p, "uc", 1)), item["id"]) + bm[1])

    # ------------------------------------------------------------------ フォームの欄
    def control(self, item: dict, value) -> None:
        self.check(item)
        c = item["control"]
        if getattr(c, "unwritable", "") and not self.force:
            raise FillError(f"{item['id']}: {c.unwritable}")
        form_field_refusal(c, item["id"], self.force)
        ff = getattr(c, "ff", {})
        if c.kind in ("checkbox", "dropdown"):
            self.warnings += form_field_notes(c, item["id"], None, self.doc.controls)
        if c.kind == "checkbox":
            on = _as_bool(value)
            self._ffres(c, ff, 1 if on else 0)
            self.patches.append(Patch(item["id"], "control", on))
            return
        if c.kind == "dropdown":
            text = _is_text_value(value)
            if text not in c.options:
                raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
            self._ffres(c, ff, c.options.index(text))
            segs = [s for s in c.result_segs if s.node is not None and s.node.kind == "txt"]
            if segs:
                self._result(c, text)
            self.patches.append(Patch(item["id"], "control", text))
            return
        text = _is_text_value(value)
        if c.maxlen and len(text) > c.maxlen:
            raise FillError(f"{item['id']}: 「{text}」は {len(text)} 字です（この欄は {c.maxlen} 字まで）")
        self.warnings += form_field_notes(c, item["id"], text, self.doc.controls)
        self._result(c, text)
        self.patches.append(Patch(item["id"], "input", text))

    def _result(self, c, text: str) -> None:
        segs = [s for s in c.result_segs if s.node is not None and s.node.kind == "txt"]
        if segs:
            first = True
            for s in segs:
                self.put(s.node.starts[0], s.node.ends[-1], self.enc(text, s.node.uc) if first else b"")
                first = False
            return
        at = getattr(c, "result_close", None)
        if at is None:
            raise FillError(f"{c.name}: 値を書く場所（フィールドの結果）がありません")
        self.put(at, at, self.enc(text, 1))

    def _ffres(self, c, ff: dict, value: int) -> None:
        if "ffres_pos" in ff:
            s, e = ff["ffres_pos"]
            old = self.blob[s:e]
            tail = b" " if old.endswith(b" ") else b""
            self.put(s, e, B + b"ffres" + str(value).encode() + tail)
        elif "fftype_pos" in ff:
            s, e = ff["fftype_pos"]
            self.put(e, e, B + b"ffres" + str(value).encode())
        else:
            raise FillError(f"{c.name or 'フォームの欄'}: 欄の状態（\\ffres）を書く場所がありません")

    def build(self) -> bytes:
        out = bytearray()
        at = 0
        for s, e, b in sorted(self.edits, key=lambda x: (x[0], x[1])):
            out += self.blob[at:s]
            out += b
            at = e
        out += self.blob[at:]
        return bytes(out)


def fill(doc: RtfDocument, values: dict, *, force: bool = False) -> FillResult:
    w = RtfWriter(doc, force)
    sp = w.sp
    problems = []
    used: dict = {}
    for key, value in values.items():
        item = sp.items.get(key)
        try:
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    item = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                p = item["para"]
                a, b = item["span"]
                w.replace_span(item, p, a, b, item["value"])
                w.patches.append(Patch(item["id"], "input", item["value"], anchor=dict(value), key=key))
                continue
            if item is None:
                raise FillError(f"{key}: そのような欄はありません{sp.items.hint(key)}")
            if item.get("control") is not None:
                w.control(item, value)
                continue
            if key in w.marks_used:
                raise FillError(f"{key}: すでに書き込んだ欄です（書き直すには、元の様式に書いてください）")
            kind = item["kind"]
            p = item["para"]
            if kind == "cell":
                text = _is_text_value(value)
                if p.text and not p.text.strip(" 　  "):
                    w.replace_span(item, p, 0, len(p.text), text)
                else:
                    w.para_end(item, p, text)
                w.patches.append(Patch(key, "input", text))
            elif kind == "checkbox":
                on = _as_bool(value)
                a, b = item["span"]
                orig = p.text[a:b]
                if box_mark(orig, on) != orig:
                    w.replace_span(item, p, a, b, box_mark(orig, on))
                w.patches.append(Patch(key, "control", on))
            elif kind == "inline":
                text = _is_text_value(value)
                a, b = item["span"]
                tabs = item.get("tabs")  # 所の後ろに残す、空欄のタブ
                pos = p.positions()
                blank = [p.segs[pos[i][0]] for i in range(a, b) if i < len(pos)]
                underlined = bool(blank) and all(sg.underline for sg in blank)
                w.replace_span(item, p, a, b, text, keep_width=underlined and bool(text), before_tab=bool(tabs) and a == b)
                if tabs and text:
                    w.warnings.append(tab_note(key, tabs))
                w.patches.append(Patch(key, "input", text))
            else:
                raise FillError(f"{key}: この種類の欄（{kind}）には書けません")
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    revision_warnings(doc, w.warnings)
    return FillResult(w.build(), w.patches, w.warnings)


def revision_warnings(doc: RtfDocument, warnings: list[str]) -> None:
    """変更履歴：承諾も却下もされていない変更履歴があること・変更履歴の記録がオンでも、書いた字は変更履歴にしていないことを伝える
    （Word・OpenDocument と同じ）。書いた字は、挿入・削除の印を打ち消して書くので（clean）、変更履歴の外のふつうの字になる。"""
    n = {k: v for k, v in doc.revisions.items() if v}
    if n:
        what = "・".join(f"{k} {v}" for k, v in n.items())
        warnings.insert(0, f"この文書には、まだ承諾も却下もされていない変更履歴（{what}）があります。ワープロで開くと変更履歴として"
                           "表示され、承諾・却下で字が変わります。書いた字は、変更履歴の外に、ふつうの字として書きました。"
                           "提出する前に、変更履歴をどうするか本人に確かめてください")
    if doc.tracking:
        warnings.append("この文書は「変更履歴の記録」がオンになっていますが、とどくの書き込みは変更履歴として記録していません"
                        "（ふつうの字として書きました。ワープロの変更履歴には、書いた所は出ません）")


# ====================================================================== 証明
@dataclass
class RtfProof:
    original_sha256: str
    new_sha256: str
    paragraphs_total: int = 0
    paragraphs_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        return (f"段落 {self.paragraphs_total:,} のうち、変わったのは書いた欄のある {self.paragraphs_changed} 段落だけ。"
                "その段落の外のバイトは、書く前と1バイトも違いません"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）"))


def _window(p) -> tuple[int, int]:
    """書き込みが起こりうる、段落のバイトの範囲：最初の字（またはフォームの欄の始まり）から、段落の終わりの印の終わりまで。"""
    los = [p.mark[0]] if p.mark else [p.end]
    for sg in p.segs:
        if sg.node is not None:
            los.append(sg.node.s)
        c = sg.field
        if c is not None and getattr(c, "field_s", None) is not None:
            los.append(c.field_s)
    return min(los), (p.mark[1] if p.mark else p.end)


def prove(original: bytes, new: bytes, written: dict) -> RtfProof:
    pr = RtfProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    a = RtfDocument(original)
    try:
        b = RtfDocument(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません：{e}")
        return pr
    depth = 0
    from .lexer import tokens

    for t in tokens(new):
        depth += 1 if t.kind == "open" else -1 if t.kind == "close" else 0
        if depth < 0:
            break
    if depth != 0:
        pr.lint_errors.append(f"書き込み後の RTF のグループ（{{ と }}）の数が合いません（{depth}）")
    pa = [p for _k, _p, bl in a.all_blocks() for p in iter_paras(bl)]
    pb = [p for _k, _p, bl in b.all_blocks() for p in iter_paras(bl)]
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.unexpected.append(f"段落の数が変わっています（{len(pa)} → {len(pb)}）")
        return pr
    sp = Spots(a)
    index = {id(p): i for i, p in enumerate(pa)}
    targets: dict[int, list] = {}
    used: dict = {}
    for fid, value in written.items():
        if isinstance(value, dict) and "@" in value:
            try:
                it = resolve_write(sp, value["@"], used)
            except WriteError as e:
                pr.unexpected.append(f"書いた所 {fid} の目印が、元の文書で決まりません：{e}")
                continue
            if it["id"] != fid or it["value"] != value["value"]:
                pr.unexpected.append(f"書いた所 {fid} が、目印から決まる所（{it['id']}「{it['value']}」）と違います")
                continue
            value = value["value"]
        else:
            it = sp.items.get(fid)
        if it is None:
            pr.unexpected.append(f"書いた欄 {fid} が、元の文書にありません")
            continue
        p = it.get("para") if it.get("control") is None else it["control"].para
        if p is None or id(p) not in index:
            pr.unexpected.append(f"書いた欄 {fid} の段落が見つかりません")
            continue
        targets.setdefault(index[id(p)], []).append((fid, it, value))
    # バイトの比べ：書いた段落の範囲の外は、書く前と後で同じ
    wins = sorted((_window(pa[i]) + (i,) for i in targets), key=lambda x: x[0])
    merged: list[list] = []
    for lo, hi, i in wins:
        if merged and lo < merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], hi)
            merged[-1][2].append(i)
        else:
            merged.append([lo, hi, [i]])
    at_a = at_b = 0
    delta = 0
    for lo, hi, idx in merged:
        if original[at_a:lo] != new[at_b:lo + delta]:
            pr.unexpected.append("書いた段落の外のバイトが変わっています")
            break
        # 書いた後の、同じ範囲の終わり：範囲の中で、終わりがいちばん後ろの段落の印の終わり
        last = max(idx, key=lambda i: _window(pa[i])[1])
        hi_b = _window(pb[last])[1]
        if hi_b < lo + delta:
            pr.unexpected.append("書いた段落の範囲を決められません")
            break
        at_a, at_b = hi, hi_b
        delta = hi_b - hi
    if original[at_a:] != new[at_b:]:
        pr.unexpected.append("書いた段落より後ろのバイトが変わっています")
    for i, items in targets.items():
        p_old, p_new = pa[i], pb[i]
        pr.paragraphs_changed += 1
        plain = [(f, it, v) for f, it, v in items if it.get("control") is None]
        if len(plain) == len(items):
            want = _expected_text(p_old, plain)
            if want is not None and p_new.text != want:
                pr.unexpected.append(f"段落 {i + 1}：書いた後の字が「{p_new.text[:60]}」で、思っていた「{want[:60]}」と違います")
    for i, (x, y) in enumerate(zip(pa, pb)):
        if i not in targets and x.text != y.text:
            pr.unexpected.append(f"書いていない段落 {i + 1} の字が変わっています（「{x.text[:30]}」→「{y.text[:30]}」）")
    return pr


# ====================================================================== api から
def api_fill(blob: bytes, name: str, values: dict, force: bool = False) -> dict:
    from ..openkinds import OpenFormatError

    try:
        res = fill(RtfDocument(blob, name), values, force=force)
    except FillError as e:
        raise OpenFormatError(str(e)) from None
    pr = prove(blob, res.blob, written_of(res.patches))
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20]}
    from ..openkinds import _formatting, _wrapped

    after: list = []  # 書いた後の文書（読み直すのは 1 回だけ）

    def reread():
        if not after:
            after.append(RtfDocument(res.blob, name))
        return after[0]

    _formatting(result, lambda: written_formatting(res.blob, name, reread()))
    # 書いたことで段落の行が折り返す見込み（目安）：証明・読み戻しは字しか見ないので、見た目の変化を伝える
    _wrapped(result, blob, "rtf", lambda: _wrap(RtfDocument(blob, name), reread()))
    readback = {key: v for key, v in values.items() if not key.startswith("@")}
    anchored = {p.key: p.id for p in res.patches if p.anchor}
    if anchored:
        result["anchored"] = anchored
        readback.update({p.id: p.value for p in res.patches if p.anchor})
    return {"blob": res.blob, "result": result, "readback": readback, "labels": {}}


def _wrap(before, after) -> list[str]:
    from .wrap import wrap_warnings

    return wrap_warnings(before, after)


def written_formatting(blob: bytes, name: str = "", doc=None) -> list[dict]:
    """書いた後の文書を読み直し、書いた所（しおり _todoku_…）の字に、消したように・目立って見える書式があれば
    [{id, formatting}]（読み戻しで伝える）。doc：もう読み直してあれば、その文書。"""
    doc = doc if doc is not None else RtfDocument(blob, name)
    colors = color_table(blob)
    out = []
    for fid, (p, a, b) in doc.marks().items():
        pos = p.positions()
        found: list[str] = []
        for k in range(a, min(b, len(pos))):
            src = p.segs[pos[k][0]].node
            for w in bad_looks(getattr(src, "look", None), colors):
                if LOOK_NAMES[w] not in found:
                    found.append(LOOK_NAMES[w])
        if found:
            out.append({"id": fid, "formatting": found})
    return out
