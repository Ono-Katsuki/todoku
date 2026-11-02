"""一太郎の文書の、字数を変える書き込み（空のマス・段落に字を足す、伏せ字を長い値に替える）と、それに合わせて直す構造。

一太郎の形式には公開された仕様がないので、集めた文書（corpus/wide_jtd の 88 件）を比べて意味を確かめた構造だけを直し、
確かめられないものがあるときは、その書き込みを断る（推測で直さない）。字数を変えるのは本文（DocumentText）だけ。

直す構造（どれも、集めた文書のすべてで形を確かめたもの）：
1. 本文の字の並び（DocumentText の「TextV.01」のセグメント）：単位の数と、単位の並び
2. 文字の書式の区切り（字の並びのすぐ後ろ）：「00＋4 バイトの数」（その数の単位を、今の書式で）と
   「FE {種類 長さ 値}… FF 00」（書式を変える）の並び。数の合計は単位の数以下で、残りは最後の書式のまま。
   字を足した所を含む区切りの数を、足した字の数だけ増やす（書式の切り替わる所にちょうど足すときは、どちらの書式か
   決められないので断る）
3. Ssmg の入れ物：セグメントの目録のバイト数・ブロックの数・ブロックの番号（足りなければ末尾にブロックを足す）と、
   先頭のブロックの数
4. ストリームの一覧（\\x04JSRV_SegmentInformation）：48 バイトの見出しと、96 バイトの項目（名前・大きさ・種類）。
   大きさが変わったストリームの大きさを直し、取り除いたストリームの項目を除いて、見出しの項目の数を減らす
   （同じ様式の組で、行・ページの印を持たない文書は、見出しが項目の数だけ違い、その項目がないことを確かめた）
5. 字を打つ位置（DocumentEditStyles の 0x2000 の区切り：ページ・位置・…）：書いた所より後ろなら、位置をずらす
6. 表の位置の表（DocumentTextPositionTables の「TCntV.01」）：表の範囲の始まりと終わり（本文の単位の位置）。
   書いた所より後ろの位置をずらす
7. 行・ページの割り付けの控え（LineMark・PageMark・PaperMark）：行ごとの単位の数とページの区切り。字数が変わると
   合わなくなり、割り付けは一太郎が字と書式から作り直すものなので、取り除く（集めた文書にも、この 3 つを持たない
   一太郎の文書があり、そのストリームの一覧は、この 3 つの項目がないだけで同じ形）
本文の字の位置を持たないことを確かめたストリーム（POSITION_FREE）は、1 バイトも変えない。それ以外のストリームがある
文書は、字の位置を持つかを確かめていないので、字数を変える書き込みを断る。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from ..cfb import CFB, CfbEditor
from ..docx.fill import FillError
from .book import SSMG, TEXT, JtdError, JtdFile, ssmg, tokens

LAYOUT = ("LineMark", "PageMark", "PaperMark")
UPDATED = {
    "DocumentText": "本文の字の並び・文字の書式の区切り・Ssmg の目録",
    "\x04JSRV_SegmentInformation": "ストリームの一覧（大きさ・取り除いた項目）",
    "DocumentEditStyles": "字を打つ位置",
    "DocumentTextPositionTables": "表の位置の表",
}
# 本文の字の位置を持たないことを確かめたストリーム（名前 → 確かめたこと）
POSITION_FREE = {
    "Font": "フォントの名前の表",
    "Header": "柱の文字（自分の字の並びだけを持つ）",
    "Footnote": "脚注の文字（自分の字の並びだけを持つ）",
    "PageLayoutStyleHeader": "ページスタイルの柱の文字（自分の字の並びだけを持つ）",
    "PageLayoutStyle": "ページのスタイル",
    "TextLayoutStyle": "文字・段落のスタイルの定義（スタイルの名前と書式）",
    "ReferenceInfo": "文書の識別の情報（すべての文書で 47 バイトの同じ形）",
    "RelatedDocuments": "関連する文書の識別の情報",
    "ThinkingTemplate": "テンプレートの情報（12 バイト）",
    "DocumentPeripheralThree": "文書の周辺の設定",
    "DocumentViewStyles": "表示の設定",
    "\x04JSRV_SummaryInformation": "文書の情報（作成者・日時など）",
    "\x05SummaryInformation": "文書のプロパティ",
    "\x05DocumentSummaryInformation": "文書のプロパティ",
    "ObjectSheets/SheetInfo": "シートの名前",
    "ObjectSheets/\x04JSRV_SegmentInformation": "シートの入れ物の一覧",
    # 字の数・表の位置の違う文書どうしで、バイト列がまったく同じ（本文の位置を持たない）
    "ItemizationData": "箇条書きの設定（字の数の違う 4 つの文書で同じ 188 バイト）",
    "OutlineTag": "アウトラインの印（字の数の違う 2 つの文書で同じ 16 バイト）",
    # 形を確かめて、位置の値が「未設定」の印（7FFFFFFF・FFFF…）だけのときに限る（_autotext_ok）
    "AutoTextInfo": "自動の文字の表（位置の値は、すべて未設定の印）",
}
FREE_PREFIX = ("DocumentMacro/",)
# 集めた文書で確かめた、本文の字の並び（TextV.01）の 1 つのセグメントの大きさ（いちばん大きいもの）。もっと大きな本文を、
# 別の形（「QLSTV.01」の一覧と、いくつものセグメントに分けた字の並び）で持つ文書があり、一太郎がどの大きさで分けるかは
# 確かめていないので、書き込み後もこの大きさまでにする
MAX_SEG_BYTES = 42_466


def _desc(name: str) -> str:
    return name.replace("\x04", "").replace("\x05", "")


# ====================================================================== 文字の書式の区切り
def parse_runs(rest: bytes) -> list[tuple[str, object]]:
    """字の並びの後ろ → [("n", 単位の数) | ("a", 書式を変えるバイト列)]。形が合わなければ JtdError。"""
    out: list[tuple[str, object]] = []
    pos = 0
    while pos < len(rest):
        b = rest[pos]
        if b == 0x00:
            if pos + 5 > len(rest):
                raise JtdError("文字の書式の区切りの数が途中で終わっています")
            out.append(("n", struct.unpack_from(">I", rest, pos + 1)[0]))
            pos += 5
        elif b == 0xFE:
            start = pos
            pos += 1
            while True:
                if pos + 2 > len(rest):
                    raise JtdError("文字の書式が途中で終わっています")
                t, ln = rest[pos], rest[pos + 1]
                pos += 2
                if t == 0xFF:
                    break
                pos += ln
            out.append(("a", rest[start:pos]))
        else:
            raise JtdError(f"文字の書式の区切りに、知らない印（0x{b:02X}）があります")
    return out


def runs_bytes(els) -> bytes:
    return b"".join(b"\x00" + struct.pack(">I", v) if k == "n" else v for k, v in els)


def run_starts(els) -> list[tuple[int, int, int]]:
    """[(要素の番号, 始まり, 終わり)]：数の要素ごとの、単位の範囲。"""
    out, pos = [], 0
    for j, (k, v) in enumerate(els):
        if k == "n":
            out.append((j, pos, pos + v))
            pos += v
    return out


def run_at(els, i: int) -> int | None:
    """単位 i を含む、数の要素の番号（数の合計より後ろ＝最後の書式のままの所なら None）。"""
    for j, a, b in run_starts(els):
        if a <= i < b:
            return j
    return None


# ====================================================================== 書き込みの単位
@dataclass
class Op:
    """k 番目の文字のセグメントの単位 [a, b) を、text にする（b == a なら a の前に足す）。"""

    k: int
    a: int
    b: int
    text: str
    key: str = ""
    inherit: str = "both"  # 足す字の書式："prev"（前の字の書式を引き継ぐ）／"both"（前後の書式が同じときだけ）

    @property
    def delta(self) -> int:
        return len(self.text) - (self.b - self.a)

    def point(self) -> int:
        """字数が変わる所：足すなら足す位置、減らすなら減らす範囲の始まり。"""
        return self.b if self.b > self.a and self.delta > 0 else (self.a + len(self.text) if self.delta < 0 else self.a)


def map_pos(ops: list[Op], p: int, unit: bool = False) -> int | None:
    """元の単位の位置 → 書き込み後の位置。減らした範囲の中（決められない）なら None。
    unit=True：位置が「その単位」を指す値（表の位置の表：行の終わり・マスの記録）。足した所にちょうど当たる値も、
    その単位が後ろへずれるので、ずらす。unit=False（字を打つ位置）：足した所にちょうど当たる値は、足した字の前のまま。"""
    out = p
    for op in ops:
        d = op.delta
        if d > 0:
            q = op.b if op.b > op.a else op.a
            if p > q or (unit and p == q):
                out += d
        elif d < 0:
            lo, hi = op.a + len(op.text), op.b
            if lo < p < hi:
                return None
            if p >= hi:
                out += d
    return out


def splice(units: tuple[int, ...], ops: list[Op]) -> list[int]:
    out = list(units)
    for op in sorted(ops, key=lambda o: -o.a):
        out[op.a:op.b] = [ord(ch) for ch in op.text]
    return out


# ====================================================================== 文書の全体で確かめること
@dataclass
class Policy:
    ok: bool
    reasons: list[str] = field(default_factory=list)
    removes: list[str] = field(default_factory=list)


def policy(doc: JtdFile) -> Policy:
    """この文書で、本文の字数を変える書き込みができるか（文書の全体で決まること）。"""
    cached = getattr(doc, "_grow_policy", None)
    if cached is not None:
        return cached
    reasons: list[str] = []
    cfb = doc.cfb
    unknown = [n for n in cfb.streams() if n not in UPDATED and n not in LAYOUT and n not in POSITION_FREE
               and not n.startswith(FREE_PREFIX)]
    if cfb.has("AutoTextInfo") and not _autotext_ok(cfb.read("AutoTextInfo")):
        unknown.append("AutoTextInfo")
    if unknown:
        reasons.append("本文の字の位置を持つかを確かめていないストリームがあります（" + "・".join(_desc(n) for n in sorted(unknown))
                       + "）。字数を変えると、その中の位置がずれるおそれがあります")
    try:
        segs = ssmg(cfb.read("DocumentText"), "DocumentText")
        texts = [s for s in segs if s.data[:8] == TEXT]
        if len(segs) != 1 or len(texts) != 1:
            reasons.append("本文（DocumentText）に、字の並びのほかのセグメントがあります（意味を確かめていません）")
        else:
            n = struct.unpack_from(">I", texts[0].data, 8)[0]
            els = parse_runs(texts[0].data[12 + 2 * n:])
            if sum(v for k, v in els if k == "n") > n:
                reasons.append("文字の書式の区切りの合計が、字の数を超えています")
    except JtdError as e:
        reasons.append(f"本文の形が想定と違います：{e}")
    if cfb.has("DocumentTextPositionTables"):
        try:
            _tcnt_positions(cfb.read("DocumentTextPositionTables"))
        except JtdError as e:
            reasons.append(f"表の位置の表（DocumentTextPositionTables）の形を確かめられません：{e}")
    try:
        _caret(cfb.read("DocumentEditStyles")) if cfb.has("DocumentEditStyles") else None
    except JtdError as e:
        reasons.append(f"字を打つ位置（DocumentEditStyles）の形を確かめられません：{e}")
    try:
        slots_of(cfb.read("\x04JSRV_SegmentInformation"))
    except (JtdError, KeyError) as e:
        reasons.append(f"ストリームの一覧（JSRV_SegmentInformation）の形を確かめられません：{e}")
    if doc.compressed:
        try:
            from .lha import jscomp_parts

            pt = jscomp_parts(CFB(doc.blob).read("JSCompDocument"))
            if pt["ext"]:
                reasons.append("圧縮した文書の LHA の見出しに、意味を確かめていない拡張の見出しがあります")
            outer = CFB(doc.blob)
            bad = [n for n in outer.streams() if n not in ("JSCompDocument", "\x04JSRV_SegmentInformation",
                                                          "\x04JSRV_SummaryInformation", "\x05SummaryInformation")]
            if bad:
                reasons.append("圧縮した文書の外側に、意味を確かめていないストリームがあります（" + "・".join(map(_desc, bad)) + "）")
            slots_of(outer.read("\x04JSRV_SegmentInformation"))
        except (JtdError, KeyError) as e:
            reasons.append(f"圧縮した文書の入れ物の形を確かめられません：{e}")
    pol = Policy(not reasons, reasons, [n for n in LAYOUT if cfb.has(n)])
    doc._grow_policy = pol
    return pol


def _autotext_ok(data: bytes) -> bool:
    """AutoTextInfo（Ssmg）：32 バイトの見出し（1・0x1C・0・項目の数・0…）と、24 バイトの項目（0x18・番号・7FFFFFFF・
    FFFF が 10 バイト・0）だけで、位置の値がすべて未設定の印であること（集めた 3 つの文書がこの形）。"""
    try:
        segs = ssmg(data, "AutoTextInfo")
    except JtdError:
        return False
    for sg in segs:
        d = sg.data
        if len(d) < 32 or struct.unpack_from(">III", d, 0) != (1, 0x1C, 0) or d[16:32] != bytes(16):
            return False
        cnt = struct.unpack_from(">I", d, 12)[0]
        if len(d) != 32 + 24 * cnt:
            return False
        for k in range(cnt):
            e = d[32 + 24 * k:56 + 24 * k]
            if e[:4] != b"\x00\x00\x00\x18" or struct.unpack_from(">H", e, 4)[0] != k \
                    or e[6:20] != b"\x7f" + b"\xff" * 13 or e[20:] != bytes(4):
                return False
    return True


# ====================================================================== 書き込む所ごとに確かめること
def check_op(doc: JtdFile, op: Op) -> str | None:
    """字数を変える書き込みが、この所でできるか。できなければ理由（日本語）。
    inherit："prev"（前の字の書式を引き継ぐ：伏せ字・空白を長い値に替える）／"both"（前後の書式が同じときだけ：空の段落に足す）。"""
    pol = policy(doc)
    if not pol.ok:
        return pol.reasons[0]
    st = doc.texts[op.k]
    if st.stream != "DocumentText":
        where = {"LayoutBoxText": "レイアウト枠", "Header": "柱", "PageLayoutStyleHeader": "ページスタイルの柱",
                 "Footnote": "脚注"}.get(st.stream, st.stream)
        return (f"{where}の中です。字数を変える書き込みは、本文だけにします（{where}の行の割り付けとの対応を確かめていないため）")
    for ch in op.text:
        o = ord(ch)
        if o < 0x20 or o > 0xFFFD or 0xD800 <= o <= 0xDFFF:
            return f"「{ch!r}」は一太郎の文書に書けない字です（改行・制御の字・ U+FFFF より後ろの字は書けません）"
    units = st.units
    toks = list(tokens(units))
    depth = 0
    starts: dict[int, tuple] = {}
    depth_at: dict[int, int] = {}
    for t in toks:
        starts[t[1]] = t
        depth_at[t[1]] = depth
        if t[0] == "open":
            depth += 1
        elif t[0] == "close":
            depth -= 1
    depth_at[len(units)] = depth
    # 消す単位は、ふつうの字だけ
    if op.delta < 0:
        for i in range(op.a + len(op.text), op.b):
            t = starts.get(i)
            if t is None or t[0] != "ch":
                return "消す所に、字でないもの（記録・制御の字）があります"
    q = op.point()
    if depth_at.get(q, 1) != 0 or (op.b > op.a and depth_at.get(op.a, 1) != 0):
        return "ルビ・均等割り・リンクなどの組の中です。組の中の字数を変えると、組の割り付けが変わるので断ります"
    seg = st.seg
    n = len(units)
    els = parse_runs(seg.data[12 + 2 * n:])
    if op.delta > 0:
        if op.inherit == "both":
            if run_at(els, q - 1 if q else 0) != run_at(els, q):
                return ("ここは文字の書式が切り替わる所です（前と後ろで書式が違う）。足す字をどちらの書式にするかを"
                        "決められないので断ります")
        # "prev"：前の字の区切り（q-1）を増やす
    elif op.delta < 0:
        counts: dict[int | None, int] = {}
        for i in range(op.a + len(op.text), op.b):
            j = run_at(els, i)
            counts[j] = counts.get(j, 0) + 1
        for j, c in counts.items():
            if j is not None and els[j][1] - c <= 0:
                return "消すと、文字の書式の区切りが 1 つなくなります（書式の並びが変わるので断ります）"
    if doc.cfb.has("DocumentTextPositionTables"):
        st_ = {t[1] for t in toks}
        for p in _tcnt_positions(doc.cfb.read("DocumentTextPositionTables")):
            if p not in st_ and p != n:
                return "表の位置の表に、字の区切りでない位置があります（意味を確かめられないので断ります）"
            if op.delta < 0 and op.a + len(op.text) < p < op.b:
                return "消す所に、表の範囲の区切りがあります"
    return None


# ====================================================================== 表の位置の表（TCntV.01）
def _tcnt_entries(data: bytes) -> list[int]:
    """TCntV.01 のセグメント → 位置の値のある、セグメントの中のバイトの位置。
    形：「TCntV.01」・2 バイトの印（1）・4 バイトの数・2 バイトの値（0x24）・29 バイトの項目（始まり・終わりの位置 4 バイトずつと、
    行・列などの値）×数・「FF FF」。"""
    if data[:8] != b"TCntV.01" or len(data) < 18:
        raise JtdError("表の位置の表のセグメントの形が想定と違います")
    flag = struct.unpack_from(">H", data, 8)[0]
    if flag != 1:
        raise JtdError("表の位置の表が、表の範囲の形（印 1）ではありません")
    cnt, ver = struct.unpack_from(">IH", data, 10)
    if ver != 0x24 or len(data) != 16 + 29 * cnt + 2 or data[-2:] != b"\xff\xff":
        raise JtdError("表の位置の表の長さ・終わりの印が想定と違います")
    out = []
    for k in range(cnt):
        out += [16 + 29 * k, 16 + 29 * k + 4]
    return out


def _tcnt_positions(stream: bytes) -> list[int]:
    out = []
    for s in ssmg(stream, "DocumentTextPositionTables"):
        for at in _tcnt_entries(s.data):
            out.append(struct.unpack_from(">I", s.data, at)[0])
    return out


def tcnt_rewrite(stream: bytes, ops: list[Op]) -> tuple[bytes, list[tuple[int, int, int]]]:
    """表の位置の表の位置を、書き込みに合わせてずらす → (新しいストリーム, [(ストリームの中の位置, 元, 新)])。"""
    out = bytearray(stream)
    changes = []
    for s in ssmg(stream, "DocumentTextPositionTables"):
        for at in _tcnt_entries(s.data):
            p = struct.unpack_from(">I", s.data, at)[0]
            q = map_pos(ops, p, unit=True)
            if q is None:
                raise FillError("表の範囲の区切りが、消す所の中にあります")
            if q != p:
                for j in range(4):
                    out[s.offset(at + j)] = q.to_bytes(4, "big")[j]
                changes.append((s.offset(at), p, q))
    return bytes(out), changes


# ====================================================================== 字を打つ位置（DocumentEditStyles）
def _caret(data: bytes) -> int | None:
    """DocumentEditStyles → 字を打つ位置の値の、ストリームの中のバイトの位置（0x2000 の区切りがなければ None）。
    形：4 バイトの見出し、区切り（2 バイトの種類・4 バイトの長さ・中身）の並び。0x2000 の中身：0x14・ページ・位置・…"""
    pos = 4
    found = None
    while pos < len(data):
        if pos + 6 > len(data):
            raise JtdError("区切りが途中で終わっています")
        t, ln = struct.unpack_from(">HI", data, pos)
        if pos + 6 + ln > len(data):
            raise JtdError("区切りの長さが範囲の外です")
        if t == 0x2000:
            if ln != 28 or struct.unpack_from(">I", data, pos + 6)[0] != 0x14:
                raise JtdError("字を打つ位置の区切りの形が想定と違います")
            found = pos + 6 + 8
        pos += 6 + ln
    return found


def caret_rewrite(data: bytes, ops: list[Op]) -> tuple[bytes, tuple[int, int] | None]:
    at = _caret(data)
    if at is None:
        return data, None
    p = struct.unpack_from(">I", data, at)[0]
    if p == 0xFFFFFFFF:
        return data, None
    q = map_pos(ops, p)
    if q is None:
        q = min(op.a + len(op.text) for op in ops if op.delta < 0 and op.a + len(op.text) < p < op.b)
        q = map_pos(ops, q)
    if q == p:
        return data, None
    out = bytearray(data)
    struct.pack_into(">I", out, at, q)
    return bytes(out), (p, q)


# ====================================================================== ストリームの一覧（JSRV_SegmentInformation）
def slots_of(data: bytes) -> tuple[int, list[tuple[int, str]]]:
    """→ (先頭の名前のない項目の数, [(項目の始まり, 名前)])。見出し 48 バイト（16 バイト目から 2 バイトずつ：項目の大きさ 96・
    名前のある項目の数）、96 バイトの項目（名前 64 バイト・…・68 バイト目から大きさ 4 バイト・種類 4 バイト）。"""
    if data[:8] != b"VDA_DOC\x00" or len(data) < 48:
        raise JtdError("ストリームの一覧の見出しが想定と違います")
    esize, count = struct.unpack_from("<HH", data, 16)
    if esize != 96 or (len(data) - 48) % 96:
        raise JtdError("ストリームの一覧の項目の大きさが想定と違います")
    total = (len(data) - 48) // 96
    lead = total - count
    if lead < 0:
        raise JtdError("ストリームの一覧の項目の数が合いません")
    out = []
    for k in range(lead, total):
        base = 48 + 96 * k
        name = data[base:base + 64].decode("utf-16-le", errors="replace").split("\x00")[0]
        if not name:
            raise JtdError("ストリームの一覧に、名前のない項目があります")
        out.append((base, name))
    return lead, out


def seginfo_rewrite(data: bytes, sizes: dict[str, int], removed: list[str]) -> bytes:
    lead, slots = slots_of(data)
    keep = bytearray(data[:48 + 96 * lead])
    n = 0
    for base, name in slots:
        if name in removed:
            continue
        item = bytearray(data[base:base + 96])
        if name in sizes:
            struct.pack_into("<I", item, 68, sizes[name])
        keep += item
        n += 1
    struct.pack_into("<H", keep, 18, n)
    return bytes(keep)


# ====================================================================== Ssmg の入れ物
def ssmg_rewrite(stream: bytes, sid: int, data: bytes) -> tuple[bytes, dict]:
    """Ssmg の入れ物の、番号 sid のセグメントを data に替える。ほかのセグメントのブロック・目録・後ろのバイトは元のまま。
    ブロックが足りなければ、ブロックの並びの末尾に足す（目録はその後ろにずれる）。余るブロックは、そのセグメントのまま持つ。"""
    count, bsize, nblocks = struct.unpack_from(">III", stream, 8)
    pos = 20 + nblocks * bsize
    entries = []
    for _ in range(count):
        s, z, ln, n, n2 = struct.unpack_from(">IIIII", stream, pos)
        blocks = list(struct.unpack_from(f">{n}I", stream, pos + 20))
        entries.append([s, z, ln, blocks, stream[pos:pos + 20 + 4 * n]])
        pos += 20 + 4 * n
    tail = stream[pos:]
    target = [e for e in entries if e[0] == sid]
    if len(target) != 1:
        raise FillError("書き込むセグメントが、目録に 1 つだけではありません")
    e = target[0]
    need = max(1, -(-len(data) // bsize))
    added = []
    while len(e[3]) < need:
        added.append(nblocks + len(added))
        e[3].append(added[-1])
    area = bytearray(stream[20:20 + nblocks * bsize]) + bytearray(bsize * len(added))
    for j, b in enumerate(e[3]):
        chunk = data[j * bsize:(j + 1) * bsize]
        area[b * bsize:(b + 1) * bsize] = chunk + bytes(bsize - len(chunk))
    head = bytearray(stream[:20])
    struct.pack_into(">I", head, 16, nblocks + len(added))
    directory = b""
    for s, z, ln, blocks, raw in entries:
        if s == sid:
            directory += struct.pack(">IIIII", s, z, len(data), len(blocks), len(blocks)) + struct.pack(f">{len(blocks)}I", *blocks)
        else:
            directory += raw
    return bytes(head) + bytes(area) + directory + tail, {"added_blocks": added, "blocks": list(e[3])}


# ====================================================================== 書き込み
@dataclass
class Rewrite:
    blob: bytes
    native: bytes
    removed: list[str]
    notes: list[str]


def rewrite(doc: JtdFile, ops: list[Op], edits: dict | None = None) -> Rewrite:
    """字数を変える書き込み（ops：本文の単位の書き換え）と、字数の同じ上書き（edits：(k, i) → 字）→ 新しいファイル。"""
    pol = policy(doc)
    if not pol.ok:
        raise FillError(pol.reasons[0])
    if not ops:
        raise FillError("字数を変える書き込みがありません")
    ops = list(ops)
    k0 = ops[0].k
    other: dict[tuple[int, int], str] = {}
    for (k, i), ch in (edits or {}).items():
        if k == k0:
            ops.append(Op(k, i, i + 1, ch, ""))
        else:
            other[(k, i)] = ch
    ops = sorted(ops, key=lambda o: (o.k, o.a))
    for x, y in zip(ops, ops[1:]):
        if x.k == y.k and (y.a < x.b or (y.a == x.b and (x.a == x.b or y.a == y.b))):
            raise FillError(f"{y.key}: ほかの書き込みと同じ所に重なりました")
    grow = [op for op in ops if op.delta]
    for op in grow:
        why = check_op(doc, op)
        if why:
            raise FillError(f"{op.key}: {why}")
    cfb = doc.cfb
    notes: list[str] = []
    st = doc.texts[ops[0].k]
    if any(op.k != st.k for op in ops):
        raise FillError("本文のほかの文字のセグメントには、まとめて書けません")
    units = st.units
    n = len(units)
    els = parse_runs(st.seg.data[12 + 2 * n:])
    new_els = [list(x) for x in els]
    for op in grow:
        if op.delta > 0:
            q = op.point()
            j = run_at(els, q - 1 if (op.inherit == "prev" and q) else q)
            if j is not None:
                new_els[j][1] += op.delta
        else:
            for i in range(op.a + len(op.text), op.b):
                j = run_at(els, i)
                if j is not None:
                    new_els[j][1] -= 1
    new_units = splice(units, ops)
    seg_data = TEXT + struct.pack(">I", len(new_units)) + struct.pack(f">{len(new_units)}H", *new_units) \
        + runs_bytes([tuple(x) for x in new_els])
    if len(seg_data) > max(MAX_SEG_BYTES, len(st.seg.data)):
        raise FillError(f"書き込み後の本文の字の並び（{len(seg_data):,} バイト）が、集めた一太郎の文書で確かめた 1 つのセグメントの"
                        f"大きさ（{MAX_SEG_BYTES:,} バイト）を超えます。大きな本文は、一太郎がいくつものセグメントに分けて持つことがあり、"
                        "どの大きさで分けるかを確かめていないので、書きません（書く値を短くするか、欄を分けて書いてください）")
    changed: dict[str, bytes] = {}
    text_stream, info = ssmg_rewrite(cfb.read("DocumentText"), st.sid, seg_data)
    changed["DocumentText"] = text_stream
    total = sum(op.delta for op in ops)
    notes.append(f"本文の字の並び：{n:,} 単位 → {len(new_units):,} 単位（{total:+d}）")
    if info["added_blocks"]:
        notes.append(f"本文の Ssmg の入れ物：ブロックを {len(info['added_blocks'])} 個足しました")
    removed: list[str] = []
    if grow:
        if cfb.has("DocumentTextPositionTables"):
            new_t, ch = tcnt_rewrite(cfb.read("DocumentTextPositionTables"), grow)
            if ch:
                changed["DocumentTextPositionTables"] = new_t
                notes.append(f"表の位置の表：位置 {len(ch)} か所をずらしました")
        if cfb.has("DocumentEditStyles"):
            new_e, moved = caret_rewrite(cfb.read("DocumentEditStyles"), grow)
            if moved:
                changed["DocumentEditStyles"] = new_e
                notes.append(f"字を打つ位置：{moved[0]} → {moved[1]}")
        removed = list(pol.removes)
        if removed:
            notes.append("行・ページの割り付けの控え（" + "・".join(removed) + "）を取り除きました")
    if other:
        by_stream: dict[str, dict[int, str]] = {}
        for (k, i), ch in other.items():
            o = doc.texts[k]
            by_stream.setdefault(o.stream, {})[o.unit_offset(i)] = ch
        for name, at in by_stream.items():
            if name in changed:
                raise FillError("本文と同じストリームの、ほかのセグメントには書けません")
            data = bytearray(cfb.read(name))
            for off, ch in at.items():
                struct.pack_into(">H", data, off, ord(ch))
            changed[name] = bytes(data)
            notes.append(f"{_desc(name)}：字数の同じ上書き {len(at)} 字")
    sizes = {name: len(data) for name, data in changed.items()}
    info_old = cfb.read("\x04JSRV_SegmentInformation")
    info_new = seginfo_rewrite(info_old, sizes, removed)
    if info_new != info_old:
        changed["\x04JSRV_SegmentInformation"] = info_new
    ed = CfbEditor(cfb.blob)
    for name, data in changed.items():
        ed.replace(name, data)
    for name in removed:
        ed.remove(name)
    native = ed.blob()
    if not doc.compressed:
        return Rewrite(native, native, removed, notes)
    return Rewrite(recompress(doc.blob, native), native, removed, notes)


def overwrite(native: bytes, doc: JtdFile, edits: dict) -> bytes:
    """字数の同じ上書き：単位の 2 バイトずつだけを書き換える。"""
    by_stream: dict[str, dict[int, str]] = {}
    for (k, i), ch in edits.items():
        st = doc.texts[k]
        by_stream.setdefault(st.stream, {})[st.unit_offset(i)] = ch
    blob = native
    for name, at in by_stream.items():
        c = CFB(blob)
        data = bytearray(c.read(name))
        for off, ch in at.items():
            struct.pack_into(">H", data, off, ord(ch))
        blob = c.replace(name, bytes(data))
    return blob


def recompress(outer_blob: bytes, native: bytes) -> bytes:
    """圧縮して保存した文書：中の文書を -lh5- で縮め直し、JSCompDocument と、外側のストリームの一覧の大きさを直す。"""
    from .lha import pack_jscomp

    outer = CFB(outer_blob)
    old = outer.read("JSCompDocument")
    new = pack_jscomp(old, native)
    ed = CfbEditor(outer_blob)
    ed.replace("JSCompDocument", new)
    info_old = outer.read("\x04JSRV_SegmentInformation")
    info_new = seginfo_rewrite(info_old, {"JSCompDocument": len(new)}, [])
    if info_new != info_old:
        ed.replace("\x04JSRV_SegmentInformation", info_new)
    return ed.blob()


__all__ = ["Op", "Policy", "Rewrite", "policy", "check_op", "rewrite", "overwrite", "recompress", "parse_runs",
           "run_at", "run_starts", "map_pos", "splice", "slots_of", "LAYOUT", "UPDATED", "POSITION_FREE", "SSMG"]
