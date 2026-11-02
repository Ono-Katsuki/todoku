"""一太郎の文書の、字数を変える書き込み（空のマス・段落に字を足す、伏せ字を長い値に替える）と、それに合わせて直す構造。

一太郎の形式には公開された仕様がないので、集めた文書（corpus/wide_jtd の 88 件）を比べて意味を確かめた構造だけを直し、
確かめられないものがあるときは、その書き込みを断る（推測で直さない）。字数を変えるのは本文（DocumentText）だけ。

直す構造（どれも、集めた文書のすべてで形を確かめたもの）：
1. 本文の字の並び（DocumentText の「TextV.01」のセグメント、または「QLSTV.01」の一覧と片）：単位の数と、単位の並び
2. 文字の書式の区切り（字の並び・片のすぐ後ろ）：「00＋4 バイトの数」（その数の単位を、今の書式で）と
   「FE {種類 長さ 値}… FF 00」（書式を変える）の並び。数の合計は単位の数以下で、残りは最後の書式のまま。
   字を足した所を含む区切りの数を、足した字の数だけ増やす。
   書式の切り替わる所（足す所の前の単位＝段落・マスの始まりの記録と、足す所の単位＝段落の終わりの印などの書式が、
   どれかの読み方（下の 9.）で違う所）にちょうど足すときは、足す字を、それだけの区切りにして、書式をすべて言い直して
   書く：足す所の前に「足す所の単位の書式を言い直す区切り（9. の片の頭と同じ、重ねた値を 1 つにまとめた区切りと、
   最後の区切りそのもの）」と「足した字の数」を入れ、その後ろは元の区切り（足す所の単位の書式を変える区切り）のまま。
   こうすると、足した字の書式は、書式を「重ねる」と読んでも「置き換える」と読んでも、前の記録の書式から引き継ぐかに
   よらず、足す所の単位の書式になり、ほかの単位の書式も元のまま（証明で 4 つの読み方とも確かめる）。
   足す字を、足す所の単位（空の段落・マスでは段落の終わりの印）の書式にするわけ：Word でも一太郎でも、空の段落に
   打った字は、段落の終わりの印の前に入り、終わりの印の書式で出る。集めた文書の書式例と、同じ様式の白紙を比べると、
   白紙で空のマスに、書式例で字を入れたマスのうち、前の記録と終わりの印の書式が違うマスは 2 つ（056d416c7448 の
   T2.R9.C2「所有権移転」・T4.R12.C2「東京都千代田区霞が関○丁目○番の土地」。白紙は 049e0b3a054a・0795105088db）で、
   どちらも字が終わりの印と同じ書式（前の記録とは、種類 3 と 20 が違う）。前の記録と同じものはなかった（ほかの埋めた
   マスは、記録・字・終わりの印がみな同じ書式）。なお、様式の見出しの字なども含めて、段落の始まりの記録の後ろに 1 つの
   書式の字があり、その記録と段落の終わりの印の書式が違う段落（1 つの字の並びの本文の 139 段落）を数えると、字が
   終わりの印と同じ書式のもの 71（表のマス 33）、前の記録と同じもの 65（表のマス 27）、どちらとも違うもの 3 で、
   見出しの字は作った人が書式を変えたものもあるので、これだけでは決められない
3. Ssmg の入れ物：セグメントの目録のバイト数・ブロックの数・ブロックの番号（足りなければ末尾にブロックを足す）・
   足した片のセグメントの目録と、先頭のセグメントの数・ブロックの数
4. ストリームの一覧（\\x04JSRV_SegmentInformation）：48 バイトの見出しと、96 バイトの項目（名前・大きさ・種類）。
   大きさが変わったストリームの大きさを直し、取り除いたストリームの項目を除いて、見出しの項目の数を減らす
   （同じ様式の組で、行・ページの印を持たない文書は、見出しが項目の数だけ違い、その項目がないことを確かめた）
5. 字を打つ位置（DocumentEditStyles の 0x2000 の区切り：ページ・位置・…）：書いた所より後ろなら、位置をずらす
6. 表の位置の表（DocumentTextPositionTables の「TCntV.01」）。書いた所より後ろの位置をずらす：
   - 表の範囲の形（印 1）：項目ごとに、表の範囲の始まりと終わり（本文の単位の位置）・2 バイト・左右の位置 2 バイトずつ・
     数 1 バイト・その数のバイト・13 バイト（集めた 13 文書の 28 項目すべてが、この長さで終わりの「FF FF」に合う）
   - しおりの形（印 0）：「MarkV.01」・しおりの番号 2 バイト・始まり 4 バイト・終わり 4 バイト（FFFFFFFF は点のしおり）。
     しおりの名札（MarkTag）が 1 つの 5 文書だけで確かめたので、その形（28 バイト）に限る。同じ様式の 3 文書で、
     同じ字（「１階　４３・００平方メートル」の頭）を指す値が 2,309・2,203・1,962 と、その前の字数の違いだけずれていて、
     本文の字の位置
7. 校正の情報（CorrectAdditionalData）：項目（番号 2 バイト・0・位置 4 バイト・…）の位置は、本文の中の、同じ番号を持つ
   校正の部品（記録の種類 0、部品の種類 0x41。番号は記録の後ろから 4 番目の単位）の記録の始まり（FFFFFFFF は本文に
   部品のない項目。8cd333a10d8a の 0x83 のように、部品があっても未設定の項目もあり、未設定のまま変えない）。文書ごとに、
   位置のあるすべての項目がこの対応になっていることを確かめてから、位置をずらす
8. 行・ページの割り付けの控え（LineMark・PageMark・PaperMark）：行ごとの単位の数とページ・用紙の区切り。字数が変わると
   合わなくなり、正しく直すには一太郎の行の割り付け（字の幅・禁則など）を再現する必要があるので、取り除く。
   集めた文書にも、この 3 つを持たない一太郎の文書（209104c1692a：文部省の交付申請の様式で、一太郎の文書の情報に
   更新回数・総編集時間・前回印刷日時があり、一太郎が保存したもの）があり、そのストリームの一覧は、この 3 つの項目が
   ないだけで同じ形
9. 大きな本文の分け方：本文の字の並びは、集めた文書では 1 つのセグメントで 42,466 バイト（10,498 単位）まで。それより
   大きな本文（61a7b1ebaf32 の 34,412 単位）は、「QLSTV.01」の一覧と、17,206 単位ずつの 2 つの片（42,526・44,286 バイト）に、
   記録の途中でも単位の数で等しく分けてあった。片の単位の数は一覧にしか書かれないので、一太郎は片ごとの数を読む。
   書き込み後の字の並びが、1 つのセグメントで確かめた大きさを超えるときは、この形に変え、どの片も確かめた大きさ
   （17,206 単位・44,286 バイト）に収まるいちばん少ない数で、単位の数で等しく分ける。分ける所より前でその片が書式を
   変えていれば、後ろの片の頭で、その書式を言い直す（それまでの区切りを 1 つにまとめた区切りと、最後の区切りそのもの）。
   書式を変える区切りを「前の書式に重ねる」と読んでも「そこで書式を置き換える」と読んでも、片の頭で書式を
   「前の片から引き継ぐ」と読んでも「初めに戻る」と読んでも、どの字の書式も元のままになる（証明で 4 つの読み方とも
   確かめる）。もとから分けて持つ本文は、片の分かれ目をそのままにし、確かめた大きさを超える片だけを分ける
本文の字の位置を持たないことを確かめたストリーム（POSITION_FREE）は、1 バイトも変えない。それ以外のストリームがある
文書は、字の位置を持つかを確かめていないので、字数を変える書き込みを断る。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from ..cfb import CFB, CfbEditor
from ..docx.fill import FillError
from .book import QLST, SSMG, TEXT, JtdError, JtdFile, Piece, TextStory, parse_runs, qlst_entries, ssmg, tokens

LAYOUT = ("LineMark", "PageMark", "PaperMark")
UPDATED = {
    "DocumentText": "本文の字の並び・文字の書式の区切り・Ssmg の目録",
    "\x04JSRV_SegmentInformation": "ストリームの一覧（大きさ・取り除いた項目）",
    "DocumentEditStyles": "字を打つ位置",
    "DocumentTextPositionTables": "表の位置の表・しおりの位置",
    "CorrectAdditionalData": "校正の情報の位置",
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
    # 枠・図・リンク：本文には置き場所の記録（枠の番号・リンクの番号を持つ部品）として入っていて、その記録は字とともに動く。
    # 同じ様式の 6 文書（017df2aef553 など）で、枠 0 の置き場所が 476・561・466・604・480・342、リンクの部品が 6,621・6,614・
    # 6,657・… と違うのに、次のストリームはバイト列がまったく同じ（219b179c3828 など 3 文書でも、Frame のほかは同じ）。
    # Frame がその 3 文書で違う値は、枠の大きさ・位置（14,570 など、本文の字数 1,668 を超える値と、負の値）
    "Frame": "枠の一覧（枠の番号・種類・大きさと位置。本文の位置を持たない）",
    "FrameName": "枠の名前の表（12 文書で同じ 12 バイト）",
    "HyperLinkFrame": "リンクの枠の表（リンクの部品の位置の違う文書どうしで同じバイト列）",
    "LayoutBox": "レイアウト枠の位置・枠線（枠の置き場所の違う文書どうしで同じバイト列）",
    "LayoutBoxText": "レイアウト枠の文字（自分の字の並びだけを持つ。字数を変える書き込みは本文だけ）",
    "LayoutBoxTextPositionTables": "レイアウト枠の文字の位置の表（枠の置き場所の違う 8 文書で同じバイト列）",
    "Figure": "図形の一覧（枠の置き場所の違う文書どうしで同じバイト列）",
    "MarkTag": "しおりの名札（名前だけ。しおりの位置は表の位置の表にあり、そちらを直す）",
    "DocumentPeripheral": "文書の周辺の設定（フォントの名前など。リンク・しおりの位置の違う 2 文書で同じ 304 バイト）",
    # 段落番号のつながり：本文の段落番号の部品（記録の種類 0、部品の種類 0x0A）の 6 番目の単位と、項目の番号でつながる
    # （6 文書のすべてで一致）。項目の 4 バイトの値は本文の位置ではない（c7137e234c6d では 855 単位の本文で 1,980、
    # 252001917699 では 2,442 単位で 3,220、dd25888df78d では 3 項目とも 1,980 で、部品は 2,425・2,538・2,716 にある）。
    # 形を確かめたときに限る（_numbering_ok）
    "NumberingLink": "段落番号のつながり（本文の部品とは番号でつながり、本文の位置を持たない）",
}
FREE_PREFIX = ("DocumentMacro/", "FigureData/")  # FigureData：図形の描画（枠の置き場所の違う 3 文書で同じバイト列）
# 集めた文書で確かめた、本文の字の並びの大きさ（いちばん大きいもの）：1 つのセグメント（TextV.01）と、分けた片
MAX_SEG_BYTES = 42_466
MAX_PIECE_UNITS = 17_206
MAX_PIECE_BYTES = 44_286


def _desc(name: str) -> str:
    return name.replace("\x04", "").replace("\x05", "")


# ====================================================================== 文字の書式の区切り
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


def attrs_of(v: bytes) -> dict[int, bytes]:
    """書式を変える要素（FE {種類 長さ 値}… FF 00）→ {種類: 値}。"""
    out: dict[int, bytes] = {}
    pos = 1
    while v[pos] != 0xFF:
        out[v[pos]] = v[pos + 2:pos + 2 + v[pos + 1]]
        pos += 2 + v[pos + 1]
    return out


def fmt_key(els, i: int) -> tuple:
    """単位 i の書式を、2 つの読み方で：(前の書式に重ねた値, 最後に効いた要素のバイト列)。i より前（i の位置まで）にある
    書式を変える要素だけが効く（数の合計より後ろの単位は、すべての要素が効く）。"""
    state: dict[int, bytes] = {}
    last = None
    pos = 0
    for k, v in els:
        if k == "n":
            if pos + v > i:
                break
            pos += v
        else:
            state.update(attrs_of(v))
            last = v
    return tuple(sorted(state.items())), last


def same_format(els, i: int, j: int) -> bool:
    """単位 i と j が、どちらの読み方でも同じ書式か。"""
    return fmt_key(els, i) == fmt_key(els, j)


def restate(before: list[bytes]) -> list[tuple[str, bytes]]:
    """片の頭で、それより前の書式を言い直す要素：重ねた値を 1 つにまとめた要素（種類の順）と、最後の要素そのもの。
    前で書式を変えていなければ、何も足さない。"""
    if not before:
        return []
    state: dict[int, bytes] = {}
    for v in before:
        state.update(attrs_of(v))
    last = before[-1]
    if attrs_of(last) == state:
        return [("a", last)]
    whole = b"\xfe" + b"".join(bytes([t, len(x)]) + x for t, x in sorted(state.items())) + b"\xff\x00"
    return [("a", whole), ("a", last)]


def boundary_at(els, q: int) -> tuple[int, int] | None:
    """単位 q の前で書式が切り替わるなら (j, k)：j は q の位置にある書式を変える区切りの、はじめの要素の番号、k は q を含む
    数の要素（なければ並びの終わり）の番号。q の前と q の書式が、どの読み方でも同じなら None。"""
    if q <= 0 or same_format(els, q - 1, q):
        return None
    pos = 0
    j = None
    for idx, (kind, v) in enumerate(els):
        if kind == "n":
            if pos + v > q:
                return (j, idx) if j is not None else None
            pos += v
        elif pos == q and j is None:
            j = idx
    return (j, len(els)) if j is not None else None


def restated_block(els, q: int, count: int) -> tuple[int, list[tuple[str, object]]]:
    """書式の切り替わる所 q に count 個の単位を足すとき、元の要素の番号 j の前に入れる区切り：q の単位の書式を言い直す
    区切り（restate：q までの区切りを重ねた値を 1 つにまとめた区切りと、最後の区切りそのもの）と「count」。その後ろの、
    q の書式を変える区切りと数は元のまま。"""
    jk = boundary_at(els, q)
    if jk is None:
        raise FillError("書式の切り替わる所ではありません")
    j, k = jk
    return j, restate([v for kind, v in els[:k] if kind == "a"]) + [("n", count)]


def split_runs(els, s: int) -> tuple[list, list]:
    """書式の区切りを、単位 s の前と後ろに分ける（後ろの頭で、前の書式を言い直す）。"""
    left: list = []
    right: list = []
    pos = 0
    for k, v in els:
        if k == "n":
            if pos + v <= s:
                left.append(("n", v))
            elif pos >= s:
                right.append(("n", v))
            else:
                left.append(("n", s - pos))
                right.append(("n", pos + v - s))
            pos += v
        else:
            (left if pos < s else right).append(("a", v))
    return left, restate([v for k, v in left if k == "a"]) + right


# ====================================================================== 書き込みの単位
@dataclass
class Op:
    """k 番目の文字のセグメントの単位 [a, b) を、text にする（b == a なら a の前に足す）。"""

    k: int
    a: int
    b: int
    text: str
    key: str = ""
    # 足す字の書式："prev"（前の字の書式を引き継ぐ）／"next"（足す所の単位の書式。前と違えば、書式を言い直して足す）
    inherit: str = "next"

    @property
    def delta(self) -> int:
        return len(self.text) - (self.b - self.a)

    def point(self) -> int:
        """字数が変わる所：足すなら足す位置、減らすなら減らす範囲の始まり。"""
        return self.b if self.b > self.a and self.delta > 0 else (self.a + len(self.text) if self.delta < 0 else self.a)


def map_pos(ops: list[Op], p: int, unit: bool = False) -> int | None:
    """元の単位の位置 → 書き込み後の位置。減らした範囲の中（決められない）なら None。
    unit=True：位置が「その単位」を指す値（表の位置の表：行の終わり・マスの記録、しおり、校正の部品の記録）。足した所に
    ちょうど当たる値も、その単位が後ろへずれるので、ずらす。unit=False（字を打つ位置）：足した所にちょうど当たる値は、
    足した字の前のまま。"""
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


def piece_of(st: TextStory, op: Op) -> int:
    """書き込みが当たる片の番号。足す所は、前の字の書式を引き継ぐなら前の字の片、そうでなければ後ろの字の片。
    書き換える・減らす範囲が片をまたぐなら FillError。"""
    pcs = st.piece_list()

    def at(i: int) -> int:
        for j, pc in enumerate(pcs):
            if pc.start <= i < pc.start + pc.n:
                return j
        return len(pcs) - 1  # 字の並びの終わり

    if op.b > op.a:
        js = {at(op.a), at(op.b - 1)}
        if len(js) != 1:
            raise FillError(f"{op.key}: 書き換える字が、本文の分けて持つ字の並びの片の分かれ目をまたいでいます")
        return js.pop()
    q = op.a
    if op.inherit == "prev" and q:
        return at(q - 1)
    if q and op.inherit == "next" and at(q - 1) != at(q):
        raise FillError(f"{op.key}: 本文の分けて持つ字の並びの片の分かれ目に、ちょうど足す所です（どちらの片の書式か決められません）")
    return at(q)


# ====================================================================== 文書の全体で確かめること
@dataclass
class Policy:
    ok: bool
    reasons: list[str] = field(default_factory=list)
    removes: list[str] = field(default_factory=list)


def body_story(doc: JtdFile) -> TextStory | None:
    """本文の字の並び（1 つだけのとき）。"""
    body = [t for t in doc.texts if t.stream == "DocumentText"]
    return body[0] if len(body) == 1 else None


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
    if cfb.has("NumberingLink") and not _numbering_ok(doc, cfb.read("NumberingLink")):
        unknown.append("NumberingLink")
    if unknown:
        reasons.append("本文の字の位置を持つかを確かめていないストリームがあります（" + "・".join(_desc(n) for n in sorted(unknown))
                       + "）。字数を変えると、その中の位置がずれるおそれがあります")
    st = body_story(doc)
    try:
        segs = ssmg(cfb.read("DocumentText"), "DocumentText")
        if st is None:
            reasons.append("本文（DocumentText）に、字の並びが 1 つだけではありません（意味を確かめていません）")
        else:
            mine = {st.sid} | {pc.seg.sid for pc in (st.pieces or [])}
            if {s.sid for s in segs} != mine:
                reasons.append("本文（DocumentText）に、字の並び（と、その片）のほかのセグメントがあります（意味を確かめていません）")
            for pc in st.piece_list():
                els = parse_runs(st.runs_of(pc))
                if sum(v for k, v in els if k == "n") > pc.n:
                    reasons.append("文字の書式の区切りの合計が、字の数を超えています")
    except JtdError as e:
        reasons.append(f"本文の形が想定と違います：{e}")
    if cfb.has("DocumentTextPositionTables"):
        try:
            _tcnt_positions(cfb.read("DocumentTextPositionTables"), cfb)
        except JtdError as e:
            reasons.append(f"表の位置の表（DocumentTextPositionTables）の形を確かめられません：{e}")
    if cfb.has("CorrectAdditionalData"):
        try:
            if st is None:
                raise JtdError("本文が 1 つではありません")
            _correct_check(cfb.read("CorrectAdditionalData"), st.units)
        except JtdError as e:
            reasons.append(f"校正の情報（CorrectAdditionalData）の位置を確かめられません：{e}")
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


NUMBERING_OBJ = 0x0A  # 段落番号の部品（記録の種類 0 の 2 番目の単位）


def _numbering_ok(doc: JtdFile, data: bytes) -> bool:
    """NumberingLink：見出し（0001・0002・全体のバイト数 4 バイト・項目の数 2 バイト）、30 バイトの項目（番号・0x18・次・前・
    4 バイトの値・…・本文の部品の番号 2 バイト・…）、終わり「FFFF 0000」の形で、項目の番号が 0 から順で、どの項目の
    部品の番号も、本文の段落番号の部品の 6 番目の単位にあること（集めた 6 文書がこの形）。"""
    if len(data) < 14 or data[:4] != b"\x00\x01\x00\x02":
        return False
    total, cnt = struct.unpack_from(">IH", data, 4)
    if total != len(data) or len(data) != 10 + 30 * cnt + 4 or data[-4:] != b"\xff\xff\x00\x00":
        return False
    st = body_story(doc)
    if st is None:
        return False
    objs = {t[2][1][5] for t in tokens(st.units)
            if t[0] == "rec" and t[2][0] == 0 and len(t[2][1]) > 5 and t[2][1][1] == NUMBERING_OBJ}
    for k in range(cnt):
        e = data[10 + 30 * k:40 + 30 * k]
        idx, size = struct.unpack_from(">HH", e, 0)
        ref = struct.unpack_from(">H", e, 22)[0]
        if idx != k or size != 0x18 or ref not in objs:
            return False
    return True


# ====================================================================== 書き込む所ごとに確かめること
def check_op(doc: JtdFile, op: Op) -> str | None:
    """字数を変える書き込みが、この所でできるか。できなければ理由（日本語）。
    inherit："prev"（前の字の書式を引き継ぐ：伏せ字・空白を長い値に替える）／"next"（足す所の単位の書式：空の段落・マスに
    足すなら段落の終わりの印、段落の頭に足すなら頭の字。前の記録と書式が違えば、書式を言い直して足す）。"""
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
    try:
        j = piece_of(st, op)
    except FillError as e:
        return str(e).split(": ", 1)[-1]
    pc = st.piece_list()[j]
    els = parse_runs(st.runs_of(pc))
    lq = q - pc.start
    if op.delta > 0:
        if op.inherit == "next" and lq > 0 and not same_format(els, lq - 1, lq):
            # 書式の切り替わる所：足す字を、足す所の単位の書式を言い直した区切りにして書く（plan_body）。足す所の単位が
            # この片にない（片・字の並びの終わり）か、字を書き換えて足すときは、言い直す書式が決まらないので断る
            if lq >= pc.n or op.b > op.a:
                return ("ここは文字の書式が切り替わる所で、足す所の後ろに、書式を合わせる字（段落の終わりの印など）が"
                        "ないので、足す字の書式を決められません")
        # "prev"：前の字の区切り（q-1）を増やす。"next"：足す所の単位の区切りを増やす（書式の切り替わる所なら、言い直す）
    elif op.delta < 0:
        counts: dict[int | None, int] = {}
        for i in range(op.a + len(op.text), op.b):
            jj = run_at(els, i - pc.start)
            counts[jj] = counts.get(jj, 0) + 1
        for jj, c in counts.items():
            if jj is not None and els[jj][1] - c <= 0:
                return "消すと、文字の書式の区切りが 1 つなくなります（書式の並びが変わるので断ります）"
    st_ = {t[1] for t in toks}
    n = len(units)
    if doc.cfb.has("DocumentTextPositionTables"):
        for p in _tcnt_positions(doc.cfb.read("DocumentTextPositionTables"), doc.cfb):
            if p not in st_ and p != n:
                return "表の位置の表に、字の区切りでない位置があります（意味を確かめられないので断ります）"
            if op.delta < 0 and op.a + len(op.text) < p < op.b:
                return "消す所に、表の範囲の区切り・しおりの位置があります"
    if doc.cfb.has("CorrectAdditionalData") and op.delta < 0:
        for p in _correct_positions(doc.cfb.read("CorrectAdditionalData")):
            if op.a + len(op.text) < p < op.b:
                return "消す所に、校正の部品があります"
    return None


# ====================================================================== 表の位置の表（TCntV.01）
def _tcnt_entries(data: bytes) -> list[tuple[int, bool]]:
    """TCntV.01 のセグメント → [(位置の値のある、セグメントの中のバイトの位置, 終わりの値か)]。
    表の範囲の形（印 1）：「TCntV.01」・印 1・4 バイトの数・2 バイトの値（0x24）・項目（始まり・終わりの位置 4 バイトずつ・
    2 バイト・左右の位置 2 バイトずつ・数 1 バイト・その数のバイト・13 バイト）×数・「FF FF」。
    しおりの形（印 0）：「TCntV.01」・印 0・「MarkV.01」・しおりの番号 0・始まり・終わり（FFFFFFFF は点のしおり。28 バイト）。"""
    if data[:8] != b"TCntV.01" or len(data) < 18:
        raise JtdError("表の位置の表のセグメントの形が想定と違います")
    flag = struct.unpack_from(">H", data, 8)[0]
    if flag == 0:
        if len(data) != 28 or data[10:18] != b"MarkV.01" or data[18:20] != b"\x00\x00":
            raise JtdError("しおりの位置の表の形が、確かめた形（しおり 1 つ・28 バイト）と違います")
        return [(20, False), (24, True)]
    if flag != 1:
        raise JtdError(f"表の位置の表の印（{flag}）が、確かめた形（1：表の範囲・0：しおり）ではありません")
    cnt, ver = struct.unpack_from(">IH", data, 10)
    if ver != 0x24:
        raise JtdError("表の位置の表の値（0x24）が想定と違います")
    out = []
    pos = 16
    for _k in range(cnt):
        if pos + 28 > len(data):
            raise JtdError("表の位置の表の項目が途中で終わっています")
        nb = data[pos + 14]
        out += [(pos, False), (pos + 4, False)]
        pos += 28 + nb
    if pos + 2 != len(data) or data[pos:] != b"\xff\xff":
        raise JtdError("表の位置の表の長さ・終わりの印が想定と違います")
    return out


def _tcnt_vals(stream: bytes) -> list[tuple[object, int, int]]:
    """[(セグメント, セグメントの中の位置, 値)]。しおりの終わりの FFFFFFFF（点のしおり）は除く。"""
    out = []
    for s in ssmg(stream, "DocumentTextPositionTables"):
        for at, is_end in _tcnt_entries(s.data):
            v = struct.unpack_from(">I", s.data, at)[0]
            if is_end and v == 0xFFFFFFFF:
                continue
            out.append((s, at, v))
    return out


def _tcnt_positions(stream: bytes, cfb: CFB | None = None) -> list[int]:
    if cfb is not None and any(s.data[8:10] == b"\x00\x00" for s in ssmg(stream, "DocumentTextPositionTables")):
        if not cfb.has("MarkTag") or not _marktag_one(cfb.read("MarkTag")):
            raise JtdError("しおりの位置の表があるのに、しおりの名札（MarkTag）が確かめた形（名札 1 つ）ではありません")
    return [v for _s, _at, v in _tcnt_vals(stream)]


def _marktag_one(data: bytes) -> bool:
    """MarkTag：「MarkV.01」・名札の数 1（2 バイト）・0（4 バイト）・名前の字数（2 バイト）・名前（UTF-16）。"""
    if data[:8] != b"MarkV.01" or len(data) < 16:
        return False
    cnt, z, ln = struct.unpack_from(">HIH", data, 8)
    return cnt == 1 and z == 0 and len(data) == 16 + 2 * ln


def tcnt_rewrite(stream: bytes, ops: list[Op]) -> tuple[bytes, list[tuple[int, int, int]]]:
    """表の位置の表の位置を、書き込みに合わせてずらす → (新しいストリーム, [(ストリームの中の位置, 元, 新)])。"""
    out = bytearray(stream)
    changes = []
    for s, at, p in _tcnt_vals(stream):
        q = map_pos(ops, p, unit=True)
        if q is None:
            raise FillError("表の範囲の区切り・しおりの位置が、消す所の中にあります")
        if q != p:
            for j in range(4):
                out[s.offset(at + j)] = q.to_bytes(4, "big")[j]
            changes.append((s.offset(at), p, q))
    return bytes(out), changes


# ====================================================================== 校正の情報（CorrectAdditionalData）
CORRECT_OBJ = 0x41  # 校正の部品（記録の種類 0 の 2 番目の単位）


def _correct_entries(data: bytes) -> list[tuple[object, int, int, int]]:
    """Ssmg のセグメント：「0001 0006」・0（4 バイト）・項目の数（2 バイト）・0（2 バイト）・32 バイトの項目（番号 2 バイト・0・
    位置 4 バイト・…）× 数・後ろ → [(セグメント, 位置のセグメントの中のバイトの位置, 番号, 位置)]。"""
    out = []
    for s in ssmg(data, "CorrectAdditionalData"):
        d = s.data
        if len(d) < 12 or d[:8] != b"\x00\x01\x00\x06\x00\x00\x00\x00" or d[10:12] != b"\x00\x00":
            raise JtdError("校正の情報の見出しが、確かめた形と違います")
        cnt = struct.unpack_from(">H", d, 8)[0]
        if len(d) < 12 + 32 * cnt:
            raise JtdError("校正の情報の項目が途中で終わっています")
        for k in range(cnt):
            base = 12 + 32 * k
            idx, z, pos = struct.unpack_from(">HHI", d, base)
            if z:
                raise JtdError("校正の情報の項目の形が、確かめた形と違います")
            out.append((s, base + 4, idx, pos))
    return out


def _correct_positions(data: bytes) -> list[int]:
    return [p for _s, _at, _i, p in _correct_entries(data) if p != 0xFFFFFFFF]


def correct_objects(units) -> dict[int, int]:
    """本文の校正の部品：番号（記録の後ろから 4 番目の単位）→ 記録の始まり。"""
    out: dict[int, int] = {}
    for t in tokens(units):
        if t[0] == "rec" and t[2][0] == 0 and len(t[2][1]) > 4 and t[2][1][1] == CORRECT_OBJ:
            out[t[2][1][-4]] = t[1]
    return out


def _correct_check(data: bytes, units) -> None:
    """位置のある項目はすべて、本文の、同じ番号の校正の部品の記録の始まりか。位置が FFFFFFFF（未設定）の項目は、本文に
    同じ番号の部品があっても未設定のままの文書がある（8cd333a10d8a の 0x83）ので、未設定のまま変えない。"""
    objs = correct_objects(units)
    seen = set()
    for _s, _at, idx, pos in _correct_entries(data):
        if idx in seen:
            raise JtdError(f"項目 {idx:#x} が 2 つあります")
        seen.add(idx)
        if pos == 0xFFFFFFFF:
            continue
        if objs.get(idx) != pos:
            raise JtdError(f"項目 {idx:#x} の位置 {pos} が、本文の同じ番号の校正の部品の記録の始まりではありません")
    if set(objs) - seen:
        raise JtdError("本文の校正の部品に、校正の情報の項目のないものがあります")


def correct_rewrite(data: bytes, ops: list[Op]) -> tuple[bytes, list[tuple[int, int]]]:
    out = bytearray(data)
    changes = []
    for s, at, _idx, p in _correct_entries(data):
        if p == 0xFFFFFFFF:
            continue
        q = map_pos(ops, p, unit=True)
        if q is None:
            raise FillError("校正の部品が、消す所の中にあります")
        if q != p:
            for j in range(4):
                out[s.offset(at + j)] = q.to_bytes(4, "big")[j]
            changes.append((p, q))
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
def ssmg_update(stream: bytes, changes: dict[int, bytes], added: list[tuple[int, bytes]] | None = None) -> tuple[bytes, dict]:
    """Ssmg の入れ物の、番号 sid のセグメントを changes[sid] に替え、added のセグメントを目録の後ろに足す。ほかのセグメントの
    ブロック・目録・後ろのバイトは元のまま。替えるセグメントは元のブロックを同じ順に使い、足りなければブロックの並びの末尾に
    足す（目録はその後ろにずれる）。余るブロックは、そのセグメントのまま持つ。足すセグメントは、末尾に足したブロックを使う。"""
    added = list(added or [])
    count, bsize, nblocks = struct.unpack_from(">III", stream, 8)
    pos = 20 + nblocks * bsize
    entries = []
    for _ in range(count):
        s, z, ln, n, n2 = struct.unpack_from(">IIIII", stream, pos)
        blocks = list(struct.unpack_from(f">{n}I", stream, pos + 20))
        entries.append([s, z, ln, blocks, stream[pos:pos + 20 + 4 * n]])
        pos += 20 + 4 * n
    tail = stream[pos:]
    sids = [e[0] for e in entries]
    for sid in changes:
        if sids.count(sid) != 1:
            raise FillError("書き込むセグメントが、目録に 1 つだけではありません")
    for sid, _d in added:
        if sid in sids:
            raise FillError("足すセグメントの番号が、目録にもうあります")
    new_blocks: list[int] = []

    def need(data: bytes) -> int:
        return max(1, -(-len(data) // bsize))

    data_of: dict[int, bytes] = {}
    for e in entries:
        if e[0] in changes:
            data_of[e[0]] = changes[e[0]]
            while len(e[3]) < need(changes[e[0]]):
                new_blocks.append(nblocks + len(new_blocks))
                e[3].append(new_blocks[-1])
    for sid, data in added:
        blocks = []
        for _ in range(need(data)):
            new_blocks.append(nblocks + len(new_blocks))
            blocks.append(new_blocks[-1])
        entries.append([sid, 0, len(data), blocks, None])
        data_of[sid] = data
    area = bytearray(stream[20:20 + nblocks * bsize]) + bytearray(bsize * len(new_blocks))
    for e in entries:
        if e[0] not in data_of:
            continue
        data = data_of[e[0]]
        for j, b in enumerate(e[3]):
            chunk = data[j * bsize:(j + 1) * bsize]
            area[b * bsize:(b + 1) * bsize] = chunk + bytes(bsize - len(chunk))
    head = bytearray(stream[:20])
    struct.pack_into(">I", head, 8, len(entries))
    struct.pack_into(">I", head, 16, nblocks + len(new_blocks))
    directory = b""
    for s, z, _ln, blocks, raw in entries:
        if s in data_of:
            directory += struct.pack(">IIIII", s, z, len(data_of[s]), len(blocks), len(blocks)) \
                + struct.pack(f">{len(blocks)}I", *blocks)
        else:
            directory += raw
    return bytes(head) + bytes(area) + directory + tail, {"added_blocks": new_blocks}


def ssmg_rewrite(stream: bytes, sid: int, data: bytes) -> tuple[bytes, dict]:
    """番号 sid のセグメントだけを data に替える（ssmg_update の 1 つだけの形）。"""
    return ssmg_update(stream, {sid: data})


# ====================================================================== 本文の字の並び：片ごとに書き、大きければ分ける
def _piece_bytes(units: list[int], els) -> int:
    return 2 * len(units) + len(runs_bytes(els))


def _divide(units: list[int], els, limit_units: int, limit_bytes: int, at_least: int) -> list[tuple[list[int], list]]:
    """1 つの片（単位と書式の区切り）を、どれも大きさの決まりに収まるいちばん少ない数（at_least 以上）で、単位の数で等しく分ける。"""
    total = len(units)
    for m in range(max(1, at_least), total + 1):
        cuts = [total * i // m for i in range(m + 1)]
        parts = []
        rest = els
        for i in range(1, m):
            left, rest = split_runs(rest, cuts[i] - cuts[i - 1])
            parts.append((units[cuts[i - 1]:cuts[i]], left))
        parts.append((units[cuts[m - 1]:], rest))
        if all(len(u) <= limit_units and _piece_bytes(u, e) <= limit_bytes for u, e in parts):
            return parts
    raise FillError("本文の字の並びを、確かめた大きさの片に分けられません")


@dataclass
class BodyPlan:
    """書き込み後の本文：分けない（seg_data）か、分ける（片の並び：(番号, 単位, 書式の区切り, 元の片の番号 or None)）。"""

    single: bytes | None
    pieces: list[tuple[int, list[int], list, int | None]]
    qlst_sid: int
    units: int
    restated: int = 0  # 書式の切り替わる所に、書式を言い直して足した所の数


def plan_body(st: TextStory, ops: list[Op], used_sids: set[int]) -> BodyPlan:
    pcs = st.piece_list()
    by_piece: dict[int, list[Op]] = {}
    for op in ops:
        by_piece.setdefault(piece_of(st, op), []).append(op)
    news: list[tuple[list[int], list, Piece]] = []
    n_restated = 0
    for j, pc in enumerate(pcs):
        local = [Op(op.k, op.a - pc.start, op.b - pc.start, op.text, op.key, op.inherit) for op in by_piece.get(j, [])]
        units = list(st.units[pc.start:pc.start + pc.n])
        els = parse_runs(st.runs_of(pc))
        new_els = [list(x) for x in els]
        restated: list[tuple[int, list]] = []  # 書式の切り替わる所に足す：(元の要素の番号, その前に入れる区切り)
        for op in local:
            if op.delta > 0:
                q = op.point()
                if op.inherit == "next" and op.a == op.b and boundary_at(els, q) is not None:
                    restated.append(restated_block(els, q, op.delta))
                    continue
                jj = run_at(els, q - 1 if (op.inherit == "prev" and q) else q)
                if jj is not None:
                    new_els[jj][1] += op.delta
            elif op.delta < 0:
                for i in range(op.a + len(op.text), op.b):
                    jj = run_at(els, i)
                    if jj is not None:
                        new_els[jj][1] -= 1
        merged = [tuple(x) for x in new_els]
        n_restated += len(restated)
        for j, block in sorted(restated, key=lambda r: -r[0]):
            merged[j:j] = block
        news.append((splice(tuple(units), local), merged, pc))
    total = sum(len(u) for u, _e, _p in news)
    if st.pieces is None:
        u, e, _pc = news[0]
        seg_data = TEXT + struct.pack(">I", len(u)) + struct.pack(f">{len(u)}H", *u) + runs_bytes(e)
        if len(seg_data) <= max(MAX_SEG_BYTES, len(st.seg.data)):
            return BodyPlan(seg_data, [], st.sid, total, n_restated)
        parts = _divide(u, e, MAX_PIECE_UNITS, MAX_PIECE_BYTES, 2)
        out = []
        nxt = max(used_sids) + 1
        for pu, pe in parts:
            out.append((nxt, pu, pe, None))
            nxt += 1
        return BodyPlan(None, out, st.sid, total, n_restated)
    out = []
    nxt = max(used_sids) + 1
    for u, e, pc in news:
        lim_u = max(MAX_PIECE_UNITS, pc.n)
        lim_b = max(MAX_PIECE_BYTES, len(pc.seg.data))
        if len(u) <= lim_u and _piece_bytes(u, e) <= lim_b:
            out.append((pc.seg.sid, u, e, pc.seg.sid))
            continue
        parts = _divide(u, e, MAX_PIECE_UNITS, MAX_PIECE_BYTES, 2)
        for i, (pu, pe) in enumerate(parts):
            if i == 0:
                out.append((pc.seg.sid, pu, pe, pc.seg.sid))
            else:
                out.append((nxt, pu, pe, None))
                nxt += 1
    return BodyPlan(None, out, st.sid, total, n_restated)


def qlst_bytes(pieces) -> bytes:
    return QLST + struct.pack(">I", len(pieces)) + b"".join(struct.pack(">II", len(u), sid) for sid, u, _e, _o in pieces) + bytes(4)


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
    n = len(st.units)
    stream = cfb.read("DocumentText")
    used = {s.sid for s in ssmg(stream, "DocumentText")}
    bp = plan_body(st, ops, used)
    changed: dict[str, bytes] = {}
    if bp.single is not None:
        text_stream, info = ssmg_update(stream, {st.sid: bp.single})
    else:
        old = {sid for sid, _u, _e, orig in bp.pieces if orig is not None}
        changes = {st.sid: qlst_bytes(bp.pieces)}
        added = []
        for sid, u, e, orig in bp.pieces:
            data = struct.pack(f">{len(u)}H", *u) + runs_bytes(e)
            if orig is not None:
                changes[sid] = data
            else:
                added.append((sid, data))
        gone = {pc.seg.sid for pc in (st.pieces or [])} - old
        if gone:
            raise FillError("本文の分けて持つ字の並びの片が、書き込みでなくなります")
        text_stream, info = ssmg_update(stream, changes, added)
        sizes = "・".join(f"{len(u):,}" for _s, u, _e, _o in bp.pieces)
        if st.pieces is None:
            notes.append(f"本文の字の並びが、1 つのセグメントで確かめた大きさ（{MAX_SEG_BYTES:,} バイト）を超えるので、"
                         f"「QLSTV.01」の一覧と {len(bp.pieces)} つの片（{sizes} 単位）に分けて持つ形にしました")
        elif len(bp.pieces) != len(st.pieces):
            notes.append(f"本文の分けて持つ字の並び：確かめた大きさを超える片を分け、{len(st.pieces)} → {len(bp.pieces)} 片"
                         f"（{sizes} 単位）にしました")
    changed["DocumentText"] = text_stream
    total = sum(op.delta for op in ops)
    notes.append(f"本文の字の並び：{n:,} 単位 → {bp.units:,} 単位（{total:+d}）")
    if bp.restated:
        notes.append(f"文字の書式の切り替わる所に字を足した {bp.restated} か所：足した字を、足す所の単位（段落の終わりの印など）の"
                     "書式をすべて言い直した区切りにして書きました")
    if info["added_blocks"]:
        notes.append(f"本文の Ssmg の入れ物：ブロックを {len(info['added_blocks'])} 個足しました")
    removed: list[str] = []
    if grow:
        if cfb.has("DocumentTextPositionTables"):
            new_t, ch = tcnt_rewrite(cfb.read("DocumentTextPositionTables"), grow)
            if ch:
                changed["DocumentTextPositionTables"] = new_t
                notes.append(f"表の位置の表：位置 {len(ch)} か所をずらしました")
        if cfb.has("CorrectAdditionalData"):
            new_c, ch = correct_rewrite(cfb.read("CorrectAdditionalData"), grow)
            if ch:
                changed["CorrectAdditionalData"] = new_c
                notes.append(f"校正の情報：位置 {len(ch)} か所をずらしました")
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
    sizes_ = {name: len(data) for name, data in changed.items()}
    info_old = cfb.read("\x04JSRV_SegmentInformation")
    info_new = seginfo_rewrite(info_old, sizes_, removed)
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
           "run_at", "run_starts", "boundary_at", "restated_block", "map_pos", "splice", "slots_of", "same_format",
           "split_runs", "restate", "attrs_of",
           "LAYOUT", "UPDATED", "POSITION_FREE", "SSMG", "MAX_SEG_BYTES", "MAX_PIECE_UNITS", "MAX_PIECE_BYTES"]
