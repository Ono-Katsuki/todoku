"""PDFに書き込む：元のバイト列の後ろに、変えたオブジェクトと新しいオブジェクトを足す（増分更新）。

- フォームの欄（AcroForm）：値（/V）と見た目（/AP）を書く。チェック・ラジオは状態（/AS）も
- 配置から見つけた所（印刷用のPDF）：書いた字を、ページの描画の後ろに足し、値の控えとして同じ名前の隠した欄を足す（pagedraw.py）。
  どのビューアーでも同じ字が見える。元のページの描画（コンテンツストリーム）は1バイトも変えない。値は機械で読める（/V）ので、読み戻せる
  電子署名のある PDF と、権限の設定がページの描画の変更を許さない PDF は、見える欄（値と見た目）として足す（_may_draw_on_page）
- 文字の見た目は、同梱の書体（Noto Sans JP）に字形のある字は、それを埋め込んで描く。ない字は、欧文は Helvetica、ハングル・中国語などは
  定義済みの CMap の字形の集まりで描く（glyphs.choose）
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field

from .doc import PdfDocument
from . import glyphs, pagedraw
from .embedfont import BUNDLED_KEY
from .form import FF_COMB, FF_EDIT, FF_MULTILINE, FF_MULTISELECT, FormField, read_fields
from .layout import View
from .objects import Name, PdfError, PdfString, Ref, Stream, dumps, text_string


class FillError(ValueError):
    pass


# 書いた字がこれより小さくなったら知らせる：6pt は、印刷物の注記に使われる下限の大きさの目安（1pt = 1/72 インチ、6pt ≒ 2.1mm）。
# 書類の集まりから決めた値ではない
MIN_READABLE_PT = 6.0


@dataclass
class Patch:
    id: str
    kind: str  # input / control
    value: object
    page: int = 0
    anchor: dict | None = None  # 見出しの字を目印にした書き込み（{"after": …, "value": …}）。証明で、同じ決め方で位置を確かめる
    key: str = ""  # 指定したときの名前（「@1」など）
    seal: dict | None = None  # 本人の印影・署名の画像を置いたとき：本人に伝える文・位置・画像の SHA-256
    circle: dict | None = None  # 印刷された字を○で囲んだとき：囲んだ字・楕円（見た目の座標）・見出し（circle.resolve）
    shown: str | None = None  # 見た目に描いた字（値と違うとき：選択肢の見せる字・書式で整えた数など）。読み戻しで見た目と照らす
    ap_text: str | None = None  # 見た目に描いた字の全体（リスト：見える行の選択肢を並べたもの）。なければ shown と照らす
    expect: dict | None = None  # 書いた状態：チェック・ラジオは {"V": 名前, "AS": [(部品の番号, 状態)]}、リストは {"I": 番号}
    rect_mm: dict | None = None  # とどくが置いた字の枠：{page, x, y, w, h（mm）, size_pt, align}


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unverified: list[str] = field(default_factory=list)
    # 作り手の権限の設定（暗号の辞書の P）が許していないのに行った書き込み：[{"op": 操作の名前, "what": 説明}]（_permissions）
    denied: list[dict] = field(default_factory=list)
    structure: list = field(default_factory=list)  # 構造の木に足した所（structadd.Placed）：読み手にどこで読まれるかの事実
    # /V から描き直す表示（PDFKit）で字の終わりが欠ける欄（広げられなかったもの）：[{"id", "value", "width_mm", "needs_mm"}]
    redraw_cut: list = field(default_factory=list)
    # 埋め込まずに描いた書体：[{"font", "kind"（standard14・cjk_predefined）, "ids"}]（_fonts_not_embedded）
    fonts_not_embedded: list = field(default_factory=list)


# ====================================================================== 増分更新
class Update:
    def __init__(self, doc: PdfDocument):
        self.doc = doc
        doc.verify_xref()  # 壊れた相互参照は、書く前に読み直す（足す相互参照に、すべての正しい位置を入れるため）
        # 新しい番号は、相互参照にある番号と、ファイルに書いてあるオブジェクトの番号（相互参照が壊れていて、後で読み直す
        # ことがある）の、どれとも重ならないようにする
        from ..lazy.filebytes import iter_matches

        scanned = max((int(m[1]) for _s, m in iter_matches(re.compile(rb"(?<![\d])(\d{1,10})\s+\d{1,5}\s+obj\b"), doc.blob)), default=0)
        self.next = max(doc.max_num, scanned) + 1
        self.objs: dict[int, tuple[int, object]] = {}  # 番号 → (世代, オブジェクト)

    def new(self, obj) -> Ref:
        n = self.next
        self.next += 1
        self.objs[n] = (0, obj)
        return Ref(n, 0)

    def put(self, ref: Ref, obj) -> None:
        self.objs[ref.num] = (ref.gen, obj)

    def get(self, ref: Ref):
        """書き換えるためのオブジェクト（すでに書き換えたものはそれを、まだならコピーを）。"""
        if ref.num in self.objs:
            return self.objs[ref.num][1]
        o = self.doc.get(ref.num)
        o = copy.deepcopy(o)
        self.objs[ref.num] = (ref.gen, o)
        return o

    def _enc(self, o, num: int, gen: int):
        """暗号化された文書：文字列を、元の文書と同じ鍵と方式（StrF）で暗号化する（オブジェクトの番号と世代ごとの鍵）。"""
        dec = self.doc.decryptor
        if dec is None:
            return o
        if isinstance(o, PdfString) or (isinstance(o, bytes) and not isinstance(o, Stream)):
            out = PdfString(dec.encrypt_string(bytes(o), num, gen))
            out.hexform = True
            return out
        if isinstance(o, list):
            return [self._enc(x, num, gen) for x in o]
        if isinstance(o, dict):
            return {k: self._enc(v, num, gen) for k, v in o.items()}
        return o

    def _body(self, num: int, gen: int, o) -> bytes:
        if isinstance(o, Stream):
            # ここで書くストリームの中身は、暗号化していないもの（新しく作った見た目など）
            data = o.raw
            d = dict(o.dict)
            if self.doc.decryptor is not None:
                data = self.doc.decryptor.encrypt_stream(data, num, gen, *self.doc.crypt_of(d, num))
                d = self._enc(d, num, gen)
            d["Length"] = len(data)
            return dumps(d) + b"\nstream\n" + data + b"\nendstream"
        return dumps(self._enc(o, num, gen))

    def build(self) -> bytes:
        orig = self.doc.blob
        out = bytearray()  # 後ろに足すバイト（元のバイトは写さない：大きな PDF は最後に流して写す）
        if not orig.endswith((b"\n", b"\r")):
            out += b"\n"
        base = len(orig)
        offsets: dict[int, tuple[int, int]] = {}
        for num in sorted(self.objs):
            gen, o = self.objs[num]
            offsets[num] = (base + len(out), gen)
            out += f"{num} {gen} obj\n".encode() + self._body(num, gen, o) + b"\nendobj\n"
        prev = _startxref(orig)
        tr = self.doc.trailer
        size = max(self.next, int(tr.get("Size", 0)))
        trailer = {"Size": size, "Root": tr["Root"]}
        for k in ("Info", "ID", "Encrypt"):
            if k in tr:
                trailer[k] = tr[k]
        packed: dict[int, tuple[int, int]] = {}
        if self.doc.repaired:
            # 元の相互参照が壊れていた（読むときに、ファイルを頭から読み直して位置を集めた）：足す相互参照に、元のすべての
            # オブジェクトの正しい位置を入れ、壊れた表（Prev）をたどらなくても開けるようにする
            for num, ent in self.doc.xref.items():
                if num in offsets:
                    continue
                if ent[0] == "o":
                    offsets[num] = (ent[1], ent[2])
                elif ent[0] == "s":
                    packed[num] = (ent[1], ent[2])
            size = max(size, max(offsets, default=0) + 1, max(packed, default=0) + 1)
            trailer["Size"] = size
        else:
            trailer["Prev"] = prev
        if packed or _uses_xref_stream(orig, prev):
            xnum = max(self.next, size)
            offsets[xnum] = (base + len(out), 0)
            nums = sorted(set(offsets) | set(packed))
            index, rows = [], bytearray()
            for n in nums:
                if index and index[-2] + index[-1] == n:
                    index[-1] += 1
                else:
                    index += [n, 1]
                if n in offsets:
                    off, gen = offsets[n]
                    rows += bytes([1]) + off.to_bytes(4, "big") + gen.to_bytes(2, "big")
                else:
                    stm, idx = packed[n]
                    rows += bytes([2]) + stm.to_bytes(4, "big") + idx.to_bytes(2, "big")
            d = dict(trailer)
            d.update({"Type": Name("XRef"), "Size": max(size, xnum + 1), "Index": index, "W": [1, 4, 2], "Length": len(rows)})
            xref_pos = base + len(out)
            out += f"{xnum} 0 obj\n".encode() + dumps(d) + b"\nstream\n" + bytes(rows) + b"\nendstream\nendobj\n"
        else:
            xref_pos = base + len(out)
            out += b"xref\n"
            if self.doc.repaired and 0 not in offsets:
                out += b"0 1\n0000000000 65535 f\r\n"  # 元の表をたどらない、全部の表：0 番は空きの項目
            nums = sorted(offsets)
            i = 0
            while i < len(nums):
                j = i
                while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
                    j += 1
                out += f"{nums[i]} {j - i + 1}\n".encode()
                for n in nums[i:j + 1]:
                    off, gen = offsets[n]
                    out += f"{off:010d} {gen:05d} n\r\n".encode()
                i = j + 1
            out += b"trailer\n" + dumps(trailer) + b"\n"
        out += f"startxref\n{xref_pos}\n%%EOF\n".encode()
        from ..lazy.filebytes import FileBytes, appended_file

        if isinstance(orig, FileBytes):
            return appended_file(orig, out)  # 元のバイトを流して写した一時ファイル（写像で開く）
        return bytes(orig) + bytes(out)


def _startxref(blob: bytes) -> int:
    m = list(re.finditer(rb"startxref\s+(\d+)", blob[-4096:]))
    if not m:
        raise PdfError("startxref がありません")
    return int(m[-1][1])


def _uses_xref_stream(blob: bytes, pos: int) -> bool:
    return not blob[pos:pos + 20].lstrip().startswith(b"xref")


# ====================================================================== 文字の見た目
HELV = {"Type": Name("Font"), "Subtype": Name("Type1"), "BaseFont": Name("Helvetica"), "Encoding": Name("WinAnsiEncoding")}
ZADB = {"Type": Name("Font"), "Subtype": Name("Type1"), "BaseFont": Name("ZapfDingbats")}
CJK = glyphs.CJK  # 字形の集まり → (資源の名前, 書体, CMap, 補遺)
cjk_font = glyphs.cjk_font


def _word_spaced(ch: str) -> bool:
    """語の間を空白で区切る文字か：東アジアの幅（UAX #11）が W・F の字（和文・中文の字）でないもの。ハングルは全角でも、語の間を
    空白で区切る。"""
    import unicodedata

    o = ord(ch)
    if 0x1100 <= o <= 0x11FF or 0x3130 <= o <= 0x318F or 0xA960 <= o <= 0xA97F or 0xAC00 <= o <= 0xD7FF:
        return True
    return unicodedata.east_asian_width(ch) not in ("W", "F")


def _latin(ch: str) -> bool:
    return glyphs.winansi(ch) is not None or ch in glyphs.EXT_CODE


def runs_for(text: str) -> list[tuple[str, str]]:
    """文字列を、同じ書体で描ける並びに分ける：(資源の名前, 文字列)。描けない字は、理由を添えて FillError。"""
    try:
        return glyphs.runs(text)
    except glyphs.GlyphError as e:
        raise FillError(str(e)) from None


def text_width(text: str, size: float) -> float:
    try:
        return glyphs.width(text) * size
    except glyphs.GlyphError as e:
        raise FillError(str(e)) from None


def _show(key: str, run: str) -> bytes:
    return b"<" + glyphs.encode_run(key, run).hex().upper().encode() + b"> Tj"


def text_ops(lines: list[str], size: float, x: float, y: float, leading: float, align_w: float | None = None,
             align: str = "center") -> bytes:
    """文字の行を描く命令。x, y は1行目の基線の始まり。align_w があれば、その幅の中で align（center・right）にそろえる。"""
    out = bytearray(b"BT\n0 g\n")
    for i, line in enumerate(lines):
        lx = x
        if align_w is not None:
            room = max(0.0, align_w - text_width(line, size))
            lx = x + (room if align == "right" else room / 2)
        ly = y - i * leading
        out += f"1 0 0 1 {lx:.2f} {ly:.2f} Tm\n".encode()
        if glyphs.shaped(line):
            try:
                out += glyphs.line_ops(line, size)
            except glyphs.GlyphError as e:
                raise FillError(str(e)) from None
            continue
        for key, run in runs_for(line):
            out += f"/{key} {size:.2f} Tf ".encode() + _show(key, run) + b"\n"
    out += b"ET\n"
    return bytes(out)


def fonts_used(texts: list[str]) -> set[str]:
    keys = set()
    for t in texts:
        for k, _ in runs_for(t):
            keys.add(k)
    return keys


def wrap(text: str, size: float, width: float) -> list[str]:
    lines = []
    for para in str(text).split("\n"):
        cur = ""
        for ch in glyphs.units(para):  # 字を組み立てる文字は、音節・記号の付いた字を切り離さない
            if cur and text_width(cur + ch, size) > width:
                # 語の間を空白で区切る文字（ラテン・キリル・ギリシャ・アルメニア・グルジア文字・ハングル・字を組み立てる文字
                # など）は、行に空白があれば、その切れ目で折り返す（UAX #14 を簡単にしたもの）。和文・中文の字（東アジアの幅が
                # W・F の字。ハングルを除く）は、どの字の間でも折り返せる
                k = cur.rfind(" ")
                if _word_spaced(ch[0]) and ch != " " and k > 0:
                    lines.append(cur[:k])
                    cur = cur[k + 1:] + ch
                    continue
                lines.append(cur)
                cur = ch
            else:
                cur += ch
        lines.append(cur)
    return lines


# 1 行に書く所：行の中の所（字の間・行の前後の空き・字の形の空欄）と下線の罫線。折り返さず、1 行に入る大きさに縮める
ONE_LINE_SHAPES = ("gap", "lead", "trail", "connector", "leader", "dash", "bracket", "blank_line", "placeholder", "underline")


def spot_multiline(spot: dict, h: float, size: float, unit: float = 1.0) -> bool:
    """折り返してよい所か：何行も入る高さ（字の 2.4 行ぶんより高い）のマス・枠だけ。行の所（基線のある所）は、いつも 1 行。"""
    if spot.get("shape") in ONE_LINE_SHAPES or spot.get("baseline") is not None:
        return False
    return h > 2.4 * min(size, 10.5 / unit)


def spot_fits(text: str, rect, size: float | None, kind: str, spot: dict | None = None) -> bool:
    """配置から見つけた欄（窓の rect・size）に、add_widget と同じ決め方で text が入るか。"""
    x0, y0, x1, y1 = rect
    w, h = x1 - x0, y1 - y0
    size = float(size or 10.5)
    if w <= 0 or h <= 0:
        return False
    try:
        layout_text(text, w, h, size, center=kind == "inline", multiline=spot_multiline(spot or {}, h, size))
    except FillError:
        return False
    return True


def layout_text(text: str, w: float, h: float, size: float, center: bool = False, multiline: bool = True,
                unit: float = 1.0):
    """枠に入るように、字の大きさと折り返しを決める。(行, 大きさ)。入らなければ FillError。
    unit：ページの UserUnit（大きさ・幅はページの座標の単位。小さくする下限の 4pt と、知らせる寸法は紙の上の pt）。"""
    s = size
    pad = min(3.0, 0.2 * w)
    while s >= 4.0 / unit - 1e-9:
        lines = wrap(text, s, w - pad) if multiline else [str(text)]
        # 1 行の欄も高さを見る：見た目のフォームは BBox（欄の四角）で切り抜かれる（ISO 32000-1 8.10.2）ので、高さを越えた字は欠けて描かれる
        if len(lines) * s * 1.15 > h + 0.5:
            s -= 0.5
            continue
        if all(text_width(line, s) <= w - pad for line in lines):
            return lines, s
        s -= 0.5
    raise FillError(f"「{text}」は、欄（幅 {w * unit:.0f}pt・高さ {h * unit:.0f}pt）に入りきりません")


# ====================================================================== 見た目のストリーム
def _ap(update: Update, w: float, h: float, ops: bytes, fonts: set[str], dr_fonts, matrix=None) -> Ref:
    """見た目のフォーム。w・h は BBox の大きさ（描く向きの座標：回した欄は、欄の Rect の幅と高さを入れ替えたもの）。
    matrix で BBox を写した四角が、欄の Rect と同じ大きさになること（ISO 32000-1 12.5.5 Algorithm 8.1：違うと、ビューアーが
    縦と横を別の倍率で伸び縮みさせて描く）。"""
    res = {}
    if fonts:
        res["Font"] = {k: (dr_fonts(k) if callable(dr_fonts) else dr_fonts[k]) for k in sorted(fonts)}
    d = {"Type": Name("XObject"), "Subtype": Name("Form"), "BBox": [0, 0, round(w, 3), round(h, 3)], "Resources": res}
    if matrix is not None:
        d["Matrix"] = list(matrix)
    return update.new(Stream(d, ops))


def _rot_matrix(rot: int, w: float, h: float):
    """ページの回転（欄の /MK の R）を打ち消して、見た目の向きで描くための行列（反時計回りに rot 度）。w・h は欄の Rect の
    幅と高さ。描く向きの BBox [0 0 vw vh]（90・270 度は vw = h・vh = w）を、ちょうど [0 0 w h] に写す。"""
    if rot == 90:
        return [0, 1, -1, 0, w, 0]
    if rot == 180:
        return [-1, 0, 0, -1, w, h]
    if rot == 270:
        return [0, -1, 1, 0, 0, h]
    return None


def norm_rot(v) -> int:
    """回転の角度（/Rotate・/MK の R）を 0・90・180・270 に（90 の倍数でなければ 0）。"""
    if isinstance(v, bool) or not isinstance(v, (int, float)) or v != int(v) or int(v) % 90:
        return 0
    return int(v) % 360


def _view_box(rot: int, w: float, h: float) -> tuple[float, float]:
    """描く向きの BBox の大きさ（90・270 度は、Rect の幅と高さを入れ替える）。"""
    return (h, w) if rot in (90, 270) else (w, h)


ALIGNS = ("left", "center", "right")


def layout_options(key: str, value) -> tuple[dict, object]:
    """書く値の、置き方の指定：{"value": 値, "align": "left"|"center"|"right", "pad_mm": 余白（mm）}（"kind" と一緒でもよい）。
    返す：({"align", "pad_mm"} のうち指定のあったもの, 値（"kind" があれば {"kind", "value"} のまま））。指定がなければ ({}, 値)。"""
    if not (isinstance(value, dict) and ("align" in value or "pad_mm" in value)):
        return {}, value
    if set(value) - {"value", "align", "pad_mm", "kind", "as"} or "value" not in value:
        raise FillError(f'{key}: 置き方を決めて書くときは {{"value": 値, "align": "left"|"center"|"right", "pad_mm": 余白}} の形で'
                        "指定してください")
    opts = {k: value[k] for k in ("align", "pad_mm") if k in value}
    check_layout_options(key, opts)
    rest = {k: v for k, v in value.items() if k in ("kind", "as", "value")}
    return opts, (rest if len(rest) > 1 else rest["value"])


def check_layout_options(key: str, opts: dict) -> None:
    al, pad = opts.get("align"), opts.get("pad_mm")
    if al is not None and al not in ALIGNS:
        raise FillError(f'{key}: align は "left"（既定）・"center"・"right" のどれかです')
    if pad is not None and not (isinstance(pad, (int, float)) and not isinstance(pad, bool) and 0 <= pad <= 20):
        raise FillError(f"{key}: pad_mm は、枠の端と字のあいだの余白（mm、0〜20）です")


def text_appearance(update: Update, rect_w: float, rect_h: float, value: str, size: float, dr_fonts: dict, *,
                    center: bool = False, comb: list | None = None, rot: int = 0, multiline: bool = True,
                    cover: bool = False, align: str = "left", unit: float = 1.0,
                    pad: float | None = None, baseline: float | None = None, frame: bytes = b"") -> tuple[Ref, float]:
    """align：枠の中の左右のそろえ方（既定は左）。pad：枠の端と字のあいだの余白（pt。既定は 1.5pt か幅の 1/10 の小さいほう）。
    baseline：字の基線の、枠の下の端からの高さ（pt。行の所は行の字と同じ基線に書く。省くと、枠の上下の中ほどにそろえる）。
    cover：欄の下地を白く塗ってから書く（伏せ字の「○○」の上に書くときに、下の字を覆う。元の字は消さない）。
    rot：欄の回転（/MK の R）。描く向きの座標（vw×vh）で字を並べ、BBox も同じ大きさにして、Matrix で Rect の向きに回す。
    frame：字の前に描く、欄の下地と枠（/MK の BG・BC。mk_frame）。"""
    vw, vh = _view_box(rot, rect_w, rect_h)
    bg = f"q\n1 g\n0 0 {vw:.2f} {vh:.2f} re\nf\nQ\n".encode() if cover else b""
    if comb:
        if glyphs.shaped(str(value)):
            raise FillError(f"「{value}」は、字を組み立てて（シェーピング）描く文字なので、1 字ずつの枠（comb）の欄には書けません")
        chars = list(str(value))
        if len(chars) > len(comb):
            raise FillError(f"「{value}」は {len(chars)} 字です（この欄は {len(comb)} 字まで）")
        s = min(size, min(b[3] - b[1] for b in comb) * 0.7)
        ops = bytearray(frame + b"/Tx BMC\nq\n")
        x_origin = comb[0][0]
        y_top = min(b[1] for b in comb)
        for ch, b in zip(chars, comb):
            bw, bh = b[2] - b[0], b[3] - b[1]
            lx = (b[0] - x_origin) + max(0.0, (bw - text_width(ch, s)) / 2)
            # 見た目の座標（y 下向き）→ 欄の中の座標（y 上向き）
            ly = vh - ((b[1] - y_top) + bh / 2 + s * 0.35)
            ops += text_ops([ch], s, lx, ly, s)
        ops += b"Q\nEMC\n"
        return _ap(update, vw, vh, bytes(ops), fonts_used(chars), dr_fonts, _rot_matrix(rot, rect_w, rect_h)), s
    given = pad is not None
    pad = min(1.5, 0.1 * vw) if pad is None else min(pad, 0.45 * vw)
    # 余白を決めたとき（pad_mm）：その余白の内側に収める（入らなければ小さくするか、断る）。枠の端を字が越えないように
    lines, s = layout_text(str(value), vw - 2 * pad if given else vw, vh, size, center, multiline, unit)
    leading = s * 1.15
    block = leading * (len(lines) - 1) + s
    y = baseline if baseline is not None else (vh + block) / 2 - s * 0.8
    if baseline is not None:
        # 行の基線にそろえても、字が枠（BBox：描く所の切り抜き）を出ない高さに収める（字の箱は基線の下に 0.12、上に 0.88 字ぶん）
        y = max(min(y, vh - s * 0.88), s * 0.12 + leading * (len(lines) - 1))
    align = "center" if center else align
    aw = vw - 2 * pad
    ops = frame + bg + b"/Tx BMC\nq\n" + text_ops(lines, s, pad, y, leading, aw if align in ("center", "right") else None,
                                          align) + b"Q\nEMC\n"
    return _ap(update, vw, vh, ops, fonts_used(lines), dr_fonts, _rot_matrix(rot, rect_w, rect_h)), s


def _color_op(c, stroke: bool) -> str | None:
    """/MK の色（1 灰・3 RGB・4 CMYK。空は色なし）→ 色の命令。"""
    ops = {1: ("g", "G"), 3: ("rg", "RG"), 4: ("k", "K")}
    if not isinstance(c, list) or len(c) not in ops or not all(isinstance(v, (int, float)) for v in c):
        return None
    return " ".join(f"{float(v):g}" for v in c) + " " + ops[len(c)][1 if stroke else 0]


def mk_frame(doc, wd: dict, vw: float, vh: float, comb_cells: int = 0) -> bytes:
    """欄の下地と枠の描画：/MK の BG・BC と /BS の W（ISO 32000-1 12.5.6.19）。見た目を描き直すビューアーと同じものを、
    とどくの見た目にも描く（描かないと、書いた欄から枠や下地が消える）。comb_cells：1 字ずつの枠の数（仕切りも描く）。"""
    mk = doc.resolve(wd.get("MK"))
    if not isinstance(mk, dict):
        return b""
    bs = doc.resolve(wd.get("BS"))
    bw = doc.resolve(bs.get("W")) if isinstance(bs, dict) else None
    bw = float(bw) if isinstance(bw, (int, float)) and not isinstance(bw, bool) and bw >= 0 else 1.0
    out = ""
    bg = _color_op([doc.resolve(v) for v in doc.resolve(mk.get("BG")) or []], False)
    if bg:
        out += f"q\n{bg}\n0 0 {vw:.2f} {vh:.2f} re\nf\nQ\n"
    bc = _color_op([doc.resolve(v) for v in doc.resolve(mk.get("BC")) or []], True)
    if bc and bw > 0:
        out += f"q\n{bc}\n{bw:g} w\n{bw / 2:g} {bw / 2:g} {vw - bw:.2f} {vh - bw:.2f} re\nS\n"
        for i in range(1, comb_cells):
            x = vw * i / comb_cells
            out += f"{x:.2f} 0 m\n{x:.2f} {vh:.2f} l\nS\n"
        out += "Q\n"
    return out.encode()


def check_appearance(update: Update, w: float, h: float, dr_fonts: dict, char: str = "4", rot: int = 0,
                     said: str | None = None) -> Ref:
    """チェックの印の見た目。w・h は欄の Rect の幅と高さ、rot は欄の回転（/MK の R）：印も欄の向きに回して描く。
    said：読み手に伝える字（印の示す字）。"""
    vw, vh = _view_box(rot, w, h)
    s = min(vw, vh) * 0.8
    ops = f"q\nBT\n0 g\n/ZaDb {s:.2f} Tf\n{(vw - s * 0.85) / 2:.2f} {(vh - s * 0.75) / 2:.2f} Td\n({char}) Tj\nET\nQ\n".encode()
    fonts = {"ZaDb"}
    if said:
        # 印の示す字を、透明な字で印の枠の中に（見た目の中に置くので、印を消して書き直すと一緒に消える）
        ops += pagedraw.said_ops(text_ops, said, vw, vh)
        fonts |= fonts_used([said])
    return _ap(update, vw, vh, ops, fonts, dr_fonts, _rot_matrix(rot, w, h))


def empty_appearance(update: Update, w: float, h: float, rot: int = 0) -> Ref:
    vw, vh = _view_box(rot, w, h)
    return _ap(update, vw, vh, b"", set(), {}, _rot_matrix(rot, w, h))


# ====================================================================== 書き込み
def _wider_hint(spot: dict, view: View) -> str:
    """字が小さくなる・入らないときに添える、広く書く方法：行の所の上下の空き（事実）と box。"""
    x0, y0, x1, y1 = spot["rect"]
    mm = view.mm
    room = ((spot.get("facts") or {}).get("room_mm")) or {}
    roomy = f"。この行の上に {room.get('above', 0):g} mm・下に {room.get('below', 0):g} mm の空き" if room else ""
    return (f"。もっと広く書くなら、枠を box で決めてください（いまの枠は 左 {x0 / mm:.1f}・上 {y0 / mm:.1f}・右 {x1 / mm:.1f}・"
            f"下 {y1 / mm:.1f} mm{roomy}）")


def _as_text(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return "" if v is None else str(v)


def _as_bool(v, on=()) -> bool:
    """チェックの欄の値：true / false と、PDF が決めた「切」の状態の名前 Off（PDF 32000-1 12.7.4.2.3）と、
    その欄の「入」の状態の名前（on）だけ（inputcheck.checkbox_value。言葉の意味では決めない）。"""
    from ..inputcheck import InputError, checkbox_value

    try:
        return checkbox_value(v, tuple(on), ("Off",))
    except InputError as e:
        raise FillError(str(e)) from None


class Filler:
    def __init__(self, doc: PdfDocument):
        self.doc = doc
        self.up = Update(doc)
        self.fields = read_fields(doc)
        self.by_name = {f.name: f for f in self.fields}
        self.dr_fonts: dict[str, Ref] = {}
        self.new_fields: list[Ref] = []
        self.page_annots: dict[int, list[Ref]] = {}
        self.pages = doc.pages()
        self.patches: list[Patch] = []
        self.warnings: list[str] = []
        self.circles: dict[int, list[dict]] = {}  # ページ → ○で囲む所（circle.resolve）
        self.covers: set[str] = set()  # 覆って書いた所（cover）の名前：覆う字に重なってよい
        self.size_override: dict[str, float] = {}  # 一続きのマスにそろえた字の大きさ（_uniform_sizes）
        from .signatures import read as read_signatures

        self.sigs = read_signatures(doc)  # 電子署名：固定された欄・認証の署名が許す変更
        self.draw_on_page = pagedraw.may_draw(doc, self.sigs)  # 書いた字をページの描画に足すか（文書の事実で決める）
        self.drawn: dict[int, list] = {}  # ページ → ページに描く値（pagedraw.Item）
        self.drawn_labels: dict[str, str] = {}  # 欄の名前 → 窓の見出し（構造の木の、隣に入れる要素を探す）
        self.layouts: dict = {}  # ページの字の並び（チェックの印の左右の字）
        self.structure: list = []  # 構造の木に足した所（structadd.Placed）
        self.struct_why: list[str] = []  # 構造の木に足せなかった理由
        self.redraw_cut: list[dict] = []  # 描き直す表示で字の終わりが欠ける欄（広げられなかったもの）
        self.widened: dict[int, tuple] = {}  # 四角を広げた部品の番号 → (元の幅, 広げた幅)
        self.pdfa = is_pdfa(doc)  # PDF/A と名乗る文書か

    # ------------------------------------------------------------------ 共通の資源
    def font(self, key: str) -> Ref:
        if key not in self.dr_fonts:
            existing = self._existing_dr_font(key)
            if existing is not None:
                self.dr_fonts[key] = existing
                return existing
            if key == "Helv":
                self.dr_fonts[key] = self.up.new(dict(HELV))
            elif key == "TdkL":
                font, tu = glyphs.ext_font()
                font["ToUnicode"] = self.up.new(Stream({}, tu))
                self.dr_fonts[key] = self.up.new(font)
            elif key == "ZaDb":
                self.dr_fonts[key] = self.up.new(dict(ZADB))
            elif key == "TdkE" or key == BUNDLED_KEY or key.startswith("TdkX"):
                self.dr_fonts[key] = self.up.new({})  # 埋め込む書体：使った字形が決まってから、仕上げで中身を書く（_embed_font）
            else:
                ros = next(r for r, v in CJK.items() if v[0] == key)
                self.dr_fonts[key] = self.up.new(cjk_font(ros))
        return self.dr_fonts[key]

    def fonts_for(self, keys):
        return self.font

    def _existing_dr_font(self, key: str):
        if key not in ("Helv", "ZaDb"):
            return None
        af = self.doc.resolve(self.doc.root.get("AcroForm"))
        dr = self.doc.resolve(af.get("DR")) if isinstance(af, dict) else None
        fonts = self.doc.resolve(dr.get("Font")) if isinstance(dr, dict) else None
        ref = fonts.get(key) if isinstance(fonts, dict) else None
        if isinstance(ref, Ref):
            fd = self.doc.resolve(ref)
            base = getattr(fd.get("BaseFont"), "name", "") if isinstance(fd, dict) else ""
            enc = getattr(self.doc.resolve(fd.get("Encoding")), "name", "")
            if (base, enc) in (("Helvetica", "WinAnsiEncoding"), ("ZapfDingbats", "")):
                return ref  # 同じ書体・同じ符号化のときだけ（符号化が違うと、同じ文字コードで別の字形になる）
        return None

    # ------------------------------------------------------------------ フォームの欄
    def set_field(self, f: FormField, value, force: bool, display: str | None = None, calc: bool = False,
                  allow_hidden: bool = False) -> None:
        """display：見た目に描く字（XFA の静的なフォームで、値は決まった形・見た目は書式で整えた字のとき）。
        calc：計算の欄を、計算し直した値で書くとき（_recalculate だけが使う）。
        allow_hidden：画面にも紙にも出ない欄（/F の Hidden・NoView、Print なし）にも書く（本人に、出ないことを伝える）。"""
        if f.readonly and not force:
            raise FillError(f"{f.name}: 読み取り専用の欄です")
        k = f.kind
        if k == "signature":
            raise FillError(f"{f.name}: 電子署名の欄には、文字を書き込めません")
        if k == "button":
            raise FillError(f"{f.name}: 押すだけのボタンで、値を持ちません")
        sig = self.sigs.locked_by(f.name) if self.sigs is not None else None
        if sig is not None:
            raise FillError(f"{f.name}: 電子署名（{sig}）で固定された欄です（署名の /Lock・FieldMDP）。書くと署名が無効と表示される"
                            "ので、読み取り専用の欄として書きません")
        fd = self.up.get(f.ref) if f.ref is not None else None
        if fd is None:
            raise FillError(f"{f.name}: 欄のオブジェクトが見つかりません")
        from . import acrojs

        node = self.doc.resolve(f.ref)
        js = acrojs.actions(self.doc, node) if f.ref is not None else {}
        if "C" in js and not calc and display is None:
            srcs = acrojs.calc_sources(js["C"])
            raise FillError(f"{f.name}: 計算で決まる欄です（{('「' + '」「'.join(srcs) + '」から計算') if srcs else '計算のスクリプト'}）。"
                            "書いても、Acrobat などで開くと計算し直されて変わるので、書きません。元になる欄に書いてください")
        hidden_ok = calc or display is not None  # 計算し直した欄・XFA の欄（XFA の側で、隠した欄を確かめる）
        if k == "checkbox":
            states = self._check_states(f, fd)
            # true / false か、その欄の「入」の状態の名前・書き出しの値（/Ja など）・Off で指す（inputcheck.checkbox_value）
            on = _as_bool(value, [st for _w, st in states] + [o[0] for o in f.options])
            if on and not states:
                raise FillError(f"{f.name}: チェックの「入」の状態の名前がありません")
            if not hidden_ok:
                self._visible(f, None, allow_hidden)
            fd["V"] = Name(states[0][1]) if on else Name("Off")
            expect = []
            for w, st in states:
                wd = self.up.get(w.ref) if w.ref is not None and w.ref != f.ref else fd
                wd["AS"] = Name(st) if on else Name("Off")
                expect.append((w.ref.num if w.ref is not None else None, wd["AS"].name))
                self._ensure_state_ap(wd, w, st, check=True)
                self._redraw_state(wd, wd["AS"].name)
            p = Patch(f.name, "control", on)
            p.expect = {"V": fd["V"].name, "AS": expect}
            self.patches.append(p)
            return
        if k == "radio":
            states = self._radio_states(f, fd)
            if not states:
                raise FillError(f"{f.name}: 選択肢の名前が書類にありません（見た目（/AP）の状態も /Opt もなく、どれを選んだかを"
                                "書けません）")
            labels = {}
            for _w, st, lab in states:
                labels[st] = lab
            listed = "、".join(f"「{x}」" for x in labels.values())
            if isinstance(value, (list, tuple, dict, set)):
                raise FillError(f"{f.name}: ひとつ選ぶ欄です。選ぶ選択肢の名前を 1 つ書いてください。選べるのは {listed}")
            off = False
            if isinstance(value, bool) or value is None:
                if value is True or not f.radio_like:
                    # はい／いいえ（true / false）では、どの部品を入にするか決まらない（部品ごとに「入」の名前が違う）
                    raise FillError(f"{f.name}: ひとつ選ぶ欄（部品ごとに「入」の名前が違う欄）です。はい／いいえ（true / false）"
                                    f"では、どれを選ぶかが決まりません。選ぶ選択肢の名前で書いてください。選べるのは {listed}")
                off = True  # チェックの欄（Radio の印なし）：どれも選ばない（すべて Off）
            want = _as_text(value)
            if not off and want == "Off" and f.radio_like and "Off" not in labels.values():
                off = True
            pick = None
            if not off:
                hit = [st for st, lab in labels.items() if lab == want] or [st for st in labels if st == want]
                if len(set(hit)) > 1:
                    raise FillError(f"{f.name}: 「{want}」に当たる選択肢が {len(set(hit))} つあり、どれか決められません"
                                    f"（状態の名前 {'、'.join(hit)}）")
                pick = hit[0] if hit else None
                if pick is None:
                    raise FillError(f"{f.name}: 「{want}」は選択肢にありません。選べるのは {listed}")
                if not hidden_ok:
                    self._visible(f, [w for w, st, _l in states if st == pick], allow_hidden)
            fd["V"] = Name(pick) if pick is not None else Name("Off")
            expect = []
            for w, st, _lab in states:
                wd = self.up.get(w.ref) if w.ref is not None and w.ref != f.ref else fd
                wd["AS"] = Name(st) if st == pick else Name("Off")
                expect.append((w.ref.num if w.ref is not None else None, wd["AS"].name))
                self._ensure_state_ap(wd, w, st, check=f.radio_like)
            p = Patch(f.name, "control", labels[pick] if pick is not None else "Off")
            p.expect = {"V": fd["V"].name, "AS": expect}
            self.patches.append(p)
            return
        if k in ("combo", "list"):
            self._set_choice(f, fd, value, js, display, allow_hidden, hidden_ok)
            return
        if isinstance(value, (list, tuple, dict, set)):
            raise FillError(f"{f.name}: この欄には 1 つの値（字か数）を書きます。リストや {{…}} の形の値は書けません"
                            f"（受け取った値：{value!r}）")
        if not hidden_ok:
            self._visible(f, None, allow_hidden)
        text = _as_text(value)
        shown_fmt = None
        fmt_js = {kk: v for kk, v in js.items() if kk in ("F", "K", "V")}
        if fmt_js and display is None:
            if all(acrojs.known(v) for v in fmt_js.values()):
                try:
                    text, shown_fmt = acrojs.commit_pair(fmt_js, text)
                except acrojs.JsError as e:
                    raise FillError(f"{f.name}: {e}") from None
            else:
                self.warnings.append(f"{f.name}: 独自のスクリプト（書式・入力の制限）は実行していません")
        multiline = bool(f.flags & FF_MULTILINE)
        if "\n" in text or "\r" in text:
            if not multiline:
                raise FillError(f"{f.name}: 1行の欄です。改行は書けません（改行の代わりに空白や「、」で区切るなど、"
                                "1行にして書いてください）")
            text = text.replace("\r\n", "\n").replace("\r", "\n")
        if f.maxlen and len(text) > f.maxlen:
            raise FillError(f"{f.name}: 「{text}」は {len(text)} 字です（この欄は {f.maxlen} 字まで）")
        shown = display if display is not None else (shown_fmt if shown_fmt is not None else text)
        try:
            keys = fonts_used([shown]) if shown else set()
        except FillError as e:
            raise FillError(f"{f.name}: {e}") from None
        fd["V"] = text_string(text)
        da = self._da_of(f)
        da_size = _da_size(da)
        for w in f.widgets:
            wd = self.up.get(w.ref) if w.ref is not None and w.ref != f.ref else fd
            x0, y0, x1, y1 = w.rect
            rot = self._widget_rot(wd)
            unit = self._unit_of(w)
            vw, vh = _view_box(rot, x1 - x0, y1 - y0)  # 描く向きの幅と高さ（欄を 90・270 度回したときは入れ替わる）
            size = da_size or max(6.0 / unit, min(12.0 / unit, vh * 0.65))
            comb = None
            pinned = False
            if shown and not multiline and not (f.flags & FF_COMB) and pagedraw.widget_drawn(self.doc, wd) is None:
                # 作り手の欄（字の見た目を使わないビューアーでも欠けないように）：入らない字は小さくし、自動の大きさは決まった大きさに
                size, pinned = _fit_redrawn(shown, size, vw, vh, _da_auto(da), unit)
                try:
                    size = layout_text(shown, vw, vh, size, _q_center(f.q), False, unit)[1]  # 見た目に描く大きさ
                except FillError:
                    pass  # 入らないことは text_appearance が断る
                font = "" if all(ord(c) < 0x80 for c in shown) else "J"
                needs = pdfkit_needs(shown, size, font)
                if needs > vw and not self._widen_for_redraw(f, w, wd, rot, needs, shown):
                    self.redraw_cut.append({"id": f.name, "value": shown, "width_mm": round(vw * unit * 25.4 / 72, 1),
                                            "needs_mm": round(needs * unit * 25.4 / 72, 1)})
                    msg = (f"{f.name}: この欄は幅 {vw * unit * 25.4 / 72:.1f} mm で、Mac のプレビューや iPhone などの標準の表示（欄の見た目を使わず"
                           f"描き直す）には {needs * unit * 25.4 / 72:.1f} mm 要ります。そこでは「{shown}」の終わりが欠けて、違う値に見えます"
                           "（ファイルの中の値は正しい）。短く書けないか確かめ、書けなければ本人に伝えてください")
                    if msg not in self.warnings:
                        self.warnings.append(msg)
                x0, y0, x1, y1 = self._rect_of(w, wd)
                vw, vh = _view_box(rot, x1 - x0, y1 - y0)
            if f.flags & FF_COMB and f.maxlen:
                cw = vw / f.maxlen
                # 使うマスは書く字の分だけ（/MaxLen の宣言の数だけ作らない。超える字は text_appearance が断る）
                comb = [(i * cw, 0, (i + 1) * cw, vh) for i in range(min(f.maxlen, max(1, len(shown) + 1)))]
            try:
                # 1 行の欄（Multiline の印なし）は折り返さない：見た目を作り直すビューアーも 1 行で描くので、1 行に収まる大きさにする
                ap, s = text_appearance(self.up, x1 - x0, y1 - y0, shown, size, self.fonts_for(keys | {"Helv"}),
                                        comb=comb, rot=rot, multiline=multiline, unit=unit,
                                        align={1: "center", 2: "right"}.get(f.q, "left"),  # /Q（0 左・1 中央・2 右）
                                        frame=mk_frame(self.doc, wd, vw, vh, f.maxlen if comb else 0))
            except FillError as e:
                raise FillError(f"{f.name}: {e}" + ("" if multiline or comb else "（1 行の欄なので、折り返して書けません）")) from None
            drawn = pagedraw.widget_drawn(self.doc, wd)
            if drawn is not None:
                # とどくがページに描いた値を直す：ページが描く XObject の中身を入れ替える（元のストリームはそのまま）
                self.up.put(drawn, self.up.objs.pop(ap.num)[1])
                ap = drawn
            wd["AP"] = {"N": ap}
            self._set_da(fd, wd, da, s, keys, bool(shown), fixed=pinned)
            if shown and not calc:
                self._report_size(f.name, s, da_size, unit)
        p = Patch(f.name, "input", text)
        p.shown = shown
        self.patches.append(p)

    @staticmethod
    def _rect_of(w, wd: dict):
        """部品の四角（書く途中で広げたら、広げた後のもの）。"""
        r = wd.get("Rect")
        if isinstance(r, list) and len(r) == 4 and all(isinstance(v, (int, float)) for v in r):
            return min(r[0], r[2]), min(r[1], r[3]), max(r[0], r[2]), max(r[1], r[3])
        return w.rect

    def _widen_for_redraw(self, f: FormField, w, wd: dict, rot: int, needs: float, shown: str) -> bool:
        """/AP を使わず /V から描き直す表示（PDFKit）は、幅が足りないと字の終わりを描かない。見えない四角（枠も下地もない部品）なら、
        字の置かれる位置を変えず、横の空いている方へ四角だけを広げて、どの表示でも字が全部出るようにする。
        広げられるのは：回していない 1 行の欄で、ほかの部品に重ならず、ページの内側に収まるとき。広げた事実は warnings で伝える。"""
        mk = self.doc.resolve(wd.get("MK"))
        if rot or (isinstance(mk, dict) and any(self.doc.resolve(mk.get(k)) for k in ("BC", "BG"))):
            return False
        x0, y0, x1, y1 = w.rect
        grow = needs - (x1 - x0) + 0.01
        if f.q == 1:
            nx0, nx1 = x0 - grow / 2, x1 + grow / 2
        elif f.q == 2:
            nx0, nx1 = x0 - grow, x1
        else:
            nx0, nx1 = x0, x1 + grow
        pg = getattr(w, "page", None)
        if not (isinstance(pg, int) and 0 <= pg < len(self.pages)):
            return False
        mb = [self.doc.resolve(v) for v in (self.doc.resolve(self.pages[pg].get("CropBox") or self.pages[pg].get("MediaBox")) or [])]
        if len(mb) == 4 and all(isinstance(v, (int, float)) for v in mb) and (nx0 < min(mb[0], mb[2]) or nx1 > max(mb[0], mb[2])):
            return False
        for other in self.fields:
            for ow in other.widgets:
                if ow is w or getattr(ow, "page", None) != pg or not ow.rect:
                    continue
                a0, b0, a1, b1 = ow.rect
                if nx0 < a1 and a0 < nx1 and y0 < b1 and b0 < y1:
                    return False
        wd["Rect"] = [round(nx0, 3), y0, round(nx1, 3), y1]
        self.widened[w.ref.num if w.ref is not None else f.ref.num] = (round(x1 - x0, 3), round(nx1 - nx0, 3))
        self.warnings.append(f"{f.name}: 欄の四角（字の見えない部分）を横に {grow * 25.4 / 72:.1f} mm 広げました。「{shown}」の位置は変わらず、"
                             "見た目を使わず描き直す表示（Mac のプレビュー・iPhone など）でも字が欠けないようにするためです")
        return True

    def _set_choice(self, f: FormField, fd: dict, value, js: dict, display, allow_hidden: bool, hidden_ok: bool) -> None:
        """一覧から選ぶ欄（コンボ・リスト）：値（/V）・選んだ番号（/I、ISO 32000-1 12.7.4.4）・見た目を合わせて書く。
        リスト：選択肢を並べて描き、選んだ行に色を付ける。複数選べるリスト（MultiSelect）は、値をリストで渡せる。"""
        from . import acrojs

        k = f.kind
        opts = [o[1] or o[0] for o in f.options]
        exports = [o[0] for o in f.options]
        multi = k == "list" and bool(f.flags & FF_MULTISELECT)
        editable = k == "combo" and bool(f.flags & FF_EDIT)
        listed = "、".join(f"「{x}」" for x in opts[:30])
        if isinstance(value, (dict, set)):
            raise FillError(f"{f.name}: 選ぶ選択肢の名前で書いてください。選べるのは {listed}")
        if isinstance(value, (list, tuple)):
            if not multi:
                raise FillError(f"{f.name}: 1 つだけ選ぶ欄です。リストでは書けません。選ぶ選択肢の名前を 1 つ書いてください")
            wants = [_as_text(x) for x in value]
            if not wants:
                raise FillError(f"{f.name}: 選ぶ選択肢を 1 つ以上書いてください")
        else:
            wants = [_as_text(value)]
            fmt_js = {kk: v for kk, v in js.items() if kk in ("F", "K", "V")}
            if fmt_js and display is None:
                if all(acrojs.known(v) for v in fmt_js.values()):
                    try:
                        wants = [acrojs.commit_pair(fmt_js, wants[0])[0]]
                    except acrojs.JsError as e:
                        raise FillError(f"{f.name}: {e}") from None
                else:
                    self.warnings.append(f"{f.name}: 独自のスクリプト（書式・入力の制限）は実行していません")
        idx: list[int] = []
        custom = None
        for t in wants:
            hit = [i for i, o in enumerate(opts) if o == t] or [i for i, e in enumerate(exports) if e == t]
            if not hit:
                if editable and len(wants) == 1:
                    custom = t  # 字を書ける一覧（Edit）：選択肢にない字
                    continue
                raise FillError(f"{f.name}: 「{t}」は選択肢にありません。選べるのは {listed}")
            if len({(opts[i], exports[i]) for i in hit}) > 1:
                raise FillError(f"{f.name}: 「{t}」に当たる選択肢が {len(hit)} つあり（"
                                + "、".join(f"{i + 1} 番目「{opts[i]}」＝{exports[i]}" for i in hit)
                                + "）、どれか決められません。見せる字で指定してください")
            if hit[0] not in idx:
                idx.append(hit[0])
        idx.sort()
        if not hidden_ok:
            self._visible(f, None, allow_hidden)
        if custom is not None:
            if "\n" in custom or "\r" in custom:
                raise FillError(f"{f.name}: 1行の欄です。改行は書けません")
            vals, labels = [custom], [custom]
        else:
            vals, labels = [exports[i] for i in idx], [opts[i] for i in idx]
        shown = display if display is not None else "、".join(labels)
        try:
            keys = fonts_used([shown] + ([] if k == "combo" else opts)) if shown else set()
        except FillError as e:
            raise FillError(f"{f.name}: {e}") from None
        if multi and isinstance(value, (list, tuple)):
            fd["V"] = [text_string(v) for v in vals]
        else:
            fd["V"] = text_string(vals[0])
        if k == "list" or "I" in fd:
            if idx and custom is None:
                fd["I"] = list(idx)  # 選んだ選択肢の番号（0 から、小さい順）
            else:
                fd.pop("I", None)
        da = self._da_of(f)
        da_size = _da_size(da)
        ap_text = None
        for w in f.widgets:
            wd = self.up.get(w.ref) if w.ref is not None and w.ref != f.ref else fd
            x0, y0, x1, y1 = w.rect
            rot = self._widget_rot(wd)
            unit = self._unit_of(w)
            if k == "combo":
                size = da_size or max(6.0 / unit, min(12.0 / unit, _view_box(rot, x1 - x0, y1 - y0)[1] * 0.65))
                ap, s = text_appearance(self.up, x1 - x0, y1 - y0, shown, size, self.fonts_for(keys | {"Helv"}),
                                        multiline=False, align={1: "center", 2: "right"}.get(f.q, "left"), rot=rot,
                                        unit=unit)
            else:
                ti = self.doc.resolve(fd.get("TI"))
                ti = int(ti) if isinstance(ti, (int, float)) and not isinstance(ti, bool) and 0 <= ti < len(opts) else 0
                ap, s, top, visible = list_appearance(self.up, x1 - x0, y1 - y0, opts, set(idx), ti, da_size or 12.0 / unit,
                                                      self.fonts_for(keys | {"Helv"}), rot=rot, unit=unit)
                if top != ti:
                    fd["TI"] = top  # 選んだ行が見える所から並べる（一番上に見せる選択肢の番号）
                ap_text = "".join(opts[i] for i in visible)
            wd["AP"] = {"N": ap}
            self._set_da(fd, wd, da, s, keys, bool(shown))
            self._report_size(f.name, s, da_size, unit)
        p = Patch(f.name, "input", list(vals) if multi and isinstance(value, (list, tuple)) else vals[0])
        p.shown = shown
        p.ap_text = ap_text
        p.expect = {"I": list(idx) if (idx and custom is None and (k == "list" or "I" in fd)) else None}
        self.patches.append(p)

    def _da_of(self, f: FormField) -> str:
        """欄の書体の指定（/DA）：欄（親の欄を含む）になければ、フォーム（AcroForm）の既定の指定。"""
        if f.da:
            return f.da
        af = self.doc.resolve(self.doc.root.get("AcroForm"))
        v = self.doc.resolve(af.get("DA")) if isinstance(af, dict) else None
        return bytes(v).decode("latin-1") if isinstance(v, (bytes, PdfString)) else ""

    def _redraw_state(self, wd: dict, state: str) -> None:
        """とどくがページに描いたチェックの印：描く XObject の中身を、いまの状態の見た目にする。"""
        drawn = pagedraw.widget_drawn(self.doc, wd)
        if drawn is None:
            return
        ap = self.doc.resolve(wd.get("AP"))
        src = self.doc.resolve(self.doc.resolve(ap.get("N")).get(state))
        self.up.put(drawn, Stream({k: v for k, v in src.dict.items() if k != "Length"}, self.doc.decode(src)))

    def _set_da(self, fd: dict, wd: dict, da: str, s: float, keys: set, drawn: bool, fixed: bool = False) -> None:
        """見た目を作り直すビューアー（NeedAppearances・欄を選んだとき）が、描いた大きさと同じ大きさで描くように、/DA の字の
        大きさを、見た目に描いた大きさにする。元が自動（0）なら 0 のまま（ビューアーが欄に収まる大きさにする）。"""
        orig = _da_size(da)
        m = re.search(r"([\d.]+)\s+Tf", da or "")
        auto = m is not None and _num(m[1]) == 0 and not fixed
        if keys - {"Helv"}:
            # 日本語などは Helvetica では描けないので、欄の書体の指定も合わせる
            # （色は元の指定のまま。Tf の命令がなければ黒で）
            name = self._da_name(_da_key(keys))
            size_text = "0" if auto else _pt(s)
            new = _da_refont(da, name)
            new = re.sub(r"([\d.]+)(\s+Tf)", lambda m: size_text + m[2], new, count=1) if new else f"/{name} {size_text} Tf 0 g"
            wd["DA"] = PdfString(new.encode("latin-1"))
            fd["DA"] = wd["DA"]
            return
        if not drawn:
            return
        # 見た目を作り直すビューアーは /DA の書体で値を描く。元の書体が埋め込まれていない・部分集合で字の割り当てが違うと、
        # 値は正しいのに別の字や×の四角に見える（とどくが描いた字とも違う）。描いた書体（Helvetica）に合わせる（色はそのまま）
        new = _da_refont(da, "Helv")
        if not (auto or ((orig is None or abs(orig - s) < 0.005) and not fixed)):
            new = _da_resize(new or da, s)
        if new is not None and new != da:
            wd["DA"] = PdfString(new.encode("latin-1"))

    def _widget_rot(self, wd) -> int:
        """欄の部品の回転（/MK の R：反時計回りの角度。ISO 32000-1 表 189）。"""
        mk = self.doc.resolve(wd.get("MK")) if isinstance(wd, dict) else None
        return norm_rot(self.doc.resolve(mk.get("R"))) if isinstance(mk, dict) else 0

    def _unit_of(self, w) -> float:
        """欄の部品があるページの UserUnit（分からなければ 1）。"""
        from .layout import user_unit

        pg = getattr(w, "page", None)
        return user_unit(self.doc, self.pages[pg]) if isinstance(pg, int) and 0 <= pg < len(self.pages) else 1.0

    def _report_size(self, name: str, s: float, da_size: float | None, unit: float = 1.0) -> None:
        """知らせる字の大きさは、紙の上の pt（UserUnit のあるページは、ページの座標の大きさ × UserUnit）。"""
        msg = None
        if s * unit < MIN_READABLE_PT:
            msg = (f"{name}: 欄に収めるため、字が {s * unit:.1f} pt になり、印刷すると読みにくい大きさです"
                   f"（{MIN_READABLE_PT:g} pt 未満）。短く書けないか確かめてください")
        elif da_size and s < da_size - 0.005:
            msg = (f"{name}: 欄に収めるため、字を {_pt(s * unit)} pt に小さくしました（書類の決めた大きさ {_pt(da_size * unit)} pt）。"
                   "欄の書体の指定（/DA）も同じ大きさにしました")
        if msg and msg not in self.warnings:
            self.warnings.append(msg)

    def _visible(self, f: FormField, widgets, allow_hidden: bool) -> None:
        """値を見せる部品が、画面にも紙にも出ない（隠した欄）なら断る（allow_hidden のときは、書いて知らせる）。
        画面には出るが紙に印刷されない欄は、書いて知らせる。"""
        gone, unprinted = f.visibility(widgets)
        if gone:
            if not allow_hidden:
                raise FillError(f"{f.name}: この欄は、画面にも紙にも出ない欄です（{gone}）。書いても本人にも受け取る人にも見えず、"
                                "紙には印刷されないので、書きません（ほかの欄を選ぶとスクリプトで出てくる欄のこともあります）。"
                                '隠れた欄と分かったうえで書くときは {"value": 値, "allow_hidden": true} で書いてください')
            self.warnings.append(f"{f.name}: 画面にも紙にも出ない欄（{gone}）に書きました。値はファイルに入りますが、"
                                 "書いた値は紙に印刷されず、画面にも出ません")
        elif unprinted:
            self.warnings.append(f"{f.name}: 画面には出ますが、紙に印刷されない欄です（{unprinted}）。印刷して出すときは、"
                                 "書いた値が紙に出ません")
        elif not any(w.on_screen for w in (f.widgets if widgets is None else widgets)):
            self.warnings.append(f"{f.name}: 紙には印刷されますが、画面には出ない欄です（画面に出さない印（/F の NoView））。"
                                 "PDF のまま出すときは、書いた値が画面で見えません")

    def _check_states(self, f: FormField, fd: dict) -> list:
        """チェックの部品ごとの「入」の状態の名前。見た目（/AP /N）にあればそれ。なければ /Opt の書き出しの値か、
        仕様の勧める Yes（ISO 32000-1 12.7.4.2.3）。"""
        opt = [o[0] for o in f.options]
        out = []
        for i, w in enumerate(f.widgets):
            st = w.on or (opt[i] if i < len(opt) and opt[i] else None) or "Yes"
            out.append((w, st))
        return out

    def _radio_states(self, f: FormField, fd: dict) -> list:
        """ラジオの部品ごとの (部品, 状態の名前, 見せる名前)。見た目の状態がない部品は、/Opt の順（状態の名前は番号）で決める。"""
        from .form import name_text

        out = []
        for i, w in enumerate(f.widgets):
            st = w.on
            if st is None:
                if i < len(f.options):
                    st = str(i)
                else:
                    continue
            lab = name_text(st)
            if f.options and st.isdigit() and int(st) < len(f.options):
                lab = f.options[int(st)][1] or f.options[int(st)][0]
            out.append((w, st, lab))
        return out

    def _ensure_state_ap(self, wd: dict, w, st: str, check: bool) -> None:
        """見た目（/AP）のない部品（/NeedAppearances の書類）：入と切の見た目を足す（どのビューアーでも同じに見えるように）。"""
        if w.on is not None:
            return
        x0, y0, x1, y1 = w.rect
        bw, bh = x1 - x0, y1 - y0
        rot = self._widget_rot(wd)  # 回した欄（/MK の R）は、印も欄の向きに回す
        mk = self.doc.resolve(wd.get("MK")) if isinstance(wd, dict) else None
        ca = self.doc.resolve(mk.get("CA")) if isinstance(mk, dict) else None
        char = bytes(ca).decode("latin-1")[:1] if isinstance(ca, (bytes, PdfString)) and ca else ""
        if not char.isalnum():
            char = "4" if check else "l"  # ZapfDingbats：4 はチェックの印、l は黒い丸
        on = check_appearance(self.up, bw, bh, self.fonts_for({"ZaDb"}), char, rot)
        off = empty_appearance(self.up, bw, bh, rot)
        wd["AP"] = {"N": {st: on, "Off": off}}

    # ------------------------------------------------------------------ 配置から見つけた欄
    def add_widget(self, page_no: int, spot: dict, value, view: View) -> None:
        """印刷用のPDFの欄に、同じ名前のフォームの欄を足す。"""
        page_ref = self.pages[page_no - 1].get("_ref")
        if page_ref is None:
            raise FillError(f"{spot['id']}: ページのオブジェクトが見つかりません")
        k = spot["kind"]
        vx0, vy0, vx1, vy1 = spot["rect"]
        x0, y0, x1, y1 = view.back_rect(vx0, vy0, vx1, vy1)
        w, h = x1 - x0, y1 - y0
        rot = view.rot
        self.drawn_labels[spot["id"]] = spot.get("label") or ""
        widget = {"Type": Name("Annot"), "Subtype": Name("Widget"), "T": text_string(spot["id"]), "F": 4, "P": page_ref}
        # 欄の説明（/TU）：読み手に読まれる字なので、窓の空きの印（［　］）と長い空白を除く（窓のための、位置の印つきの見出しは /TdkLabel に）
        label = spot.get("label") or ""
        tu = pagedraw.clean_label(label)
        if tu:
            widget["TU"] = text_string(tu)
        if label and label != tu:
            widget[pagedraw.LABEL] = text_string(label)
        if k == "checkbox":
            on = _as_bool(value)
            said = None
            if self.draw_on_page:
                # 印が示す字を、読み手に伝える（印の見た目の中に。書き直して印を消すと、説明も消える）
                said = self._check_said(page_no, spot, on)
            yes = check_appearance(self.up, w, h, self.fonts_for({"ZaDb"} | (fonts_used([said]) if said else set())), rot=rot, said=said)  # 回したページでは、印も見た目の向きに
            off = empty_appearance(self.up, w, h, rot)
            d = {**widget, "FT": Name("Btn"), "Rect": [round(x0, 3), round(y0, 3), round(x1, 3), round(y1, 3)],
                 "V": Name("Yes") if on else Name("Off"), "AS": Name("Yes") if on else Name("Off"),
                 "AP": {"N": {"Yes": yes, "Off": off}}, "MK": {"CA": PdfString(b"4"), **({"R": rot} if rot else {})},
                 "DA": PdfString(b"/ZaDb 0 Tf 0 g")}
            self._place(page_no, d, yes if on else off, spot["id"], state=True)
            self.patches.append(Patch(spot["id"], "control", on, page_no))
            return
        if isinstance(value, (list, tuple, dict, set)):
            # PDF の配置の所は、どれも 1 つの空き（1 つの値）：リストの字（「['本人']」）をそのまま書かない
            raise FillError(f"{spot['id']}: この所には 1 つの値（字か数）を書きます。リストや {{…}} の形の値は書けません"
                            f"（受け取った値：{value!r}）。空きが並ぶ所は、それぞれの id に 1 つずつ書いてください")
        text = _as_text(value)
        size = float(self.size_override.get(spot["id"]) or spot.get("size") or 10.5 / view.unit)
        comb = None
        if k == "comb":
            boxes = spot.get("box_rects") or []
            comb = []
            for b in boxes:
                bx0, by0, bx1, by1 = b
                comb.append((bx0 - vx0, by0 - vy0, bx1 - vx0, by1 - vy0))
        cover = bool(spot.get("cover"))
        unit = view.unit  # 知らせる字の大きさは、紙の上の pt（UserUnit のあるページも）
        # そろえ方：エージェントが決める（align）。既定は左（枠の左の端から、小さな決まった余白）。字の意味からは決めない
        align = spot.get("align") or "left"
        pad = spot.get("pad_mm")
        pad_pt = pad * view.mm if pad is not None else None
        multiline = spot_multiline(spot, vy1 - vy0, size, unit)
        base = spot.get("baseline")
        try:
            ap, s = text_appearance(self.up, w, h, text, size, self.fonts_for(fonts_used([text]) | {"Helv"}),
                                    align=align, comb=comb, rot=rot, multiline=multiline, cover=cover, unit=unit, pad=pad_pt,
                                    baseline=vy1 - base if base is not None and not comb else None)
        except FillError as e:
            raise FillError(f"{spot['id']}: {e}{_wider_hint(spot, view)}") from None
        if cover:
            self.covers.add(spot["id"])
            self.warnings.append(f"{spot['id']}: 元の字「{spot.get('covered', '')}」は、白い下地の欄で覆いました（元の字は消していません）")
        if text and s * unit < MIN_READABLE_PT:
            self.warnings.append(f"{spot['id']}: 欄に収めるため、字が {s * unit:.1f} pt になり、印刷すると読みにくい大きさです"
                                 f"（{MIN_READABLE_PT:g} pt 未満）。短く書けないか確かめてください{_wider_hint(spot, view)}")
        elif text and s < 0.6 * size:
            self.warnings.append(f"{spot['id']}: 欄に収めるため、字を {s * unit:.1f} pt に小さくしました（元の大きさ {size * unit:.1f} pt）")
        inner_w = (vx1 - vx0) - 2 * min(pad_pt, 0.45 * (vx1 - vx0)) if pad_pt is not None else vx1 - vx0
        if multiline and not comb:
            wrapped = len(layout_text(text, inner_w, vy1 - vy0, size, False, True, unit)[0])
            if wrapped > text.count("\n") + 1:
                self.warnings.append(f"{spot['id']}: 枠の幅に入らないので {wrapped} 行に折り返しました（字 {s * unit:.1f} pt）。"
                                     "折り返したくなければ、もっと幅のある枠を box で決めてください")
        keys = fonts_used([text]) or {"Helv"}
        da_font = self._da_name(_da_key(keys))
        d = {**widget, "FT": Name("Tx"), "Rect": [round(x0, 3), round(y0, 3), round(x1, 3), round(y1, 3)],
             "V": text_string(text), "DA": PdfString(f"/{da_font} {s:.1f} Tf 0 g".encode()), "AP": {"N": ap}}
        if comb:
            d["Ff"] = FF_COMB
            d["MaxLen"] = len(comb)
        elif "\n" in text or (multiline and len(layout_text(text, inner_w, vy1 - vy0, size, False, True, unit)[0]) > 1):
            d["Ff"] = FF_MULTILINE  # 複数行に描いた欄：見た目を作り直すソフトで 1 行に潰れないように（読み直して書き直すときも）
        if align != "left":
            d["Q"] = {"center": 1, "right": 2}[align]  # 見た目を作り直すソフトも、同じそろえ方で
        if rot:
            d["MK"] = {"R": rot}
        if cover:
            d.setdefault("MK", {})["BG"] = [1]  # 見た目を作り直すソフトでも、下地は白
        self._place(page_no, d, ap, spot["id"], state=False)
        p_ = Patch(spot["id"], "input", text, page_no, shown=text)
        # 書いた枠（ページの左上からの mm）・字の大きさ・そろえ方：エージェントが view_page で見て確かめ、違えば box・align で書き直す
        p_.rect_mm = {"page": page_no, "x": round(vx0 / view.mm, 1), "y": round(vy0 / view.mm, 1),
                      "w": round((vx1 - vx0) / view.mm, 1), "h": round((vy1 - vy0) / view.mm, 1),
                      "size_pt": round(s * unit, 1), "align": "comb" if comb else align,
                      "box_mm": [round(v / view.mm, 1) for v in (vx0, vy0, vx1, vy1)]}  # view_page の marks にそのまま渡せる
        self.patches.append(p_)

    def _check_said(self, page_no: int, spot: dict, on: bool) -> str | None:
        """チェックの印の説明。ページの字の位置だけから（印の同じ行の左右の字）。字がなければ何も言わず、知らせる。"""
        from . import circle as _circle

        lay = _circle.layout(self.doc, page_no, self.layouts)
        said = _circle.check_said(lay, spot["rect"])
        if said is not None:
            try:
                runs_for(said)
            except FillError:
                said = None  # 描ける書体がない字
        if said is None and on:
            self.warnings.append(f"{spot['id']}: チェックの印と同じ行に読める字がないので、どの選択肢の印かは、読み上げでは伝わりません（ページのこの位置を伝えてください）")
        return said

    def _said_form(self, view: View, sd: dict, key: str) -> Ref | None:
        """○が囲んだ字を、透明な字で描くフォーム。○の枠の中に。描ける書体がなければ None。"""
        from . import circle as _circle

        text = _circle.description(sd)
        x0, y0, x1, y1 = view.back_rect(*sd["rect"])
        w, h = x1 - x0, y1 - y0
        vw, vh = _view_box(view.rot, w, h)
        try:
            ops = pagedraw.said_ops(text_ops, text, vw, vh)
            ref = _ap(self.up, vw, vh, ops, fonts_used([text]), self.fonts_for(fonts_used([text])), _rot_matrix(view.rot, w, h))
        except FillError:
            sd["said_failed"] = True
            self.warnings.append(f"{sd['id']}: 囲んだ字を描ける書体がないので、読み上げでは、どの字を囲んだかが伝わりません")
            return None
        self.up.objs[ref.num][1].dict[pagedraw.SAID_KEY] = True
        return ref

    def _place(self, page_no: int, d: dict, ap: Ref, name: str, state: bool) -> None:
        """足した所を、文書に置く。ページに描くとき：字をページの描画に足し、値の控えは /Fields にだけ置く（注釈にしない。pagedraw）。
        state：チェックの印（描くものを状態の見た目の写しにして、直して書き直すときに中身だけを入れ替える）。"""
        if not self.draw_on_page:
            if not state and QUARTZ_NOTE not in self.warnings:
                self.warnings.append(QUARTZ_NOTE)
            self._attach(page_no, self.up.new(d))
            return
        drawn = ap
        if state:
            src = self.up.objs[ap.num][1]
            drawn = self.up.new(Stream(dict(src.dict), src.raw))
        d["F"] = pagedraw.ANNOT_HIDDEN
        d[pagedraw.KEY] = drawn
        self.new_fields.append(self.up.new(d))
        r = d["Rect"]
        self.drawn.setdefault(page_no, []).append(pagedraw.Item(name, drawn, r[0], r[1], r[2], r[3]))

    def _draw_page_values(self) -> None:
        """○と書いた字：ページの描画の前に「q」、後ろに描くストリームを足す（元のストリームは変えない）。
        タグのある文書では、足した字を構造の木にも足す（structadd）。"""
        from . import circle as _circle, structadd

        adder = None
        for page_no in sorted(set(self.circles) | set(self.drawn)):
            page = self.pages[page_no - 1]
            pref = page["_ref"]
            view = View(self.doc, page)
            depth = _circle.open_depth(self.doc.page_content(page))
            pd = self.up.get(pref)
            items = self.drawn.get(page_no, [])
            res = dict(self.doc.resolve(page.get("Resources")) or {})
            xo = dict(self.doc.resolve(res.get("XObject")) or {})
            first, why = structadd.first_mcid(self.doc, page) if self.draw_on_page else (None, None)
            if why and why not in self.struct_why:
                self.struct_why.append(why)
            circles = self.circles.get(page_no, [])
            data, names, seq = pagedraw.page_ops(view, depth, circles, items, xo, first, self.draw_on_page,
                                                 lambda sd, key: self._said_form(view, sd, key))
            if self.draw_on_page:
                for sd in circles:
                    if _circle.description(sd) is None:
                        self.warnings.append(f"{sd['id']}: 囲んだ字を読めないので、読み上げでは、どの字を囲んだかが伝わりません")
            if seq:
                labels = {it.name: self.drawn_labels.get(it.name, "") for it in items}
                for e in seq:
                    e["label"] = labels.get(e["id"]) or next((sd.get("text") for sd in circles if sd["id"] == e["id"]), "")
                adder = adder or structadd.Adder(self.up, self.doc)
                adder.add_page(page_no, page, seq)
            if names:
                xo.update(names)
                res["XObject"] = xo
                pd["Resources"] = res
            pre = self.up.new(Stream({}, _circle.PREFIX))
            post = self.up.new(Stream({}, data))
            cur = self.doc.resolve(pd.get("Contents"))
            if isinstance(cur, list):
                old = list(cur)
            elif pd.get("Contents") is None:
                old = []
            else:
                old = [pd.get("Contents")]
            pd["Contents"] = [pre] + old + [post]
        if adder is not None:
            adder.finish()
            self.structure = adder.placed
            msg = structadd.note(adder.placed)
            if msg:
                self.warnings.append(msg)
        for why in self.struct_why:
            self.warnings.append(f"この PDF のタグ（文書の構造の木）には、書いた字を足せませんでした（{why}）。字はページに描いてあり、値は読めますが、"
                                 "木を読む読み手（Mac のプレビューの読み上げなど）には読まれません")

    def _attach(self, page_no: int, ref: Ref, field: bool = True) -> None:
        self.page_annots.setdefault(page_no, []).append(ref)
        if field:
            self.new_fields.append(ref)

    # ------------------------------------------------------------------ 仕上げ
    def _used_gids(self, key: str) -> list[int]:
        """書いた見た目のストリームで、書体 key で描いた字形（Tj・TJ の 2 バイトずつ）。"""
        from .objects import Lexer

        out: list[int] = []
        tag = f"/{key} ".encode()
        for _num, (_gen, o) in self.up.objs.items():
            if not isinstance(o, Stream) or tag not in o.raw:
                continue
            lx = Lexer(o.raw, 0)
            cur = None
            stack: list = []
            while True:
                kind, val = lx.token()
                if kind == "eof":
                    break
                if kind == "kw":
                    if val == "Tf" and len(stack) >= 2 and isinstance(stack[-2], Name):
                        cur = stack[-2].name
                    elif val in ("Tj", "TJ") and cur == key:
                        for x in (stack[-1] if val == "TJ" and stack and isinstance(stack[-1], list) else stack[-1:]):
                            if isinstance(x, (bytes, PdfString)):
                                b = bytes(x)
                                out += [int.from_bytes(b[i:i + 2], "big") for i in range(0, len(b) - 1, 2)]
                    stack = []
                elif kind == "[":
                    arr = []
                    while True:
                        k2, v2 = lx.token()
                        if k2 in ("]", "eof"):
                            break
                        arr.append(v2)
                    stack.append(arr)
                elif kind == "<<":
                    depth = 1
                    while depth:
                        k2, _v = lx.token()
                        if k2 == "eof":
                            break
                        depth += {"<<": 1, ">>": -1}.get(k2, 0)
                else:
                    stack.append(val)
        return out

    def _embed_bundled(self) -> None:
        """TdkB：書いた見た目が使う字（CID）だけを、同梱の書体から埋め込む。PDF/A の文書には CIDSet も（PDF/A-1 の決まり）。"""
        from .embedfont import bundled_objects, compress

        t0, cid, desc, data, tu, c2g, cidset = bundled_objects(sorted(set(self._used_gids(BUNDLED_KEY))))
        flate = {"Filter": Name("FlateDecode")}
        desc["FontFile2"] = self.up.new(Stream({"Length1": len(data), **flate}, compress(data)))
        if self.pdfa:
            desc["CIDSet"] = self.up.new(Stream(dict(flate), compress(cidset)))
        cid["FontDescriptor"] = self.up.new(desc)
        cid["CIDToGIDMap"] = self.up.new(Stream(dict(flate), compress(c2g)))
        t0["DescendantFonts"] = [self.up.new(cid)]
        t0["ToUnicode"] = self.up.new(Stream(dict(flate), compress(tu)))
        self.up.put(self.dr_fonts[BUNDLED_KEY], t0)

    def _da_name(self, key: str) -> str:
        """/DA と AcroForm の /DR で、書体 key を指す名前。埋め込む書体は、この書き込みで使った字だけの部分集合なので、前の書き込みが
        同じ名前で /DR に置いた書体（別の部分集合）とは別の名前にする（元の書体は変えない。前に書いた欄の /DA も変えない）。"""
        if key not in (BUNDLED_KEY, "TdkE"):
            return key
        af = self.doc.resolve(self.doc.root.get("AcroForm"))
        dr = self.doc.resolve(af.get("DR")) if isinstance(af, dict) else None
        fonts = self.doc.resolve(dr.get("Font")) if isinstance(dr, dict) else None
        fonts = fonts if isinstance(fonts, dict) else {}
        name, n = key, 1
        while name in fonts:
            n += 1
            name = f"{key}{n}"
        return name

    def _embed_font(self, key: str = "TdkE") -> None:
        """TdkE：書いた見た目が使う字形（GID）だけを、このパソコンの TrueType の書体から埋め込む。
        TdkX…：字を組み立てて描いた書体（shape.py）。ToUnicode は、字形が表す元の字（合字なら合わせた字）にする
        （行の元の字は、見た目の ActualText にある）。"""
        from .embedfont import compress, font_objects, system_font

        shaped = key.startswith("TdkX")
        if shaped:
            from .shape import font_by_key

            sf = font_by_key(key)
            font = sf.tt
        else:
            font = system_font()
        ref = self.dr_fonts[key]
        inv: dict[int, str] = {}
        for u, g in sorted(font.cmap.items()):
            inv.setdefault(g, chr(u))
        used: dict[int, str] = {}
        for g in self._used_gids(key):
            used[g] = (sf.glyph_chars.get(g) if shaped else None) or inv.get(g, "")
        t0, cid, desc, data, tu = font_objects(font, used)
        desc["FontFile2"] = self.up.new(Stream({"Length1": len(data), "Filter": Name("FlateDecode")}, compress(data)))
        cid["FontDescriptor"] = self.up.new(desc)
        t0["DescendantFonts"] = [self.up.new(cid)]
        t0["ToUnicode"] = self.up.new(Stream({}, tu))
        self.up.put(ref, t0)
        if shaped:
            self.warnings.append(f"{sf.script.ja or sf.script.name}は、字の組み立て（シェーピング）をして、このパソコンの書体 "
                                 f"{font.name}（{font.path}）から、使った字形だけを PDF に埋め込みました。字の並びを、とどくの字の組み立てで確かめました"
                                 "（見た目そのものは確かめていません。描いた字形がいまの組み立てと同じかを、とどく自身の組み立て・"
                                 "ActualText と照らしただけです）。書いた欄は view_page でページの画像を見て、字の形が正しいかを確かめる")
            return
        names = "".join(sorted(ch for ch in used.values() if ch))
        self.warnings.append(f"「{names[:40]}」の字は、標準の書体に字形がないので、このパソコンの書体 {font.name}（{font.path}）から、"
                             "書いた字の字形だけを PDF に埋め込みました")

    def finish(self) -> bytes:
        doc = self.doc
        up = self.up
        # 先にページの描画を作る：○が示す透明な字（_said_form）も埋め込む書体を使うので、書体の中身はそのあとで書く
        if self.circles or self.drawn:
            self._draw_page_values()
        for k in list(self.dr_fonts):
            if k == BUNDLED_KEY:
                self._embed_bundled()
            elif k == "TdkE" or k.startswith("TdkX"):
                self._embed_font(k)
        for page_no, refs in self.page_annots.items():
            pref = self.pages[page_no - 1]["_ref"]
            pd = up.get(pref)
            annots = pd.get("Annots")
            if isinstance(annots, Ref):
                arr = up.get(annots)
                arr.extend(refs)
            else:
                pd["Annots"] = list(annots or []) + refs
        clear_na = self.clear_need_appearances()
        if self.new_fields or self.dr_fonts or clear_na:
            root_ref = doc.trailer["Root"]
            af_ref = doc.root.get("AcroForm")
            if isinstance(af_ref, Ref):
                af = up.get(af_ref)
            elif isinstance(af_ref, dict):
                cat = up.get(root_ref)
                af = cat["AcroForm"]
            else:
                af = {"Fields": [], "DA": PdfString(f"/{self._default_da_font()} 0 Tf 0 g".encode())}
                cat = up.get(root_ref)
                cat["AcroForm"] = up.new(af)
            if self.new_fields:
                fl = af.get("Fields")
                if isinstance(fl, Ref):
                    up.get(fl).extend(self.new_fields)
                else:
                    af["Fields"] = list(fl or []) + self.new_fields
            if self.dr_fonts:
                dr = af.get("DR")
                if isinstance(dr, Ref):
                    dr = up.get(dr)
                elif dr is None:
                    dr = {}
                    af["DR"] = dr
                fonts = dr.get("Font")
                if isinstance(fonts, Ref):
                    fonts = up.get(fonts)
                elif fonts is None:
                    fonts = {}
                    dr["Font"] = fonts
                for k, r in self.dr_fonts.items():
                    fonts.setdefault(self._da_name(k), r)
            if "DA" not in af:
                af["DA"] = PdfString(f"/{self._default_da_font()} 0 Tf 0 g".encode())
            if clear_na:
                af["NeedAppearances"] = False
        return up.build()

    def _default_da_font(self) -> str:
        """フォームの既定の /DA の書体：/DR に置いた書体（描いていない Helvetica を指さない）。"""
        return self._da_name(_da_key(set(self.dr_fonts) - {"ZaDb"}))

    def clear_need_appearances(self) -> bool:
        """NeedAppearances を false にしてよいか：とどくが描いて確かめた見た目のまま見せる。書類のほかの部品がみな見た目（/AP）を
        持つときだけ（作り直しに頼る部品があれば、そのまま残して知らせる）。電子署名のある文書と XFA は変えない
        （権限の設定は変えてよい：ページの描画ではなく、書いた欄の見せ方の指定なので）。"""
        if (not self.patches or self.sigs.signed or self.sigs.docmdp is not None or not _need_appearances(self.doc)
                or _xfa_form(self.doc) is not None):
            return False
        written = {p.id for p in self.patches}
        for f in self.fields:
            if f.name in written:
                continue
            for w in f.widgets:
                wd = self.doc.resolve(w.ref) if w.ref is not None else self.doc.resolve(f.ref)
                ap = self.doc.resolve(wd.get("AP")) if isinstance(wd, dict) else None
                if not (isinstance(ap, dict) and self.doc.resolve(ap.get("N")) is not None):
                    return False
        return True


# ページの描画に足せない文書（電子署名のある PDF など：_may_draw_on_page）では、書いた字はフォームの欄（値と見た目）として足す。
# Apple の PDF の描画（Mac のプレビュー・iPhone などの標準の表示）は見た目を使わず値から描き直すので、狭い欄では字の終わりが欠けて見える
QUARTZ_NOTE = ("この PDF は電子署名か権限の設定があるので、ページの描画は変えず、書いた字をフォームの欄として足しました。"
               "Mac のプレビューや iPhone などの標準の表示では、欄の見た目を使わず描き直すので、狭い欄の字の終わりが欠けて見えることが"
               "あります（ファイルの中の値は正しい）。確かめや印刷は、欄の見た目を使う表示（Acrobat など）で行ってください")


# 字の見た目を使わず /V から描き直すビューアー（Apple の PDFKit）で、1 行の文字の欄が 1 行で描かれる幅（Pt）：
# 字の幅 × 1.2 ＋ 4（横の余白）。入らないと語の切れ目で折り返し、2 行目は描かれない。tests/audit_findings/pdfkit/ladder.py で測った、
# 字と大きさを変えた欄のどれも、この幅なら欠けない値（書類の集まりで合わせた値ではない）
REDRAW_SCALE, REDRAW_PAD = 1.2, 4.0
# 短い値は、比より決まった余白で欠ける：「01」は 6.5pt でも 9pt でも、字の幅 ＋ 7pt の欄で初めて 2 字とも出た。同じ所で測った、短い値の下限
REDRAW_MIN_ROOM = 8.0
REDRAW_AUTO_HEIGHT = 0.8  # 自動の大きさ（/DA の 0）で PDFKit が選ぶ字の大きさの、欄の高さに対する比（同じ所で測った）


def pdfkit_needs(text: str, size: float, da_font: str = "") -> float:
    """PDFKit がこの字を 1 行で描くのに要る欄の幅。PDFKit は /DA の埋め込んだ書体を使わず、欧文を Helvetica で描く（/DA が Helvetica でも
    同梱の書体でも同じ幅に描いた。ladder.py と同じ形で測った）ので、欧文は Helvetica の幅で。そのほかは、和文の書体に置き換わるときの
    幅の上限（どの字も 1 字ぶんまで）。"""
    if all(ord(c) < 0x80 for c in text) and da_font in ("", "Helv"):
        w = sum(glyphs.helv_width(c) for c in text) * size / 1000.0
    else:
        w = len(text) * size
    return max(w * REDRAW_SCALE + REDRAW_PAD, w + REDRAW_MIN_ROOM)


def _q_center(q) -> bool:
    return q == 1


def _da_auto(da: str) -> bool:
    m = re.search(r"([\d.]+)\s+Tf", da or "")
    return m is not None and _num(m[1]) == 0


def _fit_redrawn(text: str, size: float, w: float, h: float, auto: bool, unit: float = 1.0) -> tuple[float, bool]:
    """元からある 1 行の欄：見た目を描き直すビューアーでも欠けない字の大きさ。返す：(大きさ, /DA に大きさを書くか)。
    読める大きさ（MIN_READABLE_PT）より小さくはしない（そこまで小さくしても入らなければ、大きさはそのまま）。"""
    font = "" if all(ord(c) < 0x80 for c in text) else "J"
    pinned = auto and pdfkit_needs(text, h * REDRAW_AUTO_HEIGHT, font) > w
    s = size
    floor = MIN_READABLE_PT / unit
    while s > floor and pdfkit_needs(text, s, font) > w:
        s = max(floor, s - 0.5)
    if pdfkit_needs(text, s, font) > w:
        return size, pinned
    return s, pinned or s < size


def _da_key(keys: set[str]) -> str:
    """欄の書体の指定（/DA）にする書体：見た目を作り直すソフトが使うので、日本語などの書体を先に（欧文の拡張の TdkL は最後）。"""
    for k in (BUNDLED_KEY, "TdkJ", "TdkK", "TdkG", "TdkC", "TdkE"):
        if k in keys:
            return k
    for k in sorted(keys):
        if k.startswith("TdkX"):
            return k
    if "TdkL" in keys:
        return "TdkL"
    return "Helv"


def _num(t: str) -> float | None:
    try:
        return float(t)
    except ValueError:
        return None


def _da_size(da: str) -> float | None:
    m = re.search(r"([\d.]+)\s+Tf", da or "")
    if m:
        v = _num(m[1])
        return v if v is not None and v > 0 else None
    return None


def _pt(s: float) -> str:
    """字の大きさを /DA に書く形に（10.0 → 10、7.5 → 7.5）。"""
    return f"{round(s, 2):g}"


def _da_refont(da: str, key: str) -> str | None:
    """/DA の書体の名前（Tf の命令の名前）だけを key にした指定。Tf の命令がなければ None。"""
    m = re.search(r"/[^\s/\[\]<>()]+(\s+[\d.]+\s+Tf)", da or "")
    if not m:
        return None
    return da[:m.start()] + f"/{key}" + m[1] + da[m.end():]


def _da_resize(da: str, s: float) -> str | None:
    """/DA の字の大きさ（Tf の前の数。_da_size と同じ所）だけを s にした指定。書体・色などはそのまま。"""
    m = re.search(r"([\d.]+)(\s+Tf)", da or "")
    if not m:
        return None
    return da[:m.start(1)] + _pt(s) + da[m.end(1):]


# リストの選んだ行の色：Acrobat がリストの見た目で使う色（RGB）。書類の集まりから決めた値ではない
LIST_HIGHLIGHT = (0.6, 0.757, 0.855)


def list_appearance(update: Update, w: float, h: float, labels: list[str], selected: set[int], top: int, size: float,
                    dr_fonts, rot: int = 0, unit: float = 1.0) -> tuple[Ref, float, int, list[int]]:
    """リスト（一覧から選ぶ欄で、コンボでないもの）の見た目：選択肢を上から 1 行ずつ並べ、選んだ行に色を付ける
    （ISO 32000-1 12.7.4.4 のリストの描き方。行の高さは字の大きさ）。選んだ行が欄に入らなければ、一番上に見せる行（/TI）を
    ずらす。選んだ選択肢の字が欄の幅に入らなければ、字を小さくする（4pt でも入らなければ FillError）。
    返す値：(見た目, 字の大きさ, 一番上に見せる行, 見える行の番号)。
    rot：欄の回転（/MK の R）。w・h は欄の Rect の幅と高さで、描く向きの座標（90・270 度は幅と高さが入れ替わる）で並べる。"""
    rect_w, rect_h = w, h
    w, h = _view_box(rot, w, h)
    pad = min(2.0, 0.1 * w)
    s = size
    floor = 4.0 / unit - 1e-9  # 小さくする下限：紙の上の 4pt
    while s >= floor and any(text_width(labels[i], s) > w - 2 * pad for i in selected):
        s -= 0.5
    if s < floor:
        raise FillError("選んだ選択肢の字が、リストの幅（" + f"{w * unit:.0f}pt）に入りきりません")
    line = s
    rows = max(1, int((h - 2 * pad) // line))
    if selected:
        first, last = min(selected), max(selected)
        if not (top <= first < top + rows):
            top = first
        if last >= top + rows and last - first < rows:
            top = last - rows + 1
        if any(not (top <= i < top + rows) for i in selected):
            raise FillError(f"選んだ選択肢が離れていて、リストの高さ（{rows} 行）に一度に見せられません")
    visible = list(range(top, min(len(labels), top + rows)))
    ops = bytearray(b"/Tx BMC\nq\n")
    ops += f"{pad:.2f} {pad:.2f} {w - 2 * pad:.2f} {h - 2 * pad:.2f} re W n\n".encode()
    y_top = h - pad
    for n, i in enumerate(visible):
        y = y_top - (n + 1) * line
        if i in selected:
            r, g, b = LIST_HIGHLIGHT
            ops += f"{r} {g} {b} rg\n{pad:.2f} {y:.2f} {w - 2 * pad:.2f} {line:.2f} re f\n".encode()
    for n, i in enumerate(visible):
        y = y_top - (n + 1) * line + 0.22 * s  # 基線：行の下から、字の下がり（ディセンダー）ぶん上
        if labels[i]:
            ops += text_ops([labels[i]], s, pad + 1.0, y, line)
    ops += b"Q\nEMC\n"
    keys = fonts_used([labels[i] for i in visible if labels[i]])
    return _ap(update, w, h, bytes(ops), keys, dr_fonts, _rot_matrix(rot, rect_w, rect_h)), s, top, visible


def fill(doc: PdfDocument, values: dict, *, force: bool = False, placement: bool = True) -> FillResult:
    """欄の名前（窓の［］の中）→ 値。フォームの欄はその名前、配置から見つけた欄は p{ページ}.{番号}。
    漢字の字形は、書類の言語（Catalog の /Lang）の字形の集まりから先に選ぶ（glyphs.cjk_order）。"""
    from .objects import PdfString, decode_text

    lang = doc.resolve(doc.root.get("Lang")) if hasattr(doc, "root") and isinstance(doc.root, dict) else None
    lang = decode_text(lang) if isinstance(lang, (bytes, PdfString)) else ""
    with glyphs.cjk_order(lang), glyphs.embed_first(is_pdfa(doc)):
        return _fill(doc, values, force=force, placement=placement)


def is_pdfa(doc) -> bool:
    """文書が PDF/A と名乗っているか（Catalog の /Metadata の pdfaid。ISO 19005）。"""
    try:
        st = doc.resolve(doc.root.get("Metadata"))
        return isinstance(st, Stream) and b"pdfaid" in doc.decode(st)
    except Exception:  # noqa: BLE001  読めないメタデータは、名乗っていないものとして扱う
        return False


def _fill(doc: PdfDocument, values: dict, *, force: bool = False, placement: bool = True) -> FillResult:
    from .window import build as build_window

    f = Filler(doc)
    if f.sigs.docmdp == 1:
        raise FillError("この PDF は「変更を許さない」電子署名（認証、DocMDP P=1）がされています。書き込むと署名が無効になるので、書きません")
    win = build_window(doc)
    spots = {}
    views = {}
    for pg in win["pages"]:
        for fd in pg["fields"]:
            spots[fd["id"]] = (pg["number"], fd)
    problems = []
    used: dict = {}
    circ_used: dict = {}  # ○で囲んだ所（ページ → 番号と楕円）
    layouts: dict = {}
    _uniform_sizes(f, doc, values, spots)
    xf = _xfa_form(doc)
    xfa_values = {k: v for k, v in values.items() if xf is not None and k in xf.by_id}
    for key, value in values.items():
        if key in xfa_values:
            continue
        try:
            _signed_new(f, key, value, spots)
            if xf is not None and (key.startswith("@") or (key not in f.by_name and key in spots)):
                raise FillError(f"{key}: この PDF は XFA のフォームです。XFA のフォームの欄でない所に書いても、Adobe Acrobat・"
                                "Reader では表示されないことがあるので、書きません（窓の XFA の入力欄の id を指定してください）")
            if key.startswith("@"):
                # 窓に欄として示していない所：エージェントが見出しの字を目印にして指した所（その右か下の空いている所）
                from .anchor import AnchorError, resolve_write

                if not isinstance(value, dict):
                    raise FillError(f'{key}: 欄のない所に書くときは {{"after": 見出しの字, "value": 書く字}} の形で指定してください')
                from ..sealimage import SealError, seal_key

                try:
                    skey = seal_key(value)
                except SealError as e:
                    raise FillError(f"{key}: {e}") from None
                if skey is not None:
                    # 本人の印影・署名の画像を置く（本人が同意したときだけ）：スタンプの注釈
                    _place_seal(f, key, value, skey, used, layouts, views)
                    continue
                if "circle" in value:
                    # 印刷された字を○で囲む：ページの描画の後ろに、楕円を描くストリームを足す
                    from . import circle as _circle

                    try:
                        sd = _circle.resolve(doc, value, layouts, circ_used)
                    except AnchorError as e:
                        raise FillError(f"{key}: {e}") from None
                    f.circles.setdefault(sd["page"], []).append(sd)
                    if sd.get("note"):
                        f.warnings.append(sd["note"])
                    f.patches.append(Patch(sd["id"], "circle", sd["text"], sd["page"], anchor=dict(value), key=key, circle=sd))
                    continue
                try:
                    sd = resolve_write(doc, value, used, layouts, set(f.by_name))
                except AnchorError as e:
                    raise FillError(f"{key}: {e}") from None
                page_no = sd["page"]
                if page_no not in views:
                    views[page_no] = View(doc, f.pages[page_no - 1])
                opts = {k: value[k] for k in ("align", "pad_mm") if k in value}
                check_layout_options(key, opts)
                f.add_widget(page_no, {**sd, **opts}, value["value"], views[page_no])
                if sd.get("note"):
                    f.warnings.append(f"{sd['id']}: {sd['note']}")
                f.patches[-1].anchor = dict(value)
                f.patches[-1].key = key
                continue
            if key in f.by_name or (key in spots and spots[key][1].get("source") == "acroform"):
                fld = f.by_name.get(key) or f.by_name[spots[key][1]["name"]]
                val, allow_hidden = _field_value(key, value)
                sd_ = spots.get(fld.name, (None, {}))[1]
                opts, olabs = sd_.get("options") or [], sd_.get("options_labels") or []
                if isinstance(val, str) and val not in opts and olabs.count(val) == 1:
                    val = opts[olabs.index(val)]  # 選択肢の隣に印刷された字で選んだ（「Female」→ 選択肢の名前 F）
                n_ = len(f.patches)
                f.set_field(fld, val, force, allow_hidden=allow_hidden)
                if olabs and len(f.patches) > n_ and f.patches[-1].kind == "control" and f.patches[-1].value in opts:
                    # 読み戻しと本人に伝える値は、窓の値と同じ、選んだ選択肢の隣に印刷された字（/V と /AS は expect で照らす）
                    f.patches[-1].export = f.patches[-1].value  # 計算のスクリプトに渡す値は、選択肢の名前のまま
                    f.patches[-1].value = olabs[opts.index(f.patches[-1].value)]
                continue
            if key not in spots:
                from .window import unknown_id_message

                hint = unknown_id_message(key)
                raise FillError(f"{key}: そのような欄はありません" + (f"。{hint}" if hint else ""))
            page_no, sd = spots[key]
            if page_no not in views:
                views[page_no] = View(doc, f.pages[page_no - 1])
            opts, value = layout_options(key, value)
            sd, value = as_kind({**sd, **opts}, value)
            _stacked_note(f, doc, page_no, sd, value, layouts)
            f.add_widget(page_no, sd, value, views[page_no])
        except FillError as e:
            msg = str(e)
            problems.append(msg if msg.startswith(str(key)) or ":" in msg.split(chr(10))[0][:60] else f"{key}: {msg}")
    # 書いた欄すべて（配置の所・元からあるフォームの欄）：配置の所の alternative_to は、フォームの欄の名前のこともある
    written = [spots[k][1] if k in spots else {"id": k} for k in values if k in spots or k in f.by_name]
    problems += alternative_conflicts(written)
    if xfa_values and not problems:
        try:
            _fill_xfa(f, xf, xfa_values, force)
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    if f.patches and not xfa_values:
        _recalculate(f, values)
    note = f.sigs.note()
    if note and f.patches:
        f.warnings.insert(0, note)
    if f.patches and _need_appearances(doc) and not f.clear_need_appearances():
        f.warnings.append("この PDF は、開くときにビューアーが欄の見た目を作り直す設定です（AcroForm の NeedAppearances が true）。"
                          "とどくが描いて確かめた見た目ではなく、ビューアーが値から描き直した見た目で表示・印刷されます"
                          "（字の大きさ・折り返し・書体が変わることがあります）。書いた値（/V）は確かめたとおりです")
    denied = _permissions(f)
    blob = f.finish() if f.patches else doc.blob
    if f.patches and placement:
        # 書いた字・○が、印刷された字・罫線・同じ書き込みのほかの値に重ならないか（書いた後の文書の形で）。
        # placement=False は、監査の試しの書き込み（欄の作りの確かめ）だけ
        from . import placement as _placement

        # とどくが置いた字の欄（チェックの印は、印刷された□の上に置くもの：(a)・(b) は見ない）
        created = {p.id for p in f.patches if p.kind == "input" and p.id not in f.by_name}
        probs, notes = _placement.check(doc, _placement.reopen(doc, blob), f.patches, created, f.covers)
        if probs:
            raise FillError("書いた所が、ほかの字や線に重なるので、書き出していません：\n" + "\n".join(probs[:12]))
        f.warnings += [x for x in notes if x not in f.warnings]
    fne = _fonts_not_embedded(f.patches, set(f.by_name))
    if BUNDLED_KEY in f.dr_fonts and glyphs.first_order() != glyphs.ORDER[0] \
            and any("\u4e00" <= c <= "\u9fff" for p in f.patches if isinstance(p.value, str) for c in p.value):
        f.warnings.append("この書類の言語は日本語ではありませんが、同梱の書体（Noto Sans JP）に字形のある漢字は、日本語の字形で描きました。"
                          "同じ字でも、言語によって形が少し違うことがあります（字そのものは同じで、読み上げ・検索の字も同じ）")
    cjk = [x for x in fne if x["kind"] == "cjk_predefined"]
    if cjk:
        f.warnings.append(f"{'・'.join(x['font'] for x in cjk)} の字は、書体を埋め込まず、開くソフトが持つ書体で描きます（PDF の仕様で決まった字の集まりを指す形）。"
                          "その言語の書体がないソフト・環境では、字が出なかったり別の字形になったりすることがあります（ファイルの中の値は正しい）。"
                          "相手の環境で見えるかは確かめていません")
    if fne and f.pdfa:
        f.warnings.append(f"この PDF は PDF/A と名乗っていますが、{'・'.join(x['font'] for x in fne)} の字は書体を埋め込めませんでした"
                          "（同梱の書体にもこのパソコンの書体にも字形がありません）。PDF/A の決まりに合わなくなります")
    return FillResult(blob, f.patches, f.warnings, denied=denied, structure=f.structure, redraw_cut=f.redraw_cut, fonts_not_embedded=fne)


def _fonts_not_embedded(patches: list, acro_names: set | None = None) -> list[dict]:
    """書いた字のうち、書体を埋め込まずに描いたもの（事実）：standard14 は PDF の仕様でどのソフトも代わりを持つ書体（Helvetica）、
    cjk_predefined は定義済みの CMap で字の集まりを指し、字形はソフトの持つ書体に頼るもの（平成角ゴシックなど）。埋め込んだ書体は入れない。
    xfa_template は、PDF の欄のない XFA の欄：とどくは見た目を描かず、開くソフトが様式の書体で描く。"""
    names = {v[0]: v[1] for v in CJK.values()}
    names.update({"Helv": "Helvetica", "TdkL": "Helvetica"})
    got: dict[str, list[str]] = {}
    xfa_only: list[str] = []
    for p in patches:
        text = p.shown if p.shown is not None else p.value
        if p.kind not in ("input", "calc", "xfa") or not isinstance(text, str) or not text or glyphs.shaped(text):
            continue
        if p.kind == "xfa" and acro_names is not None and p.id not in acro_names:
            xfa_only.append(p.id)
            continue
        try:
            keys = {k for k, _r in runs_for(text)}
        except FillError:
            continue
        for k in keys & set(names):
            ids = got.setdefault(names[k], [])
            if p.id not in ids:
                ids.append(p.id)
    cjk = {v[1] for v in CJK.values()}
    out = [{"font": n, "kind": "cjk_predefined" if n in cjk else "standard14", "ids": ids[:50]} for n, ids in got.items()]
    if xfa_only:
        out.append({"font": "XFA の様式の書体", "kind": "xfa_template", "ids": xfa_only[:50]})
    return out


def _natural_size(spot: dict, text: str, unit: float) -> float | None:
    """窓の欄に text を書くときの字の大きさ（add_widget・text_appearance と同じ決め方）。入らなければ None。"""
    x0, y0, x1, y1 = spot["rect"]
    w, h = x1 - x0, y1 - y0
    size = float(spot.get("size") or 10.5 / unit)
    if spot.get("kind") == "comb":
        boxes = spot.get("box_rects") or []
        return min(size, min(b[3] - b[1] for b in boxes) * 0.7) if boxes else None
    try:
        return layout_text(text, w, h, size, False, spot_multiline(spot, h, size, unit), unit)[1]
    except FillError:
        return None


def _uniform_sizes(f: Filler, doc, values: dict, spots: dict) -> None:
    """同じ行に隙間なく並ぶマス（一続きの欄：口座名義の 1 字ずつのマス）に、同じ書き込みで字を書くとき、字の大きさを
    そろえる。マスの大きさが違うと、欄ごとに字の大きさが変わり、小さい字が小書きの字（「ヤ」が「ャ」）に読めるので、
    いちばん小さい欄の大きさ（読める大きさより小さくはしない）に合わせ、そろえたことを伝える。"""
    items = []
    for k, v in values.items():
        if k not in spots or k in f.by_name or isinstance(v, (dict, list, tuple, bool)) or v is None:
            continue
        n, sd = spots[k]
        if sd.get("source") == "acroform" or sd.get("kind") not in ("text", "comb") or not sd.get("rect"):
            continue
        t = _as_text(v)
        if not t or "\n" in t:
            continue
        items.append((n, sd, t))
    if len(items) < 2:
        return
    items.sort(key=lambda x: (x[0], round(x[1]["rect"][1]), x[1]["rect"][0]))
    groups, cur = [], []
    for it in items:
        if cur:
            pn, ps, _pt_ = cur[-1]
            r0, r1 = ps["rect"], it[1]["rect"]
            same_row = it[0] == pn and abs(r0[1] - r1[1]) <= 1.0 and abs(r0[3] - r1[3]) <= 1.0
            if same_row and -1.0 <= r1[0] - r0[2] <= 2.0:
                cur.append(it)
                continue
            groups.append(cur)
        cur = [it]
    groups.append(cur)
    for g in groups:
        if len(g) < 2:
            continue
        unit = View(doc, f.pages[g[0][0] - 1]).unit
        sizes = [(sd["id"], _natural_size(sd, t, unit)) for _n, sd, t in g]
        known = [s for _i, s in sizes if s]
        if len(known) < 2 or max(known) - min(known) <= 0.5:
            continue
        target = max(min(known), MIN_READABLE_PT / unit)
        for fid, s in sizes:
            if s and s > target:
                f.size_override[fid] = target
        small = [fid for fid, s in sizes if s and s < target - 0.05]
        f.warnings.append("・".join(fid for fid, _s in sizes) + ": 同じ行に並ぶマスに書いた字の大きさを、"
                          f"{_pt(target * unit)} pt にそろえました（マスの大きさが違い、欄ごとに字の大きさが変わると、小さい字が"
                          "小書きの字（「ヤ」が「ャ」）に読めることがあるため）"
                          + (f"。{'・'.join(small)} は、読める大きさ（{MIN_READABLE_PT:g} pt）でも入らないので小さいままです。"
                             "本人に確かめてください" if small else ""))


def _stacked_note(f: Filler, doc, page_no: int, sd: dict, value, layouts: dict) -> None:
    """窓の欄（字の欄）のすぐ右か左に、印刷された字の行が縦に 2 つ以上並ぶ（「銀行／信用金庫／信用組合／農協」）とき：
    書いた字がどの行の高さに来るかで、どの行に続けて読めるかが変わる。置く所は決めずに（エージェントが決める）、
    行が並ぶことと、書いた字の高さ（mm）を知らせる。別の行に合わせるなら、box で枠を決めて書く。"""
    if sd.get("kind") != "text" or sd.get("source") == "acroform" or not sd.get("rect") or "\n" in str(value):
        return
    from . import circle as _circle

    lay = _circle.layout(doc, page_no, layouts)
    x0, y0, x1, y1 = sd["rect"]
    size = float(sd.get("size") or 10.5)
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    # 欄のあるマス（同じマスの行だけ）
    home = next((c for c in lay.index.near(cx, cy, cx, cy) if c.x0 <= cx <= c.x1 and c.y0 <= cy <= c.y1), None)
    rows = []
    for cell_, ln in _circle.lines_between(lay, y0 - 0.5, y1 + 0.5):
        if ln.vertical or not ln.text.strip():
            continue
        if home is not None and cell_ is not home:
            continue
        if (ln.x0 >= x1 - 1 and ln.x0 - x1 <= 3 * size) or (ln.x1 <= x0 + 1 and x0 - ln.x1 <= 3 * size):
            rows.append(ln)
    rows.sort(key=lambda ln: ln.y0)
    distinct = []
    for ln in rows:
        if not distinct or ln.y0 >= distinct[-1].y1 - 0.2 * ln.size:
            distinct.append(ln)
    if len(distinct) < 2:
        return

    def short(ln) -> str:
        return re.sub(r"\s+", "", ln.text)[:8]

    mm = lay.view.mm
    names = "／".join(short(ln) for ln in distinct[:6])
    level = next((ln for ln in distinct if ln.y0 <= cy <= ln.y1), None)
    f.warnings.append(f"{sd['id']}: 横に字の行が {len(distinct)} 行並ぶマス（{names}）です。書いた字は、欄の縦の中ほど（上 "
                      f"{cy / mm:.1f} mm" + (f"・「{short(level)}」の行の高さ" if level else "")
                      + "）に置きました。view_page で見て、別の行に合わせるなら box（mm）で枠を決めて書き直してください")


def _permissions(f: Filler) -> list[dict]:
    """作り手の権限の設定（暗号の辞書の P、ISO 32000-1 表 22）が許していない書き込みをしたか。

    断らずに書く：権限の設定は、作り手が「ビューアーに許す操作」を決めるもので、本人が自分の申請書に記入することまで
    禁じるものとは限らない（記入の欄があるのに、記入を許す印を付け忘れた書類もある）。ただし、受け取る側のソフトが記入を
    拒んだり、窓口が「電子で記入したもの」を受け付けなかったりする（紙に印刷して手で書く様式のこともある）ので、
    本人に伝えること（warnings → say_to_user）の先頭に入れ、書類の書く段階も下げる（grade の「権限の設定」）。"""
    dec = f.doc.decryptor
    if dec is None or not f.patches:
        return []
    ops = set()
    for p in f.patches:
        if p.kind == "seal":
            ops.add("annotate")
        elif p.kind == "circle":
            ops.add("modify")
        elif p.kind in ("input", "control") and p.id not in f.by_name:
            ops.add("create")
        else:
            ops.add("fill")
    denied = [{"op": op, "what": what} for op, what in dec.restrictions() if op in ops]
    if denied:
        f.warnings.insert(0, "この PDF は、作り手が権限の設定（暗号の辞書の P）で「" + "」「".join(d["what"] for d in denied)
                          + "」を許していません。とどくは書きましたが、受け取る側のソフトが記入を拒んだり、窓口が受け付けな"
                          "かったりすることがあります（紙に印刷して手で書く様式のこともあります）。この書き出したファイルで"
                          "提出してよいか、提出先に確かめてください")
    return denied


def _field_value(key: str, value) -> tuple:
    """フォームの欄に書く値：{"value": 値, "allow_hidden": true} の形なら、(値, True)。ほかは (値, False)。"""
    if isinstance(value, dict) and "allow_hidden" in value:
        if set(value) - {"value", "allow_hidden"} or "value" not in value or not isinstance(value["allow_hidden"], bool):
            raise FillError(f'{key}: 隠れた欄に書くときは {{"value": 値, "allow_hidden": true}} の形で指定してください')
        return value["value"], value["allow_hidden"]
    return value, False


def _need_appearances(doc) -> bool:
    af = doc.resolve(doc.root.get("AcroForm"))
    return isinstance(af, dict) and doc.resolve(af.get("NeedAppearances")) is True


def _signed_new(f: Filler, key: str, value, spots: dict) -> None:
    """認証の電子署名（DocMDP P=2・3）：新しい欄・注釈・ページの描画を足す書き込みを、書く前に断る。"""
    if f.sigs.docmdp is None or (key in f.by_name):
        return
    from .signatures import refuse_new

    if key.startswith("@") and isinstance(value, dict):
        if "seal" in value or "signature" in value:
            why = refuse_new(f.sigs, "印影・署名の画像（スタンプの注釈）", annotation=True)
        elif "circle" in value:
            why = refuse_new(f.sigs, "○（ページの描画）")
        else:
            why = refuse_new(f.sigs, "新しい欄（字を目印にした書き込み・枠）")
    elif key in spots and spots[key][1].get("source") != "acroform":
        why = refuse_new(f.sigs, "新しい欄（窓の推定した欄）")
    else:
        return
    if why:
        raise FillError(f"{key}: {why}")


def appearance_readback(doc: PdfDocument, patches: list[Patch]) -> list[dict]:
    """書いた後の文書で、書いた文字の欄の見た目（/AP）を読み直す：描いた字が、見せるはずの字と同じで、どれも描ける字形か。
    （/V の読み戻しでは、見た目の字が白く抜けても気づかないので。）空白・改行は比べない（折り返しで変わるため）。"""
    from .glyphs import appearance_text, widget_appearance

    by = {x.name: x for x in read_fields(doc)}
    out = pagedraw.read_back(doc, patches)
    for p in patches:
        fld = by.get(p.id)
        if fld is not None and p.expect:
            out += _state_readback(doc, fld, p)
        if p.kind not in ("input", "calc") or p.shown is None:
            continue
        if fld is None:
            continue
        want = re.sub(r"\s", "", str(p.ap_text if p.ap_text is not None else p.shown))
        for w in fld.widgets:
            wd = doc.resolve(w.ref) if w.ref is not None else doc.resolve(fld.ref)
            st = widget_appearance(doc, wd)
            if st is None:
                out.append({"id": p.id, "written": p.shown, "read_back": "見た目（/AP）がありません"})
                break
            got, probs = appearance_text(doc, st)
            if probs:
                out.append({"id": p.id, "written": p.shown, "read_back": "見た目の字が描けません：" + "／".join(probs)[:300]})
                break
            if re.sub(r"\s", "", got) != want:
                out.append({"id": p.id, "written": p.shown, "read_back": f"見た目の字：{got}"})
                break
    return out


def _state_readback(doc: PdfDocument, fld: FormField, p: Patch) -> list[dict]:
    """書いた後の文書で、チェック・ラジオの値（/V）と部品ごとの見た目の状態（/AS）、リストの選んだ番号（/I）を読み直し、
    書いたつもりの状態と照らす（値が合っていても、別の部品が入になっていれば、紙では別の所に印が付く）。"""
    out = []
    ex = p.expect or {}
    if "V" in ex:
        v = fld.value
        got = v.name if isinstance(v, Name) else None
        if got != ex["V"]:
            out.append({"id": p.id, "written": ex["V"], "read_back": f"値（/V）：{got}"})
        states = {w.ref.num: (w.state or "Off") for w in fld.widgets if w.ref is not None}
        for num, want in ex.get("AS") or []:
            got_as = states.get(num) if num is not None else next((w.state or "Off" for w in fld.widgets if w.ref is None), None)
            if got_as != want:
                out.append({"id": p.id, "written": want,
                            "read_back": f"部品（オブジェクト {num}）の見た目の状態（/AS）：{got_as}"})
        on = [w for w in fld.widgets if (w.state or "Off") != "Off"]
        if any(w.state != got for w in on):
            out.append({"id": p.id, "written": ex["V"],
                        "read_back": "値と違う状態の部品が入になっています：" + "、".join(str(w.state) for w in on)})
    if "I" in ex:
        node = doc.resolve(fld.ref) if fld.ref is not None else None
        got_i = doc.resolve(node.get("I")) if isinstance(node, dict) else None
        got_i = [doc.resolve(x) for x in got_i] if isinstance(got_i, list) else got_i
        if ex["I"] is not None and got_i != ex["I"]:
            out.append({"id": p.id, "written": ex["I"], "read_back": f"選んだ番号（/I）：{got_i}"})
    return out


def as_kind(sd: dict, value):
    """種類が形だけでは決まらない欄（kind_candidates のある欄）に、値からどの種類で書くかを決める。
    - {"kind": 種類, "value": 値}：その種類で書く（種類は kind_candidates のどれか）。"as" は書いた値の意味で、種類ではない
      （以前の書き方 {"as": 種類, "value": 値} は、"kind" がなく "as" が種類の候補とちょうど同じ字のときだけ、種類として受ける）
    - はい／いいえ（true / false）：チェックの印（候補に checkbox があるとき）
    - 字：候補に text があり、既定の種類がチェックの印なら、文字の欄として書く（「1」「レ」などをマスに書く）。
      既定が1字ずつ書く欄（comb）なら、1字ずつマスに書く（1つの欄として書くときは {"kind": "text", …}）
    返す値：(書く欄の情報, 書く値)。"""
    cands = sd.get("kind_candidates") or []
    if isinstance(value, dict) and "kind" not in value and "as" in value and value["as"] not in cands and "value" in value:
        return as_kind(sd, value["value"])  # as は書いた値の意味（種類の指定ではない）
    if isinstance(value, dict) and ("kind" in value or "as" in value):
        want = value.get("kind", value.get("as"))
        if not cands or want not in cands:
            raise FillError(f"{sd['id']}: 「{want}」としては書けません（書ける種類：{'・'.join(cands) or sd['kind']}）")
        if "value" not in value:
            raise FillError(f'{sd["id"]}: {{"kind": 種類, "value": 書く値}} の形で指定してください')
        return {**sd, "kind": want}, value["value"]
    if not cands:
        return sd, value
    if isinstance(value, bool):
        if "checkbox" in cands:
            return {**sd, "kind": "checkbox"}, value
        return sd, value
    if sd["kind"] == "checkbox" and "text" in cands:
        return {**sd, "kind": "text"}, value
    return sd, value


def alternative_conflicts(written: list[dict]) -> list[str]:
    """同じ所の別の読み方（alternative_to でつながる候補）の両方に書こうとしたもの：どちらか1つにする（重なって書かれるので）。
    written：書いた欄（id と、あれば alternative_to・rect）。alternative_to の先は、配置の所のことも、元からあるフォームの欄の
    名前のこともある（フォームの欄に 1 字ぶん以上重なる所）。"""
    out = []
    ids = {d["id"] for d in written}
    groups: dict[str, list[dict]] = {}
    for d in written:
        alt = d.get("alternative_to")
        if not alt:
            continue
        if alt in ids:
            out.append(f"{d['id']} と {alt} は、同じ所の別の読み方（重なる候補）です。どちらか1つにだけ書いてください")
        groups.setdefault(alt, []).append(d)
    for g in groups.values():
        for i, a in enumerate(g):
            for b in g[i + 1:]:
                if conflicting(a, b):
                    out.append(f"{a['id']} と {b['id']} は、同じ所の別の読み方（重なる候補）です。どちらか1つにだけ書いてください")
    return out


def conflicting(a: dict, b: dict) -> bool:
    """2つの欄が、同じ所の別の読み方か：一方が他方の alternative_to か、同じ候補の別の読み方どうしで重なるもの。"""
    if a.get("alternative_to") == b["id"] or b.get("alternative_to") == a["id"]:
        return True
    if a.get("alternative_to") and a.get("alternative_to") == b.get("alternative_to"):
        r, s = a["rect"], b["rect"]
        return r[0] < s[2] and s[0] < r[2] and r[1] < s[3] and s[1] < r[3]
    return False


def _place_seal(f: Filler, key: str, spec: dict, skey: str, used: dict, layouts: dict, views: dict) -> None:
    from ..sealimage import SealError, load
    from . import seal
    from .anchor import AnchorError

    try:
        img = load(spec, skey)  # 同意がなければ、ここで断る
        if not hasattr(f, "_stamp_names"):
            f._stamp_names = seal.stamp_names(f.doc)
        sd = seal.resolve(f.doc, spec, skey, img, used, layouts, set(f.by_name) | f._stamp_names)
        page_no = sd["page"]
        if page_no not in views:
            views[page_no] = View(f.doc, f.pages[page_no - 1])
        seal.write(f, page_no, sd, skey, img, views[page_no])
    except (SealError, AnchorError) as e:
        raise FillError(f"{key}: {e}") from None
    p = Patch(sd["id"], "seal", f"本人の{sd['what']}の画像", page_no, anchor=dict(spec), key=key)
    p.seal = {"told": seal.told(sd, skey), "rect": sd["rect"], "sha256": img.sha256, "what": sd["what"], "key": skey,
              "mm": sd["mm"]}
    f.patches.append(p)
    caution = seal.caution(sd["what"])
    if caution not in f.warnings:
        f.warnings.append(caution)


def _xfa_form(doc):
    """XFA のフォーム（読めなければ None：そのときは、これまでどおり PDF の欄として扱う）。"""
    from .xfa import XfaError, read

    try:
        return read(doc)
    except XfaError:
        return None


def _fill_xfa(f: Filler, xf, values: dict, force: bool) -> None:
    """XFA の欄に書く：datasets を書き換え、静的なフォームでは、同じ名前の PDF の欄の値と見た目もそろえる。"""
    from .xfafill import XfaFillError, new_datasets, packet_bytes, stream_for, write

    try:
        combined, done, notes, edit = write(xf, values, force)
    except XfaFillError as e:
        raise FillError(str(e)) from None
    doc = f.doc
    xdp = xf.xdp
    kind, data = packet_bytes(xf, combined, edit)
    if kind == "datasets":
        ref = xdp.packet_of(edit.ds).ref
        if ref is None:
            raise FillError("datasets の部品が、別のオブジェクトになっていないので書けません")
        f.up.put(ref, stream_for(doc, ref, data))
    elif kind == "xdp":
        if xdp.stream_ref is None:
            raise FillError("XFA のストリームが、別のオブジェクトになっていないので書けません")
        f.up.put(xdp.stream_ref, stream_for(doc, xdp.stream_ref, data))
    else:
        ref = f.up.new(stream_for(doc, None, new_datasets(xf, edit)))
        pair = [PdfString(b"datasets"), ref]
        if xdp.array_ref is not None:
            arr = f.up.get(xdp.array_ref)
            arr[len(arr) - 2:len(arr) - 2] = pair
        else:
            af_ref = doc.root.get("AcroForm")
            af = f.up.get(af_ref) if isinstance(af_ref, Ref) else f.up.get(doc.trailer["Root"])["AcroForm"]
            arr = list(af["XFA"])
            arr[len(arr) - 2:len(arr) - 2] = pair
            af["XFA"] = arr
    f.warnings += notes
    for key, (raw, shown) in done.items():
        x = xf.by_id[key]
        acro = f.by_name.get(key)
        if acro is not None and acro.kind not in ("button", "signature"):
            # 静的なフォーム：PDF の欄の値と見た目（書式で整えた字）も同じにする
            if x.kind == "checkbox":
                val, disp = bool(shown), None
            elif x.kind == "radio":
                val, disp = (raw or "Off"), None
            else:
                val = "\n".join(raw) if isinstance(raw, list) else (shown if shown is not None else "")
                disp = xf.display(_with_raw(x, val)) if val else None
            try:
                if x.kind == "radio" and not raw:
                    _radio_off(f, acro)
                else:
                    # 欄の値（/V）：XFA を読まず /V から描き直す表示（PDFKit）に、書式どおりの字が出るよう、文字の欄は見せる字を入れる。
                    # 決まった形の値は XFA のデータに残る（XFA を読むソフトは、そちらを使う）
                    if acro.kind == "text" and disp and not (acro.maxlen and len(disp) > acro.maxlen):
                        val = disp
                    f.set_field(acro, val, True, display=disp, calc=True)
                    f.patches.pop()
            except FillError as e:
                raise FillError(f"{key}: XFA のデータは書けますが、同じ名前の PDF の欄に書けません（{e}）") from None
        f.patches.append(Patch(key, "xfa", shown))


def _with_raw(x, raw):
    import copy

    y = copy.copy(x)
    y.raw = raw
    return y


def _radio_off(f: Filler, acro: FormField) -> None:
    fd = f.up.get(acro.ref)
    fd["V"] = Name("Off")
    for w in acro.widgets:
        wd = f.up.get(w.ref) if w.ref is not None and w.ref != acro.ref else fd
        wd["AS"] = Name("Off")


def _recalculate(f: Filler, values: dict) -> None:
    """計算の欄（AFSimple_Calculate）を、フォームの計算の順（/CO）に計算し直して書く。"""
    from . import acrojs

    doc = f.doc
    af = doc.resolve(doc.root.get("AcroForm"))
    order = [r.num for r in (doc.resolve(af.get("CO")) or []) if isinstance(r, Ref)] if isinstance(af, dict) else []
    calc = []
    for fld in f.fields:
        if fld.ref is None:
            continue
        js = acrojs.actions(doc, doc.resolve(fld.ref))
        if "C" in js:
            calc.append((order.index(fld.ref.num) if fld.ref.num in order else len(order), fld, js))
    if not calc:
        return
    current = {fld.name: fld.value_text() for fld in f.fields}
    for p in f.patches:
        current[p.id] = getattr(p, "export", p.value)
    for _i, fld, js in sorted(calc, key=lambda x: x[0]):
        r = acrojs.simple_calculate(js["C"], lambda n: current.get(n))
        if r is None:
            f.warnings.append(f"{fld.name}: 独自の計算のスクリプトは実行していません（計算の結果は、Acrobat で開いたときに決まります）")
            continue
        val = int(r) if float(r).is_integer() else round(r, 10)
        if str(current.get(fld.name)) == str(val):
            continue
        try:
            f.set_field(fld, val, force=True, calc=True)
        except FillError as e:
            f.warnings.append(f"{fld.name}: 計算した値 {val} を書けません（{e}）")
            continue
        current[fld.name] = f.patches[-1].value
        f.patches[-1].kind = "calc"
