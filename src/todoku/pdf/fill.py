"""PDFに書き込む：元のバイト列の後ろに、変えたオブジェクトと新しいオブジェクトを足す（増分更新）。

- フォームの欄（AcroForm）：値（/V）と見た目（/AP）を書く。チェック・ラジオは状態（/AS）も
- 配置から見つけた欄（印刷用のPDF）：同じ名前のフォームの欄を、その位置に足す（見た目つき）
  元のページの描画（コンテンツストリーム）は1バイトも変えない。値は機械で読める（/V）ので、読み戻せる
- 文字の見た目は、欧文は Helvetica、日本語は平成角ゴシック（UniJIS-UCS2-HW-H）、ハングル・中国語も対応する字形の集まりで描く
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field

from .doc import PdfDocument
from . import glyphs
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


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unverified: list[str] = field(default_factory=list)
    # 作り手の権限の設定（暗号の辞書の P）が許していないのに行った書き込み：[{"op": 操作の名前, "what": 説明}]（_permissions）
    denied: list[dict] = field(default_factory=list)


# ====================================================================== 増分更新
class Update:
    def __init__(self, doc: PdfDocument):
        self.doc = doc
        doc.verify_xref()  # 壊れた相互参照は、書く前に読み直す（足す相互参照に、すべての正しい位置を入れるため）
        # 新しい番号は、相互参照にある番号と、ファイルに書いてあるオブジェクトの番号（相互参照が壊れていて、後で読み直す
        # ことがある）の、どれとも重ならないようにする
        scanned = max((int(m[1]) for m in re.finditer(rb"(?<![\d])(\d{1,10})\s+\d{1,5}\s+obj\b", doc.blob)), default=0)
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
        out = bytearray(orig)
        if not out.endswith((b"\n", b"\r")):
            out += b"\n"
        offsets: dict[int, tuple[int, int]] = {}
        for num in sorted(self.objs):
            gen, o = self.objs[num]
            offsets[num] = (len(out), gen)
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
            offsets[xnum] = (len(out), 0)
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
            xref_pos = len(out)
            out += f"{xnum} 0 obj\n".encode() + dumps(d) + b"\nstream\n" + bytes(rows) + b"\nendstream\nendobj\n"
        else:
            xref_pos = len(out)
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
        return bytes(out)


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


def spot_fits(text: str, rect, size: float | None, kind: str) -> bool:
    """配置から見つけた欄（窓の rect・size）に、add_widget と同じ決め方で text が入るか。"""
    x0, y0, x1, y1 = rect
    w, h = x1 - x0, y1 - y0
    size = float(size or 10.5)
    if w <= 0 or h <= 0:
        return False
    try:
        layout_text(text, w, h, size, center=kind == "inline", multiline=h > 2.4 * min(size, 10.5))
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
        if multiline and len(lines) * s * 1.15 > h + 0.5:
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


def inline_align(spot: dict, text_w: float | None = None) -> str:
    """文字の中の空欄に書く字のそろえ方：空欄の前後の字で決める（見出しの字に寄せる）。
    前にだけ字がある（「〒［　］」「○○区［　］」）→ 左、後ろにだけ字がある（「［　］年」）→ 右（後ろの字の手前）、
    前後に字がある（「令和［　］年」）→ 中央。ただし、前後に字があっても、空欄が書く字（幅 text_w）よりずっと広い
    （3 倍と字 2 つぶんを超える：「区長宛て［　　　　　　］年」）ときは、中央に置くと前後のどちらの字からも離れて読めない
    ので、後ろの字の手前（右）にそろえる。覆って書く所（cover）は中央。"""
    if spot.get("cover"):
        return "center"
    lab = spot.get("label") or ""
    marks = list(re.finditer(r"［[^］]*］", lab))
    m = next((x for x in marks if "▶" in x.group()), marks[0] if len(marks) == 1 else None)
    if m is None:
        return "center"
    before = re.sub(r"［[^］]*］", "", lab[:m.start()]).strip(" 　")
    after = re.sub(r"［[^］]*］", "", lab[m.end():]).strip(" 　")
    # 後ろの字にくっつけてよいのは、後ろの字が単位のような短い字（区切りの前に 1〜2 字：「年」「月」「級」「円」「号」）の
    # ときだけ。ほかの字（「信用金庫」「（続柄）」）にくっつけると、書いた字がその語の一部に読める（「架空信用金庫」）ので、
    # 後ろの字との間を 1 字ぶん空ける（right_gap）
    right = "right" if unit_like(after) else "right_gap"
    if before and not after:
        return "left"
    if after and not before:
        return right
    if before and after and text_w is not None and spot.get("rect"):
        r = spot["rect"]
        size = float(spot.get("size") or 10.5)
        if (r[2] - r[0]) > 3 * text_w + 2 * size:
            return right
    return "center"


def unit_like(after: str) -> bool:
    """空欄の後ろの字が、単位のような短い字か：初めの区切り（空白・記号・行の端）までの字が 1〜2 字（「年」「月 日」
    「級」「円」「号」）。区切りの字で始まる（「（続柄）」）・3 字以上の語（「信用金庫」）は違う。"""
    from ..circlepick import sep

    k = 0
    while k < len(after) and not sep(after[k]):
        k += 1
    return 1 <= k <= 2


def text_appearance(update: Update, rect_w: float, rect_h: float, value: str, size: float, dr_fonts: dict, *,
                    center: bool = False, comb: list | None = None, rot: int = 0, multiline: bool = True,
                    cover: bool = False, align: str = "left", unit: float = 1.0) -> tuple[Ref, float]:
    """cover：欄の下地を白く塗ってから書く（伏せ字の「○○」の上に書くときに、下の字を覆う。元の字は消さない）。
    rot：欄の回転（/MK の R）。描く向きの座標（vw×vh）で字を並べ、BBox も同じ大きさにして、Matrix で Rect の向きに回す。"""
    vw, vh = _view_box(rot, rect_w, rect_h)
    bg = f"q\n1 g\n0 0 {vw:.2f} {vh:.2f} re\nf\nQ\n".encode() if cover else b""
    if comb:
        if glyphs.shaped(str(value)):
            raise FillError(f"「{value}」は、字を組み立てて（シェーピング）描く文字なので、1 字ずつの枠（comb）の欄には書けません")
        chars = list(str(value))
        if len(chars) > len(comb):
            raise FillError(f"「{value}」は {len(chars)} 字です（この欄は {len(comb)} 字まで）")
        s = min(size, min(b[3] - b[1] for b in comb) * 0.7)
        ops = bytearray(b"/Tx BMC\nq\n")
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
    lines, s = layout_text(str(value), vw, vh, size, center, multiline, unit)
    leading = s * 1.15
    block = leading * (len(lines) - 1) + s
    y = (vh + block) / 2 - s * 0.8
    pad = min(1.5, 0.1 * vw)
    align = "center" if center else align
    aw = vw - 2 * pad
    if align == "right_gap":
        # 後ろの字との間を 1 字ぶん空けて右にそろえる。空ける幅がなければ中央
        widest = max((text_width(ln_, s) for ln_ in lines), default=0.0)
        if widest + s <= aw:
            aw, align = aw - s, "right"
        else:
            align = "center"
    ops = bg + b"/Tx BMC\nq\n" + text_ops(lines, s, pad, y, leading, aw if align in ("center", "right") else None,
                                          align) + b"Q\nEMC\n"
    return _ap(update, vw, vh, ops, fonts_used(lines), dr_fonts, _rot_matrix(rot, rect_w, rect_h)), s


def check_appearance(update: Update, w: float, h: float, dr_fonts: dict, char: str = "4", rot: int = 0) -> Ref:
    """チェックの印の見た目。w・h は欄の Rect の幅と高さ、rot は欄の回転（/MK の R）：印も欄の向きに回して描く。"""
    vw, vh = _view_box(rot, w, h)
    s = min(vw, vh) * 0.8
    ops = f"q\nBT\n0 g\n/ZaDb {s:.2f} Tf\n{(vw - s * 0.85) / 2:.2f} {(vh - s * 0.75) / 2:.2f} Td\n({char}) Tj\nET\nQ\n".encode()
    return _ap(update, vw, vh, ops, {"ZaDb"}, dr_fonts, _rot_matrix(rot, w, h))


def empty_appearance(update: Update, w: float, h: float, rot: int = 0) -> Ref:
    vw, vh = _view_box(rot, w, h)
    return _ap(update, vw, vh, b"", set(), {}, _rot_matrix(rot, w, h))


# ====================================================================== 書き込み
def _as_text(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return "" if v is None else str(v)


def _as_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    t = str(v).strip().lower()
    if t in ("1", "true", "yes", "on", "x", "✓", "✔", "☑", "レ", "はい", "あり", "有", "○", "〇"):
        return True
    if t in ("0", "false", "no", "off", "", "いいえ", "なし", "無"):
        return False
    raise FillError(f"チェックの欄には、はい／いいえ（true / false）で書いてください（「{v}」）")


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
        from .signatures import read as read_signatures

        self.sigs = read_signatures(doc)  # 電子署名：固定された欄・認証の署名が許す変更

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
            elif key == "TdkE" or key.startswith("TdkX"):
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
            on = True if isinstance(value, str) and value in {st for _w, st in states} | {o[0] for o in f.options} \
                else _as_bool(value)  # 「入」の状態の名前・書き出しの値（/Ja など）で指してもよい
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
            if f.flags & FF_COMB and f.maxlen:
                cw = vw / f.maxlen
                comb = [(i * cw, 0, (i + 1) * cw, vh) for i in range(f.maxlen)]
            try:
                # 1 行の欄（Multiline の印なし）は折り返さない：見た目を作り直すビューアーも 1 行で描くので、1 行に収まる大きさにする
                ap, s = text_appearance(self.up, x1 - x0, y1 - y0, shown, size, self.fonts_for(keys | {"Helv"}),
                                        comb=comb, rot=rot, multiline=multiline, unit=unit,
                                        align={1: "center", 2: "right"}.get(f.q, "left"))  # /Q（0 左・1 中央・2 右）
            except FillError as e:
                raise FillError(f"{f.name}: {e}" + ("" if multiline or comb else "（1 行の欄なので、折り返して書けません）")) from None
            wd["AP"] = {"N": ap}
            self._set_da(fd, wd, da, s, keys, bool(shown))
            if shown and not calc:
                self._report_size(f.name, s, da_size, unit)
        p = Patch(f.name, "input", text)
        p.shown = shown
        self.patches.append(p)

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

    def _set_da(self, fd: dict, wd: dict, da: str, s: float, keys: set, drawn: bool) -> None:
        """見た目を作り直すビューアー（NeedAppearances・欄を選んだとき）が、描いた大きさと同じ大きさで描くように、/DA の字の
        大きさを、見た目に描いた大きさにする。元が自動（0）なら 0 のまま（ビューアーが欄に収まる大きさにする）。"""
        orig = _da_size(da)
        m = re.search(r"([\d.]+)\s+Tf", da or "")
        auto = m is not None and _num(m[1]) == 0
        if keys - {"Helv"}:
            # 日本語などは Helvetica では描けないので、欄の書体の指定も合わせる
            cjk = _da_key(keys)
            wd["DA"] = PdfString(f"/{cjk} {'0' if auto else _pt(s)} Tf 0 g".encode())
            fd["DA"] = wd["DA"]
            return
        if not drawn or auto or orig is None or abs(orig - s) < 0.005:
            return
        new = _da_resize(da, s)
        if new is not None:
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
        base = {"Type": Name("Annot"), "Subtype": Name("Widget"), "T": text_string(spot["id"]),
                "TU": text_string(spot.get("label") or spot["id"]), "F": 4, "P": page_ref}
        if k == "checkbox":
            on = _as_bool(value)
            yes = check_appearance(self.up, w, h, self.fonts_for({"ZaDb"}), rot=rot)  # 回したページでは、印も見た目の向きに
            off = empty_appearance(self.up, w, h, rot)
            d = {**base, "FT": Name("Btn"), "Rect": [round(x0, 3), round(y0, 3), round(x1, 3), round(y1, 3)],
                 "V": Name("Yes") if on else Name("Off"), "AS": Name("Yes") if on else Name("Off"),
                 "AP": {"N": {"Yes": yes, "Off": off}}, "MK": {"CA": PdfString(b"4"), **({"R": rot} if rot else {})},
                 "DA": PdfString(b"/ZaDb 0 Tf 0 g")}
            ref = self.up.new(d)
            self._attach(page_no, ref)
            self.patches.append(Patch(spot["id"], "control", on, page_no))
            return
        if isinstance(value, (list, tuple, dict, set)):
            # PDF の配置の所は、どれも 1 つの空き（1 つの値）：リストの字（「['本人']」）をそのまま書かない
            raise FillError(f"{spot['id']}: この所には 1 つの値（字か数）を書きます。リストや {{…}} の形の値は書けません"
                            f"（受け取った値：{value!r}）。空きが並ぶ所は、それぞれの id に 1 つずつ書いてください")
        text = _as_text(value)
        size = float(spot.get("size") or 10.5 / view.unit)
        comb = None
        if k == "comb":
            boxes = spot.get("box_rects") or []
            comb = []
            for b in boxes:
                bx0, by0, bx1, by1 = b
                comb.append((bx0 - vx0, by0 - vy0, bx1 - vx0, by1 - vy0))
        cover = bool(spot.get("cover"))
        unit = view.unit  # 知らせる字の大きさは、紙の上の pt（UserUnit のあるページも）
        # そろえ方：文字の中の空欄と、見出しに空欄の印（［　］）のある欄は、前後の字に寄せる（inline_align）
        marked = k == "inline" or "［" in (spot.get("label") or "")
        align = inline_align(spot, text_width(text, min(size, (vy1 - vy0) * 0.8)) if text else None) if marked else "left"
        ap, s = text_appearance(self.up, w, h, text, size, self.fonts_for(fonts_used([text]) | {"Helv"}),
                                align=align, comb=comb, rot=rot,
                                multiline=(vy1 - vy0) > 2.4 * min(size, 10.5 / unit), cover=cover, unit=unit)
        if cover:
            self.covers.add(spot["id"])
            self.warnings.append(f"{spot['id']}: 元の字「{spot.get('covered', '')}」は、白い下地の欄で覆いました（元の字は消していません）")
        if text and s * unit < MIN_READABLE_PT:
            self.warnings.append(f"{spot['id']}: 欄に収めるため、字が {s * unit:.1f} pt になり、印刷すると読みにくい大きさです"
                                 f"（{MIN_READABLE_PT:g} pt 未満）。もっと広い所に書くか、短く書けないか確かめてください")
        elif text and s < 0.6 * size:
            self.warnings.append(f"{spot['id']}: 欄に収めるため、字を {s * unit:.1f} pt に小さくしました（元の大きさ {size * unit:.1f} pt）")
        keys = fonts_used([text]) or {"Helv"}
        da_font = _da_key(keys)
        d = {**base, "FT": Name("Tx"), "Rect": [round(x0, 3), round(y0, 3), round(x1, 3), round(y1, 3)],
             "V": text_string(text), "DA": PdfString(f"/{da_font} {s:.1f} Tf 0 g".encode()), "AP": {"N": ap}}
        if comb:
            d["Ff"] = FF_COMB
            d["MaxLen"] = len(comb)
        if marked:
            d["Q"] = {"center": 1, "right": 2, "right_gap": 1}.get(align, 0)  # 作り直すソフトでは、くっつけない中央
        if rot:
            d["MK"] = {"R": rot}
        if cover:
            d.setdefault("MK", {})["BG"] = [1]  # 見た目を作り直すソフトでも、下地は白
        ref = self.up.new(d)
        self._attach(page_no, ref)
        self.patches.append(Patch(spot["id"], "input", text, page_no, shown=text))

    def _attach(self, page_no: int, ref: Ref, field: bool = True) -> None:
        self.page_annots.setdefault(page_no, []).append(ref)
        if field:
            self.new_fields.append(ref)

    # ------------------------------------------------------------------ 仕上げ
    def _draw_circles(self) -> None:
        """○：ページの描画の前に「q」、後ろに楕円のストリームを足す（元のストリームは変えない）。"""
        from . import circle as _circle

        for page_no, circles in sorted(self.circles.items()):
            page = self.pages[page_no - 1]
            pref = page["_ref"]
            depth = _circle.open_depth(self.doc.page_content(page))
            data = _circle.ops(View(self.doc, page), circles, depth)
            pre = self.up.new(Stream({}, _circle.PREFIX))
            post = self.up.new(Stream({}, data))
            pd = self.up.get(pref)
            cur = self.doc.resolve(pd.get("Contents"))
            if isinstance(cur, list):
                old = list(cur)
            elif pd.get("Contents") is None:
                old = []
            else:
                old = [pd.get("Contents")]
            pd["Contents"] = [pre] + old + [post]

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
        for k in list(self.dr_fonts):
            if k == "TdkE" or k.startswith("TdkX"):
                self._embed_font(k)
        if self.circles:
            self._draw_circles()
        for page_no, refs in self.page_annots.items():
            pref = self.pages[page_no - 1]["_ref"]
            pd = up.get(pref)
            annots = pd.get("Annots")
            if isinstance(annots, Ref):
                arr = up.get(annots)
                arr.extend(refs)
            else:
                pd["Annots"] = list(annots or []) + refs
        if self.new_fields or self.dr_fonts:
            root_ref = doc.trailer["Root"]
            af_ref = doc.root.get("AcroForm")
            if isinstance(af_ref, Ref):
                af = up.get(af_ref)
            elif isinstance(af_ref, dict):
                cat = up.get(root_ref)
                af = cat["AcroForm"]
            else:
                af = {"Fields": [], "DA": PdfString(b"/Helv 0 Tf 0 g")}
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
                    fonts.setdefault(k, r)
            if "DA" not in af:
                af["DA"] = PdfString(b"/Helv 0 Tf 0 g")
        return up.build()


def _da_key(keys: set[str]) -> str:
    """欄の書体の指定（/DA）にする書体：見た目を作り直すソフトが使うので、日本語などの書体を先に（欧文の拡張の TdkL は最後）。"""
    for k in ("TdkJ", "TdkK", "TdkG", "TdkC", "TdkE"):
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
    with glyphs.cjk_order(lang):
        return _fill(doc, values, force=force, placement=placement)


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
    circled_spans = None  # 同じ書き込みで○で囲む字の所（ページ → [(上, 下)]）：字の行が縦に並ぶマスの欄の高さを決める
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
                    f.patches.append(Patch(sd["id"], "circle", sd["text"], sd["page"], anchor=dict(value), key=key, circle=sd))
                    continue
                try:
                    sd = resolve_write(doc, value, used, layouts, set(f.by_name))
                except AnchorError as e:
                    raise FillError(f"{key}: {e}") from None
                page_no = sd["page"]
                if page_no not in views:
                    views[page_no] = View(doc, f.pages[page_no - 1])
                f.add_widget(page_no, sd, value["value"], views[page_no])
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
                raise FillError(f"{key}: そのような欄はありません")
            page_no, sd = spots[key]
            if page_no not in views:
                views[page_no] = View(doc, f.pages[page_no - 1])
            sd, value = as_kind(sd, value)
            if circled_spans is None:
                circled_spans = _circled_spans(doc, values, layouts)
            sd = _stacked_band(f, doc, page_no, sd, value, layouts, circled_spans)
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
    if f.patches and _need_appearances(doc):
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
    return FillResult(blob, f.patches, f.warnings, denied=denied)


def _circled_spans(doc, values: dict, layouts: dict) -> dict:
    """同じ書き込みで○で囲む字の所：ページ → [(上, 下)]（見た目の座標）。決まらない○は、ここでは除く（あとで断る）。"""
    from . import circle as _circle
    from .anchor import AnchorError

    out: dict = {}
    for _k, v in values.items():
        if not (isinstance(v, dict) and "circle" in v):
            continue
        try:
            n, _cell, ln, a, b = _circle.find(doc, v, layouts)
        except (AnchorError, Exception):  # noqa: BLE001  決まらない○は、書くときに理由を添えて断る
            continue
        r = ln.span_rect(a, b)
        out.setdefault(n, []).append((r[1], r[3]))
    return out


def _stacked_band(f: Filler, doc, page_no: int, sd: dict, value, layouts: dict, circled: dict) -> dict:
    """窓の欄（字の欄）のすぐ右か左に、印刷された字の行が縦に 2 つ以上並ぶ（「銀行／信用金庫／信用組合／農協」）とき、
    欄の縦の中ほどに書くと、たまたまその高さにある行（「信用組合」）に続けて読める。そこで書く高さを決める：
    1) 同じ書き込みで、並んだ行のどれかの字を○で囲むなら、その行の高さ
    2) なければ、並んだ行のどれの高さにもかからない、字の入る空き（帯）
    3) どちらもなければ、元のまま書いて、どの行に続けて読めるかを伝える"""
    if sd.get("kind") != "text" or sd.get("source") == "acroform" or not sd.get("rect") or "\n" in str(value):
        return sd
    from . import circle as _circle

    lay = _circle.layout(doc, page_no, layouts)
    x0, y0, x1, y1 = sd["rect"]
    size = float(sd.get("size") or 10.5)
    reach = 3 * size
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    home = next((c for c in lay.cells if c.x0 <= cx <= c.x1 and c.y0 <= cy <= c.y1), None)  # 欄のあるマス（同じマスの行だけ）
    rows = []
    for cell_, ln in _circle._lines(lay):
        if ln.vertical or not ln.text.strip() or ln.y0 < y0 - 0.5 or ln.y1 > y1 + 0.5:
            continue
        if home is not None and cell_ is not home:
            continue
        right = ln.x0 >= x1 - 1 and ln.x0 - x1 <= reach
        left = ln.x1 <= x0 + 1 and x0 - ln.x1 <= reach
        if right or left:
            rows.append(ln)
    rows.sort(key=lambda ln: ln.y0)
    distinct = []
    for ln in rows:
        if not distinct or ln.y0 >= distinct[-1].y1 - 0.2 * ln.size:
            distinct.append(ln)
    if len(distinct) < 2:
        return sd
    def _short(ln) -> str:
        return re.sub(r"\s+", "", ln.text)[:8]

    names = "／".join(_short(ln) for ln in distinct[:6])

    def band(ln):
        return (max(y0, ln.y0 - 0.15 * ln.size), min(y1, ln.y1 + 0.15 * ln.size))

    picked = [ln for ln in distinct if any(a < ln.y1 and b > ln.y0 for a, b in circled.get(page_no, []))]
    if len(picked) == 1:
        b0, b1 = band(picked[0])
        f.warnings.append(f"{sd['id']}: 横に字の行が {len(distinct)} 行並ぶ所（{names}）なので、○で囲んだ"
                          f"「{_short(picked[0])}」の行の高さに書きました")
        return {**sd, "rect": [x0, b0, x1, b1]}
    # 並んだ行の高さにかからない空き
    need = 0.9 * min(size, 10.5)
    edges = [y0] + [v for ln in distinct for v in (ln.y0 - 0.1 * ln.size, ln.y1 + 0.1 * ln.size)] + [y1]
    gaps = [(edges[i], edges[i + 1]) for i in range(0, len(edges) - 1, 2) if edges[i + 1] - edges[i] >= need]
    if gaps:
        g0, g1 = max(gaps, key=lambda g: g[1] - g[0])
        where = "上" if g1 <= distinct[0].y0 else "下" if g0 >= distinct[-1].y1 else "あいだ"
        f.warnings.append(f"{sd['id']}: 横に字の行が {len(distinct)} 行並ぶ所（{names}）なので、どの行の字にも続けて読めないよう、"
                          f"行の高さにかからない{where}の空きに書きました。選ぶ行の字を同じ書き込みで○にすると、その行の高さに書きます")
        return {**sd, "rect": [x0, g0, x1, g1]}
    mid = (y0 + y1) / 2
    level = next((ln for ln in distinct if ln.y0 <= mid <= ln.y1), None)
    f.warnings.append(f"{sd['id']}: 横に字の行が {len(distinct)} 行並ぶ所（{names}）で、どの行の高さにもかからない空きが"
                      "ありません。書いた字は欄の縦の中ほど" + (f"（「{_short(level)}」の行の高さ）" if level else "")
                      + "に置いたので、その行の字に続けて読めることがあります。選ぶ行の字を同じ書き込みで○にすると、その行の高さに"
                      "書きます。本人に確かめてください")
    return sd


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
    out = []
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
    - {"as": 種類, "value": 値}：その種類で書く（種類は kind_candidates のどれか）
    - はい／いいえ（true / false）：チェックの印（候補に checkbox があるとき）
    - 字：候補に text があり、既定の種類がチェックの印なら、文字の欄として書く（「1」「レ」などをマスに書く）。
      既定が1字ずつ書く欄（comb）なら、1字ずつマスに書く（1つの欄として書くときは {"as": "text", …}）
    返す値：(書く欄の情報, 書く値)。"""
    cands = sd.get("kind_candidates") or []
    if isinstance(value, dict) and "as" in value:
        want = value.get("as")
        if not cands or want not in cands:
            raise FillError(f"{sd['id']}: 「{want}」としては書けません（書ける種類：{'・'.join(cands) or sd['kind']}）")
        if "value" not in value:
            raise FillError(f'{sd["id"]}: {{"as": 種類, "value": 書く値}} の形で指定してください')
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
            # 静的なフォーム：PDF の欄の値（決まった形）と見た目（書式で整えた字）も同じにする
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
