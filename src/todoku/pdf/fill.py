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
from .fonts import _HELV, cid_inverse
from .form import FF_COMB, FF_RADIO, FormField, read_fields
from .layout import View
from .objects import Name, PdfError, PdfString, Ref, Stream, dumps, text_string


class FillError(ValueError):
    pass


@dataclass
class Patch:
    id: str
    kind: str  # input / control
    value: object
    page: int = 0
    anchor: dict | None = None  # 見出しの字を目印にした書き込み（{"after": …, "value": …}）。証明で、同じ決め方で位置を確かめる
    key: str = ""  # 指定したときの名前（「@1」など）
    seal: dict | None = None  # 本人の印影・署名の画像を置いたとき：本人に伝える文・位置・画像の SHA-256


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unverified: list[str] = field(default_factory=list)


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
CJK = {  # 字形の集まり → (資源の名前, 書体, CMap, 補遺)
    "Adobe-Japan1": ("TdkJ", "HeiseiKakuGo-W5", "UniJIS-UCS2-HW-H", 2),
    "Adobe-Korea1": ("TdkK", "HYGoThic-Medium", "UniKS-UCS2-H", 1),
    "Adobe-GB1": ("TdkG", "STSong-Light", "UniGB-UCS2-H", 2),
    "Adobe-CNS1": ("TdkC", "MSung-Light", "UniCNS-UCS2-H", 0),
}


def cjk_font(ros: str) -> dict:
    res, base, cmap, sup = CJK[ros]
    reg, order = ros.split("-")
    desc = {"Type": Name("FontDescriptor"), "FontName": Name(base), "Flags": 4, "FontBBox": [-100, -250, 1100, 900],
            "ItalicAngle": 0, "Ascent": 880, "Descent": -120, "CapHeight": 700, "StemV": 80}
    cid = {"Type": Name("Font"), "Subtype": Name("CIDFontType0"), "BaseFont": Name(base),
           "CIDSystemInfo": {"Registry": PdfString(reg.encode()), "Ordering": PdfString(order.encode()), "Supplement": sup},
           "FontDescriptor": desc, "DW": 1000}
    if ros == "Adobe-Japan1":
        cid["W"] = [231, 389, 500]  # 半角（英数字・カタカナ）は 500
    return {"Type": Name("Font"), "Subtype": Name("Type0"), "BaseFont": Name(f"{base}-{cmap}"), "Encoding": Name(cmap),
            "DescendantFonts": [cid]}


def _latin(ch: str) -> bool:
    try:
        ch.encode("cp1252")
        return True
    except UnicodeEncodeError:
        return False


def runs_for(text: str) -> list[tuple[str, str]]:
    """文字列を、同じ書体で描ける並びに分ける：(資源の名前, 文字列)。"""
    out: list[tuple[str, str]] = []
    for ch in text:
        if ord(ch) > 0xFFFF:
            raise FillError(f"「{ch}」（U+{ord(ch):X}）は、この書き込みでは描けない字です（基本多言語面の外の字）")
        if _latin(ch) and not ("＀" <= ch <= "￯"):
            key = "Helv"
        elif ch in cid_inverse("Adobe-Japan1"):
            key = "TdkJ"
        elif "가" <= ch <= "힯" or "ᄀ" <= ch <= "ᇿ" or "㄰" <= ch <= "㆏":
            key = "TdkK"
        elif ch in cid_inverse("Adobe-GB1"):
            key = "TdkG"
        elif ch in cid_inverse("Adobe-CNS1"):
            key = "TdkC"
        else:
            raise FillError(f"「{ch}」（U+{ord(ch):04X}）を描ける字形の集まりがありません")
        if out and out[-1][0] == key:
            out[-1] = (key, out[-1][1] + ch)
        else:
            out.append((key, ch))
    return out


def text_width(text: str, size: float) -> float:
    w = 0.0
    for key, run in runs_for(text):
        for ch in run:
            if key == "Helv":
                code = ch.encode("cp1252")[0]
                w += (_HELV[code - 32] if 32 <= code <= 126 else 556) / 1000.0
            elif key == "TdkJ" and ord(ch) < 0x80:
                w += 0.5
            elif "｡" <= ch <= "ﾟ":
                w += 0.5
            else:
                w += 1.0
    return w * size


def _show(key: str, run: str) -> bytes:
    if key == "Helv":
        data = run.encode("cp1252")
    else:
        data = run.encode("utf-16-be")
    return b"<" + data.hex().upper().encode() + b"> Tj"


def text_ops(lines: list[str], size: float, x: float, y: float, leading: float, align_w: float | None = None) -> bytes:
    """文字の行を描く命令。x, y は1行目の基線の始まり。align_w があれば、その幅の中で中央にそろえる。"""
    out = bytearray(b"BT\n0 g\n")
    for i, line in enumerate(lines):
        lx = x
        if align_w is not None:
            lx = x + max(0.0, (align_w - text_width(line, size)) / 2)
        ly = y - i * leading
        out += f"1 0 0 1 {lx:.2f} {ly:.2f} Tm\n".encode()
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
        for ch in para:
            if cur and text_width(cur + ch, size) > width:
                # 欧文は単語の切れ目で折り返す
                k = cur.rfind(" ")
                if _latin(ch) and ch != " " and k > 0:
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


def layout_text(text: str, w: float, h: float, size: float, center: bool = False, multiline: bool = True):
    """枠に入るように、字の大きさと折り返しを決める。(行, 大きさ)。入らなければ FillError。"""
    s = size
    pad = min(3.0, 0.2 * w)
    while s >= 4.0:
        lines = wrap(text, s, w - pad) if multiline else [str(text)]
        if multiline and len(lines) * s * 1.15 > h + 0.5:
            s -= 0.5
            continue
        if all(text_width(line, s) <= w - pad for line in lines):
            return lines, s
        s -= 0.5
    raise FillError(f"「{text}」は、欄（幅 {w:.0f}pt・高さ {h:.0f}pt）に入りきりません")


# ====================================================================== 見た目のストリーム
def _ap(update: Update, w: float, h: float, ops: bytes, fonts: set[str], dr_fonts, matrix=None) -> Ref:
    res = {}
    if fonts:
        res["Font"] = {k: (dr_fonts(k) if callable(dr_fonts) else dr_fonts[k]) for k in sorted(fonts)}
    d = {"Type": Name("XObject"), "Subtype": Name("Form"), "BBox": [0, 0, round(w, 3), round(h, 3)], "Resources": res}
    if matrix is not None:
        d["Matrix"] = list(matrix)
    return update.new(Stream(d, ops))


def _rot_matrix(rot: int, w: float, h: float):
    """ページの回転を打ち消して、見た目の向きで描くための行列。"""
    if rot == 90:
        return [0, 1, -1, 0, h, 0]
    if rot == 180:
        return [-1, 0, 0, -1, w, h]
    if rot == 270:
        return [0, -1, 1, 0, 0, w]
    return None


def text_appearance(update: Update, rect_w: float, rect_h: float, value: str, size: float, dr_fonts: dict, *,
                    center: bool = False, comb: list | None = None, rot: int = 0, multiline: bool = True,
                    cover: bool = False) -> tuple[Ref, float]:
    """cover：欄の下地を白く塗ってから書く（伏せ字の「○○」の上に書くときに、下の字を覆う。元の字は消さない）。"""
    vw, vh = (rect_h, rect_w) if rot in (90, 270) else (rect_w, rect_h)
    bg = f"q\n1 g\n0 0 {vw:.2f} {vh:.2f} re\nf\nQ\n".encode() if cover else b""
    if comb:
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
        return _ap(update, rect_w, rect_h, bytes(ops), fonts_used(chars), dr_fonts, _rot_matrix(rot, rect_w, rect_h)), s
    lines, s = layout_text(str(value), vw, vh, size, center, multiline)
    leading = s * 1.15
    block = leading * (len(lines) - 1) + s
    y = (vh + block) / 2 - s * 0.8
    pad = min(1.5, 0.1 * vw)
    ops = bg + b"/Tx BMC\nq\n" + text_ops(lines, s, pad, y, leading, (vw - 2 * pad) if center else None) + b"Q\nEMC\n"
    return _ap(update, rect_w, rect_h, ops, fonts_used(lines), dr_fonts, _rot_matrix(rot, rect_w, rect_h)), s


def check_appearance(update: Update, w: float, h: float, dr_fonts: dict, char: str = "4") -> Ref:
    s = min(w, h) * 0.8
    ops = f"q\nBT\n0 g\n/ZaDb {s:.2f} Tf\n{(w - s * 0.85) / 2:.2f} {(h - s * 0.75) / 2:.2f} Td\n({char}) Tj\nET\nQ\n".encode()
    return _ap(update, w, h, ops, {"ZaDb"}, dr_fonts)


def ellipse_appearance(update: Update, w: float, h: float) -> Ref:
    k = 0.5523
    rx, ry = w / 2 - 0.8, h / 2 - 0.8
    cx, cy = w / 2, h / 2
    ops = (f"q\n0 g\n0.9 w\n{cx + rx:.2f} {cy:.2f} m\n"
           f"{cx + rx:.2f} {cy + k * ry:.2f} {cx + k * rx:.2f} {cy + ry:.2f} {cx:.2f} {cy + ry:.2f} c\n"
           f"{cx - k * rx:.2f} {cy + ry:.2f} {cx - rx:.2f} {cy + k * ry:.2f} {cx - rx:.2f} {cy:.2f} c\n"
           f"{cx - rx:.2f} {cy - k * ry:.2f} {cx - k * rx:.2f} {cy - ry:.2f} {cx:.2f} {cy - ry:.2f} c\n"
           f"{cx + k * rx:.2f} {cy - ry:.2f} {cx + rx:.2f} {cy - k * ry:.2f} {cx + rx:.2f} {cy:.2f} c\nS\nQ\n").encode()
    return _ap(update, w, h, ops, set(), {})


def empty_appearance(update: Update, w: float, h: float) -> Ref:
    return _ap(update, w, h, b"", set(), {})


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

    # ------------------------------------------------------------------ 共通の資源
    def font(self, key: str) -> Ref:
        if key not in self.dr_fonts:
            existing = self._existing_dr_font(key)
            if existing is not None:
                self.dr_fonts[key] = existing
                return existing
            if key == "Helv":
                self.dr_fonts[key] = self.up.new(dict(HELV))
            elif key == "ZaDb":
                self.dr_fonts[key] = self.up.new(dict(ZADB))
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
            if base in ("Helvetica", "ZapfDingbats"):
                return ref
        return None

    # ------------------------------------------------------------------ フォームの欄
    def set_field(self, f: FormField, value, force: bool, display: str | None = None) -> None:
        """display：見た目に描く字（XFA の静的なフォームで、値は決まった形・見た目は書式で整えた字のとき）。"""
        if f.readonly and not force:
            raise FillError(f"{f.name}: 読み取り専用の欄です")
        k = f.kind
        if k == "signature":
            raise FillError(f"{f.name}: 電子署名の欄には、文字を書き込めません")
        if k == "button":
            raise FillError(f"{f.name}: 押すだけのボタンで、値を持ちません")
        fd = self.up.get(f.ref) if f.ref is not None else None
        if fd is None:
            raise FillError(f"{f.name}: 欄のオブジェクトが見つかりません")
        if k == "checkbox":
            on = _as_bool(value)
            state = None
            for w in f.widgets:
                if w.on:
                    state = w.on
            if on and state is None:
                raise FillError(f"{f.name}: チェックの「入」の見た目がありません")
            fd["V"] = Name(state) if on else Name("Off")
            for w in f.widgets:
                wd = self.up.get(w.ref) if w.ref is not None and w.ref != f.ref else fd
                wd["AS"] = Name(w.on) if (on and w.on) else Name("Off")
            self.patches.append(Patch(f.name, "control", on))
            return
        if k == "radio":
            want = _as_text(value)
            states = [(w, w.on) for w in f.widgets if w.on]
            labels = {}
            for w, st in states:
                from .form import name_text

                lab = name_text(st)
                if f.options and st.isdigit() and int(st) < len(f.options):
                    lab = f.options[int(st)][1] or f.options[int(st)][0]
                labels[st] = lab
            pick = next((st for st, lab in labels.items() if lab == want or st == want), None)
            if pick is None:
                raise FillError(f"{f.name}: 「{want}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in labels.values()))
            fd["V"] = Name(pick)
            for w, st in states:
                wd = self.up.get(w.ref) if w.ref is not None and w.ref != f.ref else fd
                wd["AS"] = Name(st) if st == pick else Name("Off")
            self.patches.append(Patch(f.name, "control", labels[pick]))
            return
        text = _as_text(value)
        from . import acrojs

        js = acrojs.actions(self.doc, self.doc.resolve(f.ref)) if f.ref is not None else {}
        fmt_js = {kk: v for kk, v in js.items() if kk in ("F", "K", "V")}
        if fmt_js and display is None:
            if all(acrojs.known(v) for v in fmt_js.values()):
                try:
                    text = acrojs.commit(fmt_js, text)
                except acrojs.JsError as e:
                    raise FillError(f"{f.name}: {e}") from None
            else:
                self.warnings.append(f"{f.name}: 独自のスクリプト（書式・入力の制限）は実行していません")
        if k in ("combo", "list"):
            opts = [o[1] or o[0] for o in f.options]
            exports = [o[0] for o in f.options]
            if text not in opts and text not in exports and not (f.flags & (1 << 18)):
                raise FillError(f"{f.name}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in opts[:30]))
            if text in opts and text not in exports:
                text = exports[opts.index(text)]
        if f.maxlen and len(text) > f.maxlen:
            raise FillError(f"{f.name}: 「{text}」は {len(text)} 字です（この欄は {f.maxlen} 字まで）")
        fd["V"] = text_string(text)
        shown = text if display is None else display
        keys = fonts_used([shown]) if shown else set()
        for w in f.widgets:
            wd = self.up.get(w.ref) if w.ref is not None and w.ref != f.ref else fd
            x0, y0, x1, y1 = w.rect
            size = _da_size(f.da) or max(6.0, min(12.0, (y1 - y0) * 0.65))
            comb = None
            if f.flags & FF_COMB and f.maxlen:
                cw = (x1 - x0) / f.maxlen
                comb = [(x0 + i * cw - x0, 0, x0 + (i + 1) * cw - x0, y1 - y0) for i in range(f.maxlen)]
            rot = int(self.doc.resolve((self.doc.resolve(wd.get("MK")) or {}).get("R")) or 0) if isinstance(wd, dict) else 0
            ap, s = text_appearance(self.up, x1 - x0, y1 - y0, shown, size, self.fonts_for(keys | {"Helv"}),
                                    comb=comb, rot=rot, multiline=bool(f.flags & (1 << 12)) or (y1 - y0) > 2.4 * size,
                                    center=f.q == 1)
            wd["AP"] = {"N": ap}
            if keys - {"Helv"}:
                # 日本語などは Helvetica では描けないので、欄の書体の指定も合わせる
                cjk = sorted(keys - {"Helv"})[0]
                wd["DA"] = PdfString(f"/{cjk} {s:.1f} Tf 0 g".encode())
                fd["DA"] = wd["DA"]
        self.patches.append(Patch(f.name, "input", text))

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
            yes = check_appearance(self.up, w, h, self.fonts_for({"ZaDb"}))
            off = empty_appearance(self.up, w, h)
            d = {**base, "FT": Name("Btn"), "Rect": [round(x0, 3), round(y0, 3), round(x1, 3), round(y1, 3)],
                 "V": Name("Yes") if on else Name("Off"), "AS": Name("Yes") if on else Name("Off"),
                 "AP": {"N": {"Yes": yes, "Off": off}}, "MK": {"CA": PdfString(b"4")},
                 "DA": PdfString(b"/ZaDb 0 Tf 0 g")}
            ref = self.up.new(d)
            self._attach(page_no, ref)
            self.patches.append(Patch(spot["id"], "control", on, page_no))
            return
        if k == "circle":
            want = _as_text(value)
            opts = spot.get("option_rects") or []
            names = [o for o, _ in opts]
            if want not in names:
                raise FillError(f"{spot['id']}: 「{want}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in names))
            kids = []
            parent = {"FT": Name("Btn"), "Ff": FF_RADIO | (1 << 14), "T": text_string(spot["id"]),
                      "TU": text_string(spot.get("label") or spot["id"]),
                      "Opt": [text_string(o) for o in names], "V": Name(str(names.index(want)))}
            pref = self.up.new(parent)
            for i, (o, r) in enumerate(opts):
                pad = 2.0
                ox0, oy0, ox1, oy1 = view.back_rect(r[0] - pad, r[1] - pad, r[2] + pad, r[3] + pad)
                ow, oh = ox1 - ox0, oy1 - oy0
                on = ellipse_appearance(self.up, ow, oh)
                off = empty_appearance(self.up, ow, oh)
                st = Name(str(i))
                kd = {"Type": Name("Annot"), "Subtype": Name("Widget"), "Parent": pref, "F": 4, "P": page_ref,
                      "Rect": [round(ox0, 3), round(oy0, 3), round(ox1, 3), round(oy1, 3)],
                      "AS": st if o == want else Name("Off"), "AP": {"N": {str(i): on, "Off": off}}}
                kref = self.up.new(kd)
                kids.append(kref)
                self._attach(page_no, kref, field=False)
            parent["Kids"] = kids
            self.new_fields.append(pref)
            self.patches.append(Patch(spot["id"], "control", want, page_no))
            return
        text = _as_text(value)
        size = float(spot.get("size") or 10.5)
        comb = None
        if k == "comb":
            boxes = spot.get("box_rects") or []
            comb = []
            for b in boxes:
                bx0, by0, bx1, by1 = b
                comb.append((bx0 - vx0, by0 - vy0, bx1 - vx0, by1 - vy0))
        cover = bool(spot.get("cover"))
        ap, s = text_appearance(self.up, w, h, text, size, self.fonts_for(fonts_used([text]) | {"Helv"}),
                                center=k in ("inline",), comb=comb, rot=rot, multiline=(vy1 - vy0) > 2.4 * min(size, 10.5),
                                cover=cover)
        if cover:
            self.warnings.append(f"{spot['id']}: 元の字「{spot.get('covered', '')}」は、白い下地の欄で覆いました（元の字は消していません）")
        if text and s < 0.6 * size:
            self.warnings.append(f"{spot['id']}: 欄に収めるため、字を {s:.1f} pt に小さくしました（元の大きさ {size:.1f} pt）")
        keys = fonts_used([text]) or {"Helv"}
        da_font = sorted(keys - {"Helv"})[0] if keys - {"Helv"} else "Helv"
        d = {**base, "FT": Name("Tx"), "Rect": [round(x0, 3), round(y0, 3), round(x1, 3), round(y1, 3)],
             "V": text_string(text), "DA": PdfString(f"/{da_font} {s:.1f} Tf 0 g".encode()), "AP": {"N": ap}}
        if comb:
            d["Ff"] = FF_COMB
            d["MaxLen"] = len(comb)
        if k == "inline":
            d["Q"] = 1
        if rot:
            d["MK"] = {"R": rot}
        if cover:
            d.setdefault("MK", {})["BG"] = [1]  # 見た目を作り直すソフトでも、下地は白
        ref = self.up.new(d)
        self._attach(page_no, ref)
        self.patches.append(Patch(spot["id"], "input", text, page_no))

    def _attach(self, page_no: int, ref: Ref, field: bool = True) -> None:
        self.page_annots.setdefault(page_no, []).append(ref)
        if field:
            self.new_fields.append(ref)

    # ------------------------------------------------------------------ 仕上げ
    def finish(self) -> bytes:
        doc = self.doc
        up = self.up
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


def _da_size(da: str) -> float | None:
    m = re.search(r"([\d.]+)\s+Tf", da or "")
    if m:
        v = float(m[1])
        return v if v > 0 else None
    return None


def fill(doc: PdfDocument, values: dict, *, force: bool = False) -> FillResult:
    """欄の名前（窓の［］の中）→ 値。フォームの欄はその名前、配置から見つけた欄は p{ページ}.{番号}。"""
    from .window import build as build_window

    if _certified_no_changes(doc):
        raise FillError("この PDF は「変更を許さない」電子署名（認証）がされています。書き込むと署名が無効になります")
    f = Filler(doc)
    win = build_window(doc)
    spots = {}
    views = {}
    for pg in win["pages"]:
        for fd in pg["fields"]:
            spots[fd["id"]] = (pg["number"], fd)
    problems = []
    used: dict = {}
    layouts: dict = {}
    xf = _xfa_form(doc)
    xfa_values = {k: v for k, v in values.items() if xf is not None and k in xf.by_id}
    for key, value in values.items():
        if key in xfa_values:
            continue
        try:
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
            if key in f.by_name:
                f.set_field(f.by_name[key], value, force)
                continue
            if key not in spots:
                raise FillError(f"{key}: そのような欄はありません")
            page_no, sd = spots[key]
            if sd.get("source") == "acroform":
                f.set_field(f.by_name[sd["name"]], value, force)
                continue
            if page_no not in views:
                views[page_no] = View(doc, f.pages[page_no - 1])
            sd, value = as_kind(sd, value)
            f.add_widget(page_no, sd, value, views[page_no])
        except FillError as e:
            problems.append(str(e))
    problems += alternative_conflicts([spots[k][1] for k in values if k in spots and spots[k][1].get("source") == "layout"])
    if xfa_values and not problems:
        try:
            _fill_xfa(f, xf, xfa_values, force)
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    if f.patches and not xfa_values:
        _recalculate(f, values)
    blob = f.finish() if f.patches else doc.blob
    return FillResult(blob, f.patches, f.warnings)


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
    """同じ所の別の読み方（alternative_to でつながる候補）の両方に書こうとしたもの：どちらか1つにする（重なって書かれるので）。"""
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
    p.seal = {"told": seal.told(sd, skey), "rect": sd["rect"], "sha256": img.sha256, "what": sd["what"], "key": skey}
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
                    f.set_field(acro, val, True, display=disp)
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
        current[p.id] = p.value
    for _i, fld, js in sorted(calc, key=lambda x: x[0]):
        r = acrojs.simple_calculate(js["C"], lambda n: current.get(n))
        if r is None:
            f.warnings.append(f"{fld.name}: 独自の計算のスクリプトは実行していません（計算の結果は、Acrobat で開いたときに決まります）")
            continue
        val = int(r) if float(r).is_integer() else round(r, 10)
        if str(current.get(fld.name)) == str(val):
            continue
        try:
            f.set_field(fld, val, force=True)
        except FillError as e:
            f.warnings.append(f"{fld.name}: 計算した値 {val} を書けません（{e}）")
            continue
        current[fld.name] = f.patches[-1].value
        f.patches[-1].kind = "calc"


def _certified_no_changes(doc: PdfDocument) -> bool:
    perms = doc.resolve(doc.root.get("Perms"))
    if not isinstance(perms, dict):
        return False
    sig = doc.resolve(perms.get("DocMDP"))
    if not isinstance(sig, dict):
        return False
    for ref in doc.resolve(sig.get("Reference")) or []:
        r = doc.resolve(ref)
        tp = doc.resolve((r or {}).get("TransformParams")) or {}
        if int(doc.resolve(tp.get("P")) or 2) == 1:
            return True
    return False
