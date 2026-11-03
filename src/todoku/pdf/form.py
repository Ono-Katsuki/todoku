"""PDFのフォーム（AcroForm）の欄を読む。

欄の名前は、親の欄の名前を「.」でつないだ完全な名前（例 topmostSubform[0].Page1[0].f1_01[0]）。
値・種類（文字・チェック・ラジオ・一覧・署名）・選択肢・文字数の上限・読み取り専用を、継承も含めて読む。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .objects import Name, PdfString, Ref, Stream, decode_text

FF_READONLY = 1
FF_REQUIRED = 2
FF_MULTILINE = 1 << 12
FF_PASSWORD = 1 << 13
FF_NOTOGGLEOFF = 1 << 14
FF_RADIO = 1 << 15
FF_PUSH = 1 << 16
FF_COMBO = 1 << 17
FF_EDIT = 1 << 18
FF_MULTISELECT = 1 << 21
FF_COMB = 1 << 24
FF_RADIOSINUNISON = 1 << 25


@dataclass
class Widget:
    ref: Ref | None
    page: int  # 0 から
    rect: tuple[float, float, float, float]  # ページの座標
    on: str | None  # チェック・ラジオの「入」の状態の名前
    state: str | None  # いまの見た目の状態（/AS）
    hidden: bool = False
    mk_caption: str = ""


@dataclass
class FormField:
    name: str
    ftype: str  # Tx / Btn / Ch / Sig
    flags: int
    value: object
    tu: str
    options: list[tuple[str, str]]  # (書き出す値, 見せる文字)
    maxlen: int | None
    widgets: list[Widget] = field(default_factory=list)
    ref: Ref | None = None
    da: str = ""
    q: int = 0

    @property
    def kind(self) -> str:
        if self.ftype == "Btn":
            if self.flags & FF_PUSH:
                return "button"
            if self.flags & FF_RADIO:
                return "radio"
            return "checkbox"
        if self.ftype == "Ch":
            return "combo" if self.flags & FF_COMBO else "list"
        if self.ftype == "Sig":
            return "signature"
        return "text"

    @property
    def readonly(self) -> bool:
        return bool(self.flags & FF_READONLY)

    @property
    def required(self) -> bool:
        return bool(self.flags & FF_REQUIRED)

    def value_text(self):
        v = self.value
        if isinstance(v, Name):
            if self.kind == "radio" and self.options and v.name.isdigit() and int(v.name) < len(self.options):
                return self.options[int(v.name)][1] or self.options[int(v.name)][0]
            return name_text(v.name)
        if isinstance(v, (bytes, PdfString)):
            return decode_text(v)
        if isinstance(v, list):
            return [decode_text(x) if isinstance(x, (bytes, PdfString)) else getattr(x, "name", x) for x in v]
        return v

    @property
    def checked(self) -> bool:
        v = self.value
        name = v.name if isinstance(v, Name) else (decode_text(v) if isinstance(v, (bytes, PdfString)) else None)
        return bool(name) and name != "Off"


def name_text(n: str) -> str:
    """名前（バイト列を latin-1 で読んだもの）を、UTF-8 として読めるならその文字にする。
    UTF-8 として読めない名前（Shift_JIS で作った名前など）は、字を推し量らず、PDF の名前の書き方（#82#CD）でバイト列のまま示す
    （name_unreadable で確かめられる。書き込みでは、この形の名前で選べる）。"""
    try:
        return n.encode("latin-1").decode("utf-8")
    except UnicodeEncodeError:
        return n
    except UnicodeDecodeError:
        return "".join(ch if 0x21 <= ord(ch) <= 0x7E and ch != "#" else f"#{ord(ch):02X}" for ch in n)


def name_unreadable(n: str) -> bool:
    """UTF-8 として読めない名前か。"""
    try:
        n.encode("latin-1").decode("utf-8")
        return False
    except UnicodeEncodeError:
        return False
    except UnicodeDecodeError:
        return True


def _txt(doc, v) -> str:
    v = doc.resolve(v)
    if isinstance(v, (bytes, PdfString)):
        return decode_text(v)
    if isinstance(v, Name):
        return v.name
    return "" if v is None else str(v)


def read_fields(doc) -> list[FormField]:
    af = doc.resolve(doc.root.get("AcroForm"))
    if not isinstance(af, dict):
        return []
    pages = doc.pages()
    annot_page: dict[int, int] = {}
    for i, p in enumerate(pages):
        for a in doc.resolve(p.get("Annots")) or []:
            if isinstance(a, Ref):
                annot_page[a.num] = i
    out: list[FormField] = []
    seen: set[int] = set()

    def walk(ref, parent_name: str, inh: dict, depth: int):
        if depth > 32:
            return
        node = doc.resolve(ref)
        if not isinstance(node, dict):
            return
        if isinstance(ref, Ref):
            if ref.num in seen:
                return
            seen.add(ref.num)
        t = _txt(doc, node.get("T")) if "T" in node else None
        name = parent_name if t is None else (f"{parent_name}.{t}" if parent_name else t)
        h = dict(inh)
        for k in ("FT", "Ff", "V", "DA", "Q", "Opt", "MaxLen"):
            if k in node:
                h[k] = doc.resolve(node[k])
        kids = doc.resolve(node.get("Kids")) or []
        field_kids = [k for k in kids if isinstance(doc.resolve(k), dict) and "T" in doc.resolve(k)]
        if field_kids:
            for k in kids:
                walk(k, name, h, depth + 1)
            return
        ft = getattr(h.get("FT"), "name", "")
        opts = []
        for o in doc.resolve(h.get("Opt")) or []:
            o = doc.resolve(o)
            if isinstance(o, list) and len(o) >= 2:
                opts.append((_txt(doc, o[0]), _txt(doc, o[1])))
            else:
                opts.append((_txt(doc, o), _txt(doc, o)))
        ml = h.get("MaxLen")
        f = FormField(name or f"（名前のない欄 {getattr(ref, 'num', '')}）", ft, int(h.get("Ff") or 0), h.get("V"),
                      _txt(doc, node.get("TU")), opts, int(ml) if isinstance(ml, (int, float)) else None,
                      ref=ref if isinstance(ref, Ref) else None, da=_txt(doc, h.get("DA")), q=int(h.get("Q") or 0))
        widgets = kids if kids else [ref]
        for w in widgets:
            wd = doc.resolve(w)
            if not isinstance(wd, dict):
                continue
            rect = [float(doc.resolve(x)) for x in doc.resolve(wd.get("Rect")) or [0, 0, 0, 0]]
            page = annot_page.get(w.num, -1) if isinstance(w, Ref) else -1
            if page < 0 and isinstance(wd.get("P"), Ref):
                page = next((i for i, p in enumerate(pages) if p.get("_ref") == wd["P"]), -1)
            ap = doc.resolve(wd.get("AP")) or {}
            n = doc.resolve(ap.get("N")) if isinstance(ap, dict) else None
            on = None
            if isinstance(n, dict):
                states = [k for k in n if k != "Off"]
                on = states[0] if states else None
            st = doc.resolve(wd.get("AS"))
            flags = int(doc.resolve(wd.get("F")) or 0)
            mk = doc.resolve(wd.get("MK")) or {}
            f.widgets.append(Widget(w if isinstance(w, Ref) else None, page,
                                    (min(rect[0], rect[2]), min(rect[1], rect[3]), max(rect[0], rect[2]), max(rect[1], rect[3])),
                                    on, st.name if isinstance(st, Name) else None, bool(flags & 2),
                                    _txt(doc, mk.get("CA")) if isinstance(mk, dict) else ""))
        out.append(f)

    for r in doc.resolve(af.get("Fields")) or []:
        walk(r, "", {}, 0)
    return out


def has_xfa(doc) -> bool:
    af = doc.resolve(doc.root.get("AcroForm"))
    return isinstance(af, dict) and af.get("XFA") is not None
