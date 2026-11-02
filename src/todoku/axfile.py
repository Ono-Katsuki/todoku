"""ActiveX の部品を、入れ物（Excel の .xlsx・.xlsm・.xls、Word の .docx・.docm・.doc）から読み、値を書き、証明する。

入れ物ごとの部品の置き場所：
- OOXML（MS-OI29500 3.4.1.1・3.6、ISO/IEC 29500 の Embedded Control Persistence）：シート（w:control）から
  activeX/activeXN.xml（ax:ocx：classid・persistence・r:id）→ activeXN.bin。persistence が
  persistStreamInit / persistStream なら .bin は CLSID（16バイト）と部品の保存の形（MS-OFORMS 2.1.1.2）、
  persistStorage なら .bin は複合ファイルで、部品は "contents" ストリーム（MS-OFORMS 2.1.1.3。フレームは "f" と "o"）、
  persistPropertyBag なら性質は ax:ocxPr の名前と値（MS-OFORMS 2.1.1.1）
  Excel は controlPr に linkedCell（つながったセル）・listFillRange（選択肢の範囲）と、控えの絵（r:id）を持つ。
  VML の v:imagedata（o:relid）も同じ控えの絵を指す
- .xls（MS-XLS 2.4.181 Obj・2.5.150 FtPictFmla・2.5.203 PictFmlaKey）：絵の OBJ（ot=8）で pictFlags.fCtl=1。
  fPrstm=1 なら "Ctls" ストリームの lPosInCtlStm から cbBufInCtlStm バイト、0 なら "MBD" と8桁の16進の名前の入れ物。
  つながったセル・選択肢の範囲は PictFmlaKey の fmlaLinkedCell・fmlaListFillRange
- .doc（MS-DOC 2.1.4・2.9.165 ODT）：CONTROL のフィールドの区切りの字の sprmCPicLocation の値の前に「_」を付けた、
  ObjectPool の中の入れ物。ODT の fOCX が立ち、fStream なら "\\003OCXDATA"、そうでなければ "contents"
"""

from __future__ import annotations

import posixpath
import re
import struct
from dataclasses import dataclass, field

from . import activex as ax
from . import axdraw

CHECK_KINDS = ("CheckBox", "ToggleButton")
TEXT_KINDS = ("TextBox",)
LIST_KINDS = ("ComboBox", "ListBox")
NUM_KINDS = ("SpinButton", "ScrollBar")
CAPTION_KINDS = ("Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "Frame")


@dataclass
class AxRef:
    """ファイルの中の1つの ActiveX の部品：置き場所と、読んだ中身。"""

    host: str  # xlsx / docx / xls / doc
    name: str
    clsid: str = ""
    progid: str = ""
    kind: str | None = None  # Forms 2.0 の種類（Forms 2.0 でなければ None）
    persistence: str = ""  # persistStreamInit / persistStream / persistStorage / persistPropertyBag / ctls / storage
    part: str = ""  # OOXML：activeX*.xml、.xls："Ctls" か MBD の入れ物、.doc：ObjectPool の入れ物
    data_part: str = ""  # 中身：activeX*.bin・"Ctls"・入れ物の中の "contents"
    stream: str = ""  # persistStorage の .bin（複合ファイル）の中のストリームの名前
    start: int = 0  # 中身の中の、部品の始まり（CLSID の後ろ・Ctls の中の位置）
    length: int | None = None  # 部品の長さ（None は中身の終わりまで）
    control: ax.Control | None = None
    frame: ax.Frame | None = None
    bag: dict | None = None  # persistPropertyBag の性質（名前 → 文字）
    images: list[str] = field(default_factory=list)  # 控えの絵の部品（OOXML）
    # 控えの絵を指している所：(種類 sheet / vml / word, その XML の部品, 部品の目印（図形の番号・要素の位置）, 絵の部品)
    pic_refs: list = field(default_factory=list)
    error: str = ""

    @property
    def caption(self) -> str:
        if self.control is not None:
            return self.control.caption
        if self.frame is not None:
            return self.frame.caption
        return ""

    @property
    def label(self) -> str:
        return ax.KIND_JA.get(self.kind or "", self.kind or "ActiveX の部品")


# ====================================================================== OOXML
def rels_of(zip_, part: str) -> dict[str, tuple[str, str]]:
    """部品の関係（rId → (種類の最後の語, 部品の名前)）。"""
    d, base = posixpath.split(part)
    rp = posixpath.join(d, "_rels", base + ".rels")
    out: dict[str, tuple[str, str]] = {}
    if not zip_.has(rp):
        return out
    for m in re.finditer(rb"<(?:\w+:)?Relationship\b([^>]*?)/?>", zip_.read(rp)):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m[1].decode("utf-8", "replace")))
        t = a.get("Target", "")
        if a.get("TargetMode") == "External":
            full = t
        elif t.startswith("/"):
            full = t[1:]
        else:
            full = posixpath.normpath(posixpath.join(d, t))
        if "Id" in a:
            out[a["Id"]] = (a.get("Type", "").rsplit("/", 1)[-1], full)
    return out


_OCX = re.compile(rb"<(?:\w+:)?ocx\b([^>]*?)(/?)>")


def _attr(attrs: bytes, name: str) -> str | None:
    m = re.search(rb"(?:\b\w+:)?\b" + name.encode() + rb'="([^"]*)"', attrs)
    return m[1].decode("utf-8", "replace") if m else None


def _unxml(s: str) -> str:
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'")
            .replace("&#xA;", "\n").replace("&#xD;", "\r").replace("&#10;", "\n").replace("&#13;", "\r")
            .replace("&amp;", "&"))


def load_ooxml(zip_, xml_part: str, host: str, name: str = "") -> AxRef:
    """activeX*.xml（ax:ocx）と、その中身（.bin か ax:ocxPr）を読む。読めないときは error に理由。"""
    ref = AxRef(host, name, part=xml_part)
    if not zip_.has(xml_part):
        ref.error = f"{xml_part} がありません"
        return ref
    xml = zip_.read(xml_part)
    m = _OCX.search(xml)
    if not m:
        ref.error = "ActiveX の部品の定義（ax:ocx）がありません"
        return ref
    attrs = m[1]
    ref.clsid = (_attr(attrs, "classid") or "").strip("{}").upper()
    ref.persistence = _attr(attrs, "persistence") or ""
    ref.kind = ax.kind_of(ref.clsid)
    rid = re.search(rb'\br:id="([^"]*)"', attrs)
    if ref.persistence == "persistPropertyBag":
        ref.bag = {_unxml(n.decode("utf-8", "replace")): _unxml(v.decode("utf-8", "replace"))
                   for n, v in re.findall(rb'<(?:\w+:)?ocxPr\b[^>]*?\b(?:\w+:)?name="([^"]*)"[^>]*?\b(?:\w+:)?value="([^"]*)"', xml)}
        if ref.kind in ax.STRUCT:
            try:
                ref.control = bag_control(ref.kind, ref.bag)
                ref.kind = ref.control.kind
            except ax.AxError as e:
                ref.error = str(e)
        return ref
    rels = rels_of(zip_, xml_part)
    if not rid or rid[1].decode() not in rels:
        ref.error = "部品の中身（activeX*.bin）への関係がありません"
        return ref
    ref.data_part = rels[rid[1].decode()][1]
    if not zip_.has(ref.data_part):
        ref.error = f"{ref.data_part} がありません"
        return ref
    data = zip_.read(ref.data_part)
    try:
        if ref.persistence in ("persistStreamInit", "persistStream"):
            prefix, _body = ax.split_stream_init(data, ref.clsid if ref.clsid in ax.CLSIDS else None)
            ref.start = len(prefix)
            if prefix:
                ref.clsid = ax.guid_str(prefix)
                ref.kind = ax.kind_of(ref.clsid) or ref.kind
            _read_control(ref, data)
        elif ref.persistence == "persistStorage":
            _read_storage(ref, data)
        else:
            ref.error = f"保存の仕方（{ref.persistence or '不明'}）が分かりません"
    except ax.AxError as e:
        ref.error = str(e)
    return ref


def _read_control(ref: AxRef, data: bytes) -> None:
    if ref.kind not in ax.STRUCT:
        if ref.kind is None:
            ref.error = "Forms 2.0 でない ActiveX の部品です（中身の形が分かりません）"
        return
    end = ref.start + ref.length if ref.length is not None else None
    ref.control = ax.parse(data, ref.kind, ref.start, end)
    ref.kind = ref.control.kind


def _read_storage(ref: AxRef, data: bytes) -> None:
    """persistStorage：.bin は複合ファイル。部品は "contents"（フレームは "f" と "o"）。"""
    from .cfb import CFB
    from .cfbwrite import storage_clsid

    c = CFB(data)
    root = storage_clsid(c, "")
    if root:
        ref.clsid = root
        ref.kind = ax.kind_of(root) or ref.kind
    names = {n.lower(): n for n in c.streams()}
    if ref.kind is None and "\x01compobj" in names:
        ref.progid = ax.compobj_progid(c.read(names["\x01compobj"]))
        ref.kind = ax.kind_of(progid=ref.progid)
    if ref.kind == "Frame" or ("f" in names and "contents" not in names):
        ref.kind = "Frame"
        ref.frame = ax.parse_frame(c.read(names["f"]), c.read(names["o"]) if "o" in names else b"")
        return
    if "contents" not in names:
        ref.error = "部品の中身（contents）がありません"
        return
    ref.stream = names["contents"]
    _read_control(ref, c.read(ref.stream))


# ---------------------------------------------------------------------- 性質の袋（persistPropertyBag）
class _BagBlock:
    """性質の袋（MS-OFORMS 2.1.1.1）を、保存の形（Block）と同じ問い合わせで読めるようにしたもの。"""

    def __init__(self, struct_name: str, values: dict):
        self.layout = ax.LAYOUTS[struct_name]
        self.values = values
        self.data_at: dict = {}
        self.counts: dict = {}

    def get(self, prop: str, struct_name: str | None = None):
        if prop in self.values:
            return self.values[prop]
        return ax.default_of(struct_name or self.layout.name, prop)

    def stored(self, prop: str) -> bool:
        return prop in self.values


def bag_control(kind: str, bag: dict) -> ax.Control:
    """性質の袋の文字（数は符号なしの10進・真偽は 0 / -1・大きさは「幅;高さ」）を、部品として読む。"""
    st = ax.STRUCT[kind]
    lay = ax.LAYOUTS[st]
    sizes = {n: s for n, _b, s in lay.data + lay.extra}
    vals: dict = {}
    for name, text in bag.items():
        if name not in sizes:
            continue
        s = sizes[name]
        if s == ax.S:
            vals[name] = text
        elif name == "Size" or s == 8:
            a, _, b = text.partition(";")
            vals[name] = (int(a or 0), int(b or 0))
        elif name in ax.CHAR:
            vals[name] = chr(int(text)) if text.strip("0") else ""
        else:
            n = int(text) if text.lstrip("-").isdigit() else 0
            if name in ax.SIGNED and n >= 1 << (8 * s - 1):
                n -= 1 << (8 * s)
            vals[name] = n
    blk = _BagBlock(st, vals)
    text = None
    if lay.textprops:
        tv = {k: (bag[k] if k == "FontName" else int(bag[k])) for k in ("FontName", "FontEffects", "FontHeight",
              "FontCharSet", "FontPitchAndFamily", "ParagraphAlign", "FontWeight") if k in bag and (k == "FontName" or bag[k].lstrip("-").isdigit())}
        text = _BagBlock("TextProps", tv)
    real = kind
    if st == "MorphData":
        real = ax.DISPLAY.get(blk.get("DisplayStyle", st), kind)
    return ax.Control(real, st, b"", blk, text)


def _xml_attr(s: str) -> bytes:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
            .replace("\n", "&#xA;").replace("\r", "&#xD;").replace("\t", "&#x9;")).encode("utf-8")


def set_bag(xml: bytes, changes: dict) -> bytes:
    """性質の袋（ax:ocxPr）の値だけを変える。既定値になる性質の要素は消し、なかった性質は最後の ocxPr の後ろに足す。"""
    for name, v in changes.items():
        pat = re.compile(rb'[ \t]*<(\w+:)?ocxPr\b[^>]*?\b(?:\w+:)?name="' + re.escape(name.encode()) + rb'"[^>]*/>')
        m = pat.search(xml)
        if v is None or v == "":
            if m:
                xml = xml[:m.start()] + xml[m.end():]
            continue
        text = _xml_attr(str(v))
        if m:
            new = re.sub(rb'(\b(?:\w+:)?value=)"[^"]*"', lambda mm: mm[1] + b'"' + text + b'"', m[0], count=1)
            xml = xml[:m.start()] + new + xml[m.end():]
            continue
        last = None
        for mm in re.finditer(rb"<(\w+:)?ocxPr\b[^>]*/>", xml):
            last = mm
        pre = (last[1] if last else b"ax:") or b""
        elem = b"<" + pre + b"ocxPr " + pre + b'name="' + _xml_attr(name) + b'" ' + pre + b'value="' + text + b'"/>'
        if last is not None:
            xml = xml[:last.end()] + elem + xml[last.end():]
        else:
            end = re.search(rb"</(\w+:)?ocx>", xml)
            if end is None:
                raise ax.AxError("性質の袋（ax:ocx）に、性質を足す場所がありません")
            xml = xml[:end.start()] + elem + xml[end.start():]
    return xml


# ---------------------------------------------------------------------- 書く（OOXML の部品）
def new_parts(zip_, ref: AxRef, changes: dict, items: list[str] | None = None,
              parts: dict | None = None, shared: set | None = None, added: dict | None = None) -> dict[str, bytes]:
    """部品の性質を変えた、OOXML の部品（.bin か .xml）と、描き直した控えの絵。parts は、すでに書き換えた部品。
    shared：ほかの部品と共有している控えの絵（Excel は同じ絵を1つの部品にまとめる）。共有している絵は書き換えず、
    この部品だけの新しい絵の部品を足して（added：名前 → 写すエントリ）、この部品からの参照だけをそこへ付け替える。"""
    parts = parts if parts is not None else {}
    added = added if added is not None else {}

    def cur(name):
        if name in parts:
            return parts[name]
        return added[name][0] if name in added else zip_.read(name)

    out: dict[str, bytes] = {}
    if ref.control is None:
        raise ax.AxError(f"{ref.label}「{ref.name}」の中身を読めないので書けません：{ref.error or '中身がありません'}")
    if ref.persistence == "persistPropertyBag":
        bag = dict(ref.bag or {})
        text_changes = {}
        for k, v in changes.items():
            d = ax.default_of(ref.control.struct, k)
            text_changes[k] = None if (v == d or v == "") else str(v)
            if text_changes[k] is None:
                bag.pop(k, None)
            else:
                bag[k] = text_changes[k]
        out[ref.part] = set_bag(cur(ref.part), text_changes)
        new_ctl = bag_control(_struct_kind(ref), bag)
    else:
        data = cur(ref.data_part)
        if ref.persistence == "persistStorage":
            from .cfb import CFB
            from .cfbwrite import replace_stream

            c = CFB(data)
            old = c.read(ref.stream)
            new = ax.set_props(old, _struct_kind(ref), changes)
            out[ref.data_part] = replace_stream(data, ref.stream, new)
            new_ctl = ax.parse(new, _struct_kind(ref))
        else:
            end = ref.start + ref.length if ref.length is not None else len(data)
            seg = ax.set_props(data, _struct_kind(ref), changes, ref.start, end)
            out[ref.data_part] = data[:ref.start] + seg + data[end:]
            new_ctl = ax.parse(seg, _struct_kind(ref))
    for img in ref.images:
        if img in out:
            continue
        pic = axdraw.picture(new_ctl, cur(img), items)
        if shared and img in shared:
            target = _split_picture(zip_, ref, img, cur, out, added)
            added[target] = (pic, img)
        else:
            out[img] = pic
    return out


def _new_name(zip_, added: dict, img: str) -> str:
    d, base = posixpath.split(img)
    m = re.match(r"(\D*)(\d*)(\.\w+)$", base)
    stem, ext = (m[1], m[3]) if m else ("image", posixpath.splitext(base)[1])
    nums = [int(x) for n in list(zip_.names()) + list(added) if posixpath.dirname(n) == d
            for x in re.findall(re.escape(stem) + r"(\d+)" + re.escape(ext) + "$", posixpath.basename(n))]
    return posixpath.join(d, f"{stem}{max(nums, default=0) + 1}{ext}")


def _add_rel(rels_xml: bytes, target: str) -> tuple[bytes, str]:
    ids = set(re.findall(rb'\bId="([^"]*)"', rels_xml))
    n = 1
    while f"rIdAx{n}".encode() in ids:
        n += 1
    rid = f"rIdAx{n}"
    rel = (f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
           f'Target="{target}"/>').encode()
    end = rels_xml.rfind(b"</Relationships>")
    if end < 0:
        raise ax.AxError("関係の部品（.rels）に、関係を足す場所がありません")
    return rels_xml[:end] + rel + rels_xml[end:], rid


def _rels_path(part: str) -> str:
    d, base = posixpath.split(part)
    return posixpath.join(d, "_rels", base + ".rels")


def _split_picture(zip_, ref: AxRef, img: str, cur, out: dict, added: dict) -> str:
    """共有している控えの絵を、この部品だけの新しい絵の部品に分ける：関係を足し、この部品からの参照の rId だけを替える。"""
    target = _new_name(zip_, added, img)
    done = False
    for kind, owner, key, pic in ref.pic_refs:
        if pic != img:
            continue
        xml = out.get(owner, cur(owner))
        rp = _rels_path(owner)
        rel_target = posixpath.relpath(target, posixpath.dirname(owner))
        rels_xml, rid = _add_rel(out.get(rp, cur(rp)), rel_target)
        if kind == "sheet":
            m = re.search(rb"<(?:\w+:)?control\b[^>]*\bshapeId=\"" + key.encode() + rb"\"[^>]*[^/]>", xml)
            pr = re.compile(rb"<(?:\w+:)?controlPr\b[^>]*>").search(xml, m.end()) if m else None
            if pr is None:
                raise ax.AxError(f"「{ref.name}」の控えの絵を指す所（controlPr）が見つかりません")
            tag = re.sub(rb'(\br:id=")[^"]*(")', lambda mm: mm[1] + rid.encode() + mm[2], pr[0], count=1)
            xml = xml[:pr.start()] + tag + xml[pr.end():]
        elif kind == "vml":
            m = None
            for mm in re.finditer(rb"<v:shape\b([^>]*)>(.*?)</v:shape>", xml, re.S):
                if re.search(rb'\b(?:id|o:spid)="_x0000_s' + key.encode() + rb'"', mm[1]):
                    m = mm
                    break
            if m is None:
                raise ax.AxError(f"「{ref.name}」の VML の図形が見つかりません")
            body = re.sub(rb'(<v:imagedata\b[^>]*\bo:relid=")[^"]*(")', lambda q: q[1] + rid.encode() + q[2], m[2], count=1)
            xml = xml[:m.start(2)] + body + xml[m.end(2):]
        else:  # word：w:control（key は w:control の r:id）の前の v:imagedata
            m = re.search(rb"<w:control\b[^>]*\br:id=\"" + key.encode() + rb"\"", xml)
            start = max(xml.rfind(b"<w:object", 0, m.start()), xml.rfind(b"<w:pict", 0, m.start())) if m else -1
            im = re.compile(rb'(<v:imagedata\b[^>]*\br:id=")([^"]*)(")').search(xml, start, m.start()) if m and start >= 0 else None
            if im is None:
                raise ax.AxError(f"「{ref.name}」の控えの絵を指す所（v:imagedata）が見つかりません")
            xml = xml[:im.start(2)] + rid.encode() + xml[im.end(2):]
        out[owner] = xml
        out[rp] = rels_xml
        done = True
    if not done:
        raise ax.AxError(f"「{ref.name}」の控えの絵を指す所が分からないので、共有している絵を分けられません")
    return target


def _struct_kind(ref: AxRef) -> str:
    """部品を読むときに渡す種類（classid から分かったもの。MorphData は DisplayStyle で読み直す）。"""
    return ax.kind_of(ref.clsid, ref.progid) or ref.kind or ""


# ====================================================================== 窓の印と、欄
def mark(ref: AxRef) -> str:
    """読み順に出す部品の状態（Excel の窓の行の中の印）。"""
    c = ref.control
    k = ref.kind
    if c is None:
        return f"〔{ref.label}〕"
    if k in CHECK_KINDS:
        st = c.checked
        return ("☑" if st else ("☐" if st is False else "▣")) + c.caption
    if k == "OptionButton":
        st = c.checked
        return ("◉" if st else "○") + c.caption
    if k == "TextBox":
        return "✎" + c.value.replace("\r\n", "⏎").replace("\n", "⏎")
    if k == "ComboBox":
        return "▼" + c.value
    if k == "ListBox":
        return "≡" + c.value
    if k in NUM_KINDS:
        return f"⇕{c.get('Position')}"
    return c.caption


def outline(ref: AxRef) -> str:
    """値を書かない部品（ボタン・フレーム）と、読めない部品：読み順の中に、何があるかだけを示す。"""
    if ref.frame is not None:
        inner = []
        for s in ref.frame.sites:
            c = s.control
            if c is None:
                inner.append(f"{s.name}（{ax.KIND_JA.get(s.kind or '', s.kind or '部品')}）")
            elif s.kind in CHECK_KINDS + ("OptionButton",):
                inner.append(("☑" if c.checked else "☐") + (c.caption or s.name))
            elif s.kind in TEXT_KINDS + LIST_KINDS:
                inner.append(f"［{s.name}＝{c.value}］" if c.value else f"［{s.name}］")
            elif s.kind in NUM_KINDS:
                inner.append(f"⇕{c.get('Position')}［{s.name}］")
            else:
                inner.append(c.caption or s.name)
        return f"〔枠「{ref.frame.caption}」：{' '.join(inner)}〕"
    if ref.error and ref.control is None:
        return f"〔{ref.label}「{ref.name}」：中身を読めません（{ref.error}）〕"
    if ref.kind == "CommandButton":
        return f"〔ボタン「{ref.caption}」〕"
    return f"〔{ref.label}「{ref.name}」〕"


def texts(ref: AxRef) -> list[str]:
    """部品の中の文字（見出し・テキストボックスなどの文字・フレームの中の部品の見出しと文字）。読み落としの検査に使う。"""
    out = []
    for c in ([ref.control] if ref.control is not None else []) + \
            ([s.control for s in ref.frame.sites if s.control is not None] if ref.frame is not None else []):
        if c.kind in CAPTION_KINDS and c.caption.strip():
            out.append(c.caption)
        if c.kind in TEXT_KINDS + LIST_KINDS and c.value.strip():
            out.append(c.value)
    if ref.frame is not None and ref.frame.caption.strip():
        out.append(ref.frame.caption)
    return out


def value_of(ref: AxRef):
    """欄の値：テキストボックス・コンボボックス・リストボックスは文字、チェック・オプションは真偽（どちらでもないは None）、
    スピン・スクロールは数。"""
    c = ref.control
    if c is None:
        return None
    k = ref.kind
    if k in CHECK_KINDS or k == "OptionButton":
        return c.checked
    if k in NUM_KINDS:
        return c.get("Position")
    if k in TEXT_KINDS + LIST_KINDS:
        return c.value
    return None


def sample_value(f: dict):
    """全入力欄への書き込みの試しに使う、部品の決まりに合う値（窓の欄 f から決める）。決められなければ None。
    オプションボタンは、組ごとに1つ（いま選んでいない最後のボタン、それが選ばれていれば最初のボタン）だけを選ぶ。"""
    k = f["kind"]
    if k == "text":
        n = f.get("maxlen") or 0
        return "テスト"[:n] if n and n < 3 else "テスト"
    if k == "checkbox":
        return not bool(f.get("value"))
    if k == "option":
        members = f.get("group_members") or [f.get("control")]
        target = members[-1] if f.get("group_selected") != members[-1] else members[0]
        return f.get("control") == target
    if k in ("select", "combo"):
        opts = f.get("options")
        if opts:
            return next((o for o in opts if o != f.get("value")), opts[0])
        return "テスト" if k == "combo" else None
    if k == "spin":
        lo, hi = f.get("min", 0), f.get("max", 100)
        return lo if f.get("value") != lo else min(hi, lo + 1)
    return None


def field_kind(ref: AxRef) -> str:
    k = ref.kind
    c = ref.control
    if k == "TextBox":
        return "text"
    if k in CHECK_KINDS:
        return "checkbox"
    if k == "OptionButton":
        return "option"
    if k == "ComboBox":
        return "select" if c is not None and (c.style == "droplist" or c.bit("MatchRequired")) else "combo"
    if k == "ListBox":
        return "multiselect" if c is not None and c.get("MultiSelect") else "select"
    if k in NUM_KINDS:
        return "spin"
    return "other"


def describe(ref: AxRef, items: list[str] | None = None) -> dict:
    """欄の一覧に添える、部品の性質（どう書けば受け付けるか）。"""
    c = ref.control
    d: dict = {"activex": ref.label, "control_kind": f"ActiveX {ref.kind}"}
    if c is None:
        d["unreadable"] = ref.error
        return d
    k = ref.kind
    if c.caption and k in CAPTION_KINDS:
        d["caption"] = c.caption
    if k == "TextBox":
        if c.get("MaxLength"):
            d["maxlen"] = c.get("MaxLength")
        if c.bit("MultiLine"):
            d["multiline"] = True
        if c.get("PasswordChar"):
            d["password"] = True
        d["write_as"] = "文字" + ("（改行も書けます）" if c.bit("MultiLine") else "（1行）")
    elif k in CHECK_KINDS:
        d["write_as"] = "TRUE（チェックあり）/ FALSE（なし）"
    elif k == "OptionButton":
        d["group"] = c.get("GroupName") or ""
        d["write_as"] = "TRUE（これを選ぶ。同じ組のほかのボタンは外れる）/ FALSE"
    elif k in LIST_KINDS:
        if items is not None:
            d["options"] = items[:500]
        must = k == "ListBox" or c.style == "droplist" or c.bit("MatchRequired")
        d["write_as"] = ("選択肢の文字" if must else "文字（選択肢にない文字も書けます）") + \
            ("" if items is not None else "（選択肢はファイルの中にありません）")
        if must and items is None:
            # 選択肢はマクロなどで後から入る：選択肢に合うかを確かめられないので、確かめずに書くとき（force）だけ書く
            d["options_unknown"] = "選択肢がファイルの中にないので、選択肢に合うかを確かめられません（確かめずに書くには force）"
        if k == "ListBox" and c.get("MultiSelect"):
            d["readonly"] = "複数選択のリストボックスは、選んだ項目がファイルに保存されないので書けません"
    elif k in NUM_KINDS:
        lo, hi = ax.value_range(c)
        d.update({"min": lo, "max": hi, "write_as": f"{lo}〜{hi} の整数"})
    if not c.bit("Enabled"):
        d["readonly"] = "使えない（Enabled が偽）ようにされた部品です"
    elif c.bit("Locked") and k in TEXT_KINDS + LIST_KINDS + CHECK_KINDS + ("OptionButton",):
        d["readonly"] = "書き換えられない（Locked が真）ようにされた部品です"
    return d


# ====================================================================== Excel（.xlsx・.xlsm）
def load_excel(wb, sh, ctl, part: str | None, pr, rels: dict) -> None:
    """シートの control（ActiveX）を読み、ctl（xlsx.workbook.Control）に入れる。pr は controlPr の要素（なければ None）。"""
    if not part:
        return
    ref = load_ooxml(wb.zip, part, "xlsx", ctl.name)
    ctl.ax = ref
    if ref.kind:
        ctl.kind = "ActiveX " + ref.kind
    if pr is not None:
        if pr.get("linkedCell"):
            ctl.link = pr.get("linkedCell")
        if pr.get("listFillRange"):
            ctl.fmla_range = pr.get("listFillRange")
        rid = pr.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if rid and rid in rels and rels[rid][0] == "image" and wb.zip.has(rels[rid][1]):
            ref.images.append(rels[rid][1])
            ref.pic_refs.append(("sheet", sh.path, ctl.shape_id, rels[rid][1]))
    c = ref.control
    if c is not None:
        if ref.kind in CAPTION_KINDS:
            ctl.text = c.caption
        if ref.kind in CHECK_KINDS + ("OptionButton",):
            st = c.checked
            ctl.checked = "Checked" if st else ("Mixed" if st is None else None)
        if ref.kind in NUM_KINDS:
            lo, hi = ax.value_range(c)
            ctl.val, ctl.vmin, ctl.vmax = c.get("Position"), lo, hi
    elif ref.frame is not None:
        ctl.text = ref.frame.caption


def excel_vml_images(wb, sh, vml_path: str | None) -> None:
    """VML の図形（v:imagedata の o:relid）が指す控えの絵も、部品の控えの絵に加える。"""
    if not vml_path or not wb.zip.has(vml_path):
        return
    vml = wb.zip.read(vml_path).decode("utf-8", errors="replace")
    rels = rels_of(wb.zip, vml_path)
    for ctl in sh.controls:
        ref = getattr(ctl, "ax", None)
        if ref is None or not ctl.shape_id:
            continue
        for m in re.finditer(r"<v:shape\b([^>]*)>(.*?)</v:shape>", vml, re.S):
            if not re.search(r'\b(?:id|o:spid)="_x0000_s' + re.escape(ctl.shape_id) + '"', m[1]):
                continue
            im = re.search(r'<v:imagedata\b[^>]*\bo:relid="([^"]*)"', m[2])
            if im and im[1] in rels and rels[im[1]][0] == "image":
                target = rels[im[1]][1]
                if wb.zip.has(target):
                    if target not in ref.images:
                        ref.images.append(target)
                    ref.pic_refs.append(("vml", vml_path, ctl.shape_id, target))
            break


def items_excel(wb, sh, ctl) -> list[str] | None:
    """コンボボックス・リストボックスの選択肢：ListFillRange のセルの見た目の文字（複数の列なら BoundColumn の列）。"""
    if not ctl.fmla_range:
        return None
    from .formula.parser import parse
    from .formula.values import to_text
    from .xlsx.calc import Calc

    try:
        from .formula.evaluator import Evaluator

        calc = Calc(wb)
        ev = Evaluator(calc, sh.name, 1, 1)
        g = ev.grid(ev.ev(parse(ctl.fmla_range.lstrip("="))))
    except Exception:  # noqa: BLE001
        return None
    c = ctl.ax.control if getattr(ctl, "ax", None) is not None else None
    col = 0
    if c is not None:
        bc = c.get("BoundColumn") or 1
        col = max(0, min(bc - 1, len(g.data[0]) - 1 if g.data else 0))
    return [to_text(row[col]) if row[col] is not None else "" for row in g.data if len(row) > col]


def excel_groups(sh) -> dict[str, list]:
    """オプションボタンの組（同じシートの、同じ GroupName のボタン。GroupName がなければシートで1組）。"""
    out: dict[str, list] = {}
    for c in sh.controls:
        ref = getattr(c, "ax", None)
        if ref is not None and ref.kind == "OptionButton" and ref.control is not None:
            out.setdefault(ref.control.get("GroupName") or "", []).append(c)
    return out


def excel_field(wb, sh, lay, b, d: dict) -> dict:
    """窓の欄の一覧に出す、ActiveX の部品の欄。d は配置から読んだ見出し（lay.describe）。"""
    from .formula.parser import addr

    ctl = b.control
    ref = ctl.ax
    items = items_excel(wb, sh, ctl) if ref.kind in LIST_KINDS else None
    link = ctl.link_cell
    fid = f"{link[0]}!{addr(*link[1:])}" if link else f"{sh.name}!{ctl.name}"
    f = {"id": fid, "cell": addr(*link[1:]) if link else None, "control": ctl.name, "kind": field_kind(ref),
         "value": value_of(ref), **d, **describe(ref, items), "provenance": "layout"}
    if ref.kind == "OptionButton" and ref.control is not None:
        grp = excel_groups(sh).get(ref.control.get("GroupName") or "", [])
        f["group_members"] = [x.name for x in grp]
        f["group_selected"] = next((x.name for x in grp if x.ax.control.checked), None)
    if ref.kind in CAPTION_KINDS and ref.control is not None and ref.control.caption:
        # 見出しのある部品（チェック・オプション）：部品の中の見出しが項目の名前（配置から読んだ見出しは区分として添える）
        near = d.get("label") if d.get("label_method") == "左の見出し" or "上の見出し" in (d.get("label_method") or "") else None
        f["label"] = f"{near}／{ref.control.caption}" if near and near != ref.control.caption else ref.control.caption
        f["label_method"] = "部品の見出し（Caption）"
    return f


@dataclass
class ExcelWrite:
    cells: dict = field(default_factory=dict)  # つながったセル → 値
    controls: list = field(default_factory=list)  # (シート, 部品, 変える性質, 選択肢)
    notes: list = field(default_factory=list)


def _cell_value_for(ref: AxRef, changes: dict, items: list[str] | None, raw_items: list | None):
    """部品の値に合わせて、つながったセルに入る値（Excel と同じ：テキストボックスは文字、チェック・オプションは TRUE/FALSE、
    スピン・スクロールは数、コンボボックス・リストボックスは選んだ項目（選択肢のセルの値））。"""
    k = ref.kind
    if "Position" in changes:
        return float(changes["Position"])
    v = changes.get("Value", "")
    if k in CHECK_KINDS + ("OptionButton",):
        return v == "1"
    if k in LIST_KINDS and items is not None and raw_items is not None and v in items:
        return raw_items[items.index(v)]
    return v


def _raw_items(wb, sh, ctl) -> list | None:
    if not ctl.fmla_range:
        return None
    from .formula.parser import parse
    from .xlsx.calc import Calc

    try:
        from .formula.evaluator import Evaluator

        ev = Evaluator(Calc(wb), sh.name, 1, 1)
        g = ev.grid(ev.ev(parse(ctl.fmla_range.lstrip("="))))
    except Exception:  # noqa: BLE001
        return None
    c = ctl.ax.control
    bc = (c.get("BoundColumn") or 1) - 1 if c is not None else 0
    return [row[bc] if len(row) > bc else None for row in g.data]


def plan_excel(wb, sheet: str, ctl, value, force: bool = False, prior: list | None = None) -> ExcelWrite:
    """ActiveX の部品 ctl に value を書くときの、部品の性質の変化（オプションボタンは同じ組のほかのボタンも外す）と、
    つながったセルの値。prior は、同じ書き込みで先に書く部品（Excel で人が順にクリックしたのと同じく、その後の状態で外す）。"""
    sh = wb.sheet(sheet)
    ref = ctl.ax
    out = ExcelWrite()
    if ref is None or ref.control is None:
        raise ax.AxError(f"「{ctl.name}」の中身を読めないので書けません：{getattr(ref, 'error', '') or '中身がありません'}")
    items = items_excel(wb, sh, ctl) if ref.kind in LIST_KINDS else None
    changes, notes = ax.plan(ref.control, value, items, force)
    out.notes += notes
    targets = [(ctl, changes)]
    if ref.kind == "OptionButton" and changes.get("Value") == "1":
        now = {}  # 先に書く部品のあとの、ボタンの Value
        for _s, c_, ch_, _i in prior or ():
            if "Value" in ch_:
                now[id(c_.ax)] = ch_["Value"]
        for other in excel_groups(sh).get(ref.control.get("GroupName") or "", []):
            on = now[id(other.ax)] == "1" if id(other.ax) in now else bool(other.ax.control.checked)
            if other is not ctl and on:
                targets.append((other, {"Value": "0"}))
    raw = _raw_items(wb, sh, ctl) if items is not None else None
    for c, ch in targets:
        out.controls.append((sheet, c, ch, items if c is ctl else None))
        if c.link_cell and wb.has_sheet(c.link_cell[0]):
            key = (wb.sheet(c.link_cell[0]).name, c.link_cell[1], c.link_cell[2])
            out.cells[key] = _cell_value_for(c.ax, ch, items if c is ctl else None, raw if c is ctl else None)
    return out


def excel_links(wb) -> dict:
    """ActiveX の部品につながったセル → (シート, 部品)。"""
    out = {}
    for sh in wb.sheets:
        for c in sh.controls:
            ref = getattr(c, "ax", None)
            if ref is not None and ref.kind in ax.VALUE_KINDS and c.link_cell and wb.has_sheet(c.link_cell[0]):
                out[(wb.sheet(c.link_cell[0]).name, c.link_cell[1], c.link_cell[2])] = (sh.name, c)
    return out


def merge_writes(writes: list) -> list:
    """同じ部品への書き込みを1つにまとめる（順は、その部品を初めに書いた所）。

    オプションボタンの組では、ほかのボタンを選ぶと、このボタンにも「Value を 0 に」が届き、名指しで外す書き込みと重なる。
    同じ部品を2度書くと、共有している控えの絵を2度分けて、1つ目の絵がどこからも指されないまま残るので、1度だけ書く。
    同じ性質に違う値を書くことになる（同じ組のボタンを2つ選んだ）なら、どちらが望みか決まらないので断る。"""
    out: dict[int, list] = {}
    for sheet, ctl, changes, items in writes:
        k = id(ctl.ax)
        if k in out:
            prev = out[k][2]
            if any(key in prev and prev[key] != v for key, v in changes.items()):
                raise ax.AxError(f"「{ctl.name}」に、ひとつの書き込みの中で違う値を2度書くことになります"
                                 "（同じ組のオプションボタンを2つ選んでいませんか）。選ぶボタンを1つにしてください")
            out[k][2] = {**prev, **changes}
            out[k][3] = items if items is not None else out[k][3]
        else:
            out[k] = [sheet, ctl, dict(changes), items]
    return [tuple(w) for w in out.values()]


def apply_excel(wb, writes: list) -> tuple[dict[str, bytes], list[dict]]:
    """部品の性質を変えた部品（.bin・.xml）と、描き直した控えの絵。(部品の名前 → 中身, 書いたことの記録)。"""
    parts: dict[str, bytes] = {}
    done = []
    users: dict[str, set] = {}
    for sh in wb.sheets:
        for c in sh.controls:
            if getattr(c, "ax", None) is not None:
                for img in c.ax.images:
                    users.setdefault(img, set()).add(id(c.ax))
    shared = {img for img, u in users.items() if len(u) > 1}
    added: dict[str, tuple[bytes, str]] = {}
    for sheet, ctl, changes, items in merge_writes(writes):
        ref = ctl.ax
        if all(ref.control.get(k) == v or (v in (None, "") and ref.control.get(k) == ax.default_of(ref.control.struct, k))
               for k, v in changes.items()):
            continue  # 変わらない
        n_added = len(added)
        new = new_parts(wb.zip, ref, changes, items, parts, shared, added)
        parts.update(new)
        done.append({"sheet": sheet, "control": ctl.name, "kind": ref.kind, "changes": dict(changes),
                     "before": {k: ref.control.get(k) for k in changes}, "parts": sorted(new) + list(added)[n_added:]})
    for name, (data, like) in added.items():
        wb.zip.add(name, data, like)  # 共有していた控えの絵から分けた、この部品だけの絵
    return parts, done


# ====================================================================== 証明（OOXML の部品）
def verify_parts(za, zb, ra: AxRef, rb: AxRef, changes: dict | None, items: list[str] | None = None) -> list[str]:
    """書き込み前後の部品（ra・rb）を、それぞれのファイル（za・zb）から読み直して比べる。changes は書き込みを指示した性質
    （None は「変わっていてはいけない」）。保存の形のバイト・入れ物（CLSID・複合ファイル・性質の袋）・控えの絵を確かめる。"""
    out: list[str] = []
    what = f"{ra.label}「{ra.name}」"
    if ra.error or rb.error:
        if (ra.error, ra.data_part) != (rb.error, rb.data_part) or (ra.data_part and za.read(ra.data_part) != zb.read(rb.data_part)):
            out.append(f"{what}：読めない部品が変わりました（{rb.error or ra.error}）")
        return out
    parts = [p for p in (ra.part, ra.data_part) if p]
    changed = [p for p in parts if za.read(p) != zb.read(p)]
    pics = [p for p in ra.images if za.has(p) and zb.has(p) and za.read(p) != zb.read(p)]
    if list(ra.images) != list(rb.images) and len(ra.images) == len(rb.images):
        pics.append("付け替え")  # 控えの絵の参照が、別の部品（分けた絵）に替わった
    if changes == "any" and ra.control is not None and rb.control is not None:
        # 書いてよい部品だが、書いた値が渡されていない：変わってよいのは値（Value・Position）だけ
        changes = {k: rb.control.get(k) for k in ("Value", "Position")
                   if any(n == k for n, _b, _s in rb.control.block.layout.data) and ra.control.get(k) != rb.control.get(k)}
    if not changes:
        if changed or pics:
            out.append(f"{what}：書き込みを指示していないのに、部品の中身か控えの絵が変わりました")
        return out
    kind = _struct_kind(ra)
    if ra.persistence == "persistPropertyBag":
        xa, xb = za.read(ra.part), zb.read(rb.part)
        strip = re.compile(rb'[ \t]*<(?:\w+:)?ocxPr\b[^>]*?\b(?:\w+:)?name="(?:' + b"|".join(re.escape(k.encode()) for k in changes)
                           + rb')"[^>]*/>')
        if strip.sub(b"", xa) != strip.sub(b"", xb):
            out.append(f"{what}：性質の袋の、指示した性質のほかが変わりました")
        for k, v in changes.items():
            got = rb.control.get(k)
            want = ax.default_of(rb.control.struct, k) if v in (None, "") else v
            if got != want:
                out.append(f"{what}：{k} が、指示した {want!r} ではなく {got!r} です")
    elif ra.persistence == "persistStorage":
        from .cfb import CFB
        from .cfbwrite import check

        da, db = za.read(ra.data_part), zb.read(rb.data_part)
        ca, cb = CFB(da), CFB(db)
        if set(ca.streams()) != set(cb.streams()):
            out.append(f"{what}：入れ物の中のストリームの構成が変わりました")
        for n in ca.streams():
            if n != ra.stream and n in cb.streams() and ca.read(n) != cb.read(n):
                out.append(f"{what}：入れ物の中の {n} が変わりました")
        out += [f"{what}：{e}" for e in check(db)]
        out += [f"{what}：{e}" for e in ax.verify(ca.read(ra.stream), cb.read(rb.stream), kind, changes)]
    else:
        da, db = za.read(ra.data_part), zb.read(rb.data_part)
        if da[:ra.start] != db[:rb.start]:
            out.append(f"{what}：部品の頭の CLSID が変わりました")
        ea = ra.start + (ra.length if ra.length is not None else len(da) - ra.start)
        eb = rb.start + (rb.length if rb.length is not None else len(db) - rb.start)
        out += [f"{what}：{e}" for e in ax.verify(da[ra.start:ea], db[rb.start:eb], kind, changes)]
        if da[ea:] != db[eb:]:
            out.append(f"{what}：部品の後ろのバイトが変わりました")
    # 控えの絵：書いた後の状態を、todoku の描き方で描いたものと、バイトまで同じであること（形の検査も）
    if len(ra.images) != len(rb.images):
        out.append(f"{what}：控えの絵の数が変わりました")
    for p0, p in zip(ra.images, rb.images):
        if not zb.has(p):
            out.append(f"{what}：控えの絵 {p} がなくなりました")
            continue
        if p0 != p and (za.has(p) or not zb.has(p0) or za.read(p0) != zb.read(p0)):
            # 共有していた絵から分けたとき：分けた絵は新しい部品で、元の絵（ほかの部品が使う）は変わっていない
            out.append(f"{what}：控えの絵の付け替え（{p0} → {p}）が、共有していた絵を分けたものになっていません")
        new = zb.read(p)
        errs = axdraw.check_picture(new)
        out += [f"{what}：控えの絵 {p}：{e}" for e in errs]
        try:
            want = axdraw.picture(rb.control, za.read(p0), items)
        except ValueError as e:
            out.append(f"{what}：控えの絵 {p} を描き直せません（{e}）")
            continue
        if new != want:
            out.append(f"{what}：控えの絵 {p} が、書いた後の状態の絵になっていません")
        if rb.kind in TEXT_KINDS + LIST_KINDS and rb.control.value:
            shown = "".join(axdraw.picture_texts(new)).replace(" ", "")
            if rb.control.value.replace("\r\n", "").replace("\n", "").replace(" ", "")[:20] not in shown and \
                    rb.control.get("PasswordChar") == "":
                out.append(f"{what}：控えの絵 {p} に、書いた文字が描かれていません")
    return out


def _ax_ref_parts(ref: AxRef) -> set[str]:
    return {p for p in (ref.part, ref.data_part, *ref.images) if p}


def rels_added_only(old: bytes, new: bytes, added) -> list[str]:
    """関係の部品が、絵の関係（分けた控えの絵 added への）を足しただけか。"""
    extra = re.findall(rb'<Relationship Id="rIdAx\d+" Type="[^"]*/image" Target="([^"]*)"/>', new)
    stripped = re.sub(rb'<Relationship Id="rIdAx\d+" Type="[^"]*/image" Target="[^"]*"/>', b"", new)
    if stripped != old:
        return ["関係の部品（.rels）が、分けた控えの絵への関係を足すほかにも変わりました"]
    names = {posixpath.basename(n) for n in added}
    return [f"関係の部品（.rels）に、分けた絵でない {t.decode()} への関係が足されました"
            for t in extra if posixpath.basename(t.decode()) not in names]


def _owner_parts(ref: AxRef) -> set[str]:
    """控えの絵を指している XML の部品と、その関係の部品（絵を分けたときに変わってよいもの）。"""
    out = set()
    for _k, owner, _key, _pic in ref.pic_refs:
        out |= {owner, _rels_path(owner)}
    return out


def prove_excel(a, b, changed_entries: list[str], allowed_controls, added: list[str] = ()) -> tuple[list[dict], list[str], set[str]]:
    """Excel（.xlsx）の ActiveX の部品の証明：(部品の変化, 想定外, 確かめた部品)。"""
    from .formula.parser import addr
    from .xlsx.calc import values_equal

    changes_out: list[dict] = []
    bad: list[str] = []
    seen_parts: set[str] = set()
    for sa in a.sheets:
        if not b.has_sheet(sa.name):
            continue
        sb = b.sheet(sa.name)
        after = {c.shape_id: c for c in sb.controls}
        for ca in sa.controls:
            ra = getattr(ca, "ax", None)
            if ra is None:
                continue
            cb = after.get(ca.shape_id)
            rb = getattr(cb, "ax", None) if cb is not None else None
            if rb is None:
                bad.append(f"{sa.name}: ActiveX の部品「{ca.name}」がなくなりました")
                continue
            seen_parts |= _ax_ref_parts(ra) | _ax_ref_parts(rb)
            want = None
            if isinstance(allowed_controls, dict):
                w = allowed_controls.get((sa.name, ca.name))
                want = w.get("ax") if isinstance(w, dict) else None
            elif allowed_controls is not None and (sa.name, ca.name) in allowed_controls:
                want = "any"
            items = items_excel(b, sb, cb) if rb.kind in LIST_KINDS else None
            problems = verify_parts(a.zip, b.zip, ra, rb, want, items)
            state_a = ax.state(ra.control) if ra.control is not None else {}
            state_b = ax.state(rb.control) if rb.control is not None else {}
            if state_a != state_b:
                changes_out.append({"sheet": sa.name, "control": ca.name, "kind": "control", "link": ca.link,
                                    "before": _state_text(state_a), "after": _state_text(state_b)})
                if want is None:
                    bad.append(f"{sa.name}: 指示していない ActiveX の部品「{ca.name}」が変わりました")
            bad += [f"{sa.name}: {p}" for p in problems]
            # つながったセルと、部品の値が合っていること（書き込みで変えた部品）
            if want and state_a != state_b and cb.link_cell and b.has_sheet(cb.link_cell[0]) and rb.control is not None:
                lc = b.sheet(cb.link_cell[0]).cells.get(cb.link_cell[1:])
                lv = lc.value if lc is not None else None
                keys = list(want) if isinstance(want, dict) else ["Position" if rb.kind in NUM_KINDS else "Value"]
                expect = _cell_value_for(rb, {k: rb.control.get(k) for k in keys}, items,
                                         _raw_items(b, sb, cb) if items is not None else None)
                if not (values_equal(lv, expect) and type(lv) is type(expect)) and not (
                        isinstance(expect, str) and to_str(lv) == expect and lv is not None and not isinstance(lv, bool)):
                    bad.append(f"{sa.name}: ActiveX の部品「{ca.name}」と、つながったセル {addr(*cb.link_cell[1:])} の値"
                               f"（{lv!r}）が合いません（部品の値 {expect!r}）")
    owners = {_rels_path(o) for sh in a.sheets for c in sh.controls if getattr(c, "ax", None) is not None
              for _k, o, _key, _p in c.ax.pic_refs}
    for n in changed_entries:
        if n in owners:
            bad += [f"{n}: {e}" for e in rels_added_only(a.zip.read(n), b.zip.read(n), added)]
    for n in list(changed_entries) + list(added):
        if ("/activeX/" in n or "/media/" in n) and n not in seen_parts:
            bad.append(f"{n}: ActiveX の部品のものでない部品（絵など）が変わったか、足されました")
    return changes_out, bad, seen_parts


def to_str(v) -> str:
    from .formula.values import to_text

    return to_text(v)


def _state_text(st: dict) -> str:
    if not st:
        return ""
    k = st.get("kind")
    if "checked" in st:
        return {True: "選んだ", False: "選んでいない", None: "どちらでもない"}[st["checked"]]
    if "position" in st:
        return f"値 {st['position']}"
    if "value" in st:
        return f"「{st['value']}」"
    return k or ""


# ====================================================================== Word（.docx・.docm・.doc）
WORD_KIND = {"text": "text", "checkbox": "checkbox", "option": "option", "select": "dropdown", "combo": "combo",
             "spin": "spin", "multiselect": "dropdown"}


def word_control(doc, container, node, part: str, p, run, hidden: bool) -> None:
    """Word の w:object・w:pict の中の w:control（ActiveX の部品）を読む。値を書ける部品は文書の入力欄（Control）に加え、
    段落にはその字（窓では［名前＝値］・☑見出し など）を置く。見出しだけの部品（ラベル・ボタン）は、段落の中の字として示す。"""
    from .docx.document import Seg

    rels = rels_of(doc.zip, part)
    rid = node.get("id")
    typ, target = rels.get(rid or "", ("", ""))
    ref = load_ooxml(doc.zip, target, "docx", node.get("name") or "") if target else \
        AxRef("docx", node.get("name") or "", error="部品の定義（activeX*.xml）への関係がありません")
    for im in container.iter("imagedata"):
        t, img = rels.get(im.get("id") or "", ("", ""))
        if t == "image" and img and doc.zip.has(img) and img not in ref.images:
            ref.images.append(img)
            ref.pic_refs.append(("word", part, rid or "", img))
    _add_word(doc, ref, node, part, p, run, hidden, len(p.segs))


def _add_word(doc, ref: AxRef, node, part: str, p, run, hidden: bool, at: int, cpref=None) -> None:
    """読んだ部品を、段落 p の at 番目の字の所に置く（.docx・.doc 共通）。"""
    from .docx.document import Control, Seg

    fk = field_kind(ref) if ref.control is not None else "other"
    if fk in WORD_KIND:
        c = Control(WORD_KIND[fk], "activex", ref.name, node, part, p)
        c.ax = ref
        c.value = value_of(ref)
        c.checked = bool(ref.control.checked) if fk in ("checkbox", "option") else False
        c.maxlen = ref.control.get("MaxLength") or None if ref.kind == "TextBox" else None
        c.help = ref.control.caption if ref.kind in CAPTION_KINDS else ""
        seg = Seg(_word_seg_text(ref), "field", cpref if cpref is not None else node, run, hidden, field=c)
        p.segs.insert(at, seg)
        c.result_segs = [seg]
        doc.controls.append(c)
    else:
        p.segs.insert(at, Seg(outline(ref) if ref.kind != "Label" or ref.control is None else ref.caption, "other",
                              cpref if cpref is not None else node, run, hidden))
    if not hasattr(doc, "ax_refs"):
        doc.ax_refs = []
    doc.ax_refs.append(ref)


def _word_seg_text(ref: AxRef) -> str:
    c = ref.control
    if ref.kind in TEXT_KINDS + LIST_KINDS:
        return c.value.replace("\r\n", "⏎").replace("\n", "⏎").replace("\r", "⏎")
    if ref.kind in NUM_KINDS:
        return str(c.get("Position"))
    return c.caption


def word_mark(c, name: str) -> str:
    """窓の段落の中の、部品の示し方。"""
    ref = c.ax
    if ref.control is None:
        return f"［{name}］"
    k = ref.kind
    if k in CHECK_KINDS:
        st = ref.control.checked
        return ("☑" if st else ("☐" if st is False else "▣")) + ref.control.caption + f"［{name}］"
    if k == "OptionButton":
        return ("◉" if ref.control.checked else "○") + ref.control.caption + f"［{name}］"
    if k in NUM_KINDS:
        return f"⇕［{name}＝{ref.control.get('Position')}］"
    v = _word_seg_text(ref)
    return f"［{name}＝{v}］" if v else f"［{name}］"


def word_groups(doc) -> dict[str, list]:
    """Word のオプションボタンの組（文書の中の、同じ GroupName のボタン）。"""
    out: dict[str, list] = {}
    for c in doc.controls:
        if c.source == "activex" and c.ax.kind == "OptionButton" and c.ax.control is not None:
            out.setdefault(c.ax.control.get("GroupName") or "", []).append(c)
    return out


def word_describe(doc, c) -> dict:
    """窓の欄の一覧に添える、Word の ActiveX の部品の性質。"""
    ref = c.ax
    d = describe(ref, None)
    if ref.kind == "OptionButton" and ref.control is not None:
        grp = word_groups(doc)
        members = grp.get(ref.control.get("GroupName") or "", [])
        d["group_members"] = [x.name for x in members]
        d["group_selected"] = next((x.name for x in members if x.ax.control.checked), None)
    return d


def word_plan(doc, c, value, force: bool = False, written: dict | None = None) -> list[tuple[object, dict, object]]:
    """Word の部品 c に value を書くときの (部品, 変える性質, 窓の値) の並び。オプションボタンは同じ組のほかのボタンも外す。
    written：同じ書き込みで先に書いた部品（id(部品の中身) → 変えた性質）。ボタンが選ばれているかは、その後の状態で見る。"""
    ref = c.ax
    if ref is None or ref.control is None:
        raise ax.AxError(f"「{c.name}」の中身を読めないので書けません：{getattr(ref, 'error', '') or '中身がありません'}")
    changes, _notes = ax.plan(ref.control, value, None, force)
    out = [(c, changes, value_after(ref, changes))]
    if ref.kind == "OptionButton" and changes.get("Value") == "1":
        for other in word_groups(doc).get(ref.control.get("GroupName") or "", []):
            prev = (written or {}).get(id(other.ax), {})
            on = prev["Value"] == "1" if "Value" in prev else bool(other.ax.control.checked)
            if other is not c and on:
                out.append((other, {"Value": "0"}, False))
    return out


def written_before(written: dict, ctl, changes: dict) -> bool:
    """同じ書き込みの中で、この部品をもう書いたか（同じ値なら True で、書き直さない）。違う値を2度書くことになるなら断る。

    部品を2度書くと、1度目に分けた控えの絵がどこからも指されずに残り、部品の中の位置も元の位置とずれるので、
    Excel で順にクリックしたのと同じ結果になるかを確かめられない。"""
    prev = written.get(id(ctl.ax))
    if prev is None:
        written[id(ctl.ax)] = dict(changes)
        return False
    if all(prev.get(k) == v for k, v in changes.items()):
        return True
    raise ax.AxError(f"「{ctl.name}」に、ひとつの書き込みの中で違う値を2度書くことになります"
                     "（同じ組のオプションボタンを2つ選んでいませんか）。選ぶボタンを1つにしてください")


def value_after(ref: AxRef, changes: dict):
    """書いた後の、窓の値（value_of と同じ形）。"""
    if "Position" in changes:
        return changes["Position"]
    v = changes.get("Value", "")
    if ref.kind in CHECK_KINDS + ("OptionButton",):
        return True if v == "1" else (False if v == "0" else None)
    return v.replace("\r\n", "\n") if ref.kind == "TextBox" else v


def word_fill(f, item: dict, value, force: bool = False) -> list:
    """Word（.docx）の部品に書く：部品の中身と控えの絵を書き換え、Filler（docx.fill）の ax_parts に入れる。書いた (id, 値) の並び。"""
    c = item["control"]
    doc = f.doc
    if not hasattr(f, "ax_parts"):
        f.ax_parts = {}
        f.ax_added = {}
    users: dict[str, set] = {}
    for r in getattr(doc, "ax_refs", []):
        for img in r.images:
            users.setdefault(img, set()).add(id(r))
    shared = {img for img, u in users.items() if len(u) > 1}
    out = []
    ids = {id(it.get("control")): k for k, it in f.sp.items.items() if it.get("control") is not None}
    if not hasattr(f, "ax_written"):
        f.ax_written = {}
    for ctl, changes, shown in word_plan(doc, c, value, force, f.ax_written):
        if written_before(f.ax_written, ctl, changes):
            if ctl is c:
                out.append((ids.get(id(ctl), ctl.name), shown))
            continue
        if all(ctl.ax.control.get(k) == v or (v in (None, "") and ctl.ax.control.get(k) == ax.default_of(ctl.ax.control.struct, k))
               for k, v in changes.items()):
            if ctl is c:
                out.append((ids.get(id(ctl), ctl.name), shown))
            continue
        new = new_parts(doc.zip, ctl.ax, changes, None, f.ax_parts, shared, f.ax_added)
        f.ax_parts.update(new)
        out.append((ids.get(id(ctl), ctl.name), shown))
    return out


def verify_word_refs(a_refs: list, b_refs: list, za, zb, written: dict, ids: dict) -> list[str]:
    """Word（.docx）の部品の証明：書き込み前後の部品を順に比べる。written は窓の欄の id → 書いた値。"""
    bad = []
    if len(a_refs) != len(b_refs):
        return [f"ActiveX の部品の数が変わりました（{len(a_refs)} → {len(b_refs)}）"]
    for ra, rb in zip(a_refs, b_refs):
        fid = ids.get(id(ra))
        changes = None
        if fid is not None and fid in written and ra.control is not None:
            try:
                changes = ax.plan(ra.control, written[fid], None, force=True)[0]
            except ax.AxError as e:
                bad.append(f"{ra.label}「{ra.name}」：書いた値 {written[fid]!r} を、部品の決まりで読めません（{e}）")
                continue
        bad += verify_parts(za, zb, ra, rb, changes)
    return bad


def reading_check(refs: list, win: dict):
    """読み落とし（ActiveX の部品の文字）：部品の中身を直接読んで拾った見出しと文字が、窓に出ているか。"""
    from .xlsx.audit import Check

    if not refs:
        return None
    raw = [(r.name, t) for r in refs for t in texts(r)]
    shown = "\n".join(line for st in win.get("stories", []) for line in st.get("reading", []))
    shown += "\n" + "\n".join(str(f.get("value") or "") + "\n" + str(f.get("caption") or "") for f in win.get("fields", []))
    flat = re.sub(r"[\s　⏎]", "", shown)
    missing = [f"{n}: {t[:40]!r}" for n, t in raw if re.sub(r"[\s　]", "", t) not in flat]
    unreadable = [f"{r.name}: {r.error}" for r in refs if r.error and r.control is None and r.frame is None]
    return Check("読み落とし（ActiveX の部品の文字）", not missing and not unreadable,
                 f"ActiveX の部品 {len(refs)} 個の見出しと文字 {len(raw) - len(missing)}/{len(raw)} 項目が窓に出ている"
                 + (f"（中身を読めない部品 {len(unreadable)} 個）" if unreadable else ""), missing + unreadable)
