"""ActiveX の部品（Microsoft Forms 2.0）の、保存された形を読み、値の性質だけを書き換える（MS-OFORMS）。

部品：テキストボックス・チェックボックス・オプションボタン・トグルボタン・コンボボックス・リストボックス（この6つは同じ形
MorphData で、DisplayStyle で種類が決まる）・スピンボタン・スクロールバー・ラベル・コマンドボタン・イメージ・フレーム。

部品の保存の形（MS-OFORMS 2.1.1.2・2.2）は、どれも次の順：
  MinorVersion（1）・MajorVersion（1）・cb（2：PropMask・DataBlock・ExtraDataBlock の大きさの和）・
  PropMask（4 か 8：既定値でない性質の印）・DataBlock（4バイト以下の性質）・ExtraDataBlock（4バイトより大きい性質）・
  StreamData（絵）・TextProps（字の書式）・そのほか（コンボボックスとリストボックスの列の幅 rgColumnInfo）
- DataBlock の性質は、部品の始まりから数えて、その大きさ（2 か 4）で割り切れる位置に置く。前に詰め物を入れる（中身は不定）。
  DataBlock の終わりは、4 で割り切れる大きさまで 0 を足す（2.1.1.2.4）
- 文字の性質（Caption・Value・GroupName など）は、DataBlock に長さと圧縮の印（CountOfBytesWithCompressionFlag：
  下の31ビットが長さ、いちばん上のビットが圧縮）、ExtraDataBlock に文字（fmString）。文字は 4 で割り切れる長さまで詰める。
  どの字も上のバイトが 0 なら、上のバイトを除いて1字1バイトで保存する（圧縮。2.4.14）。日本語は UTF-16 のまま
- 既定値（file format default）と同じ性質は保存しない（PropMask の印を消す）

書き込みは、指示した性質（Value・Position）だけを変える。DataBlock の並びが変わらなければ、変わるのはその性質のバイトと、
文字の長さに応じた cb・ExtraDataBlock の後ろへのずれだけ。性質が既定値になる・既定値でなくなるときは、仕様のとおりに
並べ直す（前の性質の位置が変わらないところまでは、詰め物も含めて元のバイトのまま）。StreamData から後ろは1バイトも変えない。
書いたあとは必ず読み直し、指示した性質のほかが変わっていないことを確かめる（verify）。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field


class AxError(ValueError):
    pass


# ====================================================================== 部品の種類
# classid（CLSID）。MS-OFORMS には一覧がないので、ファイルの中の ProgID（CompObj・.xls の strClass・.doc の CONTROL の命令）
# とも照らす。MorphData の6つの部品は、classid より DisplayStyle（2.5.20）を正とする
CLSIDS = {
    "D7053240-CE69-11CD-A777-00DD01143C57": "CommandButton",
    "8BD21D10-EC42-11CE-9E0D-00AA006002F3": "TextBox",
    "8BD21D20-EC42-11CE-9E0D-00AA006002F3": "ListBox",
    "8BD21D30-EC42-11CE-9E0D-00AA006002F3": "ComboBox",
    "8BD21D40-EC42-11CE-9E0D-00AA006002F3": "CheckBox",
    "8BD21D50-EC42-11CE-9E0D-00AA006002F3": "OptionButton",
    "8BD21D60-EC42-11CE-9E0D-00AA006002F3": "ToggleButton",
    "978C9E23-D4B0-11CE-BF2D-00AA003F40D0": "Label",
    "79176FB0-B7F2-11CE-97EF-00AA006D2776": "SpinButton",
    "DFD181E0-5E2F-11CE-A449-00AA004A803D": "ScrollBar",
    "4C599241-6926-101B-9992-00000B65C6F9": "Image",
    "6E182020-F460-11CE-9BCD-00AA00608E01": "Frame",
    "46E31370-3F7A-11CE-BED6-00AA00611080": "MultiPage",
    "EAE50EB0-4A62-11CE-BED6-00AA00611080": "TabStrip",
}
PROGIDS = {f"Forms.{k}.1": k for k in set(CLSIDS.values())}
MORPH = ("TextBox", "ListBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton")
# 保存の形（構造）：MorphData の6つはひとつの形
STRUCT = {**{k: "MorphData" for k in MORPH}, "CommandButton": "CommandButton", "Label": "Label", "Image": "Image",
          "SpinButton": "SpinButton", "ScrollBar": "ScrollBar"}
# DisplayStyle（2.5.20.1）→ 種類。7 は「選択肢から選ぶだけの」コンボボックス（fmDisplayStyleDropList）
DISPLAY = {1: "TextBox", 2: "ListBox", 3: "ComboBox", 4: "CheckBox", 5: "OptionButton", 6: "ToggleButton", 7: "ComboBox"}
# 値を書ける部品
VALUE_KINDS = ("TextBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox", "ListBox", "SpinButton", "ScrollBar")
KIND_JA = {"TextBox": "テキストボックス", "CheckBox": "チェックボックス", "OptionButton": "オプションボタン",
           "ToggleButton": "トグルボタン", "ComboBox": "コンボボックス", "ListBox": "リストボックス",
           "SpinButton": "スピンボタン", "ScrollBar": "スクロールバー", "Label": "ラベル", "CommandButton": "コマンドボタン",
           "Image": "イメージ", "Frame": "フレーム", "MultiPage": "マルチページ", "TabStrip": "タブ"}


def guid_str(b: bytes) -> str:
    """GUID の16バイト（MS-DTYP 2.3.4.2：前の3つは小さい順）を、{…} を除いた大文字の文字にする。"""
    d1, d2, d3 = struct.unpack_from("<IHH", b, 0)
    return f"{d1:08X}-{d2:04X}-{d3:04X}-{b[8:10].hex().upper()}-{b[10:16].hex().upper()}"


def guid_bytes(s: str) -> bytes:
    s = s.strip("{}")
    p = s.split("-")
    return struct.pack("<IHH", int(p[0], 16), int(p[1], 16), int(p[2], 16)) + bytes.fromhex(p[3] + p[4])


def kind_of(clsid: str | None = None, progid: str | None = None) -> str | None:
    """classid か ProgID から部品の種類。どちらも分からなければ None（Forms 2.0 でない ActiveX）。"""
    if clsid:
        k = CLSIDS.get(clsid.strip("{}").upper())
        if k:
            return k
    if progid:
        p = progid.strip()
        k = PROGIDS.get(p) or next((v for key, v in PROGIDS.items() if key.lower() == p.lower()), None)
        if k:
            return k
    return None


# ====================================================================== 文字（2.4.14）
def encode_string(s: str) -> tuple[int, bytes]:
    """文字 → (CountOfBytesWithCompressionFlag の値, fmString のバイト列)。どの字も上のバイトが 0 なら圧縮する。"""
    raw = s.encode("utf-16-le")
    if raw and all(raw[i] == 0 for i in range(1, len(raw), 2)):
        out = raw[0::2]
        return len(out) | 0x80000000, out
    return len(raw), raw


def decode_string(count: int, raw: bytes) -> str:
    if count & 0x80000000:
        return raw.decode("latin-1")
    return raw.decode("utf-16-le", errors="replace")


def _pad4(n: int) -> int:
    return (n + 3) & ~3


# ====================================================================== 保存の形の表（2.2〜2.3）
S = "S"  # 文字の性質：DataBlock に長さと圧縮の印（4バイト）、ExtraDataBlock に文字


@dataclass(frozen=True)
class Layout:
    """1つの構造の、PropMask のビットと、DataBlock・ExtraDataBlock の並び（仕様書の各節の図と欄の説明の順）。"""

    name: str
    mask_bytes: int
    data: tuple  # (性質, PropMask のビット, 大きさ（1・2・4・8 か S）)
    extra: tuple  # (性質, PropMask のビット, 大きさ（8・16 か S）)
    stream: tuple = ()  # StreamData の絵の性質（GuidAndPicture）の順：(性質, ビット)
    textprops: bool = False  # 後ろに TextProps があるか
    flags: tuple = ()  # データを持たないビット：(性質, ビット)
    reserved_one: tuple = ()  # 1 にしておくビット（MorphData の Reserved）


LAYOUTS: dict[str, Layout] = {
    # 2.2.1 CommandButton
    "CommandButton": Layout(
        "CommandButton", 4,
        (("ForeColor", 0, 4), ("BackColor", 1, 4), ("VariousPropertyBits", 2, 4), ("Caption", 3, S),
         ("PicturePosition", 4, 4), ("MousePointer", 6, 1), ("Picture", 7, 2), ("Accelerator", 8, 2), ("MouseIcon", 10, 2)),
        (("Caption", 3, S), ("Size", 5, 8)),
        stream=(("Picture", 7), ("MouseIcon", 10)), textprops=True, flags=(("TakeFocusOnClick", 9),)),
    # 2.2.3 Image
    "Image": Layout(
        "Image", 4,
        (("BorderColor", 3, 4), ("BackColor", 4, 4), ("BorderStyle", 5, 1), ("MousePointer", 6, 1),
         ("PictureSizeMode", 7, 1), ("SpecialEffect", 8, 1), ("Picture", 10, 2), ("PictureAlignment", 11, 1),
         ("VariousPropertyBits", 13, 4), ("MouseIcon", 14, 2)),
        (("Size", 9, 8),),
        stream=(("Picture", 10), ("MouseIcon", 14)), flags=(("AutoSize", 2), ("PictureTiling", 12))),
    # 2.2.4 Label
    "Label": Layout(
        "Label", 4,
        (("ForeColor", 0, 4), ("BackColor", 1, 4), ("VariousPropertyBits", 2, 4), ("Caption", 3, S),
         ("PicturePosition", 4, 4), ("MousePointer", 6, 1), ("BorderColor", 7, 4), ("BorderStyle", 8, 2),
         ("SpecialEffect", 9, 2), ("Picture", 10, 2), ("Accelerator", 11, 2), ("MouseIcon", 12, 2)),
        (("Caption", 3, S), ("Size", 5, 8)),
        stream=(("Picture", 10), ("MouseIcon", 12)), textprops=True),
    # 2.2.5 MorphData（テキストボックス・リストボックス・コンボボックス・チェックボックス・オプションボタン・トグルボタン）
    "MorphData": Layout(
        "MorphData", 8,
        (("VariousPropertyBits", 0, 4), ("BackColor", 1, 4), ("ForeColor", 2, 4), ("MaxLength", 3, 4),
         ("BorderStyle", 4, 1), ("ScrollBars", 5, 1), ("DisplayStyle", 6, 1), ("MousePointer", 7, 1),
         ("PasswordChar", 9, 2), ("ListWidth", 10, 4), ("BoundColumn", 11, 2), ("TextColumn", 12, 2),
         ("ColumnCount", 13, 2), ("ListRows", 14, 2), ("cColumnInfo", 15, 2), ("MatchEntry", 16, 1), ("ListStyle", 17, 1),
         ("ShowDropButtonWhen", 18, 1), ("DropButtonStyle", 20, 1), ("MultiSelect", 21, 1), ("Value", 22, S),
         ("Caption", 23, S), ("PicturePosition", 24, 4), ("BorderColor", 25, 4), ("SpecialEffect", 26, 4),
         ("MouseIcon", 27, 2), ("Picture", 28, 2), ("Accelerator", 29, 2), ("GroupName", 32, S)),
        (("Size", 8, 8), ("Value", 22, S), ("Caption", 23, S), ("GroupName", 32, S)),
        stream=(("MouseIcon", 27), ("Picture", 28)), textprops=True, reserved_one=(31,)),
    # 2.2.7 ScrollBar
    "ScrollBar": Layout(
        "ScrollBar", 4,
        (("ForeColor", 0, 4), ("BackColor", 1, 4), ("VariousPropertyBits", 2, 4), ("MousePointer", 4, 1), ("Min", 5, 4),
         ("Max", 6, 4), ("Position", 7, 4), ("PrevEnabled", 9, 4), ("NextEnabled", 10, 4), ("SmallChange", 11, 4),
         ("LargeChange", 12, 4), ("Orientation", 13, 4), ("ProportionalThumb", 14, 2), ("Delay", 15, 4),
         ("MouseIcon", 16, 2)),
        (("Size", 3, 8),),
        stream=(("MouseIcon", 16),)),
    # 2.2.8 SpinButton
    "SpinButton": Layout(
        "SpinButton", 4,
        (("ForeColor", 0, 4), ("BackColor", 1, 4), ("VariousPropertyBits", 2, 4), ("Min", 5, 4), ("Max", 6, 4),
         ("Position", 7, 4), ("PrevEnabled", 8, 4), ("NextEnabled", 9, 4), ("SmallChange", 10, 4), ("Orientation", 11, 4),
         ("Delay", 12, 4), ("MouseIcon", 13, 2), ("MousePointer", 14, 1)),
        (("Size", 3, 8),),
        stream=(("MouseIcon", 13),)),
    # 2.3 TextProps（字の書式）
    "TextProps": Layout(
        "TextProps", 4,
        (("FontName", 0, S), ("FontEffects", 1, 4), ("FontHeight", 2, 4), ("FontCharSet", 4, 1),
         ("FontPitchAndFamily", 5, 1), ("ParagraphAlign", 6, 1), ("FontWeight", 7, 2)),
        (("FontName", 0, S),)),
    # 2.2.5.6 MorphDataColumnInfo（コンボボックス・リストボックスの列の幅）
    "ColumnInfo": Layout("ColumnInfo", 4, (("ColumnWidth", 0, 4),), ()),
    # 2.2.10.1 FormControl（フレームの "f" の頭）
    "Form": Layout(
        "Form", 4,
        (("BackColor", 1, 4), ("ForeColor", 2, 4), ("NextAvailableID", 3, 4), ("BooleanProperties", 6, 4),
         ("BorderStyle", 7, 1), ("MousePointer", 8, 1), ("ScrollBars", 9, 1), ("GroupCnt", 13, 4), ("MouseIcon", 15, 2),
         ("Cycle", 16, 1), ("SpecialEffect", 17, 1), ("BorderColor", 18, 4), ("Caption", 19, S), ("Font", 20, 2),
         ("Picture", 21, 2), ("Zoom", 22, 4), ("PictureAlignment", 23, 1), ("PictureSizeMode", 25, 1),
         ("ShapeCookie", 26, 4), ("DrawBuffer", 27, 4)),
        (("DisplayedSize", 10, 8), ("LogicalSize", 11, 8), ("ScrollPosition", 12, 8), ("Caption", 19, S)),
        flags=(("PictureTiling", 24),)),
    # 2.2.10.12 OleSiteConcreteControl（フレームの中の部品ごとの、名前・番号・位置など）
    "Site": Layout(
        "Site", 4,
        (("Name", 0, S), ("Tag", 1, S), ("ID", 2, 4), ("HelpContextID", 3, 4), ("BitFlags", 4, 4),
         ("ObjectStreamSize", 5, 4), ("TabIndex", 6, 2), ("ClsidCacheIndex", 7, 2), ("GroupID", 9, 2),
         ("ControlTipText", 11, S), ("RuntimeLicKey", 12, S), ("ControlSource", 13, S), ("RowSource", 14, S)),
        (("Name", 0, S), ("Tag", 1, S), ("Position", 8, 8), ("ControlTipText", 11, S), ("RuntimeLicKey", 12, S),
         ("ControlSource", 13, S), ("RowSource", 14, S))),
}

# 符号つきの性質（ほかは符号なし）
SIGNED = {"Min", "Max", "Position", "SmallChange", "LargeChange", "PrevEnabled", "NextEnabled", "TextColumn",
          "ColumnCount", "ColumnWidth", "ID", "HelpContextID", "TabIndex", "Zoom", "Orientation"}
# 1字の性質（Unicode の1字）
CHAR = {"Accelerator", "PasswordChar"}

# 既定値（2.5 の各節の file format default）。部品の種類ごとに違うものは DEFAULT_BY_KIND
DEFAULT = {
    "Caption": "", "Value": "", "GroupName": "", "FontName": "MS Sans Serif", "Name": "", "Tag": "", "ControlTipText": "",
    "RuntimeLicKey": "", "ControlSource": "", "RowSource": "",
    "MaxLength": 0, "BorderStyle": 0, "ScrollBars": 0, "DisplayStyle": 1, "MousePointer": 0, "PasswordChar": "",
    "ListWidth": 0, "BoundColumn": 1, "TextColumn": -1, "ColumnCount": 1, "ListRows": 8, "cColumnInfo": 0,
    "MatchEntry": 2, "ListStyle": 0, "ShowDropButtonWhen": 0, "DropButtonStyle": 1, "MultiSelect": 0,
    "PicturePosition": 0x00070001, "Accelerator": "", "Min": 0, "Position": 0, "PrevEnabled": 1, "NextEnabled": 1,
    "SmallChange": 1, "LargeChange": 1, "Orientation": -1, "ProportionalThumb": 0xFFFF, "Delay": 50,
    "FontEffects": 0, "FontHeight": 160, "FontCharSet": 1, "FontPitchAndFamily": 0, "ParagraphAlign": 1,
    "FontWeight": 400, "PictureSizeMode": 0, "PictureAlignment": 2, "ColumnWidth": -1,
    "ID": 0, "HelpContextID": 0, "BitFlags": 0x33, "ObjectStreamSize": 0, "TabIndex": -1, "ClsidCacheIndex": 0x7FFF,
    "GroupID": 0, "BooleanProperties": 0x04, "Zoom": 100,
}
DEFAULT_BY_KIND = {
    # 2.5.96 VariousPropertyBits・2.5.3 BackColor・2.5.31 ForeColor・2.5.7 BorderColor・2.5.8 BorderStyle・2.5.80 SpecialEffect・2.5.46 Max
    "MorphData": {"VariousPropertyBits": 0x2C80081B, "BackColor": 0x80000005, "ForeColor": 0x80000008,
                  "BorderColor": 0x80000006, "SpecialEffect": 2},
    "CommandButton": {"VariousPropertyBits": 0x0000001B, "BackColor": 0x8000000F, "ForeColor": 0x80000012},
    "Label": {"VariousPropertyBits": 0x0080001B, "BackColor": 0x8000000F, "ForeColor": 0x80000012,
              "BorderColor": 0x80000006, "SpecialEffect": 0},
    "Image": {"VariousPropertyBits": 0x0000001B, "BackColor": 0x8000000F, "BorderColor": 0x80000006, "BorderStyle": 1,
              "SpecialEffect": 0},
    "ScrollBar": {"VariousPropertyBits": 0x0000001B, "BackColor": 0x8000000F, "ForeColor": 0x80000012, "Max": 32767},
    "SpinButton": {"VariousPropertyBits": 0x0000001B, "BackColor": 0x8000000F, "ForeColor": 0x80000012, "Max": 100},
    "Form": {"BackColor": 0x8000000F, "ForeColor": 0x80000012, "BorderColor": 0x80000012, "SpecialEffect": 0},
}

# VariousPropertiesBitfield（2.5.96.1）のビット
VPB = {"Enabled": 1, "Locked": 2, "BackStyle": 3, "ColumnHeads": 10, "IntegralHeight": 11, "MatchRequired": 12,
       "Alignment": 13, "Editable": 14, "WordWrap": 23, "AutoSize": 28, "MultiLine": 31}


def default_of(struct_name: str, prop: str):
    return DEFAULT_BY_KIND.get(struct_name, {}).get(prop, DEFAULT.get(prop))


# ====================================================================== 読む
@dataclass
class Block:
    """1つの構造（PropMask・DataBlock・ExtraDataBlock）を読んだもの。位置はすべて、渡したバイト列の頭からの位置。"""

    layout: Layout
    base: int
    minor: int
    major: int
    cb: int
    mask: int
    values: dict = field(default_factory=dict)  # 保存されている性質 → 値
    data_at: dict = field(default_factory=dict)  # DataBlock の欄 → (位置, 大きさ)
    extra_at: dict = field(default_factory=dict)  # ExtraDataBlock の欄 → (位置, 中身の大きさ, 詰め物を含む大きさ)
    counts: dict = field(default_factory=dict)  # 文字の性質 → CountOfBytesWithCompressionFlag の値
    data_start: int = 0
    data_end: int = 0  # DataBlock の終わり（ExtraDataBlock の始まり）
    end: int = 0  # ExtraDataBlock の終わり

    def stored(self, prop: str) -> bool:
        bit = next((b for n, b, _ in self.layout.data + self.layout.extra if n == prop), None)
        if bit is None:
            bit = next((b for n, b in self.layout.flags if n == prop), None)
        return bit is not None and bool(self.mask >> bit & 1)

    def get(self, prop: str, struct_name: str | None = None):
        if prop in self.values:
            return self.values[prop]
        return default_of(struct_name or self.layout.name, prop)


def _num(buf: bytes, at: int, size: int, signed: bool) -> int:
    return int.from_bytes(buf[at:at + size], "little", signed=signed)


def parse_block(buf: bytes, base: int, layout: Layout) -> Block:
    """buf の base から、layout の形の構造を読む。cb と並びが合わなければ AxError。"""
    if base + 4 + layout.mask_bytes > len(buf):
        raise AxError(f"{layout.name} の頭が途中で切れています")
    minor, major, cb = struct.unpack_from("<BBH", buf, base)
    mask = int.from_bytes(buf[base + 4:base + 4 + layout.mask_bytes], "little")
    blk = Block(layout, base, minor, major, cb, mask)
    p = base + 4 + layout.mask_bytes
    blk.data_start = p
    for name, bit, size in layout.data:
        if not mask >> bit & 1:
            continue
        n = 4 if size == S else size
        p = base + ((p - base + n - 1) // n) * n  # 部品の始まりから数えて、大きさで割り切れる位置
        if p + n > len(buf):
            raise AxError(f"{layout.name} の {name} が途中で切れています")
        blk.data_at[name] = (p, n)
        if size == S:
            blk.counts[name] = struct.unpack_from("<I", buf, p)[0]
        elif name in CHAR:
            blk.values[name] = chr(_num(buf, p, n, False)) if _num(buf, p, n, False) else ""
        else:
            blk.values[name] = _num(buf, p, n, name in SIGNED)
        p += n
    p = blk.data_start + _pad4(p - blk.data_start)  # DataBlock の大きさは 4 で割り切れる
    blk.data_end = p
    for name, bit, size in layout.extra:
        if not mask >> bit & 1:
            continue
        if size == S:
            count = blk.counts.get(name, 0)
            n = count & 0x7FFFFFFF
            if p + n > len(buf):
                raise AxError(f"{layout.name} の {name} の文字が途中で切れています")
            blk.values[name] = decode_string(count, bytes(buf[p:p + n]))
            blk.extra_at[name] = (p, n, _pad4(n))
            p += _pad4(n)
        else:
            if p + size > len(buf):
                raise AxError(f"{layout.name} の {name} が途中で切れています")
            if size == 8:
                blk.values[name] = struct.unpack_from("<ii", buf, p)
            else:
                blk.values[name] = bytes(buf[p:p + size])
            blk.extra_at[name] = (p, size, size)
            p += size
    blk.end = p
    if cb != p - (base + 4):
        raise AxError(f"{layout.name} の大きさ（cb={cb}）が、読んだ性質の大きさ（{p - base - 4}）と合いません")
    return blk


def _skip_picture(buf: bytes, p: int, what: str) -> int:
    """GuidAndPicture（2.4.8：CLSID 16・Preamble 4・Size 4・絵）を読み飛ばす。"""
    if p + 24 > len(buf):
        raise AxError(f"{what} の絵が途中で切れています")
    pre, size = struct.unpack_from("<II", buf, p + 16)
    if pre != 0x0000746C:
        raise AxError(f"{what} の絵の印（Preamble）が違います")
    if p + 24 + size > len(buf):
        raise AxError(f"{what} の絵が途中で切れています")
    return p + 24 + size


@dataclass
class Control:
    """1つの部品を読んだもの（MS-OFORMS 2.2 の …Control）。"""

    kind: str  # TextBox など（MorphData は DisplayStyle で決めた種類）
    struct: str  # MorphData / CommandButton / Label / Image / SpinButton / ScrollBar
    data: bytes  # 部品の始まりから、読んだ終わりまで（後ろに続くバイトは含めない）
    block: Block
    text: Block | None = None  # TextProps
    columns: list = field(default_factory=list)  # rgColumnInfo（Block）
    end: int = 0  # data の中の、この部品の終わり（TextProps・rgColumnInfo の後ろ）
    rest: bytes = b""  # 部品の後ろに続く、読まなかったバイト（そのまま残す）

    def get(self, prop: str):
        return self.block.get(prop, self.struct)

    @property
    def caption(self) -> str:
        return self.get("Caption") or ""

    @property
    def value(self) -> str:
        return self.get("Value") or ""

    @property
    def various(self) -> int:
        return self.get("VariousPropertyBits") or 0

    def bit(self, name: str) -> bool:
        return bool(self.various >> VPB[name] & 1)

    @property
    def checked(self) -> bool | None:
        """チェックボックス・オプションボタン・トグルボタンの状態（2.5.95：「1」選ぶ・「0」選ばない・ほかはどちらでもない）。"""
        v = self.value
        return True if v == "1" else (False if v == "0" else None)

    @property
    def size(self) -> tuple[int, int]:
        """大きさ（HIMETRIC：0.01 mm）。"""
        return tuple(self.block.values.get("Size", (0, 0)))

    @property
    def font(self) -> dict:
        t = self.text
        g = (lambda k: t.get(k)) if t is not None else (lambda k: DEFAULT.get(k))
        return {"name": g("FontName"), "height": g("FontHeight"), "weight": g("FontWeight"), "charset": g("FontCharSet"),
                "effects": g("FontEffects"), "align": g("ParagraphAlign")}

    @property
    def style(self) -> str:
        """コンボボックスの形：combo（字を打てる）/ droplist（選ぶだけ）。"""
        return "droplist" if self.get("DisplayStyle") == 7 else "combo"


def parse(data: bytes, kind: str | None = None, start: int = 0, end: int | None = None) -> Control:
    """部品の保存の形（data の start から）を読む。kind は classid・ProgID から分かった種類（MorphData は確かめに使うだけ）。"""
    buf = bytes(data[:end]) if end is not None else bytes(data)
    st = STRUCT.get(kind or "")
    if st is None:
        raise AxError(f"{KIND_JA.get(kind or '', kind or '不明な部品')}は、値や見出しを読む形がありません")
    lay = LAYOUTS[st]
    blk = parse_block(buf, start, lay)
    p = blk.end
    m = blk.mask
    for name, bit in lay.stream:
        if m >> bit & 1:
            p = _skip_picture(buf, p, name)
    text = None
    if lay.textprops and p + 8 <= len(buf):
        text = parse_block(buf, p, LAYOUTS["TextProps"])
        p = text.end
    real = kind
    if st == "MorphData":
        ds = blk.get("DisplayStyle", st)
        real = DISPLAY.get(ds, kind if kind in MORPH else "TextBox")
    ctl = Control(real or st, st, buf, blk, text)
    if st == "MorphData":
        for _ in range(blk.get("cColumnInfo", st) or 0):
            ci = parse_block(buf, p, LAYOUTS["ColumnInfo"])
            ctl.columns.append(ci)
            p = ci.end
    ctl.end = p
    ctl.rest = buf[p:]
    return ctl


# ====================================================================== 書く
def _encode_value(name: str, size, value) -> bytes:
    """DataBlock の欄のバイト列（文字の性質は、長さと圧縮の印）。"""
    if size == S:
        return struct.pack("<I", encode_string(value)[0])
    if name in CHAR:
        return struct.pack("<H", ord(value) if value else 0)
    return int(value).to_bytes(size, "little", signed=name in SIGNED)


def rebuild(buf: bytes, blk: Block, changes: dict) -> bytes:
    """blk（buf の中の構造）の性質を changes（性質 → 新しい値）のとおりに変えた、構造（頭から ExtraDataBlock の終わりまで）のバイト列。

    既定値と同じ値は保存しない（印を消す）。DataBlock の並びが元と同じところまでは、詰め物も含めて元のバイトのまま。
    変えない文字の性質は、ExtraDataBlock の中で元の詰め物ごと移す。"""
    lay = blk.layout
    base = blk.base
    new_vals = dict(blk.values)
    mask = blk.mask
    for name, v in changes.items():
        entry = next(((b, s) for n, b, s in lay.data + lay.extra if n == name), None)
        if entry is None:
            raise AxError(f"{lay.name} に {name} という性質はありません")
        bit, _size = entry
        if v is None or v == default_of(lay.name, name) or (isinstance(v, str) and v == ""):
            mask &= ~(1 << bit)
            new_vals.pop(name, None)
        else:
            mask |= 1 << bit
            new_vals[name] = v
    for b in lay.reserved_one:
        mask |= 1 << b
    out = bytearray(buf[base:base + 4 + lay.mask_bytes])
    out[4:4 + lay.mask_bytes] = mask.to_bytes(lay.mask_bytes, "little")
    # DataBlock
    old_seq = [(n, blk.data_at[n][0] - base) for n, b, _ in lay.data if blk.mask >> b & 1]
    k = 0
    same = True
    for name, bit, size in lay.data:
        if not mask >> bit & 1:
            continue
        n = 4 if size == S else size
        rel = len(out)
        pos = ((rel + n - 1) // n) * n
        if same and k < len(old_seq) and old_seq[k] == (name, pos):
            out += buf[base + rel:base + pos]  # 元と同じ位置：詰め物も元のまま
        else:
            same = False
            out += bytes(pos - rel)
        k += 1
        if name in changes:
            out += _encode_value(name, size, new_vals[name])
        else:
            p0, n0 = blk.data_at[name]
            out += buf[p0:p0 + n0]
    data_len = len(out) - (4 + lay.mask_bytes)
    end_pad = _pad4(data_len) - data_len
    if same and k == len(old_seq) and blk.data_end - blk.data_start == data_len + end_pad:
        out += buf[base + len(out):base + len(out) + end_pad]
    else:
        out += bytes(end_pad)
    # ExtraDataBlock
    for name, bit, size in lay.extra:
        if not mask >> bit & 1:
            continue
        if name in changes or name not in blk.extra_at:
            if size == S:
                _c, raw = encode_string(new_vals[name])
                if name in blk.extra_at and blk.extra_at[name][1] == len(raw):
                    p0, n0, padded = blk.extra_at[name]
                    out += raw + buf[p0 + n0:p0 + padded]  # 同じ長さ：詰め物は元のまま（文字のバイトだけが変わる）
                else:
                    out += raw + bytes(_pad4(len(raw)) - len(raw))  # 長さが変わる：詰め物は 0（元の文字を残さない）
            elif size == 8:
                out += struct.pack("<ii", *new_vals[name])
            else:
                out += bytes(new_vals[name])
        else:
            p0, _n, padded = blk.extra_at[name]
            out += buf[p0:p0 + padded]
    struct.pack_into("<H", out, 2, len(out) - 4)
    if len(out) - 4 > 0xFFFF:
        raise AxError(f"{lay.name} が大きくなりすぎます（cb が 65535 を超えます）")
    return bytes(out)


def set_props(data: bytes, kind: str | None, changes: dict, start: int = 0, end: int | None = None) -> bytes:
    """部品（data の start〜end）の性質だけを変えたバイト列（start〜end の部分）。後ろの StreamData・TextProps・列の幅は変えない。
    書いたあとに読み直し、指示した性質になっていて、ほかの性質が変わっていないことを確かめる。"""
    seg = bytes(data[start:end] if end is not None else data[start:])
    before = parse(seg, kind)
    head = rebuild(seg, before.block, changes)
    new = head + seg[before.block.end:]
    problems = verify(seg, new, kind, changes)
    if problems:
        raise AxError("部品の書き換えを読み直すと合いません：" + "；".join(problems))
    return new


def verify(old: bytes, new: bytes, kind: str | None, changes: dict) -> list[str]:
    """old を changes のとおりに変えたのが new か：指示した性質は指示した値（既定値なら保存されていない）、
    ほかの性質は保存のされ方まで同じ、StreamData から後ろ（TextProps・列の幅・続くバイト）はバイト一致。"""
    out = []
    try:
        a = parse(old, kind)
        b = parse(new, kind)
    except AxError as e:
        return [f"部品を読めません：{e}"]
    if (a.block.minor, a.block.major) != (b.block.minor, b.block.major):
        out.append("版（MinorVersion・MajorVersion）が変わりました")
    lay = a.block.layout
    for name, bit, _size in lay.data + lay.extra:
        if name in changes:
            want = changes[name]
            got = b.get(name)
            dflt = default_of(lay.name, name)
            if want is None or want == "":
                want = dflt
            if got != want:
                out.append(f"{name} が、指示した {want!r} ではなく {got!r} です")
            if b.block.stored(name) != (want != dflt):
                out.append(f"{name} の保存のされ方（既定値なら保存しない）が仕様と違います")
            continue
        if a.block.stored(name) != b.block.stored(name) or a.block.values.get(name) != b.block.values.get(name):
            out.append(f"指示していない性質 {name} が変わりました")
        elif name in a.block.data_at and name not in a.block.counts:
            pa, na = a.block.data_at[name]
            pb, nb = b.block.data_at[name]
            if old[pa:pa + na] != new[pb:pb + nb]:
                out.append(f"指示していない性質 {name} のバイトが変わりました")
    for name, bit in lay.flags:
        if (a.block.mask >> bit & 1) != (b.block.mask >> bit & 1):
            out.append(f"指示していない性質 {name} の印が変わりました")
    known = {bit for _n, bit, _s in lay.data + lay.extra} | {bit for _n, bit in lay.flags}
    other = ~sum(1 << x for x in known) & ((1 << (8 * lay.mask_bytes)) - 1)
    if a.block.mask & other != b.block.mask & other:
        out.append("PropMask の、性質に当たらないビットが変わりました")
    if old[a.block.end:] != new[b.block.end:]:
        out.append("StreamData から後ろ（絵・字の書式・列の幅）が変わりました")
    # DataBlock：並びが元と同じところ（変えた性質の欄の手前まで）は、詰め物も含めてバイト一致
    seq_a = [(n, a.block.data_at[n][0]) for n, bt, _ in lay.data if a.block.mask >> bt & 1]
    seq_b = [(n, b.block.data_at[n][0]) for n, bt, _ in lay.data if b.block.mask >> bt & 1]
    prefix = a.block.data_end
    for i in range(max(len(seq_a), len(seq_b))):
        ea = seq_a[i] if i < len(seq_a) else None
        eb = seq_b[i] if i < len(seq_b) else None
        if ea != eb:
            prefix = min(x[1] for x in (ea, eb) if x is not None)
            break
    allowed = {2, 3} | set(range(4, a.block.data_start))  # cb と PropMask（印の変化は上で確かめた）
    for name in changes:
        if name in a.block.data_at:
            p, n = a.block.data_at[name]
            allowed.update(range(p, p + n))
    if seq_a == seq_b and a.block.data_end != b.block.data_end:
        out.append("DataBlock の大きさが変わりました")
    if old[:2] != new[:2] or any(old[i] != new[i] for i in range(a.block.data_start, min(prefix, len(old), len(new)))
                                 if i not in allowed):
        out.append("DataBlock の、指示した性質のほかのバイトが変わりました")
    return out


# ====================================================================== 入れ物（ホストごとの包み方）
def split_stream_init(blob: bytes, clsid: str | None = None) -> tuple[bytes, bytes]:
    """OOXML の persistStreamInit / persistStream の部品（activeX*.bin）：先頭の CLSID（16バイト）と、部品の保存の形。
    Excel・Word は、部品の classid と同じ CLSID を先に書いてから、部品が自分を書く（OleSaveToStream と同じ形）。
    先頭が classid と違うときは、CLSID はないものとして全体を部品とする。"""
    if len(blob) >= 16 and (clsid is None or blob[:16] == guid_bytes(clsid)):
        if clsid is not None or guid_str(blob[:16]) in CLSIDS:
            return blob[:16], blob[16:]
    return b"", blob


def compobj_progid(data: bytes) -> str:
    """CompObj（MS-OLEDS 2.3.8）の AnsiUserType の後ろの、ClipboardFormat・Reserved1（ProgID が入っている）を読む。
    仕様では Reserved1 は「無視してよい」ものなので、種類の決め手にはせず、classid が分からないときの手がかりにだけ使う。"""
    try:
        p = 28
        n = struct.unpack_from("<I", data, p)[0]
        p += 4 + n
        marker = struct.unpack_from("<i", data, p)[0]
        p += 4
        if marker in (-1, -2):
            p += 4
        elif marker > 0:
            p += marker
        n = struct.unpack_from("<I", data, p)[0]
        if 0 < n <= 0x28:
            return data[p + 4:p + 4 + n].rstrip(b"\x00").decode("latin-1")
    except struct.error:
        pass
    return ""


# ====================================================================== フレーム（2.1.2・2.2.10）
@dataclass
class Site:
    """フレームの中の部品の場所（OleSiteConcreteControl）と、"o" の中のその部品の保存の形。"""

    name: str
    kind: str | None
    id: int
    block: Block
    at: int  # "o" の中の位置
    size: int  # ObjectStreamSize
    position: tuple[int, int]  # フレームの左上からの位置（HIMETRIC）
    control: Control | None = None
    error: str = ""


# 2.4.5 FormEmbeddedActiveXControlCached：ClsidCacheIndex → 種類
CACHED = {7: "Form", 12: "Image", 14: "Frame", 15: "MorphData", 16: "SpinButton", 17: "CommandButton", 18: "TabStrip",
          21: "Label", 23: "TextBox", 24: "ListBox", 25: "ComboBox", 26: "CheckBox", 27: "OptionButton",
          28: "ToggleButton", 47: "ScrollBar", 57: "MultiPage"}


@dataclass
class Frame:
    caption: str
    block: Block
    sites: list[Site]
    sites_at: int = 0  # "f" の中の、Sites の始まり


def _skip_font(buf: bytes, p: int) -> int:
    """GuidAndFont（2.4.7）：FontGUID のあとに StdFont（2.4.12）か TextProps。"""
    g = guid_str(buf[p:p + 16])
    p += 16
    if g == "0BE35203-8F91-11CE-9DE3-00AA004BB851":
        face_len = buf[p + 10]
        return p + 11 + face_len
    if g == "AFC20920-DA4E-11CE-B943-00AA006887B4":
        return parse_block(buf, p, LAYOUTS["TextProps"]).end
    raise AxError("フレームの字の書式（GuidAndFont）の種類が分かりません")


def parse_frame(f: bytes, o: bytes) -> Frame:
    """フレーム（親の部品）の "f"（FormControl）と "o"（中の部品の保存の形を順に並べたもの）を読む。"""
    blk = parse_block(f, 0, LAYOUTS["Form"])
    p = blk.end
    if blk.mask >> 15 & 1:  # MouseIcon
        p = _skip_picture(f, p, "MouseIcon")
    if blk.mask >> 20 & 1:  # Font
        p = _skip_font(f, p)
    if blk.mask >> 21 & 1:  # Picture
        p = _skip_picture(f, p, "Picture")
    # FormSiteData（2.2.10.6）
    flags = blk.get("BooleanProperties", "Form")
    if not flags >> 15 & 1:  # FORM_FLAG_DONTSAVECLASSTABLE でなければ、クラスの表の数がある
        n_cls = struct.unpack_from("<H", f, p)[0]
        p += 2
        for _ in range(n_cls):
            _ver, cb = struct.unpack_from("<HH", f, p)
            p += 4 + cb
    n_sites, n_bytes = struct.unpack_from("<II", f, p)
    p += 8
    start = p
    count = 0
    while count < n_sites:
        _depth, toc = f[p], f[p + 1]
        p += 2
        if toc & 0x80:
            p += 1
            count += toc & 0x7F
        else:
            count += 1
    p = start + _pad4(p - start)
    sites: list[Site] = []
    at = 0
    sites_at = p
    for _ in range(n_sites):
        sb = parse_block(f, p, LAYOUTS["Site"])
        p = sb.end
        idx = sb.get("ClsidCacheIndex", "Site")
        kind = CACHED.get(idx)
        if kind == "MorphData":
            kind = "TextBox"
        size = sb.get("ObjectStreamSize", "Site")
        s = Site(sb.get("Name", "Site"), kind, sb.get("ID", "Site"), sb, at, size,
                 tuple(sb.values.get("Position", (0, 0))))
        if kind in STRUCT and size and at + size <= len(o):
            try:
                s.control = parse(o[at:at + size], kind)
            except AxError as e:
                s.error = str(e)
        elif kind in ("Frame", "MultiPage"):
            s.error = "フレームの中のフレームは、読まずに飛ばしました"
        at += size if kind not in ("Frame", "MultiPage", "Form") else 0
        sites.append(s)
    if p - start != n_bytes:
        raise AxError("フレームの部品の一覧（SiteData）の大きさが合いません")
    return Frame(blk.get("Caption", "Form") or "", blk, sites, sites_at)


# ====================================================================== 値（窓と書き込みの共通）
def state(ctl: Control) -> dict:
    """部品の状態を、窓・証明に出す形で。"""
    k = ctl.kind
    d: dict = {"kind": k}
    if k in ("Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton"):
        d["caption"] = ctl.caption
    if k in ("TextBox", "ComboBox", "ListBox"):
        d["value"] = ctl.value
    if k in ("CheckBox", "OptionButton", "ToggleButton"):
        d["checked"] = ctl.checked
    if k == "OptionButton" or (k == "CheckBox" and ctl.get("GroupName")):
        d["group"] = ctl.get("GroupName") or ""
    if k in ("SpinButton", "ScrollBar"):
        d["position"] = ctl.get("Position")
        d["min"] = ctl.get("Min")
        d["max"] = ctl.get("Max")
    return d


def value_range(ctl: Control) -> tuple[int, int]:
    """スピンボタン・スクロールバーの値の範囲（Min と Max の小さい方〜大きい方。2.5.67）。"""
    a, b = ctl.get("Min"), ctl.get("Max")
    return (min(a, b), max(a, b))


def as_checked(value) -> bool:
    """チェックの部品の値：true / false と、部品の Value に入る値「1」「0」（MS-OFORMS 2.2.5.1 CheckBox の Value）だけ
    （inputcheck.checkbox_value。言葉の意味では決めない）。"""
    from .inputcheck import InputError, checkbox_value

    try:
        return checkbox_value(value, ("1",), ("0",))
    except InputError as e:
        raise AxError(str(e)) from None


def plan(ctl: Control, value, items: list[str] | None = None, force: bool = False) -> tuple[dict, list[str]]:
    """部品に value を書くときに変える性質（性質 → 新しい値）と、知らせること。書けない値なら AxError。
    items：コンボボックス・リストボックスの選択肢（Excel の ListFillRange などで分かるとき。分からなければ None）。"""
    k = ctl.kind
    notes: list[str] = []
    if k not in VALUE_KINDS:
        raise AxError(f"{KIND_JA.get(k, k)}は、値を書く部品ではありません")
    if not ctl.bit("Enabled") and not force:
        raise AxError(f"{KIND_JA[k]}は、使えない（Enabled が偽）ようにされています")
    if ctl.bit("Locked") and k in ("TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton") and not force:
        raise AxError(f"{KIND_JA[k]}は、書き換えられない（Locked が真）ようにされています")
    if k == "TextBox":
        text = "" if value is None else (("TRUE" if value else "FALSE") if isinstance(value, bool) else
                                          (str(int(value)) if isinstance(value, float) and value.is_integer() else str(value)))
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        if "\n" in text and not ctl.bit("MultiLine"):
            raise AxError("改行のある文字は、複数行（MultiLine）のテキストボックスにだけ書けます")
        ml = ctl.get("MaxLength")
        if ml and len(text) > ml:
            raise AxError(f"「{text}」は {len(text)} 字です（このテキストボックスは {ml} 字まで）")
        if ctl.bit("MultiLine"):
            text = text.replace("\n", "\r\n")  # Forms 2.0 は、複数行の改行を CR LF で持つ
        return {"Value": text}, notes
    if k in ("CheckBox", "OptionButton", "ToggleButton"):
        return {"Value": "1" if as_checked(value) else "0"}, notes
    if k in ("ComboBox", "ListBox"):
        if k == "ListBox" and ctl.get("MultiSelect"):
            raise AxError("複数選択のリストボックスは、選んだ項目がファイルに保存されないので書けません")
        text = "" if value is None else (str(int(value)) if isinstance(value, float) and value.is_integer() else str(value))
        must = k == "ListBox" or ctl.style == "droplist" or ctl.bit("MatchRequired")
        if items is not None:
            if text not in items and (must or text):
                if must:
                    shown = "、".join(f"「{x}」" for x in items[:15])
                    raise AxError(f"「{text}」は選択肢にありません。選べるのは {shown}")
                notes.append(f"「{text}」は選択肢にない文字です（この{KIND_JA[k]}は、選択肢にない文字も受け付けます）")
        elif must and text:
            if not force:
                raise AxError(f"この{KIND_JA[k]}は選択肢からしか選べませんが、選択肢がファイルの中にないので、確かめられません")
            notes.append(f"選択肢がファイルの中にないので、「{text}」が選択肢にあるかを確かめていません")
        return {"Value": text}, notes
    # スピンボタン・スクロールバー
    try:
        v = float(value)
    except (TypeError, ValueError):
        raise AxError(f"{KIND_JA[k]}には数を書いてください（「{value}」）") from None
    lo, hi = value_range(ctl)
    if v != int(v) or not lo <= v <= hi:
        raise AxError(f"{KIND_JA[k]}の値は {lo}〜{hi} の整数です（{value!r} は範囲外）")
    return {"Position": int(v)}, notes
