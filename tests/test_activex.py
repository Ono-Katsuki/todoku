"""ActiveX の部品（Microsoft Forms 2.0）の読み書きと証明（MS-OFORMS・MS-OI29500・MS-XLS・MS-DOC）。

部品の保存の形は、このファイルの中で仕様書のとおりに組み立てる（DataBlock の性質は、部品の頭から数えて大きさで
割り切れる位置に置き、DataBlock は 4 の倍数、ExtraDataBlock の文字は 4 の倍数まで詰める。PropMask のビットは
仕様書の各節の図の番号）。とどくの表（activex.LAYOUTS）は使わずに組み立て、とどくが同じに読むかを確かめる。
"""

from __future__ import annotations

import io
import re
import struct
import zipfile
from pathlib import Path

import pytest

from todoku import activex as ax
from todoku import axdraw

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
SPEC = CORPUS / "specs" / "MS-OFORMS.txt"

GUID = {k: v for v, k in ax.CLSIDS.items()}


# ====================================================================== 仕様書どおりに組み立てる
def sstr(s: str) -> tuple[int, bytes]:
    """2.4.14：どの字も上のバイトが 0 なら圧縮（1字1バイト）、そうでなければ UTF-16。"""
    raw = s.encode("utf-16-le")
    if all(raw[i] == 0 for i in range(1, len(raw), 2)):
        return len(raw) // 2 | 0x80000000, raw[0::2]
    return len(raw), raw


def pack_control(minor: int, major: int, mask: int, mask_bytes: int, data: list[tuple[int, bytes]],
                 extra: list[bytes], tail: bytes = b"", pad: int = 0x00) -> bytes:
    """data：DataBlock の性質の (大きさ, バイト列) を仕様書の順に。extra：ExtraDataBlock の性質のバイト列（文字は詰め物を含めて）。
    pad：DataBlock の途中の詰め物の値（仕様では不定。Office は前の値の残りを書くことがある）。"""
    body = bytearray(mask.to_bytes(mask_bytes, "little"))
    for size, raw in data:
        while (4 + len(body)) % size:
            body.append(pad)
        body += raw
    while (len(body) - mask_bytes) % 4:
        body.append(0)
    for raw in extra:
        body += raw
    return struct.pack("<BBH", minor, major, len(body)) + bytes(body) + tail


def padded(raw: bytes, fill: int = 0) -> bytes:
    return raw + bytes([fill]) * ((-len(raw)) % 4)


def textprops(font: str = "ＭＳ ゴシック", height: int = 220, charset: int = 128) -> bytes:
    """2.3.1 TextProps：FontName（ビット0）・FontHeight（ビット2）・FontCharSet（ビット4）。"""
    cnt, raw = sstr(font)
    return pack_control(0, 2, 0x01 | 0x04 | 0x10, 4,
                        [(4, struct.pack("<I", cnt)), (4, struct.pack("<I", height)), (1, bytes([charset]))],
                        [padded(raw)])


def morph(display: int, value: str | None = None, caption: str | None = None, group: str | None = None,
          size=(3000, 600), various: int | None = None, max_length: int | None = None, pad: int = 0) -> bytes:
    """2.2.5 MorphDataControl：DisplayStyle（ビット6）・Size（ビット8）・Value（22）・Caption（23）・GroupName（32）・
    Reserved（31、必ず1）・VariousPropertyBits（0）・MaxLength（3）。"""
    mask = (1 << 8) | (1 << 31)
    data: list[tuple[int, bytes]] = []
    extra: list[bytes] = [struct.pack("<ii", *size)]
    if various is not None:
        mask |= 1
        data.append((4, struct.pack("<I", various)))
    if max_length is not None:
        mask |= 1 << 3
        data.append((4, struct.pack("<I", max_length)))
    if display != 1:
        mask |= 1 << 6
        data.append((1, bytes([display])))
    for bit, s in ((22, value), (23, caption), (32, group)):
        if s:
            mask |= 1 << bit
            cnt, raw = sstr(s)
            data.append((4, struct.pack("<I", cnt)))
            extra.append(padded(raw, 0xAB))  # 詰め物は不定（ここでは 0xAB にして、書き換えで残るかを確かめる）
    return pack_control(0, 2, mask, 8, data, extra, textprops() + b"", pad=pad)


def spin(kind: str, position: int, lo: int | None = None, hi: int | None = None, size=(500, 1000)) -> bytes:
    """2.2.8 SpinButton / 2.2.7 ScrollBar：Size（ビット3）・Min（5）・Max（6）・Position（7）。"""
    mask = 1 << 3
    data = []
    for bit, v in ((5, lo), (6, hi), (7, position or None)):
        if v is not None:
            mask |= 1 << bit
            data.append((4, struct.pack("<i", v)))
    return pack_control(0, 2, mask, 4, data, [struct.pack("<ii", *size)])


def caption_control(kind: str, caption: str, size=(3000, 600)) -> bytes:
    """2.2.1 CommandButton / 2.2.4 Label：Caption（ビット3）・Size（ビット5）。後ろに TextProps。"""
    cnt, raw = sstr(caption)
    return pack_control(0, 2, (1 << 3) | (1 << 5), 4, [(4, struct.pack("<I", cnt))],
                        [padded(raw), struct.pack("<ii", *size)], textprops())


# ====================================================================== 仕様書と照らす
@pytest.fixture(scope="module")
def spec() -> str:
    if not SPEC.exists():
        pytest.skip("MS-OFORMS の仕様書がありません（uv run python scripts/fetch_specs.py）")
    return SPEC.read_text(encoding="utf-8")


def _section(spec: str, title: str) -> str:
    m = re.search(r"\n2\.[0-9.]+ +" + re.escape(title) + r"\n", spec)
    assert m, title
    nxt = re.search(r"\n2\.[0-9.]+ +[A-Z][\w ()]*\n", spec[m.end():])
    return spec[m.end():m.end() + (nxt.start() if nxt else 6000)]


@pytest.mark.parametrize("name,title", [("CommandButton", "CommandButton"), ("Image", "Image"), ("Label", "Label"),
                                        ("MorphData", "MorphData"), ("ScrollBar", "ScrollBar"),
                                        ("SpinButton", "SpinButton"), ("TextProps", "TextProps")])
def test_layout_tables_match_spec(spec, name, title):
    """PropMask のビットの順と、DataBlock・ExtraDataBlock の性質の並びと大きさが、原本の各節と同じ。"""
    pm = re.findall(r"\n(?:[A-Za-z] - )?(\w+) \((\d+) bits?\):", _section(spec, title + "PropMask"))
    bits, pos = {}, 0
    for nm, n in pm:
        bits[nm] = pos
        pos += int(n)
    assert pos == 8 * ax.LAYOUTS[name].mask_bytes
    lay = ax.LAYOUTS[name]
    data = [(n, s) for n, s in re.findall(r"\n(\w+) \((\d+ bytes?|variable)\):", _section(spec, title + "DataBlock"))
            if not n.startswith("Padding")]
    assert [n for n, _b, _s in lay.data] == [n for n, _s in data]
    for (n, b, s), (_n2, s2) in zip(lay.data, data):
        assert int(s2.split()[0]) == (4 if s == ax.S else s), n
        assert bits["f" + n] == b, n
    extra = re.findall(r"\n(\w+) \((\d+ bytes?|variable)\):", _section(spec, title + "ExtraDataBlock"))
    assert [n for n, _b, _s in lay.extra] == [n for n, _s in extra]
    for n, b, _s in lay.extra:
        assert bits["f" + n] == b
    for n, b in lay.flags:
        assert bits["f" + n] == b


def test_various_bits_and_defaults_match_spec(spec):
    seg = _section(spec, "VariousPropertiesBitfield")
    pos, at = 0, {}
    for nm, n in re.findall(r"\n(?:[A-Za-z] - )?(\w+) \((\d+) bits?\):", seg):
        at[nm] = pos
        pos += int(n)
    assert pos == 32
    assert {k: at[k] for k in ax.VPB} == ax.VPB
    assert "0x2C80081B" in _section(spec, "VariousPropertyBits") and ax.DEFAULT_BY_KIND["MorphData"]["VariousPropertyBits"] == 0x2C80081B
    assert "0x00000064 (100)" in _section(spec, "Max") and ax.DEFAULT_BY_KIND["SpinButton"]["Max"] == 100
    ds = _section(spec, "fmDisplayStyle")
    for name, v in (("Text", 1), ("List", 2), ("Combo", 3), ("CheckBox", 4), ("OptionButton", 5), ("Toggle", 6),
                    ("DropList", 7)):
        assert re.search(r"fmDisplayStyle" + name + r"\s+0x0" + str(v), ds)
    cached = _section(spec, "FormEmbeddedActiveXControlCached")
    for idx, kind in ax.CACHED.items():
        assert re.search(rf"\n{idx}\n{kind}(?:\n|$)", cached), (idx, kind)


# ====================================================================== 保存の形
def test_string_compression_examples():
    # 3.1：「ABC」は圧縮して3バイト、「地球」は圧縮できず4バイト
    assert ax.encode_string("ABC") == (3 | 0x80000000, b"ABC")
    assert ax.encode_string("地球") == (4, "地球".encode("utf-16-le"))
    assert ax.encode_string("") == (0, b"")


def test_spec_example_commandbutton():
    # 3.2 の例：Caption「CommandButton1」（圧縮）・MousePointer 0x63・MouseIcon（絵）・大きさ 4445×1270
    picture = b"\x00\x00\x02\x00\x01\x00\x20\x20" + bytes(0x2FE - 8)
    stream = ax.guid_bytes("0BE35204-8F91-11CE-9DE3-00AA004BB851") + struct.pack("<II", 0x746C, 0x2FE) + picture
    cnt, raw = sstr("CommandButton1")
    b = pack_control(0, 2, (1 << 3) | (1 << 5) | (1 << 6) | (1 << 10), 4,
                     [(4, struct.pack("<I", cnt)), (1, b"\x63"), (2, b"\xff\xff")],
                     [padded(raw), struct.pack("<ii", 0x115D, 0x4F6)], stream + textprops("Tahoma", 160, 0))
    assert b[:4] == b"\x00\x02\x24\x00"  # cbCommandButton = 0x24（例と同じ）
    assert b[8:16] == bytes.fromhex("0e00008063" "00" "ffff")  # DataBlock：Caption・MousePointer・詰め物・MouseIcon
    c = ax.parse(b, "CommandButton")
    assert c.caption == "CommandButton1" and c.size == (4445, 1270) and c.get("MousePointer") == 0x63
    assert c.font["name"] == "Tahoma" and c.rest == b""


def test_parse_morph_kinds_and_values():
    tb = ax.parse(morph(1, "山田　太郎"), "TextBox")
    assert (tb.kind, tb.value) == ("TextBox", "山田　太郎")
    cb = ax.parse(morph(4, "1", "同意する"), "CheckBox")
    assert (cb.kind, cb.checked, cb.caption) == ("CheckBox", True, "同意する")
    ob = ax.parse(morph(5, "0", "男", "性別"), "OptionButton")
    assert (ob.kind, ob.checked, ob.get("GroupName")) == ("OptionButton", False, "性別")
    assert ax.parse(morph(5, None, "女", "性別"), "OptionButton").checked is None  # 値なし＝どちらでもない
    assert ax.parse(morph(7, "東京"), "ComboBox").style == "droplist"
    assert ax.parse(morph(3, "東京"), "ComboBox").style == "combo"
    assert ax.parse(morph(6, "1", "表示"), "ToggleButton").checked is True
    # DisplayStyle が正：classid が TextBox でも、DisplayStyle 4 ならチェックボックス
    assert ax.parse(morph(4, "0", "x"), "TextBox").kind == "CheckBox"


def test_rewrite_same_length_changes_only_value_bytes():
    old = morph(1, "AB", max_length=10)
    new = ax.set_props(old, "TextBox", {"Value": "CD"})
    diff = [i for i in range(len(old)) if old[i] != new[i]]
    c = ax.parse(new, "TextBox")
    assert c.value == "CD" and len(new) == len(old)
    p, n, _ = c.block.extra_at["Value"]
    assert all(p <= i < p + n for i in diff)  # 変わったのは文字のバイトだけ（詰め物 0xAB も元のまま）
    assert ax.verify(old, new, "TextBox", {"Value": "CD"}) == []


def test_rewrite_japanese_grows_and_keeps_tail():
    old = morph(1, "A", max_length=40)
    new = ax.set_props(old, "TextBox", {"Value": "申請者　山田"})
    c = ax.parse(new, "TextBox")
    assert c.value == "申請者　山田" and c.block.counts["Value"] == 12  # UTF-16・圧縮なし
    assert c.get("MaxLength") == 40 and c.font["name"] == "ＭＳ ゴシック"
    assert new[-len(textprops()):] == old[-len(textprops()):]  # TextProps はそのまま
    assert struct.unpack_from("<H", new, 2)[0] == struct.unpack_from("<H", old, 2)[0] + 8  # cb は文字の伸びた分だけ
    assert ax.verify(old, new, "TextBox", {"Value": "申請者　山田"}) == []


def test_rewrite_add_and_remove_value_relayouts_datablock():
    old = morph(4, None, "同意する", "組1", various=0x2C80081B | 0x8000)
    on = ax.set_props(old, "CheckBox", {"Value": "1"})
    c = ax.parse(on, "CheckBox")
    assert c.checked is True and c.caption == "同意する" and c.get("GroupName") == "組1"
    assert c.block.stored("Value") and ax.verify(old, on, "CheckBox", {"Value": "1"}) == []
    # 前に置かれた性質（VariousPropertyBits・DisplayStyle）の位置とバイトは元のまま
    assert on[12:17] == old[12:17]
    off = ax.set_props(on, "CheckBox", {"Value": ""})  # 既定値（空の文字）は保存しない
    assert not ax.parse(off, "CheckBox").block.stored("Value")
    assert off == old  # 足して消すと、元のバイト列に戻る（見出し・組の文字の詰め物も元のまま移る）


def test_position_default_is_not_stored():
    old = spin("SpinButton", 5, 0, 10)
    zero = ax.set_props(old, "SpinButton", {"Position": 0})
    assert not ax.parse(zero, "SpinButton").block.stored("Position")
    assert len(zero) == len(old) - 4
    back = ax.set_props(zero, "SpinButton", {"Position": 7})
    assert ax.parse(back, "SpinButton").get("Position") == 7


def test_verify_catches_tampering():
    old = morph(1, "AB")
    new = bytearray(ax.set_props(old, "TextBox", {"Value": "XY"}))
    new[-3] ^= 0xFF  # TextProps のバイトを壊す
    assert any("StreamData" in p for p in ax.verify(old, bytes(new), "TextBox", {"Value": "XY"}))
    other = ax.set_props(old, "TextBox", {"MaxLength": 3})
    assert any("MaxLength" in p for p in ax.verify(old, other, "TextBox", {}))


def test_plan_checks_rules():
    tb = ax.parse(morph(1, "", max_length=3), "TextBox")
    with pytest.raises(ax.AxError):
        ax.plan(tb, "1234")
    with pytest.raises(ax.AxError):
        ax.plan(tb, "a\nb")  # 1行のテキストボックス
    dl = ax.parse(morph(7, ""), "ComboBox")
    with pytest.raises(ax.AxError):
        ax.plan(dl, "大阪", items=["東京", "名古屋"])
    assert ax.plan(dl, "東京", items=["東京", "名古屋"])[0] == {"Value": "東京"}
    with pytest.raises(ax.AxError):
        ax.plan(dl, "東京")  # 選択肢がファイルにない・選ぶだけのコンボボックス
    sp = ax.parse(spin("SpinButton", 1, 0, 10), "SpinButton")
    with pytest.raises(ax.AxError):
        ax.plan(sp, 11)
    assert ax.plan(sp, "3")[0] == {"Position": 3}


# ====================================================================== 控えの絵
def test_emf_and_wmf_are_well_formed_and_carry_the_text():
    c = ax.parse(morph(1, "山田太郎"), "TextBox")
    emf = axdraw.to_emf(axdraw.render(c, 120, 24, 96.0))
    recs = axdraw.emf_records(emf)
    assert recs[0][0] == 1 and recs[-1][0] == 0x0E
    assert axdraw.emf_texts(emf) == ["山田太郎"]
    wmf = axdraw.to_wmf(axdraw.render(c, 120, 24, 96.0), (0, 0, 1800, 360), 1440)
    assert axdraw.wmf_texts(wmf) == ["山田太郎"] and axdraw.check_picture(wmf) == []
    chk = ax.parse(morph(4, "1", "同意する"), "CheckBox")
    assert axdraw.emf_texts(axdraw.to_emf(axdraw.render(chk, 100, 20, 96.0))) == ["同意する"]


# ====================================================================== 複合ファイルの小さなストリーム
def _cfb(streams: dict[str, bytes], clsid: str | None = None) -> bytes:
    """版3（512 バイトのセクタ）の、ルートの直下にストリームを並べた複合ファイル（MS-CFB 2.2〜2.6）。"""
    ss = 512
    names = list(streams)
    mini = bytearray()
    starts = {}
    for n in names:
        d = streams[n]
        starts[n] = len(mini) // 64
        mini += d + bytes((-len(d)) % 64)
    n_mini = len(mini) // 64
    minifat = []
    for n in names:
        k = -(-len(streams[n]) // 64)
        s = starts[n]
        minifat += [s + i + 1 for i in range(k - 1)] + [0xFFFFFFFE]
    minifat += [0xFFFFFFFF] * ((-len(minifat)) % 128)
    sectors = []
    # 0: FAT、1: ディレクトリ、2: ミニ FAT、3〜: ミニストリーム
    dir_entries = 1 + len(names)
    dir_sec = bytearray(ss)

    def entry(i, name, kind, child, left, right, start, size):
        raw = name.encode("utf-16-le") + b"\x00\x00"
        e = raw + bytes(64 - len(raw)) + struct.pack("<HBB", len(raw), kind, 1) + struct.pack("<III", left, right, child)
        e += bytes(16 + 4 + 16) + struct.pack("<IQ", start, size)
        dir_sec[i * 128:(i + 1) * 128] = e

    ms_secs = -(-len(mini) // ss)
    entry(0, "Root Entry", 5, 1 if names else 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 3 if mini else 0xFFFFFFFE, len(mini))
    if clsid:
        dir_sec[80:96] = ax.guid_bytes(clsid)  # ルートの入れ物の CLSID（MS-CFB 2.6.1）
    for i, n in enumerate(names):
        entry(i + 1, n, 2, 0xFFFFFFFF, 0xFFFFFFFF, i + 2 if i + 1 < len(names) else 0xFFFFFFFF, starts[n], len(streams[n]))
    fat = [0xFFFFFFFD, 0xFFFFFFFE, 0xFFFFFFFE] + [3 + i + 1 for i in range(ms_secs - 1)] + [0xFFFFFFFE]
    fat += [0xFFFFFFFF] * (128 - len(fat))
    head = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(16) + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
    head += struct.pack("<IIIIIIIII", 0, 1, 1, 0, 4096, 2, 1, 0xFFFFFFFE, 0)
    head += struct.pack("<109I", 0, *([0xFFFFFFFF] * 108))
    head += bytes(512 - len(head))
    body = struct.pack("<128I", *fat) + bytes(dir_sec) + struct.pack("<128I", *minifat[:128]) + bytes(mini) + bytes((-len(mini)) % ss)
    del sectors, dir_entries, n_mini
    return head + body


def test_cfb_resize_small_streams():
    from todoku.cfb import CFB
    from todoku.cfbwrite import check, replace_stream

    blob = _cfb({"contents": b"A" * 100, "other": b"B" * 70})
    big = replace_stream(blob, "contents", b"C" * 300)  # ミニセクタが足りない：ミニストリームの後ろに足す
    c = CFB(big)
    assert c.read("contents") == b"C" * 300 and c.read("other") == b"B" * 70 and check(big) == []
    small = replace_stream(big, "contents", b"D" * 10)
    c = CFB(small)
    assert c.read("contents") == b"D" * 10 and c.read("other") == b"B" * 70 and check(small) == []
    reg = replace_stream(small, "contents", b"E" * 5000)  # ふつうのセクタへ
    assert CFB(reg).read("contents") == b"E" * 5000 and CFB(reg).read("other") == b"B" * 70 and check(reg) == []
    back = replace_stream(reg, "contents", b"F" * 77)  # ふつうのセクタからミニストリームへ
    assert CFB(back).read("contents") == b"F" * 77 and CFB(back).read("other") == b"B" * 70 and check(back) == []


# ====================================================================== Excel（.xlsx）
def emf_stub(w: int = 120, h: int = 24) -> bytes:
    """元の控えの絵：ヘッダー（MS-EMF 2.2.9、拡張つき 108 バイト）と EOF だけの EMF。装置 1920×1080 px・508×285 mm。"""
    frame = (0, 0, round(w * 508 * 100 / 1920), round(h * 285 * 100 / 1080))
    eof = struct.pack("<IIIII", 0x0E, 20, 0, 16, 20)
    head = struct.pack("<II", 1, 108) + struct.pack("<iiii", 0, 0, w - 1, h - 1) + struct.pack("<iiii", *frame)
    head += struct.pack("<IIIIHHIII", 0x464D4520, 0x10000, 108 + len(eof), 2, 1, 0, 0, 0, 0)
    head += struct.pack("<ii", 1920, 1080) + struct.pack("<ii", 508, 285) + struct.pack("<III", 0, 0, 0)
    head += struct.pack("<ii", 508000, 285000)
    return head + eof


def ax_xlsx(controls: list[dict], cells: dict[str, str] | None = None, bag: bool = False, share_picture: bool = False) -> bytes:
    """ActiveX の部品を置いたブック（MS-OI29500 3.4.1.1・3.6.1.2、ISO/IEC 29500 の control・controlPr・VML）。
    controls：{"name", "clsid", "bin", "row", "col", "link", "fill"}。bag=True は性質の袋（persistPropertyBag）で保存する
    （そのときは "bag"：性質の名前 → 文字）。"""
    cells = cells or {}
    files: dict[str, bytes] = {}
    ns = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>',
          '<Default Extension="emf" ContentType="image/x-emf"/>',
          '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>']
    rels = ['<Relationship Id="rIdV" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" '
            'Target="../drawings/vmlDrawing1.vml"/>']
    vml_rels, shapes, ctls = [], [], []
    for i, c in enumerate(controls, 1):
        sid = 1024 + i
        ct.append(f'<Override PartName="/xl/activeX/activeX{i}.xml" ContentType="application/vnd.ms-office.activeX+xml"/>')
        if bag:
            props = "".join(f'<ax:ocxPr ax:name="{k}" ax:value="{v}"/>' for k, v in c["bag"].items())
            files[f"xl/activeX/activeX{i}.xml"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ax:ocx ax:classid="{{{c["clsid"]}}}" '
                f'ax:persistence="persistPropertyBag" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" '
                f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{props}</ax:ocx>').encode()
        else:
            files[f"xl/activeX/activeX{i}.xml"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ax:ocx ax:classid="{{{c["clsid"]}}}" '
                'ax:persistence="persistStreamInit" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>').encode()
            files[f"xl/activeX/_rels/activeX{i}.xml.rels"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX{i}.bin"/>'
                '</Relationships>').encode()
            files[f"xl/activeX/activeX{i}.bin"] = ax.guid_bytes(c["clsid"]) + c["bin"]
        rels.append(f'<Relationship Id="rIdC{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" '
                    f'Target="../activeX/activeX{i}.xml"/>')
        img = 1 if share_picture else i  # Excel は同じ絵を1つの部品・1つの関係にまとめることがある
        if img == i:
            files[f"xl/media/image{i}.emf"] = emf_stub()
            rels.append(f'<Relationship Id="rIdI{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                        f'Target="../media/image{i}.emf"/>')
            vml_rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                            f'Target="../media/image{i}.emf"/>')
        r, col = c["row"] - 1, c["col"] - 1
        shapes.append(f'<v:shape id="{c["name"]}" o:spid="_x0000_s{sid}" type="#_x0000_t201" style="position:absolute;'
                      f'width:90pt;height:18pt"><v:imagedata o:relid="rId{img}" o:title=""/><x:ClientData ObjectType="Pict">'
                      f'<x:Anchor>{col}, 2, {r}, 2, {col + 2}, 2, {r}, 20</x:Anchor><x:CF>Pict</x:CF></x:ClientData></v:shape>')
        attrs = f' linkedCell="{c["link"]}"' if c.get("link") else ""
        attrs += f' listFillRange="{c["fill"]}"' if c.get("fill") else ""
        ctls.append(f'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                    f'<mc:Choice Requires="x14"><control shapeId="{sid}" r:id="rIdC{i}" name="{c["name"]}">'
                    f'<controlPr defaultSize="0" autoLine="0"{attrs} r:id="rIdI{img}"><anchor moveWithCells="1">'
                    f'<from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>0</xdr:rowOff></from>'
                    f'<to><xdr:col>{col + 2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>200000</xdr:rowOff></to>'
                    f'</anchor></controlPr></control></mc:Choice><mc:Fallback><control shapeId="{sid}" r:id="rIdC{i}" '
                    f'name="{c["name"]}"/></mc:Fallback></mc:AlternateContent>')
    rows: dict[int, list[str]] = {}
    for a1, v in cells.items():
        m = re.fullmatch(r"([A-Z]+)(\d+)", a1)
        rows.setdefault(int(m[2]), []).append(f'<c r="{a1}" t="inlineStr"><is><t>{v}</t></is></c>')
    sheet_data = "".join(f'<row r="{r}">{"".join(cs)}</row>' for r, cs in sorted(rows.items()))
    files["xl/worksheets/sheet1.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet {ns} '
        'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
        f'<sheetData>{sheet_data}</sheetData><legacyDrawing r:id="rIdV"/><controls>{"".join(ctls)}</controls></worksheet>').encode()
    files["xl/worksheets/_rels/sheet1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels) + "</Relationships>").encode()
    files["xl/drawings/vmlDrawing1.vml"] = (
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:x="urn:schemas-microsoft-com:office:excel">' + "".join(shapes) + "</xml>").encode()
    files["xl/drawings/_rels/vmlDrawing1.vml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(vml_rels) + "</Relationships>").encode()
    files["xl/workbook.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook {ns}><sheets>'
                                '<sheet name="申請" sheetId="1" r:id="rId1"/></sheets></workbook>').encode()
    files["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/></Relationships>').encode()
    files["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>').encode()
    files["[Content_Types].xml"] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.'
                                    'openxmlformats.org/package/2006/content-types">' + "".join(ct) + "</Types>").encode()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ["[Content_Types].xml"] + [n for n in files if n != "[Content_Types].xml"]:
            z.writestr(n, files[n])
    return buf.getvalue()


def form_controls() -> list[dict]:
    return [
        {"name": "TextBox1", "clsid": GUID["TextBox"], "bin": morph(1, None, max_length=20), "row": 1, "col": 2, "link": "H1"},
        {"name": "CheckBox1", "clsid": GUID["CheckBox"], "bin": morph(4, "0", "同意する"), "row": 3, "col": 2, "link": "H3"},
        {"name": "OptionButton1", "clsid": GUID["OptionButton"], "bin": morph(5, "1", "男", "性別"), "row": 5, "col": 2,
         "link": "H5"},
        {"name": "OptionButton2", "clsid": GUID["OptionButton"], "bin": morph(5, "0", "女", "性別"), "row": 5, "col": 4,
         "link": "H6"},
        {"name": "ComboBox1", "clsid": GUID["ComboBox"], "bin": morph(7, None), "row": 7, "col": 2, "link": "H7",
         "fill": "E1:E3"},
        {"name": "SpinButton1", "clsid": GUID["SpinButton"], "bin": spin("SpinButton", 1, 0, 10), "row": 9, "col": 2,
         "link": "H9"},
        {"name": "Label1", "clsid": GUID["Label"], "bin": caption_control("Label", "申請者の住所"), "row": 11, "col": 2},
        {"name": "CommandButton1", "clsid": GUID["CommandButton"], "bin": caption_control("CommandButton", "入力チェック"),
         "row": 13, "col": 2},
    ]


@pytest.fixture()
def xlsx_blob() -> bytes:
    return ax_xlsx(form_controls(), {"A1": "氏名", "A3": "確認", "A5": "性別", "A7": "都道府県", "A9": "人数",
                                     "E1": "東京", "E2": "大阪", "E3": "名古屋"})


def test_xlsx_read_controls(xlsx_blob):
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(xlsx_blob, "a.xlsx")
    by = {c.name: c for c in wb.sheet("申請").controls}
    assert by["TextBox1"].kind == "ActiveX TextBox" and by["TextBox1"].link == "H1"
    assert by["CheckBox1"].text == "同意する" and by["CheckBox1"].checked is None
    assert by["OptionButton1"].checked == "Checked" and by["OptionButton1"].ax.control.get("GroupName") == "性別"
    assert by["ComboBox1"].fmla_range == "E1:E3" and by["ComboBox1"].ax.control.style == "droplist"
    assert (by["SpinButton1"].val, by["SpinButton1"].vmin, by["SpinButton1"].vmax) == (1, 0, 10)
    assert by["Label1"].text == "申請者の住所" and by["CommandButton1"].text == "入力チェック"
    assert by["TextBox1"].ax.images == ["xl/media/image1.emf"]  # controlPr と VML の両方が同じ絵を指す


def test_xlsx_window_shows_controls(xlsx_blob):
    from todoku.xlsx import window
    from todoku.xlsx.workbook import Workbook

    doc = window.build(Workbook(xlsx_blob, "a.xlsx"))
    sh = doc["sheets"][0]
    f = {x.get("control"): x for x in sh["fields"]}
    assert f["TextBox1"]["id"] == "申請!H1" and f["TextBox1"]["kind"] == "text" and f["TextBox1"]["label"] == "氏名"
    assert f["CheckBox1"]["kind"] == "checkbox" and f["CheckBox1"]["value"] is False
    assert f["OptionButton2"]["kind"] == "option" and f["OptionButton2"]["group"] == "性別"
    assert f["ComboBox1"]["options"] == ["東京", "大阪", "名古屋"] and f["ComboBox1"]["kind"] == "select"
    assert f["SpinButton1"]["min"] == 0 and f["SpinButton1"]["max"] == 10
    text = "\n".join(x["text"] for x in sh["reading"])
    assert "☐同意する［H3］" in text and "◉男［H5］" in text and "○女［H6］" in text and "申請者の住所" in text
    md = window.to_markdown(doc)
    assert "ActiveX のテキストボックス「TextBox1」" in md and "入力チェック" in md


def test_xlsx_fill_and_prove(xlsx_blob):
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(xlsx_blob, "a.xlsx")
    res = fill(wb, {"申請!TextBox1": "山田　太郎", "申請!CheckBox1": True, "申請!OptionButton2": True,
                    "申請!ComboBox1": "大阪", "申請!SpinButton1": 7})
    allowed_c = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}
    pr = prove(xlsx_blob, res.blob, res.allowed_inputs, allowed_c)
    assert pr.ok, pr.unexpected + pr.lint_errors
    after = Workbook(res.blob, "a.xlsx")
    by = {c.name: c for c in after.sheet("申請").controls}
    cells = after.sheet("申請").cells
    assert by["TextBox1"].ax.control.value == "山田　太郎" and cells[(1, 8)].value == "山田　太郎"
    assert by["CheckBox1"].checked == "Checked" and cells[(3, 8)].value is True
    assert by["OptionButton2"].checked == "Checked" and by["OptionButton1"].checked is None or \
        by["OptionButton1"].ax.control.checked is False
    assert cells[(5, 8)].value is False and cells[(6, 8)].value is True
    assert by["ComboBox1"].ax.control.value == "大阪" and cells[(7, 8)].value == "大阪"
    assert by["SpinButton1"].val == 7 and cells[(9, 8)].value == 7
    # 控えの絵は書いた状態で描き直され、書かなかった部品の絵は元のまま
    z0, z1 = zipfile.ZipFile(io.BytesIO(xlsx_blob)), zipfile.ZipFile(io.BytesIO(res.blob))
    assert axdraw.emf_texts(z1.read("xl/media/image1.emf")) == ["山田　太郎"]
    assert z0.read("xl/media/image7.emf") == z1.read("xl/media/image7.emf")
    assert z0.read("xl/activeX/activeX8.bin") == z1.read("xl/activeX/activeX8.bin")


def test_xlsx_write_through_linked_cell(xlsx_blob):
    from todoku.xlsx.edit import fill
    from todoku.xlsx.workbook import Workbook

    res = fill(Workbook(xlsx_blob, "a.xlsx"), {"申請!H1": "佐藤", "申請!H3": "TRUE"})
    after = Workbook(res.blob, "a.xlsx")
    by = {c.name: c for c in after.sheet("申請").controls}
    assert by["TextBox1"].ax.control.value == "佐藤" and by["CheckBox1"].checked == "Checked"


def test_xlsx_refuses_bad_values(xlsx_blob):
    from todoku.xlsx.edit import EditError, fill
    from todoku.xlsx.workbook import Workbook

    for values in ({"申請!ComboBox1": "福岡"}, {"申請!SpinButton1": 11}, {"申請!CommandButton1": "x"},
                   {"申請!TextBox1": "あ" * 21}):
        with pytest.raises(EditError):
            fill(Workbook(xlsx_blob, "a.xlsx"), values)


def test_xlsx_proof_detects_unrequested_changes(xlsx_blob):
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove
    from todoku.xlsx.workbook import Workbook
    from todoku.zipraw import RawZip

    res = fill(Workbook(xlsx_blob, "a.xlsx"), {"申請!TextBox1": "山田"})
    allowed_c = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}
    z = RawZip(res.blob)
    data = bytearray(z.read("xl/activeX/activeX2.bin"))  # 書いていないチェックボックスを壊す
    data[-1] ^= 0x01
    z.replace("xl/activeX/activeX2.bin", bytes(data))
    pr = prove(xlsx_blob, z.build(), res.allowed_inputs, allowed_c)
    assert not pr.ok and any("CheckBox1" in u for u in pr.unexpected)
    z = RawZip(res.blob)
    z.replace("xl/media/image1.emf", emf_stub())  # 控えの絵を古いままにする
    pr = prove(xlsx_blob, z.build(), res.allowed_inputs, allowed_c)
    assert not pr.ok and any("控えの絵" in u for u in pr.unexpected)


def test_xlsx_property_bag():
    blob = ax_xlsx([{"name": "CheckBox1", "clsid": GUID["CheckBox"], "row": 2, "col": 2, "link": "H2",
                     "bag": {"DisplayStyle": "4", "Caption": "確認しました", "Value": "0", "Size": "3000;600"}}], bag=True)
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(blob, "b.xlsx")
    c = wb.sheet("申請").controls[0]
    assert c.kind == "ActiveX CheckBox" and c.text == "確認しました" and c.checked is None
    res = fill(wb, {"申請!CheckBox1": True})
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"})
    assert pr.ok, pr.unexpected
    assert b'ax:name="Value" ax:value="1"' in zipfile.ZipFile(io.BytesIO(res.blob)).read("xl/activeX/activeX1.xml")
    assert Workbook(res.blob).sheet("申請").controls[0].checked == "Checked"


def test_xlsx_shared_picture_is_split():
    """Excel は同じ控えの絵を1つの部品にまとめる：書いた部品だけの絵を足し、共有の絵（ほかの部品が使う）は元のまま。"""
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove
    from todoku.xlsx.workbook import Workbook

    ctls = [{"name": "OptionButton1", "clsid": GUID["OptionButton"], "bin": morph(5, "0", "はい", "回答"), "row": 1, "col": 2},
            {"name": "OptionButton2", "clsid": GUID["OptionButton"], "bin": morph(5, "0", "いいえ", "回答"), "row": 2, "col": 2}]
    blob = ax_xlsx(ctls, share_picture=True)
    res = fill(Workbook(blob, "s.xlsx"), {"申請!OptionButton1": True})
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"})
    assert pr.ok, pr.unexpected + pr.lint_errors
    z0, z1 = zipfile.ZipFile(io.BytesIO(blob)), zipfile.ZipFile(io.BytesIO(res.blob))
    assert z1.read("xl/media/image1.emf") == z0.read("xl/media/image1.emf")
    by = {c.name: c for c in Workbook(res.blob).sheet("申請").controls}
    assert by["OptionButton1"].ax.images != by["OptionButton2"].ax.images
    assert axdraw.emf_texts(z1.read(by["OptionButton1"].ax.images[0])) == ["はい"]


def test_xlsx_shared_picture_is_split_once_per_control():
    """選んだボタンが外すボタンを、名指しでも外すと、同じ部品に「Value を 0 に」が2度届く。部品は1度だけ書き、
    共有の絵から分けた絵がどこからも指されないまま残らないこと。同じ組のボタンを2つ選ぶのは断る。"""
    import re

    from todoku.xlsx.edit import EditError, fill
    from todoku.xlsx.proof import prove
    from todoku.xlsx.workbook import Workbook

    ctls = [{"name": f"OptionButton{i}", "clsid": GUID["OptionButton"], "bin": morph(5, "1" if i == 1 else "0", t, "回答"),
             "row": i, "col": 2} for i, t in ((1, "はい"), (2, "いいえ"), (3, "わからない"))]
    blob = ax_xlsx(ctls, share_picture=True)
    res = fill(Workbook(blob, "s.xlsx"), {"申請!OptionButton2": True, "申請!OptionButton1": False})
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"})
    assert pr.ok, pr.unexpected + pr.lint_errors
    assert sorted(p.address for p in res.patches if p.kind == "control") == ["OptionButton1", "OptionButton2"]
    z1 = zipfile.ZipFile(io.BytesIO(res.blob))
    targets = {t.rsplit("/", 1)[-1] for n in z1.namelist() if n.endswith(".rels")
               for t in re.findall(r'Target="([^"]*)"', z1.read(n).decode())}
    assert {n.rsplit("/", 1)[-1] for n in z1.namelist() if "/media/" in n} <= targets  # どの絵も、どこかから指されている
    by = {c.name: c for c in Workbook(res.blob).sheet("申請").controls}
    assert [by[f"OptionButton{i}"].ax.control.get("Value") for i in (1, 2, 3)] == ["0", "1", "0"]
    # 2 を選んでから 3 を選ぶ：3 が 2 を外すので、2 には「選ぶ」と「外す」の両方が届く。どちらが望みか決まらないので断る
    with pytest.raises(EditError, match="2度"):
        fill(Workbook(blob, "s.xlsx"), {"申請!OptionButton2": True, "申請!OptionButton3": True})


# ====================================================================== Word（.docx）
def wmf_stub(tw: int = 1800, th: int = 360) -> bytes:
    """元の控えの絵：置き場所の頭（MS-WMF 2.3.2.3、22バイト）と META_HEADER・META_EOF だけの WMF。"""
    ph = struct.pack("<IHhhhhHI", 0x9AC6CDD7, 0, 0, 0, tw, th, 1440, 0)
    chk = 0
    for w in struct.unpack("<10H", ph):
        chk ^= w
    eof = struct.pack("<IH", 3, 0)
    head = struct.pack("<HHHIHIH", 1, 9, 0x0300, (18 + len(eof)) // 2, 0, 3, 0)
    return ph + struct.pack("<H", chk) + head + eof


def compobj(progid: str) -> bytes:
    """MS-OLEDS 2.3.8 CompObjStream：ヘッダー 28 バイト・AnsiUserType・ClipboardFormat（なし）・Reserved1（ProgID）。"""
    user = b"Microsoft Forms 2.0\x00"
    pid = progid.encode() + b"\x00"
    return bytes(28) + struct.pack("<I", len(user)) + user + struct.pack("<I", 0) + struct.pack("<I", len(pid)) + pid


def ax_docx(controls: list[dict], share_picture: bool = False) -> bytes:
    """ActiveX の部品を置いた Word の文書（w:object の中の v:shape と w:control。部品の中身は persistStorage の複合ファイル）。"""
    files: dict[str, bytes] = {}
    rels, paras = [], []
    ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>',
          '<Default Extension="wmf" ContentType="image/x-wmf"/>',
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>']
    for i, c in enumerate(controls, 1):
        clsid = GUID[c["kind"]]
        ct.append(f'<Override PartName="/word/activeX/activeX{i}.xml" ContentType="application/vnd.ms-office.activeX+xml"/>')
        files[f"word/activeX/activeX{i}.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ax:ocx ax:classid="{{{clsid}}}" ax:persistence="persistStorage" '
            'r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>').encode()
        files[f"word/activeX/_rels/activeX{i}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX{i}.bin"/>'
            '</Relationships>').encode()
        files[f"word/activeX/activeX{i}.bin"] = _cfb({"\x01CompObj": compobj(f"Forms.{c['kind']}.1"), "contents": c["bin"]}, clsid)
        img = 1 if share_picture else i
        if not share_picture or i == 1:
            files[f"word/media/image{img}.wmf"] = wmf_stub()
            rels.append(f'<Relationship Id="rIdI{img}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                        f'Target="media/image{img}.wmf"/>')
        rels.append(f'<Relationship Id="rIdC{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" '
                    f'Target="activeX/activeX{i}.xml"/>')
        paras.append(f'<w:p><w:r><w:t>{c.get("before", "")}</w:t></w:r><w:r><w:object w:dxaOrig="1800" w:dyaOrig="360">'
                     f'<v:shape id="_x0000_i{1024 + i}" type="#_x0000_t75" style="width:90pt;height:18pt" o:ole="">'
                     f'<v:imagedata r:id="rIdI{img}" o:title=""/></v:shape>'
                     f'<w:control r:id="rIdC{i}" w:name="{c["name"]}" w:shapeid="_x0000_i{1024 + i}"/></w:object></w:r></w:p>')
    files["word/document.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>'
        + "".join(paras) + "</w:body></w:document>").encode()
    files["word/_rels/document.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels) + "</Relationships>").encode()
    files["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>').encode()
    files["[Content_Types].xml"] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.'
                                    'openxmlformats.org/package/2006/content-types">' + "".join(ct) + "</Types>").encode()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ["[Content_Types].xml"] + [n for n in files if n != "[Content_Types].xml"]:
            z.writestr(n, files[n])
    return buf.getvalue()


def word_controls() -> list[dict]:
    return [{"name": "TextBox1", "kind": "TextBox", "bin": morph(1, "", max_length=30), "before": "氏名："},
            {"name": "OptionButton1", "kind": "OptionButton", "bin": morph(5, "1", "はい", "回答"), "before": "回答："},
            {"name": "OptionButton2", "kind": "OptionButton", "bin": morph(5, "0", "いいえ", "回答")},
            {"name": "CheckBox1", "kind": "CheckBox", "bin": morph(4, "0", "同意する")},
            {"name": "Label1", "kind": "Label", "bin": caption_control("Label", "記入上の注意")}]


@pytest.mark.parametrize("share", [False, True])
def test_docx_read_fill_prove(share):
    from todoku.docx.document import DocxDocument
    from todoku.docx.fill import fill, written_of
    from todoku.docx.proof import prove
    from todoku.docx.window import build, to_markdown

    blob = ax_docx(word_controls(), share_picture=share)
    win = build(DocxDocument(blob, "a.docx"))
    f = {x["id"]: x for x in win["fields"]}
    assert f["TextBox1"]["kind"] == "text" and f["TextBox1"]["label"].startswith("氏名")
    assert f["OptionButton1"]["value"] is True and f["OptionButton2"]["group"] == "回答"
    md = to_markdown(win)
    assert "◉はい［OptionButton1］" in md and "○いいえ［OptionButton2］" in md and "記入上の注意" in md
    res = fill(DocxDocument(blob, "a.docx"), {"TextBox1": "山田　花子", "OptionButton2": True, "CheckBox1": True})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors
    after = {x["id"]: x["value"] for x in build(DocxDocument(res.blob, "a.docx"))["fields"]}
    assert after == {"TextBox1": "山田　花子", "OptionButton1": False, "OptionButton2": True, "CheckBox1": True}
    z = zipfile.ZipFile(io.BytesIO(res.blob))
    pics = [axdraw.wmf_texts(z.read(n)) for n in z.namelist() if n.endswith(".wmf")]
    assert ["山田　花子"] in pics
    if share:  # 共有していた絵は元のまま残り、書いた部品ごとに新しい絵ができる
        assert z.read("word/media/image1.wmf") == wmf_stub() and len(pics) > 1


@pytest.mark.parametrize("share", [False, True])
def test_docx_same_control_is_written_once(share):
    """選んだボタンが外すボタンを、名指しでも外す（同じ値）なら1度だけ書く。同じ組のボタンを2つ選ぶのは断る。"""
    from todoku.docx.document import DocxDocument
    from todoku.docx.fill import FillError, fill, written_of
    from todoku.docx.proof import prove

    blob = ax_docx(word_controls(), share_picture=share)
    res = fill(DocxDocument(blob, "a.docx"), {"OptionButton2": True, "OptionButton1": False})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors
    z = zipfile.ZipFile(io.BytesIO(res.blob))
    rels = z.read("word/_rels/document.xml.rels").decode()
    assert all(n.rsplit("/", 1)[-1] in rels for n in z.namelist() if "/media/" in n)  # 指されない絵を残さない
    blob3 = ax_docx(word_controls() + [{"name": "OptionButton3", "kind": "OptionButton",
                                        "bin": morph(5, "0", "わからない", "回答")}], share_picture=share)
    with pytest.raises(FillError, match="2度"):
        fill(DocxDocument(blob3, "a.docx"), {"OptionButton2": True, "OptionButton3": True})


def test_docx_proof_detects_wrong_value():
    from todoku.docx.document import DocxDocument
    from todoku.docx.fill import fill, written_of
    from todoku.docx.proof import prove

    blob = ax_docx(word_controls())
    res = fill(DocxDocument(blob, "a.docx"), {"TextBox1": "山田"})
    w = written_of(res.patches)
    w["TextBox1"] = "佐藤"  # 書いたと称する値が、部品の中身と違う
    assert not prove(blob, res.blob, w).ok


@pytest.mark.skipif(not (CORPUS / "wide_activex").exists(), reason="ActiveX の部品のある様式（corpus/wide_activex）がありません")
def test_corpus_activex_audits():
    from todoku.docx import audit as daudit
    from todoku.xlsx import audit as xaudit

    files = sorted(p for p in (CORPUS / "wide_activex").rglob("*") if p.suffix.lower() in (".xlsm", ".xlsx", ".docm", ".docx"))
    files.append(CORPUS / "wide" / "www.city.nagoya.jp" / "r8mousikomisyo.xlsm")
    for path in files:
        if not path.exists():
            continue
        au = (daudit if path.suffix.lower().startswith(".doc") else xaudit).run(path.read_bytes(), path.name)
        bad = {c.name: c.items[:3] for c in au.checks if not c.ok and ("ActiveX" in c.name or "読み落とし" in c.name)}
        assert not bad, (path.name, bad)


# ====================================================================== 古い Excel（.xls）
def xls_pict_obj(oid: int, progid: str, pos: int, size: int, link: bytes | None):
    """ActiveX の部品の OBJ（MS-XLS 2.4.181）：FtCmo（ot=8）・FtCf・FtPioGrbit（fCtl・fPrstm）・FtPictFmla
    （ObjFmla：PtgTbl と PictFmlaEmbedInfo、lPosInCtlStm・cbBufInCtlStm・PictFmlaKey）・ftEnd。"""
    from test_xls_controls import Obj, ft, obj_fmla

    cls = progid.encode("latin-1")
    fmla = struct.pack("<H", 5) + bytes(4) + b"\x02" + bytes(4) + bytes([3, len(cls), 0, 0]) + cls
    fmla += b"\0" * (len(fmla) % 2)
    body = struct.pack("<H", len(fmla)) + fmla + struct.pack("<II", pos, size)
    body += struct.pack("<I", 0) + obj_fmla(link) + obj_fmla(None)  # PictFmlaKey：鍵なし・つながったセル・選択肢の範囲なし
    o = Obj(0x08, oid)
    o.add(ft(0x07, struct.pack("<H", 0x0002)))
    o.add(ft(0x08, struct.pack("<H", 0x0010 | 0x0020)))
    o.add(ft(0x09, body))
    return o


def ax_xls() -> bytes:
    from test_xls_controls import ptg_ref, workbook

    s1 = ax.guid_bytes(GUID["TextBox"]) + morph(1, "", max_length=10)
    s2 = ax.guid_bytes(GUID["CheckBox"]) + morph(4, "0", "同意する")
    shapes = [{"obj": xls_pict_obj(1, "Forms.TextBox.1", 0, len(s1), ptg_ref(0, 5)), "anchor": (1, 0, 3, 1), "name": "TextBox1"},
              {"obj": xls_pict_obj(2, "Forms.CheckBox.1", len(s1), len(s2), ptg_ref(1, 5)), "anchor": (1, 2, 3, 3),
               "name": "CheckBox1"}]
    data = workbook([{"name": "申込", "cells": {(0, 0): "氏名", (1, 0): "確認"}, "shapes": shapes}])
    assert len(data) < 4096
    return _cfb({"Workbook": data, "Ctls": s1 + s2})


def test_xls_read_fill_prove():
    from todoku.xls.book import XlsWorkbook
    from todoku.xlsx import window
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove

    blob = ax_xls()
    wb = XlsWorkbook(blob, "a.xls")
    by = {c.name: c for c in wb.sheet("申込").controls}
    assert by["TextBox1"].kind == "ActiveX TextBox" and by["TextBox1"].link == "F1"
    assert by["CheckBox1"].text == "同意する" and by["CheckBox1"].checked is None
    f = {x.get("control"): x for x in window.build(wb)["sheets"][0]["fields"]}
    assert f["TextBox1"]["id"] == "申込!F1" and f["CheckBox1"]["kind"] == "checkbox"
    res = fill(XlsWorkbook(blob, "a.xls"), {"申込!TextBox1": "山田　太郎", "申込!CheckBox1": True})
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"})
    assert pr.ok, pr.unexpected + pr.lint_errors
    after = XlsWorkbook(res.blob, "a.xls")
    by2 = {c.name: c for c in after.sheet("申込").controls}
    assert by2["TextBox1"].ax.control.value == "山田　太郎" and by2["CheckBox1"].checked == "Checked"
    assert after.sheet("申込").cells[(1, 6)].value == "山田　太郎" and after.sheet("申込").cells[(2, 6)].value is True
    # テキストボックスが伸びたので、後ろの部品（チェックボックス）の lPosInCtlStm が直っている
    assert by2["CheckBox1"].ax.xls["pos"] > by["CheckBox1"].ax.xls["pos"]


# ====================================================================== 古い Word（.doc）
def _cfb_tree(streams: dict[str, bytes], clsids: dict[str, str] | None = None) -> bytes:
    """入れ物（"入れ物/ストリーム"）もある複合ファイル（版3）。4096 バイト以上はふつうのセクタ、未満はミニストリーム。"""
    ss, ms = 512, 64
    clsids = clsids or {}
    nodes: list[dict] = [{"name": "Root Entry", "kind": 5, "kids": []}]
    index = {"": 0}
    for path in streams:
        parts = path.split("/")
        for depth in range(1, len(parts)):
            p = "/".join(parts[:depth])
            if p not in index:
                index[p] = len(nodes)
                nodes.append({"name": parts[depth - 1], "kind": 1, "kids": []})
                nodes[index["/".join(parts[:depth - 1])]]["kids"].append(index[p])
        index[path] = len(nodes)
        nodes.append({"name": parts[-1], "kind": 2, "kids": [], "data": streams[path]})
        nodes[index["/".join(parts[:-1])]]["kids"].append(index[path])
    sectors: list[bytes] = []
    fat: list[int] = []

    def chain(blob: bytes) -> int:
        n = -(-len(blob) // ss)
        first = len(sectors)
        for k in range(n):
            sectors.append(blob[k * ss:(k + 1) * ss].ljust(ss, b"\0"))
            fat.append(first + k + 1 if k + 1 < n else 0xFFFFFFFE)
        return first if n else 0xFFFFFFFE

    mini, minifat = bytearray(), []
    for nd in nodes:
        d = nd.get("data")
        if d is None:
            continue
        if len(d) >= 4096:
            nd["start"] = chain(d)
        else:
            k = -(-len(d) // ms)
            first = len(mini) // ms
            mini += d.ljust(k * ms, b"\0")
            minifat += [first + i + 1 for i in range(k - 1)] + [0xFFFFFFFE] if k else []
            nd["start"] = first if k else 0xFFFFFFFE
    minifat += [0xFFFFFFFF] * ((-len(minifat)) % 128)
    mf_start = chain(struct.pack(f"<{len(minifat)}I", *minifat)) if minifat else 0xFFFFFFFE
    root_start = chain(bytes(mini)) if mini else 0xFFFFFFFE
    for nd in nodes:  # 兄弟は右へ1列につなぐ（赤黒木としては不正確でも、たどる順は同じ）
        for a, b in zip(nd["kids"], nd["kids"][1:]):
            nodes[a]["right"] = b
    dirs = bytearray()
    for i, nd in enumerate(nodes):
        raw = (nd["name"] + "\0").encode("utf-16-le")
        start = root_start if i == 0 else nd.get("start", 0)
        size = len(mini) if i == 0 else len(nd.get("data", b""))
        clsid = ax.guid_bytes(clsids[next(p for p, k in index.items() if k == i)]) if \
            next(p for p, k in index.items() if k == i) in clsids else bytes(16)
        dirs += (raw.ljust(64, b"\0") + struct.pack("<HBB", len(raw), nd["kind"], 1)
                 + struct.pack("<III", 0xFFFFFFFF, nd.get("right", 0xFFFFFFFF), nd["kids"][0] if nd["kids"] else 0xFFFFFFFF)
                 + clsid + bytes(4 + 16) + struct.pack("<IQ", start, size))
    dirs += bytes((-len(dirs)) % ss)
    dir_start = chain(bytes(dirs))
    n_fat = 1
    while n_fat * 128 < len(sectors) + n_fat:
        n_fat += 1
    fat_at = list(range(len(sectors), len(sectors) + n_fat))
    fat += [0xFFFFFFFD] * n_fat
    fat += [0xFFFFFFFF] * (n_fat * 128 - len(fat))
    head = (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(16) + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
            + struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, 4096, mf_start, len(minifat) // 128 if minifat else 0,
                          0xFFFFFFFE, 0)
            + struct.pack("<109I", *(fat_at + [0xFFFFFFFF] * (109 - n_fat))))
    return head + b"".join(sectors) + struct.pack(f"<{len(fat)}I", *fat)


def ax_doc(controls: list[dict]) -> bytes:
    """ActiveX の部品を置いた Word 97 の文書（MS-DOC）。部品ごとに、段落「見出し」＋ CONTROL のフィールド
    （0x13 命令 0x14 絵の字 0x01 0x15）。区切りの字の sprmCPicLocation（0x6A03）は ObjectPool の入れ物の番号、
    絵の字には sprmCFSpec（0x0855）。入れ物には \\003ObjInfo（ODT：fOCX）・\\003OCXNAME・\\001CompObj・contents。"""
    text, runs, pool = "", [], {}
    for i, c in enumerate(controls):
        num = 1000 + i
        text += c.get("before", "")
        text += "\x13 CONTROL Forms." + c["kind"] + ".1 \\s \x14"
        runs.append((len(text) - 1, struct.pack("<HI", 0x6A03, num)))
        text += "\x01"
        runs.append((len(text) - 1, struct.pack("<HB", 0x0855, 1) + struct.pack("<HI", 0x6A03, 0)))
        text += "\x15\r"
        st = f"ObjectPool/_{num}"
        pool[st + "/\x01CompObj"] = compobj(f"Forms.{c['kind']}.1")
        pool[st + "/\x03ObjInfo"] = struct.pack("<HH", 0x1000, 3)  # ODTPersist1.fOCX（ビット12）・cf
        pool[st + "/\x03OCXNAME"] = (c["name"] + "\0").encode("utf-16-le")
        pool[st + "/contents"] = c["bin"]
    fc0 = 0x400
    wd = bytearray(0x400) + text.encode("utf-16-le")
    wd += bytes((-len(wd)) % 512)
    # CHPX の FKP（2.9.44）：rgfc（crun+1）・rgb（CHPX の位置/2）・CHPX（cb と grpprl）・最後のバイトが crun
    cuts = sorted({0, len(text)} | {cp for cp, _ in runs} | {cp + 1 for cp, _ in runs})
    page = bytearray(512)
    crun = len(cuts) - 1
    struct.pack_into(f"<{crun + 1}I", page, 0, *(fc0 + 2 * x for x in cuts))
    top = 511
    props = dict(runs)
    for k, cp in enumerate(cuts[:-1]):
        g = props.get(cp)
        if g is None:
            continue
        top -= len(g) + 1
        top -= top % 2
        page[top] = len(g)
        page[top + 1:top + 1 + len(g)] = g
        page[4 * (crun + 1) + k] = top // 2
    page[511] = crun
    chpx_pn = len(wd) // 512
    wd += page
    # PAPX の FKP（2.9.175）：段落はみな書式なし（istd 0）
    pcuts = [0] + [i + 1 for i, ch in enumerate(text) if ch == "\r"]
    page = bytearray(512)
    n = len(pcuts) - 1
    struct.pack_into(f"<{n + 1}I", page, 0, *(fc0 + 2 * x for x in pcuts))
    page[500] = 1  # PAPX：cb=1（grpprl なし。istd の2バイト）
    for k in range(n):
        page[4 * (n + 1) + 13 * k] = 250
    page[511] = n
    papx_pn = len(wd) // 512
    wd += page
    table = bytearray()
    clx_at = len(table)
    table += b"\x02" + struct.pack("<I", 12 + 4) + struct.pack("<II", 0, len(text)) + struct.pack("<HIH", 0, fc0, 0)
    chpx_at = len(table)
    table += struct.pack("<III", fc0, fc0 + 2 * len(text), chpx_pn)
    papx_at = len(table)
    table += struct.pack("<III", fc0, fc0 + 2 * len(text), papx_pn)
    fib = bytearray(900)
    struct.pack_into("<HH", fib, 0, 0xA5EC, 0x00C1)
    struct.pack_into("<H", fib, 0x0A, 0x0200)  # fWhichTblStm：1Table
    struct.pack_into("<H", fib, 32, 14)
    struct.pack_into("<H", fib, 62, 22)
    lw = [0] * 22
    lw[0], lw[3] = len(wd), len(text)
    struct.pack_into("<22i", fib, 64, *lw)
    struct.pack_into("<H", fib, 152, 93)
    from todoku.doc.book import FC_NAMES

    def put(name, fc, lcb):
        k = FC_NAMES.index(name)
        struct.pack_into("<II", fib, 154 + 8 * k, fc, lcb)

    put("Clx", clx_at, 1 + 4 + 16)
    put("PlcfBteChpx", chpx_at, 12)
    put("PlcfBtePapx", papx_at, 12)
    wd[:900] = fib
    return _cfb_tree({"WordDocument": bytes(wd), "1Table": bytes(table), **pool})


def test_doc_read_fill_prove():
    from todoku.doc.audit import prove
    from todoku.doc.book import DocFile
    from todoku.doc.fill import fill
    from todoku.docx.fill import written_of
    from todoku.docx.window import build, to_markdown

    blob = ax_doc([{"name": "TextBox1", "kind": "TextBox", "bin": morph(1, "", max_length=20), "before": "氏名"},
                   {"name": "CheckBox1", "kind": "CheckBox", "bin": morph(4, "0", "同意する")}])
    doc = DocFile(blob, "a.doc")
    win = build(doc)
    f = {x["id"]: x for x in win["fields"]}
    assert f["TextBox1"]["kind"] == "text" and f["CheckBox1"]["value"] is False
    assert "☐同意する［CheckBox1］" in to_markdown(win)
    res = fill(DocFile(blob, "a.doc"), {"TextBox1": "鈴木　一郎", "CheckBox1": True})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors
    after = {x["id"]: x["value"] for x in build(DocFile(res.blob, "a.doc"))["fields"]}
    assert after == {"TextBox1": "鈴木　一郎", "CheckBox1": True}
    w = written_of(res.patches)
    w["TextBox1"] = "佐藤"
    assert not prove(blob, res.blob, w).ok
