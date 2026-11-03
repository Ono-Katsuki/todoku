"""字を組み立てて（シェーピング）描く文字：アラビア文字（ペルシャ語・ウルドゥー語も）・ヘブライ文字・インド系の文字・タイ文字。

1 行の字を、見た目の順の字形（字形の番号と位置）にする。標準ライブラリだけで、次の決まりに従う：
- 並びの向き：Unicode 双方向アルゴリズム（UAX #9）。段落の向きは最初の強い字、数字は左から右。明示の埋め込み・上書き・
  隔離の字（LRE・RLO・RLI など）は扱わず、断る
- アラビア文字の結合：Unicode の ArabicShaping.txt の結合の型（data/shaping.json.gz に入れた。scripts/make_shaping_data.py
  が UCD から作る）で、語頭・語中・語末・独立の形を決め、GSUB の isol・fina・medi・init、rlig（ラーム・アリフ）・calt・liga を
  当てる。記号（ハラカート）は GPOS の mark・mkmk で付ける
- インド系の文字：Microsoft の「Developing OpenType Fonts for <文字>」（新しい仕様：dev2・bng2 など）の音節の組み立て。
  音節の切り方・基の子音・レーフ・前に置く母音記号・前に置く子音の並べ替えは、HarfBuzz の Indic の組み立てと同じにする
  （Unicode の IndicSyllabicCategory・IndicPositionalCategory を使う）。機能は locl・ccmp → nukt・akhn・rphf・rkrf・pref・
  blwf・abvf・half・pstf・vatu・cjct（1 つずつ）→ 並べ直し → init・pres・abvs・blws・psts・haln・calt・clig・liga・rlig・
  rclt、GPOS は abvm・blwm・dist・kern・mark・mkmk・curs
- タイ文字：サラ・アム（ำ）を分け、GPOS の mark・mkmk で記号を付ける
確かに組み立てられないもの（書体にない字形・壊れた音節・書体にない語中形・できないラーム・アリフ・レーフ・半分の形など）は、
その字の並びを名指しして断る（GlyphError）。点線の丸（U+25CC）を入れて描くことはしない。

書体：このパソコンにある、GSUB を持ち、埋め込みを許す TrueType（glyf）の書体を、文字ごとの候補から選ぶ。
環境変数 TODOKU_PDF_FONT_<文字>（ARABIC・HEBREW・DEVANAGARI・BENGALI・GURMUKHI・GUJARATI・ORIYA・TAMIL・TELUGU・
KANNADA・THAI・LAO）で書体のファイル（.ttc なら「パス#番号」）を指せる。空にすると、その文字は描かない。
"""

from __future__ import annotations

import bisect
import gzip
import hashlib
import json
import os
import re
import struct
import unicodedata
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

from . import otl
from .glyphs import GlyphError


class ShapeError(GlyphError):
    pass


# ====================================================================== Unicode の文字データ
@lru_cache(maxsize=1)
def _data() -> dict:
    p = Path(__file__).parent / "data" / "shaping.json.gz"
    d = json.loads(gzip.decompress(p.read_bytes()))
    for k in ("joining", "indic_cat", "indic_pos", "scripts"):
        rows = d[k]
        d[k] = ([r[0] for r in rows], rows)
    d["mirror"] = {a: b for a, b in d["mirror"]}
    d["brackets"] = {a: (b, t) for a, b, t in d["brackets"]}
    return d


def _range_value(name: str, cp: int):
    starts, rows = _data()[name]
    i = bisect.bisect_right(starts, cp) - 1
    if i >= 0 and rows[i][0] <= cp <= rows[i][1]:
        return rows[i][2]
    return None


def script_of(ch: str) -> str:
    """Scripts.txt の文字の種類（組み立てる文字・Common・Inherited のほかは "Other"）。"""
    return _range_value("scripts", ord(ch)) or "Other"


def joining_type(cp: int) -> str:
    v = _range_value("joining", cp)
    if v is not None:
        return v
    return "T" if unicodedata.category(chr(cp)) in ("Mn", "Me", "Cf") else "U"


# ====================================================================== 文字ごとの決まり
@dataclass(frozen=True)
class IndicConfig:
    virama: int
    reph_pos: str  # before_post・after_sub・before_sub・after_main・after_post
    reph_mode: str  # implicit・explicit
    blwf_post_only: bool
    half_check: bool  # 音節の中の見える半分の形の印（ヴィラーマ）を、組み立てられなかったものとして断る


@dataclass(frozen=True)
class ScriptInfo:
    name: str
    kind: str  # arabic・indic・simple
    tags: tuple[str, ...]
    env: str
    rtl: bool
    fonts: tuple[tuple[str, str], ...]  # (ファイル, PostScript 名（空ならどれでも）)
    indic: IndicConfig | None = None
    ja: str = ""


_MAC = "/System/Library/Fonts/"
_SUP = "/System/Library/Fonts/Supplemental/"
_NOTO = ("/usr/share/fonts/truetype/noto/", "/usr/share/fonts/noto/", "/usr/share/fonts/google-noto/")
_WIN = "C:\\Windows\\Fonts\\"


def _noto(name: str) -> tuple[tuple[str, str], ...]:
    return tuple((d + name, "") for d in _NOTO)


SCRIPTS: dict[str, ScriptInfo] = {s.name: s for s in [
    ScriptInfo("Arabic", "arabic", ("arab",), "ARABIC", True,
               ((_SUP + "Tahoma.ttf", ""), (_SUP + "Arial Unicode.ttf", ""), *_noto("NotoNaskhArabic-Regular.ttf"),
                *_noto("NotoSansArabic-Regular.ttf"), ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", ""),
                (_WIN + "tahoma.ttf", ""), (_WIN + "arial.ttf", ""), (_MAC + "NotoNastaliq.ttc", "NotoNastaliqUrdu"),
                *_noto("NotoNastaliqUrdu-Regular.ttf")), ja="アラビア文字"),
    ScriptInfo("Hebrew", "simple", ("hebr",), "HEBREW", True,
               ((_SUP + "Tahoma.ttf", ""), (_SUP + "Arial Unicode.ttf", ""), *_noto("NotoSansHebrew-Regular.ttf"),
                ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", ""), (_WIN + "tahoma.ttf", ""), (_WIN + "arial.ttf", "")),
               ja="ヘブライ文字"),
    ScriptInfo("Devanagari", "indic", ("dev2",), "DEVANAGARI", False,
               ((_SUP + "Devanagari Sangam MN.ttc", "DevanagariSangamMN"), *_noto("NotoSansDevanagari-Regular.ttf"),
                (_WIN + "Nirmala.ttf", ""), (_WIN + "Nirmala.ttc", "NirmalaUI"), (_WIN + "mangal.ttf", "")),
               IndicConfig(0x094D, "before_post", "implicit", False, True), "デーヴァナーガリー文字"),
    ScriptInfo("Bengali", "indic", ("bng2",), "BENGALI", False,
               ((_SUP + "Bangla Sangam MN.ttc", "BanglaSangamMN"), (_SUP + "Bangla MN.ttc", "BanglaMN"),
                *_noto("NotoSansBengali-Regular.ttf"), (_WIN + "Nirmala.ttf", ""), (_WIN + "vrinda.ttf", "")),
               IndicConfig(0x09CD, "after_sub", "implicit", False, True), "ベンガル文字"),
    ScriptInfo("Gurmukhi", "indic", ("gur2",), "GURMUKHI", False,
               ((_MAC + "MuktaMahee.ttc", "MuktaMahee-Regular"), (_SUP + "Gurmukhi Sangam MN.ttc", "GurmukhiSangamMN"),
                (_SUP + "Gurmukhi MN.ttc", "GurmukhiMN"), *_noto("NotoSansGurmukhi-Regular.ttf"), (_WIN + "Nirmala.ttf", ""),
                (_WIN + "raavi.ttf", "")),
               IndicConfig(0x0A4D, "before_sub", "implicit", False, False), "グルムキー文字"),
    ScriptInfo("Gujarati", "indic", ("gjr2",), "GUJARATI", False,
               ((_SUP + "Gujarati Sangam MN.ttc", "GujaratiSangamMN"), *_noto("NotoSansGujarati-Regular.ttf"),
                (_WIN + "Nirmala.ttf", ""), (_WIN + "shruti.ttf", "")),
               IndicConfig(0x0ACD, "before_post", "implicit", False, True), "グジャラート文字"),
    ScriptInfo("Oriya", "indic", ("ory2",), "ORIYA", False,
               ((_MAC + "NotoSansOriya.ttc", "NotoSansOriya"), (_SUP + "Oriya Sangam MN.ttc", "OriyaSangamMN"),
                *_noto("NotoSansOriya-Regular.ttf"), (_WIN + "Nirmala.ttf", ""), (_WIN + "kalinga.ttf", "")),
               IndicConfig(0x0B4D, "after_main", "implicit", False, True), "オリヤー文字"),
    ScriptInfo("Tamil", "indic", ("tml2",), "TAMIL", False,
               ((_SUP + "Tamil MN.ttc", "TamilMN"), *_noto("NotoSansTamil-Regular.ttf"), (_WIN + "Nirmala.ttf", ""),
                (_WIN + "latha.ttf", "")),
               IndicConfig(0x0BCD, "after_post", "implicit", False, False), "タミル文字"),
    ScriptInfo("Telugu", "indic", ("tel2",), "TELUGU", False,
               ((_SUP + "Telugu Sangam MN.ttc", "TeluguSangamMN"), (_SUP + "Telugu MN.ttc", "TeluguMN"),
                *_noto("NotoSansTelugu-Regular.ttf"), (_WIN + "Nirmala.ttf", ""), (_WIN + "gautami.ttf", "")),
               IndicConfig(0x0C4D, "after_post", "explicit", True, True), "テルグ文字"),
    ScriptInfo("Kannada", "indic", ("knd2",), "KANNADA", False,
               ((_MAC + "NotoSansKannada.ttc", "NotoSansKannada-Regular"), (_SUP + "Kannada Sangam MN.ttc", "KannadaSangamMN"),
                *_noto("NotoSansKannada-Regular.ttf"), (_WIN + "Nirmala.ttf", ""), (_WIN + "tunga.ttf", "")),
               IndicConfig(0x0CCD, "after_post", "implicit", True, True), "カンナダ文字"),
    ScriptInfo("Thai", "simple", ("thai",), "THAI", False,
               ((_SUP + "Tahoma.ttf", ""), *_noto("NotoSansThai-Regular.ttf"), (_WIN + "tahoma.ttf", ""),
                (_WIN + "LeelawUI.ttf", "")), ja="タイ文字"),
    ScriptInfo("Lao", "simple", ("lao ",), "LAO", False,
               ((_SUP + "Lao Sangam MN.ttf", ""), (_SUP + "Lao MN.ttc", "LaoMN"), *_noto("NotoSansLao-Regular.ttf"),
                (_WIN + "LaoUI.ttf", "")), ja="ラーオ文字"),
]}
# 組み立てが要るが、まだ確かめていない文字（書体があっても断る）
UNSUPPORTED = {"Malayalam": "マラヤーラム文字", "Sinhala": "シンハラ文字", "Khmer": "クメール文字", "Myanmar": "ミャンマー文字",
               "Syriac": "シリア文字", "Thaana": "ターナ文字", "Nko": "ンコ文字"}
COMPLEX = set(SCRIPTS) | set(UNSUPPORTED)

# Ra（レーフになる子音）：HarfBuzz の Indic の表と同じ
_RA = {0x0930, 0x09B0, 0x09F0, 0x0A30, 0x0AB0, 0x0B30, 0x0BB0, 0x0C30, 0x0CB0, 0x0D30}
_NAME = {"isol": "独立形", "init": "語頭形", "medi": "語中形", "fina": "語末形"}


def _ja(script: str) -> str:
    s = SCRIPTS.get(script)
    return (s.ja if s else UNSUPPORTED.get(script, script)) or script


def _u(text: str) -> str:
    return " ".join(f"U+{ord(c):04X}" for c in text)


# ====================================================================== 書体
class ShapeFont:
    def __init__(self, tt, script: ScriptInfo, tag: str):
        self.tt = tt
        self.script = script
        self.tag = tag
        t = tt.t
        lens = (len(t.get("GDEF", b"")), len(t.get("GSUB", b"")), len(t.get("GPOS", b"")))
        self.gdef = otl.GDEF(None if lens in otl.GDEF_BLOCKLIST else t.get("GDEF"))
        self.gsub = otl.Layout(t.get("GSUB"), True)
        self.gpos = otl.Layout(t.get("GPOS"), False)
        self.sub_feats = self.gsub.features(tag)
        self.pos_feats = self.gpos.features(self.gpos.script((tag, "DFLT")))
        self.engine = otl.Shaper(self.gdef, self.gsub, self.gpos, tt.num)
        self.key = "TdkX" + hashlib.sha1(f"{tt.path}#{tt.offset}".encode()).hexdigest()[:6].upper()
        self.glyph_chars: dict[int, str] = {}  # ToUnicode：字形 → 元の字（最初に描いたもの）

    @property
    def name(self) -> str:
        return self.tt.name

    def gid(self, cp: int) -> int | None:
        g = self.tt.cmap.get(cp)
        return g if g and g < self.tt.num else None


def _open_face(path: str, want: str, script: ScriptInfo):
    from .embedfont import TrueType, collection_faces

    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError:
        return None
    faces = collection_faces(data)
    picks = []
    for off in faces:
        try:
            tt = TrueType(path, data, off)
        except Exception:  # noqa: BLE001
            continue
        picks.append(tt)
    if want:
        picks = [tt for tt in picks if tt.name == want] or []
    for tt in picks:
        if not tt.embeddable:
            continue
        gsub = otl.Layout(tt.t.get("GSUB"), True)
        tag = gsub.script(script.tags)
        if tag is None:
            continue
        sf = ShapeFont(tt, script, tag)
        if script.kind == "arabic" and not all(f in sf.sub_feats for f in ("init", "medi", "fina")):
            continue
        if script.name in ("Thai", "Lao") and "mark" not in sf.pos_feats:
            continue
        return sf
    return None


_REGISTRY: dict[str, ShapeFont] = {}  # 開いた書体：資源の名前（TdkX…）→ 書体
_OPENED: dict[tuple, list] = {}  # (文字, 環境変数) → [開いた書体, 次に試す候補の番号]


def _env_font(path_idx: str, info: ScriptInfo) -> ShapeFont | None:
    path, _, idx = path_idx.partition("#")
    from .embedfont import TrueType, collection_faces

    try:
        data = Path(path).read_bytes()
        off = collection_faces(data)[int(idx or 0)]
        tt = TrueType(path, data, off)
    except (OSError, ValueError, IndexError):
        return None
    return _open_face(path, tt.name, info)


def nth_font(script: str, i: int) -> ShapeFont | None:
    """その文字を組み立てられる i 番目の書体（使う順。要るときに開く）。環境変数で指したら、その書体だけ。"""
    info = SCRIPTS[script]
    env = os.environ.get("TODOKU_PDF_FONT_" + info.env)
    st = _OPENED.setdefault((script, env), [[], 0])
    fonts, nxt = st
    if env is not None:
        if nxt == 0:
            st[1] = 1
            sf = _env_font(env, info) if env else None
            if sf is not None:
                fonts.append(sf)
                _REGISTRY[sf.key] = sf
        return fonts[i] if i < len(fonts) else None
    while len(fonts) <= i and st[1] < len(info.fonts):
        path, want = info.fonts[st[1]]
        st[1] += 1
        if not os.path.exists(path):
            continue
        sf = _open_face(path, want, info)
        if sf is not None and all(f.key != sf.key for f in fonts):
            fonts.append(sf)
            _REGISTRY[sf.key] = sf
    return fonts[i] if i < len(fonts) else None


def font_for(script: str) -> ShapeFont | None:
    return nth_font(script, 0)


def font_by_key(key: str) -> ShapeFont | None:
    return _REGISTRY.get(key)


def font_by_name(name: str) -> ShapeFont | None:
    """PostScript 名の書体（このプロセスで組み立てに使ったもの。なければ候補を開いて探す）。"""
    for sf in _REGISTRY.values():
        if sf.name == name:
            return sf
    for s in SCRIPTS:
        i = 0
        while True:
            sf = nth_font(s, i)
            if sf is None:
                break
            if sf.name == name:
                return sf
            i += 1
    return None


# ====================================================================== 双方向（UAX #9）
_EXPLICIT = {"LRE", "RLE", "LRO", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"}


def bidi_levels(text: str, para: int | None = None) -> tuple[list[int], int]:
    """1 行（1 つの段落）の字ごとの並びの深さ（偶数：左から右、奇数：右から左）と、段落の深さ。
    UAX #9 の P2・P3、W1〜W7、N0〜N2、I1・I2、L1（明示の字は扱わない）。"""
    n = len(text)
    orig = [unicodedata.bidirectional(c) or "L" for c in text]
    for i, t in enumerate(orig):
        if t in _EXPLICIT:
            raise ShapeError(f"「{text}」の {_u(text[i])}（双方向の埋め込み・上書き・隔離の字）は、扱えません")
    if para is None:
        para = 0
        for t in orig:
            if t in ("L",):
                break
            if t in ("R", "AL"):
                para = 1
                break
    sos = "R" if para % 2 else "L"
    t = list(orig)
    # BN（ZWJ など）は前の字の種類として扱う（X9 で取り除くのと同じ結果になる）
    prev = sos
    for i in range(n):
        if t[i] in ("NSM", "BN"):
            t[i] = prev if not (t[i] == "NSM" and prev in ("B", "S")) else "ON"
        prev = t[i]
    # W2・W3
    last = sos
    for i in range(n):
        if t[i] in ("L", "R", "AL"):
            last = t[i]
        elif t[i] == "EN" and last == "AL":
            t[i] = "AN"
    t = ["R" if x == "AL" else x for x in t]
    # W4
    for i in range(1, n - 1):
        if t[i] == "ES" and t[i - 1] == "EN" and t[i + 1] == "EN":
            t[i] = "EN"
        elif t[i] == "CS" and t[i - 1] == t[i + 1] and t[i - 1] in ("EN", "AN"):
            t[i] = t[i - 1]
    # W5
    i = 0
    while i < n:
        if t[i] == "ET":
            j = i
            while j < n and t[j] == "ET":
                j += 1
            if (i > 0 and t[i - 1] == "EN") or (j < n and t[j] == "EN"):
                for k in range(i, j):
                    t[k] = "EN"
            i = j
        else:
            i += 1
    # W6
    t = ["ON" if x in ("ES", "ET", "CS") else x for x in t]
    # W7
    last = sos
    for i in range(n):
        if t[i] in ("L", "R"):
            last = t[i]
        elif t[i] == "EN" and last == "L":
            t[i] = "L"
    e = "R" if para % 2 else "L"

    def strong(x):
        return "R" if x in ("R", "EN", "AN") else ("L" if x == "L" else None)

    # N0：対になる括弧（BD16）
    br = _data()["brackets"]
    stack, pairs = [], []
    for i, c in enumerate(text):
        if t[i] != "ON":
            continue
        cp = ord(c)
        info = br.get(cp)
        if info is None:
            continue
        mate, kind = info
        if kind == "o":
            if len(stack) >= 63:
                break
            stack.append((mate, i))
        else:
            for k in range(len(stack) - 1, -1, -1):
                m, j = stack[k]
                if m == cp or (m, cp) in ((0x232A, 0x3009), (0x3009, 0x232A)):
                    pairs.append((j, i))
                    del stack[k:]
                    break
    for a, b in sorted(pairs):
        inside = {strong(t[k]) for k in range(a + 1, b)} - {None}
        if e in inside:
            d = e
        elif inside:
            d = None
            for k in range(a - 1, -1, -1):
                s = strong(t[k])
                if s:
                    d = s
                    break
            d = d or sos
        else:
            continue
        t[a] = t[b] = d
    # N1・N2
    ni = ("B", "S", "WS", "ON")
    i = 0
    while i < n:
        if t[i] in ni:
            j = i
            while j < n and t[j] in ni:
                j += 1
            before = strong(t[i - 1]) if i > 0 else sos
            after = strong(t[j]) if j < n else sos
            d = before if before == after else e
            for k in range(i, j):
                t[k] = d
            i = j
        else:
            i += 1
    levels = []
    for x in t:
        lv = para
        if para % 2 == 0:
            lv += 1 if x == "R" else (2 if x in ("AN", "EN") else 0)
        else:
            lv += 1 if x in ("L", "EN", "AN") else 0
        levels.append(lv)
    # L1：区切り・行の終わりの空白は段落の深さ
    k = n
    while k > 0 and orig[k - 1] in ("WS", "BN") or (k > 0 and orig[k - 1] in _EXPLICIT):
        k -= 1
        levels[k] = para
    for i in range(n):
        if orig[i] in ("S", "B"):
            levels[i] = para
            k = i
            while k > 0 and orig[k - 1] in ("WS", "BN"):
                k -= 1
                levels[k] = para
    return levels, para


def reorder(levels: list[int]) -> list[int]:
    """L2：見た目の順（左から右）の、論理の位置。"""
    idx = list(range(len(levels)))
    if not levels:
        return idx
    hi = max(levels)
    lo = min(lv for lv in levels if lv % 2) if any(lv % 2 for lv in levels) else hi + 1
    for lv in range(hi, lo - 1, -1):
        i = 0
        while i < len(idx):
            if levels[idx[i]] >= lv:
                j = i
                while j < len(idx) and levels[idx[j]] >= lv:
                    j += 1
                idx[i:j] = idx[i:j][::-1]
                i = j
            else:
                i += 1
    return idx


# ====================================================================== 行の組み立て
@dataclass
class Seg:
    key: str
    shaped: bool
    logical: str
    width: float  # 1000 分の 1（字の大きさ 1 のとき 1000）
    glyphs: list = field(default_factory=list)  # 組み立てた字形：(gid, x, y, 幅)（1000 分の 1、区切りの始まりから）
    text: str = ""  # 組み立てない字（見た目の順）
    font: ShapeFont | None = None


@dataclass
class Line:
    text: str
    segs: list[Seg]
    width: float


@lru_cache(maxsize=8192)
def needs(text: str) -> bool:
    """組み立て（シェーピング）が要る字があるか。"""
    for c in text:
        if c.isascii():
            continue
        if script_of(c) in COMPLEX or unicodedata.bidirectional(c) in ("R", "AL", "AN"):
            return True
    return False


def _env_key() -> tuple:
    return tuple(os.environ.get("TODOKU_PDF_FONT_" + s.env) for s in SCRIPTS.values())


def line(text: str) -> Line:
    """1 行を組み立てる。組み立てられなければ ShapeError（理由と、その字の並び）。"""
    return _line(text, _env_key())


@lru_cache(maxsize=4096)
def _line(text: str, _env) -> Line:
    """書体は、文字ごとの候補を順に試す：ある候補で組み立てられない字の並びがあれば、その文字だけ次の候補で
    組み立て直す（どの候補でもだめなら、最初の候補の理由で断る）。"""
    levels, _para = bidi_levels(text)
    n = len(text)
    # 文字の種類：Common・Inherited は前（なければ後ろ）の組み立てる文字に寄せる
    scr = [script_of(c) for c in text]
    res = list(scr)
    last = None
    for i in range(n):
        if res[i] in COMPLEX:
            last = res[i]
        elif res[i] in ("Common", "Inherited") and last is not None:
            res[i] = last
    nxt = None
    for i in range(n - 1, -1, -1):
        if res[i] in COMPLEX:
            nxt = res[i]
        elif res[i] in ("Common", "Inherited") and nxt is not None:
            res[i] = nxt
    for i, s in enumerate(res):
        if s in UNSUPPORTED:
            raise ShapeError(f"「{text}」の「{text[i]}」（{_u(text[i])}）は、{_ja(s)}です。この文字の組み立て（シェーピング）は"
                             "まだ確かめていないので、PDF に描きません")
    used = [s for s in dict.fromkeys(res) if s in SCRIPTS]
    for s in used:
        if nth_font(s, 0) is None:
            i = next(k for k, x in enumerate(res) if x == s)
            c = text[i]
            env = "TODOKU_PDF_FONT_" + SCRIPTS[s].env
            raise ShapeError(f"「{c}」（{_u(c)}）は{_ja(s)}で、字の組み立て（シェーピング）が要りますが、このパソコンに、"
                             f"組み立ての表（GSUB）を持ち埋め込める {_ja(s)}の TrueType の書体が見つかりません"
                             f"（環境変数 {env} で書体を指せます）")
    pick = {s: 0 for s in used}
    first_err: dict[str, ShapeError] = {}
    while True:
        try:
            return _build(text, levels, scr, res, {s: nth_font(s, pick[s]) for s in used})
        except ShapeError as e:
            s = getattr(e, "script", None)
            if s not in pick:
                raise
            first_err.setdefault(s, e)
            pick[s] += 1
            if nth_font(s, pick[s]) is None:
                raise first_err[s] from None


def _build(text: str, levels: list[int], scr: list[str], res: list[str], fonts: dict) -> Line:
    from .glyphs import choose

    n = len(text)
    kinds = []
    for i, c in enumerate(text):
        s = res[i]
        if s in fonts:
            sf = fonts[s]
            if scr[i] in ("Common", "Inherited") and unicodedata.category(c) not in ("Mn", "Me") \
                    and c not in "\u200c\u200d" and sf.gid(_mirror(c) if levels[i] % 2 else ord(c)) is None:
                kinds.append(None)  # 組み立てる書体にない記号・数字：ふつうの書体で
                continue
            kinds.append(s)
        else:
            kinds.append(None)
    # 区切り：(深さ, 種類) が同じ続き
    runs = []
    for i in range(n):
        k = (levels[i], kinds[i])
        if runs and runs[-1][0] == k and runs[-1][2] == i:
            runs[-1][2] = i + 1
        else:
            runs.append([k, i, i + 1])
    segs_logical = []
    for (lv, kind), a, b in runs:
        sub = text[a:b]
        rtl = lv % 2 == 1
        if kind is None:
            chars = [chr(_mirror(c)) if rtl else c for c in sub]
            if rtl:
                chars.reverse()
            vis = "".join(chars)
            for c in vis:
                choose(c)  # 描けない字は GlyphError
            segs_logical.append((lv, None, sub, vis, None))
        else:
            sf = fonts[kind]
            try:
                gl, w = shape_run(sf, sub, rtl)
            except ShapeError as e:
                e.script = kind  # type: ignore[attr-defined]
                raise
            segs_logical.append((lv, kind, sub, None, (sf, gl, w)))
    # 見た目の順（区切りの単位で L2）
    order = reorder([s[0] for s in segs_logical])
    segs = []
    total = 0.0
    for k in order:
        lv, kind, sub, vis, shaped = segs_logical[k]
        if kind is None:
            groups: list[list] = []
            for c in vis:
                key, _code, w = choose(c)
                if groups and groups[-1][0] == key:
                    groups[-1][1] += c
                    groups[-1][2] += w
                else:
                    groups.append([key, c, w])
            for key, s, w in groups:
                segs.append(Seg(key, False, s, w, text=s))
                total += w
        else:
            sf, gl, w = shaped
            segs.append(Seg(sf.key, True, sub, w, glyphs=gl, font=sf))
            total += w
    return Line(text, segs, total)


def _mirror(c: str) -> int:
    return _data()["mirror"].get(ord(c), ord(c))


# ====================================================================== 1 つの区切りを組み立てる
_KEEP_DECOMPOSED = {0x0931, 0x09DC, 0x09DD, 0x0B94}  # HarfBuzz の Indic：分けない字


# 記号を並べる順（HarfBuzz の「modified combining class」：ヘブライ文字は SBL の順、アラビア文字はシャッダを先に、
# テルグ文字の長さの記号・タイ文字のサラ・ウ／ウーを前に。hb-unicode.hh と同じ）
_MCC = {10: 22, 11: 15, 12: 16, 13: 17, 14: 23, 15: 18, 16: 19, 17: 20, 18: 21, 19: 14, 20: 24, 21: 12, 22: 25, 23: 13,
        24: 10, 25: 11, 26: 26, 27: 28, 28: 29, 29: 30, 30: 31, 31: 32, 32: 33, 33: 27, 34: 34, 35: 35, 36: 36, 84: 4, 91: 5,
        103: 3, 130: 132, 132: 131}
# アラビア文字の、前に動かす記号（Arabic Mark Transient Reordering の MCM）
_MCM = {0x0654, 0x0655, 0x0658, 0x06DC, 0x06E3, 0x06E7, 0x06E8, 0x08CA, 0x08CB, 0x08CD, 0x08CE, 0x08CF, 0x08D3, 0x08F3}


def _mcc(cp: int) -> int:
    c = unicodedata.combining(chr(cp))
    return _MCC.get(c, c)


def _is_mark(cp: int) -> bool:
    return unicodedata.category(chr(cp)) in ("Mn", "Mc", "Me")


def _normalize(sf: ShapeFont, text: str) -> list[tuple[int, list[int]]]:
    """字 → [(字, 元の位置の並び)]。HarfBuzz の正規化と同じ：分ける（インド系の文字は、書体に字形があっても分けられれば
    分ける。ほかは書体に字形がなければ分ける）→ 記号を並べる → 書体に字形があれば合わせる。"""
    out: list[tuple[int, list[int]]] = []
    indic = sf.script.kind == "indic"
    for i, c in enumerate(text):
        cp = ord(c)
        if indic or sf.gid(cp) is None:
            parts = _decompose(c)
            if len(parts) > 1 and not (indic and cp in _KEEP_DECOMPOSED) and all(sf.gid(ord(p)) for p in parts):
                for p in parts:
                    out.append((ord(p), [i]))
                continue
        out.append((cp, [i]))
    # 記号の並び（0 でない結合の順の続きを、安定に並べる）
    k = 0
    while k < len(out):
        if _mcc(out[k][0]) == 0:
            k += 1
            continue
        j = k
        while j < len(out) and _mcc(out[j][0]) != 0:
            j += 1
        if j - k <= 32:
            seg = sorted(out[k:j], key=lambda x: _mcc(x[0]))
            if sf.script.kind == "arabic":
                seg = _arabic_mcm(seg)
            out[k:j] = seg
        k = j
    if len(out) < 2:
        return out
    # 合わせる：基の字（starter）と、間を挟まれていない記号
    res = [out[0]]
    starter = 0
    for cp, src in out[1:]:
        if _is_mark(cp):
            prev = res[-1][0]
            a = res[starter][0]
            comp = None
            if starter == len(res) - 1 or _mcc(prev) < _mcc(cp):
                if not (indic and _is_mark(a)):
                    if indic and a == 0x09AF and cp == 0x09BC:
                        comp = 0x09DF
                    else:
                        n = unicodedata.normalize("NFC", chr(a) + chr(cp))
                        comp = ord(n) if len(n) == 1 else None
                if comp is not None and sf.gid(comp) is None:
                    comp = None
            if comp is not None:
                res[starter] = (comp, res[starter][1] + src)
                continue
            if starter < len(res) - 1 and _mcc(prev) > _mcc(cp):
                starter = len(res)
        res.append((cp, src))
        if _mcc(cp) == 0:
            starter = len(res) - 1
    return res


def _arabic_mcm(seg: list) -> list:
    """HarfBuzz の reorder_marks_arabic：ccc 220・230 の MCM を、記号の並びの先に動かす。"""
    seg = list(seg)
    start = 0
    i = 0
    for cc in (220, 230):
        while i < len(seg) and _mcc(seg[i][0]) < cc:
            i += 1
        if i == len(seg):
            break
        if _mcc(seg[i][0]) > cc:
            continue
        j = i
        while j < len(seg) and _mcc(seg[j][0]) == cc and seg[j][0] in _MCM:
            j += 1
        if i == j:
            continue
        seg = seg[:start] + seg[i:j] + seg[start:i] + seg[j:]
        start += j - i
        i = j
    return seg


def _decompose(c: str) -> list[str]:
    d = unicodedata.decomposition(c)
    if not d or d.startswith("<"):
        return [c]
    out = []
    for h in d.split():
        out += _decompose(chr(int(h, 16)))
    return out


def _thai_sara_am(items: list[tuple[int, list[int]]]) -> list[tuple[int, list[int]]]:
    """タイ文字・ラーオ文字のサラ・アム（U+0E33・U+0EB3）を、ニッカヒット + サラ・アーに分け、ニッカヒットを前の上の記号
    （声調の記号など）の前に動かす（HarfBuzz の Thai と同じ）。"""
    out: list[tuple[int, list[int]]] = []

    def above(cp):
        x = cp & ~0x0080
        return 0x0E34 <= x <= 0x0E37 or 0x0E47 <= x <= 0x0E4E or x in (0x0E31, 0x0E3B)

    for cp, src in items:
        if (cp & ~0x0080) == 0x0E33:
            nik = (cp - 0x0E33 + 0x0E4D, src)
            aa = (cp - 1, src)
            start = len(out)
            while start > 0 and above(out[start - 1][0]):
                start -= 1
            out.insert(start, nik)
            out.append(aa)
        else:
            out.append((cp, src))
    return out


def shape_run(sf: ShapeFont, text: str, rtl: bool) -> tuple[list, float]:
    """組み立てる文字の 1 つの区切り（論理の順）→ (見た目の順の字形 [(gid, x, y, 幅)]（1000 分の 1）, 幅)。"""
    items = _normalize(sf, text)
    if sf.script.name in ("Thai", "Lao"):
        items = _thai_sara_am(items)
    buf: list[otl.Info] = []
    for cp, src in items:
        if rtl:
            m = _data()["mirror"].get(cp)
            if m is not None and sf.gid(m) is not None:
                cp = m
        cat = unicodedata.category(chr(cp))
        g = sf.gid(cp)
        if g is None:
            if cp in (0x200C, 0x200D):
                g = -1
            elif cat == "Cf":
                continue  # 字形のない書式の字（LRM・RLM など）：組み立てに使わない
            else:
                c = chr(cp)
                raise ShapeError(f"「{text}」の「{c}」（{_u(c)}）の字形が、書体 {sf.name} にありません")
        elif not sf.tt.glyph(g) and not chr(cp).isspace() and cat not in ("Cf", "Zs"):
            c = chr(cp)
            raise ShapeError(f"「{text}」の「{c}」（{_u(c)}）の字形の中身が、書体 {sf.name} にありません（白く抜けます）")
        inf = otl.Info(g, cp, src[0], "".join(text[i] for i in src))
        inf.extra = None
        inf.cat = ""
        sf.engine.set_class(inf, synth_mark=cat == "Mn")  # GDEF のないとき：Mn だけを記号に（HarfBuzz と同じ）
        buf.append(inf)
    sf.engine.rtl = rtl
    kind = sf.script.kind
    if kind == "arabic":
        buf, gpos_stages, zero = _arabic(sf, buf, text)
    elif kind == "indic":
        buf, gpos_stages, zero = _indic(sf, buf, text)
    else:
        buf, gpos_stages, zero = _simple(sf, buf, text)
    # 字形の確かめ：.notdef・書体にない字形・点線の丸を入れていない
    for inf in buf:
        if inf.gid == 0 or inf.gid >= sf.tt.num:
            raise ShapeError(f"「{text}」の「{inf.chars}」が、書体 {sf.name} の字形にない形（.notdef）になります")
    dotted = sf.gid(0x25CC)
    if dotted is not None and "\u25cc" not in text and any(inf.gid == dotted for inf in buf):
        raise ShapeError(f"「{text}」の組み立てで、点線の丸（U+25CC）が入ります（壊れた並び）")
    for inf in buf:
        inf.xa = sf.tt.adv[inf.gid] if inf.gid >= 0 else 0
        inf.ya = inf.xo = inf.yo = 0
        inf.attach = -1
        inf.atype = 0
    if gpos_stages:
        sf.engine.run(buf, gpos_stages, False)
    # ZWJ・ZWNJ（字形のないもの・あるもの）は描かない（ActualText に残る）
    keep = [k for k, inf in enumerate(buf) if inf.cp not in (0x200C, 0x200D) or inf.ligated]
    if len(keep) != len(buf):
        remap = {old: new for new, old in enumerate(keep)}
        nb = []
        for k in keep:
            inf = buf[k]
            if inf.attach >= 0:
                inf.attach = remap.get(inf.attach, -1)
                if inf.attach < 0:
                    inf.atype = 0
            nb.append(inf)
        buf = nb
    placed, width = otl.positions(buf, rtl, zero)
    sc = 1000.0 / sf.tt.upem
    out = []
    for inf, x, y in placed:
        out.append((inf.gid, x * sc, y * sc, round(sf.tt.width(inf.gid), 1)))
        if inf.chars and inf.gid not in sf.glyph_chars:
            sf.glyph_chars[inf.gid] = inf.chars
    return out, width * sc


def _stages(sf: ShapeFont, spec: list, gsub: bool, masks: dict[str, int]) -> list[otl.Stage]:
    """spec：[(pause, [(機能, 全体か, auto_zwj, auto_zwnj, 音節ごと)])] → 段階。同じ機能は最初の段階だけ。"""
    feats = sf.sub_feats if gsub else sf.pos_feats
    seen = set()
    out = []
    for pause, fl in spec:
        items: dict[int, list] = {}
        for tag, _glob, azwj, azwnj, per in fl:
            if tag in seen:
                continue
            seen.add(tag)
            if tag not in feats:
                continue
            if tag not in masks:
                masks[tag] = 1 << len(masks)
            for li in feats[tag]:
                if li in items:
                    it = items[li]
                    it[1] |= masks[tag]
                    it[2] = it[2] and azwj
                    it[3] = it[3] and azwnj
                else:
                    items[li] = [li, masks[tag], azwj, azwnj, per, tag]
        out.append(otl.Stage([tuple(items[k]) for k in sorted(items)], pause))
    return out


def _global_mask(spec: list, masks: dict[str, int]) -> int:
    m = 0
    for _p, fl in spec:
        for tag, glob, *_r in fl:
            if glob and tag in masks:
                m |= masks[tag]
    return m


_GPOS_COMMON = [(None, [("abvm", True, True, True, False), ("blwm", True, True, True, False),
                        ("mark", True, False, False, False), ("mkmk", True, False, False, False),
                        ("curs", True, True, True, False), ("dist", True, True, True, False),
                        ("kern", True, True, True, False)])]
_GSUB_COMMON = [("ccmp", True, True, True, False), ("locl", True, True, True, False), ("rlig", True, True, True, False),
                ("calt", True, True, True, False), ("clig", True, True, True, False), ("liga", True, True, True, False),
                ("rclt", True, True, True, False)]


def _run_gsub(sf: ShapeFont, buf, spec, masks, local: dict[int, int] | None = None):
    stages = _stages(sf, spec, True, masks)
    gm = _global_mask(spec, masks)
    for k, inf in enumerate(buf):
        inf.mask = gm | (local.get(k, 0) if local else 0)
    try:
        return sf.engine.run(buf, stages, True)
    except GlyphError:
        raise
    except (ValueError, IndexError, KeyError, struct.error) as e:
        raise ShapeError(f"書体 {sf.name} の組み立ての表（GSUB）を読めません（{e}）") from None


def _gpos(sf: ShapeFont, masks: dict[str, int]):
    st = _stages(sf, _GPOS_COMMON, False, masks)
    gm = _global_mask(_GPOS_COMMON, masks)
    for s in st:
        s.items = [(li, gm, a, b, p, t) for li, _m, a, b, p, t in s.items]
    return st


def _simple(sf: ShapeFont, buf, text):
    masks: dict[str, int] = {}
    spec = [(None, [("ccmp", True, True, True, False), ("locl", True, True, True, False)]), (None, list(_GSUB_COMMON))]
    buf = _run_gsub(sf, buf, spec, masks)
    if any(inf.is_mark for inf in buf) and "mark" not in sf.pos_feats:
        marks = "".join(inf.chars for inf in buf if inf.is_mark)
        raise ShapeError(f"「{text}」の記号（{_u(marks)}）を付ける位置の表（GPOS の mark）が、書体 {sf.name} にありません")
    pm: dict[str, int] = {}
    st = _gpos(sf, pm)
    for inf in buf:
        inf.mask = -1
    return buf, st, True


# ------------------------------------------------------------------ アラビア文字
def _arabic(sf: ShapeFont, buf, text):
    types = [joining_type(inf.cp) if inf.cp >= 0 else "U" for inf in buf]
    forms: list[str | None] = [None] * len(buf)
    prev = None
    for i, t in enumerate(types):
        if t == "T":
            continue
        if prev is not None and types[prev] in ("D", "L", "C") and t in ("D", "R", "C"):
            forms[prev] = {"isol": "init", "fina": "medi"}.get(forms[prev] or "", forms[prev])
            forms[i] = "fina"
        else:
            forms[i] = "isol" if t in ("D", "R", "L", "C") else None
        prev = i
    masks: dict[str, int] = {}
    m3 = (False, True, False)  # 全体でない・auto_zwj・auto_zwnj（HarfBuzz の Arabic の形の機能）
    spec = [(None, [("ccmp", True, True, True, False), ("locl", True, True, True, False)])]
    for f in ("isol", "fina", "fin2", "fin3", "medi", "med2", "init"):
        spec.append((None, [(f, *m3, False)]))
    spec.append((None, [("rlig", True, False, True, False)]))
    spec.append((None, [("rclt", True, False, True, False), ("calt", True, False, True, False)]))
    spec.append((None, list(_GSUB_COMMON)))
    _stages(sf, spec, True, masks)  # 目印の割り当て
    local = {}
    for k, f in enumerate(forms):
        if f and f in masks:
            local[k] = masks[f]
    for k, inf in enumerate(buf):
        inf.extra = None
        inf.feat = set()
        inf.pos = k  # 元の並びの番号
    want = {k: f for k, f in enumerate(forms)}
    src_cp = [inf.cp for inf in buf]
    out = _run_gsub(sf, buf, spec, masks, local)
    # 確かめる：決めた形の機能で、字形が置き換わったか（合字になったものは、合字で）
    by_pos = {}
    for inf in out:
        by_pos.setdefault(inf.pos, inf)
    for k, f in want.items():
        if f in (None, "isol") or types[k] == "C":
            continue
        inf = by_pos.get(k)
        if inf is None:
            continue
        if f not in inf.feat and not inf.ligated:
            c = chr(src_cp[k])
            raise ShapeError(f"「{text}」の「{c}」（{_u(c)}）の{_NAME[f]}が、書体 {sf.name} にありません（つながった形で描けません）")
    if any(inf.is_mark for inf in out) and "mark" not in sf.pos_feats:
        marks = "".join(inf.chars for inf in out if inf.is_mark)
        raise ShapeError(f"「{text}」の記号（{_u(marks)}）を付ける位置の表（GPOS の mark）が、書体 {sf.name} にありません")
    pm: dict[str, int] = {}
    st = _gpos(sf, pm)
    for inf in out:
        inf.mask = -1
    return out, st, True


# ------------------------------------------------------------------ インド系の文字
# 位置（HarfBuzz の indic_position_t と同じ順）
(P_START, P_RA, P_PRE_M, P_PRE_C, P_BASE_C, P_AFTER_MAIN, P_ABOVE_C, P_BEFORE_SUB, P_BELOW_C, P_AFTER_SUB,
 P_BEFORE_POST, P_POST_C, P_AFTER_POST, P_FINAL_C, P_SMVD, P_END) = range(16)

_ISC = {
    "Avagraha": "Y", "Bindu": "m", "Brahmi_Joining_Number": "O", "Cantillation_Mark": "A", "Consonant": "C",
    "Consonant_Dead": "C", "Consonant_Final": "K", "Consonant_Head_Letter": "C", "Consonant_Initial_Postfixed": "C",
    "Consonant_Killer": "M", "Consonant_Medial": "K", "Consonant_Placeholder": "O", "Consonant_Preceding_Repha": "F",
    "Consonant_Prefixed": "X", "Consonant_Subjoined": "K", "Consonant_Succeeding_Repha": "K",
    "Consonant_With_Stacker": "S", "Gemination_Mark": "m", "Invisible_Stacker": "H", "Joiner": "j",
    "Modifying_Letter": "X", "Non_Joiner": "z", "Nukta": "N", "Number": "O", "Number_Joiner": "O", "Pure_Killer": "M",
    "Register_Shifter": "r", "Syllable_Modifier": "m", "Tone_Letter": "X", "Tone_Mark": "N", "Virama": "H",
    "Visarga": "m", "Vowel": "V", "Vowel_Dependent": "M", "Vowel_Independent": "V",
}
_IPC = {
    "Left": P_PRE_C, "Top": P_ABOVE_C, "Bottom": P_BELOW_C, "Right": P_POST_C, "Bottom_And_Right": P_POST_C,
    "Left_And_Right": P_POST_C, "Top_And_Bottom": P_BELOW_C, "Top_And_Bottom_And_Left": P_BELOW_C,
    "Top_And_Bottom_And_Right": P_POST_C, "Top_And_Left": P_ABOVE_C, "Top_And_Left_And_Right": P_POST_C,
    "Top_And_Right": P_POST_C, "Overstruck": P_AFTER_MAIN, "Visual_Order_Left": P_PRE_M,
}
_CONSONANT = set("CRSKVOo")  # HarfBuzz の CONSONANT_FLAGS（C・Ra・CS・CM・V・PLACEHOLDER・DOTTEDCIRCLE）


def _matra_pos(cp: int, pos: int, script: str) -> int:
    if pos == P_PRE_C:
        return P_PRE_M
    if pos == P_POST_C:
        return {"Devanagari": P_AFTER_SUB, "Telugu": P_BEFORE_SUB if cp <= 0x0C42 else P_AFTER_SUB,
                "Kannada": P_BEFORE_SUB if (cp < 0x0CC3 or cp > 0x0CD6) else P_AFTER_SUB,
                "Sinhala": P_AFTER_SUB}.get(script, P_AFTER_POST)
    if pos == P_ABOVE_C:
        return {"Devanagari": P_AFTER_SUB, "Gurmukhi": P_AFTER_POST, "Gujarati": P_AFTER_SUB, "Oriya": P_AFTER_MAIN,
                "Tamil": P_AFTER_SUB, "Telugu": P_BEFORE_SUB, "Kannada": P_BEFORE_SUB}.get(script, P_AFTER_SUB)
    if pos == P_BELOW_C:
        return {"Devanagari": P_AFTER_SUB, "Bengali": P_AFTER_SUB, "Gurmukhi": P_AFTER_POST, "Gujarati": P_AFTER_POST,
                "Oriya": P_AFTER_SUB, "Tamil": P_AFTER_POST, "Telugu": P_BEFORE_SUB, "Kannada": P_BEFORE_SUB,
                "Malayalam": P_AFTER_POST}.get(script, P_AFTER_SUB)
    return pos


def indic_props(cp: int, script: str) -> tuple[str, int]:
    """字 → (音節の種類の記号, 位置)。HarfBuzz の Indic の表と同じ決め方（UCD の IndicSyllabicCategory・
    IndicPositionalCategory から）。"""
    if cp == 0x200C:
        return "z", P_END
    if cp == 0x200D:
        return "j", P_END
    if cp == 0x25CC:
        return "o", P_BASE_C
    if cp in (0x00A0, 0x2010, 0x2011, 0x2012, 0x2013, 0x2014, 0x2015, 0x2022, 0x25FB, 0x25FC, 0x25FD, 0x25FE):
        cat = "O"
    else:
        cat = _ISC.get(_range_value("indic_cat", cp) or "", "X")
    if cp in _RA and cat == "C":
        cat = "R"
    if cp in (0x0953, 0x0954):
        cat = "m"
    pos = _IPC.get(_range_value("indic_pos", cp) or "", P_END)
    if cat in _CONSONANT:
        pos = P_BASE_C
    elif cat == "M":
        pos = _matra_pos(cp, pos, script)
    elif cat in ("m", "A", "Y"):
        pos = P_SMVD
    if cp == 0x0B01:
        pos = P_BEFORE_SUB
    return cat, pos


# 音節の形（HarfBuzz の hb-ot-shaper-indic-machine.rl と同じ文法）
_c = "[CR]"
_n = "(?:z?r)?(?:NN?)?"
_z = "[jz]"
_reph = "(?:RH|F)"
_sm = "[mp]"
_cn = f"{_c}j?{_n}"
_matra = f"{_z}*(?:M|{_sm}?P)N?H?"
_tail = f"(?:{_z}?{_sm}{_sm}?z?)?[AD]*"
_halant = f"(?:{_z}?H(?:jN?)?)"
_final_halant = f"(?:{_halant}|Hz)"
_hm = f"(?:{_final_halant}|(?:{_matra})*)"
_cst = f"(?:{_halant}{_cn})*K?{_hm}{_tail}"
_SYL = [
    ("consonant", re.compile(f"[FS]?{_cn}{_cst}")),
    ("vowel", re.compile(f"{_reph}?V{_n}(?:j|{_cst})")),
    ("standalone", re.compile(f"(?:[FS]?O|{_reph}?o){_n}{_cst}")),
    ("symbol", re.compile(f"YN?{_tail}")),
    ("broken", re.compile(f"{_reph}?{_n}{_cst}")),
]


def syllables(cats: str) -> list[tuple[str, int, int]]:
    """音節の切れ目：[(種類, 始まり, 終わり)]。いちばん長く合う形（同じ長さなら先の形）。"""
    out = []
    i = 0
    n = len(cats)
    while i < n:
        best = ("other", i + 1)
        for end in range(min(n, i + 40), i, -1):
            hit = None
            for kind, rx in _SYL:
                if rx.fullmatch(cats, i, end):
                    hit = kind
                    break
            if hit is not None:
                best = (hit, end)
                break
        out.append((best[0], i, best[1]))
        i = best[1]
    return out


def _indic(sf: ShapeFont, buf, text):
    cfg = sf.script.indic
    script = sf.script.name
    eng = sf.engine
    for inf in buf:
        inf.cat, inf.pos = indic_props(inf.cp, script) if inf.cp >= 0 else ("X", P_END)
    cats = "".join(inf.cat for inf in buf)
    syl = syllables(cats)
    for si, (kind, a, b) in enumerate(syl, 1):
        if kind == "broken":
            raise ShapeError(f"「{text}」の「{''.join(x.chars for x in buf[a:b])}」（{_u(''.join(x.chars for x in buf[a:b]))}）"
                             f"は、{_ja(script)}の音節の形になっていません（基の字のない記号など）。組み立てられないので描きません")
        for k in range(a, b):
            buf[k].syl = si
            buf[k].skind = kind
    masks: dict[str, int] = {}
    PS = True  # 音節ごと
    basic = ["nukt", "akhn", "rphf", "rkrf", "pref", "blwf", "abvf", "half", "pstf", "vatu", "cjct"]
    glob = {"nukt", "akhn", "rkrf", "vatu", "cjct"}
    virama = sf.gid(cfg.virama)

    def would(tag, gids):
        return tag in sf.sub_feats and eng.would_substitute(sf.sub_feats[tag], gids)

    def initial(b):
        # 子音の位置（下に付く形・後ろに付く形があるか）を書体で決める
        for inf in b:
            if inf.pos == P_BASE_C and virama is not None and inf.cat in "CR":
                g = inf.gid
                if (would("blwf", [virama, g]) or would("blwf", [g, virama]) or would("vatu", [virama, g])
                        or would("vatu", [g, virama])):
                    inf.pos = P_BELOW_C
                elif would("pstf", [virama, g]) or would("pstf", [g, virama]):
                    inf.pos = P_POST_C
                elif would("pref", [virama, g]) or would("pref", [g, virama]):
                    inf.pos = P_POST_C
        out = []
        k = 0
        while k < len(b):
            s = b[k].syl
            j = k
            while j < len(b) and b[j].syl == s:
                j += 1
            seg = b[k:j]
            if seg and seg[0].skind in ("consonant", "vowel", "standalone"):
                seg = _initial_reorder(sf, seg, masks, cfg, would, text)
            out += seg
            k = j
        return out

    def final(b):
        out = []
        k = 0
        while k < len(b):
            s = b[k].syl
            j = k
            while j < len(b) and b[j].syl == s:
                j += 1
            seg = b[k:j]
            if seg and seg[0].skind in ("consonant", "vowel", "standalone"):
                seg = _final_reorder(sf, seg, masks, cfg, text, out[-1] if out else None)
            out += seg
            k = j
        return out

    spec = [(None, [("locl", True, True, True, PS), ("ccmp", True, True, True, PS)])]
    for i, f in enumerate(basic):
        spec.append((initial if i == 0 else None, [(f, f in glob, False, False, PS)]))
    spec.append((final, [("init", False, False, False, PS), ("pres", True, False, False, PS),
                         ("abvs", True, False, False, PS), ("blws", True, False, False, PS),
                         ("psts", True, False, False, PS), ("haln", True, False, False, PS)] + _GSUB_COMMON))
    _stages(sf, spec, True, masks)
    for f in basic + ["init"]:
        masks.setdefault(f, 1 << len(masks))
    out = _run_gsub(sf, buf, spec, masks)
    # 確かめる：レーフ・半分の形
    _indic_checks(sf, out, text, cfg, masks)
    pm: dict[str, int] = {}
    st = _gpos(sf, pm)
    for inf in out:
        inf.mask = -1
    return out, st, False


def _is_halant(inf) -> bool:
    return inf.cat == "H" and not inf.ligated


def _initial_reorder(sf: ShapeFont, b: list, masks: dict, cfg: IndicConfig, would, text: str) -> list:
    """HarfBuzz の initial_reordering_consonant_syllable（新しい仕様）。"""
    start, end = 0, len(b)
    base = end
    has_reph = False
    limit = start
    chars = "".join(x.chars for x in b)
    # 音節の頭の Ra + ヴィラーマ + 子音は、レーフ（上に付く r）にする文字（デーヴァナーガリー文字など）
    expect_reph = (cfg.reph_mode == "implicit" and cfg.half_check and end - start >= 3 and b[0].cat == "R"
                   and b[1].cat == "H" and b[2].cat not in "jz")
    if "rphf" in sf.sub_feats and start + 3 <= end and (
            (cfg.reph_mode == "implicit" and b[start + 2].cat not in "jz") or
            (cfg.reph_mode == "explicit" and b[start + 2].cat == "j")):
        gl = [b[start].gid, b[start + 1].gid, b[start + 2].gid]
        if would("rphf", gl[:2]) or (cfg.reph_mode == "explicit" and would("rphf", gl)):
            limit += 2
            while limit < end and b[limit].cat in "jz":
                limit += 1
            base = start
            has_reph = True
    elif b[start].cat == "F":
        limit += 1
        while limit < end and b[limit].cat in "jz":
            limit += 1
        base = start
        has_reph = True
    # 基の子音：後ろから
    i = end
    seen_below = False
    while True:
        i -= 1
        if b[i].cat in _CONSONANT:
            if b[i].pos != P_BELOW_C and (b[i].pos != P_POST_C or seen_below):
                base = i
                break
            if b[i].pos == P_BELOW_C:
                seen_below = True
            base = i
        else:
            if start < i and b[i].cat == "j" and b[i - 1].cat == "H":
                break
        if i <= limit:
            break
    if has_reph and base == start and limit - base <= 2:
        has_reph = False  # ほかに子音がない：Ra が基の子音
        expect_reph = False
    if expect_reph and not has_reph:
        raise ShapeError(f"「{text}」の「{chars}」（{_u(chars)}）のレーフ（上に付く r）を、書体 {sf.name} の表（rphf）で"
                         "作れません")
    for k in range(start, base):
        b[k].pos = min(P_PRE_C, b[k].pos)
    if base < end:
        b[base].pos = P_BASE_C
    for k in range(base + 1, end):
        if b[k].cat == "M":
            for j in range(k + 1, end):
                if b[j].cat in _CONSONANT:
                    b[j].pos = P_FINAL_C
                    break
            break
    if has_reph:
        b[start].pos = P_RA
    # 記号を前の字に付ける
    last_pos = P_START
    for k in range(start, end):
        c = b[k].cat
        if c in "jzNrKH":
            b[k].pos = last_pos
            if c == "H" and b[k].pos == P_PRE_M:
                for j in range(k, start, -1):
                    if b[j - 1].pos != P_PRE_M:
                        b[k].pos = b[j - 1].pos
                        break
        elif b[k].pos != P_SMVD:
            if c == "P" and k > start and b[k - 1].cat == "m":
                b[k - 1].pos = b[k].pos
            last_pos = b[k].pos
    last = base
    for k in range(base + 1, end):
        if b[k].cat in _CONSONANT:
            for j in range(last + 1, k):
                if b[j].pos < P_SMVD:
                    b[j].pos = b[k].pos
            last = k
        elif b[k].cat in "MP":
            last = k
    # 並べ替え（位置の順、同じなら元の順）
    order = sorted(range(start, end), key=lambda k: (b[k].pos, k))
    b = [b[k] for k in order]
    first_left = last_left = end
    base = end
    for k in range(start, end):
        if b[k].pos == P_BASE_C:
            base = k
            break
        if b[k].pos == P_PRE_M:
            if first_left == end:
                first_left = k
            last_left = k
    if first_left < last_left:
        b[first_left:last_left + 1] = b[first_left:last_left + 1][::-1]
        i = first_left
        for j in range(first_left, last_left + 1):
            if b[j].cat in "MP":
                b[i:j + 1] = b[i:j + 1][::-1]
                i = j + 1
    # 目印
    for k in range(start, end):
        if b[k].pos != P_RA:
            break
        b[k].mask |= masks["rphf"]
    pre = masks["half"] | (0 if cfg.blwf_post_only else masks["blwf"])
    for k in range(start, base):
        b[k].mask |= pre
    post = masks["blwf"] | masks["abvf"] | masks["pstf"]
    for k in range(base + 1, end):
        b[k].mask |= post
    if "pref" in sf.sub_feats and base + 2 < end:
        for k in range(base + 1, end - 1):
            if would("pref", [b[k].gid, b[k + 1].gid]):
                b[k].mask |= masks["pref"]
                b[k + 1].mask |= masks["pref"]
                break
    for k in range(base + 1, end):
        if b[k].cat in "jz":
            nonj = b[k].cat == "z"
            j = k
            while True:
                j -= 1
                if nonj:
                    b[j].mask &= ~masks["half"]
                if not (j > start and b[j].cat not in _CONSONANT):
                    break
    for inf in b:
        inf.feat = set()
    return b


def _final_reorder(sf: ShapeFont, b: list, masks: dict, cfg: IndicConfig, text: str, prev) -> list:
    """HarfBuzz の final_reordering_syllable_indic。"""
    script = sf.script.name
    start, end = 0, len(b)
    virama = sf.gid(cfg.virama)
    for inf in b:
        if virama is not None and inf.gid == virama and inf.ligated and inf.multiplied:
            inf.cat = "H"
            inf.flags &= ~(otl.LIGATED | otl.MULTIPLIED)
    try_pref = "pref" in sf.sub_feats
    base = start
    while base < end:
        if b[base].pos >= P_BASE_C:
            if try_pref and base + 1 < end:
                for k in range(base + 1, end):
                    if b[k].mask & masks["pref"]:
                        if not (b[k].flags & otl.SUBSTITUTED and b[k].ligated and not b[k].multiplied):
                            base = k
                            while base < end and _is_halant(b[base]):
                                base += 1
                            if base < end:
                                b[base].pos = P_BASE_C
                            try_pref = False
                        break
                if base == end:
                    break
            if start < base and b[base].pos > P_BASE_C:
                base -= 1
            break
        base += 1
    if base == end and start < base and b[base - 1].cat == "j" and not b[base - 1].ligated:
        base -= 1
    if base < end:
        while start < base and b[base].cat in "NH" and not b[base].ligated:
            base -= 1

    def one_of(inf, cats):
        return inf.cat in cats and not inf.ligated

    # 前に置く母音記号
    if start + 1 < end and start < base:
        new_pos = base - 2 if base == end else base - 1
        if script not in ("Malayalam", "Tamil"):
            while True:
                while new_pos > start and not one_of(b[new_pos], "MPH"):
                    new_pos -= 1
                if _is_halant(b[new_pos]) and b[new_pos].pos != P_PRE_M:
                    if new_pos + 1 < end and b[new_pos + 1].cat == "j" and new_pos > start:
                        new_pos -= 1
                        continue
                else:
                    new_pos = start
                break
        if start < new_pos and b[new_pos].pos != P_PRE_M:
            i = new_pos
            while i > start:
                if b[i - 1].pos == P_PRE_M:
                    old = i - 1
                    if old < base <= new_pos:
                        base -= 1
                    tmp = b.pop(old)
                    b.insert(new_pos, tmp)
                    new_pos -= 1
                i -= 1
    # レーフ
    if start + 1 < end and b[start].pos == P_RA and ((b[start].cat == "F") ^ (b[start].ligated and not b[start].multiplied)):
        rp = cfg.reph_pos
        new_reph = None
        if rp != "after_post":
            k = start + 1
            while k < base and not _is_halant(b[k]):
                k += 1
            if k < base and _is_halant(b[k]):
                if k + 1 < base and b[k + 1].cat in "jz":
                    k += 1
                new_reph = k
            if new_reph is None and rp == "after_main":
                k = base
                while k + 1 < end and b[k + 1].pos <= P_AFTER_MAIN:
                    k += 1
                if k < end:
                    new_reph = k
            if new_reph is None and rp == "after_sub":
                k = base
                while k + 1 < end and b[k + 1].pos not in (P_POST_C, P_AFTER_POST, P_SMVD):
                    k += 1
                if k < end:
                    new_reph = k
        if new_reph is None:
            k = start + 1
            while k < base and not _is_halant(b[k]):
                k += 1
            if k < base and _is_halant(b[k]):
                if k + 1 < base and b[k + 1].cat in "jz":
                    k += 1
                new_reph = k
        if new_reph is None:
            k = end - 1
            while k > start and b[k].pos == P_SMVD:
                k -= 1
            if _is_halant(b[k]):
                for i in range(base + 1, k):
                    if b[i].cat in "MP":
                        k -= 1
            new_reph = k
        reph = b.pop(start)
        b.insert(new_reph, reph)
        if start < base <= new_reph:
            base -= 1
    # 前に置く子音（pref）
    if try_pref and base + 1 < end:
        for i in range(base + 1, end):
            if b[i].mask & masks["pref"]:
                if b[i].ligated and not b[i].multiplied:
                    new_pos = base
                    if script not in ("Malayalam", "Tamil"):
                        while new_pos > start and not one_of(b[new_pos - 1], "MPH"):
                            new_pos -= 1
                    if new_pos > start and _is_halant(b[new_pos - 1]):
                        if new_pos < end and b[new_pos].cat in "jz":
                            new_pos += 1
                    tmp = b.pop(i)
                    b.insert(new_pos, tmp)
                    if new_pos <= base < i:
                        base += 1
                break
    # 語頭（前が文字・記号でない）の、前に置く母音記号に init
    if b[start].pos == P_PRE_M:
        if prev is None or unicodedata.category(chr(prev.cp)) not in (
                "Cf", "Cn", "Co", "Cs", "Ll", "Lm", "Lo", "Lt", "Lu", "Mc", "Me", "Mn"):
            b[start].mask |= masks["init"]
    return b


def _indic_checks(sf: ShapeFont, buf, text: str, cfg: IndicConfig, masks) -> None:
    virama = sf.gid(cfg.virama)
    k = 0
    while k < len(buf):
        s = buf[k].syl
        j = k
        while j < len(buf) and buf[j].syl == s:
            j += 1
        seg = buf[k:j]
        chars = "".join(x.chars for x in seg)
        if seg and seg[0].skind in ("consonant", "vowel"):
            if any(x.pos == P_RA and not x.ligated for x in seg):
                raise ShapeError(f"「{text}」の「{chars}」（{_u(chars)}）のレーフ（上に付く r）が、書体 {sf.name} で"
                                 "作れませんでした")
            if cfg.half_check and virama is not None:
                for i, x in enumerate(seg):
                    if x.gid == virama and i < len(seg) - 1:
                        rest = seg[i + 1:]
                        if any(r.cat in _CONSONANT for r in rest) and not any(r.cp == 0x200C for r in rest[:1]):
                            raise ShapeError(f"「{text}」の「{chars}」（{_u(chars)}）は、書体 {sf.name} で子音の組み合わせ"
                                             "（半分の形・結合の形）にならず、ヴィラーマ（ハラント）が見えたままになります")
        k = j


# ====================================================================== 折り返しの単位
def units(text: str) -> list[str]:
    """折り返してよい単位（字の組み立てで切り離せない並び）。組み立てる文字のない字の並びは 1 字ずつ。"""
    if not needs(text):
        return list(text)
    out: list[str] = []
    n = len(text)
    i = 0
    while i < n:
        s = script_of(text[i])
        info = SCRIPTS.get(s)
        if info is not None and info.kind == "indic":
            j = i
            while j < n and script_of(text[j]) in (s, "Inherited") or (j < n and text[j] in "\u200c\u200d"):
                j += 1
            cats = "".join(indic_props(ord(c), s)[0] for c in text[i:j])
            for _k, a, b in syllables(cats):
                out.append(text[i + a:i + b])
            i = j
            continue
        j = i + 1
        while j < n and (unicodedata.category(text[j]) in ("Mn", "Mc", "Me") or text[j] in "\u200c\u200d"):
            j += 1
        out.append(text[i:j])
        i = j
    return out


def verify(text: str, drawn: list[tuple[str, int]]) -> str | None:
    """見た目のストリームに描いた、組み立てた書体の字形の並び（書体の名前, 字形）が、text をいま組み立てたものと同じか。
    違えば理由。"""
    try:
        ln = line(text)
    except GlyphError as e:
        return f"「{text}」を組み立てられません：{e}"
    want = []
    for sg in ln.segs:
        if sg.shaped:
            want += [(sg.font.name, g[0]) for g in sg.glyphs]
    if want != drawn:
        return f"「{text}」の描いた字形の並びが、組み立てた字形と違います"
    return None
