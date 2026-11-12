"""文字コードの印のない 8 ビットの字（書庫の中のファイルの名前・メールの字・OLE の表示名）の読み方を、規則で決める。

文字のファイル（datafile/textenc）と同じ決め方（どの段も、規則として書けることだけを使う。「日本語のはず」と置かない）：

1. ASCII の字だけなら ASCII。
2. UTF-8 として厳密に読める名前は UTF-8（UTF-8 は、先頭のバイトと続くバイトの形が決まっている）。
3. 残りは、**バイトからは文字コードが決まらない**。まとめて（1 つの書庫の名前の全部・1 通のメールの字の全部）、読める文字コードを
   すべて試す：多バイト（Shift_JIS・EUC-JP・Big5・GBK・UHC）、Windows などの 1 バイト、DOS の OEM のコードページ（ZIP の決まりの
   CP437 と、Windows の ZIP・LZH が名前に使う地域ごとのコードページ：CP850・CP866 など）。厳密に、きれいに（割り当てのない字・
   C1 の制御文字なしで）読めた読み方が 1 つだけなら、それ。読んだ字の違うものが 2 つ以上なら、決まった順（多バイト → 1 バイト → OEM）の
   いちばん前のもので字を示し、ほかを候補として示す。どれが正しいかは、エージェントが字を読んで決める（encoding で渡すか、
   正しい読み方の名前でそのまま指す）。順は決め方ではない（決められないときに、どれで示すかだけ）。
4. どれでも読めなければ、形式の決まりの文字コード（ZIP は CP437）か Latin-1 で、バイトを 1 字ずつ字にして、そう伝える。

同じファイルの中に文字コードの宣言があるとき（declared。メールのコードページ・複合ファイルの要約の情報のコードページ）は、
ASCII でない字の全部がその文字コードで厳密にきれいに読めれば、それ（UTF-8 の形より先）。
エージェントが選んだ文字コード（encoding）は、印のない字の全部がその文字コードで厳密にきれいに読めるときだけ使う（読めなければ NameCodeError）。

表（dBASE・Lotus 1-2-3・Excel 2〜4）の 8 ビットの字は、1 つのファイルの字の全部が同じ文字コード（decide_whole）：字ごとには決めない
（UTF-8 も、全部の字が UTF-8 の形のときだけ）。項目の幅いっぱいの字（dBASE）は、終わりの字が幅で切れていることがあるので、
終わりの、途中までの字だけは許す。
"""

from __future__ import annotations

import codecs
from contextvars import ContextVar
from dataclasses import dataclass, field

from .datafile.textenc import ISO2022_MS, MULTI, NAME_JA, SINGLE, canonical, clean


class NameCodeError(ValueError):
    pass


def _known(names: tuple) -> tuple:
    out = []
    for n in names:
        try:
            codecs.lookup(n)
        except LookupError:
            continue
        out.append(n)
    return tuple(out)


# DOS・Windows の OEM のコードページ（Python が表を持つもの）。CP437 は ZIP の決まり（APPNOTE 4.4.4）の文字コード
OEM = _known(("cp437", "cp850", "cp858", "cp852", "cp855", "cp857", "cp860", "cp861", "cp862", "cp863", "cp864", "cp865",
              "cp866", "cp869", "cp737", "cp775", "cp720", "cp1125"))
ORDER = MULTI + tuple(e for e in SINGLE if e not in MULTI) + tuple(e for e in OEM if e not in SINGLE)
_ONE_BYTE = frozenset(e for e in ORDER if e not in MULTI)  # 1 バイトが 1 字の文字コード（前後のバイトで読み方が変わらない）

# Windows のコードページの番号 → Python の codec（番号がそのまま「cp番号」でないもの）
_CODEPAGES = {65001: "utf-8", 65000: "utf-7", 1200: "utf-16-le", 1201: "utf-16-be", 12000: "utf-32-le", 12001: "utf-32-be",
              20127: "ascii", 932: "cp932", 936: "gbk", 949: "cp949", 950: "cp950", 1361: "johab", 54936: "gb18030",
              20936: "gb2312", 52936: "hz", 50220: ISO2022_MS, 50221: ISO2022_MS, 50222: ISO2022_MS,
              50225: "iso2022_kr", 51932: "euc_jp", 20932: "euc_jp", 51949: "euc_kr", 51936: "gb2312", 20866: "koi8_r",
              21866: "koi8_u", 10000: "mac_roman", 10001: "shift_jis", 10006: "mac_greek", 10007: "mac_cyrillic",
              10029: "mac_latin2", 10079: "mac_iceland", 10081: "mac_turkish",
              **{28590 + n: f"iso8859_{n}" for n in (1, 2, 3, 4, 5, 6, 7, 8, 9, 13, 15)}, 28603: "iso8859_13",
              28605: "iso8859_15", 28606: "iso8859_16"}


def codec_of_codepage(cp: int) -> str | None:
    """Windows のコードページの番号 → Python の codec の名前。表を持たない番号は None（ほかの番号の表で読まない）。"""
    name = _CODEPAGES.get(cp, f"cp{cp}")
    try:
        codecs.lookup(name)
    except LookupError:
        return None
    return name


# GDI の字体の文字セット（LOGFONT の lfCharSet。MS-WMF 2.1.1.5）→ 文字コード。1（DEFAULT）・2（SYMBOL）は、文字コードを決めない
CHARSETS = {0: "cp1252", 77: "mac_roman", 128: "cp932", 129: "cp949", 130: "johab", 134: "gbk", 136: "cp950", 161: "cp1253",
            162: "cp1254", 163: "cp1258", 177: "cp1255", 178: "cp1256", 186: "cp1257", 204: "cp1251", 222: "cp874",
            238: "cp1250", 255: "cp437"}


def read_ansi(raw: bytes, charset: int | None = None) -> tuple[str, str]:
    """絵（WMF・EMF）の 8 ビットの字：(字, 決められないときの読み方の説明。決まれば "")。字体の文字セット charset の文字コードで
    厳密に読めれば、それ。文字セットが文字コードを決めない（DEFAULT・字体なし）か、読めなければ、decide の決め方。
    SYMBOL の字体の字は、字形の番号なので、バイトのまま（Latin-1）。"""
    if raw.isascii():
        return raw.decode("ascii"), ""
    if charset == 2:
        return raw.decode("latin-1"), ""
    enc = CHARSETS.get(charset)
    if enc is not None:
        try:
            return raw.decode(enc), ""
        except UnicodeDecodeError:
            pass
    reading = decide([raw])
    text, how = reading.text(raw)
    return text, (how if reading.undetermined else "")


def label(enc: str) -> str:
    return NAME_JA.get(enc, enc.upper() if enc.startswith(("cp", "iso", "koi")) else enc)


def _strict(raw: bytes, enc: str) -> str | None:
    """厳密に、きれいに読めた字（割り当てのない字・C1 の制御文字があれば None）。C0 の制御文字（0x00〜0x1F）は見ない：
    ここで試す文字コードのどれでも同じ字で、多バイトの字の中にも出ないので、文字コードを決める手がかりにならない。"""
    try:
        t = raw.decode(enc)
    except (UnicodeDecodeError, LookupError):
        return None
    return t if clean(t) else None


def _utf8(raw: bytes) -> str | None:
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return None


def _clip(s: str, n: int = 40) -> str:
    return s if len(s) <= n else s[:n] + "…"


@dataclass
class Reading:
    """ひとまとまりの、印のない字の読み方。encoding・basis は、ASCII でも UTF-8 でもない字のもの（なければ空）。"""

    encoding: str = ""
    basis: str = ""  # one（読める文字コードが 1 つだけ）・order（決められない：順のいちばん前）・given（エージェントが選んだ）・
    #                  declared（同じファイルの宣言）・none（どれでも読めない：fallback で 1 字ずつ）
    candidates: list = field(default_factory=list)  # [{"encoding", "label", "sample"}]（読んだ字の違う、ほかの読み方）
    own_sample: str = ""  # 候補の sample と同じ字を、encoding で読んだもの
    count: int = 0  # encoding で読んだ字の数（名前の数）
    declared_by: str = ""
    declared_note: str = ""  # ファイルの宣言を使わなかったとき、その事実（表の字：decide_whole）

    @property
    def undetermined(self) -> bool:
        return self.basis in ("order", "none")

    def text(self, raw: bytes) -> tuple[str, str]:
        """(字, 読み方の説明)。"""
        if raw.isascii():
            return raw.decode("ascii"), "ASCII"
        if self.basis == "given":
            return raw.decode(self.encoding), f"{label(self.encoding)}（エージェントが選んだ）"
        if self.basis == "declared":
            return raw.decode(self.encoding), f"{label(self.encoding)}（{self.declared_by}）"
        u = _utf8(raw)
        if u is not None:
            return u, "UTF-8（バイトの形）"
        if self.basis == "none":
            return raw.decode(self.encoding, "replace"), f"{label(self.encoding)}（どの文字コードでも、きれいに読めません）"
        how = {"one": "ほかの文字コードでは読めない", "order": "決められない"}[self.basis]
        return raw.decode(self.encoding), f"{label(self.encoding)}（{how}）"

    def info(self) -> dict:
        """窓の JSON に入れる、読み方の決め方（ASCII・UTF-8 のほかの字があるときだけ）。"""
        if not self.basis:
            return {}
        out = {"encoding": self.encoding, "label": label(self.encoding), "basis": self.basis, "count": self.count}
        if self.basis == "given":
            out["chosen_by"] = "agent"
        if self.declared_note:
            out["declared"] = self.declared_note
        if self.candidates:
            out["sample"] = self.own_sample
            out["candidates"] = list(self.candidates)
        return out

    def note_lines(self, what: str = "名前", how_to_choose: str = "") -> list[str]:
        """Markdown の窓の注意（決められないとき・エージェントが選んだとき）。"""
        if self.basis == "given":
            return [f"- {what}の文字コードは、エージェントが選んで指定した {label(self.encoding)} で読みました"
                    f"（{self.count} 個。厳密に読めることは確かめました）。{self.declared_note}"]
        if self.basis == "none":
            return [f"- **{what}の文字コードを決められませんでした**：{self.count} 個の{what}は、文字コードがファイルに書かれていず、"
                    f"UTF-8・多バイト・1 バイトのどの文字コードでも、まとめてきれいには読めません。{label(self.encoding)} で"
                    "バイトを 1 字ずつ字にしています（字が正しいとは限りません）。"]
        if self.basis != "order":
            return []
        out = [f"- **{what}の文字コードを規則では決められません**：{self.count} 個の{what}は、文字コードがファイルに書かれていません。"
               f"{self.declared_note}{label(self.encoding)} で読みましたが、次の文字コードでも、きれいに読めます（字の違う最初の{what}を並べます）。"
               f"意味の通る方を選んでください。{how_to_choose}",
               f"  - {self.encoding}（いまの読み方）：{self.own_sample}"]
        out += [f"  - {c['encoding']}：{c['sample']}" for c in self.candidates]
        return out


def decide(raws, encoding: str | None = None, declared=(), fallback: str = "latin-1") -> Reading:
    """raws（印のないバイトの字の並び）の読み方。declared：[(文字コード, どこの宣言か)]（前から順に、読めた最初のもの）。"""
    todo = list(dict.fromkeys(r for r in raws if not r.isascii()))
    if encoding:
        enc = canonical(encoding)
        if enc is None:
            raise NameCodeError(f"文字コード「{encoding}」を知りません（例：UTF-8・Shift_JIS・EUC-KR・GBK・Big5・CP850・CP866）")
        if not todo:
            raise NameCodeError("この書類には、文字コードを選べる字がありません（どの字も ASCII か、文字コードが書かれています）")
        for r in todo:
            if _strict(r, enc) is None:
                raise NameCodeError(f"指定された文字コード（{label(enc)}）では、文字コードの印のない字を厳密に読めません"
                                    f"（{r[:24]!r}）。窓に並んだ候補から選んでください")
        return Reading(enc, "given", count=len(todo))
    if todo:
        for enc, where in declared:
            if enc and _strict(b"\n".join(todo), enc) is not None:
                return Reading(enc, "declared", count=len(todo), declared_by=where)
    rest = [r for r in todo if _utf8(r) is None]
    if not rest:
        return Reading()
    joined = b"\n".join(rest)
    got: list[tuple[str, str]] = []
    for enc in ORDER:
        t = _strict(joined, enc)
        if t is not None and all(t != u for _e, u in got):
            got.append((enc, t))
    if not got:
        return Reading(fallback, "none", count=len(rest))
    enc, text = got[0]
    if len(got) == 1:
        return Reading(enc, "one", count=len(rest))
    own = text.split("\n")
    cands, k0 = [], None
    for e2, t2 in got[1:]:
        other = t2.split("\n")
        k = next((i for i, (a, b) in enumerate(zip(own, other)) if a != b), 0)
        k0 = k if k0 is None else k0
        # 見本は、字の違う最初のもの。いまの読み方の見本（own_sample）と別のものなら、いまの読み方も添える
        cands.append({"encoding": e2, "label": label(e2),
                      "sample": _clip(other[k]) + ("" if k == k0 else f"（いまの読み方では {_clip(own[k])}）")})
    return Reading(enc, "order", cands, _clip(own[k0]), len(rest))


# ====================================================================== 表の字：1 つのファイルに 1 つの文字コード
_TABLE: ContextVar[tuple | None] = ContextVar("todoku_table_encoding", default=None)  # (文字コード, そのファイルの印の集まり)


def _blob_key(blob) -> tuple:
    b = blob if isinstance(blob, bytes) else bytes(blob)
    return len(b), hash(b)


def use_table_encoding(encoding: str | None, blob=b"") -> None:
    """表（dBASE・Lotus 1-2-3・Excel 2〜4）のファイル blob の字の、エージェントが選んだ文字コード（None：選んでいない）。
    選択は、そのファイルの中身にだけ当てはめる（ほかの表・文書に埋め込んだ表は、規則で読む）。同じ文字コードで続けて呼ぶと、
    当てはめるファイルを足す（書いた後のファイルを、書く前と同じ文字コードで読み戻すため）。"""
    if not encoding:
        _TABLE.set(None)
        return
    cur = _TABLE.get()
    keys = cur[1] if cur and cur[0] == encoding else frozenset()
    _TABLE.set((encoding, keys | {_blob_key(blob)}))


def table_encoding(blob=b"") -> str | None:
    cur = _TABLE.get()
    return cur[0] if cur and _blob_key(blob) in cur[1] else None


def _strict_cut(raw: bytes, enc: str) -> str | None:
    """幅で切れているかもしれない字：終わりの、途中までの字（多バイトの字の前のバイトだけ）は読まずに、残りを厳密に読む。"""
    try:
        t = codecs.getincrementaldecoder(enc)("strict").decode(raw, False)
    except (UnicodeDecodeError, LookupError):
        return None
    return t if clean(t) else None


def decide_whole(raws, cut=(), encoding: str | None = None, declared=(), fallback: str = "latin-1",
                 empty_ok: bool = False) -> Reading:
    """1 つのファイルの 8 ビットの字の全部（raws と cut）を、同じ 1 つの文字コードで読む読み方。encoding は、いつも入る
    （ASCII だけなら "ascii"、basis は空）。cut：項目の幅いっぱいの字（終わりの字が切れていることがある）。
    declared：[(文字コード, どこの宣言か)]。前から順に、全部を厳密にきれいに読めた最初のもの。なければ、全部が UTF-8 の形なら
    UTF-8（basis は shape）。ほかは decide と同じ（決まった順に試し、読めた最初の読み方で示し、ほかを候補に並べる）。
    empty_ok：ASCII でない字が 1 つもなくても、encoding を受け取る（書ける形式：書く値の文字コードを、エージェントが選ぶ）。"""
    whole = list(dict.fromkeys(r for r in raws if not r.isascii()))
    seen = set(whole)
    cuts = [r for r in dict.fromkeys(r for r in cut if not r.isascii()) if r not in seen]
    count = len(whole) + len(cuts)
    joined, cut_joined = b"\n".join(whole), b"\n".join(cuts)

    kinds: bytes | None = None  # 字のバイトの種類

    def read(enc: str) -> str | None:
        nonlocal kinds
        if enc in _ONE_BYTE:  # 1 バイトが 1 字の文字コード：読めるか・きれいかは、バイトの種類（256 までの字）だけで決まる
            if kinds is None:
                kinds = bytes(sorted(set(joined) | set(cut_joined)))
            if _strict(kinds, enc) is None:
                return None
            a = joined.decode(enc)
            return (f"{a}\n{cut_joined.decode(enc)}" if whole else cut_joined.decode(enc)) if cuts else a
        a = _strict(joined, enc) if whole else ""
        if a is None or not cuts:
            return a
        b = _strict(cut_joined, enc)  # どの字も切れていなければ、まとめて読める
        if b is None:
            parts = []
            for r in cuts:
                t = _strict_cut(r, enc)
                if t is None:
                    return None
                parts.append(t)
            b = "\n".join(parts)
        return f"{a}\n{b}" if whole else b

    said = [(enc, where) for enc, where in declared if enc]
    if encoding:
        enc = canonical(encoding)
        if enc is None:
            raise NameCodeError(f"文字コード「{encoding}」を知りません（例：UTF-8・Shift_JIS・EUC-KR・GBK・Big5・CP850・CP866）")
        if not count and not empty_ok:
            raise NameCodeError("この書類には、文字コードを選べる字がありません（どの字も ASCII です）")
        if count and read(enc) is None:
            raise NameCodeError(f"指定された文字コード（{label(enc)}）では、この書類の字を厳密に読めません。窓に並んだ候補から選んでください")
        note = "".join(f"ファイルに書かれた文字コードは {label(e)}（{w}）です。" for e, w in said[:1] if e != enc)
        return Reading(enc, "given", count=count, declared_note=note)
    if not count:  # ASCII の字だけ：宣言があれば、書くときの文字コードは宣言のもの
        return Reading(said[0][0], "declared", declared_by=said[0][1]) if said else Reading("ascii")
    for enc, where in said:
        if read(enc) is not None:
            return Reading(enc, "declared", count=count, declared_by=where)
    note = "".join(f"ファイルに書かれた文字コード {label(e)}（{w}）では、字を厳密に読めません。" for e, w in said)
    if read("utf-8") is not None:
        return Reading("utf-8", "shape", count=count, declared_note=note)
    got: list[tuple[str, str]] = []
    for enc in ORDER:
        t = read(enc)
        if t is not None and all(t != u for _e, u in got):
            got.append((enc, t))
    if not got:
        return Reading(fallback, "none", count=count, declared_note=note)
    enc, text = got[0]
    if len(got) == 1:
        return Reading(enc, "one", count=count, declared_note=note)
    own = text.split("\n")
    cands, k0 = [], None
    for e2, t2 in got[1:]:
        other = t2.split("\n")
        k = next((i for i, (a, b) in enumerate(zip(own, other)) if a != b), 0)
        k0 = k if k0 is None else k0
        cands.append({"encoding": e2, "label": label(e2),
                      "sample": _clip(other[k]) + ("" if k == k0 else f"（いまの読み方では {_clip(own[k])}）")})
    return Reading(enc, "order", cands, _clip(own[k0]), count, declared_note=note)


def table_reading(blob, raws, cut=(), declared=(), empty_ok: bool = False) -> Reading:
    """表のファイル blob の字の読み方：エージェントがこのファイルに選んだ文字コード（use_table_encoding）があれば、それで厳密に
    読めるときだけ使う。"""
    return decide_whole(raws, cut, table_encoding(blob), declared, empty_ok=empty_ok)


def table_window(doc: dict, reading: "Reading | None") -> None:
    """窓の JSON に、表の字の文字コードの決め方を入れる（規則で決まらなかったとき・エージェントが選んだとき・宣言を使わなかったとき）。"""
    if reading is None or not (reading.basis in ("order", "none", "given") or reading.declared_note):
        return
    doc["text_encoding"] = reading.info()
    if reading.basis == "given" and not reading.count:
        doc["text_encoding"]["notes"] = [f"- このファイルに ASCII でない字はありません。書く値は、エージェントが選んで指定した文字コード"
                                         f"（{label(reading.encoding)}）のバイトにします。{reading.declared_note}"]
        return
    doc["text_encoding"]["notes"] = reading.note_lines(
        "値", "選んだ文字コードの名前を、このファイルのどの道具にも encoding で渡します。") or (
        [f"- {reading.declared_note}{label(reading.encoding)} で読みました"
         f"（{'字の全部が UTF-8 の形' if reading.basis == 'shape' else 'ほかの文字コードでは読めない'}）。"])


def readings(raw: bytes) -> list[str]:
    """raw の、きれいに読める読み方の全部（字の違うもの）。正しい読み方の名前で指されたときに、項目を見つけるため。"""
    out: list[str] = []
    for enc in ("utf-8",) + ORDER:
        t = _strict(raw, enc)
        if t is not None and t not in out:
            out.append(t)
    return out
