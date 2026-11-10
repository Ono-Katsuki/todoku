"""Word（.docx）の文書を読む：本文・表・ヘッダー/フッター・脚注・コメント・テキストボックス・段落番号・
旧来のフォームの欄（FORMTEXT など）・コンテンツコントロール・文書の保護。

段落の文字は、どの要素（w:t など）のどのバイト位置から来たかを覚えておき、書き込みを小さな差分にする。
"""

from __future__ import annotations

import contextlib
import contextvars
import re
from dataclasses import dataclass, field

from ..zipraw import RawZip
from .xmlpos import Node, parse

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"

WINGDINGS = {0xF06F: "□", 0xF0A8: "□", 0xF071: "❑", 0xF072: "❒", 0xF06E: "■", 0xF0FE: "☑", 0xF0FD: "☒", 0xF078: "☒",
             0xF0FC: "✓", 0xF0FB: "✗", 0xF06C: "●", 0xF0A1: "○", 0xF09F: "•", 0xF0A7: "▪", 0xF0D8: "➢", 0xF0E8: "➔",
             0xF0B2: "◆", 0xF075: "◆", 0xF0B7: "•"}
SYMBOL_FONT = {0xF0B7: "•", 0xF0A7: "§", 0xF0D7: "×", 0xF0B4: "×", 0xF0AE: "→", 0xF0AC: "←", 0xF0B0: "°", 0xF0B1: "±"}


# 箇条書きの印のうち、チェックの欄（□）とみなすもの。チェックした印は、同じフォントで ■（☐ は ☑。Wingdings は ☑ の字）
BULLET_BOXES = frozenset("□☐❑❒")


class DocxError(ValueError):
    pass


@dataclass
class Seg:
    """段落の文字の1片。node は書き換える要素（w:t・w:sym など）。"""

    text: str
    kind: str  # t / tab / br / sym / field / note / other
    node: Node | None = None
    run: Node | None = None
    hidden: bool = False
    underline: bool = False
    field: "Control | None" = None
    look: tuple | None = None  # 字の書式の事実（取り消し線か, 字の色 "RRGGBB" か None）：docx/fontfacts.py。ない形式は None


def _tab_stops(ppr) -> list[tuple[str, str, str]]:
    """w:pPr の w:tabs：[(位置, 種類, リーダー)]。"""
    tabs = ppr.find("tabs") if ppr is not None else None
    if tabs is None:
        return []
    return [(t.get("pos") or "", t.get("val") or "", t.get("leader") or "") for t in tabs.findall("tab")]


def _lay_stops(stops: dict, more: list) -> None:
    """タブ位置を重ねる：同じ位置は後のものが勝ち、clear はその位置を消す。"""
    for pos, val, leader in more:
        if val == "clear":
            stops.pop(pos, None)
        else:
            stops[pos] = leader


@dataclass
class Para:
    node: Node
    part: str
    segs: list[Seg]
    style: str = ""
    label: str = ""  # 段落番号（「１．」「(1)」「①」など）
    cell: "Cell | None" = None
    boxes: list["Block"] = field(default_factory=list)  # 段落の中のテキストボックスの中身
    leaders: tuple = ()  # 段落のタブ位置のうち、線（リーダー）を引くものの、線の種類（タブのある段落だけ）

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.segs if not s.hidden)

    def raw_text(self) -> str:
        """XML の文字の要素から来た字だけ（非表示の字も含む。窓の読み落としの検査に使う）。"""
        return "".join(s.text for s in self.segs if s.kind in ("t", "sym", "math", "ruby", "del"))

    def deleted_text(self) -> str:
        """変更履歴で削除された字（w:del の w:delText、RTF の \\deleted）。窓では本文に入れず、印をつけて示す。"""
        return "".join(s.text for s in self.segs if s.kind == "del")

    def hidden_text(self) -> str:
        return "".join(s.text for s in self.segs if s.hidden and s.kind in ("t", "sym"))

    def positions(self) -> list[tuple[int, int]]:
        """text の1字ごとに、(seg の番号, seg の中の位置)。"""
        out = []
        for i, s in enumerate(self.segs):
            if s.hidden:
                continue
            for k in range(len(s.text)):
                out.append((i, k))
        return out


@dataclass
class Cell:
    node: Node
    row: int
    col: int
    span: int = 1
    vmerge: str = ""  # "restart" / "continue" / ""
    blocks: list = field(default_factory=list)
    table: "Table | None" = None
    shade: str = ""

    def paras(self) -> list[Para]:
        out = []
        for b in self.blocks:
            if isinstance(b, Para):
                out.append(b)
            elif isinstance(b, Table):
                for row in b.rows:
                    for c in row:
                        out.extend(c.paras())
        return out

    def text(self) -> str:
        return " ".join(p.text.strip() for p in self.paras() if p.text.strip())


@dataclass
class Table:
    node: Node
    part: str
    rows: list[list[Cell]] = field(default_factory=list)


@dataclass
class Control:
    """入力欄：旧来のフォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN）またはコンテンツコントロール。"""

    kind: str  # text / checkbox / dropdown / date / combo / richtext / picture
    source: str  # formfield / sdt
    name: str
    node: Node  # fldChar（begin）または w:sdt
    part: str
    para: Para | None = None
    value: object = None
    options: list[str] = field(default_factory=list)
    checked: bool = False
    placeholder: bool = False
    maxlen: int | None = None
    date_format: str = ""
    date_lid: str = ""  # 日付の言語（コンテンツコントロールの w:date/w:lid。旧来の欄は、欄の字の言語：_run_lang）
    date_lid_from: str = ""  # 旧来の欄の言語が決まった所（字の要素・スタイル・文書の既定・themeFontLang）
    date_calendar: str = ""  # 日付の暦（w:date/w:calendar：gregorian・japan・taiwan・thai・korea・hijri・hebrew・saka など）
    result_segs: list[Seg] = field(default_factory=list)
    end: Node | None = None
    lock: str = ""
    help: str = ""
    ax: object = None  # ActiveX の部品（source が activex のとき、axfile.AxRef）
    multiline: bool = True  # 書式なしテキストのコントロール（w:text）：w:multiLine（既定は「いいえ」。ECMA-376 第1部 17.5.2.44）
    # 旧来のフォームの欄の設定（.docx の w:ffData・.doc の FFData・RTF の \formfield）
    text_type: str = ""  # 字の欄の種類（TEXT_TYPES のどれか。字の欄でなければ空）
    text_format: str = ""  # 字の欄の書式（w:format。数の「#,##0」・字の「UPPERCASE」など。Word が欄を出るときに当てる）
    expression: str = ""  # 計算の欄の式（w:default）
    disabled: bool = False  # 入力を受け付けない設定（w:enabled が「いいえ」・fProt・\ffprot1）
    calc_on_exit: bool = False  # 欄を出るときに、計算の欄を計算し直す設定（w:calcOnExit・fRecalc・\ffrecalc1）


# 旧来のフォームの字の欄の種類（w:textInput/w:type、ST_FFTextType）。.doc の iTypeTxt・RTF の \fftypetxt は、この並びの番号（0〜5）
TEXT_TYPES = ("regular", "number", "date", "currentDate", "currentTime", "calculated")


def form_field_fixed(c) -> str:
    """旧来のフォームの欄のうち、人が入力しない欄の理由（入力する欄なら空の字）。ファイルに書かれた欄の設定だけで決める。"""
    if c is None or c.source != "formfield":
        return ""
    if c.text_type == "calculated":
        return f"計算の欄です（式「{c.expression}」の結果を Word が入れる欄で、入力する所ではありません）"
    if c.text_type == "currentDate":
        return "Word が今日の日付を入れる欄です（入力する所ではありません）"
    if c.text_type == "currentTime":
        return "Word がいまの時刻を入れる欄です（入力する所ではありません）"
    if c.disabled:
        return "入力を受け付けない設定の欄です（フォームの欄の設定で、入力できないようになっています）"
    return ""


def form_field_refusal(c, fid: str, force: bool) -> None:
    """人が入力しない欄（form_field_fixed）には書かない（force なら書く）。"""
    from .fill import FillError

    why = form_field_fixed(c)
    if why and not force:
        raise FillError(f"{fid}: {why}")


def form_field_notes(c, fid: str, text: str | None, controls) -> list[str]:
    """旧来のフォームの欄に書いたときに伝えること：数の欄に数字のない字を書いた・計算の欄を計算し直していない。
    text は書いた字（チェック・選択肢の欄は None）。"""
    out = []
    if c.text_type == "number" and text is not None and not any(ch.isdigit() for ch in text):
        out.append(f"{fid}: この欄は「数」の欄と設定されています。書いた「{text}」には数字がありません（字はそのまま書きました）")
    if c.calc_on_exit:
        calc = [x for x in controls if x.source == "formfield" and x.text_type == "calculated"]
        if calc:
            names = "、".join(f"{x.name or '（名前なし）'}（式「{x.expression}」）" for x in calc[:8])
            out.append(f"{fid}: この欄は、Word で欄を出るときに計算の欄を計算し直す設定です。とどくは計算の欄を計算し直していません"
                       f"（計算の欄 {len(calc)} か所：{names}。いま表示される値は、書く前のままです。Word で開いてフィールドを更新すると計算し直されます）")
    return out


Block = Para | Table


class Numbering:
    """段落番号（numbering.xml）。"""

    def __init__(self, root: Node | None):
        self.abstract: dict[str, dict[int, dict]] = {}
        self.nums: dict[str, tuple[str, dict[int, dict]]] = {}
        self.counters: dict[str, list[int]] = {}
        if root is None:
            return
        top = root.kids[0] if root.kids else None
        if top is None:
            return
        for an in top.findall("abstractNum"):
            aid = an.get("abstractNumId")
            self.abstract[aid] = {int(lv.get("ilvl", 0)): self._lvl(lv) for lv in an.findall("lvl")}
            link = an.find("numStyleLink")
            if link is not None:
                self.abstract[aid]["_link"] = link.get("val")
        for n in top.findall("num"):
            nid = n.get("numId")
            an = n.find("abstractNumId")
            over = {}
            for lo in n.findall("lvlOverride"):
                il = int(lo.get("ilvl", 0))
                d = {}
                so = lo.find("startOverride")
                if so is not None:
                    d["start"] = int(so.get("val", 1))
                lv = lo.find("lvl")
                if lv is not None:
                    d.update(self._lvl(lv))
                over[il] = d
            self.nums[nid] = (an.get("val") if an is not None else "", over)

    @staticmethod
    def _lvl(lv: Node) -> dict:
        d = {}
        for k in ("start", "numFmt", "lvlText", "lvlRestart"):
            x = lv.find(k)
            if x is not None:
                d[k] = x.get("val")
        if "start" in d:
            d["start"] = int(d["start"] or 1)
        return d

    def fmt(self, num_id: str, ilvl: int) -> str:
        """段落番号の書式（decimal・bullet など）。上書き（lvlOverride の lvl）があれば、そちら。"""
        if num_id not in self.nums:
            return ""
        aid, over = self.nums[num_id]
        lv = dict(self.abstract.get(aid, {}).get(ilvl, {}))
        lv.update({k: v for k, v in over.get(ilvl, {}).items() if k != "start"})
        return lv.get("numFmt") or "decimal"

    def label(self, num_id: str, ilvl: int) -> str:
        if num_id not in self.nums or num_id == "0":
            return ""
        aid, over = self.nums[num_id]
        levels = dict(self.abstract.get(aid, {}))
        key = aid
        cnt = self.counters.setdefault(key, [0] * 10)
        for i in range(ilvl + 1, 10):
            cnt[i] = 0
        lv = dict(levels.get(ilvl, {}))
        lv.update({k: v for k, v in over.get(ilvl, {}).items() if k != "start"})
        start = over.get(ilvl, {}).get("start", lv.get("start", 1))
        if cnt[ilvl] == 0:
            cnt[ilvl] = start
        else:
            cnt[ilvl] += 1
        text = lv.get("lvlText") or ""
        fmt = lv.get("numFmt") or "decimal"
        self.last_unshown = []  # この番号の中の、再現できない番号の書式（段階は、この事実で決める）
        if fmt == "bullet":
            return _bullet(text)
        if fmt == "none":
            return re.sub(r"%\d", "", text)

        def rep(m):
            i = int(m[1]) - 1
            lvi = dict(levels.get(i, {}))
            lvi.update({k: v for k, v in over.get(i, {}).items() if k != "start"})
            n = cnt[i] if cnt[i] else lvi.get("start", 1)
            f_ = lvi.get("numFmt") or "decimal"
            if _number(n, f_) is None:
                self.last_unshown.append(f_)
            return number_text(n, f_)

        return re.sub(r"%(\d)", rep, text)


def _bullet(t: str) -> str:
    if not t:
        return "・"
    o = ord(t[0])
    if 0xF000 <= o <= 0xF0FF:
        return WINGDINGS.get(o) or SYMBOL_FONT.get(o) or "・"
    return t


# 段落番号の書式（ECMA-376 第1部 17.18.59 ST_NumberFormat）。番号の字が Word と同じと確かめられる範囲だけを字にし、
# それ以外（書式を知らない・その番号の字の決まりを確かめていない）は、1, 2, 3 にせず〔…は再現できません〕と示す。
_KANJI = "〇一二三四五六七八九"
_AIUEO = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
_AIUEO_HALF = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ"
_IROHA = "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセスン"
# 半角のカタカナには ヰ・ヱ がない（JIS X 0201）。その番号は再現できないので None
_IROHA_HALF = [{"ヰ": None, "ヱ": None}.get(c, _AIUEO_HALF[_AIUEO.index(c)] if c in _AIUEO else None) for c in _IROHA]
_DAIJI_JA = "〇壱弐参四伍六七八九"  # japaneseLegal（日本の大字）
_DAIJI_ZH_S = "零壹贰叁肆伍陆柒捌玖"  # chineseLegalSimplified
_DAIJI_ZH_T = "零壹貳參肆伍陸柒捌玖"  # ideographLegalTraditional
_STEMS = "甲乙丙丁戊己庚辛壬癸"
_BRANCHES = "子丑寅卯辰巳午未申酉戌亥"
_GANADA = "가나다라마바사아자차카타파하"
_CHOSUNG = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ"
_KO_SINO = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
_KO_NATIVE = ["하나", "둘", "셋", "넷", "다섯", "여섯", "일곱", "여덟", "아홉", "열"]
# ロシア語の字の番号：Й・Ъ・Ы・Ь を除く 28 字（Ё も除く）。Word・LibreOffice の russianUpper/Lower と同じ並び
_RU_UPPER = "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ"
_HEBREW_ALEF = "אבגדהוזחטיכלמנסעפצקרשת"  # hebrew2：22 字（語末の形は使わない）
_HE_ONES = ["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"]
_HE_TENS = ["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"]
_HE_HUNDREDS = ["", "ק", "ר", "ש", "ת"]
_TH_DIGITS = ["", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"]


def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"),
            (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def _kanji(n: int) -> str | None:
    """位取りの漢数字（japaneseCounting：十・百・千の前の「一」は書かない）。1〜9999。"""
    if not 1 <= n <= 9999:
        return None
    out = []
    for digit, unit in zip(f"{n:04d}", ("千", "百", "十", "")):
        k = int(digit)
        if k:
            out.append(("" if k == 1 and unit else _KANJI[k]) + unit)
    return "".join(out)


def _counting_99(n: int, ones: list[str] | str, ten: str) -> str | None:
    """1〜99 の位取りの数（十の前の「一」は書かない：十・十一・二十・二十一）。"""
    if not 1 <= n <= 99:
        return None
    t, o = divmod(n, 10)
    return (("" if t == 1 else ones[t]) + ten if t else "") + (ones[o] if o else "")


def _digits(n: int, table: str | list[str]) -> str:
    return "".join(table[int(c)] for c in str(n))


def _cycle(chars, n: int, repeat: bool = False) -> str | None:
    """字の並びで1つずつ。repeat：並びを超えたら同じ字を重ねる（ECMA-376 17.18.59 の aiueo・iroha は「1つの字の1回以上の
    くり返し」：ｱ…ﾝ、ｱｱ…）。repeat でない並びを超える番号は、確かめていないので None。"""
    if n < 1 or (n > len(chars) and not repeat):
        return None
    ch = chars[(n - 1) % len(chars)]
    return None if ch is None else ch * ((n - 1) // len(chars) + 1)


def _hebrew1(n: int) -> str | None:
    """ヘブライ文字の数（hebrew1：1〜499。15 は טו、16 は טז（神の名を避ける書き方））。"""
    if not 1 <= n <= 499:
        return None
    h, rest = divmod(n, 100)
    t, o = divmod(rest, 10)
    if rest == 15:
        tail = "טו"
    elif rest == 16:
        tail = "טז"
    else:
        tail = _HE_TENS[t] + _HE_ONES[o]
    return _HE_HUNDREDS[h] + tail


def _thai_counting(n: int) -> str | None:
    """タイ語の数の言葉（thaiCounting：1〜99。十の位の 1 は สิบ、2 は ยี่สิบ、一の位の 1 は十のあとで เอ็ด）。"""
    if not 1 <= n <= 99:
        return None
    t, o = divmod(n, 10)
    head = "" if not t else ("สิบ" if t == 1 else ("ยี่สิบ" if t == 2 else _TH_DIGITS[t] + "สิบ"))
    tail = "" if not o else ("เอ็ด" if o == 1 and t else _TH_DIGITS[o])
    return head + tail


def _letters(n: int, base: str) -> str:
    """A, B, …, Z, AA, BB, …（Word の upperLetter・lowerLetter：同じ字を重ねる）。"""
    return base[(n - 1) % len(base)] * ((n - 1) // len(base) + 1)


def number_unshown(n: int, fmt: str) -> str:
    return f"〔{n}：番号の書式「{fmt}」は再現できません〕"


def _number(n: int, fmt: str) -> str | None:
    if fmt in ("decimal", "decimalHalfWidth"):
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt in ("decimalFullWidth", "decimalFullWidth2"):
        return str(n).translate(str.maketrans("0123456789", "０１２３４５６７８９"))
    if fmt in ("decimalEnclosedCircle", "decimalEnclosedCircleChinese"):
        return chr(0x2460 + n - 1) if 1 <= n <= 20 else (chr(0x3251 + n - 21) if 21 <= n <= 35 else str(n))
    if fmt == "decimalEnclosedParen":
        return chr(0x2474 + n - 1) if 1 <= n <= 20 else f"({n})"
    if fmt == "decimalEnclosedFullstop":
        return chr(0x2488 + n - 1) if 1 <= n <= 20 else f"{n}."
    if fmt == "numberInDash":
        return f"- {n} -"
    if fmt == "upperRoman":
        return _roman(n) if n >= 1 else None
    if fmt == "lowerRoman":
        return _roman(n).lower() if n >= 1 else None
    if fmt == "upperLetter":
        return _letters(n, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") if n >= 1 else None
    if fmt == "lowerLetter":
        return _letters(n, "abcdefghijklmnopqrstuvwxyz") if n >= 1 else None
    if fmt == "ordinal":
        return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"
    # 日本語
    if fmt == "aiueo":
        return _cycle(_AIUEO_HALF, n, True)
    if fmt == "aiueoFullWidth":
        return _cycle(_AIUEO, n, True)
    if fmt == "iroha":
        return _cycle(_IROHA_HALF, n, True)
    if fmt == "irohaFullWidth":
        return _cycle(_IROHA, n, True)
    if fmt == "japaneseCounting":
        return _kanji(n)
    if fmt == "japaneseLegal":
        return _DAIJI_JA[n] if 1 <= n <= 9 else None  # 10 からの書き方（拾・壱拾）は確かめていない
    if fmt == "japaneseDigitalTenThousand":
        return _digits(n, _KANJI) if 1 <= n <= 9999 else None
    if fmt in ("ideographDigital", "taiwaneseDigital", "koreanDigital2"):
        return _digits(n, _KANJI) if n >= 1 else None
    # 中国語
    if fmt in ("chineseCounting", "taiwaneseCounting"):
        return _counting_99(n, _KANJI, "十")  # 100 から（一百・一百零一）は確かめていない
    if fmt in ("chineseCountingThousand", "taiwaneseCountingThousand"):
        return _KANJI[n] if 1 <= n <= 9 else None
    if fmt == "chineseLegalSimplified":
        return _DAIJI_ZH_S[n] if 1 <= n <= 9 else None
    if fmt == "ideographLegalTraditional":
        return _DAIJI_ZH_T[n] if 1 <= n <= 9 else None
    if fmt == "ideographTraditional":
        return _cycle(_STEMS, n)
    if fmt == "ideographZodiac":
        return _cycle(_BRANCHES, n)
    if fmt == "ideographZodiacTraditional":
        return _STEMS[(n - 1) % 10] + _BRANCHES[(n - 1) % 12] if 1 <= n <= 60 else None  # 干支（甲子〜癸亥）
    if fmt == "ideographEnclosedCircle":
        return chr(0x3280 + n - 1) if 1 <= n <= 10 else None  # ㊀〜㊉
    # 韓国語
    if fmt == "ganada":
        return _cycle(_GANADA, n)
    if fmt == "chosung":
        return _cycle(_CHOSUNG, n)
    if fmt == "koreanDigital":
        return _KO_SINO[n] if 1 <= n <= 9 else None
    if fmt == "koreanCounting":
        return _counting_99(n, _KO_SINO, "십")
    if fmt == "koreanLegal":
        return _KO_NATIVE[n - 1] if 1 <= n <= 10 else None
    # ロシア語・ヘブライ語・インド・タイ
    if fmt == "russianUpper":
        return _cycle(_RU_UPPER, n)
    if fmt == "russianLower":
        return _cycle(_RU_UPPER.lower(), n)
    if fmt == "hebrew1":
        return _hebrew1(n)
    if fmt == "hebrew2":
        return _cycle(_HEBREW_ALEF, n)
    if fmt == "hindiNumbers":
        return "".join(chr(0x0966 + int(c)) for c in str(n)) if n >= 0 else None  # デーヴァナーガリーの数字
    if fmt == "thaiNumbers":
        return "".join(chr(0x0E50 + int(c)) for c in str(n)) if n >= 0 else None  # タイ数字
    if fmt == "thaiCounting":
        return _thai_counting(n)
    # arabicAlpha・arabicAbjad・hindiVowels・hindiConsonants・hindiCounting・thaiLetters・vietnameseCounting・
    # cardinalText・ordinalText・dollarText・bahtText・hex・chicago・custom など：字の並び（字の形・除く字）を確かめていない
    return None


def number_text(n: int, fmt: str) -> str:
    """段落番号の n 番目の字。再現できない書式・番号は、1, 2, 3 にせず〔n：番号の書式「…」は再現できません〕。"""
    got = _number(n, fmt)
    return got if got is not None else number_unshown(n, fmt)


def _lang_attr(fmt: str) -> str:
    """日付の書式に元号の記号（g・e）があれば東アジアの言語（w:lang の w:eastAsia）、なければ w:val（ECMA-376 第1部 17.3.2.20）。"""
    return "eastAsia" if re.search(r"[ge]", re.sub(r"'[^']*'", "", fmt or "")) else "val"


def _lang_of(rpr, attr: str) -> str:
    lang = rpr.find("lang") if rpr is not None else None
    return (lang.get(attr) or "") if lang is not None else ""


def _run_lang(doc, begin, fmt: str) -> tuple[str, str]:
    """旧来の日付の欄の言語と、どこで決まったか：(言語, 決まった所)。

    字の要素の言語は、スタイルの階層（ECMA-376 第1部 17.7.2）の順に決まる：字の要素の w:rPr/w:lang → 字のスタイル（w:rStyle と
    その basedOn）→ 段落のスタイル（w:pStyle、なければ既定の段落スタイル。basedOn も）→ 文書の既定（w:docDefaults/w:rPrDefault）。
    どこにもなければ、文書の設定 w:themeFontLang（17.15.1.88。Word が書類を作った PC の言語を書く。決まった所として伝える）。
    それもなければ ""（言語の指定なし：数だけの書式は書け、名前・元号の書式は書かない）。"""
    attr = _lang_attr(fmt)
    run = getattr(begin, "parent", None)
    rpr = run.find("rPr") if run is not None and hasattr(run, "find") else None
    got = _lang_of(rpr, attr)
    if got:
        return got, "字の要素の w:lang"
    st = doc.lang_styles() if doc is not None else {}

    def chain(sid: str) -> str:
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            v = st.get("langs", {}).get(sid, {}).get(attr, "")
            if v:
                return v
            sid = st.get("based", {}).get(sid, "")
        return ""

    rs = rpr.find("rStyle") if rpr is not None else None
    rsid = (rs.get("val") or "") if rs is not None else ""
    got = chain(rsid)
    if got:
        return got, f"字のスタイル「{rsid}」の w:lang"
    p = run
    while p is not None and getattr(p, "local", "") != "p":
        p = getattr(p, "parent", None)
    ppr = p.find("pPr") if p is not None else None
    ps = ppr.find("pStyle") if ppr is not None else None
    psid = ((ps.get("val") or "") if ps is not None else "") or st.get("para_default", "")
    got = chain(psid)
    if got:
        return got, f"段落のスタイル「{psid}」の w:lang"
    got = st.get("default", {}).get(attr, "")
    if got:
        return got, "文書の既定（w:docDefaults）の w:lang"
    got = st.get("theme", {}).get(attr, "")
    if got:
        return got, "文書の設定 w:themeFontLang"
    return "", ""


_PARTIAL: contextvars.ContextVar = contextvars.ContextVar("todoku_docx_partial", default=None)


@contextlib.contextmanager
def partial_base(base: dict):
    """この中で開く文書は、書類の一部（lazy/docxdoc）として読む：書き込みの試しの開き直しも、番号は書類全体と同じ。"""
    tok = _PARTIAL.set(base)
    try:
        yield
    finally:
        _PARTIAL.reset(tok)


class DocxDocument:
    kind = "docx"

    def __init__(self, blob: bytes, filename: str = "", base: dict | None = None):
        """base：書類の一部だけを読むとき（lazy/docxdoc）の、その前までの数（段落・表・欄・埋め込み・グラフ・絵の番号、段落番号の数え）と、
        読むヘッダーなどの部品（stories）。番号が書類全体を読んだときと同じになる。"""
        self.blob = blob
        self.filename = filename
        self.zip = RawZip(blob)
        base = base or _PARTIAL.get() or {}
        self.partial = bool(base)
        self.para_base = base.get("para", 0)
        self.table_base = base.get("table", 0)
        self.control_names = list(base.get("control_names", ()))  # 前までの欄の名前（ctl{番号}・同じ名前の #2 を、全体と同じに決める）
        self.embeds: list[dict] = [dict(e) for e in base.get("embeds", ())]  # 埋め込んだ OLE の部品（word/embeddings の部品）
        self.charts: list[dict] = [{"n": i + 1, "stub": True} for i in range(base.get("charts", 0))]  # グラフ（word/charts の部品）：本文の〔グラフ1〕の所
        self.diagrams: list[dict] = [{"n": i + 1, "stub": True} for i in range(base.get("diagrams", 0))]  # SmartArt（word/diagrams のデータの部品）：本文の〔SmartArt1〕の所
        self.pictures: list[dict] = [{"n": i + 1, "stub": True} for i in range(base.get("pictures", 0))]  # 絵（a:blip・v:imagedata）：説明（代わりの字）と飾りの印。段階とは別の注意に使う
        self.main = self._main_part()
        self.parts: dict[str, tuple[bytes, Node]] = {}
        self.rels = self._rels(self.main)
        self.has_macros = any(n.lower().endswith("vbaproject.bin") for n in self.zip.names())
        styles = self._load("word/styles.xml")
        self.style_names: dict[str, str] = {}
        self.style_num: dict[str, tuple[str, int]] = {}
        self.style_based: dict[str, str] = {}
        self.style_outline: dict[str, int] = {}  # スタイルの見出しの段（w:outlineLvl。0〜8 が見出し、9 は本文）
        if styles is not None:
            for st in styles.kids[0].findall("style") if styles.kids else []:
                sid = st.get("styleId")
                nm = st.find("name")
                self.style_names[sid] = nm.get("val") if nm is not None else sid
                b = st.find("basedOn")
                if b is not None:
                    self.style_based[sid] = b.get("val")
                ppr = st.find("pPr")
                ol = ppr.find("outlineLvl") if ppr is not None else None
                if ol is not None and str(ol.get("val", "")).isdigit():
                    self.style_outline[sid] = int(ol.get("val"))
                npr = ppr.find("numPr") if ppr is not None else None
                if npr is not None:
                    nid = npr.find("numId")
                    il = npr.find("ilvl")
                    self.style_num[sid] = (nid.get("val") if nid is not None else "", int(il.get("val")) if il is not None else 0)
        self.numbering = Numbering(self._load("word/numbering.xml"))
        if base.get("counters"):
            self.numbering.counters = {k: list(v) for k, v in base["counters"].items()}
        settings = self._load("word/settings.xml")
        self.protection = ""
        self.protection_enforced = False
        self.track_revisions = False  # 変更履歴の記録（w:trackRevisions）がオンか
        if settings is not None and settings.kids:
            tr_ = settings.kids[0].find("trackRevisions")
            self.track_revisions = tr_ is not None and str(tr_.get("val", "1")).lower() not in ("0", "false", "off")
            dp = settings.kids[0].find("documentProtection")
            if dp is not None:
                self.protection = dp.get("edit", "")
                self.protection_enforced = dp.get("enforcement", "0") in ("1", "true", "on")
        self.controls: list[Control] = []
        # 段落をまたいで開いているフィールド（目次・索引・参考文献など、結果が何段落にもなるもの）：本文の流れごとに持ち越す
        self._carry: list = []
        # 書類の一部だけを読むとき（lazy/docxdoc が数えたもの）：その前から開いているフィールドと、
        # この部分の終わりで開いているフィールドが後ろで閉じるか
        self._carry_base = tuple(base.get("fields", ()))
        self._carry_ends = base.get("fields_end")
        self.body = self._story(self.main)
        self.stories: list[tuple[str, str, list]] = []  # (種類, 部品, ブロック)
        for rid, (typ, target) in self.rels.items():
            kind = typ.rsplit("/", 1)[-1]
            if kind in ("header", "footer", "footnotes", "endnotes", "comments"):
                name = _join(self.main, target)
                if self.partial and name not in base.get("stories", ()):
                    continue
                if self.zip.has(name):
                    self.stories.append((kind, name, self._story(name)))
        if self.partial:  # 前までの分の控え（番号のためだけ）を除く
            self.embeds_all = self.embeds
            self.embeds = [e for e in self.embeds if not e.get("stub")]
            self.charts = [e for e in self.charts if not e.get("stub")]
            self.diagrams = [e for e in self.diagrams if not e.get("stub")]
            self.pictures = [e for e in self.pictures if not e.get("stub")]

    # ------------------------------------------------------------------ 部品

    @property
    def vba(self):
        """マクロ（VBA）のプロジェクト：モジュールとソース（todoku.vba.Project）。マクロがなければ None。"""
        if getattr(self, "_vba", None) is None and self.has_macros:
            from ..vba import read_project

            try:
                self._vba = read_project(self._vba_blob())
            except Exception as e:  # noqa: BLE001
                self._vba = e
        v = getattr(self, "_vba", None)
        return v if v is not None and not isinstance(v, Exception) else None

    def _vba_blob(self) -> bytes:
        name = next(n for n in self.zip.names() if n.lower().endswith("vbaproject.bin"))
        return self.zip.read(name)

    def _main_part(self) -> str:
        rels = self.zip.read("_rels/.rels") if self.zip.has("_rels/.rels") else b""
        m = re.search(rb'Target="/?([^"]*document[^"]*\.xml)"', rels)
        name = m[1].decode() if m else "word/document.xml"
        if not self.zip.has(name):
            raise DocxError("Word の文書の本体（document.xml）がありません")
        return name

    def _note_link(self, node: Node, part: str, segs: list) -> None:
        """w:hyperlink の先（r:id の関係、または w:anchor の文書の中の位置）。"""
        rid = node.get("id")
        target = self._rels(part).get(rid, ("", ""))[1] if rid else ""
        anchor = node.get("anchor")
        if anchor:
            target = f"{target}#{anchor}"
        if target:
            self.__dict__.setdefault("links", []).append(("".join(s.text for s in segs), target))

    def _note_field_link(self, instr: str, segs: list) -> None:
        from ..links import field_target

        target = field_target(instr)
        if target:
            self.__dict__.setdefault("links", []).append(("".join(s.text for s in segs), target))

    def _rels(self, part: str) -> dict[str, tuple[str, str]]:
        d, base = part.rsplit("/", 1) if "/" in part else ("", part)
        rp = f"{d}/_rels/{base}.rels" if d else f"_rels/{base}.rels"
        out = {}
        if self.zip.has(rp):
            for m in re.finditer(rb"<(?:\w+:)?Relationship\b([^>]*)/?>", self.zip.read(rp)):
                a = dict(re.findall(r'(\w+)="([^"]*)"', m[1].decode("utf-8", "replace")))
                if "Id" in a:
                    out[a["Id"]] = (a.get("Type", ""), a.get("Target", ""))
        return out

    def _load(self, name: str) -> Node | None:
        if name in self.parts:
            return self.parts[name][1]
        if not self.zip.has(name):
            return None
        xml = self.zip.read(name)
        root = parse(xml)
        self.parts[name] = (xml, root)
        return root

    def xml(self, part: str) -> bytes:
        return self.parts[part][0]

    def lang_styles(self) -> dict:
        """字の言語（w:lang）の決まり方の材料：スタイルごとの w:lang・basedOn・既定の段落スタイル・文書の既定・themeFontLang。"""
        memo = getattr(self, "_lang_styles", None)
        if memo is not None:
            return memo

        def attrs(el) -> dict:
            return {a: el.get(a) for a in ("val", "eastAsia", "bidi") if el is not None and el.get(a)}

        out: dict = {"langs": {}, "based": {}, "para_default": "", "default": {}, "theme": {}}
        root = self._load("word/styles.xml")
        top = root.kids[0] if root is not None and root.kids else None
        if top is not None:
            dd = top.find("docDefaults")
            rd = dd.find("rPrDefault") if dd is not None else None
            rpr = rd.find("rPr") if rd is not None else None
            out["default"] = attrs(rpr.find("lang") if rpr is not None else None)
            for st in top.findall("style"):
                sid = st.get("styleId") or ""
                b = st.find("basedOn")
                if b is not None:
                    out["based"][sid] = b.get("val") or ""
                rp = st.find("rPr")
                out["langs"][sid] = attrs(rp.find("lang") if rp is not None else None)
                if st.get("type") == "paragraph" and st.get("default") in ("1", "true", "on"):
                    out["para_default"] = sid
        settings = self._load("word/settings.xml")
        if settings is not None and settings.kids:
            out["theme"] = attrs(settings.kids[0].find("themeFontLang"))
        self._lang_styles = out
        return out

    def font_styles(self) -> dict:
        """字の書式の決まり方の材料：スタイルごとの (基のスタイル, 取り消し線, 字の色)・既定の段落スタイル・文書の既定。"""
        memo = self.__dict__.get("_font_styles")
        if memo is not None:
            return memo
        out: dict = {"styles": {}, "para_default": "", "default": (None, None)}
        root = self._load("word/styles.xml")
        top = root.kids[0] if root is not None and root.kids else None
        if top is not None:
            dd = top.find("docDefaults")
            rd = dd.find("rPrDefault") if dd is not None else None
            out["default"] = _look_of(rd.find("rPr") if rd is not None else None)
            for st in top.findall("style"):
                b = st.find("basedOn")
                out["styles"][st.get("styleId") or ""] = (b.get("val") or "" if b is not None else "", *_look_of(st.find("rPr")))
                if st.get("type") == "paragraph" and st.get("default") in ("1", "true", "on"):
                    out["para_default"] = st.get("styleId") or ""
        self.__dict__["_font_styles"] = out
        self.__dict__["_look_memo"] = {}
        return out

    def run_look(self, rpr, para_style: str) -> tuple:
        """字の書式の事実（取り消し線か, 字の色）：文書の既定 → 段落のスタイル → 字のスタイル → 字の書式の順で、近いものが勝つ
        （ISO 29500-1 17.7.2。スタイルの取り消し線の打ち消し合いは見ない）。表のスタイルの字の書式は見ない。"""
        fs = self.font_styles()
        rs = rpr.find("rStyle") if rpr is not None else None
        rstyle = rs.get("val") or "" if rs is not None else ""
        memo = self.__dict__["_look_memo"]
        base = memo.get((para_style, rstyle))
        if base is None:
            strike, color = fs["default"]

            def walk(sid: str) -> None:
                nonlocal strike, color
                chain: list = []
                while sid in fs["styles"] and sid not in chain and len(chain) < 32:
                    chain.append(sid)
                    sid = fs["styles"][sid][0]
                for x in reversed(chain):
                    _b, s_, c_ = fs["styles"][x]
                    strike = strike if s_ is None else s_
                    color = color if c_ is None else c_

            walk(para_style or fs["para_default"])
            walk(rstyle)
            base = memo[(para_style, rstyle)] = (bool(strike), color)
        s_, c_ = _look_of(rpr)
        return (base[0] if s_ is None else s_, base[1] if c_ is None else c_)

    # ------------------------------------------------------------------ 本文
    def _story(self, part: str) -> list:
        root = self._load(part)
        if root is None or not root.kids:
            return []
        top = root.kids[0]
        container = top.find("body") if top.local == "document" else top
        if container is None:
            return []
        if top.local in ("footnotes", "endnotes", "comments"):
            out = []
            for item in top.kids:
                if item.local in ("footnote", "endnote", "comment"):
                    if item.get("type") in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    self._carry = []
                    out += self._blocks(item, part)
            return out
        self._carry = [_carried_field(f) for f in self._carry_base] if part == self.main else []
        out = self._blocks(container, part)
        if part == self.main and self._carry_ends is not None and len(self._carry_ends) == len(self._carry):
            for f, closed in reversed(list(zip(self._carry, self._carry_ends))):  # 内側のフィールドから
                if closed:
                    _mark_held(f)
        return out

    def _blocks(self, node: Node, part: str, cell: Cell | None = None) -> list:
        out = []
        for ch in node.kids:
            loc = ch.local
            if loc == "p":
                out.append(self._para(ch, part, cell))
            elif loc == "tbl":
                out.append(self._table(ch, part))
            elif loc == "sdt":
                content = ch.find("sdtContent")
                ctl = self._sdt_control(ch, part, None)
                inner = self._blocks(content, part, cell) if content is not None else []
                if ctl is not None:
                    ctl.result_segs = [s for b in inner if isinstance(b, Para) for s in b.segs]
                    ctl.para = next((b for b in inner if isinstance(b, Para)), None)
                    for b in inner:
                        if isinstance(b, Para):
                            for s in b.segs:
                                s.field = s.field or ctl
                out += inner
            elif loc in ("customXml", "ins", "moveTo"):
                out += self._blocks(ch, part, cell)
        return out

    def _table(self, node: Node, part: str) -> Table:
        t = Table(node, part)
        for r, tr in enumerate([x for x in node.kids if x.local in ("tr", "sdt", "customXml")]):
            rows = [tr] if tr.local == "tr" else [x for x in tr.iter("tr")]
            for tr2 in rows:
                row = []
                col = 0
                for tc in [x for x in tr2.kids if x.local == "tc"] + [x for s in tr2.kids if s.local in ("sdt", "customXml") for x in s.iter("tc")]:
                    pr = tc.find("tcPr")
                    span = 1
                    vm = ""
                    shade = ""
                    if pr is not None:
                        gs = pr.find("gridSpan")
                        if gs is not None:
                            span = int(gs.get("val", 1) or 1)
                        v = pr.find("vMerge")
                        if v is not None:
                            vm = v.get("val", "continue") or "continue"
                        sh = pr.find("shd")
                        if sh is not None:
                            shade = sh.get("fill", "") or ""
                    c = Cell(tc, len(t.rows), col, span, vm, table=t, shade=shade)
                    c.blocks = self._blocks(tc, part, c)
                    row.append(c)
                    col += span
                t.rows.append(row)
        return t

    def _para(self, node: Node, part: str, cell: Cell | None) -> Para:
        p = Para(node, part, [], cell=cell)
        ppr = node.find("pPr")
        num = None
        if ppr is not None:
            ps = ppr.find("pStyle")
            if ps is not None:
                p.style = ps.get("val", "")
            npr = ppr.find("numPr")
            if npr is not None:
                nid = npr.find("numId")
                il = npr.find("ilvl")
                num = (nid.get("val") if nid is not None else "", int(il.get("val")) if il is not None else 0)
        if num is None and p.style:
            s = p.style
            seen = 0
            while s and seen < 20:
                if s in self.style_num:
                    num = self.style_num[s]
                    break
                s = self.style_based.get(s, "")
                seen += 1
        if num is not None and num[0]:
            p.label = self.numbering.label(num[0], num[1])
            if getattr(self.numbering, "last_unshown", None):
                p.number_unshown = list(self.numbering.last_unshown)  # 再現できない番号の書式（窓では〔n：番号の書式「…」は再現できません〕）
            if p.label in BULLET_BOXES and self.numbering.fmt(num[0], num[1]) == "bullet":
                # 箇条書きの印が □ の段落（「□ Ⅰ 負担上限月額に関する認定」）：チェックの欄にできる（印は段落番号の定義の字）
                p.bullet = {"num": num[0], "ilvl": num[1], "direct": ppr is not None and ppr.find("numPr") is not None}
        self._inline(node, p, part, {"stack": self._carry})
        self._carry_over(p, part)
        if any(s.kind == "tab" for s in p.segs):
            p.leaders = self._leaders(ppr, p.style)
        return p

    def _carry_over(self, p: Para, part: str) -> None:
        """段落の終わりで、まだ閉じていないフィールドを次の段落へ持ち越す（フィールドの始まり・区切り・終わりの印は、
        本文の流れの順に対になる。ECMA-376 第1部 17.16.2：結果は何段落にもなれる）。
        結果の途中のフィールド（フォームの欄でないもの）は、この段落のぶんの結果の字を控え、終わりの印が来たときに
        フィールドの結果の印を付ける（終わりの印のないフィールドは、いままでどおり、ふつうの字として読む）。
        フォームの欄は、いままでどおり、始まりの段落の中だけで読む（後ろの段落の終わりの印は、対として数えるだけ）。"""
        for f in self._carry:
            if f.get("ghost"):
                continue
            if f.get("ctl") is not None:
                f["ghost"] = True
                continue
            if f["phase"] == "result":
                f.setdefault("held", []).append((p.segs[f.get("n0", 0):], part, p))
                f["n0"] = 0
            f["carried"] = True

    def tab_styles(self) -> dict:
        """タブ位置の決まり方の材料：スタイルごとの (基のスタイル, タブ位置の並び [(位置, 種類, リーダー)])・既定の段落スタイル。"""
        memo = self.__dict__.get("_tab_styles")
        if memo is not None:
            return memo
        out: dict = {"styles": {}, "para_default": ""}
        root = self._load("word/styles.xml")
        top = root.kids[0] if root is not None and root.kids else None
        if top is not None:
            for st in top.findall("style"):
                b = st.find("basedOn")
                out["styles"][st.get("styleId") or ""] = (b.get("val") or "" if b is not None else "", _tab_stops(st.find("pPr")))
                if st.get("type") == "paragraph" and st.get("default") in ("1", "true", "on"):
                    out["para_default"] = st.get("styleId") or ""
        self.__dict__["_tab_styles"] = out
        return out

    def _leaders(self, ppr, style: str) -> tuple:
        """段落のタブ位置（スタイルの基から順に重ね、w:val="clear" の位置は消す。ISO 29500-1 17.3.1.38）のうち、
        線（リーダー）を引くものの、線の種類（dot・hyphen・underscore・heavy・middleDot）。"""
        ts = self.tab_styles()
        chain, s = [], style or ts["para_default"]
        while s and s in ts["styles"] and s not in chain and len(chain) < 20:
            chain.append(s)
            s = ts["styles"][s][0]
        stops: dict = {}
        for sid in reversed(chain):
            _lay_stops(stops, ts["styles"][sid][1])
        _lay_stops(stops, _tab_stops(ppr))
        return tuple(sorted({x for x in stops.values() if x and x != "none"}))

    def _inline(self, node: Node, p: Para, part: str, st: dict, deleted: bool = False) -> None:
        for ch in node.kids:
            loc = ch.local
            if loc == "r":
                n0 = len(p.segs)
                self._run(ch, p, part, st, deleted)
                if len(p.segs) > n0:
                    look = self.run_look(ch.find("rPr"), p.style)
                    for s in p.segs[n0:]:
                        s.look = look
            elif loc in ("hyperlink", "smartTag", "customXml", "ins", "moveTo", "fldSimple", "bdo", "dir"):
                n0 = len(p.segs)
                self._inline(ch, p, part, st, deleted)
                if loc == "hyperlink":
                    self._note_link(ch, part, p.segs[n0:])
                if loc == "fldSimple":
                    self._note_field_link(ch.get("instr", ""), p.segs[n0:])
                if loc == "fldSimple":
                    _mark_result(p.segs[n0:], ch.get("instr", ""), ch, part, p, ch.get("fldLock", "") in ("1", "true", "on"))
            elif loc in ("del", "moveFrom"):
                self._inline(ch, p, part, st, True)
            elif loc == "sdt":
                ctl = self._sdt_control(ch, part, p)
                content = ch.find("sdtContent")
                if content is not None:
                    n0 = len(p.segs)
                    self._inline(content, p, part, st, deleted)
                    if ctl is not None:
                        ctl.result_segs = p.segs[n0:]
                        for s in ctl.result_segs:
                            s.field = ctl
                        if not ctl.result_segs:
                            p.segs.append(Seg("", "field", content, None, field=ctl))
                            ctl.result_segs = [p.segs[-1]]
            elif loc == "oMath" or loc == "oMathPara":
                txt = "".join(n.text for n in ch.iter("t"))
                if txt:
                    p.segs.append(Seg(txt, "math", None))

    def _run(self, r: Node, p: Para, part: str, st: dict, deleted: bool) -> None:
        rpr = r.find("rPr")
        hidden = deleted
        underline = False
        sym_font = ""
        if rpr is not None:
            v = rpr.find("vanish")
            if v is not None and v.get("val", "true") not in ("0", "false", "off"):
                hidden = True
            u = rpr.find("u")
            if u is not None and u.get("val", "single") not in ("none", "0"):
                underline = True
            rf = rpr.find("rFonts")
            if rf is not None:
                sym_font = rf.get("ascii", "") or rf.get("hAnsi", "") or ""
        stack: list = st["stack"]
        for ch in r.kids:
            loc = ch.local
            # 前の段落から続く命令の途中では、命令の字（instrText）だけを命令として読む（本文の字は落とさない）
            in_instr = bool(stack) and stack[-1]["phase"] == "instr" and not stack[-1].get("carried")
            if loc == "fldChar":
                typ = ch.get("fldCharType")
                if typ == "begin":
                    stack.append({"phase": "instr", "instr": "", "begin": ch, "ctl": None, "run": r,
                                  "locked": ch.get("fldLock", "") in ("1", "true", "on")})
                elif typ == "separate" and stack:
                    f = stack[-1]
                    f["phase"] = "result"
                    if f.get("carried") and _is_form_instr(f["instr"]):
                        f["ghost"] = True  # 命令が段落をまたぐフォームの欄：欄にしない（いままでどおり）
                    elif not f.get("carried"):
                        f["ctl"] = self._form_field(f, part, p)
                    f["n0"] = len(p.segs)
                elif typ == "end" and stack:
                    f = stack.pop()
                    if f.get("ghost"):
                        continue
                    if f.get("carried") and f["phase"] == "instr":
                        continue  # 結果のないフィールドが段落をまたいで終わった
                    if f.get("ctl") is None and f["phase"] == "instr":
                        f["ctl"] = self._form_field(f, part, p)
                        f["n0"] = len(p.segs)
                    _mark_held(f)  # 前の段落までの結果の字
                    ctl = f.get("ctl")
                    if ctl is None and not f.get("carried") and len(p.segs) == f.get("n0", len(p.segs)):
                        # 結果の字のない EQ フィールド（古い形のふりがな）：命令の中の字が見える字
                        from ..eqfield import eq_text

                        got = eq_text(f["instr"])
                        if got is not None:
                            p.segs.append(Seg(got[0], "ruby" if got[1] is not None else "other", None, r))
                            if got[1]:
                                p.segs += [Seg("（", "other", None, r), Seg(got[1], "ruby", None, r), Seg("）", "other", None, r)]
                    if ctl is None and f["phase"] == "result":
                        self._note_field_link(f["instr"], p.segs[f.get("n0", len(p.segs)):])
                    if ctl is None and f["phase"] == "result":
                        # 差し込み・計算などのフィールドの結果（MERGEFIELD・DATE・REF・= など）：Word が更新すると作り直す字
                        _mark_result(p.segs[f.get("n0", len(p.segs)):], f["instr"], f["begin"], part, p, f["locked"])
                    if ctl is not None:
                        ctl.end = ch
                        ctl.result_segs = p.segs[f.get("n0", len(p.segs)):]
                        for s in ctl.result_segs:
                            s.field = s.field or ctl
                        if ctl.kind == "checkbox":
                            p.segs.append(Seg("☑" if ctl.checked else "☐", "field", ch, r, field=ctl))
                        elif not ctl.result_segs:
                            p.segs.append(Seg("", "field", ch, r, field=ctl))
                            ctl.result_segs = [p.segs[-1]]
                continue
            if loc == "instrText":
                if stack:
                    stack[-1]["instr"] += ch.text
                continue
            if in_instr:
                continue
            if loc == "t":
                p.segs.append(Seg(ch.text, "t", ch, r, hidden, underline))
            elif loc == "delText":
                p.segs.append(Seg(ch.text, "del", ch, r, True))  # 変更履歴で削除された字（窓では印をつけて示す）
            elif loc == "tab" or loc == "ptab":
                p.segs.append(Seg("\t", "tab", ch, r, hidden, underline))
            elif loc in ("br", "cr"):
                p.segs.append(Seg("\n", "br", ch, r, hidden))
            elif loc == "noBreakHyphen":
                p.segs.append(Seg("‑", "other", ch, r, hidden))
            elif loc == "sym":
                code = int(ch.get("char", "0") or "0", 16)
                font = ch.get("font", "") or sym_font
                if code < 0xF000:
                    code2 = code + 0xF000
                else:
                    code2 = code
                if "wingding" in font.lower():
                    txt = WINGDINGS.get(code2, "◆")
                elif "symbol" in font.lower():
                    txt = SYMBOL_FONT.get(code2, chr(code2 - 0xF000) if code2 >= 0xF020 else "・")
                else:
                    txt = chr(code) if code < 0xE000 else "・"
                p.segs.append(Seg(txt, "sym", ch, r, hidden, underline))
            elif loc in ("footnoteReference", "endnoteReference"):
                p.segs.append(Seg(f"※{ch.get('id', '')}", "note", ch, r, hidden))
            elif loc in ("drawing", "pict", "object", "AlternateContent"):
                self._boxes(ch, p, part)
                axc = next(iter(ch.iter("control")), None) if loc in ("object", "pict") else None
                if axc is not None:
                    from ..axfile import word_control

                    word_control(self, ch, axc, part, p, r, hidden)  # ActiveX の部品（w:control）
                    continue
                for mark in self._graphics(ch, part):
                    p.segs.append(Seg(mark, "other", None, r, hidden))
                ole = next(iter(ch.iter("OLEObject")), None)
                e = self._add_embed(ole, part, p, ch) if ole is not None else None
                alt = [n.get("descr") or n.get("title") for n in ch.iter("docPr") if n.get("descr") or n.get("title")]
                if e is not None and "alt" not in e:
                    e["alt"] = next((n.get("alt") for n in ch.iter("shape") if n.get("alt")), "") or (alt[0] if alt else "")
                if e is not None:
                    p.segs.append(Seg(f"〔埋め込み{e['n']}：{e['label']}〕", "other", None, r, hidden))
                elif self._picture(ch, p, alt[0] if alt else ""):
                    pic = self.pictures[-1]
                    mark = f"〔図: {pic['alt']}〕" if pic["alt"] else f"〔図{pic['n']}（説明なし）〕" if not pic["decorative"] else ""
                    if mark:
                        p.segs.append(Seg(mark, "other", None, r, hidden))
                elif alt:
                    p.segs.append(Seg(f"〔図: {alt[0]}〕", "other", None, r, hidden))
            elif loc == "ruby":
                base = ch.find("rubyBase")
                rt = ch.find("rt")
                bt = "".join(n.text for n in base.iter("t")) if base is not None else ""
                rtt = "".join(n.text for n in rt.iter("t")) if rt is not None else ""
                p.segs.append(Seg(bt, "ruby", None, r, hidden))
                if rtt.strip():  # 空白だけのふりがなは、字がない（窓に「（　）」を作らない。空欄の候補にもしない）
                    p.segs.append(Seg("（", "other", None, r, hidden))
                    p.segs.append(Seg(rtt, "ruby", None, r, hidden))
                    p.segs.append(Seg("）", "other", None, r, hidden))

    def _add_embed(self, ole: Node, part: str, p: Para, holder: Node | None = None) -> dict | None:
        """埋め込んだ OLE の部品（o:OLEObject）を、文書の埋め込みの一覧に足す：部品の置き場所（word/embeddings/…）を関係から引く。
        holder（w:object など）の中の絵（v:imagedata・a:blip）は、Word が表示・印刷に使う表の絵として覚える（images）。"""
        rid = ole.get("id")
        rels = self._rels(part)
        typ, target = rels.get(rid or "", ("", ""))
        if not target or ole.get("Type", "Embed") != "Embed":
            return None
        name = _join(part, target)
        if not self.zip.has(name):
            return None
        for e in self.embeds:
            if e["storage"] == name:
                return e
        images = []
        for n in list(holder.iter("imagedata")) + list(holder.iter("blip")) if holder is not None else []:
            t = rels.get(n.get("id") or n.get("embed") or "", ("", ""))[1]
            img = _join(part, t) if t else ""
            if img and self.zip.has(img) and img not in images:
                images.append(img)
        prog = ole.get("ProgID") or ""
        from ..embeds.ole import label_of

        label = label_of(prog)
        e = {"n": len(self.embeds) + 1, "storage": name, "prog": prog, "label": label, "para": p, "part": part,
             "images": images, "aspect": ole.get("DrawAspect") or "Content"}
        self.embeds.append(e)
        return e

    def _picture(self, node: Node, p: Para, alt: str) -> bool:
        """図の中に絵（a:blip・v:imagedata）があれば、絵の一覧に足す。代わりの中身（Fallback）とテキストボックスの中は数えない。"""
        found = None
        for n in list(node.iter("blip")) + list(node.iter("imagedata")):
            if not any(x.local in ("Fallback", "txbxContent") for x in _ancestors(n, node)):
                found = n
                break
        if found is None:
            return False
        rid = found.get("embed") or found.get("id") or ""
        target = self._rels(p.part).get(rid, ("", ""))[1] if rid else ""
        media = _join(p.part, target) if target else ""
        pr = next(iter(node.iter("docPr")), None)
        deco = pr is not None and any(x.local == "decorative" and x.get("val") in ("1", "true") for x in pr.iter("decorative"))
        if not alt:
            alt = next((s.get("alt") for s in node.iter("shape") if s.get("alt")), "") or ""  # VML の絵の説明
        ext = next(iter(node.iter("extent")), None)
        size = None
        if ext is not None and (ext.get("cx") or "").isdigit() and (ext.get("cy") or "").isdigit():
            size = [round(int(ext.get("cx")) / 36000, 1), round(int(ext.get("cy")) / 36000, 1)]  # EMU → mm
        self.pictures.append({"n": len(self.pictures) + 1, "alt": alt.strip(), "decorative": deco, "para": p,
                              "media": media if self.zip.has(media) else "", "size_mm": size})
        return True

    def _graphics(self, node: Node, part: str) -> list[str]:
        """図の中のグラフ（c:chart・新しい種類の cx:chart）と SmartArt（dgm:relIds）を一覧に足し、本文に出す印を返す。
        AlternateContent の代わりの中身（Fallback）は数えない。"""
        out = []
        rels = None
        for g in node.iter("graphicData"):
            if any(x.local in ("Fallback", "txbxContent") for x in _ancestors(g, node)):
                continue  # 代わりの中身・テキストボックスの中（中身を段落として読むときに数える）
            for k in g.kids:
                if k.local not in ("chart", "relIds"):
                    continue
                rels = rels if rels is not None else self._rels(part)
                rid = k.get("id") if k.local == "chart" else k.get("dm")
                target = _join(part, rels.get(rid or "", ("", ""))[1]) if rid in rels else ""
                if not target or not self.zip.has(target):
                    continue
                pr = next((n for n in node.iter("docPr")), None)
                alt = (pr.get("descr") or pr.get("title") or "") if pr is not None else ""
                if k.local == "chart":
                    self.charts.append({"n": len(self.charts) + 1, "part": target, "alt": alt})
                    out.append(f"〔グラフ{len(self.charts)}〕")
                else:
                    self.diagrams.append({"n": len(self.diagrams) + 1, "part": target, "alt": alt})
                    out.append(f"〔SmartArt{len(self.diagrams)}〕")
        return out

    def embedded_books(self) -> list[tuple[dict, object]]:
        """埋め込んだ Excel のブック：(埋め込みの情報, ブック)。.xlsx の部品も、OLE の入れ物（.bin）の中のブックも読む。"""
        out = []
        for e in self.embeds:
            wb = self.embedded_book(e)
            if wb is not None:
                out.append((e, wb))
        return out

    def embedded_book(self, e: dict):
        """埋め込み1つの Excel のブック（Excel のブックでない・開けないときは None）。"""
        from ..cfb import CFB
        from ..xls.book import XlsWorkbook
        from ..xlsx.workbook import Workbook

        try:
            data = self.zip.read(e["storage"])
            if data[:4] == b"PK\x03\x04":
                return Workbook(data, e["storage"])
            if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                c = CFB(data)
                names = {n.lower(): n for n in c.streams()}
                if "workbook" in names or "book" in names:
                    return XlsWorkbook(data, e["storage"])
                if "package" in names:
                    return Workbook(c.read(names["package"]), e["storage"])
        except Exception:  # noqa: BLE001
            return None
        return None

    def _boxes(self, node: Node, p: Para, part: str) -> None:
        """図形・テキストボックスの中の文字（段落の後ろに付ける）。描画キャンバス・グループの中のテキストボックスも、
        すべて読む。AlternateContent は、Word が表示する Choice の方だけを読む（Fallback は同じ中身の代わりなので数えない）。
        テキストボックスの中のテキストボックスは、外の箱の段落の中で読む（ここでは二重に数えない）。"""
        for tx in node.iter("txbxContent"):
            ups = list(_ancestors(tx, node))
            if any(a.local in ("Fallback", "txbxContent") for a in ups):
                continue
            keep, self._carry = self._carry, []  # テキストボックスの中は、別の流れ
            try:
                p.boxes += self._blocks(tx, part)
            finally:
                self._carry = keep

    # ------------------------------------------------------------------ 入力欄
    def _form_field(self, f: dict, part: str, p: Para) -> Control | None:
        instr = f["instr"].strip()
        word = instr.split()[0].upper() if instr.split() else ""
        if word not in ("FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN"):
            return None
        ff = f["begin"].find("ffData")
        name = ""
        ctl = Control({"FORMTEXT": "text", "FORMCHECKBOX": "checkbox", "FORMDROPDOWN": "dropdown"}[word], "formfield",
                      "", f["begin"], part, p)
        if ff is not None:
            n = ff.find("name")
            name = n.get("val", "") if n is not None else ""
            en, coe = ff.find("enabled"), ff.find("calcOnExit")
            ctl.disabled = en is not None and en.get("val", "true") in ("0", "false", "off")  # ない・値なしは「入力できる」
            ctl.calc_on_exit = coe is not None and coe.get("val", "true") not in ("0", "false", "off")
            ht = ff.find("helpText") or ff.find("statusText")
            if ht is not None:
                ctl.help = ht.get("val", "")
            cb = ff.find("checkBox")
            if cb is not None:
                c = cb.find("checked")
                d = cb.find("default")
                if c is not None:
                    ctl.checked = c.get("val", "1") not in ("0", "false")
                elif d is not None:
                    ctl.checked = d.get("val", "0") not in ("0", "false")
            dd = ff.find("ddList")
            if dd is not None:
                ctl.options = [x.get("val", "") for x in dd.findall("listEntry")]
                res = dd.find("result")
                idx = int(res.get("val", 0)) if res is not None else 0
                ctl.value = ctl.options[idx] if 0 <= idx < len(ctl.options) else None
            ti = ff.find("textInput")
            if ti is not None:
                ml = ti.find("maxLength")
                if ml is not None and ml.get("val"):
                    ctl.maxlen = int(ml.get("val"))
                tp = ti.find("type")
                ctl.text_type = tp.get("val", "regular") if tp is not None else "regular"
                if ctl.text_type == "date":
                    ctl.kind = "date"
                fmt = ti.find("format")
                if fmt is not None:
                    ctl.date_format = ctl.text_format = fmt.get("val", "")
                dv = ti.find("default")
                if ctl.text_type == "calculated" and dv is not None:
                    ctl.expression = dv.get("val", "")
                if ctl.kind == "date":
                    ctl.date_lid, ctl.date_lid_from = _run_lang(self, f["begin"], ctl.date_format)
        ctl.name = name
        self.controls.append(ctl)
        return ctl

    def _sdt_control(self, sdt: Node, part: str, p: Para | None) -> Control | None:
        pr = sdt.find("sdtPr")
        if pr is None:
            return None
        kind = "richtext"
        options: list[str] = []
        date_fmt = date_lid = date_cal = ""
        checked = False
        multiline = True
        for k in pr.kids:
            loc = k.local
            if loc == "text":
                kind = "text"
                multiline = (k.get("multiLine") or "").lower() in ("1", "true", "on")
            elif loc == "date":
                kind = "date"
                df = k.find("dateFormat")
                date_fmt = df.get("val", "") if df is not None else ""
                lid_el, cal_el = k.find("lid"), k.find("calendar")
                date_lid = lid_el.get("val", "") if lid_el is not None else ""
                date_cal = cal_el.get("val", "") if cal_el is not None else ""
            elif loc in ("dropDownList", "comboBox"):
                kind = "dropdown" if loc == "dropDownList" else "combo"
                options = [li.get("displayText") or li.get("value") or "" for li in k.findall("listItem")]
            elif loc == "checkbox" and k.tag.startswith("{" + W14):
                kind = "checkbox"
                c = k.find("checked")
                checked = c is not None and c.get("val", "1") in ("1", "true")
            elif loc == "picture":
                kind = "picture"
            elif loc in ("docPartObj", "docPartList", "bibliography", "citation", "group", "equation"):
                return None
        alias = pr.find("alias")
        tag = pr.find("tag")
        lock = pr.find("lock")
        ctl = Control(kind, "sdt", (alias.get("val") if alias is not None else "") or (tag.get("val") if tag is not None else ""),
                      sdt, part, p, options=options, checked=checked, date_format=date_fmt,
                      placeholder=pr.find("showingPlcHdr") is not None, lock=lock.get("val", "") if lock is not None else "",
                      date_lid=date_lid, date_calendar=date_cal, multiline=multiline)
        self.controls.append(ctl)
        return ctl

    # ------------------------------------------------------------------ 全体
    def all_blocks(self):
        """(種類, 部品, ブロック)。本文が先、ヘッダー・フッター・脚注・コメントが後。"""
        yield "body", self.main, self.body
        for kind, part, blocks in self.stories:
            yield kind, part, blocks


def _look_of(rpr) -> tuple:
    """字の書式の要素の、取り消し線（w:strike・w:dstrike）と字の色（w:color の val。auto は黒）：指定がなければ None。"""
    if rpr is None:
        return None, None
    strike = None
    for name in ("strike", "dstrike"):
        el = rpr.find(name)
        if el is not None:
            on = el.get("val", "true") not in ("0", "false", "off")
            strike = on or bool(strike)
    c = rpr.find("color")
    color = None
    if c is not None:
        v = (c.get("val") or "").upper()
        color = "000000" if v == "AUTO" else v if len(v) == 6 and all(ch in "0123456789ABCDEF" for ch in v) else None
    return strike, color


def _is_form_instr(instr: str) -> bool:
    w = instr.split()
    return bool(w) and w[0].upper() in ("FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN")


def _carried_field(f: dict) -> dict:
    """書類の一部だけを読むとき、その前から開いているフィールド（lazy/docxdoc が数えた phase・instr・locked）。"""
    form = _is_form_instr(f.get("instr", ""))
    return {"phase": f["phase"], "instr": f.get("instr", ""), "begin": None, "ctl": None, "run": None,
            "locked": bool(f.get("locked")), "carried": True, "n0": 0, "ghost": form and f["phase"] == "result"}


def _mark_held(f: dict) -> None:
    """段落をまたぐフィールドの、前の段落までの結果の字に、フィールドの結果の印を付ける。"""
    for segs, part, p in f.pop("held", ()):
        _mark_result(segs, f["instr"], f["begin"], part, p, f["locked"])


def _mark_result(segs: list, instr: str, node: Node, part: str, p: Para, locked: bool) -> None:
    """フィールドの結果の字（フォームの欄でないもの）に、そのフィールドの印（kind "field" の Control。文書の入力欄の一覧には
    入れない）を付ける。結果の字は、Word がフィールドを更新すると作り直す（ECMA-376 第1部 17.16）ので、入力欄の推定・
    字を目印にした書き込みの対象にしない（固定した（fldLock）フィールドは、そのことを添える）。"""
    instr = " ".join(str(instr).split())
    if not instr:
        return
    fc = Control("field", "field-result", instr.split()[0].upper(), node, part, p)
    fc.instr = instr
    fc.locked = locked
    for s in segs:
        if s.field is None:
            s.field = fc


def _ancestors(n: Node, stop: Node):
    """n の親を、stop まで順に。"""
    x = n.parent
    while x is not None and x is not stop:
        yield x
        x = x.parent


def iter_paras(blocks):
    for b in blocks:
        if isinstance(b, Para):
            yield b
            if b.boxes:
                yield from iter_paras(b.boxes)
        elif isinstance(b, Table):
            for row in b.rows:
                for c in row:
                    yield from iter_paras(c.blocks)


def _join(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    d = base.rsplit("/", 1)[0] if "/" in base else ""
    parts = (d + "/" + target).split("/") if d else target.split("/")
    out = []
    for x in parts:
        if x == "..":
            if out:
                out.pop()
        elif x and x != ".":
            out.append(x)
    return "/".join(out)
