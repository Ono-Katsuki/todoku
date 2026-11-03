"""表示形式のロケール・暦・数字の指定（[$-…]・[DBNum…]・B1/B2）。

Excel の表示形式の [$-NNCCLLLL]（ECMA-376 第1部 18.8.31 numFmts の「ロケール」の記法。各バイトの意味は、Microsoft の
「Changing Numeral Shapes in Excel」（Office Global Experience の記事：[$-nnccLCID]、nn = 数字の形、cc = 暦、LCID = ロケール）と、
LibreOffice の Excel 互換の読み込み（svl/source/numbers/zformat.cxx の ImpObtainCalendarAndNumerals・LocaleType）に合わせる）：
- 下位2バイト LLLL：ロケール（LCID。言語の名前・曜日の名前・和暦／民国の年の決め方に使う）
- 3バイト目 CC：暦（00・01・02 グレゴリオ暦、03 和暦、04 台湾（民国）、06 ヒジュラ暦、07 タイの仏暦、17 ウンム・アル＝クラー暦 ほか）
- 上位バイト NN：数字の形（02 アラビア・インド数字、03 ペルシャの数字、04 デーヴァナーガリー数字、0D タイ数字 ほか）
BCP 47 の名前（[$-ja-JP]・[$-ja-JP-x-gannen] など）も同じ意味で読む。

ここにない組み合わせ・確かめていない書き方は、推測して字にしない（Unsupported を投げ、窓は元の値と一緒にそう示す）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .values import Unsupported


@dataclass(frozen=True)
class Loc:
    """1つのセクションの、ロケール・暦・数字の指定。"""

    lang: str | None = None  # None：ロケールの指定なし（Excel を使う PC の設定。todoku は日本語の PC として扱う）
    calendar: int = 0  # 暦（CC）。0 は既定（グレゴリオ暦）
    numerals: int = 0  # 数字の形（NN）。0 は既定
    dbnum: int = 0  # [DBNum1]〜[DBNum4]
    gannen: bool = False  # [$-ja-JP-x-gannen]：元号の最初の年を「元年」と書く
    hijri_prefix: bool = False  # セクションの先頭の B2（ヒジュラ暦）
    text: str = ""  # 元の指定（知らせる文に使う）


# ---------------------------------------------------------------- ロケール（LCID と BCP 47 の名前）
# LCID の主言語（下位 10 ビット）。[MS-LCID] Windows Language Code Identifier Reference の表から、ここで名前を持つ言語だけ
_PRIMARY = {0x11: "ja", 0x12: "ko", 0x1E: "th", 0x01: "ar", 0x0C: "fr", 0x07: "de", 0x0A: "es", 0x16: "pt", 0x19: "ru", 0x09: "en"}
# 中国語は、地域で字が違う（[MS-LCID]：0x0404 台湾・0x0C04 香港・0x1404 マカオは繁体字、0x0804 中国・0x1004 シンガポールは簡体字）
_ZH = {0x0404: "zh-TW", 0x0C04: "zh-TW", 0x1404: "zh-TW", 0x0804: "zh-CN", 0x1004: "zh-CN"}
_TAGS = {"ja": "ja", "ko": "ko", "th": "th", "ar": "ar", "fr": "fr", "de": "de", "es": "es", "pt": "pt", "ru": "ru", "en": "en"}


def _lang_of_lcid(lcid: int) -> str | None:
    if lcid == 0:
        return None
    if lcid in (0xF800, 0xF400):
        raise Unsupported("PC の設定の長い日付・時刻の形式（[$-F800]・[$-F400]）")
    if lcid & 0x3FF == 0x04:
        if lcid in _ZH:
            return _ZH[lcid]
        return f"lcid:{lcid:04X}"
    return _PRIMARY.get(lcid & 0x3FF, f"lcid:{lcid:04X}")


def _lang_of_tag(tag: str) -> tuple[str, bool]:
    parts = tag.split("-")
    low = [p.lower() for p in parts]
    gannen = False
    if "x" in low:
        i = low.index("x")
        priv = low[i + 1:]
        if priv == ["gannen"] and low[0] == "ja":
            gannen = True
        elif priv in (["sysdate"], ["systime"]):
            raise Unsupported(f"PC の設定の日付・時刻の形式（[$-{tag}]）")
        else:
            raise Unsupported(f"ロケールの指定 [$-{tag}]")
        low = low[:i]
    if len(low) > 2 or any(not re.fullmatch(r"[a-z0-9]{1,8}", p) for p in low):
        raise Unsupported(f"ロケールの指定 [$-{tag}]")
    lang = low[0]
    if lang == "zh":
        region = low[1] if len(low) > 1 else ""
        if region in ("tw", "hk", "mo", "hant"):
            return "zh-TW", gannen
        if region in ("cn", "sg", "hans"):
            return "zh-CN", gannen
        raise Unsupported(f"ロケールの指定 [$-{tag}]")
    return _TAGS.get(lang, f"tag:{lang}"), gannen


_LOC_RE = re.compile(r"\[\$([^\]\-]*)-([^\]]*)\]")


def parse_section(sec: str) -> tuple[str, Loc]:
    """セクションから [$…-…]・[DBNum…]・先頭の B1/B2 を取り出す。通貨記号（[$¥-411] の ¥）は字として残す。"""
    lang, cal, num, gannen, texts = None, 0, 0, False, []
    seen: str | None = None

    def loc(m):
        nonlocal lang, cal, num, gannen, seen
        cur, spec = m[1], m[2]
        texts.append(m[0])
        if cur and re.fullmatch(r"(?i)x-euro[12]?", spec):
            spec = ""  # [$€-x-euro2] など：ユーロの記号の書き方の違いだけ（ロケールは決めない）
        if spec:
            if seen is not None:
                if spec.lower() != seen:
                    raise Unsupported(f"1つのセクションに違うロケールの指定（{''.join(texts)}）")
                return '"' + cur.replace('"', "") + '"' if cur else ""
            seen = spec.lower()
            if re.fullmatch(r"[0-9A-Fa-f]{1,8}", spec):
                v = int(spec, 16)
                lang = _lang_of_lcid(v & 0xFFFF)
                cal = (v >> 16) & 0xFF
                num = (v >> 24) & 0xFF
            else:
                lang, gannen = _lang_of_tag(spec)
        return '"' + cur.replace('"', "") + '"' if cur else ""

    sec = _LOC_RE.sub(loc, sec)
    m = re.search(r"\[\$([^\]]*)\]", sec)  # [$€] のように、ロケールのない通貨記号
    if m:
        sec = sec.replace(m[0], '"' + m[1].replace('"', "") + '"')
    dbnum = 0
    for m in re.finditer(r"(?i)\[DBNum(\d)\]", sec):
        dbnum = int(m[1])
        texts.append(m[0])
    sec = re.sub(r"(?i)\[DBNum\d\]", "", sec)
    if dbnum and dbnum not in (1, 2, 3, 4):
        raise Unsupported(f"[DBNum{dbnum}]")
    hijri = False
    m = re.match(r"^((?:\[[^\]]*\])*)\s*B([12])", sec)
    if m:
        hijri = m[2] == "2"
        texts.append("B" + m[2])
        sec = m[1] + sec[m.end():]
    return sec, Loc(lang, cal, num, dbnum, gannen, hijri, "".join(texts))


# ---------------------------------------------------------------- 数字の形（NN）
# LibreOffice zformat.cxx の数字の形の番号（0x02〜0x13）と、Unicode の各文字の 0〜9（各文字の Nd の並び）
_DIGIT_ZERO = {
    0x02: 0x0660,  # アラビア・インド数字
    0x03: 0x06F0,  # ペルシャ（拡張アラビア・インド）数字
    0x04: 0x0966,  # デーヴァナーガリー数字
    0x05: 0x09E6,  # ベンガル数字
    0x06: 0x0A66,  # グルムキー数字
    0x07: 0x0AE6,  # グジャラート数字
    0x08: 0x0B66,  # オリヤー数字
    0x09: 0x0BE6,  # タミル数字
    0x0A: 0x0C66,  # テルグ数字
    0x0B: 0x0CE6,  # カンナダ数字
    0x0C: 0x0D66,  # マラヤーラム数字
    0x0D: 0x0E50,  # タイ数字
    0x0E: 0x0ED0,  # ラーオ数字
    0x0F: 0x0F20,  # チベット数字
    0x10: 0x1040,  # ミャンマー数字
    0x12: 0x17E0,  # クメール数字
    0x13: 0x1810,  # モンゴル数字
}
# CJK の数字の番号：0x1B〜0x1D は日本語の [DBNum1]〜[DBNum3] と同じ（zformat.cxx：NatNum1〜3）
_CJK_JA = {0x1B: 1, 0x1C: 2, 0x1D: 3}


def numeral_mode(loc: Loc) -> tuple[int, int]:
    """(数字の形, DBNum)。両方の指定・意味を確かめていない番号は Unsupported。"""
    n, db = loc.numerals, loc.dbnum
    if n in (0, 1):
        n = 0
    elif n in _CJK_JA:
        if db:
            raise Unsupported(f"数字の形と [DBNum] の両方の指定（{loc.text}）")
        n, db = 0, _CJK_JA[n]
    elif n not in _DIGIT_ZERO:
        raise Unsupported(f"数字の形の番号 0x{n:02X}（{loc.text}）")
    if db:
        lang = loc.lang
        if lang not in (None, "ja"):
            raise Unsupported(f"日本語でないロケールの [DBNum{db}]（{loc.text}）")
        if db == 4:
            raise Unsupported("[DBNum4]（日本語）")
        if n:
            raise Unsupported(f"数字の形と [DBNum] の両方の指定（{loc.text}）")
    return n, db


def shape_digits(s: str, numerals: int) -> str:
    """0〜9 を、その数字の形の字にする（区切りの字はそのまま：Excel の区切りは PC の設定で決まり、ロケールでは変わらない）。"""
    if not numerals:
        return s
    if numerals in (0x02, 0x03) and ("," in s or "." in s):
        # アラビア文字の数字には、区切りの字（٬ ٫）が別にある。Excel がどちらを出すかを確かめていないので、字にしない
        raise Unsupported(f"アラビア・インド数字（0x{numerals:02X}）の、桁区切り・小数点の字")
    z = _DIGIT_ZERO[numerals]
    return "".join(chr(z + ord(ch) - 48) if "0" <= ch <= "9" else ch for ch in s)


# ---------------------------------------------------------------- 漢数字（[DBNum1]〜[DBNum3]、日本語）
# 日本語版 Excel の [DBNum1]（漢数字）・[DBNum2]（大字）・[DBNum3]（全角の数字）。
# 「G/標準」では位取り（十・百・千・万）をつけ、0・# の桁の記号では1桁ずつ字にする（LibreOffice の対応：
# DBNum1 → NatNum4・DBNum2 → NatNum5・DBNum3 → NatNum3、日付は DBNum1〜3 → NatNum1〜3）
_KAN = "〇一二三四五六七八九"
_DAIJI = "〇壱弐参四伍六七八九"
_ZEN = "０１２３４５６７８９"


def dbnum_digits(s: str, db: int) -> str:
    """1桁ずつ字にする（0 や # の桁の記号の表示形式）。"""
    if db == 2 and "0" in s:
        raise Unsupported("[DBNum2]（大字）の 0 の字")
    table = {1: _KAN, 2: _DAIJI, 3: _ZEN}[db]
    out = []
    for ch in s:
        if "0" <= ch <= "9":
            out.append(table[ord(ch) - 48])
        elif ch in "-":
            out.append(ch)
        else:
            raise Unsupported(f"[DBNum{db}] の数の中の「{ch}」")
    return "".join(out)


def dbnum_units(n: int, db: int) -> str:
    """位取りをつけて字にする（「G/標準」・日付の年月日）。1 〜 9999,9999,9999 の整数だけ。"""
    if n <= 0 or n >= 10 ** 12:
        raise Unsupported(f"[DBNum{db}] の位取りで書く数 {n}")
    if db == 1:
        digits, units, groups, keep_one = _KAN, ("千", "百", "十", ""), ("億", "万", ""), False
    elif db == 2:
        digits, units, groups, keep_one = _DAIJI, ("阡", "百", "拾", ""), ("億", "萬", ""), True
    else:
        digits, units, groups, keep_one = _ZEN, ("千", "百", "十", ""), ("億", "万", ""), False
    s = f"{n:012d}"
    out = []
    for gi in range(3):
        g = s[gi * 4:(gi + 1) * 4]
        if g == "0000":
            continue
        for k, ch in enumerate(g):
            d = ord(ch) - 48
            if d == 0:
                continue
            if d == 1 and units[k] and not keep_one:
                out.append(units[k])
            else:
                out.append(digits[d] + units[k])
        out.append(groups[gi])
    return "".join(out)


# ---------------------------------------------------------------- 月・曜日の名前
# 各ロケールの月・曜日の名前（Excel が [$-LCID] で使う Windows の各国語の名前：LOCALE_SMONTHNAME・LOCALE_SABBREVMONTHNAME・
# LOCALE_SDAYNAME・LOCALE_SABBREVDAYNAME。.NET の CultureInfo(...).DateTimeFormat の MonthNames・DayNames と同じ値）。
# 確かめられなかった省略形（None）は、推測して出さない。曜日は日曜から。
_EN_M = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
_EN_D = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
_NAMES: dict[str, dict[str, list[str] | None]] = {
    "en": {"mmmm": _EN_M, "mmm": [m[:3] for m in _EN_M], "dddd": _EN_D, "ddd": [d[:3] for d in _EN_D]},
    "ja": {"mmmm": [f"{i}月" for i in range(1, 13)], "mmm": [f"{i}月" for i in range(1, 13)],
           "dddd": [c + "曜日" for c in "日月火水木金土"], "ddd": list("日月火水木金土")},
    "zh-CN": {"mmmm": ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"],
              "mmm": [f"{i}月" for i in range(1, 13)],
              "dddd": ["星期" + c for c in "日一二三四五六"], "ddd": ["周" + c for c in "日一二三四五六"]},
    "zh-TW": {"mmmm": ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"],
              "mmm": [f"{i}月" for i in range(1, 13)],
              "dddd": ["星期" + c for c in "日一二三四五六"], "ddd": ["週" + c for c in "日一二三四五六"]},
    "ko": {"mmmm": [f"{i}월" for i in range(1, 13)], "mmm": None,
           "dddd": [c + "요일" for c in "일월화수목금토"], "ddd": list("일월화수목금토")},
    "th": {"mmmm": ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน",
                    "ตุลาคม", "พฤศจิกายน", "ธันวาคม"],
           "mmm": ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."],
           "dddd": ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"],
           "ddd": ["อา.", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส."]},
    # アラビア語：グレゴリオ暦の月の名前は国で違う（エジプトの يناير・レバント地方の كانون الثاني など）ので、出さない
    "ar": {"mmmm": None, "mmm": None, "dddd": ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], "ddd": None},
    "fr": {"mmmm": ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"],
           "mmm": ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."],
           "dddd": ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"],
           "ddd": ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."]},
    "de": {"mmmm": ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
           "mmm": ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"],
           "dddd": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"],
           "ddd": ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"]},
    "es": {"mmmm": ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
           "mmm": None, "dddd": ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], "ddd": None},
    "pt": {"mmmm": ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"],
           "mmm": None, "dddd": ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"],
           "ddd": None},
    # ロシア語の月：日（d）と並ぶときは生格（ноября）、月だけのときは主格（ноябрь）（Windows の GetDateFormat と同じ使い分け）
    "ru": {"mmmm": ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"],
           "mmmm_gen": ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"],
           "mmm": None, "dddd": ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], "ddd": None},
}
# 「aaa」「aaaa」（日本語版の曜日の記号）：日本語・中国語・韓国語のロケールだけ
_AAA = {None: "ja", "ja": "ja", "ko": "ko", "zh-CN": "zh-CN", "zh-TW": "zh-TW"}


def name_table(lang: str | None, kind: str, with_day: bool = False) -> list[str]:
    """月・曜日の名前の表（kind：mmm・mmmm・ddd・dddd・aaa・aaaa）。ロケールなし（PC の設定）は、日本語版 Excel と同じ：
    mmm・ddd は英語、aaa は日本語。"""
    if kind in ("aaa", "aaaa"):
        base = _AAA.get(lang)
        if base is None:
            raise Unsupported(f"このロケールの曜日の記号 {kind}")
        return _NAMES[base]["ddd" if kind == "aaa" else "dddd"]
    t = _NAMES.get(lang or "en")
    if t is None:
        raise Unsupported(f"ロケール {lang} の月・曜日の名前")
    if kind == "mmmm" and with_day and "mmmm_gen" in t:
        kind = "mmmm_gen"
    names = t.get(kind)
    if names is None:
        raise Unsupported(f"ロケール {lang} の {kind}（名前の省略形）")
    return names


# ---------------------------------------------------------------- 暦
def hijri_tabular(y: int, m: int, d: int) -> tuple[int, int, int]:
    """グレゴリオ暦の日付を、表で決まるヒジュラ暦（Excel の Hijri・Windows の HijriCalendar の「クウェートの方法」：
    30 年で 11 回のうるう年（2・5・7・10・13・16・18・21・24・26・29 年）、1 年 1 月 1 日 = 622 年 7 月 16 日（ユリウス暦））にする。"""
    import datetime as dt
    import math

    epoch = 1948439.5
    jd = dt.date(y, m, d).toordinal() + 1721424.5

    def to_jd(hy, hm, hd):
        return hd + math.ceil(29.5 * (hm - 1)) + (hy - 1) * 354 + math.floor((3 + 11 * hy) / 30) + epoch - 1

    hy = math.floor((30 * (jd - epoch) + 10646) / 10631)
    hm = min(12, math.ceil((jd - (29 + to_jd(hy, 1, 1))) / 29.5) + 1)
    hd = int(jd - to_jd(hy, hm, 1) + 1)
    return int(hy), int(hm), hd
