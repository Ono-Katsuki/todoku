"""表示形式（"#,##0" や "ggge年m月d日" など）で数値を文字列にする。

TEXT関数と、AIに渡す「見えている値」の両方で使う。再現できない形式は Unsupported を投げる。
"""

from __future__ import annotations

import math
import re
from functools import lru_cache
from decimal import ROUND_HALF_UP, Decimal

from .dates import serial_parts, serial_to_date, to_wareki, weekday_sun0
from .locale_fmt import Loc, dbnum_digits, dbnum_units, hijri_tabular, name_table, numeral_mode, parse_section, shape_digits
from .values import Unsupported, dec15, number_to_text

# 組み込みの表示形式ID（ECMA-376 第1部 18.8.30）。ファイルには番号だけがあり、形式の字（formatCode）はない。
# 14（短い日付）・22（短い日付と時刻）と、27〜36・50〜58（東アジアの日付・時刻）は、開く PC の地域の設定で形式が変わる
# （18.8.30 は、この番号を zh-tw・zh-cn・ja-jp・ko-kr・th-th で別々に定め、ほかの地域の形は定めない。14 は PC の「短い日付」：
# 日本語の PC は yyyy/m/d、米国は m/d/yyyy、ドイツは dd.mm.yyyy）。todoku は書類を受け取る人の PC を知らないので、この番号は
# 特定の地域の字にしない：format_value は LocaleDependent を投げ、窓には地域によらない書き方（ISO 8601）を添えて、見え方が
# 変わることを示す。ブックが形式の字を書いていれば（styles.xml の numFmts・BIFF の FORMAT）、そちらを使う。
LOCALE_BUILTIN = frozenset({14, 22, *range(27, 37), *range(50, 59)})
# todoku の印（Excel の形式の字ではない）。後ろの ISO の形は、日付の形式と分かるため（is_date_format・入力の読み取り）
_BUILTIN_TAG = "[$-x-builtin-{}]yyyy-mm-dd"
_BUILTIN_RE = re.compile(r"\[\$-x-builtin-(\d+)\]")
BUILTIN = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
    **{i: _BUILTIN_TAG.format(i) for i in sorted(LOCALE_BUILTIN)},
}
# 日本語の PC（ja-JP）での、地域で変わる番号の形（ブックが日本語の PC の設定と分かるときに使う。18.8.30 の ja-jp の表）
JA_JP_BUILTIN = {
    14: "yyyy/m/d", 22: "yyyy/m/d h:mm",
    27: '[$-411]ge.m.d', 28: '[$-411]ggge"年"m"月"d"日"', 29: '[$-411]ggge"年"m"月"d"日"',
    30: "m/d/yy", 31: 'yyyy"年"m"月"d"日"', 32: 'h"時"mm"分"', 33: 'h"時"mm"分"ss"秒"',
    34: 'yyyy"年"m"月"', 35: 'm"月"d"日"', 36: '[$-411]ge.m.d',
    50: '[$-411]ge.m.d', 51: '[$-411]ggge"年"m"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
    54: '[$-411]ggge"年"m"月"d"日"', 55: 'yyyy"年"m"月"', 56: 'm"月"d"日"', 57: '[$-411]ge.m.d',
    58: '[$-411]ggge"年"m"月"d"日"',
}


class LocaleDependent(Unsupported):
    """組み込みの表示形式で、開く PC の地域の設定で見え方が変わるもの。shown：地域によらない書き方（ISO 8601）。"""

    def __init__(self, fmt_id: int, shown: str):
        super().__init__(f"組み込みの表示形式（番号 {fmt_id}）は、開く PC の地域の設定で見え方が変わる"
                         f"（地域によらない書き方では {shown}）")
        self.fmt_id = fmt_id
        self.shown = shown


def builtin_id(fmt: str | None) -> int | None:
    """PC の地域の設定で変わる組み込みの表示形式（の印）なら、その番号。"""
    m = _BUILTIN_RE.search(fmt or "")
    return int(m[1]) if m else None


def iso_text(x: float, date1904: bool | None = None) -> str:
    """シリアル値の、地域によらない書き方（ISO 8601：yyyy-mm-dd、時刻があれば Thh:mm:ss。1900 年の数え方の 1 未満は時刻だけ）。"""
    from .dates import current_1904

    if x < 0:
        return number_to_text(x)
    day = math.floor(x)
    secs = int(Decimal(repr((x - day) * 86400)).quantize(Decimal(1), rounding=ROUND_HALF_UP))
    if secs >= 86400:
        day, secs = day + 1, secs - 86400
    hh, rem = divmod(secs, 3600)
    t = f"{hh:02d}:{rem // 60:02d}:{rem % 60:02d}"
    if day == 0 and not (current_1904() if date1904 is None else date1904):
        return t
    y, m, d = serial_parts(day, date1904)
    return f"{y:04d}-{m:02d}-{d:02d}" + (f"T{t}" if secs else "")

_DATE_CODE_RE = re.compile(r"(?i)(?<![\\])(yy|y|d|h|s|g|e|am/pm)")


def _strip_literals(fmt: str) -> str:
    return re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", fmt)


def is_date_format(fmt: str) -> bool:
    sec = fmt.split(";")[0]
    if re.search(r"(?i)\[(h+|m+|s+)\]", sec):
        return True  # [h]:mm などの経過時間
    core = _strip_literals(sec)
    core = re.sub(r"(?i)general", "", core)  # 「標準」の G・e を日付の記号と取り違えない
    core = re.sub(r"[eE][+-]", "", core)  # 指数表示（0.00E+00）
    return bool(re.search(r"(?i)[ydhsgeb]", core)) or bool(re.search(r"(?i)m", core) and not re.search(r"[0#?]", core))


def split_sections(fmt: str) -> list[str]:
    return list(_split_sections(fmt))


@lru_cache(maxsize=4096)
def _split_sections(fmt: str) -> tuple[str, ...]:
    """表示形式を ; で分ける（同じ形式は何度も出てくるので、分けた結果を控える）。"""
    out, buf, q = [], [], False
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            q = not q
        elif ch == "\\" and not q and i + 1 < len(fmt):
            buf.append(fmt[i:i + 2])
            i += 2
            continue
        elif ch == ";" and not q:
            out.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    out.append("".join(buf))
    return tuple(out)


def format_value(v, fmt: str, date1904: bool | None = None) -> str:
    """セルの値を表示形式で文字列にする。date1904：ブックの日付の数え方（None は計算中のブックの数え方。dates.py）。
    再現できない表示形式は、推測した字を返さずに Unsupported を投げる。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, str):
        secs = split_sections(fmt)
        if len(secs) >= 4:
            t = parse_section(secs[3])[0]
            return _literal(t).replace("@", v) if "@" in t else _literal(t)
        if "@" in fmt and len(secs) == 1:
            return _literal(parse_section(fmt)[0]).replace("@", v)
        return v
    x = float(v)
    bid = builtin_id(fmt)
    if bid is not None:
        raise LocaleDependent(bid, iso_text(x, date1904))
    if fmt in ("General", "general", ""):
        return _general(x)
    secs = _split_sections(fmt)
    if len(secs) == 1 and "@" in _literal_free(secs[0]):
        # 文字列の区画（@）だけの表示形式：数は標準の表示（Excel で、数のセルに「文字列」の書式を付けたとき）
        return _general(x)
    sec = None
    neg_has_own = False
    if _conditional(fmt):
        # 条件のあるセクション（ECMA-376 第1部 18.8.31：[<=999]000;000-0000 など）。負の数の符号の付け方は確かめていない
        sec = _conditional_section(secs, x, fmt)
        secs = (sec,)
    if len(secs) == 1:
        sec = secs[0]
    elif len(secs) >= 2:
        if x > 0 or (x == 0 and len(secs) < 3):
            sec = secs[0]
        elif x < 0:
            sec = secs[1]
            neg_has_own = True
        else:
            sec = secs[2]
    if neg_has_own:
        x = abs(x)
    sec, kind, loc = _clean_section(sec)
    numerals, db = numeral_mode(loc)
    if kind == "general":
        if db:
            if x != int(x) or x < 0:
                raise Unsupported(f"[DBNum{db}] の標準の表示で、整数でない数・負の数")
            return dbnum_units(int(x), db)
        return shape_digits(_general(x), numerals)
    if kind == "date":
        return _format_date(x, sec, loc, numerals, db, date1904)
    if re.search(r"(?i)general", _literal_free(sec)):
        # 字と「標準」を組み合わせたセクション（"計"General など）：標準の表示の前後に字を付ける
        if len(re.findall(r"(?i)general", sec)) != 1 or len(re.findall(r"(?i)general", _literal_free(sec))) != 1:
            raise Unsupported(f"標準（General）を含む表示形式: {sec}")
        head, tail = re.split(r"(?i)general", sec, maxsplit=1)
        g = dbnum_units(int(x), db) if db and x == int(x) and x > 0 else (_raise("[DBNum] と標準") if db else _general(x))
        return _literal(head) + shape_digits(g, numerals) + _literal(tail)
    if db:
        return _format_number(x, sec, lambda t: dbnum_digits(t, db))
    return _format_number(x, sec, (lambda t: shape_digits(t, numerals)) if numerals else None)


_COND = re.compile(r"\[(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)\]")


def _conditional_section(secs: tuple[str, ...], x: float, fmt: str) -> str:
    """条件のあるセクションから、値に使うセクションを選ぶ（条件を取り除いて返す）：1つ目・2つ目の条件を順に見て、
    合うものを使う。どちらにも合わなければ、3つ目のセクション（2つ目に条件がなければ2つ目）。"""
    if x < 0:
        raise Unsupported(f"条件付きの表示形式の負の数: {fmt}")
    ops = {"<=": x.__le__, ">=": x.__ge__, "<>": x.__ne__, "<": x.__lt__, ">": x.__gt__, "=": x.__eq__}
    conds = []
    for sec in secs[:4]:
        m = _COND.search(sec)
        if m is None and re.search(r"\[(<|>|=)", sec):
            raise Unsupported(f"読めない条件: {fmt}")
        conds.append((m, _COND.sub("", sec, count=1) if m else sec))
    if len(secs) > 3 or (len(secs) == 3 and conds[2][0] is not None):
        raise Unsupported(f"条件付きの表示形式の形: {fmt}")
    for m, body in conds[:2]:
        if m is not None and ops[m[1]](float(m[2])):
            return body
    if len(conds) >= 3:
        return conds[2][1]
    if len(conds) == 2 and conds[1][0] is None:
        return conds[1][1]
    raise Unsupported(f"どの条件にも合わない値: {fmt}")


@lru_cache(maxsize=4096)
def _conditional(fmt: str) -> bool:
    return any(re.search(r"\[(<|>|=)", s) for s in _split_sections(fmt))


@lru_cache(maxsize=4096)
def _clean_section(sec: str) -> tuple[str, str, Loc]:
    """セクションから色・ロケール・漢数字の指定を取り出し、種類（general・date・number）を決める（同じセクションは控える）。"""
    sec = re.sub(r"(?i)\[(red|black|blue|green|white|yellow|magenta|cyan|color\s*\d+)\]", "", sec)
    sec, loc = parse_section(sec)  # [$-411]・[$-ja-JP-x-gannen]・[$¥-411]（通貨記号は字）・[DBNum1]・B2
    if sec.strip().lower() == "general":
        return sec, "general", loc
    return sec, "date" if (is_date_format(sec) or loc.hijri_prefix) else "number", loc


def _general(x: float) -> str:
    if x == 0:
        return "0"
    a = abs(x)
    if a >= 1e11 or a < 1e-9:
        s = f"{x:.5E}"
        mant, exp = s.split("E")
        mant = mant.rstrip("0").rstrip(".")
        return f"{mant}E{exp[0]}{exp[1:].lstrip('0').rjust(2, '0')}"
    # 標準の表示は11桁まで
    digits = 10 - int(math.floor(math.log10(a))) if a >= 1 else 10
    d = dec15(x).quantize(Decimal(1).scaleb(-max(0, min(digits, 10))), rounding=ROUND_HALF_UP)
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


def _literal(sec: str) -> str:
    out = []
    i = 0
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.index('"', i + 1)
            out.append(sec[i + 1:j])
            i = j + 1
            continue
        if ch == "\\" and i + 1 < len(sec):
            out.append(sec[i + 1])
            i += 2
            continue
        if ch == "_" and i + 1 < len(sec):
            out.append(" ")
            i += 2
            continue
        if ch == "*" and i + 1 < len(sec):
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _tokens(sec: str) -> list[tuple[str, str]]:
    """('lit', 文字) か ('code', 記号) の列。"""
    out = []
    i = 0
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.index('"', i + 1)
            out.append(("lit", sec[i + 1:j]))
            i = j + 1
        elif ch == "\\" and i + 1 < len(sec):
            out.append(("lit", sec[i + 1]))
            i += 2
        elif ch == "_" and i + 1 < len(sec):
            out.append(("lit", " "))
            i += 2
        elif ch == "*" and i + 1 < len(sec):
            i += 2
        elif ch == "[":
            j = sec.index("]", i)
            out.append(("code", sec[i:j + 1]))
            i = j + 1
        else:
            out.append(("code", ch))
            i += 1
    return out


def _format_number(x: float, sec: str, shape=None) -> str:
    """shape：数の部分の字を変える（数字の形・漢数字）。"""
    toks = _tokens(sec)
    codes = "".join(t for k, t in toks if k == "code")
    if "E" in codes.upper() and re.search(r"[0#]E[+-]", codes, re.I):
        raise Unsupported(f"指数の表示形式: {sec}")
    if "/" in codes and re.search(r"[0#?]\s*/\s*[0#?\d]", codes):
        raise Unsupported(f"分数の表示形式: {sec}")
    pct = codes.count("%")
    x = x * (100 ** pct)
    # 末尾のカンマ（千単位で割る）
    m = re.search(r"([0#?])(,+)(?![0#?])", codes)
    if m:
        x = x / (1000 ** len(m[2]))
    int_part, _, frac_part = codes.partition(".")
    frac_digits = len(re.findall(r"[0#?]", frac_part))
    thousands = bool(re.search(r"[0#?],[0#?]", int_part))
    min_int = len(re.findall(r"0", int_part.replace(",", "")))
    neg = x < 0
    d = dec15(abs(x)).quantize(Decimal(1).scaleb(-frac_digits), rounding=ROUND_HALF_UP)
    s = format(d, "f")
    ip, _, fp = s.partition(".")
    ip = ip.lstrip("0")
    int_pat = re.findall(r"[0#?]", int_part.replace(",", ""))
    pad = ""
    if len(ip) < len(int_pat):
        # 足りない桁：0 は「0」、? は空白、# は何も出さない
        pad = "".join({"0": "0", "?": " ", "#": ""}[ch] for ch in int_pat[:len(int_pat) - len(ip)])
    if thousands and ip:
        ip = f"{int(ip):,}"
    ip = pad + ip
    # 小数部の末尾の0：# は消し、? は空白にする（桁をそろえるため）
    if frac_digits:
        frac_pat = re.findall(r"[0#?]", frac_part)
        digits = list(fp)
        k = len(digits) - 1
        while k >= 0 and digits[k] == "0" and frac_pat[k] in "#?":
            digits[k] = " " if frac_pat[k] == "?" else ""
            k -= 1
        fp = "".join(digits)
    # 整数部の桁の記号のあいだに字がある（「000-0000」の郵便番号など）：桁を右から1つずつ記号に当てる
    code_idx = [i for i, (k, t) in enumerate(toks) if k == "code"]
    dot = next((i for i in code_idx if toks[i][1] == "."), len(toks))
    int_ph = [i for i in code_idx if i < dot and toks[i][1] in "0#?"]
    frac_ph = [i for i in code_idx if i > dot and toks[i][1] in "0#?"]
    if frac_ph and any(not (toks[i][0] == "code" and toks[i][1] in "0#?") for i in range(frac_ph[0], frac_ph[-1] + 1)):
        raise Unsupported(f"小数部の桁のあいだに字のある表示形式: {sec}")
    if int_ph and any(not (toks[i][0] == "code" and toks[i][1] in "0#?,") for i in range(int_ph[0], int_ph[-1] + 1)):
        if thousands:
            raise Unsupported(f"桁区切りと、桁のあいだの字のある表示形式: {sec}")
        digits_ = ip[len(pad):] if pad and ip.startswith(pad) else ip.lstrip(" ")
        n = len(int_pat)
        assign = []
        for i in range(n):
            j = len(digits_) - (n - i)
            assign.append(digits_[j] if j >= 0 else {"0": "0", "?": " ", "#": ""}[int_pat[i]])
        if len(digits_) > n:
            assign[0] = digits_[:len(digits_) - n + 1]
        frac_text = ("." + fp if frac_digits and (fp or "." in codes) else ("." if "." in codes and not fp else ""))
        out, ii = [], 0
        for i, (k, t) in enumerate(toks):
            if k == "lit":
                out.append(t)
            elif i == dot:
                out.append(shape(frac_text) if shape is not None and frac_text != "." else frac_text)
            elif t in "0#?":
                if i < dot:
                    out.append(shape(assign[ii]) if shape is not None else assign[ii])
                    ii += 1
            elif t == "," or (t.startswith("[") and t.endswith("]")):
                continue
            else:
                out.append(t)
        res = "".join(out)
        return ("-" + res) if neg and float(d) != 0 else res
    number = ip + ("." + fp if frac_digits and (fp or "." in codes) else ("." if "." in codes and not fp else ""))
    if shape is not None:
        number = shape(number)
    # 記号の位置に数字を差し込み、前後の文字はそのまま残す
    out = []
    placed = False
    for k, t in toks:
        if k == "lit":
            out.append(t)
        elif t in "0#?,." or (t.startswith("[") and t.endswith("]")):
            if not placed and t in "0#?.":
                out.append(number)
                placed = True
        elif t == "%":
            out.append("%")
        else:
            out.append(t)
    if not placed and re.search(r"[0#?]", codes):
        out.insert(0, number)
    # 桁の記号のないセクション（0;-0;"ゼロ" の「"ゼロ"」など）は、字だけを出す（Excel も数を出さない）
    res = "".join(out)
    if neg and float(d) != 0:
        res = "-" + res
    return res


# 暦（CC）：グレゴリオ暦として年月日を出す番号（00 既定・01 各国語・02 米国の英語）。09〜0C もグレゴリオ暦だが、名前の言語が違う
_GREGORIAN = (0x00, 0x01, 0x02)
_GREGORIAN_OTHER_NAMES = (0x09, 0x0A, 0x0B, 0x0C)


def _format_date(x: float, sec: str, loc: Loc = Loc(), numerals: int = 0, db: int = 0, date1904: bool | None = None) -> str:
    if x < 0:
        raise Unsupported("負の日付")
    if re.search(r"(?i)am/pm", sec):
        sec = re.sub(r"(?i)am/pm", "\x01", sec)
    toks = _tokens(sec)
    has_ampm = any(t == "\x01" for k, t in toks if k == "code")
    cal = 6 if loc.hijri_prefix else loc.calendar
    if cal not in _GREGORIAN + _GREGORIAN_OTHER_NAMES + (0x03, 0x04, 0x06, 0x07):
        raise Unsupported(f"暦の番号 0x{cal:02X}（{loc.text}）")
    lang = loc.lang
    y, mo, d = serial_parts(x, date1904)
    frac = x - math.floor(x)
    # 秒の小数（ss.00）：小数の桁数で丸める
    codes = [t for k, t in toks if k == "code"]
    sub = 0
    for i, t in enumerate(codes):
        if t == "." and i + 1 < len(codes) and codes[i + 1] == "0":
            j = i + 1
            while j < len(codes) and codes[j] == "0":
                j += 1
            sub = j - i - 1
            break
    secs_total = Decimal(str(frac * 86400)).quantize(Decimal(1).scaleb(-sub), rounding=ROUND_HALF_UP)
    whole = int(secs_total)
    sub_digits = format(secs_total - whole, "f")[2:2 + sub].ljust(sub, "0") if sub else ""
    hh, rem = divmod(whole, 3600)
    mi, ss = divmod(rem, 60)
    try:
        date = serial_to_date(x, date1904) if (date1904 or int(x) not in (0, 60)) else None
    except ValueError:
        date = None
    sun0 = weekday_sun0(x, date1904)

    def num(n: int, width: int = 1) -> str:
        """日付・時刻の数（年月日・時分秒）を、数字の形・漢数字の指定どおりに字にする。"""
        if db:
            if width > 1:
                raise Unsupported(f"[DBNum{db}] の、桁をそろえる日付の記号（2桁）")
            return dbnum_units(n, db)
        return shape_digits(str(n).rjust(width, "0"), numerals)

    def calendar_ymd(part: int = 0) -> tuple[int, int, int]:
        """暦の年月日（y・m・d の記号に使う。part：0 年・1 月・2 日。和暦・民国の暦の月と日は、グレゴリオ暦と同じ）。"""
        if cal == 6:
            if date is None:
                raise Unsupported("ヒジュラ暦の、1900年3月1日より前の日付")
            return hijri_tabular(date.year, date.month, date.day)
        if cal in (0x03, 0x04) and part == 0:
            raise Unsupported(f"和暦・民国の暦の指定（{loc.text}）での年の記号 y")
        if cal == 0 and lang == "ar":
            # アラビア語（サウジアラビアなど）の既定の暦は、PC ごとにグレゴリオ暦かヒジュラ暦（ウンム・アル＝クラー暦）かが違う
            raise Unsupported(f"暦を指定しないアラビア語のロケール（{loc.text}）の年月日")
        if cal == 7:
            return y + 543, mo, d
        return y, mo, d

    def era() -> tuple[str, str, int, str]:
        """(元号, 略号, 年, 種類)：和暦（ja）か民国（zh-TW）。"""
        jp = cal == 0x03 or (cal in _GREGORIAN and lang in (None, "ja"))
        roc = cal == 0x04 or (cal in _GREGORIAN and lang == "zh-TW")
        if date is None:
            raise Unsupported("1900年3月1日より前の日付の元号")
        if jp:
            name, abbr, n = to_wareki(date)
            if not name:
                raise Unsupported("明治より前の日付の元号")
            return name, abbr, n, "jp"
        if roc:
            if date.year <= 1911:
                raise Unsupported("中華民国の暦より前の日付")
            return "", "", date.year - 1911, "roc"
        raise Unsupported(f"このロケール・暦（{loc.text or '指定なし'}）の元号の記号 g・e")

    def names(kind: str, with_day: bool = False) -> list[str]:
        if cal in _GREGORIAN_OTHER_NAMES or cal == 6:
            raise Unsupported(f"この暦（0x{cal:02X}）の月・曜日の名前")
        return name_table(lang, kind, with_day)

    # 記号を連続でまとめる
    groups: list[tuple[str, str]] = []
    for k, t in toks:
        if k == "code" and groups and groups[-1][0] == "code" and groups[-1][1][-1].lower() == t.lower() and t.isalpha():
            groups[-1] = ("code", groups[-1][1] + t)
        else:
            groups.append((k, t))
    has_day = any(k == "code" and t.lower() in ("d", "dd") for k, t in groups)
    out = []
    skip_zero = False
    for idx, (k, t) in enumerate(groups):
        if k == "lit":
            out.append(t)
            continue
        low = t.lower()
        if skip_zero and t == "0":
            continue
        skip_zero = False
        if t == "\x01":
            out.append("AM" if hh < 12 else "PM")
        elif t == "." and sub and idx + 1 < len(groups) and groups[idx + 1] == ("code", "0"):
            out.append("." + shape_digits(sub_digits, numerals) if not db else _raise(f"[DBNum{db}] の秒の小数"))
            skip_zero = True
        elif low.startswith("y"):
            cy = calendar_ymd(0)[0]
            if db:
                raise Unsupported(f"[DBNum{db}] の年の記号 y（西暦）")
            out.append(num(cy % 100, 2) if len(t) <= 2 else num(cy, 4))
        elif low.startswith("b") and len(t) in (2, 4):
            if date is None:
                raise Unsupported("1900年3月1日より前の日付の仏暦")
            by = date.year + 543  # 仏暦（タイ）：西暦 + 543
            if db:
                raise Unsupported(f"[DBNum{db}] の仏暦の記号 b")
            out.append(num(by % 100, 2) if len(t) == 2 else num(by, 4))
        elif low.startswith("m"):
            # 直前が h か直後が s なら「分」
            prev = next((g[1].lower() for g in reversed(groups[:idx]) if g[0] == "code" and g[1].strip() and g[1][0].isalpha()), "")
            nxt = next((g[1].lower() for g in groups[idx + 1:] if g[0] == "code" and g[1].strip() and g[1][0].isalpha()), "")
            if (prev.startswith("h") or nxt.startswith("s")) and len(t) <= 2:
                out.append(num(mi, 2) if len(t) >= 2 else num(mi))
            elif len(t) <= 2:
                cm = calendar_ymd(1)[1]
                out.append(num(cm, 2) if len(t) == 2 else num(cm))
            else:
                if cal not in _GREGORIAN + (0x03, 0x04, 0x07):
                    raise Unsupported(f"この暦（0x{cal:02X}）の月の名前")
                if len(t) == 5:
                    if lang not in (None, "en"):
                        raise Unsupported(f"ロケール {lang} の月の頭文字 mmmmm")
                    out.append(_EN_MONTHS[mo - 1][0])
                elif len(t) > 5:
                    raise Unsupported(f"月の記号 {t}")
                else:
                    out.append(names("mmm" if len(t) == 3 else "mmmm", has_day)[mo - 1])
        elif low.startswith("d"):
            if len(t) <= 2:
                cd = calendar_ymd(2)[2]
                out.append(num(cd, 2) if len(t) == 2 else num(cd))
            else:
                if int(x) == 0 and not date1904:
                    raise Unsupported("1900年1月0日の曜日")
                out.append(names("ddd" if len(t) == 3 else "dddd")[sun0])
        elif low.startswith("a") and len(t) >= 3:
            if int(x) == 0 and not date1904:
                raise Unsupported("1900年1月0日の曜日")
            out.append(names("aaa" if len(t) == 3 else "aaaa")[sun0])
        elif low.startswith("h"):
            h = hh % 12 or 12 if has_ampm else hh
            out.append(num(h, 2) if len(t) >= 2 else num(h))
        elif low.startswith("s"):
            out.append(num(ss, 2) if len(t) >= 2 else num(ss))
        elif low.startswith("g"):
            name, abbr, _n, kind = era()
            if kind != "jp":
                raise Unsupported(f"民国の元号の記号 g（{loc.text}）")
            out.append(name if len(t) >= 3 else (name[:1] if len(t) == 2 else abbr))
        elif low.startswith("e"):
            _name, _abbr, n, kind = era()
            if n == 1 and loc.gannen and kind == "jp":
                nxt_lit = next((g[1] for g in groups[idx + 1:]), "")
                if not nxt_lit.startswith("年"):
                    # 「元年」と続かない書き方（R元.5.1 など）で Excel がどう見せるかは確かめていない
                    raise Unsupported(f"「年」の続かない元年の表示（{loc.text}）")
                out.append("元")  # [$-ja-JP-x-gannen]：元号の最初の年は「元年」
            else:
                out.append(num(n, 2) if len(t) >= 2 else num(n))
        elif t.startswith("["):
            raise Unsupported(f"経過時間の表示形式: {sec}")
        elif t.isascii() and t.isalpha():
            raise Unsupported(f"日付の表示形式の記号「{t}」")
        else:
            out.append(t)
    return "".join(out)


_EN_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]


def _raise(msg: str):
    raise Unsupported(msg)


def number_display(x: float) -> str:
    return number_to_text(x)


def _literal_free(sec: str) -> str:
    """区画の、引用符の中と \ の次の字を除いた部分（@ が書式の記号として使われているかを見る）。"""
    out, i = [], 0
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            i = len(sec) if j < 0 else j + 1
            continue
        if ch == "\\":
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)
