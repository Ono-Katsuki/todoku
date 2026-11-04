"""VBA の組み込みの関数と定数（Microsoft の VBA 言語リファレンスの規則どおり）。

今の時刻・日付（Now・Date・Time・Timer）は、計算のたびに変わる値なので、使ったことを記録する（結果は揮発性）。
人の答えが要るもの（MsgBox の「はい／いいえ」・InputBox）は、渡された答えを順に使い、なければ NeedsInput。
"""

from __future__ import annotations

import datetime as dt
import math
import unicodedata

from . import parse as P
from .interp import EMPTY, MISSING, NOTHING, NULL, Collection, Dictionary, NeedsInput, Unsupported, VArray, VbaError, VDate, compare, is_numeric_str, num_result, parse_date_text, serial_from, to_bool, to_datetime, to_double, to_long, to_str

CONSTANTS = {
    "vbcrlf": "\r\n", "vbnewline": "\r\n", "vbcr": "\r", "vblf": "\n", "vbtab": "\t", "vbnullstring": "", "vbnullchar": "\0",
    "vbok": 1, "vbcancel": 2, "vbabort": 3, "vbretry": 4, "vbignore": 5, "vbyes": 6, "vbno": 7,
    "vbokonly": 0, "vbokcancel": 1, "vbabortretryignore": 2, "vbyesnocancel": 3, "vbyesno": 4, "vbretrycancel": 5,
    "vbcritical": 16, "vbquestion": 32, "vbexclamation": 48, "vbinformation": 64, "vbdefaultbutton1": 0,
    "vbdefaultbutton2": 256, "vbdefaultbutton3": 512, "vbsystemmodal": 4096,
    "vbbinarycompare": 0, "vbtextcompare": 1,
    "vbuppercase": 1, "vblowercase": 2, "vbpropercase": 3, "vbwide": 4, "vbnarrow": 8, "vbkatakana": 16, "vbhiragana": 32,
    "vbunicode": 64, "vbfromunicode": 128,
    "vbsunday": 1, "vbmonday": 2, "vbtuesday": 3, "vbwednesday": 4, "vbthursday": 5, "vbfriday": 6, "vbsaturday": 7,
    "vbusesystemdayofweek": 0, "vbfirstjan1": 1,
    "vbempty": 0, "vbnull": 1, "vbinteger": 2, "vblong": 3, "vbsingle": 4, "vbdouble": 5, "vbcurrency": 6, "vbdate": 7,
    "vbstring": 8, "vbobject": 9, "vberror": 10, "vbboolean": 11, "vbvariant": 12, "vbarray": 8192,
    "xlup": -4162, "xldown": -4121, "xltoleft": -4159, "xltoright": 4161,
    "xlcalculationmanual": -4135, "xlcalculationautomatic": -4105, "xlcalculationsemiautomatic": 2,
    "xlpastevalues": -4163, "xlpasteall": -4104, "xlpasteformats": -4122, "xlpasteformulas": -4123,
    "xlpastespecialoperationnone": -4142, "xlnone": -4142, "xlautomatic": -4105,
    "xlwhole": 1, "xlpart": 2, "xlvalues": -4163, "xlformulas": -4123, "xlbyrows": 1, "xlbycolumns": 2,
    "xlnext": 1, "xlprevious": 2, "xlcelltypelastcell": 11, "xlcelltypeconstants": 2, "xlcelltypeformulas": -4123,
    "xlcelltypeblanks": 4, "xlcelltypevisible": 12, "xlsheetvisible": -1, "xlsheethidden": 0, "xlsheetveryhidden": 2,
    "xlshiftup": -4162, "xlshifttoleft": -4159, "xlshiftdown": -4121, "xlshifttoright": -4161,
    "xlyes": 1, "xlno": 2, "xlguess": 0, "xlascending": 1, "xldescending": 2, "xlcsv": 6, "xlworkbookdefault": 51,
    "xlopenxmlworkbookmacroenabled": 52, "xlr1c1": -4150, "xla1": 1, "xlcenter": -4108, "xlleft": -4131, "xlright": -4152,
    "xlcontinuous": 1, "xlthin": 2, "xlmedium": -4138, "xlthick": 4, "xledgebottom": 9, "xledgetop": 8,
    "xledgeleft": 7, "xledgeright": 10, "xlinsidehorizontal": 12, "xlinsidevertical": 11,
}

NOARG = {"now", "date", "time", "timer", "doevents", "rnd", "beep", "err"}


def _arg(rt, args: list[P.Arg], i: int, fr, default=MISSING):
    if i >= len(args) or args[i].value is None:
        return default
    return rt.value(args[i].value, fr)


def _named(rt, args: list[P.Arg], name: str, fr, pos: int, default=MISSING):
    for a in args:
        if a.name == name:
            return rt.value(a.value, fr)
    pos_args = [a for a in args if a.name is None]
    if pos < len(pos_args) and pos_args[pos].value is not None:
        return rt.value(pos_args[pos].value, fr)
    return default


def call(rt, name: str, args: list[P.Arg], fr, node):
    f = FUNCS.get(name)
    if f is None:
        if name in NOARG:
            f = FUNCS[name]
        else:
            raise Unsupported(f"関数「{name}」", node.line, fr.module)
    return f(rt, args, fr, node)


FUNCS: dict = {}


def fn(*names):
    def deco(f):
        for n in names:
            FUNCS[n] = f
        return f

    return deco


def _vals(rt, args, fr):
    return [rt.value(a.value, fr) if a.value is not None else MISSING for a in args]


# ---------------------------------------------------------------------- 文字列
@fn("len")
def f_len(rt, args, fr, node):
    v = _arg(rt, args, 0, fr)
    if v is NULL:
        return NULL
    return len(to_str(v))


@fn("lenb")
def f_lenb(rt, args, fr, node):
    return len(to_str(_arg(rt, args, 0, fr)).encode("utf-16-le"))


@fn("left", "left$")
def f_left(rt, args, fr, node):
    s, n = to_str(_arg(rt, args, 0, fr)), to_long(_arg(rt, args, 1, fr))
    if n < 0:
        raise VbaError(5)
    return s[:n]


@fn("right", "right$")
def f_right(rt, args, fr, node):
    s, n = to_str(_arg(rt, args, 0, fr)), to_long(_arg(rt, args, 1, fr))
    if n < 0:
        raise VbaError(5)
    return s[len(s) - n:] if n else ""


@fn("mid", "mid$")
def f_mid(rt, args, fr, node):
    s = to_str(_arg(rt, args, 0, fr))
    start = to_long(_arg(rt, args, 1, fr))
    ln = _arg(rt, args, 2, fr)
    if start < 1:
        raise VbaError(5)
    if ln is MISSING:
        return s[start - 1:]
    n = to_long(ln)
    if n < 0:
        raise VbaError(5)
    return s[start - 1:start - 1 + n]


@fn("instr")
def f_instr(rt, args, fr, node):
    vals = _vals(rt, args, fr)
    if len(vals) >= 3 and not isinstance(vals[0], str):
        start, s, sub = to_long(vals[0]), to_str(vals[1]), to_str(vals[2])
        mode = vals[3] if len(vals) > 3 else MISSING
    else:
        start, s, sub = 1, to_str(vals[0]), to_str(vals[1])
        mode = vals[2] if len(vals) > 2 else MISSING
    if start < 1:
        raise VbaError(5)
    text = (mode is not MISSING and to_long(mode) == 1) or (mode is MISSING and rt._text_compare(fr))
    if text:
        s, sub = s.lower(), sub.lower()
    if sub == "":
        return start if start <= len(s) + 1 else 0
    return s.find(sub, start - 1) + 1


@fn("instrrev")
def f_instrrev(rt, args, fr, node):
    s, sub = to_str(_arg(rt, args, 0, fr)), to_str(_arg(rt, args, 1, fr))
    start = _arg(rt, args, 2, fr, -1)
    start = to_long(start) if start is not MISSING else -1
    hay = s if start == -1 else s[:start]
    return hay.rfind(sub) + 1


@fn("replace")
def f_replace(rt, args, fr, node):
    s, find, rep = to_str(_arg(rt, args, 0, fr)), to_str(_arg(rt, args, 1, fr)), to_str(_arg(rt, args, 2, fr))
    start = _arg(rt, args, 3, fr, 1)
    count = _arg(rt, args, 4, fr, -1)
    mode = _arg(rt, args, 5, fr, 0)
    start = to_long(start) if start is not MISSING else 1
    count = to_long(count) if count is not MISSING else -1
    s = s[start - 1:]
    if find == "":
        return s
    if mode is not MISSING and to_long(mode) == 1:
        import re

        return re.sub(re.escape(find), lambda m: rep, s, count=0 if count < 0 else count, flags=re.I)
    return s.replace(find, rep, count if count >= 0 else -1)


@fn("trim", "trim$")
def f_trim(rt, args, fr, node):
    return to_str(_arg(rt, args, 0, fr)).strip(" ")


@fn("ltrim", "ltrim$")
def f_ltrim(rt, args, fr, node):
    return to_str(_arg(rt, args, 0, fr)).lstrip(" ")


@fn("rtrim", "rtrim$")
def f_rtrim(rt, args, fr, node):
    return to_str(_arg(rt, args, 0, fr)).rstrip(" ")


@fn("ucase", "ucase$")
def f_ucase(rt, args, fr, node):
    return to_str(_arg(rt, args, 0, fr)).upper()


@fn("lcase", "lcase$")
def f_lcase(rt, args, fr, node):
    return to_str(_arg(rt, args, 0, fr)).lower()


@fn("space", "space$")
def f_space(rt, args, fr, node):
    n = to_long(_arg(rt, args, 0, fr))
    rt.make_string(n)
    return " " * n


@fn("string", "string$")
def f_string(rt, args, fr, node):
    n = to_long(_arg(rt, args, 0, fr))
    c = _arg(rt, args, 1, fr)
    ch = chr(to_long(c)) if not isinstance(c, str) else c[:1]
    rt.make_string(n)
    return ch * n


@fn("strreverse")
def f_strreverse(rt, args, fr, node):
    return to_str(_arg(rt, args, 0, fr))[::-1]


@fn("chr", "chr$", "chrw", "chrw$")
def f_chr(rt, args, fr, node):
    n = to_long(_arg(rt, args, 0, fr))
    if node.target.name.startswith("chrw") or n < 0x80:
        return chr(n & 0xFFFF)
    try:
        return bytes([n >> 8, n & 0xFF]).decode("cp932") if n > 0xFF else bytes([n]).decode("cp932")
    except UnicodeDecodeError:
        raise VbaError(5)


@fn("asc", "ascw")
def f_asc(rt, args, fr, node):
    s = to_str(_arg(rt, args, 0, fr))
    if not s:
        raise VbaError(5)
    if node.target.name == "ascw":
        return ord(s[0])
    b = s[0].encode("cp932", errors="replace")
    v = int.from_bytes(b, "big")
    return v - 0x10000 if len(b) == 2 and v > 0x7FFF else v


@fn("strconv")
def f_strconv(rt, args, fr, node):
    s = to_str(_arg(rt, args, 0, fr))
    conv = to_long(_arg(rt, args, 1, fr))
    if conv & 1:
        s = s.upper()
    if conv & 2:
        s = s.lower()
    if conv & 3 == 3:
        s = " ".join(w[:1].upper() + w[1:].lower() for w in s.split(" "))
    if conv & 4:  # vbWide
        s = _to_wide(s)
    if conv & 8:  # vbNarrow
        s = _to_narrow(s)
    if conv & 16:  # vbKatakana
        s = "".join(chr(ord(c) + 0x60) if "ぁ" <= c <= "ゖ" else c for c in s)
    if conv & 32:  # vbHiragana
        s = "".join(chr(ord(c) - 0x60) if "ァ" <= c <= "ヶ" else c for c in s)
    if conv & (64 | 128):
        raise Unsupported("StrConv の vbUnicode・vbFromUnicode（バイト列の変換）", node.line, fr.module)
    return s


def _to_wide(s: str) -> str:
    out = []
    for c in s:
        o = ord(c)
        if 0x21 <= o <= 0x7E:
            out.append(chr(o + 0xFEE0))
        elif c == " ":
            out.append("　")
        else:
            out.append(unicodedata.normalize("NFKC", c) if "｡" <= c <= "ﾟ" else c)
    # 半角カナの濁点の合成（ｶﾞ → ガ）は NFKC が行う
    return unicodedata.normalize("NFC", "".join(out))


def _to_narrow(s: str) -> str:
    from ..formula.functions import _asc_text  # ASC 関数と同じ変換

    return _asc_text(s)


@fn("strcomp")
def f_strcomp(rt, args, fr, node):
    a, b = to_str(_arg(rt, args, 0, fr)), to_str(_arg(rt, args, 1, fr))
    mode = _arg(rt, args, 2, fr, 0)
    text = mode is not MISSING and to_long(mode) == 1
    return compare(a, b, text)


@fn("split")
def f_split(rt, args, fr, node):
    s = to_str(_arg(rt, args, 0, fr))
    d = _arg(rt, args, 1, fr, " ")
    d = to_str(d) if d is not MISSING else " "
    limit = _arg(rt, args, 2, fr, -1)
    limit = to_long(limit) if limit is not MISSING else -1
    if s == "":
        return VArray([0], [-1])
    parts = s.split(d, limit - 1) if limit > 0 else s.split(d) if d else [s]
    return VArray([0], [len(parts) - 1], {(i,): p for i, p in enumerate(parts)})


@fn("join")
def f_join(rt, args, fr, node):
    a = _arg(rt, args, 0, fr)
    d = _arg(rt, args, 1, fr, " ")
    d = to_str(d) if d is not MISSING else " "
    if not isinstance(a, VArray):
        raise VbaError(13)
    return d.join(to_str(x) for x in a.flat())


@fn("format", "format$")
def f_format(rt, args, fr, node):
    v = _arg(rt, args, 0, fr)
    fmt = _arg(rt, args, 1, fr, "")
    fmt = to_str(fmt) if fmt is not MISSING else ""
    return vba_format(v, fmt, node, fr)


def vba_format(v, fmt: str, node=None, fr=None) -> str:
    if v is EMPTY or v is None:
        return ""
    if fmt == "":
        return to_str(v)
    low = fmt.lower()
    named = {"general number": None, "currency": "#,##0", "fixed": "0.00", "standard": "#,##0.00", "percent": "0.00%",
             "yes/no": None, "true/false": None, "on/off": None, "short date": "yyyy/mm/dd", "long date": "yyyy年m月d日",
             "short time": "hh:mm", "long time": "h:mm:ss"}
    if low in named:
        if low == "general number":
            return to_str(to_double(v))
        if low in ("yes/no", "true/false", "on/off"):
            b = to_bool(v)
            return {"yes/no": ("Yes", "No"), "true/false": ("True", "False"), "on/off": ("On", "Off")}[low][0 if b else 1]
        fmt = named[low]
    if any(t in low for t in ("yy", "mm", "dd", "hh", "nn", "ss", "m/", "/d", "aaa", "ggg", "ee")) or \
            (isinstance(v, VDate) and not any(ch in fmt for ch in "0#")):
        d = to_datetime(v if not isinstance(v, str) else (parse_date_text(v) or to_double(v)))
        return _format_date(d, fmt)
    from ..formula.numfmt import format_value

    try:
        x = to_double(v)
    except VbaError:
        return to_str(v)
    return format_value(x, fmt)


def _format_date(d: dt.datetime, fmt: str) -> str:
    import re

    youbi = "月火水木金土日"
    eras = [(dt.datetime(2019, 5, 1), "令和", "R"), (dt.datetime(1989, 1, 8), "平成", "H"), (dt.datetime(1926, 12, 25), "昭和", "S"),
            (dt.datetime(1912, 7, 30), "大正", "T"), (dt.datetime(1868, 1, 25), "明治", "M")]
    era = next((e for e in eras if d >= e[0]), None)
    tokens = re.findall(r'"[^"]*"|\\.|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|aaaa|aaa|ggg|gg|g|ee|e|AM/PM|am/pm|.', fmt)
    out = []
    prev_h = False
    for t in tokens:
        tl = t.lower()
        if t.startswith('"'):
            out.append(t[1:-1])
        elif t.startswith("\\"):
            out.append(t[1:])
        elif tl == "yyyy":
            out.append(f"{d.year:04d}")
        elif tl == "yy":
            out.append(f"{d.year % 100:02d}")
        elif tl in ("mm", "m"):
            if prev_h:
                out.append(f"{d.minute:02d}" if tl == "mm" else str(d.minute))
            else:
                out.append(f"{d.month:02d}" if tl == "mm" else str(d.month))
        elif tl == "mmmm":
            out.append(d.strftime("%B"))
        elif tl == "mmm":
            out.append(d.strftime("%b"))
        elif tl == "dd":
            out.append(f"{d.day:02d}")
        elif tl == "d":
            out.append(str(d.day))
        elif tl == "dddd":
            out.append(d.strftime("%A"))
        elif tl == "ddd":
            out.append(d.strftime("%a"))
        elif tl == "aaaa":
            out.append(youbi[d.weekday()] + "曜日")
        elif tl == "aaa":
            out.append(youbi[d.weekday()])
        elif tl in ("hh", "h"):
            out.append(f"{d.hour:02d}" if tl == "hh" else str(d.hour))
        elif tl in ("nn", "n"):
            out.append(f"{d.minute:02d}" if tl == "nn" else str(d.minute))
        elif tl in ("ss", "s"):
            out.append(f"{d.second:02d}" if tl == "ss" else str(d.second))
        elif tl in ("ggg", "gg"):
            out.append(era[1] if era else "")
        elif tl == "g":
            out.append(era[2] if era else "")
        elif tl in ("ee", "e"):
            n = d.year - era[0].year + 1 if era else d.year
            out.append(f"{n:02d}" if tl == "ee" else str(n))
        elif tl == "am/pm":
            out.append(("AM" if d.hour < 12 else "PM") if t == "AM/PM" else ("am" if d.hour < 12 else "pm"))
        else:
            out.append(t)
        if tl in ("hh", "h"):
            prev_h = True
        elif tl not in (":",) and t.strip():
            prev_h = False if tl not in ("mm", "m") else prev_h
    return "".join(out)


# ---------------------------------------------------------------------- 変換・数
@fn("cstr")
def f_cstr(rt, args, fr, node):
    return to_str(_arg(rt, args, 0, fr))


@fn("cint")
def f_cint(rt, args, fr, node):
    x = to_long(_arg(rt, args, 0, fr))
    if not -32768 <= x <= 32767:
        raise VbaError(6)
    return x


@fn("clng", "clnglng", "cbyte")
def f_clng(rt, args, fr, node):
    return to_long(_arg(rt, args, 0, fr))


@fn("cdbl", "csng", "ccur", "cdec")
def f_cdbl(rt, args, fr, node):
    return to_double(_arg(rt, args, 0, fr))


@fn("cbool")
def f_cbool(rt, args, fr, node):
    return to_bool(_arg(rt, args, 0, fr))


@fn("cdate", "datevalue")
def f_cdate(rt, args, fr, node):
    v = _arg(rt, args, 0, fr)
    if isinstance(v, str):
        d = parse_date_text(v)
        if d is None:
            raise VbaError(13)
        return VDate(math.floor(d)) if node.target.name == "datevalue" else d
    return VDate(to_double(v))


@fn("cvar")
def f_cvar(rt, args, fr, node):
    return _arg(rt, args, 0, fr)


@fn("cverr")
def f_cverr(rt, args, fr, node):
    from ..formula.values import ExcelError

    n = to_long(_arg(rt, args, 0, fr))
    code = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?", 2036: "#NUM!", 2042: "#N/A"}.get(n)
    if code is None:
        raise Unsupported("CVErr の番号", node.line, fr.module)
    return ExcelError(code)


@fn("val")
def f_val(rt, args, fr, node):
    import re

    s = to_str(_arg(rt, args, 0, fr)).replace(" ", "").replace("\t", "")
    m = re.match(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eEdD][+-]?\d+)?|&[Hh][0-9A-Fa-f]+|&[Oo]?[0-7]+", s)
    if not m:
        return 0
    t = m.group(0)
    if t[:2].lower() == "&h":
        return int(t[2:], 16)
    if t[:1] == "&":
        return int(t.lstrip("&Oo"), 8)
    x = float(t.replace("d", "e").replace("D", "e"))
    return num_result(x)


@fn("int")
def f_int(rt, args, fr, node):
    return num_result(float(math.floor(to_double(_arg(rt, args, 0, fr)))))


@fn("fix")
def f_fix(rt, args, fr, node):
    return num_result(float(math.trunc(to_double(_arg(rt, args, 0, fr)))))


@fn("round")
def f_round(rt, args, fr, node):
    x = to_double(_arg(rt, args, 0, fr))
    n = _arg(rt, args, 1, fr, 0)
    n = to_long(n) if n is not MISSING else 0
    from decimal import ROUND_HALF_EVEN, Decimal

    q = Decimal(1).scaleb(-n)
    return num_result(float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_EVEN)))  # VBA の Round は銀行型


@fn("abs")
def f_abs(rt, args, fr, node):
    return num_result(abs(to_double(_arg(rt, args, 0, fr))))


@fn("sgn")
def f_sgn(rt, args, fr, node):
    x = to_double(_arg(rt, args, 0, fr))
    return (x > 0) - (x < 0)


@fn("sqr")
def f_sqr(rt, args, fr, node):
    x = to_double(_arg(rt, args, 0, fr))
    if x < 0:
        raise VbaError(5)
    return math.sqrt(x)


@fn("exp")
def f_exp(rt, args, fr, node):
    return math.exp(to_double(_arg(rt, args, 0, fr)))


@fn("log")
def f_log(rt, args, fr, node):
    x = to_double(_arg(rt, args, 0, fr))
    if x <= 0:
        raise VbaError(5)
    return math.log(x)


@fn("hex", "hex$")
def f_hex(rt, args, fr, node):
    x = to_long(_arg(rt, args, 0, fr))
    return format(x & 0xFFFFFFFF if x < 0 else x, "X")


@fn("oct", "oct$")
def f_oct(rt, args, fr, node):
    x = to_long(_arg(rt, args, 0, fr))
    return format(x & 0xFFFFFFFF if x < 0 else x, "o")


@fn("rnd")
def f_rnd(rt, args, fr, node):
    raise Unsupported("Rnd（乱数）は、計算のたびに変わるので再現しません", node.line, fr.module)


# ---------------------------------------------------------------------- 日付
def _volatile(rt):
    rt.host.volatile = True


@fn("now")
def f_now(rt, args, fr, node):
    _volatile(rt)
    return serial_from(rt.host.now())


@fn("date", "date$")
def f_date(rt, args, fr, node):
    _volatile(rt)
    n = rt.host.now()
    return serial_from(dt.datetime(n.year, n.month, n.day))


@fn("time", "time$")
def f_time(rt, args, fr, node):
    _volatile(rt)
    n = rt.host.now()
    return VDate((n.hour * 3600 + n.minute * 60 + n.second) / 86400)


@fn("timer")
def f_timer(rt, args, fr, node):
    _volatile(rt)
    n = rt.host.now()
    return float(n.hour * 3600 + n.minute * 60 + n.second)


@fn("year")
def f_year(rt, args, fr, node):
    return to_datetime(_date_arg(rt, args, fr)).year


@fn("month")
def f_month(rt, args, fr, node):
    return to_datetime(_date_arg(rt, args, fr)).month


@fn("day")
def f_day(rt, args, fr, node):
    return to_datetime(_date_arg(rt, args, fr)).day


@fn("hour")
def f_hour(rt, args, fr, node):
    return to_datetime(_date_arg(rt, args, fr)).hour


@fn("minute")
def f_minute(rt, args, fr, node):
    return to_datetime(_date_arg(rt, args, fr)).minute


@fn("second")
def f_second(rt, args, fr, node):
    return to_datetime(_date_arg(rt, args, fr)).second


def _date_arg(rt, args, fr, i=0):
    v = _arg(rt, args, i, fr)
    if isinstance(v, str):
        d = parse_date_text(v)
        if d is None:
            raise VbaError(13)
        return d
    return v


@fn("weekday")
def f_weekday(rt, args, fr, node):
    d = to_datetime(_date_arg(rt, args, fr))
    first = _arg(rt, args, 1, fr, 1)
    first = to_long(first) if first is not MISSING else 1
    sunday_based = (d.weekday() + 1) % 7 + 1  # 日曜 = 1
    return (sunday_based - first) % 7 + 1


@fn("dateserial")
def f_dateserial(rt, args, fr, node):
    y, m, d = (to_long(_arg(rt, args, i, fr)) for i in range(3))
    if 0 <= y <= 29:
        y += 2000
    elif 30 <= y <= 99:
        y += 1900
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    base = dt.datetime(y, m, 1)
    return serial_from(base + dt.timedelta(days=d - 1))


@fn("timeserial")
def f_timeserial(rt, args, fr, node):
    h, m, s = (to_long(_arg(rt, args, i, fr)) for i in range(3))
    return VDate((h * 3600 + m * 60 + s) / 86400)


@fn("dateadd")
def f_dateadd(rt, args, fr, node):
    unit = to_str(_arg(rt, args, 0, fr)).lower()
    n = to_double(_arg(rt, args, 1, fr))
    d = to_datetime(_date_arg(rt, args, fr, 2))
    if unit in ("yyyy", "q", "m"):
        months = int(n) * {"yyyy": 12, "q": 3, "m": 1}[unit]
        y, mo = divmod(d.month - 1 + months, 12)
        y += d.year
        mo += 1
        import calendar

        day = min(d.day, calendar.monthrange(y, mo)[1])
        return serial_from(d.replace(year=y, month=mo, day=day))
    secs = {"d": 86400, "y": 86400, "w": 86400, "ww": 7 * 86400, "h": 3600, "n": 60, "s": 1}.get(unit)
    if secs is None:
        raise VbaError(5)
    return serial_from(d + dt.timedelta(seconds=int(n) * secs))


@fn("datediff")
def f_datediff(rt, args, fr, node):
    unit = to_str(_arg(rt, args, 0, fr)).lower()
    a = to_datetime(_date_arg(rt, args, fr, 1))
    b = to_datetime(_date_arg(rt, args, fr, 2))
    if unit == "yyyy":
        return b.year - a.year
    if unit == "q":
        return (b.year - a.year) * 4 + (b.month - 1) // 3 - (a.month - 1) // 3
    if unit == "m":
        return (b.year - a.year) * 12 + b.month - a.month
    if unit in ("d", "y"):
        return (b.date() - a.date()).days
    if unit == "w":
        return (b.date() - a.date()).days // 7
    if unit == "ww":
        sa = a.date() - dt.timedelta(days=(a.weekday() + 1) % 7)
        sb = b.date() - dt.timedelta(days=(b.weekday() + 1) % 7)
        return (sb - sa).days // 7
    if unit == "h":
        return int((b.replace(minute=0, second=0) - a.replace(minute=0, second=0)).total_seconds() // 3600)
    if unit == "n":
        return int((b.replace(second=0) - a.replace(second=0)).total_seconds() // 60)
    if unit == "s":
        return int((b - a).total_seconds())
    raise VbaError(5)


@fn("datepart")
def f_datepart(rt, args, fr, node):
    unit = to_str(_arg(rt, args, 0, fr)).lower()
    d = to_datetime(_date_arg(rt, args, fr, 1))
    return {"yyyy": d.year, "q": (d.month - 1) // 3 + 1, "m": d.month, "y": d.timetuple().tm_yday, "d": d.day,
            "w": (d.weekday() + 1) % 7 + 1, "h": d.hour, "n": d.minute, "s": d.second}.get(unit, 0)


@fn("isdate")
def f_isdate(rt, args, fr, node):
    v = _arg(rt, args, 0, fr)
    return isinstance(v, VDate) or (isinstance(v, str) and parse_date_text(v) is not None)


@fn("monthname")
def f_monthname(rt, args, fr, node):
    return f"{to_long(_arg(rt, args, 0, fr))}月"


@fn("weekdayname")
def f_weekdayname(rt, args, fr, node):
    n = to_long(_arg(rt, args, 0, fr))
    abbr = _arg(rt, args, 1, fr, False)
    name = "日月火水木金土"[(n - 1) % 7]
    return name if abbr is not MISSING and to_bool(abbr) else name + "曜日"


# ---------------------------------------------------------------------- 型・配列
@fn("isempty")
def f_isempty(rt, args, fr, node):
    return _arg(rt, args, 0, fr) is EMPTY


@fn("isnull")
def f_isnull(rt, args, fr, node):
    return _arg(rt, args, 0, fr) is NULL


@fn("ismissing")
def f_ismissing(rt, args, fr, node):
    v = rt.eval(args[0].value, fr) if args else MISSING
    return v is MISSING


@fn("isnumeric")
def f_isnumeric(rt, args, fr, node):
    v = _arg(rt, args, 0, fr)
    if isinstance(v, bool) or v is EMPTY:
        return True
    if isinstance(v, (int, float)):
        return True
    return isinstance(v, str) and is_numeric_str(v)


@fn("isarray")
def f_isarray(rt, args, fr, node):
    return isinstance(_arg(rt, args, 0, fr), VArray)


@fn("isobject")
def f_isobject(rt, args, fr, node):
    v = rt.eval(args[0].value, fr)
    return v is NOTHING or (not isinstance(v, (int, float, str, bool, VArray)) and v not in (EMPTY, NULL, MISSING))


@fn("iserror")
def f_iserror(rt, args, fr, node):
    from ..formula.values import ExcelError

    return isinstance(_arg(rt, args, 0, fr), ExcelError)


@fn("typename")
def f_typename(rt, args, fr, node):
    v = rt.eval(args[0].value, fr)
    if v is EMPTY:
        return "Empty"
    if v is NULL:
        return "Null"
    if v is NOTHING:
        return "Nothing"
    if isinstance(v, bool):
        return "Boolean"
    if isinstance(v, int):
        return "Long" if not -32768 <= v <= 32767 else "Integer"
    if isinstance(v, VDate):
        return "Date"
    if isinstance(v, float):
        return "Double"
    if isinstance(v, str):
        return "String"
    if isinstance(v, VArray):
        return "Variant()"
    if isinstance(v, Collection):
        return "Collection"
    if isinstance(v, Dictionary):
        return "Dictionary"
    return getattr(v, "type_name", type(v).__name__)


@fn("vartype")
def f_vartype(rt, args, fr, node):
    v = rt.eval(args[0].value, fr)
    if v is EMPTY:
        return 0
    if v is NULL:
        return 1
    if isinstance(v, bool):
        return 11
    if isinstance(v, int):
        return 3
    if isinstance(v, VDate):
        return 7
    if isinstance(v, float):
        return 5
    if isinstance(v, str):
        return 8
    if isinstance(v, VArray):
        return 8204
    return 9


@fn("array")
def f_array(rt, args, fr, node):
    vals = _vals(rt, args, fr)
    return VArray([0], [len(vals) - 1], {(i,): v for i, v in enumerate(vals)})


@fn("ubound", "lbound")
def f_bound(rt, args, fr, node):
    a = _arg(rt, args, 0, fr)
    dim = _arg(rt, args, 1, fr, 1)
    dim = to_long(dim) if dim is not MISSING else 1
    if not isinstance(a, VArray):
        raise VbaError(13)
    if not a.lbounds or not 1 <= dim <= len(a.lbounds):
        raise VbaError(9)
    return a.ubounds[dim - 1] if node.target.name == "ubound" else a.lbounds[dim - 1]


@fn("iif")
def f_iif(rt, args, fr, node):
    c = _arg(rt, args, 0, fr)
    t, f = _arg(rt, args, 1, fr), _arg(rt, args, 2, fr)  # VBA の IIf は、両方の式を計算する
    return t if to_bool(c) else f


@fn("choose")
def f_choose(rt, args, fr, node):
    i = to_long(_arg(rt, args, 0, fr))
    if not 1 <= i < len(args):
        return NULL
    return _arg(rt, args, i, fr)


@fn("switch")
def f_switch(rt, args, fr, node):
    for k in range(0, len(args) - 1, 2):
        if to_bool(_arg(rt, args, k, fr)):
            return _arg(rt, args, k + 1, fr)
    return NULL


@fn("nz")
def f_nz(rt, args, fr, node):
    v = _arg(rt, args, 0, fr)
    return v if v not in (NULL, EMPTY) else _arg(rt, args, 1, fr, "")


@fn("createobject")
def f_createobject(rt, args, fr, node):
    cls = to_str(_arg(rt, args, 0, fr)).lower()
    if cls == "scripting.dictionary":
        return Dictionary()
    return rt.host.create_object(cls, node, fr)


@fn("doevents")
def f_doevents(rt, args, fr, node):
    return 0


@fn("beep")
def f_beep(rt, args, fr, node):
    return EMPTY


@fn("environ", "environ$")
def f_environ(rt, args, fr, node):
    raise NeedsInput("environ", f"マクロがパソコンの設定（{to_str(_arg(rt, args, 0, fr))}）を読もうとしています", node.line)


@fn("dir", "dir$")
def f_dir(rt, args, fr, node):
    p = _arg(rt, args, 0, fr)
    if p is MISSING:
        return ""
    path = to_str(p)
    if path in rt.inputs or path in rt.files_out:
        return path.replace("\\", "/").rsplit("/", 1)[-1]
    raise NeedsInput("file", f"マクロがファイル「{path}」があるかを調べています", node.line)


@fn("eof")
def f_eof(rt, args, fr, node):
    return rt.eof(to_long(_arg(rt, args, 0, fr)))


@fn("freefile")
def f_freefile(rt, args, fr, node):
    n = 1
    while n in rt.open_files:
        n += 1
    return n


@fn("msgbox")
def f_msgbox(rt, args, fr, node):
    prompt = to_str(_named(rt, args, "prompt", fr, 0))
    buttons = _named(rt, args, "buttons", fr, 1, 0)
    b = to_long(buttons) if buttons is not MISSING else 0
    rt.messages.append(prompt)
    kind = b & 7
    if kind == 0:
        return 1  # OK だけ：押すのは OK
    if rt.answers:
        a = rt.answers.pop(0)
        return _answer_code(a)
    raise NeedsInput("msgbox", prompt, node.line)


def _answer_code(a) -> int:
    if isinstance(a, int):
        return a
    t = str(a).strip().lower()
    return {"ok": 1, "cancel": 2, "abort": 3, "retry": 4, "ignore": 5, "yes": 6, "no": 7, "はい": 6, "いいえ": 7,
            "キャンセル": 2}.get(t, 1)


@fn("inputbox")
def f_inputbox(rt, args, fr, node):
    prompt = to_str(_named(rt, args, "prompt", fr, 0))
    if rt.answers:
        return to_str(rt.answers.pop(0))
    raise NeedsInput("inputbox", prompt, node.line)


@fn("callbyname")
def f_callbyname(rt, args, fr, node):
    raise Unsupported("CallByName", node.line, fr.module)


@fn("application.run", "run")
def f_run(rt, args, fr, node):
    name = to_str(_arg(rt, args, 0, fr))
    rest = [_arg(rt, args, i, fr) for i in range(1, len(args))]
    return rt.run(name.rsplit("!", 1)[-1].strip("'"), rest)
