"""表示形式（"#,##0" や "ggge年m月d日" など）で数値を文字列にする。

TEXT関数と、AIに渡す「見えている値」の両方で使う。再現できない形式は Unsupported を投げる。
"""

from __future__ import annotations

import math
import re
from functools import lru_cache
from decimal import ROUND_HALF_UP, Decimal

from .dates import serial_parts, serial_to_date, to_wareki
from .values import Unsupported, dec15, number_to_text

# 組み込みの表示形式ID（ECMA-376 18.8.30 と日本語版の既定）
BUILTIN = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "yyyy/m/d", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "yyyy/m/d h:mm",
    27: '[$-411]ge.m.d', 28: '[$-411]ggge"年"m"月"d"日"', 29: '[$-411]ggge"年"m"月"d"日"',
    30: "m/d/yy", 31: 'yyyy"年"m"月"d"日"', 32: 'h"時"mm"分"', 33: 'h"時"mm"分"ss"秒"',
    34: 'yyyy"年"m"月"', 35: 'm"月"d"日"', 36: '[$-411]ge.m.d',
    37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
    50: '[$-411]ge.m.d', 51: '[$-411]ggge"年"m"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
    54: '[$-411]ggge"年"m"月"d"日"', 55: 'yyyy"年"m"月"', 56: 'm"月"d"日"', 57: '[$-411]ge.m.d',
    58: '[$-411]ggge"年"m"月"d"日"',
}

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
    return bool(re.search(r"(?i)[ydhsge]", core)) or bool(re.search(r"(?i)m", core) and not re.search(r"[0#?]", core))


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


def format_value(v, fmt: str) -> str:
    """セルの値を表示形式で文字列にする。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, str):
        secs = split_sections(fmt)
        if len(secs) >= 4:
            return _literal(secs[3]).replace("@", v) if "@" in secs[3] else _literal(secs[3])
        if "@" in fmt and len(secs) == 1:
            return _literal(fmt).replace("@", v)
        return v
    x = float(v)
    if fmt in ("General", "general", ""):
        return _general(x)
    secs = _split_sections(fmt)
    if len(secs) == 1 and "@" in _literal_free(secs[0]):
        # 文字列の区画（@）だけの表示形式：数は標準の表示（Excel で、数のセルに「文字列」の書式を付けたとき）
        return _general(x)
    sec = None
    neg_has_own = False
    # 条件付きセクション [>=100] などは未対応
    if _conditional(fmt):
        raise Unsupported(f"条件付きの表示形式: {fmt}")
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
    sec, kind = _clean_section(sec)
    if kind == "general":
        return _general(x)
    if kind == "date":
        return _format_date(x, sec)
    return _format_number(x, sec)


@lru_cache(maxsize=4096)
def _conditional(fmt: str) -> bool:
    return any(re.search(r"\[(<|>|=)", s) for s in _split_sections(fmt))


@lru_cache(maxsize=4096)
def _clean_section(sec: str) -> tuple[str, str]:
    """セクションから色・ロケール・漢数字の指定を除き、種類（general・date・number）を決める（同じセクションは控える）。"""
    sec = re.sub(r"(?i)\[(red|black|blue|green|white|yellow|magenta|cyan|color\s*\d+)\]", "", sec)
    sec = re.sub(r"\[\$-[^\]]*\]", "", sec)  # [$-411] [$-ja-JP] などのロケール指定
    sec = re.sub(r"\[\$([^\]-]+)(-[^\]]*)?\]", lambda m: '"' + m[1] + '"', sec)  # [$¥-411] は通貨記号
    sec = re.sub(r"\[DBNum\d\]", "", sec, flags=re.I)
    if sec.strip().lower() == "general":
        return sec, "general"
    return sec, "date" if is_date_format(sec) else "number"


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


def _format_number(x: float, sec: str) -> str:
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
    number = ip + ("." + fp if frac_digits and (fp or "." in codes) else ("." if "." in codes and not fp else ""))
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
    if not placed:
        out.insert(0, number)
    res = "".join(out)
    if neg and float(d) != 0:
        res = "-" + res
    return res


def _format_date(x: float, sec: str) -> str:
    if x < 0:
        raise Unsupported("負の日付")
    toks = _tokens(sec)
    has_ampm = any(t.upper() in ("AM/PM", "A/P") for k, t in toks if k == "code") or re.search(r"(?i)am/pm", sec)
    if re.search(r"(?i)am/pm", sec):
        sec2 = re.sub(r"(?i)am/pm", "\x01", sec)
        toks = _tokens(sec2)
    y, mo, d = serial_parts(x)
    frac = x - math.floor(x)
    secs_total = int(Decimal(str(frac * 86400)).quantize(Decimal(1), rounding=ROUND_HALF_UP))
    hh, rem = divmod(secs_total, 3600)
    mi, ss = divmod(rem, 60)
    try:
        date = serial_to_date(x) if int(x) not in (0, 60) else None
    except ValueError:
        date = None
    era_name, era_abbr, era_year = to_wareki(date) if date else ("", "", y)

    # 記号を連続でまとめる
    groups: list[tuple[str, str]] = []
    for k, t in toks:
        if k == "code" and groups and groups[-1][0] == "code" and groups[-1][1][-1].lower() == t.lower() and t.isalpha():
            groups[-1] = ("code", groups[-1][1] + t)
        else:
            groups.append((k, t))
    out = []
    for idx, (k, t) in enumerate(groups):
        if k == "lit":
            out.append(t)
            continue
        low = t.lower()
        if t == "\x01":
            out.append("AM" if hh < 12 else "PM")
        elif low.startswith("y"):
            out.append(f"{y % 100:02d}" if len(t) <= 2 else f"{y:04d}")
        elif low.startswith("m"):
            # 直前が h か直後が s なら「分」
            prev = next((g[1].lower() for g in reversed(groups[:idx]) if g[0] == "code" and g[1].strip() and g[1][0].isalpha()), "")
            nxt = next((g[1].lower() for g in groups[idx + 1:] if g[0] == "code" and g[1].strip() and g[1][0].isalpha()), "")
            if prev.startswith("h") or nxt.startswith("s"):
                out.append(f"{mi:02d}" if len(t) >= 2 else str(mi))
            elif len(t) == 1:
                out.append(str(mo))
            elif len(t) == 2:
                out.append(f"{mo:02d}")
            else:
                import calendar

                out.append(calendar.month_abbr[mo] if len(t) == 3 else calendar.month_name[mo])
        elif low.startswith("d"):
            if len(t) == 1:
                out.append(str(d))
            elif len(t) == 2:
                out.append(f"{d:02d}")
            else:
                wd = ["月", "火", "水", "木", "金", "土", "日"][date.weekday()] if date else ""
                out.append(wd)
        elif low.startswith("a") and len(t) >= 3:
            wd = ["月", "火", "水", "木", "金", "土", "日"][date.weekday()] if date else ""
            out.append(wd + ("曜日" if len(t) >= 4 else ""))
        elif low.startswith("h"):
            h = hh % 12 or 12 if has_ampm else hh
            out.append(f"{h:02d}" if len(t) >= 2 else str(h))
        elif low.startswith("s"):
            out.append(f"{ss:02d}" if len(t) >= 2 else str(ss))
        elif low.startswith("g"):
            out.append(era_name if len(t) >= 3 else (era_name[:1] if len(t) == 2 else era_abbr))
        elif low.startswith("e"):
            out.append(f"{era_year:02d}" if len(t) >= 2 else str(era_year))
        elif t.startswith("["):
            raise Unsupported(f"経過時間の表示形式: {sec}")
        else:
            out.append(t)
    return "".join(out)


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
