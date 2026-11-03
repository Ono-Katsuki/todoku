"""한글の段落の番号（문단 번호・개요 번호）の書式：numFormat（KS X 6101 の NumberType1/2）と「^1.」の形の書式の字。"""

from __future__ import annotations

_HANGUL_SYL = "가나다라마바사아자차카타파하"
_HANGUL_JAMO = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ"
_CIRCLED_SYL = "㉮㉯㉰㉱㉲㉳㉴㉵㉶㉷㉸㉹㉺㉻"
_CIRCLED_JAMO = "㉠㉡㉢㉣㉤㉥㉦㉧㉨㉩㉪㉫㉬㉭"
_IDEO = "一二三四五六七八九十"
_CIRCLED_IDEO = "㊀㊁㊂㊃㊄㊅㊆㊇㊈㊉"
_PHONETIC = ["일", "이", "삼", "사", "오", "육", "칠", "팔", "구", "십"]


def _roman(n: int) -> str:
    out = []
    for v, s in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"),
                 (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= v:
            out.append(s)
            n -= v
    return "".join(out)


def _cycle(chars, n: int) -> str:
    return chars[(n - 1) % len(chars)] if n >= 1 else str(n)


def _latin(n: int) -> str:
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out or "A"


def number(n: int, fmt: str) -> str:
    fmt = (fmt or "DIGIT").upper()
    if fmt == "DIGIT":
        return str(n)
    if fmt == "CIRCLED_DIGIT":
        return chr(0x2460 + n - 1) if 1 <= n <= 20 else str(n)
    if fmt == "ROMAN_CAPITAL":
        return _roman(n)
    if fmt == "ROMAN_SMALL":
        return _roman(n).lower()
    if fmt == "LATIN_CAPITAL":
        return _latin(n)
    if fmt == "LATIN_SMALL":
        return _latin(n).lower()
    if fmt == "CIRCLED_LATIN_CAPITAL":
        return chr(0x24B6 + n - 1) if 1 <= n <= 26 else _latin(n)
    if fmt == "CIRCLED_LATIN_SMALL":
        return chr(0x24D0 + n - 1) if 1 <= n <= 26 else _latin(n).lower()
    if fmt == "HANGUL_SYLLABLE":
        return _cycle(_HANGUL_SYL, n)
    if fmt == "CIRCLED_HANGUL_SYLLABLE":
        return _cycle(_CIRCLED_SYL, n)
    if fmt == "HANGUL_JAMO":
        return _cycle(_HANGUL_JAMO, n)
    if fmt == "CIRCLED_HANGUL_JAMO":
        return _cycle(_CIRCLED_JAMO, n)
    if fmt == "HANGUL_PHONETIC":
        return _PHONETIC[n - 1] if 1 <= n <= 10 else str(n)
    if fmt == "IDEOGRAPH":
        return _IDEO[n - 1] if 1 <= n <= 10 else str(n)
    if fmt == "CIRCLED_IDEOGRAPH":
        return _CIRCLED_IDEO[n - 1] if 1 <= n <= 10 else str(n)
    return str(n)


def label(pattern: str, counters: list[int], formats: dict[int, str]) -> str:
    """「^1.^2」の ^N を、N 番目の段の番号（その段の書式）にした字。"""
    out = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "^" and i + 1 < len(pattern) and pattern[i + 1].isdigit():
            lv = int(pattern[i + 1])
            n = counters[lv] if lv < len(counters) else 0
            out.append(number(max(n, 1), formats.get(lv, "DIGIT")))
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)
