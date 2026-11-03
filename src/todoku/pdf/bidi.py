"""右から左へ書く字（アラビア文字・ヘブライ文字など）の行を、見た目の順（左から右）から、読む順（論理順）に戻す。

PDF の字は、描いた位置（見た目）で並べると、右から左へ書く字は逆になる。Unicode の双方向アルゴリズム（UAX #9）の
レベルを、見た目の順の字に当て、L2 の並べ替え（いちばん高いレベルから、奇数のレベルまで、そのレベル以上の続きを逆にする）を
もう一度行うと、読む順に戻る（L2 の並べ替えは、レベルが字と一緒に動くので、2回で元に戻る）。

字の向きの種類は unicodedata.bidirectional（Unicode の文字データベース）で決める。数（EN・AN）とその間の区切り（ES・CS・ET）は、
右から左の中でも左から右のまま（レベル 2）。結合する記号（NSM）は、前の字と同じ扱い。

段落の向き（P2・P3：最初の強い字）：見た目の順では、読む順の最初の字は、右から左の段落なら右の端、左から右の段落なら左の端にある。
左の端の強い字が L で、右の端の強い字が R・AL のときは、どちらの段落としても筋が通るので、多いほうの向きにする。
"""

from __future__ import annotations

import unicodedata

_R = ("R", "AL")
_NUM = ("EN", "AN")


def bidi_class(text: str) -> str:
    """字（グリフの字。2字以上のこともある）の向きの種類：L・R・NUM・NSM・N（どちらでもない）。"""
    for ch in text:
        c = unicodedata.bidirectional(ch)
        if c == "NSM":
            continue
        if c == "L":
            return "L"
        if c in _R:
            return "R"
        if c in _NUM:
            return "NUM"
        if c in ("ES", "CS", "ET"):
            return c
        return "N"
    return "NSM" if text else "N"


def has_rtl(texts: list[str]) -> bool:
    return any(bidi_class(t) == "R" for t in texts)


def paragraph_rtl(classes: list[str]) -> bool:
    strong = [c for c in classes if c in ("L", "R")]
    if not strong:
        return False
    if strong[0] == "R":
        return True  # 左の端が右から左の字：右から左の段落（左から右の段落なら、左の端は読む順の最初の字）
    if strong[-1] == "L":
        return False
    return strong.count("R") > strong.count("L")


def levels(classes: list[str], rtl: bool) -> list[int]:
    """見た目の順の字の向きの種類 → 埋め込みのレベル（明示の埋め込みのない、1つの段落として）。"""
    n = len(classes)
    cls = list(classes)
    # NSM：前の字（なければ段落の向き）と同じ
    for i, c in enumerate(cls):
        if c == "NSM":
            cls[i] = cls[i - 1] if i else ("R" if rtl else "L")
    # W4・W5：数の間の区切り（1つ）と、数に接する ET は数
    for i, c in enumerate(cls):
        if c in ("ES", "CS") and 0 < i < n - 1 and cls[i - 1] == "NUM" and cls[i + 1] == "NUM":
            cls[i] = "NUM"
    for i, c in enumerate(cls):
        if c == "ET":
            j = i
            while j < n and cls[j] == "ET":
                j += 1
            if (i > 0 and cls[i - 1] == "NUM") or (j < n and cls[j] == "NUM"):
                for k in range(i, j):
                    cls[k] = "NUM"
    cls = ["N" if c in ("ES", "CS", "ET") else c for c in cls]
    # 数のレベル：左から右の段落では、見た目で左の強い字が L（なければ右の強い字が L か、どちらもない）なら L の続き
    nearest_left: list[str | None] = []
    cur = None
    for c in cls:
        nearest_left.append(cur)
        if c in ("L", "R"):
            cur = c
    nearest_right: list[str | None] = [None] * n
    cur = None
    for i in range(n - 1, -1, -1):
        nearest_right[i] = cur
        if cls[i] in ("L", "R"):
            cur = cls[i]
    num_l = [False] * n
    if not rtl:
        for i, c in enumerate(cls):
            if c == "NUM":
                lft, rgt = nearest_left[i], nearest_right[i]
                num_l[i] = lft == "L" or (lft is None and rgt in ("L", None))
    # N1・N2：どちらでもない字は、両側の強い字（数は R とみなす。L の続きの数は L）が同じ向きならその向き、違えば段落の向き
    def side(i: int) -> str:
        c = cls[i]
        if c == "NUM":
            return "L" if num_l[i] else "R"
        return c

    eff = [side(i) if cls[i] != "N" else "N" for i in range(n)]
    i = 0
    while i < n:
        if eff[i] != "N":
            i += 1
            continue
        j = i
        while j < n and eff[j] == "N":
            j += 1
        a = eff[i - 1] if i > 0 else ("R" if rtl else "L")
        b = eff[j] if j < n else ("R" if rtl else "L")
        d = a if a == b else ("R" if rtl else "L")
        for k in range(i, j):
            eff[k] = d
        i = j
    out = []
    for i in range(n):
        c = cls[i]
        if not rtl:
            if c == "NUM":
                out.append(0 if num_l[i] else 2)
            else:
                out.append(1 if eff[i] == "R" else 0)
        else:
            if c in ("NUM", "L"):
                out.append(2)
            else:
                out.append(2 if eff[i] == "L" else 1)
    return out


def reorder(units: list, lv: list[int]) -> list:
    """L2：いちばん高いレベルから、いちばん低い奇数のレベルまで、そのレベル以上の続きを逆にする。"""
    if not lv:
        return list(units)
    items = list(zip(units, lv))
    hi = max(lv)
    lo_odd = min((x for x in lv if x % 2), default=None)
    if lo_odd is None:
        return list(units)
    for level in range(hi, lo_odd - 1, -1):
        i = 0
        while i < len(items):
            if items[i][1] >= level:
                j = i
                while j < len(items) and items[j][1] >= level:
                    j += 1
                items[i:j] = items[i:j][::-1]
                i = j
            else:
                i += 1
    return [u for u, _ in items]


def visual_to_logical(texts: list[str]) -> tuple[list[int], bool]:
    """見た目の順（左から右）の字（グリフごと）→ 読む順の並び（もとの番号）と、段落が右から左か。"""
    classes = [bidi_class(t) if not t.isspace() else "N" for t in texts]
    rtl = paragraph_rtl(classes)
    lv = levels(classes, rtl)
    return reorder(list(range(len(texts))), lv), rtl
