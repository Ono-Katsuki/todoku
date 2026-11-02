"""OpenDocument の数式（OpenFormula、ODF 1.3 第4部）を、todoku の数式の部品（Excel の書き方）で計算できる形にする。

書き換えるのは、書き方の違う所だけ：
- 参照 [.A1]・[.$A$1:.B2]・[Sheet1.A1]・['シート 1'.A1] → A1・$A$1:B2・Sheet1!A1・'シート 1'!A1
- 関数の引数の区切り「;」→「,」、配列の定数 {1;2|3;4}（「;」は列、「|」は行）→ {1,2;3,4}
- 名前空間の頭（of:・oooc:）。msoxl:（Excel の書き方の数式）はそのまま
書き換えられない書き方（3次元の参照・ほかのファイルの参照・参照の結合「~」・共通部分「!」）は、計算できないものとして返す。
"""

from __future__ import annotations

import re

from ..formula.parser import col_to_num


class OdfFormulaError(ValueError):
    pass


_REF_PART = re.compile(r"^(\$?)(?:'((?:[^']|'')*)'|([^.\[\]':]*))\.(\$?[A-Za-z]{1,3}\$?\d+|\$?[A-Za-z]{1,3}|\$?\d+|#REF!)$")


def _sheet_q(name: str) -> str:
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def ref_part(t: str) -> tuple[str | None, str]:
    """「Sheet1.A1」「.A1」「$'シート'.$B$2」→ (シートの名前, 番地)。"""
    m = _REF_PART.match(t.strip())
    if not m:
        raise OdfFormulaError(f"参照を読めません：[{t}]")
    sheet = m[2].replace("''", "'") if m[2] is not None else m[3]
    return (sheet or None), m[4]


def ref(body: str) -> str:
    """角かっこの中（[ ] は除く）の参照を、Excel の書き方にする。"""
    if "#" in body.split(".")[0] and not body.startswith("."):
        raise OdfFormulaError("ほかのファイルの参照は、計算できません")
    parts = _split_colon(body)
    if len(parts) == 1:
        sh, a = ref_part(parts[0])
        return (_sheet_q(sh) + "!" if sh else "") + a
    if len(parts) != 2:
        raise OdfFormulaError(f"参照を読めません：[{body}]")
    s1, a1 = ref_part(parts[0])
    s2, a2 = ref_part(parts[1])
    if s2 and s1 and s1 != s2:
        raise OdfFormulaError("3次元の参照（シートをまたぐ範囲）は、計算できません")
    sh = s1 or s2
    return (_sheet_q(sh) + "!" if sh else "") + a1 + ":" + a2


def _split_colon(body: str) -> list[str]:
    """引用符の外の「:」で分ける。"""
    out, cur, q = [], [], False
    for ch in body:
        if ch == "'":
            q = not q
        if ch == ":" and not q:
            out.append("".join(cur))
            cur = []
            continue
        cur.append(ch)
    out.append("".join(cur))
    return out


def to_excel(src: str) -> str:
    """table:formula の値 → Excel の書き方の数式（先頭の「=」は付けない）。"""
    s = src.strip()
    m = re.match(r"^([A-Za-z][\w.-]*):", s)
    ns = ""
    if m and not s.startswith(("=", "[")):
        ns = m[1].lower()
        s = s[m.end():]
    if s.startswith("="):
        s = s[1:]
    if ns == "msoxl":
        return s
    out = []
    i = 0
    n = len(s)
    depth_brace = 0
    while i < n:
        ch = s[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(s[i:j + 1])
            i = j + 1
            continue
        if ch == "[":
            j = i + 1
            q = False
            while j < n and (s[j] != "]" or q):
                if s[j] == "'":
                    q = not q
                j += 1
            if j >= n:
                raise OdfFormulaError("参照の「]」がありません")
            out.append(ref(s[i + 1:j]))
            i = j + 1
            continue
        if ch == "#":
            em = re.match(r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A)", s[i:])
            if em:
                out.append(em[0])  # エラーの値
                i += len(em[0])
                continue
        if ch == "{":
            depth_brace += 1
            out.append("{")
            i += 1
            continue
        if ch == "}":
            depth_brace -= 1
            out.append("}")
            i += 1
            continue
        if ch == ";":
            out.append("," if depth_brace == 0 else ",")
            i += 1
            continue
        if ch == "|" and depth_brace > 0:
            out.append(";")
            i += 1
            continue
        if ch == "~":
            raise OdfFormulaError("参照の結合（~）は、計算できません")
        if ch == "!":
            raise OdfFormulaError("参照の共通部分（!）は、計算できません")
        if ch == "'":
            # 引用符で囲んだシートの名前（参照の外）は使わない書き方
            raise OdfFormulaError("角かっこの外の引用符は読めません")
        out.append(ch)
        i += 1
    return "".join(out)


def cell_address(text: str) -> tuple[str | None, int, int]:
    """「Sheet1.B3」「$Sheet1.$B$3」→ (シート, 行, 列)。"""
    sh, a = ref_part(text)
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", a)
    if not m:
        raise OdfFormulaError(f"セルの番地を読めません：{text}")
    return sh, int(m[2]), col_to_num(m[1])


def range_address(text: str) -> tuple[str | None, int, int, int, int]:
    """「Sheet1.A1:Sheet1.H40」「Sheet1.A1:.H40」→ (シート, 行1, 列1, 行2, 列2)。"""
    parts = _split_colon(text.strip())
    sh, r1, c1 = cell_address(parts[0])
    if len(parts) == 1:
        return sh, r1, c1, r1, c1
    p2 = parts[1]
    if not p2.startswith(".") and "." not in p2:
        p2 = "." + p2
    _sh2, r2, c2 = cell_address(p2)
    return sh, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
