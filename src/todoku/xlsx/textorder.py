"""文字の並び順で結果が変わる数式を見つける。

MATCH（照合の種類を省く・1・-1）・VLOOKUP・HLOOKUP（近似一致）・LOOKUP は、表が並んでいる前提の二分探索で、文字どうしの比べ方は
Excel の動く環境（OS・言語）の順で決まる。ASCII でない字を含む表で文字を探すと、同じファイルが環境で別の行を選び、#N/A にもなる。
確かめた例（scripts/eval の dry run の niigata_490358.xlsx）：MATCH("処遇改善加算Ⅰロ", {Ⅰイ, Ⅰロ, Ⅱイ, Ⅱロ, Ⅲ, Ⅳ}) は、ファイルの保存された値では 2 番目、
Excel for Mac で計算し直すと 6 番目、「Ⅲ」は #N/A。どちらも「誤り」とは決められないので、事実として伝える。

見つけ方は数式の作りと、保存された値だけ：表の字に ASCII でない字があり、探す値が空でない文字（または数式の結果で分からない）こと。"""

from __future__ import annotations

from ..formula.parser import Binary, Bool, Func, Num, Postfix, RefNode, Unary, addr

LOOKUPS = ("MATCH", "VLOOKUP", "HLOOKUP", "LOOKUP")


def _approx_call(n: Func) -> bool:
    """定数で近似一致と分かる呼び出し。"""
    a = n.args
    if n.name == "LOOKUP":
        return len(a) >= 2
    if n.name == "MATCH":
        if len(a) < 2:
            return False
        return len(a) < 3 or (isinstance(a[2], Num) and float(a[2].value) != 0)
    if len(a) < 3:
        return False
    if len(a) < 4:
        return True
    f = a[3]
    return (isinstance(f, Bool) and bool(f.value)) or (isinstance(f, Num) and float(f.value) != 0)


def _keys_range(n: Func) -> RefNode | None:
    tbl = n.args[1]
    return tbl if isinstance(tbl, RefNode) and tbl.kind in ("range", "cell") else None


def _has_wide_text(wb, sheet, node: RefNode, fn: str, memo: dict) -> bool:
    k = (sheet, node.r1, node.c1, node.r2, node.c2, fn)
    if k in memo:
        return memo[k]
    r1, r2, c1, c2 = node.r1, node.r2, node.c1, node.c2
    if fn == "VLOOKUP":
        c2 = c1
    elif fn == "HLOOKUP":
        r2 = r1
    sh = wb.sheet(sheet)
    found = False
    if sh is not None:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = sh.cells.get((r, c))
                if cell is not None and isinstance(cell.value, str) and not cell.value.isascii():
                    found = True
                    break
            if found:
                break
    memo[k] = found
    return found


def _calls(ast):
    stack = [ast]
    while stack:
        n = stack.pop()
        if isinstance(n, Func):
            if n.name in LOOKUPS:
                yield n
            stack.extend(n.args)
        elif isinstance(n, Binary):
            stack.extend((n.left, n.right))
        elif isinstance(n, (Unary, Postfix)):
            stack.append(n.operand)


def _key_is_text(wb, sheet: str, key) -> bool:
    if isinstance(key, RefNode) and key.kind == "cell":
        sh = wb.sheet(key.sheet or sheet)
        cell = sh.cells.get((key.r1, key.c1)) if sh is not None else None
        v = cell.value if cell is not None else None
        return isinstance(v, str) and v != ""
    return not isinstance(key, Num)


def text_order_cells(wb, sh) -> list[dict]:
    """[{cell, function, table}]：文字の並び順で結果が変わる数式のセル。"""
    out = []
    memo: dict = {}
    for (r, c), cell in sorted(sh.cells.items()):
        f = cell.formula if cell.has_formula else None
        if not f or not any(x in f.upper() for x in LOOKUPS):
            continue
        try:
            ast = cell.ast
        except Exception:  # noqa: BLE001  読めない数式は「確かめられない数式」で伝える
            continue
        if ast is None:
            continue
        for n in _calls(ast):
            if not _approx_call(n):
                continue
            tbl = _keys_range(n)
            if tbl is None:
                continue
            tsheet = tbl.sheet or sh.name
            if _key_is_text(wb, sh.name, n.args[0]) and _has_wide_text(wb, tsheet, tbl, n.name, memo):
                out.append({"cell": addr(r, c), "function": n.name, "table": f"{tsheet}!{addr(tbl.r1, tbl.c1)}:{addr(tbl.r2, tbl.c2)}"})
                break
    return out
