"""記号の書体（Wingdings・Symbol など：symfont）のセルの字。

セルの値（cell.value）は、ファイルに保存された字のまま：計算・マクロ・比べるのは、保存された字。画面・紙に見える字は、
その字を書体の対応表で読んだ字（Wingdings の「þ」は、印のついた箱）で、窓・欄・見出し・書き込みの指定はこちらで扱う。
字の数は変わらない（symfont.text）。書くときは、見える字を、その書体でその形になる保存の字に直す（encode_targets）。

記号の書体：セルの書式の書体（Style.sym）と、字ごとの書体（.xlsx の <rFont>、.xls の字ごとの書式の字体）。対応表のある書体
（symfont.known）か、文字セットが記号（SYMBOL_CHARSET＝2）の書体。"""

from __future__ import annotations

import copy
import html
import re

from .. import symfont
from ..formula.parser import addr

_RFONT = re.compile(rb'<(?:[\w.-]+:)?rFont\b[^>]*?\bval="([^"]*)"')
_CHARSET2 = re.compile(rb'<(?:[\w.-]+:)?charset\b[^>]*?\bval="2"')


def cell_font(wb, cell) -> str:
    """セルの書式の書体が記号の書体なら、その名前。ほかは ""。"""
    return getattr(wb.style(cell.style), "sym", "") or ""


def _runs(wb, cell) -> list | None:
    """字ごとの書体 [(字の数, 書体)]（書体：記号の書体の名前・""＝記号の書体でない・None＝セルの書体）。なければ None。"""
    fn = getattr(wb, "sym_runs", None)
    return fn(cell) if fn is not None else None


def char_fonts(wb, cell) -> list[str] | None:
    """文字のセルの、1字ごとの記号の書体（記号の書体でない字は ""）。記号の書体の字のないセル・文字でないセルは None。"""
    v = cell.value
    if not isinstance(v, str) or not v:
        return None
    f = cell_font(wb, cell)
    runs = None if cell.has_formula else _runs(wb, cell)
    out: list[str] = []
    for n, font in runs or ():
        out += [f if font is None else font] * n
    if runs is None or len(out) != len(v):
        return [f] * len(v) if f else None
    return out if any(out) else None


def _read(s: str, f: str) -> str:
    return symfont.text(s, f, None, (symfont.key(f),)) if f else s


def decode(s: str, fonts: list[str]) -> str:
    """保存された字 s を、1字ごとの書体 fonts で読んだ、見える字。"""
    return "".join(_read(ch, f) for ch, f in zip(s, fonts))


def seen_cell(wb, cell):
    """セルの写しで、値（文字）を見える字にしたもの（raw：元のセル）。記号の書体の字のないセルは、そのセルのまま。"""
    if cell is None or not getattr(wb, "has_sym", False):
        return cell
    fonts = char_fonts(wb, cell)
    if fonts is None:
        return cell
    new = decode(cell.value, fonts)
    if new == cell.value:
        return cell
    c = copy.copy(cell)
    c.value = new
    c.raw = cell
    return c


def shown(wb, cell, out: str) -> str:
    """表示形式を通した字 out を、セルの書体で読んだ字。値そのままの字（文字のセル）は字ごとの書体で、ほか（数・日付・
    数式の結果・表示形式の足した字）はセルの書式の書体で読む。"""
    if not cell.has_formula and out == cell.value:
        fonts = char_fonts(wb, cell)
        return decode(out, fonts) if fonts is not None else out
    return _read(out, cell_font(wb, cell))


def prop_font(rpr: bytes | None, f: str) -> str:
    """字の書式（<rPr> のバイト）の決める記号の書体。書体の指定のない書式・書式のない字は、セルの書体 f。"""
    m = _RFONT.search(rpr) if rpr else None
    if m is None:
        return f
    name = html.unescape(m.group(1).decode("utf-8", "replace"))
    return name if name and (_CHARSET2.search(rpr) or symfont.known(name)) else ""


def _enc(ch: str, f: str, where: str, kept: list[str]) -> str:
    if not f:
        return ch
    code = symfont.code_of(f, ch)
    if code is not None:
        return chr(code)
    drawn = _read(ch, f)
    if drawn == ch:
        if not ch.isspace():
            kept.append(ch)
        return ch
    as_ = f"記号「{drawn}」" if not 0xE000 <= ord(drawn) <= 0xF8FF else f"記号（番号 0x{ord(drawn) - 0xF000:02X} の形）"
    raise ValueError(f"{where}: この字の書体は記号の書体（{f}）です。「{ch}」を書くと、その書体の{as_}の形で表示されるので、"
                     "書いていません。記号を書くなら、その記号の字（窓に出ている字）を指定してください。ふつうの字を書く所か、本人に確かめてください")


def _old_props(wb, s: str, r: int, c: int, raw_old: str, f: str):
    """字の一部を書き換えるセルの、(元の字ごとの書式, 書式 → 書体)。字ごとの書式のないセル・その書式を残して書けないセルは None
    （セルの書式の書体で読む）。.xlsx は <rPr>、.xls（BIFF8）は書体の番号。"""
    if getattr(wb, "kind", "xlsx") == "xls":
        isst = getattr(wb, "rich_cells", {}).get((s, r, c))
        if isst is None or isst not in wb.sst_runs or wb.shared_strings[isst] != raw_old:
            return None
        from .richtext import xls_runs

        return xls_runs(raw_old, wb.sst_runs[isst]), (lambda p: f if p is None else wb.font_sym(p))
    if not hasattr(wb, "zip"):
        return None
    from .richtext import cell_rst

    try:
        rst = cell_rst(wb, s, r, c)
    except (ValueError, KeyError):
        return None
    if not rst.rich or rst.text != raw_old:
        return None
    return rst.char_props(), (lambda p: prop_font(p, f))


def encode_targets(wb, targets: dict, partial: set, warnings: list[str]) -> dict:
    """書く値（見える字）を、保存する字に直す（targets を書き換える）。記号の書体の字にならない所は、そのまま。
    返すのは、保存する字が見える字と違うセルの、見える字（(シート, 行, 列) → 字：読み戻し・本人に伝える字に使う）。

    書き換えない字（元の見える字と同じ所）は、保存された字のまま。新しい字は、その所の書体（字の一部を書き換えるセルは、
    書式を残す決まり richtext.splice で決まる字ごとの書体。ほかはセルの書式の書体）で：その書体にその形の字があれば、その番号の字
    （Excel で打った字と同じ U+00xx）。その書体が読み替えない字（かな・漢字・U+2611 など）は、そのまま（ほかの書体で描かれる）。
    その書体では別の記号の形になる字（英字・数字など）は、書かずに断る（ValueError）。文字でない値（数・日付・論理値）は
    そのまま書き、どう見えるかを warnings に足す。"""
    seen: dict = {}
    if not getattr(wb, "has_sym", False):
        return seen
    from .richtext import _alignment, splice_props

    for (s, r, c), v in list(targets.items()):
        sh = wb.sheet(s)
        cell = sh.cells.get((r, c))
        f = getattr(wb.style(sh.style_of(r, c)), "sym", "") or ""
        runs = _runs(wb, cell) if cell is not None and not cell.has_formula else None
        if (not f and runs is None) or v is None:
            continue
        where = f"{s}!{addr(r, c)}"
        if not isinstance(v, str):
            if f:
                from ..formula.numfmt import format_value

                try:
                    as_ = "「" + _read(format_value(v, wb.style(sh.style_of(r, c)).num_fmt, getattr(wb, "date1904", False)), f) + "」"
                except Exception:  # noqa: BLE001  再現できない表示形式
                    as_ = "その書体の記号"
                warnings.append(f"{where}: このセルの書体は記号の書体（{f}）なので、書いた値は{as_}の形で表示されます"
                                "（値は書いたとおり）。記号そのものを書くなら、その記号の字を指定してください")
            continue
        raw_old = cell.value if cell is not None and isinstance(cell.value, str) and not cell.has_formula else ""
        new_fonts = [f] * len(v)
        seen_old = _read(raw_old, f)
        got = _old_props(wb, s, r, c, raw_old, f) if runs is not None and (s, r, c) in partial else None
        if got is not None:
            props, font_of = got
            seen_old = "".join(_read(ch, font_of(p)) for ch, p in zip(raw_old, props))
            try:
                new_fonts = [font_of(p) for p in splice_props(seen_old, props, v)[0]]
            except ValueError:
                continue  # 書式の違う字の境目への書き足し：書式を残す所（plan_inputs）が、同じ理由で断る
        if not any(new_fonts):
            continue
        kept: list[str] = []
        out: list[str] = []
        for op, i1, i2, j1, j2 in _alignment(seen_old, v):
            if op == "equal":
                out.append(raw_old[i1:i2])
            elif op != "delete":
                out.append("".join(_enc(ch, new_fonts[j], where, kept) for j, ch in zip(range(j1, j2), v[j1:j2])))
        stored = "".join(out)
        if got is not None and [font_of(p) for p in splice_props(raw_old, props, stored)[0]] != new_fonts:
            raise ValueError(f"{where}: 記号の書体の字とほかの書体の字のまざったセルで、書き換えた字の書体が一つに決まらないので、"
                             "書いていません（old・new で、書き換える字を短く指定してください）")
        if kept:
            names = "・".join(dict.fromkeys(x for x in new_fonts if x))
            warnings.append(f"{where}: 「{''.join(dict.fromkeys(kept))[:40]}」は、記号の書体（{names}）が記号に読み替えない字です。"
                            "ほかの書体で描かれます（LibreOffice で確かめた動きで、Excel では確かめていません）")
        if stored != v:
            seen[(s, r, c)] = v
            targets[(s, r, c)] = stored
    return seen


def sample_values(wb, sheet: str, field: dict, vals: list) -> list:
    """検査（audit）が、文字の中の所に書く試しの値：記号の書体の字の所は、その書体にある記号（いまの字と違うもの）にする。
    ほかの所は vals のまま。"""
    from ..formula.parser import parse_cell

    cell = wb.sheet(sheet).cells.get(parse_cell(field["cell"]))
    fonts = char_fonts(wb, cell) if cell is not None else None
    if not fonts:
        return vals
    now = seen_cell(wb, cell).value
    out = list(vals)
    for k, sp in enumerate(field.get("spans") or []):
        at = sp.get("at")
        if k >= len(out) or not isinstance(at, int):
            continue
        f = fonts[min(at, len(fonts) - 1)]
        if not f:
            continue
        for code in range(0x21, 0x7F):
            ch = _read(chr(code), f)
            if ch != chr(code) and symfont.code_of(f, ch) == code and ch != now[at:at + 1]:
                out[k] = ch
                break
    return out
