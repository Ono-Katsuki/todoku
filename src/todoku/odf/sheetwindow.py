"""OpenDocument の表計算の窓：Excel の窓と同じ書き方（［B3］は入力欄、〈C5＝値〉は計算結果、☐☑はチェック）。

違いは：
- 数式は、ファイルに書かれた OpenFormula のまま出す（todoku が計算に使う Excel の書き方も添える）
- 書くと計算結果が変わる欄は、どの数式が変わるか、todoku が計算し直せるかを示す（計算し直せない数式は、書いた後、
  保存された計算結果が古いままになる）
- マクロは LibreOffice Basic のソースとそのきっかけ
"""

from __future__ import annotations

from ..formula.parser import RefNode, addr, iter_refs, parse_cell
from ..xlsx import window as xw
from .basic import macro_info
from .sheet import OdsWorkbook
from .xml import top


def dependents(wb: OdsWorkbook) -> tuple[dict, list]:
    """セル → それを（直接・間接に）使う数式のセル。計算できない数式（参照が分からないもの）の一覧も返す。"""
    direct: dict[tuple[str, int, int], set] = {}
    ranges: list[tuple[str, int, int, int, int, tuple]] = []
    unknown = []
    for sh in wb.sheets:
        for (r, c), cell in sh.cells.items():
            if not cell.has_formula:
                continue
            me = (sh.name, r, c)
            if cell.ast is None:
                unknown.append(me)
                continue
            for ref in iter_refs(cell.ast):
                if not isinstance(ref, RefNode):
                    continue
                s = ref.sheet or sh.name
                if (ref.r2 - ref.r1 + 1) * (ref.c2 - ref.c1 + 1) <= 2000:
                    for rr in range(ref.r1, ref.r2 + 1):
                        for cc in range(ref.c1, ref.c2 + 1):
                            direct.setdefault((s, rr, cc), set()).add(me)
                else:
                    ranges.append((s, ref.r1, ref.c1, ref.r2, ref.c2, me))

    def users(key):
        out = set(direct.get(key, ()))
        s, r, c = key
        for s2, r1, c1, r2, c2, me in ranges:
            if s2 == s and r1 <= r <= r2 and c1 <= c <= c2:
                out.add(me)
        return out

    def walk(key):
        seen = set()
        todo = [key]
        while todo:
            k = todo.pop()
            for u in users(k):
                if u not in seen:
                    seen.add(u)
                    todo.append(u)
        return seen

    return {"walk": walk}, unknown


def build_wb(wb: OdsWorkbook) -> dict:
    doc = xw.build(wb)
    doc["kind"] = "fods" if wb.pkg.flat else "ods"
    doc["kind_ja"] = "OpenDocument の表計算（.fods・1つの XML）" if wb.pkg.flat else "OpenDocument の表計算（.ods）"
    doc["notation"]["〈C5＝値〉"] = "計算結果（数式。ファイルに保存された結果）"
    dep, unknown = dependents(wb)
    feeds_total = 0
    for sh_doc, sh in zip(doc["sheets"], wb.sheets):
        for item in sh_doc.get("cells", []):
            r, c = parse_cell(item["cell"])
            f = wb.odf_formula.get((sh.name, r, c))
            if f:
                cell = sh.cells.get((r, c))
                item["formula"] = f
                if cell is not None and cell.ast is not None:
                    item["formula_excel"] = "=" + cell.formula
                elif cell is not None:
                    item["formula_unsupported"] = cell.parse_error or "計算できない書き方"
        for f in sh_doc.get("fields", []):
            if not f.get("cell"):
                continue
            r, c = parse_cell(f["cell"])
            users = dep["walk"]((sh.name, r, c))
            if not users:
                continue
            feeds_total += 1
            cells = sorted(users)
            bad = [u for u in cells if wb.sheet(u[0]).cells[(u[1], u[2])].ast is None]
            names = "、".join(f"{s}!{addr(rr, cc)}" if s != sh.name else addr(rr, cc) for s, rr, cc in cells[:6]) + \
                ("ほか" if len(cells) > 6 else "")
            note = f"書くと、数式 {names} の計算結果が変わります（todoku が計算し直して書きます）"
            if bad:
                note = (f"書くと、数式 {names} の計算結果が変わります。そのうち "
                        + "、".join(addr(rr, cc) for _s, rr, cc in bad[:6])
                        + " は todoku では計算し直せないので、保存された計算結果が古いままになります")
            f["feeds_formulas"] = [f"{s}!{addr(rr, cc)}" for s, rr, cc in cells]
            f["write_as"] = (f.get("write_as", "") + "。" if f.get("write_as") else "") + note
    doc["feeds_total"] = feeds_total
    doc["unknown_formulas"] = [f"{s}!{addr(r, c)}" for s, r, c in unknown]
    doc["unread"] = [*wb.unexpanded, *(doc.get("unread") or [])]
    doc["macros"] = macro_info(wb.pkg, top(wb.pkg.tree(wb.part)))
    return doc


def build(blob: bytes, name: str) -> dict:
    return build_wb(OdsWorkbook(blob, name))


def markdown_parts(doc: dict) -> list[tuple[str, str]]:
    parts = xw.markdown_parts(doc)
    extra = [f"- 形式: {doc['kind_ja']}。数式は、ファイルに書かれた OpenFormula（of:=…）のまま JSON の窓に出しています"
             "（〈 〉の値は、ファイルに保存された計算結果です）。"]
    if doc.get("unknown_formulas"):
        extra.append(f"- todoku では計算し直せない数式が {len(doc['unknown_formulas'])} 個あります（"
                     + "、".join(doc["unknown_formulas"][:8]) + ("ほか" if len(doc["unknown_formulas"]) > 8 else "")
                     + "）。この数式の元になるセルに書くと、保存された計算結果が古いままになります。")
    out = []
    for i, (title, text) in enumerate(parts):
        lines = text.split("\n")
        if len(lines) > 2 and lines[2].startswith("- 元ファイルの SHA-256"):
            lines[2] += f"、{doc['kind_ja']}"
            if i == 0:
                lines[3:3] = extra
        out.append((title, "\n".join(lines)))
    feeds = [(sh["name"], f) for sh in doc["sheets"] for f in sh.get("fields", []) if f.get("feeds_formulas")]
    if feeds:
        sec = ["", "## 書くと計算結果が変わる欄", "",
               "次の欄に書くと、数式の計算結果が変わります。todoku は、計算し直せる数式は計算し直して、結果と見える字も書きます。", ""]
        sec += [f"- {f['id']}（{f.get('label') or '見出しなし'}）：{f['write_as'].split('。')[-1]}" for _s, f in feeds[:300]]
        if len(feeds) > 300:
            sec.append(f"- ほか {len(feeds) - 300} 欄")
        t, text = out[-1]
        out[-1] = (t, text + "\n".join(sec) + "\n")
    return out


def grade_sheets(blob: bytes, name: str):
    """Excel と同じ決め方（入力欄の印・規則では決められないセル・数式の照合）に、OpenDocument の表計算の事情を足す。"""
    from ..grade import Reason, _final, _macro_reasons, _write_macros, image_only_sheets, sheet_input_reasons
    from ..xlsx.calc import verify

    wb = OdsWorkbook(blob, name)
    doc = build_wb(wb)
    reasons: list = []
    write: list = []
    sheet_input_reasons(doc, reasons, write, protect="シートの保護")
    rep = verify(wb).summary()
    if rep["unsupported"]:
        reasons.append(Reason("C", "確かめられない数式", f"{rep['unsupported']} 個の数式は、規則どおりに計算し直して確かめられません"
                              "（その数式の元になるセルに書くと、保存された計算結果が古いままになります）", doc.get("unknown_formulas", [])))
    if rep["mismatch"]:
        reasons.append(Reason("C", "保存された値と合わない数式", f"{rep['mismatch']} 個のセルで、保存された計算結果が数式と一致しません"))
    if rep["unsupported"]:
        write.append(Reason("B", "確かめられない数式", f"{rep['unsupported']} 個の数式は計算し直せないので、書いた値から計算される所を確かめられません"))
    reasons += _macro_reasons(doc.get("macros"))
    write += _write_macros(doc.get("macros"))
    for u in doc.get("unread") or []:
        reasons.append(Reason("F", "読めない部分", u))
    image_only = set(image_only_sheets(wb))
    for sh_obj in wb.sheets:
        if sh_obj.name in image_only:
            reasons.append(Reason("E", "画像だけのシート", f"シート「{sh_obj.name}」の中身は画像で、セルの文字がありません", [sh_obj.name]))
            write.append(Reason("D", "画像だけのシート", f"シート「{sh_obj.name}」は画像だけで、書くセルがありません。手で書く紙として扱います",
                                [sh_obj.name]))
        ctl = [s for s in sh_obj.shapes if s.kind == "control"]
        if ctl:
            write.append(Reason("D", "フォームの部品", f"シート「{sh_obj.name}」のフォームの部品 {len(ctl)} 個には、いまは書けません"
                                  "（つながったセルがあれば、そのセルに書きます）", [s.name for s in ctl]))
    from ..formula.parser import addr
    from ..grade import picture_alerts

    pictures = []
    for sh_obj in wb.sheets:
        for n, s in enumerate((s for s in sh_obj.shapes if s.kind == "picture"), 1):
            at = f"の {addr(*s.anchor)} 付近" if s.anchor else ""
            pictures.append({"id": f"{sh_obj.name}!{s.name or f'図{n}'}", "alt": (s.alt or "").strip() or None,
                             "decorative": getattr(s, "decorative", False), "where": f"シート「{sh_obj.name}」{at}"})
    g = _final(reasons, write)
    g.alerts = picture_alerts(pictures)
    return g
