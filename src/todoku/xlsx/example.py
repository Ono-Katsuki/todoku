"""記入例を「Excelが出した答え」として使い、記入と再計算の正しさを確かめる。

行政は空の様式と、同じ様式に記入した記入例を一緒に公開していることが多い。記入例はExcelで保存されて
いるので、数式セルにはExcelの計算結果が、チェックボックスにはExcelが書いた状態が残っている。
空の様式に記入例の入力値だけを書き込み、こちらで計算し直した値とチェックボックスの部品が、
記入例と一致するかを全数で比べる。ほかの表計算ソフトを正解にしない（ソフトごとに読み込みや計算の違いがあるため）。
"""

from __future__ import annotations

from ..book import open_book

from dataclasses import dataclass, field

from ..formula.parser import addr
from .calc import values_equal
from .edit import fill
from .proof import Proof, prove
from .workbook import Workbook


@dataclass
class ExampleCheck:
    inputs: int
    formulas: int
    mismatches: list[tuple[str, str, str, object, object]] = field(default_factory=list)
    structure_differs: list[str] = field(default_factory=list)  # 数式が違い、比較から外したセル
    template_differs: list[str] = field(default_factory=list)  # 見出しの文字など、様式の版の違い
    not_input_fields: list[str] = field(default_factory=list)  # 記入例で書かれているのに、窓が入力欄と示さなかったセル
    agent_fields: list[str] = field(default_factory=list)  # 記入例で書かれたセルのうち、窓が「規則では決められない」と示したセル
    updated_formulas: int = 0
    controls: int = 0  # 比べたチェックボックス
    control_mismatches: list[tuple[str, str, object, object]] = field(default_factory=list)
    parts_compared: int = 0  # 書き換えたチェックボックスの部品のうち、記入例と比べたもの
    parts_identical: int = 0
    unverified: list[str] = field(default_factory=list)
    proof: Proof | None = None
    # 数式の参照先に、数式が記入例と違うセル（structure_differs）がある：計算の条件が同じでないので比べない
    depends_on_structure: list[tuple[str, str, str, object, object]] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return (not self.mismatches and not self.control_mismatches and not self.not_input_fields
                and self.proof is not None and self.proof.ok)

    def headline(self) -> str:
        s = (f"記入例の入力 {self.inputs} か所を空の様式に記入 → 計算結果の更新 {self.updated_formulas} 個。"
             f"数式 {self.formulas:,} 個のうち、Excel（記入例）の計算結果と一致 {self.formulas - len(self.mismatches):,}・"
             f"不一致 {len(self.mismatches)}")
        if self.agent_fields:
            s += f"。記入例で書かれたうち {len(self.agent_fields)} か所は、窓が「規則では入力欄か決められない（エージェントが判断）」と示したセル"
        if self.depends_on_structure:
            s += f"（参照先の数式が記入例と違うので比べなかった数式 {len(self.depends_on_structure)} 個）"
        if self.controls:
            s += (f"。チェックボックス {self.controls} 個の状態が一致 {self.controls - len(self.control_mismatches)}・"
                  f"不一致 {len(self.control_mismatches)}")
        if self.parts_compared:
            s += f"（書き換えた部品 {self.parts_identical}/{self.parts_compared} が記入例とバイト一致）"
        return s


def pair_sheets(blank: Workbook, example: Workbook) -> dict[str, str]:
    """白紙のシート → 記入例のシート。同じ名前のシート、なければ見出しの文字が同じ位置に6割以上あるシート
    （記入例ではシートの名前を「記入例」「6-month budget」などに変えることがある）。"""
    out: dict[str, str] = {}
    names = {se.name for se in example.sheets}
    for sb in blank.sheets:
        if sb.name in names:
            out[sb.name] = sb.name
    used = set(out.values())
    for sb in blank.sheets:
        if sb.name in out:
            continue
        tb = {k: c.value for k, c in sb.cells.items() if isinstance(c.value, str) and c.value.strip()}
        if not tb:
            continue
        best, score = None, 0.0
        for se in example.sheets:
            if se.name in used:
                continue
            same = sum(1 for k, v in tb.items() if (ce := se.cells.get(k)) is not None and ce.value == v)
            if same / len(tb) > score:
                best, score = se.name, same / len(tb)
        if best is not None and score >= 0.6:
            out[sb.name] = best
            used.add(best)
    return out


def example_inputs(blank: Workbook, example: Workbook) -> tuple[dict[str, object], list[str], list[str]]:
    """記入例のうち、空の様式と値が違う「数式でないセル」と、リンクのないチェックボックスの状態を取り出す。

    返すもの: (書き込む値, 数式が違うセル, 様式の版の違い)。様式の版の違い（見出しの文字が違うなど）も、
    比べる条件をそろえるために書き込むが、入力の数には数えない。
    """
    from .window import input_cells

    values: dict[str, object] = {}
    structure: list[str] = []
    template: list[str] = []
    pairs = pair_sheets(blank, example)
    for sb in blank.sheets:
        if sb.name not in pairs:
            structure.append(sb.name)
            continue
        se = example.sheet(pairs[sb.name])
        fields = input_cells(blank, sb) if not sb.state.startswith("non-") else {}
        for k in sorted(set(sb.cells) | set(se.cells)):
            cb, ce = sb.cells.get(k), se.cells.get(k)
            fb = cb.formula if cb else None
            fe = ce.formula if ce else None
            if fb != fe:
                structure.append(f"{sb.name}!{addr(*k)}")
                continue
            if fe is not None or (ce is not None and ce.array_master) or (cb is not None and cb.array_master):
                continue
            vb = cb.value if cb else None
            ve = ce.value if ce else None
            if not (values_equal(vb, ve) and type(vb) is type(ve)):
                key = f"{sb.name}!{addr(*k)}"
                values[key] = ve
                locked = sb.protected and blank.style(sb.style_of(*k)).locked
                if (isinstance(vb, str) and vb.strip() and k not in fields) or (locked and k not in fields):
                    # 見出しの文字が違う、またはExcelでも書き込めないロックされたセル：様式の版の違い
                    template.append(key)
        # リンク先のセルがないチェックボックス（#REF! など）は、名前で状態を書き込む
        after = {c.shape_id: c for c in se.controls}
        for c in sb.controls:
            ce = after.get(c.shape_id)
            if ce is None or c.kind != "CheckBox" or ce.checked == c.checked:
                continue
            if not (c.link_cell and blank.has_sheet(c.link_cell[0])):
                values[f"{sb.name}!{c.name}"] = ce.is_checked
    return values, structure, template


def check(blank_blob: bytes, example_blob: bytes) -> ExampleCheck:
    from .window import input_cells

    blank = open_book(blank_blob)
    example = open_book(example_blob)
    values, structure, template = example_inputs(blank, example)
    res = fill(blank, values, force=True)
    allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
    allowed_controls = {(p.sheet, p.address) for p in res.patches if p.kind == "control"}
    pr = prove(blank_blob, res.blob, allowed, allowed_controls)
    ours = open_book(res.blob)
    templ = set(template)
    out = ExampleCheck(inputs=len(values) - len(templ), formulas=0, structure_differs=structure, template_differs=template,
                       updated_formulas=sum(p.kind == "formula" for p in res.patches),
                       unverified=res.unverified, proof=pr)
    # 記入例で書かれた欄が、窓で入力欄として示されているか（入力欄の見落とし）
    from .window import undecided_cells

    fields = {sh.name: input_cells(blank, sh) for sh in blank.sheets if not sh.state.startswith("non-")}
    undecided = {sh.name: undecided_cells(blank, sh, fields[sh.name]) for sh in blank.sheets if sh.name in fields}
    linked = {(blank.sheet(c.link_cell[0]).name, c.link_cell[1], c.link_cell[2])
              for sh in blank.sheets for c in sh.controls if c.link_cell and blank.has_sheet(c.link_cell[0])}
    for key in values:
        if key in templ:
            continue
        sheet, a = key.rsplit("!", 1)
        try:
            from ..formula.parser import parse_cell

            r, c = parse_cell(a)
        except ValueError:
            continue  # チェックボックスの名前
        if (r, c) not in fields.get(sheet, {}) and (sheet, r, c) not in linked:
            if (r, c) in undecided.get(sheet, {}):
                out.agent_fields.append(key)  # 窓は「規則では決められない（エージェントが判断）」と示している
            else:
                out.not_input_fields.append(key)
    skip = set(structure)
    for bname, ename in pair_sheets(blank, example).items():
        se = example.sheet(ename)
        if not ours.has_sheet(bname):
            continue
        so = ours.sheet(bname)
        for k, ce in se.cells.items():
            if not ce.has_formula or f"{bname}!{addr(*k)}" in skip:
                continue
            co = so.cells.get(k)
            got = co.value if co else None
            if not values_equal(ce.value, got):
                m = (bname, ce.address, ce.formula or "", ce.value, got)
                if co is not None and _depends_on(ours, bname, co, skip):
                    out.depends_on_structure.append(m)
                    continue
                out.mismatches.append(m)
            out.formulas += 1
        mine = {c.shape_id: c for c in so.controls}
        for ce in se.controls:
            co = mine.get(ce.shape_id)
            if co is None or ce.kind != "CheckBox":
                continue
            out.controls += 1
            if (co.checked or None) != (ce.checked or None):
                out.control_mismatches.append((se.name, ce.name, ce.checked, co.checked))
    for part in res.changed_parts:
        if (part.endswith(".vml") or "/ctrlProps/" in part) and example.zip.has(part):
            out.parts_compared += 1
            out.parts_identical += ours.zip.read(part) == example.zip.read(part)
    return out


def _depends_on(wb: Workbook, sheet: str, cell, targets: set[str], limit: int = 50_000) -> bool:
    """数式 cell が（数式をたどって）targets（「シート!A1」）のどれかを参照するか。"""
    from ..formula.parser import RefNode, iter_refs

    seen: set[tuple[str, int, int]] = set()
    stack = [(sheet, cell)]
    while stack:
        sh_name, c = stack.pop()
        if c.ast is None:
            continue
        for ref in iter_refs(c.ast):
            if not isinstance(ref, RefNode):
                continue
            name = ref.sheet or sh_name
            if not wb.has_sheet(name):
                continue
            sh = wb.sheet(name)
            r2, c2 = min(ref.r2, sh.max_row if hasattr(sh, "max_row") else ref.r2), ref.c2
            for (r, col), cc in sh.cells.items():
                if not (ref.r1 <= r <= r2 and ref.c1 <= col <= c2):
                    continue
                key = (sh.name, r, col)
                if key in seen:
                    continue
                seen.add(key)
                if f"{sh.name}!{addr(r, col)}" in targets:
                    return True
                if cc.has_formula:
                    stack.append((sh.name, cc))
                if len(seen) > limit:
                    return True  # たどりきれない：比べない側にしておく
    return False
