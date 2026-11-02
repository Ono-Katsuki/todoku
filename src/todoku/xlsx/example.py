"""記入例を「Excelが出した答え」として使い、記入と再計算の正しさを確かめる。

行政は空の様式と、同じ様式に記入した記入例を一緒に公開していることが多い。記入例はExcelで保存されて
いるので、数式セルにはExcelの計算結果が、チェックボックスにはExcelが書いた状態が残っている。
空の様式に記入例の入力値だけを書き込み、こちらで計算し直した値とチェックボックスの部品が、
記入例と一致するかを全数で比べる。ほかの表計算ソフトを正解にしない（ソフトごとに読み込みや計算の違いがあるため）。
"""

from __future__ import annotations

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
    updated_formulas: int = 0
    controls: int = 0  # 比べたチェックボックス
    control_mismatches: list[tuple[str, str, object, object]] = field(default_factory=list)
    parts_compared: int = 0  # 書き換えたチェックボックスの部品のうち、記入例と比べたもの
    parts_identical: int = 0
    unverified: list[str] = field(default_factory=list)
    proof: Proof | None = None

    @property
    def ok(self) -> bool:
        return (not self.mismatches and not self.control_mismatches and not self.not_input_fields
                and self.proof is not None and self.proof.ok)

    def headline(self) -> str:
        s = (f"記入例の入力 {self.inputs} か所を空の様式に記入 → 計算結果の更新 {self.updated_formulas} 個。"
             f"数式 {self.formulas:,} 個のうち、Excel（記入例）の計算結果と一致 {self.formulas - len(self.mismatches):,}・"
             f"不一致 {len(self.mismatches)}")
        if self.controls:
            s += (f"。チェックボックス {self.controls} 個の状態が一致 {self.controls - len(self.control_mismatches)}・"
                  f"不一致 {len(self.control_mismatches)}")
        if self.parts_compared:
            s += f"（書き換えた部品 {self.parts_identical}/{self.parts_compared} が記入例とバイト一致）"
        return s


def example_inputs(blank: Workbook, example: Workbook) -> tuple[dict[str, object], list[str], list[str]]:
    """記入例のうち、空の様式と値が違う「数式でないセル」と、リンクのないチェックボックスの状態を取り出す。

    返すもの: (書き込む値, 数式が違うセル, 様式の版の違い)。様式の版の違い（見出しの文字が違うなど）も、
    比べる条件をそろえるために書き込むが、入力の数には数えない。
    """
    from .window import input_cells

    values: dict[str, object] = {}
    structure: list[str] = []
    template: list[str] = []
    for sb in blank.sheets:
        try:
            se = example.sheet(sb.name)
        except KeyError:
            structure.append(sb.name)
            continue
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

    blank = Workbook(blank_blob)
    example = Workbook(example_blob)
    values, structure, template = example_inputs(blank, example)
    res = fill(blank, values, force=True)
    allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
    allowed_controls = {(p.sheet, p.address) for p in res.patches if p.kind == "control"}
    pr = prove(blank_blob, res.blob, allowed, allowed_controls)
    ours = Workbook(res.blob)
    templ = set(template)
    out = ExampleCheck(inputs=len(values) - len(templ), formulas=0, structure_differs=structure, template_differs=template,
                       updated_formulas=sum(p.kind == "formula" for p in res.patches),
                       unverified=res.unverified, proof=pr)
    # 記入例で書かれた欄が、窓で入力欄として示されているか（入力欄の見落とし）
    fields = {sh.name: input_cells(blank, sh) for sh in blank.sheets if not sh.state.startswith("non-")}
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
            out.not_input_fields.append(key)
    skip = set(structure)
    for se in example.sheets:
        if not ours.has_sheet(se.name):
            continue
        so = ours.sheet(se.name)
        for k, ce in se.cells.items():
            if not ce.has_formula or f"{se.name}!{addr(*k)}" in skip:
                continue
            out.formulas += 1
            co = so.cells.get(k)
            got = co.value if co else None
            if not values_equal(ce.value, got):
                out.mismatches.append((se.name, ce.address, ce.formula or "", ce.value, got))
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
