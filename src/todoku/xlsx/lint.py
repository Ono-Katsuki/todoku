"""書き出したファイルが、Excelで「修復」なしに開ける構造になっているかを確かめる。

ほかのソフトで開いて試すのではなく、Excelが開くときに求める条件（ECMA-376 の要素の順序、
セルの並び、型と値の対応、パッケージの関係）を直接検査する。
"""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from ..formula.parser import parse_cell
from ..formula.values import ERROR_CODES
from ..zipraw import RawZip

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# CT_Worksheet の子要素の順序
WORKSHEET_ORDER = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
    "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
    "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins",
    "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors",
    "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
    "webPublishItems", "tableParts", "extLst",
]
# CT_Workbook の子要素の順序
WORKBOOK_ORDER = [
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
    "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
    "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
]
CELL_ORDER = ["f", "v", "is", "extLst"]


@dataclass
class LintReport:
    errors: list[str] = field(default_factory=list)
    checked_parts: int = 0

    @property
    def ok(self) -> bool:
        return not self.errors


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def _check_order(children: list[str], order: list[str], where: str, errors: list[str]) -> None:
    pos = {n: i for i, n in enumerate(order)}
    last = -1
    for name in children:
        if name not in pos:
            continue  # 拡張要素（mc:AlternateContent など）は順序の対象外
        if pos[name] < last:
            errors.append(f"{where}: 要素 <{name}> の位置が規格の順序に合いません")
        last = max(last, pos[name])


def lint(blob: bytes, parts: list[str] | None = None) -> LintReport:
    """parts を指定すると、その部分だけを詳しく調べる（書き換えた部分の検査）。"""
    if blob[:4] == b"\xd0\xcf\x11\xe0":
        from ..xls.proof import lint as lint_xls

        r = lint_xls(blob)
        return LintReport(errors=r.errors, checked_parts=r.checked_parts)
    z = RawZip(blob)
    rep = LintReport()
    names = set(z.names())
    # ---------------------------------------------------- パッケージ
    ct = ET.fromstring(z.read("[Content_Types].xml"))
    defaults = {d.get("Extension", "").lower() for d in ct.iter(CT + "Default")}
    overrides = {o.get("PartName", "").lstrip("/") for o in ct.iter(CT + "Override")}
    for n in names:
        if n.endswith("/") or n == "[Content_Types].xml" or n.startswith("[trash]/"):
            continue  # [trash]/ はOfficeの差分保存の名残で、Excel自身が書き、読むときは無視する
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        if n not in overrides and ext not in defaults:
            rep.errors.append(f"{n}: [Content_Types].xml に種類が登録されていません")
    # 登録だけあって部品がないもの（docProps/app.xml など）は、Excelは読むときに無視するので問題にしない
    for n in names:
        if not n.endswith(".rels"):
            continue
        base_dir = posixpath.dirname(posixpath.dirname(n))
        for r in ET.fromstring(z.read(n)).iter(REL + "Relationship"):
            if r.get("TargetMode") == "External":
                continue
            t = r.get("Target", "")
            full = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base_dir, t))
            if full not in names and not t.startswith("#"):
                rep.errors.append(f"{n}: 参照先 {t} がありません")

    # ---------------------------------------------------- 中身
    targets = parts if parts is not None else [n for n in names if n.endswith(".xml") or n.endswith(".rels")]
    sst_count = None
    for n in names:
        if n.endswith("sharedStrings.xml"):
            sst_count = len(ET.fromstring(z.read(n)).findall(M + "si"))
    for n in sorted(targets):
        if n not in names:
            continue
        rep.checked_parts += 1
        raw = z.read(n)
        try:
            from .workbook import parse_xml

            root = parse_xml(raw)
        except ET.ParseError as e:
            rep.errors.append(f"{n}: XMLとして読めません（{e}）")
            continue
        tag = _local(root.tag)
        if tag == "worksheet":
            _check_order([_local(c.tag) for c in root], WORKSHEET_ORDER, n, rep.errors)
            _lint_sheet(n, root, sst_count, rep.errors)
        elif tag == "workbook":
            _check_order([_local(c.tag) for c in root], WORKBOOK_ORDER, n, rep.errors)
    return rep


def _lint_sheet(name: str, root: ET.Element, sst_count: int | None, errors: list[str]) -> None:
    sd = root.find(M + "sheetData")
    if sd is None:
        errors.append(f"{name}: sheetData がありません")
        return
    last_row = 0
    for row in sd:
        if _local(row.tag) != "row":
            continue
        r = row.get("r")
        if r is None:
            last_row += 1
            rn = last_row
        else:
            rn = int(r)
            if rn <= last_row:
                errors.append(f"{name}: 行 {rn} が順番どおりに並んでいません（直前は {last_row}）")
            last_row = rn
        last_col = 0
        for c in row:
            if _local(c.tag) != "c":
                continue
            ref = c.get("r")
            if ref:
                cr, cc = parse_cell(ref)
                if cr != rn:
                    errors.append(f"{name}: セル {ref} が行 {rn} の中にあります")
                if cc <= last_col:
                    errors.append(f"{name}: セル {ref} が列の順番どおりに並んでいません")
                last_col = cc
            else:
                last_col += 1
            _check_order([_local(x.tag) for x in c], CELL_ORDER, f"{name} {ref}", errors)
            t = c.get("t", "n")
            v = c.find(M + "v")
            vt = v.text if v is not None else None
            if t == "s":
                if vt is None or not vt.isdigit():
                    errors.append(f"{name} {ref}: 共有文字列の番号がありません")
                elif sst_count is not None and int(vt) >= sst_count:
                    errors.append(f"{name} {ref}: 共有文字列の番号 {vt} が範囲外です")
            elif t == "b":
                if vt not in (None, "0", "1"):
                    errors.append(f"{name} {ref}: 論理値が 0/1 ではありません（{vt}）")
            elif t == "e":
                if vt is not None and vt not in ERROR_CODES:
                    errors.append(f"{name} {ref}: 不明なエラー値 {vt}")
            elif t == "inlineStr":
                if c.find(M + "is") is None:
                    errors.append(f"{name} {ref}: inlineStr なのに文字列がありません")
            elif t == "n":
                if vt not in (None, ""):
                    try:
                        float(vt)
                    except ValueError:
                        errors.append(f"{name} {ref}: 数値のセルに数値でない値 {vt!r}")
            elif t not in ("str", "d"):
                errors.append(f"{name} {ref}: 不明な型 t={t}")
            if t != "inlineStr" and c.find(M + "is") is not None:
                errors.append(f"{name} {ref}: 文字列(is)があるのに型が inlineStr ではありません")
    if root.find(M + "mergeCells") is not None:
        for m in root.find(M + "mergeCells"):
            ref = m.get("ref", "")
            if not re.fullmatch(r"[A-Z]+\d+(:[A-Z]+\d+)?", ref):  # 1つのセルだけの結合も、Excelは受け付ける
                errors.append(f"{name}: 結合の範囲 {ref} が不正です")
