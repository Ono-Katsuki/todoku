"""テスト用の架空の様式を作る（個人情報なし）。

openpyxl で作ったあと、とどくの計算エンジンで計算結果（<v>）を書き込む（todoku.xlsx.edit.bake）。
計算結果が正しいことは、tests/test_xlsx.py で手計算の値と照らして確かめる。
使い方: uv run python scripts/make_fixtures.py
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Protection, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures"


def kami_excel(path: Path) -> None:
    """方眼紙・セル結合・入力規則・名前付き範囲・端数処理を含む、架空の「請求明細」様式。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "明細書"
    thin = Side(style="thin")
    for col in range(1, 41):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = 2.6
    ws.merge_cells("A1:AN1")
    ws["A1"] = "介 護 給 付 費 明 細 書（架空の様式）"
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A3:H3"); ws["A3"] = "事 業 所 番 号"
    ws.merge_cells("I3:T3"); ws["I3"] = "1234567890"
    ws.merge_cells("A4:H4"); ws["A4"] = "地 域 区 分"
    ws.merge_cells("I4:T4"); ws["I4"] = "3級地"
    ws.merge_cells("U4:Z4"); ws["U4"] = "単　価"
    ws.merge_cells("AA4:AF4"); ws["AA4"] = "=VLOOKUP(I4,単価表!A2:B9,2,FALSE)"
    ws.merge_cells("A5:H5"); ws["A5"] = "サービス提供年月"
    ws.merge_cells("I5:T5"); ws["I5"] = "=DATE(2026,10,1)"
    ws["I5"].number_format = '[$-ja-JP]ggge"年"m"月"'

    headers = [("A7:H7", "サービス内容"), ("I7:N7", "単位数"), ("O7:R7", "回数"), ("S7:X7", "サービス単位数"), ("Y7:AF7", "摘要")]
    for rng, text in headers:
        ws.merge_cells(rng)
        ws[rng.split(":")[0]] = text
    rows = [("生活介護５", 663, 20), ("送迎加算Ⅰ", 21, 38), ("食事提供体制加算", 30, 18), ("", None, None), ("", None, None)]
    for i, (name, unit, times) in enumerate(rows):
        r = 8 + i
        ws.merge_cells(f"A{r}:H{r}"); ws[f"A{r}"] = name or None
        ws.merge_cells(f"I{r}:N{r}"); ws[f"I{r}"] = unit
        ws.merge_cells(f"O{r}:R{r}"); ws[f"O{r}"] = times
        ws.merge_cells(f"S{r}:X{r}"); ws[f"S{r}"] = f'=IF(OR(I{r}="",O{r}=""),"",I{r}*O{r})'
        ws.merge_cells(f"Y{r}:AF{r}")
        for c in ("A", "I", "O", "Y"):
            ws[f"{c}{r}"].protection = Protection(locked=False)

    ws.merge_cells("A14:R14"); ws["A14"] = "給付単位数"
    ws.merge_cells("S14:X14"); ws["S14"] = "=SUM(S8:S12)"
    ws.merge_cells("A15:R15"); ws["A15"] = "処遇改善加算（41.8%）"
    ws.merge_cells("S15:X15"); ws["S15"] = "=ROUND(S14*処遇改善率,0)"
    ws.merge_cells("A16:R16"); ws["A16"] = "単位数合計"
    ws.merge_cells("S16:X16"); ws["S16"] = "=S14+S15"
    ws.merge_cells("A17:R17"); ws["A17"] = "費用合計（円）"
    ws.merge_cells("S17:X17"); ws["S17"] = "=ROUNDDOWN(S16*AA4,0)"
    ws.merge_cells("A18:R18"); ws["A18"] = "給付率（%）"
    ws.merge_cells("S18:X18"); ws["S18"] = 90
    ws["S18"].protection = Protection(locked=False)
    ws.merge_cells("A19:R19"); ws["A19"] = "請求額（円）"
    ws.merge_cells("S19:X19"); ws["S19"] = "=ROUNDDOWN(S17*S18/100,0)"
    ws.merge_cells("A20:R20"); ws["A20"] = "利用者負担額（円）"
    ws.merge_cells("S20:X20"); ws["S20"] = "=MIN(S17-S19,上限月額)"
    ws.merge_cells("A21:R21"); ws["A21"] = "確認"
    ws.merge_cells("S21:AF21"); ws["S21"] = '=IF(S19+S20=S17,"OK","差額あり："&TEXT(S17-S19-S20,"#,##0")&"円")'
    ws.merge_cells("A22:R22"); ws["A22"] = "提供月の和暦"
    ws.merge_cells("S22:AF22"); ws["S22"] = '=TEXT(I5,"ggge年m月")'
    ws.merge_cells("A23:R23"); ws["A23"] = "加算の件数"
    ws.merge_cells("S23:X23"); ws["S23"] = '=COUNTIF(A8:A12,"*加算*")'
    ws.merge_cells("A24:R24"); ws["A24"] = "加算分の単位"
    ws.merge_cells("S24:X24"); ws["S24"] = '=SUMIF(A8:A12,"*加算*",S8:S12)'
    ws.merge_cells("A25:R25"); ws["A25"] = "1回あたり平均"
    ws.merge_cells("S25:X25"); ws["S25"] = "=IFERROR(ROUND(S14/SUM(O8:O12),2),0)"
    ws.merge_cells("A26:R26"); ws["A26"] = "印"
    for r in range(3, 27):
        for c in range(1, 33):
            ws.cell(r, c).border = Border(top=thin, bottom=thin, left=thin, right=thin)

    dv = DataValidation(type="list", formula1="=単価表!$A$2:$A$9", allow_blank=True,
                        showErrorMessage=True, errorStyle="stop", errorTitle="地域区分", error="一覧から選んでください")
    dv.add("I4")
    ws["I4"].protection = Protection(locked=False)
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="whole", operator="between", formula1="0", formula2="31", allow_blank=True,
                         promptTitle="回数", prompt="1か月の利用回数（0〜31）")
    dv2.add("O8:O12")
    ws.add_data_validation(dv2)
    ws.protection.sheet = True

    m = wb.create_sheet("単価表")
    m["A1"], m["B1"] = "地域区分", "単価"
    for i, (k, v) in enumerate([("1級地", 11.20), ("2級地", 10.96), ("3級地", 10.90), ("4級地", 10.72),
                                ("5級地", 10.60), ("6級地", 10.36), ("7級地", 10.18), ("その他", 10.00)]):
        m.cell(i + 2, 1, k)
        m.cell(i + 2, 2, v)
    m["D1"], m["E1"] = "処遇改善率", 0.418
    m["D2"], m["E2"] = "上限月額", 9300
    wb.defined_names["処遇改善率"] = DefinedName("処遇改善率", attr_text="単価表!$E$1")
    wb.defined_names["上限月額"] = DefinedName("上限月額", attr_text="単価表!$E$2")
    wb.save(path)


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    from todoku.xlsx.edit import bake
    from todoku.xlsx.workbook import Workbook

    with tempfile.TemporaryDirectory() as td:
        for name, make in (("kami_meisai.xlsx", kami_excel),):
            raw = Path(td) / name
            make(raw)
            res = bake(Workbook(raw.read_bytes(), name))
            (OUT / name).write_bytes(res.blob)
            print("作成:", OUT / name, f"（計算結果 {len(res.patches)} 個を書き込み）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
