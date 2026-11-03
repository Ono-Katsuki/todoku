"""ロケールつきの表示形式（和暦・元年・漢数字・タイの仏暦・ヒジュラ暦・アラビア・インドの数字・台湾の民国）を、
todoku の表示形式と、Excel が見せる字（下の expect。Excel for Microsoft 365 の表示）とで比べる。
同じ値・形式の xlsx も作り、窓に出る字も確かめる。

使い方: PYTHONPATH=src python locale_formats.py OUT.xlsx
"""

from __future__ import annotations

import sys

from todoku.formula.numfmt import format_value

# 2019-05-01 = シリアル 43586、2025-11-03 = 45964
CASES = [
    # (値, 形式, Excel の表示)
    (43586, '[$-ja-JP-x-gannen]ggge"年"m"月"d"日"', "令和元年5月1日"),
    (43586, '[$-411]ggge"年"m"月"d"日"', "令和1年5月1日"),
    (45964, '[$-411]ggge"年"m"月"d"日"', "令和7年11月3日"),
    (45964, '[DBNum1][$-411]ggge"年"m"月"d"日"', "令和七年十一月三日"),
    (45964, '[DBNum3][$-411]0', "４５９６４"),
    (1234, "[DBNum1]0", "一二三四"),
    (1234, "[DBNum2]0", "壱弐参四"),
    (45964, "[$-th-TH]d mmmm bbbb", "3 พฤศจิกายน 2568"),
    (45964, "[$-1070000]d/mm/yyyy", "3/11/2568"),  # 07 = タイ仏暦
    (45964, "[$-D07041E]d mmm yy", "๓ พ.ย. ๖๘"),  # D0 = タイ数字, 07 = 仏暦, 041E = th-TH
    (45964, "B2dd/mm/yyyy", "12/05/1447"),  # ヒジュラ暦（Excel の Hijri：1447年5月12日 前後）
    (1234.5, "[$-2000401]#,##0.00", "١٬٢٣٤٫٥٠"),  # 02 = アラビア・インド数字
    (1234.5, "[$-3000439]#,##0.00", "१,२३४.५०"),  # 03 = デーヴァナーガリー数字（hi-IN）
    (45964, '[$-404]e"年"m"月"d"日"', "114年11月3日"),  # 台湾：民国（Excel は e を中華民国の年に）
    (45964, '[$-412]yyyy"년" m"월" d"일" aaa', "2025년 11월 3일 월"),
    (45964, "[$-ar-SA]dddd", "الاثنين"),
    (45964, "[$-ru-RU]dddd, d mmmm yyyy", "понедельник, 3 ноября 2025"),
    (45964, "[$-de-DE]dddd, d. mmmm yyyy", "Montag, 3. November 2025"),
    (1234567.891, "[$-de-DE]#,##0.00", "1.234.567,89"),  # 区切りは OS の設定。de-DE の PC では
]


def main(out):
    rows = []
    for v, f, exp in CASES:
        try:
            got = format_value(v, f)
        except Exception as e:  # noqa: BLE001
            got = f"<{type(e).__name__}: {e}>"
        st = "OK" if got == exp else "DIFF"
        rows.append((v, f, exp, got, st))
        print(f"{st:4} {f!r:42} excel={exp!r:30} todoku={got!r}")
    try:
        import openpyxl
    except ImportError:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "fmt"
    for i, (v, f, exp, _g, _s) in enumerate(rows, 1):
        ws.cell(i, 1, f)
        c = ws.cell(i, 2, v)
        c.number_format = f
        ws.cell(i, 3, exp)
    wb.save(out)
    print("saved", out)


if __name__ == "__main__":
    main(sys.argv[1])
