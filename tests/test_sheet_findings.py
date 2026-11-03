"""表計算の監査で見つかったことの回帰テスト（合成のブックで）：

1. 1904 年の日付の数え方（workbookPr date1904・BIFF8 の DATEMODE）を、書くときも見せるときも数式でも使う
2. ロケール・暦・数字の指定のある表示形式（[$-…]・[DBNum…]・B2）を Excel と同じ字にし、再現できないものは推測しない
3. 保護されたシートで、見出しのすぐ右のロックされた欄：見出しはその欄につき、右の補助のセルにつかない
4. 読み戻し・本人に伝える値：日付・数を、ブックの数え方と表示形式でそろえる。計算し直したセルも伝える
5. .ods：入力規則（整数 1〜10）を守り、窓に出す。ありもしないグラフの行を出さない

tests/fixtures/date1904.xls は、下の _build で作ったブック（1904 年の数え方）を LibreOffice で Excel 97 の形式にしたもの。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from todoku import api
from todoku.formula.numfmt import format_value
from todoku.formula.values import Unsupported

FIX = Path(__file__).parent / "fixtures"
openpyxl = pytest.importorskip("openpyxl")


def _build(path: Path, date1904: bool = False) -> Path:
    """申請のシート：B5 開始日（日付の書式・ロックのまま）・C2:C6 はロックを外した補助の値（D 列の数式が使う）・
    B12 受付日（日付の書式・ロックを外す）・B8 金額 =B6*B7・B4 人数（入力規則 整数 1〜10）。"""
    from openpyxl import Workbook
    from openpyxl.styles import Protection
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "申請"
    for r, lab in enumerate(("氏名", "区分", "人数", "開始日", "単価", "数量", "金額"), 2):
        ws.cell(r, 1, lab)
    ws["B5"].number_format = "yyyy/m/d"
    ws["A12"] = "受付日"
    ws["B12"].number_format = "yyyy/m/d"
    ws["B12"].protection = Protection(locked=False)
    ws["B6"], ws["B7"], ws["B8"] = 1000, 2, "=B6*B7"
    for r in range(2, 7):
        ws.cell(r, 3, r)
        ws.cell(r, 4, f"=C{r}*2")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="10", showErrorMessage=True)
    dv.add("B4")
    ws.add_data_validation(dv)
    ws["A14"] = "日付の計算"
    ws["B14"] = "=DATE(2026,11,3)"
    ws["C14"] = "=YEAR(B12)"
    ws["D14"] = '=TEXT(B14,"yyyy/m/d")'
    ws.protection.sheet = True
    for a in ("B2", "B3", "B4", "B6", "B7", "C2", "C3", "C4", "C5", "C6"):
        ws[a].protection = Protection(locked=False)
    wb.save(path)
    if date1904:
        items = {n: zipfile.ZipFile(path).read(n) for n in zipfile.ZipFile(path).namelist()}
        w = items["xl/workbook.xml"].decode()
        w = re.sub(r"<workbookPr(\s[^>]*)?/>", '<workbookPr\\1 date1904="1"/>', w) if "<workbookPr" in w else \
            w.replace("<sheets>", '<workbookPr date1904="1"/><sheets>')
        items["xl/workbook.xml"] = w.encode()
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n, b in items.items():
                z.writestr(n, b)
    return path


def _cell(path: Path, sheet: str, a: str):
    from todoku.book import open_book
    from todoku.formula.parser import parse_cell

    wb = open_book(path.read_bytes(), path.name)
    c = wb.sheet(sheet).cells.get(parse_cell(a))
    return wb, (c.value if c is not None else None)


# ====================================================================== 1. 1904 年の日付
@pytest.mark.parametrize("date1904, serial", [(False, 46329.0), (True, 44867.0)])
def test_date_system_write_and_readback(tmp_path, date1904, serial):
    src = _build(tmp_path / "a.xlsx", date1904)
    r = api.fill(str(src), {"申請!B12": "令和8年11月3日"}, str(tmp_path / "b.xlsx"))
    wb, v = _cell(tmp_path / "b.xlsx", "申請", "B12")
    assert wb.date1904 is date1904 and v == serial  # 1904 年の数え方では 1462 日小さい
    assert r["proof_ok"] and r["read_back_ok"], r
    assert "受付日：2026/11/3" in r["say_to_user"], r["say_to_user"]


def test_date_system_in_formulas_and_display(tmp_path):
    from todoku.book import open_book
    from todoku.xlsx.calc import Calc

    src = _build(tmp_path / "a.xlsx", True)
    r = api.fill(str(src), {"申請!B12": "2026-11-03"}, str(tmp_path / "b.xlsx"))
    wb = open_book((tmp_path / "b.xlsx").read_bytes(), "b.xlsx")
    calc = Calc(wb)
    assert calc.cell_value("申請", 14, 2) == 44867.0  # DATE は、ブックの数え方のシリアル値
    assert calc.cell_value("申請", 14, 3) == 2026.0  # YEAR(1904 年の数え方のシリアル値)
    assert calc.cell_value("申請", 14, 4) == "2026/11/3"  # TEXT も同じ数え方
    assert format_value(0, "yyyy/m/d", True) == "1904/1/1" and format_value(1462, "yyyy/m/d", False) == "1904/1/1"
    assert format_value(44867, "yyyy/m/d aaa", True) == "2026/11/3 火"
    assert r["read_back_ok"]


def test_date_system_xls_datemode(tmp_path):
    """Excel 97 の形式の DATEMODE（1904 年）：書く値も、見せる字も、1904 年の数え方。"""
    src = tmp_path / "a.xls"
    src.write_bytes((FIX / "date1904.xls").read_bytes())
    r = api.fill(str(src), {"申請!B12": "2026/11/3"}, str(tmp_path / "b.xls"))
    wb, v = _cell(tmp_path / "b.xls", "申請", "B12")
    assert wb.date1904 and v == 44867.0
    assert r["proof_ok"] and r["read_back_ok"] and "受付日：2026/11/3" in r["say_to_user"], r


# ====================================================================== 2. ロケール・暦・数字
# Excel（日本語版、Microsoft 365）が見せる字。2019/5/1 = 43586、2025/11/3（月）= 45964
LOCALE_CASES = [
    (43586, '[$-ja-JP-x-gannen]ggge"年"m"月"d"日"', "令和元年5月1日"),
    (43586, '[$-411]ggge"年"m"月"d"日"', "令和1年5月1日"),
    (45964, '[DBNum1][$-411]ggge"年"m"月"d"日"', "令和七年十一月三日"),
    (45964, '[DBNum3][$-411]0', "４５９６４"),
    (1234, "[DBNum1]0", "一二三四"),
    (1234, "[DBNum2]0", "壱弐参四"),
    (1234, "[DBNum1]General", "千二百三十四"),
    (1234, "[DBNum2]General", "壱阡弐百参拾四"),
    (1234, "[DBNum3]General", "千２百３十４"),
    (12345, "[DBNum1]General", "一万二千三百四十五"),
    (45964, "[$-th-TH]d mmmm bbbb", "3 พฤศจิกายน 2568"),
    (45964, "[$-1070000]d/mm/yyyy", "3/11/2568"),
    (45964, "[$-D07041E]d mmm yy", "๓ พ.ย. ๖๘"),
    (45964, "B2dd/mm/yyyy", "12/05/1447"),
    (45964, "[$-60000]dd/mm/yyyy", "12/05/1447"),
    (1234, "[$-2000401]0", "١٢٣٤"),
    (1234, "[$-4000439]0", "१२३४"),
    (45964, '[$-404]e"年"m"月"d"日"', "114年11月3日"),
    (45964, '[$-412]yyyy"년" m"월" d"일" aaa', "2025년 11월 3일 월"),
    (45964, "[$-ar-SA]dddd", "الاثنين"),
    (45964, "[$-ru-RU]dddd, d mmmm yyyy", "понедельник, 3 ноября 2025"),
    (45964, "[$-ru-RU]mmmm yyyy", "ноябрь 2025"),
    (45964, "[$-de-DE]dddd, d. mmmm yyyy", "Montag, 3. November 2025"),
    (45964, "[$-fr-FR]ddd d mmm yyyy", "lun. 3 nov. 2025"),
    (45964, "[$-es-ES]dddd d mmmm", "lunes 3 noviembre"),
    (45964, "[$-pt-BR]dddd", "segunda-feira"),
    (45964, "[$-zh-CN]dddd mmmm", "星期一 十一月"),
    (45964, "[$-zh-TW]ddd", "週一"),
    (45964, "ddd dddd aaa aaaa mmm", "Mon Monday 月 月曜日 Nov"),  # ロケールなし（日本語版）：ddd・mmm は英語
    (45964, "[$-411]ddd dddd mmm", "月 月曜日 11月"),
    (45964 + 43201.5 / 86400, "h:mm:ss.00", "12:00:01.50"),
    (1, "yyyy/m/d ddd", "1900/1/1 Sun"),
    (1234.5, '[$€-x-euro2] #,##0.00', "€ 1,234.50"),
    (1234, '[$¥-411]#,##0', "¥1,234"),
    # 桁の記号のあいだの字・条件のあるセクション（郵便番号・電話番号の様式）・桁の記号のないセクション・字と標準
    (1234567, "000-0000", "123-4567"),
    (9012345678, "000-0000-0000", "090-1234-5678"),
    (12, "[<=999]000;000-0000", "012"),
    (1234567, "[<=999]000;000-0000", "123-4567"),
    (1234, "[<=999]000;[<=9999]000-00;000-0000", "012-34"),
    (150, '[Red][>=100]0"点";[Blue]0"点"', "150点"),
    (0, '0;-0;"ゼロ"', "ゼロ"),
    (5, '"計"General"件"', "計5件"),  # 1900 年の数え方の 1900/1/1 は、Excel では日曜
]


@pytest.mark.parametrize("v, fmt, shown", LOCALE_CASES)
def test_locale_formats(v, fmt, shown):
    assert format_value(v, fmt) == shown


# 確かめていない組み合わせは、推測した字を返さない
@pytest.mark.parametrize("v, fmt", [
    (1234.5, "[$-2000401]#,##0.00"),  # アラビア文字の数字の区切りの字
    (45964, "[$-170000]yyyy"),  # ウンム・アル＝クラー暦
    (45964, "[$-80000]yyyy"),  # ヘブライ暦
    (45964, "[$-F800]dddd"),  # PC の設定の長い日付
    (45964, "[$-x-sysdate]dddd"),
    (45964, "[$-ko-KR]mmm"),  # 省略形を確かめていない
    (45964, "[$-409]ggge"),  # 英語のロケールの元号
    (1234, "[DBNum4]0"),
    (1234, "[DBNum1][$-804]0"),  # 中国語の漢数字
    (45964, "[$-ar-SA]yyyy"),  # 暦を指定しないアラビア語の年
    (43586, '[$-ja-JP-x-gannen]ge.m.d'),  # 「年」の続かない元年
    (-5, "[<=999]000;000-0000"),  # 条件付きの表示形式の負の数
    (50, '[>=100]0;[<10]"少"'),  # どの条件にも合わない
])
def test_locale_formats_unsupported(v, fmt):
    with pytest.raises(Unsupported):
        format_value(v, fmt)


def test_unsupported_format_is_flagged_in_window_and_grade(tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "表示"
    ws["A1"], ws["B1"] = "金額", 1234.5
    ws["B1"].number_format = "[$-2000401]#,##0.00"
    ws["A2"], ws["B2"] = "年", 45964
    ws["B2"].number_format = "[$-th-TH]d mmmm bbbb"
    wb.save(tmp_path / "f.xlsx")
    _k, doc, parts = api.window(str(tmp_path / "f.xlsx"))
    md = "\n".join(t for _n, t in parts)
    assert "〔表示形式「[$-2000401]#,##0.00」は再現できません：元の値 1234.5〕" in md
    assert "3 พฤศจิกายน 2568" in md
    assert [u["cell"] for u in doc["sheets"][0]["unshown_formats"]] == ["B1"]
    g = api.grade(str(tmp_path / "f.xlsx"))
    r = next(x for x in g["reasons"] if x["title"] == "再現できない表示形式")
    assert g["grade"] == "B" and r["grade"] == "B" and r["where"] == ["表示!B1"]


# ====================================================================== 3. 保護されたシートのロックされた欄
def test_locked_label_cell_keeps_its_label(tmp_path):
    src = _build(tmp_path / "a.xlsx")
    fields = {f["id"]: f for f in api.fields(str(src))}
    b5 = fields["申請!B5"]
    assert b5["label"] == "開始日" and b5["readonly"] and "パスワードが要る" in b5["protected"]
    assert not fields["申請!C5"].get("label")  # 右の補助のセル（D5 の数式が使う）は、見出しを取らない
    _k, doc, parts = api.window(str(src))
    md = "\n".join(t for _n, t in parts)
    assert "開始日 ｜ ［B5：保護されたセル］ ｜ ［C5＝5］" in md
    g = api.grade(str(src))
    w = next(x for x in g["write"]["reasons"] if x["title"] == "保護されたセルの欄")
    assert g["write"]["grade"] == "B" and w["where"] == ["申請!B5"]
    with pytest.raises(api.ApiError):
        api.fill(str(src), {"申請!B5": "2026/11/3"}, str(tmp_path / "b.xlsx"))  # Excel でも書けない


# ====================================================================== 4. 読み戻しと本人に伝える値
def test_say_to_user_shows_cell_format_and_recalculated(tmp_path):
    src = _build(tmp_path / "a.xlsx")
    r = api.fill(str(src), {"申請!B4": "３", "申請!B6": 2500}, str(tmp_path / "b.xlsx"))
    say = r["say_to_user"]
    assert r["read_back_ok"], r["read_back_differences"]
    assert "人数：3\n" in say and "単価：2500" in say and "3.0" not in say
    assert "再計算で変わった所" in say and "金額（申請!B8）" in say and "→ 5000" in say, say
    assert any(x["id"] == "申請!B8" and x["after"] == "5000" for x in r["recalculated_cells"])


# ====================================================================== 5. .ods
def test_ods_whole_number_rule_and_no_bogus_chart(tmp_path):
    from test_openformats import ODS_AUTO, make_ods

    val = ('<table:content-validations><table:content-validation table:name="val1" '
           'table:condition="of:cell-content-is-whole-number() and cell-content-is-between(1,10)" table:allow-empty-cell="false">'
           '<table:error-message table:message-type="stop" table:display="true"/></table:content-validation></table:content-validations>')
    sheet = ('<table:table table:name="申請"><table:table-column table:number-columns-repeated="3"/>'
             '<table:table-row><table:table-cell office:value-type="string" calcext:value-type="string"><text:p>人数</text:p></table:table-cell>'
             '<table:table-cell table:style-name="ce1" table:content-validation-name="val1"/><table:table-cell/></table:table-row>'
             '</table:table>')
    src = tmp_path / "a.ods"
    src.write_bytes(make_ods(sheet, ODS_AUTO, val))
    for bad in (11, 0, 2.5):
        with pytest.raises(api.ApiError):
            api.fill(str(src), {"申請!B1": bad}, str(tmp_path / "b.ods"))
    r = api.fill(str(src), {"申請!B1": 5}, str(tmp_path / "b.ods"))
    assert r["proof_ok"] and r["read_back_ok"] and "人数：5\n" in r["say_to_user"] + "\n", r["say_to_user"]
    _k, doc, parts = api.window(str(src))
    md = "\n".join(t for _n, t in parts)
    assert "整数 1〜10" in md and "None" not in md
    assert "グラフ" not in md and "AttributeError" not in md
