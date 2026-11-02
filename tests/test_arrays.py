"""動的配列の関数・XLOOKUP の一致モード・通貨の書式（Microsoft の関数の説明の例と、その規則どおりの結果）。"""

from __future__ import annotations

from pathlib import Path

import pytest

from todoku.formula.evaluator import Evaluator
from todoku.formula.parser import parse
from todoku.formula.values import Arr, ExcelError
from todoku.xlsx.calc import Calc, values_equal
from todoku.xlsx.workbook import load

KAMI = Path(__file__).parent / "fixtures" / "kami_meisai.xlsx"


def ev(expr: str, array: bool = True):
    wb = load(str(KAMI))
    v = Evaluator(Calc(wb), wb.sheets[0].name, 1, 1, array=array).run(parse(expr))
    return v


def rows(v) -> list[list]:
    assert isinstance(v, Arr), v
    return [[x.code if isinstance(x, ExcelError) else x for x in r] for r in v.data]


def code(v) -> str:
    assert isinstance(v, ExcelError), v
    return v.code


def test_sequence():
    assert rows(ev("SEQUENCE(2,3)")) == [[1, 2, 3], [4, 5, 6]]
    assert rows(ev("SEQUENCE(3,1,10,-5)")) == [[10], [5], [0]]
    assert code(ev("SEQUENCE(0)")) == "#CALC!"
    assert code(ev("SEQUENCE(-1)")) == "#VALUE!"


def test_filter():
    data = '{"Apple",10;"Pear",0;"Kiwi",5}'
    assert rows(ev(f"FILTER({data},INDEX({data},0,2)>0)")) == [["Apple", 10], ["Kiwi", 5]]
    assert ev(f'FILTER({data},INDEX({data},0,2)>99,"なし")') == "なし"
    assert code(ev(f"FILTER({data},INDEX({data},0,2)>99)")) == "#CALC!"
    assert rows(ev("FILTER({1,2,3},{TRUE,FALSE,TRUE})")) == [[1, 3]]  # 列を選ぶ
    assert code(ev("FILTER({1;2;3},{TRUE;FALSE})")) == "#VALUE!"  # 大きさが合わない


def test_unique():
    assert rows(ev('UNIQUE({"a";"A";"b";"a"})')) == [["a"], ["b"]]  # 大文字・小文字は区別しない
    assert rows(ev('UNIQUE({"a";"A";"b";"a"},FALSE,TRUE)')) == [["b"]]  # 1回だけ出てくるもの
    assert rows(ev("UNIQUE({1,1,2},TRUE)")) == [[1, 2]]
    assert rows(ev('UNIQUE({1;"1";1})')) == [[1], ["1"]]  # 数と文字列は別


def test_sort_and_sortby():
    assert rows(ev("SORT({3;1;2})")) == [[1], [2], [3]]
    assert rows(ev("SORT({3;1;2},1,-1)")) == [[3], [2], [1]]
    assert rows(ev('SORT({2;"b";TRUE;"a";1})')) == [[1], [2], ["a"], ["b"], [True]]  # 数値・文字列・論理値の順
    assert rows(ev('SORT({"b",2;"a",1;"c",1},{2,1},{1,-1})')) == [["c", 1], ["a", 1], ["b", 2]]
    assert rows(ev('SORTBY({"x";"y";"z"},{3;1;2})')) == [["y"], ["z"], ["x"]]
    assert rows(ev('SORT({1,3,2},1,1,TRUE)')) == [[1, 2, 3]]  # 列を並べ替える


def test_reshape():
    a = "{1,2,3;4,5,6;7,8,9}"
    assert rows(ev(f"TOCOL({a})")) == [[x] for x in range(1, 10)]
    assert rows(ev(f"TOCOL({a},0,TRUE)")) == [[1], [4], [7], [2], [5], [8], [3], [6], [9]]
    assert rows(ev("TOROW({1;2})")) == [[1, 2]]
    assert rows(ev(f"TAKE({a},2)")) == [[1, 2, 3], [4, 5, 6]]
    assert rows(ev(f"TAKE({a},-1)")) == [[7, 8, 9]]
    assert rows(ev(f"TAKE({a},,-1)")) == [[3], [6], [9]]
    assert rows(ev(f"DROP({a},1,1)")) == [[5, 6], [8, 9]]
    assert code(ev(f"DROP({a},3)")) == "#CALC!"
    assert rows(ev(f"CHOOSECOLS({a},1,-1)")) == [[1, 3], [4, 6], [7, 9]]
    assert rows(ev(f"CHOOSEROWS({a},-1,1)")) == [[7, 8, 9], [1, 2, 3]]
    assert rows(ev("VSTACK({1,2},{3})")) == [[1, 2], [3, "#N/A"]]
    assert rows(ev("HSTACK({1;2},{3})")) == [[1, 3], [2, "#N/A"]]
    assert rows(ev("EXPAND({1,2},2,3,0)")) == [[1, 2, 0], [0, 0, 0]]
    assert rows(ev("WRAPROWS({1,2,3,4,5},2)")) == [[1, 2], [3, 4], [5, "#N/A"]]
    assert rows(ev("WRAPCOLS({1,2,3},2,0)")) == [[1, 3], [2, 0]]
    assert rows(ev("TRANSPOSE({1,2;3,4})")) == [[1, 3], [2, 4]]


@pytest.mark.parametrize("expr, expected", [
    ('XLOOKUP("b",{"a";"b";"c"},{1;2;3})', 2),
    ('XLOOKUP("z",{"a";"b"},{1;2},"なし")', "なし"),
    ("XLOOKUP(25,{10;20;30},{1;2;3},,-1)", 2),  # 完全一致か、次に小さい値
    ("XLOOKUP(25,{10;20;30},{1;2;3},,1)", 3),  # 完全一致か、次に大きい値
    ('XLOOKUP("B*",{"a";"bc";"bd"},{1;2;3},,2)', 2),  # ワイルドカード
    ('XLOOKUP("b*",{"a";"b*";"bc"},{1;2;3})', 2),  # 完全一致では * もそのままの字
    ('XLOOKUP("b",{"b";"a";"b"},{1;2;3},,0,-1)', 3),  # 末尾から
    ("XLOOKUP(20,{10;20;30;40},{1;2;3;4},,0,2)", 2),  # 二分探索（昇順）
    ("XLOOKUP(20,{40;30;20;10},{1;2;3;4},,0,-2)", 3),  # 二分探索（降順）
    ("XMATCH(25,{10;20;30},1)", 3),
    ('XMATCH("c",{"a","b","c"})', 3),
])
def test_xlookup_modes(expr, expected):
    assert values_equal(expected, ev(expr, array=False)), (expr, ev(expr, array=False))


def test_xlookup_returns_row_and_reference():
    assert rows(ev('XLOOKUP("b",{"a";"b"},{1,2;3,4})')) == [[3, 4]]
    assert code(ev('XLOOKUP("z",{"a";"b"},{1;2})', array=False)) == "#N/A"


@pytest.mark.parametrize("expr, expected", [
    ("DOLLAR(1234.567,2)", "$1,234.57"), ("DOLLAR(1234.567,-2)", "$1,200"), ("DOLLAR(-1234.567,-2)", "($1,200)"),
    ("DOLLAR(-0.123,4)", "($0.1230)"), ("DOLLAR(99.888)", "$99.89"),
    ("FIXED(1234.567,1)", "1,234.6"), ("FIXED(1234.567,-1)", "1,230"), ("FIXED(-1234.567,-1,TRUE)", "-1230"),
    ("FIXED(44.332)", "44.33"), ("YEN(1234.567)", "¥1,235"),
])
def test_currency_text(expr, expected):
    assert ev(expr, array=False) == expected


def test_scalar_functions_work_element_by_element_in_array_formulas():
    assert rows(ev('MID("ab1.5",SEQUENCE(5),1)')) == [["a"], ["b"], ["1"], ["."], ["5"]]
    assert rows(ev('LEN({"a","bb";"ccc",""})')) == [[1, 2], [3, 0]]
    # gov.wales の様式の数式：文字の中の数字と小数点だけを取り出して数にする
    f = ('IFERROR(LET(_xlpm.txt,"約12.5kW",_xlpm.chars,MID(_xlpm.txt,SEQUENCE(LEN(_xlpm.txt)),1),'
         '_xlpm.num,TEXTJOIN("",TRUE,IF((_xlpm.chars>="0")*(_xlpm.chars<="9")+(_xlpm.chars="."),_xlpm.chars,"")),'
         '--_xlpm.num),"")')
    assert ev(f) == 12.5


# ====================================================================== 配列数式の結果の書き込み
def _array_book() -> bytes:
    """B1:B3 は昔の配列数式 {=A1:A3*2}、C1 は動的配列 =SEQUENCE(A4)（こぼれる範囲 C1:C3）。"""
    import io
    import zipfile

    ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' \
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    sheet = (f'<?xml version="1.0" encoding="UTF-8"?><worksheet {ns}><sheetData>'
             '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B3">A1:A3*2</f><v>2</v></c>'
             '<c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn.SEQUENCE(A4)</f><v>1</v></c></row>'
             '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>4</v></c><c r="C2"><v>2</v></c></row>'
             '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>6</v></c><c r="C3"><v>3</v></c></row>'
             '<row r="4"><c r="A4"><v>3</v></c></row></sheetData></worksheet>')
    meta = ('<?xml version="1.0" encoding="UTF-8"?><metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1">'
            '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" '
            'splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
            '</metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
            '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
            '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        z.writestr("xl/workbook.xml", f'<?xml version="1.0" encoding="UTF-8"?><workbook {ns}><sheets>'
                   '<sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>'
                   '</Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        z.writestr("xl/metadata.xml", meta)
    return buf.getvalue()


def test_array_formula_members_are_written():
    from todoku.xlsx.calc import verify
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove
    from todoku.xlsx.workbook import Workbook

    blob = _array_book()
    wb = Workbook(blob)
    assert verify(wb).summary()["mismatch"] == 0
    res = fill(wb, {"S!A2": 5})
    assert not res.unverified and not res.full_calc_on_load, res.unverified
    after = Workbook(res.blob)
    assert after.sheet("S").cells[(2, 2)].value == 10  # 配列数式 {=A1:A3*2} の2つ目
    assert verify(after).summary()["mismatch"] == 0
    pr = prove(blob, res.blob, {(p.sheet, p.address) for p in res.patches if p.kind == "input"})
    assert pr.ok, (pr.unexpected, pr.lint_errors)


def test_dynamic_array_spill_change_is_left_to_excel():
    from todoku.xlsx.edit import fill
    from todoku.xlsx.workbook import Workbook

    res = fill(Workbook(_array_book()), {"S!A4": 4})  # SEQUENCE(4)：こぼれる範囲が C1:C4 に広がる
    assert any("C1" in u and "動的配列" in u for u in res.unverified), res.unverified
    assert res.full_calc_on_load
