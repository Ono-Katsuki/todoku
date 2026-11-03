"""Excel の様式を、データ（入力欄）と処理（入力につながる数式）に分ける・試算する。"""

from __future__ import annotations

import io

import pytest

from todoku import api

openpyxl = pytest.importorskip("openpyxl")


def _form(tmp_path):
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "計算"
    ws["A1"], ws["A2"], ws["A3"], ws["A4"] = "単価", "数量", "小計", "合計（税込）"
    ws["B1"], ws["B2"] = 100, 2  # openpyxl は空のセルを書き出さないので、初めの値を入れておく
    ws["B3"] = "=B1*B2"
    ws["B4"] = "=ROUND(B3*1.1,0)"
    ws["D1"] = "=1+1"  # 入力につながらない数式
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="1000000", showErrorMessage=True)
    dv.add("B1")
    dv.add("B2")
    ws.add_data_validation(dv)
    p = tmp_path / "form.xlsx"
    buf = io.BytesIO()
    wb.save(buf)
    p.write_bytes(buf.getvalue())
    return str(p)


def test_inputs_and_steps_in_words(tmp_path):
    f = _form(tmp_path)
    m = api.calc_model(f)
    assert m["inputs"] == 2 and m["steps"] == 2 and m["results"] == 1  # D1 は入力につながらない
    g = m["result_groups"][0]
    assert g["cells"] == ["計算!B4"] and g["in_words"] == "＝ROUND(「小計」*1.1,0)"
    e = api.calc_model(f, "計算!B4")["explain"]
    assert e["from"][0]["in_words"] == "＝「単価」*「数量」"
    with pytest.raises(api.ApiError):
        api.calc_model(f, "計算!D1")


def test_try_calculation_does_not_write(tmp_path):
    f = _form(tmp_path)
    before = open(f, "rb").read()
    r = api.calculate(f, {"計算!B1": 120, "計算!B2": 3})
    got = {c["cell"]: c["after"] for c in r["changed"]}
    assert str(got["計算!B3"]).strip() == "360" and str(got["計算!B4"]).strip() == "396"
    assert "合計（税込）：" in r["say_to_user"] and open(f, "rb").read() == before
    with pytest.raises(api.ApiError):
        api.calculate(f, {"計算!B1": -5})  # 入力規則に合わない
