"""大きな配列：作らずに、頼まれたマスだけ計算する。

- 小さな配列でも「大きな配列の道」を通すと、いままでの結果と 1 つも違わない（test_arrays をそのまま走らせる）
- シートに入る大きさなら計算できる。集計は式で答えるか、少しずつ読む。終わらなければ、進んだ所から続けられる
"""

from __future__ import annotations

import json
import random
import time
import tracemalloc
from fractions import Fraction

import pytest

from todoku import limits
from todoku.formula import values
from todoku.formula.evaluator import Evaluator
from todoku.formula.lazyarr import LazyArr, Poly, ScanStore, _power_sum
from todoku.formula.parser import parse
from todoku.formula.values import ExcelError, Unfinished

from test_arrays import *  # noqa: F401,F403  いまの配列の試験を、大きな配列の道でも走らせる
import test_arrays as _base


class _Ctx:
    date1904 = False

    def cell_value(self, sheet, r, c):
        return None

    def bounds(self, sheet):
        return (1, 1)

    def resolve_name(self, name, sheet, here):
        return None

    def has_sheet(self, sheet):
        return True


def calc(text: str, ctx=None):
    return Evaluator(ctx or _Ctx(), "S", 1, 1, 0, 0, True).run(parse(text))


@pytest.fixture(autouse=True)
def every_array_is_lazy(request, monkeypatch):
    """test_arrays から来た試験は、どの大きさの配列も LazyArr にして走らせる。"""
    if request.node.function.__module__ == _base.__name__:
        monkeypatch.setattr(values, "LAZY_ABOVE", 0)


def flat(v):
    return [x.code if isinstance(x, ExcelError) else x for row in v.data for x in row]


# ====================================================================== 値は作らずに読む
FULL = "ROW(1:1048576)*COLUMN(A:XFD)"


@pytest.mark.parametrize("text,rows,cols", [("SEQUENCE(1048576,16384)", 1048576, 16384), ("EXPAND(1,1048576,16384,0)", 1048576, 16384),
                                            (FULL, 1048576, 16384), ("WRAPROWS(SEQUENCE(1048576),16384)", 64, 16384),
                                            ("-ROW(1:1048576)", 1048576, 1), ("ABS(ROW(1:1048576))", 1048576, 1)])
def test_arrays_that_fit_the_sheet_exist(text, rows, cols):
    tracemalloc.start()
    t0 = time.time()
    v = calc(text)
    assert isinstance(v, LazyArr)
    assert (v.rows, v.cols) == (rows, cols)
    v.at(rows - 1, cols - 1)
    peak = tracemalloc.get_traced_memory()[1]
    tracemalloc.stop()
    assert peak < 8 << 20 and time.time() - t0 < 5


@pytest.mark.parametrize("text", ["SEQUENCE(1048577)", "SEQUENCE(1,16385)", "EXPAND(1,1048577)", "EXPAND(1,1,16385)", "VSTACK(SEQUENCE(1048576),1)"])
def test_past_the_sheet_is_still_num(text):
    assert getattr(calc(text), "code", None) == "#NUM!"


def test_cells_of_the_big_arrays():
    v = calc(FULL)
    assert v.at(1048575, 16383) == 1048576 * 16384
    assert v.at(4, 2) == 15
    s = calc("SEQUENCE(1048576,16384,5,2)")
    assert s.at(0, 0) == 5 and s.at(1, 0) == 5 + 16384 * 2 and s.at(1048575, 16383) == 5 + 2 * (1048576 * 16384 - 1)
    assert calc('EXPAND(7,1048576,16384,"x")').at(100, 100) == "x"
    assert calc("EXPAND(7,1048576,16384)").at(100, 100).code == "#N/A"
    assert calc("INDEX(" + FULL + ",1048576,16384)") == 1048576 * 16384
    assert calc("INDEX(" + FULL + ",0,3)").at(9, 0) == 30
    assert getattr(calc("TRANSPOSE(SEQUENCE(1048576))"), "code", None) == "#NUM!"  # 1,048,576 列はシートに入らない
    assert calc("TRANSPOSE(SEQUENCE(16384,1000))").at(5, 2) == 2 * 1000 + 6
    d = calc("DROP(SEQUENCE(1048576,2),5)")
    assert d.at(0, 0) == 11 and d.rows == 1048571
    t = calc("TAKE(SEQUENCE(1048576,16384),-1)")
    assert t.at(0, 16383) == 1048576 * 16384
    assert calc('IF(ROW(1:1048576)>2,"a","b")').at(2, 0) == "a"
    assert calc("IFERROR(1/(ROW(1:1048576)-1),-1)").at(0, 0) == -1
    assert calc("LET(_xlpm.x,SEQUENCE(1048576,16384),INDEX(_xlpm.x,2,1))") == 16385


def test_textsplit_and_makearray():
    v = calc('TEXTSPLIT("a,b;c",",",";")')
    assert flat(v) == ["a", "b", "c", "#N/A"]
    v = calc('TEXTSPLIT("x"&REPT(",",8000)&REPT(";",16000),",",";")')  # 字の上限の中で、1 億マスを超える
    assert (v.rows, v.cols) == (16001, 8001) and isinstance(v, LazyArr)
    assert v.at(0, 0) == "x" and v.at(0, 8000) == "" and v.at(5, 5).code == "#N/A"
    m = calc("MAKEARRAY(1048576,16384,LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r*_xlpm.c))")
    assert isinstance(m, LazyArr) and m.at(1048575, 16383) == 1048576 * 16384
    assert flat(calc("MAKEARRAY(2,2,LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r+_xlpm.c))")) == [2, 3, 3, 4]
    assert flat(calc("_xlfn.MAKEARRAY(2,2,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r*_xlpm.c))")) == [1, 2, 2, 4]
    assert calc('_xlfn.TEXTSPLIT("a b"," ")').at(0, 1) == "b"


# ====================================================================== 集計：式で答える
@pytest.mark.parametrize("text", [
    "SUM(ROW(1:300)*COLUMN(A:CV))", "COUNT(ROW(1:300)*COLUMN(A:CV))", "MAX(ROW(1:300)*COLUMN(A:CV))", "MIN(ROW(1:300)*COLUMN(A:CV))",
    "AVERAGE(ROW(1:300)*COLUMN(A:CV))", "SUM(ROW(1:300)^2)", "SUM(-ROW(1:300)+COLUMN(A:CV)*3-7)", "COUNTA(SEQUENCE(30,100))",
    "SUM(SEQUENCE(30,100,5,3))", "MAX(SEQUENCE(30,100,5,-3))", "MIN(SEQUENCE(30,100,5,-3))", "AVERAGE(SEQUENCE(300,10,2,2))",
    "SUM(EXPAND(SEQUENCE(3),100,50,2))", "COUNT(EXPAND(SEQUENCE(3),100,50,\"x\"))", "COUNTA(EXPAND(SEQUENCE(3),100,50,\"x\"))",
    "SUM(EXPAND(SEQUENCE(3),100,50))", "LARGE(SEQUENCE(30,100),17)", "SMALL(SEQUENCE(30,100),17)", "LARGE(ROW(1:300)*COLUMN(A:CV),5)",
    "SMALL(ROW(1:300)*COLUMN(A:CV),5)", "VAR(ROW(1:300)*COLUMN(A:CV))", "VAR.P(SEQUENCE(30,100))", "STDEV(SEQUENCE(30,100))",
    "SUMPRODUCT(ROW(1:300)*COLUMN(A:CV))", "SUMPRODUCT(ROW(1:300),ROW(1:300)+1)", "COUNTIF(SEQUENCE(30,100),\">50\")", "SUMIF(SEQUENCE(30,100),\">50\")",
    "PRODUCT(SEQUENCE(30))", "SUM(TRANSPOSE(SEQUENCE(30,100)))", "SUM(TOCOL(SEQUENCE(30,100)))", "SUM(TAKE(SEQUENCE(300,10),100))",
    "SUM(DROP(SEQUENCE(300,10),7,2))", "SUM(VSTACK(SEQUENCE(100),SEQUENCE(100,2)))", "SUM(HSTACK(SEQUENCE(100),SEQUENCE(50)))",
    "SUM(INDEX(SEQUENCE(300,10),0,3))", "MAX(FILTER(SEQUENCE(1000),MOD(SEQUENCE(1000),3)=0))", "SUM(UNIQUE(MOD(SEQUENCE(1000),7)))",
    "SUM(SORT(SEQUENCE(1000),1,-1))", "SUM(MOD(ROW(1:1000),7))", "AND(SEQUENCE(1000))", "OR(SEQUENCE(1000)-1)", "XOR(SEQUENCE(1000))",
    "COUNTBLANK(EXPAND(1,100,50,\"\"))", "AVERAGEIF(SEQUENCE(1000),\">10\")", "COUNTIFS(SEQUENCE(1000),\">10\",SEQUENCE(1000),\"<20\")",
    "SUMIFS(SEQUENCE(1000),SEQUENCE(1000),\">10\")", "MAXIFS(SEQUENCE(1000),SEQUENCE(1000),\"<20\")", "SUM(ROW(1:1000)/2)", "SUM(ROW(1:1000)%)",
])
def test_lazy_equals_lists(monkeypatch, text):
    want = calc(text)  # リストの道
    monkeypatch.setattr(values, "LAZY_ABOVE", 0)
    got = calc(text)
    if isinstance(want, ExcelError):
        assert isinstance(got, ExcelError) and got.code == want.code
    else:
        assert values.to_15(float(got)) == values.to_15(float(want)), (got, want)


def test_full_sheet_aggregates_answer_at_once():
    t0 = time.time()
    r, c = 1048576, 16384
    assert calc(f"COUNT({FULL})") == r * c
    assert calc(f"MAX({FULL})") == r * c
    assert calc(f"MIN({FULL})") == 1
    assert calc(f"SUM({FULL})") == float(r * (r + 1) // 2 * (c * (c + 1) // 2))
    assert calc(f"AVERAGE({FULL})") == float(r * (r + 1) // 2 * (c * (c + 1) // 2)) / (r * c)
    assert calc("SUM(SEQUENCE(1048576,16384))") == float(r * c * (r * c + 1) // 2)
    assert calc("LARGE(SEQUENCE(1048576,16384),1)") == r * c
    assert calc("SMALL(SEQUENCE(1048576,16384),5)") == 5
    assert calc("LARGE(SEQUENCE(1048576,16384),3)") == r * c - 2
    assert calc("COUNT(EXPAND(1,1048576,16384,0))") == r * c
    assert calc("SUM(EXPAND(1,1048576,16384,0))") == 1
    assert calc("SUM(EXPAND(1,1048576,16384,2))") == 1 + 2 * (r * c - 1)
    assert calc('COUNT(EXPAND(1,1048576,16384,"x"))') == 1
    assert calc('COUNTA(EXPAND(1,1048576,16384,"x"))') == r * c
    assert calc("SUM(EXPAND(1,1048576,16384))").code == "#N/A"
    assert time.time() - t0 < 2


def test_sum_within_double_range_is_bit_exact():
    # どの途中の和も double でちょうど表せる：式の答えと順に足した答えがビットまで同じ
    s = sum(float(i) * float(j) for i in range(1, 1001) for j in range(1, 101))
    assert calc("SUM(ROW(1:1000)*COLUMN(A:CV))") == s


def test_power_sums():
    for p in range(0, 9):
        for n in (0, 1, 2, 10, 37):
            assert _power_sum(n, p) == sum(i ** p for i in range(n)), (n, p)


def test_poly_substitute():
    random.seed(1)
    p = Poly({(1, 1): Fraction(3), (2, 0): Fraction(1, 2), (0, 0): Fraction(7)})
    q = p.substitute(2, 5)

    def at(poly, i, j):
        return sum(v * i ** a * j ** b for (a, b), v in poly.t.items())

    for _ in range(20):
        i, j = random.randrange(9), random.randrange(9)
        assert at(q, i, j) == at(p, i + 2, j + 5)


# ====================================================================== 順に読む：進み具合と続き
def _slow():
    return "SUM(MOD(ROW(1:1048576)*COLUMN(A:E),7))"


def test_chunked_aggregate_stops_says_how_far_and_resumes(monkeypatch):
    b = limits.Budget()
    b.calc_left = 0.05
    ctx = _Ctx()
    store = ctx.scan_store = ScanStore()
    with limits.call_budget(b):
        with pytest.raises(Unfinished) as e:
            calc(_slow(), ctx)
    assert 0 < e.value.done < e.value.total == 1048576 * 5
    assert "万" not in str(e.value) and f"{e.value.total:,}" in str(e.value) and "続き" in str(e.value)
    rec = next(iter(store.data.values()))
    assert rec["k"] == e.value.done and json.loads(json.dumps(rec))["state"] == rec["state"]
    # 控えから：すでに読んだ所は読まない。結果は最初から計算したのと同じ
    limit_small = "SUM(MOD(ROW(1:70000),7))"
    ctx2 = _Ctx()
    ctx2.scan_store = ScanStore()
    b = limits.Budget()
    b.calc_left = 0.0
    with limits.call_budget(b):
        with pytest.raises(Unfinished) as e2:
            calc(limit_small, ctx2)
    # 時間の予算 0：はじめのまとまり（4,096 マス）で止まる。Windows の時計は 15.6 ms 刻みで、経過が 0 に見えて 1 まとまり多く進むことがある
    assert e2.value.done in (4096, 8192)
    ctx2.scan_seq = {}
    b = limits.Budget()
    b.calc_left = 1000
    with limits.call_budget(b):
        got = calc(limit_small, ctx2)
    assert got == sum(i % 7 for i in range(1, 70001))
    assert next(iter(ctx2.scan_store.data.values()))["done"]


def test_memory_stays_small_on_a_chunked_aggregate():
    b = limits.Budget()
    b.calc_left = 1000
    tracemalloc.start()
    with limits.call_budget(b):
        got = calc("SUM(MOD(ROW(1:200000),7))")
    peak = tracemalloc.get_traced_memory()[1]
    tracemalloc.stop()
    assert got == sum(i % 7 for i in range(1, 200001))
    assert peak < 4 << 20


def test_filter_unique_sort_on_big_arrays_keep_only_the_numbers():
    v = calc("FILTER(SEQUENCE(100000),MOD(SEQUENCE(100000),2)=0)")
    assert v.rows == 50000 and v.at(1, 0) == 4
    u = calc("UNIQUE(MOD(SEQUENCE(100000),3))")
    assert flat(u) == [1, 2, 0]
    s = calc("SORT(SEQUENCE(100000),1,-1)")
    assert isinstance(s, LazyArr) and s.at(0, 0) == 100000 and s.at(99999, 0) == 1


def test_job_runs_without_a_time_limit_and_budget_reaches_the_deep_thread():
    from todoku.xlsx.calc import _run_deep

    b = limits.Budget()
    b.calc_left = float("inf")
    with limits.call_budget(b):
        assert _run_deep(limits.budget) is b


def test_sum_past_double_range_follows_excel_row_order(monkeypatch):
    # 2^53 を超える和：Excel と同じに行の順に足す（実機の Excel で 15 桁まで一致を確かめた形）
    b = limits.Budget()
    b.calc_left = 1000
    with limits.call_budget(b):
        got = calc("SUM(ROW(1:1000)*COLUMN(A:CV)*123456789012)")
    s = 0.0
    for i in range(1, 1001):
        for j in range(1, 101):
            s += float(i) * j * 123456789012.0
    assert got == s
    with limits.call_budget(b):
        got = calc("SUM(ROW(1:1048576)*COLUMN(A:XFD))")  # Excel も持てない大きさ：正確な和
    assert got == float((1048576 * 1048577 // 2) * (16384 * 16385 // 2))
