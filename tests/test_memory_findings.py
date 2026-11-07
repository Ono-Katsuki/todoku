"""メモリと時間の監査（tests/audit_findings/mem）で見つけた所：宣言された数を信じない・2 乗にしない。小さな入力と、小さくした上限で確かめる。"""

from __future__ import annotations

import io
import os
import random
import sys
import time
import zipfile
from pathlib import Path

import pytest

pytest.importorskip("fcntl", reason="評価の道具は POSIX（fcntl）を使う。Windows では動かさない")

from todoku import api, limits
from todoku.limits import Budget, LimitExceeded, call_budget

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "audit_findings" / "mem"))
os.environ.setdefault("CRAFT_ONLY", "1")
import crafted  # noqa: E402

from test_example import make_xlsx  # noqa: E402


def _case(tmp_path, name: str, ext: str) -> str:
    p = tmp_path / f"{name}.{ext}"
    p.write_bytes(crafted.CASES[name][0]())
    return str(p)


def _sheet_cells(rows: int, row) -> bytes:
    return crafted.xlsx(crafted.sheet(cells="".join(row(i) for i in range(3, rows + 3))))


# ====================================================================== 数式の参照を区間で持つ（積み上げの合計）
def _naive_readers(boxes: list[list[tuple]], first: int):
    out: dict = {}
    for order, bs in enumerate(boxes):
        seen = set()
        for r1, c1, r2, c2 in bs:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) in seen:
                        continue
                    seen.add((r, c))
                    d = out.setdefault((r, c), {"cells": [], "n": 0})
                    d["n"] += 1
                    if len(d["cells"]) < first:
                        d["cells"].append(f"S!A{order + 1}")
    return out


@pytest.mark.parametrize("seed", range(6))
def test_readers_answers_like_cell_by_cell(seed):
    from todoku.xlsx.readers import Readers

    rng = random.Random(seed)
    boxes = []
    for _ in range(40):
        one = []
        for _ in range(rng.randint(1, 3)):  # 同じ数式の重なる範囲も 1 つに数える
            r1, c1 = rng.randint(1, 20), rng.randint(1, 8)
            one.append((r1, c1, r1 + rng.randint(0, 15), c1 + (rng.randint(0, 3) if seed % 3 else rng.randint(60, 80))))
        boxes.append(one)
    rd = Readers(5, 10**6, lambda *a: None)
    for order, bs in enumerate(boxes):
        o = rd.add_formula("S", order + 1, 1)
        for b in bs:
            rd.add_box(*b, o)
    want = _naive_readers(boxes, 5)
    assert set(rd.places()) == set(want)
    for k, v in want.items():
        assert rd.get(k) == v and k in rd
    assert rd.get((999, 999)) is None and (999, 999) not in rd


def test_readers_room_notes_wide_boxes_and_keeps_small():
    from todoku.xlsx.readers import Readers

    rd = Readers(5, 100)
    rd.add_box(1, 1, 5, 1, rd.add_formula("S", 1, 3))
    rd.add_box(1, 2, 1000, 2, rd.add_formula("S", 2, 3))
    assert sorted(rd.places()) == [(r, 1) for r in range(1, 6)]
    assert rd.wide == [(1, 2, 1000, 2)]  # 広い範囲は範囲の所（xlsx/places.py）
    assert rd.get((900, 2))["n"] == 1  # 広げなくても、読む数式は答えられる


def test_running_totals_model_is_interval_based(tmp_path):
    # 累計 SUM(C$3:Ci) が 3000 行：マスごとの控えなら 450 万。1 行の中の数と、数式のつながりは正しいまま
    n = 3000
    data = _sheet_cells(n, lambda i: f'<row r="{i}"><c r="C{i}"><v>1</v></c><c r="D{i}"><f>SUM(D$2:D{i - 1})+C{i}+B{i}</f></c></row>')
    p = tmp_path / "t.xlsx"
    p.write_bytes(data)
    t0 = time.time()
    m = api.calc_model(str(p))
    assert time.time() - t0 < 30
    assert isinstance(m, dict)


def test_model_lists_cap_with_counts():
    from todoku.xlsx import model as xm

    idx = xm._Index({("S", r, 1) for r in range(1, 201)})
    spans = idx.spans("S", 1, 1, 200, 1)
    assert spans == [(0, 200)]
    flags = bytearray(b"\x01" * 200)
    nxt, upto = xm._marks(flags)
    got, total = xm._listed(spans, nxt, lambda i: f"{i:03d}", upto)
    assert total == 200 and len(got) == xm.LISTED and got == sorted(got)


def test_model_small_results_are_exact():
    # 小さな様式：入力 → 処理 → 結果が、いままでと同じに決まる
    from todoku.xlsx.model import _Index, _Stabber, _merged

    assert _merged([(5, 9), (0, 2), (2, 3), (8, 12)]) == [(0, 3), (5, 12)]
    st = _Stabber(16)
    st.add(0, 8, 100)
    st.add(4, 16, 200)
    assert sorted(st.take(5)) == [100, 200]
    assert st.take(5) == []  # 1 度取ったものは消える
    idx = _Index({("A", 1, 1), ("A", 3, 1), ("A", 2, 2), ("B", 1, 1)})
    got = idx.spans("A", 1, 1, 3, 2)
    assert len(got) == 1 and got[0][1] - got[0][0] == 3  # 隣の列まで続く区間は 1 つ
    assert idx.spans("A", 2, 1, 2, 1) == []


@pytest.mark.parametrize("seed", range(5))
def test_stabber_matches_brute_force(seed):
    from todoku.xlsx.model import _Stabber

    rng = random.Random(seed)
    n = rng.randint(1, 60)
    iv = []
    st = _Stabber(n)
    for who in range(40):
        a = rng.randrange(n)
        b = rng.randint(a + 1, n)
        iv.append((a, b, who))
        st.add(a, b, who)
    taken = set()
    for p in rng.sample(range(n), n):
        got = set(st.take(p))
        assert got - taken == {w for a, b, w in iv if a <= p < b} - taken
        taken |= got
    assert taken == {w for _a, _b, w in iv}


def test_tag_end_handles_quotes_and_unterminated():
    from todoku.docx.xmlpos import _tag_end

    assert _tag_end(b'<a b="x>y" c=\'>\'/>z', 0) == 18
    assert _tag_end(b'<a b="x', 0) == 7


def test_many_rows_of_blanks_sections_and_heights_stay_linear(tmp_path):
    # 数式の列と「？」のある行が 4000 行：見出しの探索・遠い見出しまでの高さを、行ごとに上まで足さない
    p = tmp_path / "t.xlsx"
    p.write_bytes(_sheet_cells(4000, lambda i: f'<row r="{i}"><c r="A{i}" t="inlineStr"><is><t>品{i}</t></is></c><c r="D{i}"><f>B{i}*2</f></c></row>'))
    t0 = time.time()
    assert api.summary(str(p))["sheets"]
    assert time.time() - t0 < 30


def test_rows_and_cols_mm_match_the_plain_sum():
    from todoku.xlsx.window import col_mm, cols_mm, row_mm, rows_mm
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(make_xlsx({"S": ({"A1": 1}, [])}))
    sh = wb.sheets[0]
    sh.row_heights.update({3: 30.0, 500: 7.5, 900: 40.0})
    sh.hidden_rows.update({10, 600})
    sh.col_widths.update({2: 20.0, 100: 3.0})
    sh.hidden_cols.add(5)
    for a, b in ((1, 5), (1, 1029), (2, 3072), (499, 1525)):
        assert rows_mm(sh, a, b) == pytest.approx(sum(row_mm(sh, r) for r in range(a, b + 1)))
        assert cols_mm(wb, sh, a, b) == pytest.approx(sum(col_mm(wb, sh, c) for c in range(a, b + 1)))


def test_existing_in_uses_cells_and_merges_only():
    from todoku.xlsx.workbook import Cell, Sheet

    sh = Sheet("S", "x")
    sh.cells[(5, 2)] = Cell("S", 5, 2)
    sh.cells[(1, 1)] = Cell("S", 1, 1)
    sh.merges.append((9, 3, 9, 4))
    assert sh.existing_in(1, 1, 1048576, 16384) == [(1, 1), (5, 2), (9, 3)]
    assert sh.existing_in(2, 1, 8, 16384) == [(5, 2)]
    sh.cells[(6, 2)] = Cell("S", 6, 2)  # 増えれば索引を作り直す
    assert sh.existing_in(2, 2, 8, 2) == [(5, 2), (6, 2)]


# ====================================================================== 表計算の宣言
@pytest.mark.parametrize("name", ["xlsx_array_ref_all", "xlsx_merge_plus_validation", "xlsx_lastcell_plus_formula_all", "xlsx_control_anchor_all",
                                  "xlsx_cols_many", "xlsx_table_all", "xlsx_chart_pt_idx", "xlsx_chart_ptcount", "xlsx_chart_multilvl_ptcount",
                                  "xlsx_chartex_idx"])
def test_declared_size_does_not_take_memory(tmp_path, name):
    p = _case(tmp_path, name, "xlsx")
    t0 = time.time()
    assert api.summary(p)["sheets"]
    assert time.time() - t0 < 30


def test_chart_sparse_points_keep_values_and_numbers():
    from todoku.embeds.chart import Chart, Series, chart_dict

    s = Series("s")
    s.vals = {0: "1", 2_000_000_000: "2"}
    d = chart_dict(Chart(series=[s]), 1)["series"][0]
    assert d["values"] == ["1", "2"] and d["point_numbers"] == [1, 2_000_000_001] and d["points_declared"] == 2_000_000_001
    s2 = Series("s")
    s2.vals = {0: "1", 3: "2"}
    d2 = chart_dict(Chart(series=[s2]), 1)["series"][0]
    assert d2["values"] == ["1", "", "", "2"] and "point_numbers" not in d2  # ふつうのグラフは、いままでどおり点の順


def test_multilevel_labels_same_as_before():
    import xml.etree.ElementTree as ET

    from todoku.embeds.chart import _cache_values

    C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
    cache = ET.fromstring(f'<c:multiLvlStrCache {C}><c:ptCount val="3"/><c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:lvl>'
                          '<c:lvl><c:pt idx="0"><c:v>X</c:v></c:pt><c:pt idx="2"><c:v>Y</c:v></c:pt></c:lvl></c:multiLvlStrCache>')
    assert _cache_values(cache, False, False) == ({0: "X ／ a", 1: "X", 2: "Y ／ c"}, 3)


# ====================================================================== 式：字の上限と配列の上限
class _Ctx:
    def __init__(self):
        self.vals = {}

    def cell_value(self, sheet, r, c):
        return self.vals.get((r, c))

    def bounds(self, sheet):
        return (1, 1)

    def has_sheet(self, sheet):
        return True

    def resolve_name(self, *a):
        return None


def _calc(text, array=True):
    from todoku.formula.evaluator import Evaluator
    from todoku.formula.parser import parse

    return Evaluator(_Ctx(), "S", 1, 1, 0, 0, array).ev(parse(text))


@pytest.mark.parametrize("text", ['REPT("x",4000000000)', 'REPT(REPT("x",1000000),5000)', 'SUBSTITUTE(REPT("a",100),"a",REPT("b",1000))',
                                  'REPT("x",32767)&"y"', 'REPLACE(REPT("x",32767),1,1,"yy")'])
def test_text_over_cell_limit_is_value_error(text):
    assert getattr(_calc(text), "code", None) == "#VALUE!"


def test_text_at_cell_limit_is_kept():
    assert _calc('LEN(REPT("x",32767))') == 32767
    assert _calc('LEN(SUBSTITUTE("ab","a",REPT("b",32766)))') == 32767


@pytest.mark.parametrize("text", ["SEQUENCE(1048577)", "SEQUENCE(1,16385)", "EXPAND(1,1048577)", "EXPAND(1,1,16385)"])
def test_arrays_over_the_sheet_are_num(text):
    assert getattr(_calc(text), "code", None) == "#NUM!"


@pytest.mark.parametrize("text", ["SEQUENCE(1048576,16384)", "EXPAND(1,1048576,16384,0)", "ROW(1:1048576)*COLUMN(A:XFD)", "WRAPROWS(SEQUENCE(2),16384000)"])
def test_arrays_that_fit_the_sheet_cost_no_memory(text):
    from todoku.formula.values import Arr, ExcelError

    r = _calc(text)
    assert isinstance(r, Arr) or (isinstance(r, ExcelError) and r.code == "#NUM!")


def test_small_arrays_are_unchanged():
    assert _calc("SUM(SEQUENCE(1000))") == 500500


# ====================================================================== 要素の予算
def test_count_part_counts_each_part_once():
    b = Budget(items=100)
    with call_budget(b):
        b.count_part("a.xml", 60)
        b.count_part("a.xml", 60)  # 検査で読み直しても数えない
        b.count_part("a.xml", 70)
        with pytest.raises(LimitExceeded, match="b.xml"):
            b.count_part("b.xml", 40, "b.xml の要素")


def test_zip_part_elements_are_counted_before_the_tree():
    blob = io.BytesIO()
    with zipfile.ZipFile(blob, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("a.xml", "<r>" + "<c/>" * 10000 + "</r>")
    from todoku.zipraw import RawZip

    with call_budget(Budget(items=1000)):
        with pytest.raises(LimitExceeded, match="a.xml"):
            RawZip(blob.getvalue()).read("a.xml")
    with call_budget(Budget(items=100000)):
        assert len(RawZip(blob.getvalue()).read("a.xml")) > 1000


def test_csv_cells_are_counted(tmp_path):
    from todoku.datafile.csvfile import CsvFile

    with call_budget(Budget(items=1000)):
        with pytest.raises(LimitExceeded, match="CSV"):
            CsvFile(b"," * 5000 + b"\n", "a.csv")


# ====================================================================== 窓の控え
def test_big_window_is_held_without_a_copy(monkeypatch, tmp_path):
    monkeypatch.setattr(api, "_WINDOWS_COPY", 0)
    api._WINDOWS.clear()
    p = tmp_path / "a.csv"
    p.write_text("id,name\n1,山田\n", encoding="utf-8")
    first = api.window(str(p))
    held = next(iter(api._WINDOWS.values()))[1]
    assert isinstance(held, api._Held)
    assert api.window(str(p)) is first  # 写さない
    assert api.fields(str(p))
    api._WINDOWS.clear()


# ====================================================================== 繰り返しの宣言（ODS・ODT）・結合（HWPX）・.doc
def test_ods_hidden_columns_stop_at_the_sheet_edge(tmp_path):
    p = _case(tmp_path, "ods_column_hidden_many", "ods")
    from todoku.odf.sheet import OdsWorkbook

    wb = OdsWorkbook(Path(p).read_bytes(), "a.ods")
    assert all(c <= 16384 for sh in wb.sheets for c in sh.hidden_cols)


def test_odt_repeated_rows_are_bounded_and_reported(tmp_path):
    p = _case(tmp_path, "odt_rows_cols_amp", "odt")
    _k, doc, _parts = api.window(p)
    assert not doc.get("unread") and "繰り返す" in doc["repeats"][0]


def test_hwpx_row_span_stops_at_the_table(tmp_path):
    p = _case(tmp_path, "hwpx_rowspan", "hwpx")
    t0 = time.time()
    assert api.summary(p)
    assert time.time() - t0 < 30


def test_doc_character_count_is_bound_by_the_stream(tmp_path):
    p = _case(tmp_path, "doc_ccp_huge", "doc")
    t0 = time.time()
    assert api.summary(p)
    assert time.time() - t0 < 30


def test_xls_repeated_colinfo_is_read_once(tmp_path):
    p = _case(tmp_path, "xls_colinfo_many", "xls")
    t0 = time.time()
    assert api.summary(p)
    assert time.time() - t0 < 30


# ====================================================================== 文字のファイル
def test_csv_fields_stop_at_the_window_limit_and_share_context():
    from todoku.datafile.csvfile import CsvFile
    from todoku.datafile.window import MAX_FIELDS, build_csv

    win = build_csv(CsvFile(b"a,b\n" + b",\n" * (MAX_FIELDS * 3), "a.csv"))
    assert len(win["fields"]) == MAX_FIELDS and win["fields_truncated"] is True
    wide = build_csv(CsvFile(b"," * 500 + b"\n" + b"," * 500 + b"\n", "b.csv"))
    assert wide["fields"][0]["context"] is wide["fields"][1]["context"]


def test_in_order_search_finds_out_of_order_and_misses():
    from todoku.datafile.audit import _InOrder

    s = _InOrder("aa\nbb\ncc\n")
    assert s.has("bb") and s.has("cc") and s.has("aa") and not s.has("zz")


def test_txt_anchor_gives_up_on_repetition_quickly():
    from todoku.txt.audit import anchored_check

    blob = ("同じ行です。" * 10 + "\n").encode() * 20000
    t0 = time.time()
    chk = anchored_check(blob, "a.txt")
    assert chk.ok and time.time() - t0 < 30


# ====================================================================== RTF
def test_rtf_runs_of_unicode_chars_share_one_seg_and_fill_exact_bytes():
    from todoku.rtf.document import RtfDocument, paragraphs

    src = b"{\\rtf1\\ansi{\\fonttbl{\\f0 MS Mincho;}}\\pard\\u12371?\\u12428?abc\\par}"
    d = RtfDocument(src)
    p = paragraphs(d)[0]
    assert p.text == "これabc" and len(p.segs) == 1
    node = p.segs[0].node
    assert [src[a:b] for a, b in zip(node.starts, node.ends)] == [b"\\u12371?", b"\\u12428?", b"a", b"b", b"c"]


def test_rtf_plain_runs_match_char_positions():
    from todoku.rtf.document import RtfDocument, paragraphs

    src = b"{\\rtf1\\ansi\\pard abc\ndef\\par}"
    p = paragraphs(RtfDocument(src))[0]
    assert p.text == "abcdef"
    pos = [(a, b) for s in p.segs for a, b in zip(s.node.starts, s.node.ends)]
    assert [src[a:b] for a, b in pos] == [b"a", b"b", b"c", b"d", b"e", b"f"]


def test_rtf_table_grid_is_looked_up():
    from todoku.rtf.document import RtfDocument

    rows = b"".join(b"\\trowd" + b"".join(b"\\cellx%d" % (1000 + r * 30 + i) for i in range(1, 30)) + b"\\intbl x\\cell\\row\n" for r in range(300))
    t0 = time.time()
    assert RtfDocument(b"{\\rtf1\\ansi " + rows + b"}")
    assert time.time() - t0 < 30


# ====================================================================== マクロ
def _macro(body: str):
    from todoku.vba import _PROC, Module, Project
    from todoku.vbarun import run_macro
    from todoku.xlsx.workbook import Workbook

    src = "Sub Test()\n" + body + "\nEnd Sub\n"
    wb = Workbook(make_xlsx({"入力": ({"A1": 1, "A2": 2}, [])}))
    wb.has_macros = True
    wb._vba = Project(932, "VBAProject", [Module("Module1", "Module1", 0, "procedural", 'Attribute VB_Name = "Module1"\n' + src,
                                                 [m.group(2) for m in _PROC.finditer(src)])])
    wb.sheets[0].code_name = "Sheet1"
    return run_macro(wb, "Test")


@pytest.mark.parametrize("body", ["Do\nLoop", 'Dim c\nFor Each c In Range("A1:XFD1048576")\nNext', "Dim i\nFor i = 1 To 2000000000\nNext"])
def test_macro_empty_loops_reach_the_limit(body):
    from todoku.vbarun.interp import Unsupported

    t0 = time.time()
    with pytest.raises(Unsupported, match="繰り返し"):
        _macro(body)
    assert time.time() - t0 < 60


def test_macro_loop_with_body_is_counted_once_per_statement():
    assert _macro("Dim i, n\nFor i = 1 To 1000000\nn = n + 1\nNext\nRange(\"B1\").Value = n") is not None


def test_macro_hasformula_on_the_whole_sheet_is_immediate():
    t0 = time.time()
    _macro('Dim v\nv = Range("A1:XFD1048576").HasFormula')
    assert time.time() - t0 < 5


@pytest.mark.parametrize("body", ['Dim s\ns = String(10000000, "a")\ns = Replace(s, "a", String(1000, "b"))',
                                  'Dim s\ns = String(10000000, ",")\nDim v\nv = Split(s, ",")'])
def test_macro_amplification_is_checked_before_making_it(body):
    from todoku.vbarun.interp import Unsupported

    with pytest.raises(Unsupported):
        _macro(body)


def test_macro_replace_still_works():
    from todoku.vba import _PROC  # noqa: F401

    _macro('Dim s\ns = Replace("aXa", "a", "bb")\nRange("B1").Value = s')


# ====================================================================== 壊れた入力：理由を伝える
def test_broken_xml_inside_xlsx_is_api_error(tmp_path):
    z0 = zipfile.ZipFile(io.BytesIO(crafted.XL))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i in z0.infolist():
            z.writestr(i, b"<worksheet><broken" if i.filename == crafted.SH else z0.read(i.filename))
    p = tmp_path / "a.xlsx"
    p.write_bytes(out.getvalue())
    with pytest.raises(api.ApiError, match="ParseError"):
        api.window(str(p))
