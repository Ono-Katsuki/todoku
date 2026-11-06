"""採点 v2（scripts/eval/v2）の試験。作った小さな書類だけを使う。とっておいた書類・結果には触れない。
比べる道具の venv（TODOKU_BASELINE_PY）・LibreOffice がないときは、その試験を飛ばす。"""

import datetime
import importlib
import importlib.util
import io
import json
import os
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import pytest

V2 = Path(__file__).resolve().parents[1] / "scripts" / "eval" / "v2"
sys.path.insert(0, str(V2))
for m in ("common", "truth", "read_compare", "runner", "stage2_select", "write_compare", "make_signed"):  # v1 と同じ名前の部品が先に読まれていれば外す
    sys.modules.pop(m, None)
common = importlib.import_module("common")
T = importlib.import_module("truth")
RC = importlib.import_module("read_compare")
RN = importlib.import_module("runner")
SS = importlib.import_module("stage2_select")
assert Path(common.__file__).parent == V2
# ほかの試験が読む v1 の同じ名前の部品と混ざらないように、名前の表から外す（ここでは手元の参照を使う）
for m in ("common", "truth", "read_compare", "runner", "stage2_select", "write_compare", "make_signed"):
    sys.modules.pop(m, None)
WC = None  # write_compare は下で読む
sys.path.remove(str(V2))

BPY = os.environ.get("TODOKU_BASELINE_PY", "/private/tmp/claude-501/-Users-onokatsuki-Documents-GitHub-claude---/fdcc494c-06ce-4f3c-bc44-7cab0977bf2d/scratchpad/m/ev/baseline_venv/bin/python")
need_base = pytest.mark.skipif(not Path(BPY).exists(), reason="比べる道具の venv がない")
need_lo = pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice がない")


# ---------------------------------------------------------------- T1：CSV の正解はマスの字
def test_csv_cells_are_values_not_syntax():
    cells, d = T.csv_cells('a,"b,c","He said ""hi"""\r\n1,2,3\r\n')
    assert d == ","
    assert cells == ["a", "b,c", 'He said "hi"', "1", "2", "3"]


def test_csv_table_reader_and_raw_reader_both_score_perfect():
    raw = 'name;city\r\n"Müller, A";Zürich\r\n'
    cells, d = T.csv_cells(raw)
    assert d == ";"
    # 生の字を返す道具も、表に読んだ道具も、中身は全部ある
    assert RC.recalls(cells, raw)["char"] == 1.0 and RC.recalls(cells, raw)["bigram"] == 1.0
    table = "| name | city |\n| Müller, A | Zürich |"
    assert RC.recalls(cells, table)["char"] == 1.0 and RC.recalls(cells, table)["bigram"] == 1.0


def test_csv_missing_cell_still_fails():
    """弱めていない：マスが 1 つ欠ければ落ちる。v1 の生の行の正解でなくても同じ。"""
    cells, _ = T.csv_cells("name,city\nAlice,Paris\nBob,Berlin\n")
    assert RC.recalls(cells, "name city Alice Paris Bob")["char"] < RC.EXACT


def test_v1_truth_was_unfair_to_parsed_readers():
    """v1 の形（生の行）では、中身の全部がある表が落ちる。v2 で通る。これが直した理由。"""
    raw = 'a,"b,c",d\n1,2,3\n'
    v1_units = raw.splitlines()
    table = "a | b,c | d\n1 | 2 | 3"
    assert RC.recalls(v1_units, table)["char"] < RC.EXACT
    assert RC.recalls(T.csv_cells(raw)[0], table)["char"] == 1.0


def test_csv_delimiter_tsv_and_decimal_comma():
    assert T.csv_delimiter("a\tb\n1\t2\n", ".tsv") == "\t"
    assert T.csv_delimiter("x;y\n1,5;2,5\n3,5;4,5\n") == ";"
    # 引用符の中のタブ・改行に引きずられない
    assert T.csv_delimiter('a,b,c,d,e\n1,2,3,4,"x\ty\nz\tw\nu\tv"\n') == ","


def test_semicolon_csv_comma_splitter_loses_decimal_commas():
    """; 区切りの小数のコンマは中身。コンマで切った道具は、その字を落とす（v2 でも落ちる）。"""
    cells, d = T.csv_cells("a;b\n85,3;73,3\n85,0;87,9\n")
    assert d == ";"
    assert RC.recalls(cells, "a;b 85 3 73 3 85 0 87 9".replace(";", " "))["char"] < RC.EXACT


# ---------------------------------------------------------------- T2：文字コード
def test_decode_valid_utf8_is_never_guessed_as_cp949():
    b = "Zürich;Genève\n".encode()
    t, how = T.decode_text(b, "text/csv")
    assert (t, how) == ("Zürich;Genève\n", "utf8")
    mojibake = b.decode("cp949", "replace")
    assert RC.recalls([t], mojibake)["char"] < RC.EXACT  # 誤った復号の道具は落ちる（v1 はこれが正解だった）
    assert RC.recalls([t], t)["char"] == 1.0


def test_decode_bom_declared_unverified():
    assert T.decode_text("é,x".encode("utf-8-sig"))[1] == "bom"
    assert T.decode_text("é,x".encode("utf-16"))[1] == "bom"
    t, how = T.decode_text("Peñalolén".encode("latin-1"), "text/csv; charset=latin1")
    assert how == "declared" and t == "Peñalolén"
    t, how = T.decode_text("Pe\x96a €".encode("cp1252", "replace"), "text/csv;charset=ISO-8859-1")
    assert how == "declared" and "€" in t  # WHATWG：iso-8859-1 の表示は windows-1252
    t, how = T.decode_text("Peñalolén".encode("latin-1"), "text/csv")
    assert how == "unverified" and t.replace("\ufffd", "") == "Pealoln"


def test_unverified_splits_units_at_unknown_bytes(tmp_path):
    """分からない字をはさんだ前後の字は、2 字の並びの正解にしない。"""
    p = tmp_path / "x.csv"
    p.write_bytes("Pe\xf1a,b\n1,2\n".encode("latin-1"))
    t = T.build(str(p), "csv", "d2", tmp_path, sys.executable, None, None, "text/csv")
    assert "Pe" in t["units"] and "a" in t["units"]
    assert RC.recalls(t["units"], "Peña b 1 2")["bigram"] == 1.0


def test_unverified_scores_ascii_only_for_every_tool(tmp_path):
    p = tmp_path / "x.csv"
    p.write_bytes("name,city\nPeñalolén,Chile\n".encode("latin-1"))
    t = T.build(str(p), "csv", "d1", tmp_path, sys.executable, None, None, "text/csv")
    assert t["encoding_source"] == "unverified"
    assert "ñ" not in "".join(t["units"]) and "\ufffd" not in "".join(t["units"])
    assert RC.recalls(t["units"], "name city Peñalolén Chile")["char"] == 1.0
    assert RC.recalls(t["units"], "name city Chile")["char"] < RC.EXACT


# ---------------------------------------------------------------- 向きの制御
def test_bidi_controls_are_not_content():
    assert common.norm("\u202bسلام\u202c \ufeffabc\u200f") == "سلامabc"
    units = ["\u202bسلام\u202c", "\u202adef\u202c"]
    assert RC.recalls(units, "سلام def")["char"] == 1.0
    assert RC.recalls(units, "سلام")["char"] < RC.EXACT  # 本当の字が欠ければ落ちる


def test_pdf_units_split_at_direction_runs_but_keep_order_inside():
    t = {"kind": "pdf", "units": ["\u202bسلام\u202c\u202a123\u202c\u202bعليكم\u202c"]}
    units = RC._recall_units(t)
    assert units == ["سلام", "123", "عليكم"]
    assert RC.recalls(units, "عليكم 123 سلام")["bigram"] == 1.0  # 並びの順は問わない
    assert RC.recalls(units, "سلام 123 " + "عليكم"[::-1])["bigram"] < RC.EXACT  # 並びの中の字の順が逆なら落ちる


def test_zero_width_joiner_stays_content():
    assert common.norm("می\u200cخواهم") != common.norm("میخواهم")  # ZWNJ は意味のある字。除かない


# ---------------------------------------------------------------- R1・R3
def test_exact_read_bigram_zero_fails_and_none_passes():
    base = {"opened": True, "ref_chars": 10, "char": 1.0}
    assert RC.exact_read({**base, "bigram": 0.0}) is False
    assert RC.exact_read({**base, "bigram": None}) is True
    assert RC.exact_read({**base, "bigram": 0.996}) is True
    assert RC.exact_read({**base, "bigram": 0.99}) is False


def test_image_text_rule():
    scan = {"kind": "pdf", "pages_text_chars": [10, 0], "ocr_units": ["x" * 2000]}
    text = {"kind": "pdf", "pages_text_chars": [1900, 0], "ocr_units": ["x" * 2000]}
    assert RC.image_text(scan) is True
    assert RC.image_text(text) is False
    assert RC.image_text({"kind": "pdf", "pages_text_chars": [10], "ocr_units": []}) is False  # 字がどちらにもない
    assert RC.image_text({"kind": "pdf", "pages_text_chars": [10]}) is False  # 文字認識なし：決めない
    s = RC.score_tool({**scan, "units": ["CamScanner"], "path": "x", "probes": {"pages": 1}}, "CamScanner", "pypdf", {})
    assert s["exact_text_layer"] is True and s["exact_read"] is False and s["image_text"] is True


# ---------------------------------------------------------------- T3・T4・T5：画面の字
@need_base
@need_lo
def test_libreoffice_truth_is_what_is_shown(tmp_path):
    code = """
import openpyxl, sys, datetime
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S"
v = datetime.datetime(2024, 1, 15)
rows = [("a", 51583.1377, "#,##0"), ("b", 0.2695, "0.0%"), ("c", v, "dd/mm/yyyy"), ("d", v, "d-mmm-yy"), ("e", datetime.datetime(2024, 3, 20), "mm-dd-yy")]
for i, (k, val, f) in enumerate(rows, 1):
    ws.cell(i, 1, k); c = ws.cell(i, 2, val); c.number_format = f
wb.save(sys.argv[1])
"""
    x = tmp_path / "t.xlsx"
    subprocess.run([BPY, "-I", "-c", code, str(x)], check=True, timeout=60)
    os.environ["EVAL_LOCK"] = str(tmp_path / "lock")
    lo = T.LibreOffice(tmp_path)
    t = T.build(str(x), "xlsx", "doc1", tmp_path, BPY, lo, None)
    d = t["sheets"][0]["display"]
    assert d["B1"] == "51,583"  # v1 は 51583.1377
    assert d["B2"] == "27.0%"  # v1 は 26.95%
    assert d["B3"] == "15/01/2024"  # v1 は 2024/01/15
    assert d["B4"] == "15-Jan-24"  # ja-JP では 15-1月-24
    assert d["B5"] in ("3/20/2024", "03/20/24", "03/20/2024")
    # 環境で変わる日付は字の比べから除く。ほかは残す
    assert t["locale_dates_excluded"] == 1 and "51,583" in t["units"] and d["B5"] not in t["units"]
    # 日付の並びのどれでも画面の字に数える。違う日付は数えない
    cat = lambda text: RC.score_sheet(t, RC.Hay(text), text)["displayed"]  # noqa: E731
    base = "51,583\n27.0%\n15/01/2024\n15-Jan-24\n"
    for shown in ("2024-03-20", "2024/3/20", "3/20/2024", "20.03.2024", "2024年3月20日", "2024-03-20T00:00:00"):
        assert cat(base + shown) == {"n": 5, "ok": 5}, shown
    assert cat(base + "2024-03-21") == {"n": 5, "ok": 4}
    # 元の値だけの道具は落ちる
    assert cat("51583.1377\n0.2695\n2024-01-15\n2024-01-15\n2024-03-20")["ok"] <= 2


def test_value_sensitivity_passes_raw_reader_but_not_primary():
    t = {"kind": "xlsx", "path": "x", "units": ["Total", "51,583"],
         "sheets": [{"name": "S", "display": {"A1": "Total", "B1": "51,583"}, "fresh": {"A1": "Total", "B1": "51,583"}}],
         "probes": {"sheets": [{"name": "S", "state": "visible", "cells": {"A1": {"raw": "Total", "t": "s"}, "B1": {"raw": "51583.1377", "t": "n", "fk": "number", "fmt": "#,##0"}}, "merges": []}]}}
    raw = RC.score_tool(t, "Total\t51583.1377", "openpyxl", {})
    assert raw["exact_read"] is False and raw["exact_value"] is True
    wrong = RC.score_tool(t, "Total\t99999.1377", "openpyxl", {})
    assert wrong["exact_value"] is False
    shown = RC.score_tool(t, "Total\t51,583", "openpyxl", {})
    assert shown["exact_read"] is True and shown["exact_value"] is True


def test_date_count_does_not_match_inside_longer_numbers():
    h = RC.Hay("no 12024-01-155 here")
    assert h.date_count(24, 1, 15) == 0
    assert RC.serial_date("45306") == (24, 1, 15)
    assert RC.serial_date("abc") is None


# ---------------------------------------------------------------- 書く：中身だけを「消えた」に数える
def _xlsx_like_excel(path: Path, formula: bool = False, cached: bool = True, extra: dict | None = None):
    """本物の Excel の保存にある、中身でない部品（extLst・absPath・calcChain・printerSettings・フォルダの項目）を持つ小さな xlsx。"""
    wbx = ('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
           'xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac">'
           '<mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="C:\\Users\\x\\"/></mc:Choice></mc:AlternateContent>'
           '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>'
           '<extLst><ext uri="{1}"><xcalcf:calcFeatures xmlns:xcalcf="urn:x"/></ext></extLst></workbook>')
    f = '<c r="B1"><f>1+1</f>' + ("<v>2</v>" if cached else "") + "</c>" if formula else ""
    sheet = ('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">'
             f'<c r="A1" t="inlineStr"><is><t>Name</t></is></c>{f}</row></sheetData><pageMargins left="1" right="1" top="1" bottom="1" header="0" footer="0"/></worksheet>')
    styles = ('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/></font></fonts>'
              '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>'
              '<extLst><ext uri="{2}"><x14:slicerStyles xmlns:x14="urn:y" defaultSlicerStyle="z"/></ext></extLst></styleSheet>')
    parts = {"[Content_Types].xml": ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                                     '<Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>'
                                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                                     '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                                     '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'),
             "_rels/.rels": ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
                             'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
             "xl/workbook.xml": wbx, "xl/worksheets/sheet1.xml": sheet, "xl/styles.xml": styles,
             "xl/_rels/workbook.xml.rels": ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
                                            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" '
                                            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
             "xl/calcChain.xml": '<?xml version="1.0"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
             "xl/printerSettings/printerSettings1.bin": "x", "[trash]/0000.dat": "x", "xl/": ""}
    parts.update(extra or {})
    with zipfile.ZipFile(path, "w") as z:
        for n, b in parts.items():
            z.writestr(n, b)


def _write_check(tmp_path, src: Path, val="TEST-12345") -> dict:
    out = tmp_path / ("out" + src.suffix)
    tg = json.dumps({"id": "S!C5"})
    kind = src.suffix[1:]
    r = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "write", "openpyxl", str(src), str(out), kind, tg, val], capture_output=True, text=True, timeout=120)
    assert "error" not in json.loads(r.stdout.splitlines()[-1]), r.stdout
    r = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "check", str(src), str(out), kind, tg, val], capture_output=True, text=True, timeout=120)
    return json.loads(r.stdout.splitlines()[-1])


@need_base
def test_openpyxl_not_failed_for_non_content_parts(tmp_path):
    src = tmp_path / "a.xlsx"
    _xlsx_like_excel(src)
    c = _write_check(tmp_path, src)
    assert c["landed"] is True
    assert c["lost"] == {}, c  # v1：parts・extLst・alternate_content・printer_settings で必ず落ちていた


@need_base
def test_real_loss_still_counts(tmp_path):
    """弱めていない：図形が消えれば落ちる。"""
    drawing = '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:sp macro=""/></xdr:wsDr>'
    src = tmp_path / "b.xlsx"
    _xlsx_like_excel(src, extra={"xl/drawings/drawing1.xml": drawing})
    c = _write_check(tmp_path, src)
    assert c["lost"].get("shapes") == 1 and c["lost"].get("drawings") == 1


@need_base
def test_cached_formula_value_loss_is_counted(tmp_path):
    """W5：openpyxl の保存で、数式の保存された値は消える。v1 は <v></v> を値ありに数えて見逃した。"""
    src = tmp_path / "c.xlsx"
    _xlsx_like_excel(src, formula=True, cached=True)
    c = _write_check(tmp_path, src)
    assert c["lost"].get("formula_cached_values") == 1, c


@need_base
def test_byte_keeping_edit_is_not_counted_as_loss(tmp_path):
    """同じ確かめで、部品をそのまま残す書き方（とどくの形）は何も消えない。"""
    src = tmp_path / "d.xlsx"
    _xlsx_like_excel(src, formula=True, cached=True)
    out = tmp_path / "d_out.xlsx"
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w") as zo:
        for it in zi.infolist():
            b = zi.read(it.filename)
            if it.filename == "xl/worksheets/sheet1.xml":
                b = b.replace(b"</row>", b'<c r="C1" t="inlineStr"><is><t>TEST-12345</t></is></c></row>')
            zo.writestr(it, b)
    tg = json.dumps({"id": "S!C1"})
    r = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "check", str(src), str(out), "xlsx", tg, "TEST-12345"], capture_output=True, text=True, timeout=120)
    c = json.loads(r.stdout.splitlines()[-1])
    assert c["landed"] is True and c["lost"] == {}


@need_base
def test_docx_zip_folder_entries_and_trash_are_not_content(tmp_path):
    code = """
import docx, sys, zipfile
d = docx.Document(); t = d.add_table(rows=1, cols=1); d.save(sys.argv[1])
zi = zipfile.ZipFile(sys.argv[1]); o = zipfile.ZipFile(sys.argv[2], "w")
for it in zi.infolist(): o.writestr(it, zi.read(it.filename))
o.writestr("word/", ""); o.writestr("[trash]/0001.dat", "x"); o.close()
"""
    a, b = tmp_path / "a.docx", tmp_path / "b.docx"
    subprocess.run([BPY, "-I", "-c", code, str(a), str(b)], check=True, timeout=60)
    out = tmp_path / "o.docx"
    tg = json.dumps({"id": "T1.R1.C1"})
    r = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "write", "python-docx", str(b), str(out), "docx", tg, "TEST-12345"], capture_output=True, text=True, timeout=60)
    assert "error" not in json.loads(r.stdout.splitlines()[-1]), r.stdout
    r = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "check", str(b), str(out), "docx", tg, "TEST-12345"], capture_output=True, text=True, timeout=60)
    c = json.loads(r.stdout.splitlines()[-1])
    assert c["landed"] is True and c["lost"] == {}, c


# ---------------------------------------------------------------- 実行の守り
def test_nice_prefix_at_nice_20():
    """v1 の大量の失敗の原因：nice 20 で os.nice(0) が PermissionError。v2 は読むだけ。"""
    code = f"import os,sys; sys.path.insert(0,{str(V2)!r}); import common; print(common.nice_prefix()); print(common.run_watched(['/bin/echo','ok'], limit_s=10, heavy=False)['stdout'].strip())"
    r = subprocess.run(["nice", "-n", "20", sys.executable, "-I", "-c", code], capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    assert r.stdout.split("\n")[:2] == ["[]", "ok"]
    v1 = subprocess.run(["nice", "-n", "20", sys.executable, "-I", "-c", "import os; os.nice(0)"], capture_output=True, text=True, timeout=30)
    if sys.platform == "darwin":
        assert "Operation not permitted" in v1.stderr  # 根本の原因の再現（macOS）


def test_run_watched_survives_killpg_eperm(monkeypatch):
    def fake(pid, sig):
        raise PermissionError(1, "Operation not permitted")

    monkeypatch.setattr(common.os, "killpg", fake)
    r = common.run_watched(["/bin/sleep", "2"], limit_s=0.2, heavy=False)
    assert r["reason"] == "timeout"


def test_failguard_stops_on_mass_failure():
    g = common.FailGuard(window=50, limit=0.2, min_docs=20)
    for _ in range(19):
        g.add(True)  # 20 件までは判断しない
    with pytest.raises(common.MassFailure):
        g.add(True)
    g = common.FailGuard()
    for i in range(500):
        g.add(i % 10 == 0)  # 10% は止めない


WORKER = """
import json, sys, time
from pathlib import Path
a = sys.argv
path = a[a.index("--one") + 1]; out = Path(a[a.index("--out") + 1])
sys.path.insert(0, a[a.index("--v2") + 1])
from common import atomic_json, doc_id
mode = Path(path).read_text()
if mode == "hang": time.sleep(60)
if mode == "crash": raise PermissionError(1, "Operation not permitted")
atomic_json(out / "scores" / f"{doc_id(path)}.json", {"id": doc_id(path), "path": path, "tools": {"todoku": {"opened": True}}})
"""


def _drive(tmp_path, modes: list[str], limit=30.0):
    w = tmp_path / "w.py"
    w.write_text(WORKER, encoding="utf-8")
    base = tmp_path / "out"
    (base / "scores").mkdir(parents=True, exist_ok=True)
    (tmp_path / "docs").mkdir(exist_ok=True)
    files = []
    for i, m in enumerate(modes):
        p = tmp_path / "docs" / f"d{i}.txt"
        p.write_text(m)
        files.append(str(p))
    shown = []
    rc, note = common.drive_docs(w, files, base, ["--out", str(base), "--v2", str(V2)], limit,
                                 lambda p, why: {"id": common.doc_id(p), "path": p, "unscored": why, "tools": {}},
                                 lambda rec: "bad" if rec.get("unscored") else None, lambda p, r: None, lambda i, n, p, r: shown.append(p))
    return rc, note, base, files


def test_drive_time_limit_records_unscored(tmp_path):
    rc, _n, base, files = _drive(tmp_path, ["ok", "hang", "ok"], limit=1.5)
    assert rc == 0
    rec = json.loads((base / "scores" / f"{common.doc_id(files[1])}.json").read_text())
    assert rec["unscored"] == "time_limit" and rec["tools"] == {}


def test_drive_mass_failure_aborts_and_keeps_no_scores(tmp_path):
    """v1 の 09:31 の形：全部が同じ理由で落ちる。rc は 0 にならず、記録は scores に入らず、続きからやり直せる。"""
    rc, note, base, files = _drive(tmp_path, ["crash"] * 40)
    assert rc == common.ABORT_RC and "大量の失敗" in note
    assert (base / "ABORTED").exists()
    assert not list((base / "scores").glob("*.json"))
    # 原因を直して（ここでは書類を普通に）もう一度：全部やり直される
    for f in files:
        Path(f).write_text("ok")
    rc, _n, base, _f = _drive_again(tmp_path, base, files)
    assert rc == 0 and len(list((base / "scores").glob("*.json"))) == 40 and not (base / "ABORTED").exists()


def test_mass_failure_quarantines_bad_records(tmp_path):
    """記録は書けたが中身が失敗（道具がどれも開けない）の形：大量なら止め、scores に残さない。v2 の試運転で実際に起きた形。"""
    rc, note, base, files = _drive_bad(tmp_path, 25)
    assert rc == common.ABORT_RC
    assert not list((base / "scores").glob("*.json"))
    assert len(list((base / "quarantine").rglob("*.json"))) == 20


def _drive_bad(tmp_path, n):
    w = tmp_path / "w.py"
    w.write_text(WORKER, encoding="utf-8")
    base = tmp_path / "out"
    (base / "scores").mkdir(parents=True)
    (tmp_path / "docs").mkdir()
    files = []
    for i in range(n):
        p = tmp_path / "docs" / f"d{i}.txt"
        p.write_text("ok")
        files.append(str(p))
    rc, note = common.drive_docs(w, files, base, ["--out", str(base), "--v2", str(V2)], 30, lambda p, why: {"unscored": why, "tools": {}}, lambda rec: "開けない", lambda p, r: None, lambda i, n, p, r: None)
    return rc, note, base, files


def _drive_again(tmp_path, base, files):
    w = tmp_path / "w.py"
    rc, note = common.drive_docs(w, files, base, ["--out", str(base), "--v2", str(V2)], 30, lambda p, why: {"unscored": why, "tools": {}}, lambda rec: None, lambda p, r: None, lambda i, n, p, r: None)
    return rc, note, base, files


def test_drive_single_failure_retries_then_unscored(tmp_path):
    rc, _n, base, files = _drive(tmp_path, ["crash", "ok", "ok"])
    assert rc == 3  # 1 回目：scores に入れない
    assert not (base / "scores" / f"{common.doc_id(files[0])}.json").exists()
    rc, _n = common.drive_docs(tmp_path / "w.py", files, base, ["--out", str(base), "--v2", str(V2)], 30,
                               lambda p, why: {"id": common.doc_id(p), "path": p, "unscored": why, "tools": {}}, lambda rec: None, lambda p, r: None, lambda i, n, p, r: None)
    rec = json.loads((base / "scores" / f"{common.doc_id(files[0])}.json").read_text())
    assert rec["unscored"] == "failed"


def test_cleanup_refuses_without_verified(tmp_path):
    (tmp_path / "docs").mkdir()
    (tmp_path / "docs" / "x.pdf").write_text("x")
    class A:
        out = str(tmp_path)
    assert RN.cmd_cleanup(A) == 1
    assert (tmp_path / "docs" / "x.pdf").exists()
    (tmp_path / "VERIFIED").write_text("{}")
    (tmp_path / "counts_v2.json").write_text("{}")
    assert RN.cmd_cleanup(A) == 0
    assert not (tmp_path / "docs").exists()


def test_verify_rejects_mass_failure_state(tmp_path):
    (tmp_path / "docs_map.json").write_text(json.dumps({"A": {"path": "/x/a.pdf"}, "B": {"path": "/x/b.pdf"}}))
    probs, _st = RN.verify_state(tmp_path)
    assert any("採点の記録がない" in p for p in probs)
    for sub in ("read", "write"):
        (tmp_path / sub / "scores").mkdir(parents=True)
    for n in ("a.pdf", "b.pdf"):
        did = common.doc_id(f"/x/{n}")
        (tmp_path / "read" / "scores" / f"{did}.json").write_text(json.dumps({"kind": "pdf", "tools": {"todoku": {"opened": False}, "pypdf": {"opened": False}}}))
        (tmp_path / "write" / "scores" / f"{did}.json").write_text(json.dumps({"kind": "pdf", "runs": {}}))
    probs, _st = RN.verify_state(tmp_path)
    assert any("開けた割合" in p for p in probs) and any("0 件" in p for p in probs)


def test_guard_out_refuses_real_places(tmp_path):
    with pytest.raises(SystemExit):
        RN.guard_out(common.ROOT / "world" / "portal")
    with pytest.raises(SystemExit):
        RN.guard_out(tmp_path / "film" / "x")
    assert RN.guard_out(tmp_path / "ok") == (tmp_path / "ok").resolve()


# ---------------------------------------------------------------- 数の形
FRAMES = {"crawl": "X", "countries": {"AA": {"families": {"pdf": {"frame": 100, "pool": 100}, "csv": {"frame": 50, "pool": 50}}}, "BB": {"families": {"pdf": {"frame": 10, "pool": 10}}}}, "wikidata": {}}


def _sub_scores():
    docs = [{"id": f"AA-pdf-{i}", "country": "AA", "family": "pdf", "domain_source": "suffix", "pi": 0.1, "block": i} for i in range(4)]
    docs += [{"id": f"AA-csv-{i}", "country": "AA", "family": "csv", "domain_source": "suffix", "pi": 0.2, "block": i} for i in range(2)]
    docs += [{"id": "BB-pdf-0", "country": "BB", "family": "pdf", "domain_source": "suffix", "pi": 1.0, "block": 9}]
    s = {d["id"]: {"read_ok": True, "read_todoku": True, "read_base": i % 2 == 0, "write_todoku": i % 2 == 1, "write_base": False, "kind": "pdf"} for i, d in enumerate(docs)}
    s["AA-csv-1"] = {"read_ok": True, "read_todoku": False, "read_base": True, "kind": "csv"}  # 書く比べの対象でない
    return docs, s


def test_estimate_counts_no_field_as_not_writable_and_shapes():
    docs, s = _sub_scores()
    res = RN.estimate(docs, s, FRAMES)
    w = res["world"]
    assert w["read_todoku"]["n"] == 7 and w["write_todoku"]["n"] == 7  # 書く欄のない書類も分母に入る
    assert w["write_todoku"]["frame"] == 160
    frag = RN.portal_fragment(res, "t")
    assert set(frag["final"]) == {"version", "read", "write"} and set(frag["final"]["read"]) == {"reach", "new", "base", "lost"}
    assert set(frag["final"]["read"]["reach"]) == {"n", "ci"}
    assert frag["countries"]["AA"]["counts"]["write"].keys() == {"reach", "new"}
    fc = RN.film_counts(res, "provisional")
    assert fc["total"] == sum(fc["countries"].values()) and isinstance(fc["countries_surveyed"], int)


def test_estimate_excludes_unscored_from_denominator():
    docs, s = _sub_scores()
    s["AA-pdf-0"] = {"read_ok": False, "read_todoku": False, "read_base": False, "write_unscored": True}
    res = RN.estimate(docs, s, FRAMES)
    assert res["world"]["read_todoku"]["n"] == 6 and res["world"]["write_todoku"]["n"] == 6


def test_one_score_unscored_is_never_exact():
    d = RN.one_score({"unscored": "time_limit", "tools": {"todoku": {"opened": True, "exact_read": False, "unscored": True}}}, {})
    assert d["read_ok"] is False and d["read_todoku"] is False


def test_portal_shape_checker_detects_mismatch(tmp_path):
    spec = importlib.util.spec_from_file_location("check_shapes_v2", V2 / "check_shapes.py")
    cs = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cs)

    assert cs.same_keys({"a": {"n": 1, "ci": [1, 2]}}, {"a": {"n": 2, "ci": [0, 5]}}) == []
    assert cs.same_keys({"a": {"n": 1}}, {"a": {}}) and cs.same_keys({"n": 1}, {"n": None})


# ---------------------------------------------------------------- 第 2 段の標本
def test_stage2_allocation_only_from_frame_sizes_and_fixed():
    st = [{"country": "US", "family": "pdf", "domain_source": "suffix", "pool": 7000, "frame": 800000, "blocks": 3000},
          {"country": "JP", "family": "pdf", "domain_source": "suffix", "pool": 9000, "frame": 300000, "blocks": 900},
          {"country": "AD", "family": "pdf", "domain_source": "suffix", "pool": 60, "frame": 60, "blocks": 2},
          {"country": "TV", "family": "pdf", "domain_source": "suffix", "pool": 10, "frame": 10, "blocks": 1}]
    a = SS.allocate(st, 100, 2)
    assert {s["country"] for s in a} == {"US", "JP", "AD"}  # 候補が 50 以下の層は測らない
    assert sum(s["n"] for s in a) == 100
    by = {s["country"]: s["n"] for s in a}
    assert by["AD"] == 2 and by["US"] > by["JP"] > 2
    assert SS.allocate(st, 100, 2) == a


def test_stage2_select_excludes_seen_and_is_deterministic():
    cands = [{"digest": f"D{i}", "block": i % 3, "block_stratum": "core"} for i in range(30)]
    excl = {f"D{i}" for i in range(10)}
    base = {"country": "AA", "family": "pdf", "domain_source": "suffix"}
    a, pool = SS.select_stratum(cands, 5, 20261107, excl, {"core": 3}, {"core": 12}, base, "")
    b, _ = SS.select_stratum(list(reversed(cands)), 5, 20261107, excl, {"core": 3}, {"core": 12}, base, "")
    assert pool == 20 and [x["digest"] for x in a] == [x["digest"] for x in b]
    assert not ({x["digest"] for x in a} & excl)
    assert abs(a[0]["pi"] - (3 / 12) * (5 / 20)) < 1e-12
    other, _ = SS.select_stratum(cands, 5, 1, excl, {"core": 3}, {"core": 12}, base, "")
    assert [x["digest"] for x in other] != [x["digest"] for x in a]


def test_stage2_plan_committed_matches_frames():
    """計画のファイルは world/frames.json からの決まった計算：書き換えると分かる。"""
    plan = json.loads(SS.PLAN.read_text(encoding="utf-8"))
    frames = json.loads((SS.W / "frames.json").read_text(encoding="utf-8"))
    use = SS.allocate(SS.strata_of(frames), plan["n_total"], plan["k_min"], plan["pool_min"])
    assert [(s["country"], s["family"], s["domain_source"], s["n"]) for s in use] == [(s["country"], s["family"], s["domain_source"], s["n"]) for s in plan["allocation"]]
    assert plan["seed"] == 20261107 and plan["n_total"] == 900


def test_exclusion_includes_every_sample_row():
    ex = SS.exclusion_digests()
    rows = [json.loads(x) for x in (SS.W / "sample.jsonl").read_text(encoding="utf-8").splitlines() if x.strip()]
    assert all(r["digest"] in ex for r in rows)


FROZEN = Path("/private/tmp/claude-501/-Users-onokatsuki-Documents-GitHub-claude---/fdcc494c-06ce-4f3c-bc44-7cab0977bf2d/scratchpad/final/world")


@pytest.mark.skipif(not (FROZEN / "scores.json").exists(), reason="固定版の記録がない")
def test_estimator_reproduces_published_numbers():
    """数の計算は固定版と同じ：固定版の scores.json から、公開した final と国の数がそのまま出る。"""
    sub = json.loads((FROZEN / "subset.json").read_text(encoding="utf-8"))["docs"]
    scores = json.loads((FROZEN / "scores.json").read_text(encoding="utf-8"))["scores"]
    frames = json.loads((RN.W / "frames.json").read_text(encoding="utf-8"))
    res = RN.estimate(sub, scores, frames, exclude_unscored=False)
    real = json.loads((RN.W / "portal" / "data.json").read_text(encoding="utf-8"))
    frag = RN.portal_fragment(res, real["final"]["version"])
    assert frag["final"] == real["final"]
    for cc, c in frag["countries"].items():
        assert c["counts"] == real["countries"][cc]["counts"], cc
    assert RN.film_counts(res, "final")["world_point"] == 1822350


def test_jp_forms_page_list_fixed_and_refusal_rule():
    spec = importlib.util.spec_from_file_location("jp_forms_v2", V2 / "jp_forms.py")
    jf = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(jf)
    pages = [ln.split("\t") for ln in jf.PAGES.read_text(encoding="utf-8").splitlines() if not ln.startswith("#")]
    assert {p[2] for p in pages} == {"www.pref.fukushima.lg.jp", "www.city.akita.lg.jp", "www.city.fukuoka.lg.jp", "www.city.kashiwa.lg.jp"}
    assert not ({p[2] for p in pages} & jf.seen_hosts([]))
    for why in ("robots", "captcha", "login", "http_403", "http_429", "http_401"):
        assert jf.REFUSED.match(why)
    assert not jf.REFUSED.match("http_404")
    html = b'<a href="a.pdf">x</a><a href="/b.XLSX">y</a><a href="c.html">z</a>'
    assert jf.links_of("https://x.example/d/p.html", html, "text/html") == ["https://x.example/b.XLSX", "https://x.example/d/a.pdf"]


# ---------------------------------------------------------------- 追加の直し（R7〜R9・inline の所・選び直し）
SRC = Path(common.ROOT) / "src"


def test_radio_marks_count_as_state():
    assert RC.checks([{"checked": True}, {"checked": False}], "◉［r1＝A］ ｜ ○［r2］")["ok"] == 2
    assert RC.checks([{"checked": True}, {"checked": False}], "☑ ☐")["ok"] == 2  # 比べる道具の印も同じ
    assert RC.checks([{"checked": True}], "○［r1］")["ok"] == 0  # 状態が違えば落ちる


def test_decoder_keeps_literal_bar_and_blank_places():
    win = "## 本文\n\n- 区分｜種別 の順に書く［P1.c10？］\n- a ｜ b\n- 氏名（［P2.c3］）を書く\n- ab（［　］）cd\n- ［P1.c2？…］続き\n"
    t = RC.clean_todoku(RC.todoku_content(win))
    assert "区分｜種別" in t and "a   b" in t.replace("｜", "X")
    assert "ab（　）cd" in t
    assert "…続き" in t and "P1.c2" not in t and "P2.c3" not in t and "［" not in t
    # 本当に欠けた字は、これでも落ちる
    assert RC.recalls(["区分｜種別"], RC.clean_todoku(RC.todoku_content("## 本文\n\n- 区分 種別\n")))["char"] < 1


def test_merged_row_matches_through_blank_place():
    h = RC.Hay(RC.clean_todoku("- 5: 結合 ｜ ab（［　］）cd\n"))
    assert any("ab( )cd" in ln and "結合" in ln for ln in h.lines)


def test_form_mark_inside_word_does_not_break_bigrams_but_printed_box_stays():
    assert RC.recalls(["YesNo"], RC.clean_todoku(RC.drop_ctrl_marks("Yes☐［Check14］No")))["bigram"] == 1.0
    assert RC.recalls(["YesNo"], RC.clean_todoku("Yes☐［Check14］No"))["bigram"] < 1.0  # 直す前の形
    assert "☐" in RC.drop_ctrl_marks("☐［P3.c1？］ はい") and "☐" in RC.drop_ctrl_marks("☐［p1.t10-20］")
    assert "◉" not in RC.drop_ctrl_marks("◉［r1＝A］")
    assert RC.CHECK_OFF.search("☐［Check14］")  # 状態の数えは元の字のまま


@need_base
def test_inline_places_picked_only_when_no_other_and_all_tools_measured(tmp_path):
    code = """
import docx, openpyxl, sys
d = docx.Document(); d.add_paragraph("氏名（　　　　）を書く"); d.save(sys.argv[1] + "/a.docx")
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S"; ws["A1"] = "氏名（　　　　）"; wb.save(sys.argv[1] + "/b.xlsx")
"""
    subprocess.run([BPY, "-I", "-c", code, str(tmp_path)], check=True, timeout=60)
    tw = [sys.executable, "-I", str(V2 / "todoku_write.py"), str(SRC)]
    for name, tool, kind in (("a.docx", "python-docx", "docx"), ("b.xlsx", "openpyxl", "xlsx")):
        p = str(tmp_path / name)
        pick = json.loads(subprocess.run([*tw, p, "pick"], capture_output=True, text=True, timeout=120).stdout.splitlines()[-1])
        assert pick["id"] and pick["inline"] is True, pick  # 前の定義では「書ける欄がない」
        out = str(tmp_path / ("t_" + name))
        res = json.loads(subprocess.run([*tw, p, "fill", out, pick["id"], "TEST-12345"], capture_output=True, text=True, timeout=120).stdout.splitlines()[-1])
        assert not res.get("error") and res.get("proof_ok") is True and Path(out).exists(), res
        chk = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "check", p, out, kind, json.dumps(pick), "TEST-12345"], capture_output=True, text=True, timeout=120)
        assert json.loads(chk.stdout.splitlines()[-1])["landed"] is True
        o2 = str(tmp_path / ("b_" + name))
        w = json.loads(subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "write", tool, p, o2, kind, json.dumps(pick), "TEST-12345"], capture_output=True, text=True, timeout=120).stdout.splitlines()[-1])
        assert not w.get("error"), w
        chk = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "check", p, o2, kind, json.dumps(pick), "TEST-12345"], capture_output=True, text=True, timeout=120)
        assert json.loads(chk.stdout.splitlines()[-1])["landed"] is True


@need_base
def test_rule_cell_chosen_before_inline_and_dark_cell_skipped(tmp_path):
    code = """
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white
c = canvas.Canvas(sys.argv[1])
c.drawString(72, 760, "Name"); c.setFillColor(black); c.rect(150, 750, 300, 24, stroke=1, fill=1)
c.setFillColor(black); c.drawString(72, 700, "City"); c.setFillColor(white); c.rect(150, 690, 300, 24, stroke=1, fill=1)
c.save()
"""
    p = tmp_path / "k.pdf"
    subprocess.run([BPY, "-I", "-c", code, str(p)], check=True, timeout=60)
    tw = [sys.executable, "-I", str(V2 / "todoku_write.py"), str(SRC), str(p), "pick"]
    pick = json.loads(subprocess.run(tw, capture_output=True, text=True, timeout=120).stdout.splitlines()[-1])
    assert pick["inline"] is False and pick["rect"][1] > 100  # 黒いマス（上）でなく白いマス（下）
    again = json.loads(subprocess.run([*tw, json.dumps([pick["id"]])], capture_output=True, text=True, timeout=120).stdout.splitlines()[-1])
    assert again["id"] != pick["id"] and again["inline"] is True  # 断られたら次の所へ。黒いマスには戻らない


def _load_wc():
    sys.path.insert(0, str(V2))
    for m in ("common", "make_signed", "write_compare"):
        sys.modules.pop(m, None)
    try:
        return importlib.import_module("write_compare")
    finally:
        sys.path.remove(str(V2))
        for m in ("common", "make_signed", "write_compare"):
            sys.modules.pop(m, None)


def test_refused_place_is_repicked_for_every_tool(tmp_path, monkeypatch):
    wc = _load_wc()
    calls = []

    def fake(cmd, **kw):
        c = [str(x) for x in cmd]
        calls.append(c)
        if "pick" in c:
            i = c.index("pick")
            done = json.loads(c[i + 1]) if len(c) > i + 1 else []
            pid = "narrow" if not done else "wide"
            return {"stdout": json.dumps({"id": pid, "kind": "text", "inline": False}), "returncode": 0}
        if "fill" in c:
            pid = c[c.index("fill") + 2]
            if pid == "narrow":
                return {"stdout": json.dumps({"error": "ApiError: 入りきりません"}), "returncode": 0}
            Path(c[c.index("fill") + 1]).write_bytes(b"x")
            return {"stdout": json.dumps({"proof_ok": True}), "returncode": 0}
        if "write" in c:
            Path(c[c.index("write") + 3]).write_bytes(b"x")
            return {"stdout": json.dumps({"tool": c[c.index("write") + 1], "target": c[c.index("write") + 5]}), "returncode": 0}
        if "check" in c:
            return {"stdout": json.dumps({"landed": True, "lost": {}}), "returncode": 0}
        return {"stdout": "{}", "returncode": 0}

    monkeypatch.setattr(wc, "run_watched", fake)
    doc = tmp_path / "a.docx"
    doc.write_bytes(b"x")

    class A:
        todoku_python = sys.executable
        baseline_python = sys.executable
        todoku_src = "s"
        pkr = None

    rec = wc.one(str(doc), A, tmp_path / "o")
    assert rec["pick_refused"] == ["narrow"] and rec["target"]["id"] == "wide"
    wr = [c for c in calls if "write" in c]
    assert wr and all(json.loads(c[c.index("write") + 5])["id"] == "wide" for c in wr)  # 比べる道具も同じ所
    assert rec["runs"]["todoku:ascii"]["write_ok"] and rec["runs"]["python-docx:ascii"]["write_ok"]


def test_unverified_file_is_inspected_but_never_write_ok(tmp_path, monkeypatch):
    wc = _load_wc()

    def fake(cmd, **kw):
        c = [str(x) for x in cmd]
        if "pick" in c:
            return {"stdout": json.dumps({"id": "A", "kind": "text"}), "returncode": 0}
        if "fill" in c:
            o = Path(c[c.index("fill") + 1])
            wc.unverified_file(o).write_bytes(b"x")
            return {"stdout": json.dumps({"proof_ok": False}), "returncode": 0}
        if "check" in c:
            return {"stdout": json.dumps({"landed": True, "lost": {}}), "returncode": 0}
        return {"stdout": "{}", "returncode": 0}

    monkeypatch.setattr(wc, "run_watched", fake)
    monkeypatch.setitem(wc.TOOLS, "docx", ())
    doc = tmp_path / "a.docx"
    doc.write_bytes(b"x")

    class A:
        todoku_python = sys.executable
        baseline_python = sys.executable
        todoku_src = "s"
        pkr = None

    r = wc.one(str(doc), A, tmp_path / "o")["runs"]["todoku:ascii"]
    assert r["unverified"] is True and r["check"]["landed"] is True and r["write_ok"] is False


def test_previous_definition_kept_beside():
    run = {"todoku:ascii": {"write_ok": True}, "openpyxl:ascii": {"write_ok": False}}
    d = RN.one_score({"tools": {"todoku": {"opened": True}}}, {"target": {"inline": True}, "runs": run})
    assert d["write_todoku"] is True and d["write_todoku_prev"] is False
    d = RN.one_score({"tools": {"todoku": {"opened": True}}}, {"target": {"inline": False}, "runs": run})
    assert d["write_todoku_prev"] is True


@need_base
def test_color_not_normal_counts_beside(tmp_path):
    code = """
import docx, sys
from docx.shared import RGBColor
d = docx.Document()
r = d.add_paragraph().add_run("ふつうの色の長い本文の字の並び"); r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
r = d.add_paragraph().add_run("赤い字"); r.font.color.rgb = RGBColor(0xFF, 0, 0)
d.save(sys.argv[1])
"""
    x, o = tmp_path / "c.docx", tmp_path / "p.json"
    subprocess.run([BPY, "-I", "-c", code, str(x)], check=True, timeout=60)
    subprocess.run([BPY, "-I", str(V2 / "probes.py"), str(x), "docx", str(o)], check=True, timeout=60)
    pr = json.loads(o.read_text())
    assert pr["color_normal"] == "333333"
    cat = RC.score_docx({"probes": pr}, RC.Hay("赤い字 色:FF0000"), "")
    assert cat["color"]["n"] == 2 and cat["color_not_normal"] == {"n": 1, "ok": 1}


@need_base
@need_lo
def test_doc_check_finds_value_in_table_and_rejects_absent(tmp_path):
    """.doc の確かめ（W7）。値がないときは落ちる。"""
    x = tmp_path / "t.docx"
    subprocess.run([BPY, "-I", "-c", "import docx,sys; d=docx.Document(); t=d.add_table(rows=1,cols=1); t.cell(0,0).text='TEST-12345'; d.save(sys.argv[1])", str(x)], check=True, timeout=60)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{tmp_path}/prof", "--headless", "--norestore", "--convert-to", "doc:MS Word 97", "--outdir", str(tmp_path), str(x)], capture_output=True, timeout=180)
    doc = str(tmp_path / "t.doc")
    tg = json.dumps({"id": "T1.R1.C1"})
    for val, want in (("TEST-12345", True), ("ABSENT-999", False)):
        r = subprocess.run([BPY, "-I", str(V2 / "baseline_write.py"), "check", doc, doc, "doc", tg, val], capture_output=True, text=True, timeout=400)
        assert json.loads(r.stdout.splitlines()[-1])["landed"] is want
