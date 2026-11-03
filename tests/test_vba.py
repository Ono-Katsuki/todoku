"""マクロ（VBA）のソースの取り出し：Microsoft の仕様書 MS-OVBA の例（3.2 圧縮の例、3.1.6 PROJECT ストリームの例）と同じ結果になる。"""

from __future__ import annotations

import glob
import zipfile

import pytest

from todoku.vba import decompress, decrypt, read_project


@pytest.mark.parametrize("comp, plain", [
    ("01 19 B0 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 73 74 75 76 2E", b"abcdefghijklmnopqrstuv."),
    ("01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 61 6B 6C 00 20 6D 6E 6F "
     "70 06 71 02 70 04 00 72 73 74 75 76 10 77 78 79 7A 00 2C",
     b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"),
    ("01 03 B0 02 61 45 00", b"a" * 73),  # 16×4＋9 バイト
])
def test_decompress_spec_examples(comp, plain):
    assert decompress(bytes.fromhex(comp)) == plain


def test_decrypt_spec_examples():
    assert decrypt(bytes.fromhex("0705D8E3D8EDDBF1DBF1DBF1DBF1")) == b"\x00\x00\x00\x00"  # CMG：保護なし
    assert decrypt(bytes.fromhex("0E0CD1ECDFF4E7F5E7F5E7")) == b"\x00"  # DPB：パスワードなし
    assert decrypt(bytes.fromhex("1517CAF1D6F9D7F9D706")) == b"\xff"  # GC：見える


@pytest.mark.corpus
def test_corpus_macro_sources():
    files = [f for f in sorted(glob.glob("corpus/**/*.xlsm", recursive=True))]
    if not files:
        pytest.skip("corpus なし")
    for f in files:
        p = read_project(zipfile.ZipFile(f).read("xl/vbaProject.bin"))
        assert p.code_page == 932 and p.modules
        assert any(m.procedures for m in p.modules), f  # どのファイルにも、手続きのあるモジュールがある
        for m in p.modules:
            assert m.source.startswith("Attribute VB_Name"), (f, m.name)  # ソースの頭は、いつもモジュールの名前


def _project():
    from todoku.vba import Module, Project

    src_sheet = ('Attribute VB_Name = "Sheet1"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                 '    Range("B1").Value = Target.Value * 2\nEnd Sub\n')
    src_mod = 'Attribute VB_Name = "Module1"\nSub 印刷()\nAttribute 印刷.VB_ProcData.VB_Invoke_Func = " \\n14"\n    ActiveSheet.PrintOut\nEnd Sub\n'
    return Project(932, "VBAProject", [Module("Sheet1", "Sheet1", 0, "document", src_sheet, ["Worksheet_Change"]),
                                         Module("Module1", "Module1", 0, "procedural", src_mod, ["印刷"])])


def test_macro_section_lists_events_buttons_and_sources():
    from todoku.macros import change_event_sheets, macro_info, markdown

    info = macro_info(_project(), True, {"Sheet1": "入力"}, [{"sheet": "入力", "control": "Button 1", "text": "印刷",
                                                               "macro": "[0]!Module1.印刷"}])
    assert [e["when"] for e in info["events"]] == ["このシートのセルの値を変えたとき"] and info["events"][0]["on_change"]
    assert change_event_sheets(info) == {"入力": ["Sheet1.Worksheet_Change"]}
    md = "\n".join(markdown(info))
    assert "値を変えると動く" in md and "Module1.印刷" in md and 'Range("B1").Value = Target.Value * 2' in md
    assert "Attribute" not in md  # 編集の画面に出ない設定の行は、窓にも出さない


def test_fill_warns_about_change_event_macros():
    from test_example import HEAD, VALS, make_xlsx
    from todoku.xlsx.edit import fill
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(make_xlsx({"入力": (HEAD, VALS)}))
    wb.has_macros = True
    wb._vba = _project()
    wb.sheets[0].code_name = "Sheet1"
    # マクロ（Range("B1").Value = Target.Value * 2）は動かす。文字「A」の2倍は、Excel でも型が合わずに止まる
    res = fill(wb, {"入力!A2": "A"})
    assert res.macro["reproduced"] and any("Excel でも止まります" in w and "'13'" in w for w in res.warnings), res.warnings
    assert [p.kind for p in res.patches] == ["input"]  # 止まったので、マクロはほかのセルを書いていない
    res = fill(wb, {"入力!A2": 21})
    assert {p.address: p.kind for p in res.patches} == {"A2": "input", "B1": "macro"}
