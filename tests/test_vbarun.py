"""マクロ（VBA）を動かす仕組み：言語の規則（VBA 言語リファレンスのとおり）と、マクロから見える Excel。"""

from __future__ import annotations

import pytest

from todoku.vba import Module, Project
from todoku.vbarun import NeedsInput, Unsupported, change_events, run_macro
from todoku.xlsx.workbook import Workbook

from test_example import make_xlsx


def book(source: str, cells: dict | None = None, sheet_module: str = "") -> Workbook:
    """セルとマクロを持つブック。source は標準モジュール Module1、sheet_module はシート「入力」のモジュール Sheet1。"""
    from todoku.vba import _PROC

    wb = Workbook(make_xlsx({"入力": (cells or {}, [])}))

    def procs(src):
        return [m.group(2) for m in _PROC.finditer(src)]

    mods = [Module("Module1", "Module1", 0, "procedural", 'Attribute VB_Name = "Module1"\n' + source, procs(source))]
    if sheet_module:
        mods.append(Module("Sheet1", "Sheet1", 0, "document", 'Attribute VB_Name = "Sheet1"\n' + sheet_module,
                           procs(sheet_module)))
    wb.has_macros = True
    wb._vba = Project(932, "VBAProject", mods)
    wb.sheets[0].code_name = "Sheet1"
    return wb


def out(wb, macro="Test", **kw):
    r = run_macro(wb, macro, **kw)
    return {f"{s}!{chr(64 + c)}{r_}": v for (s, r_, c), v in r.changed.items()}, r


def test_language_rules():
    src = '''
Function Twice(ByRef n As Long) As Long
    n = n * 2
    Twice = n + 1
End Function

Sub Test()
    Dim i As Long, s As String, total As Double, a(1 To 3) As Long, arr() As String
    For i = 1 To 3
        a(i) = i * 10
        total = total + a(i)
    Next i
    Range("A1").Value = total                     ' 60
    s = ""
    Dim k As Variant
    For Each k In Array("x", "y", "z")
        s = s & UCase(k)
    Next
    Range("A2").Value = s                         ' XYZ
    Select Case total
        Case Is < 10: Range("A3").Value = "small"
        Case 50 To 70: Range("A3").Value = "mid"
        Case Else: Range("A3").Value = "big"
    End Select
    ReDim arr(0 To 1)
    arr(0) = "a": arr(1) = "b"
    ReDim Preserve arr(0 To 2)
    arr(2) = "c"
    Range("A4").Value = Join(arr, "-")            ' a-b-c
    Dim n As Long: n = 5
    Range("A5").Value = Twice(n)                  ' 11（参照渡しで n は 10）
    Range("A6").Value = n                         ' 10
    Range("A7").Value = 7 \\ 2 & "," & 7 Mod 3 & "," & -7 \\ 2   ' 3,1,-3
    Range("A8").Value = Round(2.5) & "," & Round(3.5) & "," & Int(-2.5) & "," & Fix(-2.5)  ' 2,4,-3,-2
    Range("A9").Value = Mid("申請書類", 2, 2) & InStr("abcabc", "c") & Len("とどく")  ' 請書33
    Range("A10").Value = Format(1234.5, "#,##0.0") & "|" & Format(DateSerial(2026, 4, 1), "yyyy/mm/dd ggge年")
    With Range("B1")
        .Value = "with"
    End With
    Dim c As New Collection
    c.Add "one", "k1": c.Add "two"
    Range("B2").Value = c("k1") & c(2) & c.Count
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("x") = 1: d.Add "y", 2
    Range("B3").Value = d.Count & d("y") & d.Exists("z")
End Sub
'''
    got, _ = out(book(src))
    assert got["入力!A1"] == 60 and got["入力!A2"] == "XYZ" and got["入力!A3"] == "mid" and got["入力!A4"] == "a-b-c"
    assert got["入力!A5"] == 11 and got["入力!A6"] == 10
    assert got["入力!A7"] == "3,1,-3" and got["入力!A8"] == "2,4,-3,-2"
    assert got["入力!A9"] == "請書33" and got["入力!A10"] == "1,234.5|2026/04/01 令和8年"
    assert got["入力!B1"] == "with" and got["入力!B2"] == "onetwo2" and got["入力!B3"] == "22False"


def test_error_handling():
    src = '''
Sub Test()
    On Error Resume Next
    Dim x As Double
    x = 1 / 0
    Range("A1").Value = Err.Number                ' 11
    Err.Clear
    On Error GoTo Handler
    x = CInt("abc")
    Range("A2").Value = "not reached"
    Exit Sub
Handler:
    Range("A3").Value = "caught " & Err.Number    ' caught 13
    Resume Done
Done:
    Range("A4").Value = "done"
End Sub
'''
    got, _ = out(book(src))
    assert got["入力!A1"] == 11 and "入力!A2" not in got and got["入力!A3"] == "caught 13" and got["入力!A4"] == "done"


def test_excel_cells_end_and_recalc():
    cells = {"A1": "名前", "A2": "a", "A3": "b", "A4": "c", "B1": 5}
    src = '''
Sub Test()
    Dim last As Long
    last = Cells(Rows.Count, 1).End(xlUp).Row     ' 4
    Cells(last + 1, 1).Value = "d"
    Range("C1").Value = last
    Range("C2").Value = WorksheetFunction.CountA(Range("A1:A10"))  ' 5
    Range("C3").Value = Range("A2").Offset(1, 0).Value & Range("A1").Resize(2, 1).Cells(2, 1).Value  ' ba
    Dim cell As Range, s As String
    For Each cell In Range("A2:A4")
        s = s & cell.Value
    Next
    Range("C4").Value = s                         ' abc
    Range("A3").ClearContents
    Range("C5").Value = IsEmpty(Range("A3").Value)
End Sub
'''
    got, _ = out(book(src, cells))
    assert got["入力!A5"] == "d" and got["入力!C1"] == 4 and got["入力!C2"] == 5
    assert got["入力!C3"] == "ba" and got["入力!C4"] == "abc" and got["入力!A3"] is None and got["入力!C5"] is True


def test_change_event_runs_like_a_person_typing():
    sheet = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("A1:A10")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Target.Offset(0, 1).Value = UCase(Target.Value) & "!"
    Application.EnableEvents = True
End Sub
'''
    wb = book("Sub Test()\nEnd Sub\n", {}, sheet)
    r = change_events(wb, [("入力", 2, 1, "abc"), ("入力", 2, 3, "zzz")])
    assert r.changed == {("入力", 2, 2): "ABC!"}  # A2 に入れたときだけ、B2 をマクロが書く（C2 は範囲の外）


def test_needs_answers_instead_of_guessing():
    src = '''
Sub Test()
    If MsgBox("消しますか？", vbYesNo) = vbYes Then Range("A1").Value = "消した" Else Range("A1").Value = "残した"
    MsgBox "終わりました"
End Sub
'''
    with pytest.raises(NeedsInput):
        run_macro(book(src), "Test")  # はい・いいえは、勝手に決めない
    got, r = out(book(src), answers=["いいえ"])
    assert got["入力!A1"] == "残した" and r.messages == ["消しますか？", "終わりました"]


def test_file_output_and_unsupported():
    src = '''
Sub Test()
    Open "out.csv" For Output As #1
    Print #1, "名前,数"
    Write #1, "とどく", 3
    Close #1
End Sub
Sub Fmt()
    Range("A1").Interior.Color = vbRed
End Sub
'''
    _got, r = out(book(src))
    assert r.files["out.csv"].decode("cp932") == '名前,数\r\n"とどく",3\r\n'
    with pytest.raises(Unsupported):
        run_macro(book(src), "Fmt")  # 書式の変更は再現しない（はっきり断る）


def test_conditional_compilation():
    """#Const・#If・#ElseIf・#Else・#End If（MS-VBAL 3.4）。手続きの外の宣言も、手続きの中の文も選ぶ。"""
    src = '''
#Const 版 = 2
#If Mac Then
Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal c As String) As LongPtr
#ElseIf VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

#If Mac Then
Private Type 結果
    x As LongPtr
End Type
#Else
Private Type 結果
    x As Long
End Type
#End If

Sub Test()
#If 版 >= 2 And Not Mac Then
    Range("A1").Value = "新しい版"
#ElseIf 版 = 1 Then
    Range("A1").Value = "古い版"
#Else
    この行は読まない ((
#End If
#If 決めていない Then
    Range("A2").Value = "決めていない名前は Empty"
#Else
    Range("A2").Value = "偽"
#End If
#If Win64 Then
    Range("A3").Value = 64
#Else
    Range("A3").Value = 32
#End If
    Range("A4").Value = 1 + _
        2
End Sub
'''
    got, _r = out(book(src))
    assert got == {"入力!A1": "新しい版", "入力!A2": "偽", "入力!A3": 64, "入力!A4": 3}


def test_conditional_compilation_rules():
    from todoku.vbarun.parse import VbaSyntaxError, parse_module, preprocess, project_constants

    # 行の番号は変えない（使わない行と指示の行は空の行）
    src = "#If Mac Then\nA = 1\n#Else\nB = 2\n#End If\nC = 3"
    assert preprocess(src).split("\n") == ["", "", "", "B = 2", "", "C = 3"]
    # 続きの行（ _）も、ひとまとまりの行として扱う
    assert preprocess("#If Mac Then\nx = 1 + _\n  2\n#End If\ny = 1").split("\n") == ["", "", "", "", "y = 1"]
    # 入れ子：外側が使われないときは、内側の #Else も使わない
    nested = preprocess("#If Mac Then\n#If VBA7 Then\na\n#Else\nb\n#End If\n#End If\nc").split("\n")
    assert nested == ["", "", "", "", "", "", "", "c"]
    # 使われない側の #Const は効かない
    assert "z" in preprocess("#If Mac Then\n#Const K = 1\n#End If\n#If K Then\ny\n#Else\nz\n#End If")
    # プロジェクトの引数（PROJECTCONSTANTS）
    assert project_constants("DEBUG_MODE = 1 : 版 = -2") == {"debug_mode": 1, "版": -2}
    assert "on" in preprocess("#If DEBUG_MODE Then\non\n#End If", {"debug_mode": 1})
    with pytest.raises(VbaSyntaxError, match="閉じていません"):
        parse_module("M", "#If Mac Then\nSub A()\nEnd Sub")
    with pytest.raises(VbaSyntaxError, match="#If」のない"):
        parse_module("M", "#End If")


def test_declarations_and_debug_print():
    """WithEvents の変数・配列を返す手続き・Debug.Print の並び（「;」「,」・Spc・Tab）。"""
    from todoku.vbarun.parse import FileStmt, parse_module

    m = parse_module("C", '''Option Explicit
Public WithEvents TextBox As MSForms.TextBox
Attribute TextBox.VB_VarHelpID = -1
Private WithEvents btn As MSForms.CommandButton, n As Long
Implements IHandler
Public Event Changed(ByVal v As Long)

Function check() As String()
    Dim a() As String
    check = a
End Function

Public Property Get Years() As Long()
End Property
''')
    assert [v.name for d in m.decls for v in d.vars] == ["textbox", "btn", "n"]
    assert {"check", "get:years"} <= set(m.procs)
    p = parse_module("M", '''Sub A()
    Debug.Print "緯度: " & x; "度" & y, z; Spc(2); Tab(10); Tab; "終"
    Debug.Print
End Sub''')
    st = p.procs["a"].body[0]
    assert isinstance(st, FileStmt) and st.kind == "debugprint"
    kinds = [it if isinstance(it, str) else (it[0] if isinstance(it, tuple) else "式") for it in st.args[1]]
    assert kinds == ["式", ";", "式", ",", "式", ";", "spc", ";", "tab", ";", "tab", ";", "式"]

    src = '''
Function Bump() As Long
    Range("A1").Value = Range("A1").Value + 1
    Bump = 1
End Function
Sub Test()
    Debug.Print "回数"; Bump(); Spc(Bump())
    Open "out.txt" For Output As #1
    Print #1, "a"; Spc(2); "b"; Tab(8); "c"
    Print #1, "abcdefghij"; Tab(3); "d"
    Close #1
End Sub
'''
    got, r = out(book(src, {"A1": 0}))
    assert got["入力!A1"] == 2  # Debug.Print の式も計算する（呼んだ関数が動く）
    assert r.files["out.txt"].decode("cp932") == "a  b   c\r\nabcdefghij\r\n  d\r\n"


def test_fill_runs_change_macros_and_proves_their_cells():
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove

    sheet = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$A$2" Then Exit Sub
    Application.EnableEvents = False
    Range("B2").Value = Target.Value * 2
    Application.EnableEvents = True
End Sub
'''
    wb = book("Sub Test()\nEnd Sub\n", {"A1": "数", "B1": "2倍"}, sheet)
    blob = wb.zip.original
    res = fill(wb, {"入力!A2": 21})
    kinds = {p.address: p.kind for p in res.patches}
    assert kinds == {"A2": "input", "B2": "macro"} and res.macro["reproduced"], (kinds, res.macro, res.warnings)
    after = Workbook(res.blob)
    assert after.sheet("入力").cells[(2, 2)].value == 42
    pr = prove(blob, res.blob, res.allowed_inputs, set(), {(p.sheet, p.address) for p in res.patches if p.kind == "macro"})
    assert pr.ok, pr.unexpected


def test_api_run_macro_writes_proof_and_asks_before_guessing(tmp_path):
    from todoku import api

    src = '''
Sub 集計()
    If MsgBox("集計しますか？", vbYesNo) = vbNo Then Exit Sub
    Range("B1").Value = WorksheetFunction.Sum(Range("A1:A3"))
End Sub
'''
    wb = book(src, {"A1": 1, "A2": 2, "A3": 3})
    path = tmp_path / "m.xlsm"
    # ブックの中身（マクロの部品つき）は、組み立てたものを使う：vba を差し込んだブックを、api が開けるように保存する
    path.write_bytes(wb.zip.original)
    import todoku.xlsx.workbook as W

    orig = W.Workbook.vba
    W.Workbook.vba = property(lambda self: wb._vba)  # type: ignore[assignment]
    try:
        W.Workbook.has_macros = True  # type: ignore[attr-defined]
        r = api.run_macro(str(path), "Module1.集計", str(tmp_path / "out.xlsm"))
        assert r["done"] is False and r["needs"] == "msgbox" and r["prompt"] == "集計しますか？"
        r = api.run_macro(str(path), "Module1.集計", str(tmp_path / "out.xlsm"), answers=["はい"])
        assert r["done"] and r["proof_ok"] and r["changed"] == {"入力!B1": 6}, r
    finally:
        W.Workbook.vba = orig  # type: ignore[assignment]
        del W.Workbook.has_macros
