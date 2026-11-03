"""日本語以外の書類でも、書き込む所を同じ見つけ方で示す（字の意味ではなく、字の形と並びだけで決める）。

- 字の集まり（chars.py）：Unicode の資料（一般カテゴリ・BidiBrackets.txt・PropList.txt）から作ったもの
- 文字の中の所：2つ以上の空白の並び・下線の字・点線・ダッシュ・かっこの中の空白・チェックの箱の字
- Excel の見出し：いちばん近い、印刷された字をそのまま（右から左へ表示するシートでは、画面の右）
"""

from __future__ import annotations

import io
import unicodedata

import pytest

from todoku import chars
from todoku.docx.document import DocxDocument
from todoku.docx.fill import fill as docx_fill
from todoku.docx.window import build as docx_build
from todoku.xlsx import window
from todoku.xlsx.edit import fill as xlsx_fill
from todoku.xlsx.workbook import Workbook

from test_docx import make_docx, p

openpyxl = pytest.importorskip("openpyxl")


# ====================================================================== 字の集まり
def test_char_tables_cite_unicode():
    # 字の表は、動かす Python の unicodedata と同じ版でなければならない（pyproject の requires-python で 3.11 に固定）。
    # Python の版を上げるときは、その版で scripts/make_chars.py を動かして chars.py の表を作り直す
    assert chars.UCD_VERSION == unicodedata.unidata_version,         f"chars.py は Unicode {chars.UCD_VERSION}、この Python は {unicodedata.unidata_version}：scripts/make_chars.py で作り直す"
    # かっこの組（BidiBrackets.txt）：ASCII・全角・CJK のかっこ
    for o, c in ("()", "[]", "{}", "（）", "［］", "【】", "「」", "〔〕", "《》"):
        assert chars.closing(o) == c
    assert chars.closing("<") is None  # 不等号は Bidi_Paired_Bracket ではない
    # コロン：Po で名前に COLON。؛（ARABIC SEMICOLON）と ։（ARMENIAN FULL STOP）はコロンではない
    assert all(chars.is_label_colon(ch) for ch in ":：﹕︓፥᠄܃")
    assert not chars.is_label_colon("؛") and not chars.is_label_colon("։") and not chars.is_label_colon(";")
    assert chars.is_sentence_terminal("։") and chars.is_sentence_terminal("।") and chars.is_sentence_terminal("。")
    assert chars.is_terminal_punctuation("、") and chars.is_terminal_punctuation("؛") and not chars.is_terminal_punctuation("_")
    # 下線の字（Pc）・空白（Zs）
    assert all(chars.is_connector(ch) for ch in "_‿⁀⁔︳︴﹍﹎﹏＿")
    assert all(chars.is_space(ch) for ch in "   　") and not chars.is_space("\t")
    # チェックの箱の字
    assert set("☐□▢◻◽❏❐❑❒⬜") == set(chars.EMPTY_BOXES) and set("☑☒") == set(chars.CHECKED_BOXES)


def test_norm_label_keeps_latin_word_spaces():
    assert window.norm_label("I am a b c") == "I am a b c"
    assert window.norm_label("[ ] Yes [ ] No") == "[ ] Yes [ ] No"
    assert window.norm_label("事 業 所 番 号") == "事業所番号"  # 全角の字の字間を空けた見出しは、これまでどおり詰める


# ====================================================================== Excel
def _xlsx(build) -> bytes:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _fields(blob: bytes) -> tuple[dict, dict]:
    doc = window.build(Workbook(blob, "t.xlsx"))
    return doc, {f["id"]: f for sh in doc["sheets"] for f in sh["fields"]}


TEXTS = {
    "A1": "中華民國　　年　　月　　日", "A2": "　　년　　월　　일", "A3": "　　省　　市　　县", "A4": "Name ......",
    "A5": "Nom : …………", "A6": "नाम ......", "A7": "[ ] Yes [ ] No", "A8": "☐ Male ☐ Female", "A9": "( ) Male ( ) Female",
}


def test_excel_places_in_other_languages():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        for a, t in TEXTS.items():
            ws[a] = t

    blob = _xlsx(build)
    doc, f = _fields(blob)
    assert f["S!A1"]["blanks"] == 3 and f["S!A1"]["agent"] and [s["printed"] for s in f["S!A1"]["spans"]] == ["　　"] * 3
    assert f["S!A1"]["spans"][0]["before"] == "中華民國" and f["S!A1"]["spans"][0]["width"] == 2.0
    assert f["S!A2"]["blanks"] == 3 and f["S!A3"]["blanks"] == 3
    for a in ("A4", "A5", "A6"):  # 点線：規則では決めない所（「？」）
        assert f[f"S!{a}"]["blanks"] == 1 and f[f"S!{a}"]["agent"], f[f"S!{a}"]
    for a in ("A7", "A8", "A9"):  # かっこの中の空白・チェックの箱の字：規則の所
        assert f[f"S!{a}"]["blanks"] == 2 and not f[f"S!{a}"].get("agent"), f[f"S!{a}"]
    reading = {x.get("row"): x["text"] for x in doc["sheets"][0]["reading"]}
    assert reading[1] == "〔A1〕中華民國［？　　］年［？　　］月［？　　］日"  # 「？」の所の印に、元の空白をそのまま入れる
    assert reading[7] == "〔A7〕[［　］] Yes [［　］] No"
    assert reading[8] == "〔A8〕☐［　］ Male ☐［　］ Female"  # 箱の字は消さずに示す
    res = xlsx_fill(Workbook(blob, "t.xlsx"), {"S!A1": ["115", "11", "3"], "S!A8": ["☑", None], "S!A4": [" Lee"]})
    sh = Workbook(res.blob, "t.xlsx").sheet("S")
    assert sh.cells[(1, 1)].value == "中華民國115年11月3日"
    assert sh.cells[(8, 1)].value == "☑ Male ☐ Female" and sh.cells[(4, 1)].value == "Name  Lee"


def test_excel_label_is_nearest_printed_text_and_rtl_sheet():
    def build(wb):
        ws = wb.active
        ws.title = "ar"
        ws.sheet_view.rightToLeft = True
        ws["A1"], ws["A2"], ws["C2"] = "الاسم", "المبلغ", "دينار"
        ws.protection.sheet = True
        for a in ("B1", "B2"):
            ws[a].protection = openpyxl.styles.Protection(locked=False)

    doc, f = _fields(_xlsx(build))
    sh = doc["sheets"][0]
    assert sh["right_to_left"] is True
    # 列 A は画面の右端：見出しは列の番号の小さい側（画面の右）の、いちばん近い字をそのまま
    assert f["ar!B1"]["label"] == "الاسم" and f["ar!B1"]["label_method"] == "右の見出し"
    assert f["ar!B2"]["label"] == "المبلغ" and f["ar!B2"]["after"] == "دينار"  # すぐ後ろの字はそのまま（単位かどうかは決めない）
    md = window.to_markdown(doc)
    assert "右から左へ表示するシート" in md


def test_excel_unit_words_are_not_dropped_from_labels():
    # 「円」のような字も、いちばん近い字なら見出しになる（語の一覧で除かない）。字にはさまれた欄の並びは、並びのまま示す
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"], ws["C1"] = "金額", "円"
        ws["A2"], ws["C2"], ws["E2"], ws["G2"] = "令和", "年", "月", "日"
        ws.protection.sheet = True
        for a in ("B1", "B2", "D2", "F2"):
            ws[a].protection = openpyxl.styles.Protection(locked=False)

    _doc, f = _fields(_xlsx(build))
    assert f["S!B1"]["label"] == "金額" and f["S!B1"]["after"] == "円"
    assert f["S!D2"]["label"] == "令和［　］年［▶］月［　］日"


# ====================================================================== Word
def test_word_places_in_other_languages():
    body = "".join(p(t) for t in TEXTS.values())
    blob = make_docx(body)
    win = docx_build(DocxDocument(blob, "w.docx"))
    by = {f["id"]: f for f in win["fields"]}
    n = {k: [f for f in win["fields"] if f["id"].startswith(f"P{k}.") and not f["id"][len(f"P{k}."):].startswith("a")]
         for k in range(1, 10)}
    assert len(n[1]) == 3 and all(f.get("agent") for f in n[1])  # 中華民國　　年　　月　　日
    assert len(n[2]) == 3 and len(n[3]) == 3  # 　　년　　월　　일 ・ 　　省　　市　　县
    for k in (4, 5, 6):
        assert len(n[k]) == 1 and n[k][0].get("agent") and "点線" in n[k][0]["why"]
    assert len(n[7]) == 2 and not any(f.get("agent") for f in n[7])  # [ ] Yes [ ] No
    assert [by["P8.1"]["kind"], by["P8.1"]["label"], by["P8.2"]["label"]] == ["checkbox", "Male", "Female"]
    assert len(n[9]) == 2  # ( ) Male ( ) Female
    res = docx_fill(DocxDocument(blob, "w.docx"), {"P1.1": "115", "P1.2": "11", "P1.3": "3", "P8.1": True})
    lines = [ln for st in docx_build(DocxDocument(res.blob, "w.docx"))["stories"] for ln in st["reading"]]
    assert any(ln.startswith("中華民國［P1.1＝115］年［P1.2＝11］月［P1.3＝3］日") for ln in lines), lines
    assert any(ln.startswith("☑［P8.1］ Male") for ln in lines), lines


def test_word_long_line_of_blanks_is_linear():
    # 改行のない長い段落（.txt の 1 行など）：所の数に比例する時間で作る（所ごとに段落をたどり直さない）
    import time

    def run(n: int) -> float:
        body = p("Nom : ………… ( ) ☐ x ○○ " * n)
        doc = DocxDocument(make_docx(body), "l.docx")
        t = time.perf_counter()
        win = docx_build(doc)
        assert len(win["fields"]) >= 3 * n
        return time.perf_counter() - t

    small, big = run(500), run(4000)
    assert big < 8 * small * 3, (small, big)  # 8 倍の長さで、時間は 8 倍の数倍まで（2 乗なら 64 倍）
