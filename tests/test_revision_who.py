"""変更履歴の「だれが・いつ」と、まだ承諾も却下もされていない挿入の印（.docx・.doc・RTF・OpenDocument で同じ決まり）。

削除された字は本文に入れず、段落の後ろに〔変更履歴で削除された字（…）: 字〕。挿入された字は本文の字として読み、段落の字が
みな同じ挿入なら段落の前に〔変更履歴で挿入された段落（…）〕、一部だけなら段落の後ろに〔うち変更履歴で挿入された字（…）：「字」〕。
変更した人・日時は、ファイルに書いてある値のまま（書いていなければ出さない）。
"""

from __future__ import annotations

import os
import re
import struct

import pytest
from test_openformats import _package, make_rtf, sj
from test_revisions_and_looks import DOC_REV, _large, _lines, _plain, _range_lines, _reasons, _word6_rev
from test_word_findings import docx, p, r, write

from todoku import api, lazy
from todoku.docx.revfacts import dttm

B = "\\"
NEW = "まだ承諾も却下もされていない"
A1 = 'w:author="田中 一郎" w:date="2026-01-05T10:30:00Z"'
A2 = 'w:author="Sato" w:date="2026-02-01T09:00:00+09:00"'
W1 = "変更した人：田中 一郎・日時：2026-01-05T10:30:00Z"
W2 = "変更した人：Sato・日時：2026-02-01T09:00:00+09:00"


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()


def ins(attrs: str, *runs: str, tag: str = "ins") -> str:
    return f'<w:{tag} w:id="1" {attrs}>{"".join(runs)}</w:{tag}>'


def gone(attrs: str, text: str, tag: str = "del") -> str:
    return f'<w:{tag} w:id="2" {attrs}><w:r><w:delText xml:space="preserve">{text}</w:delText></w:r></w:{tag}>'


def _detail(path) -> list[str]:
    return [x["detail"] for x in api.grade(str(path))["reasons"] if x["title"] == "変更履歴"]


# ====================================================================== 日時（Word の DTTM）
def test_dttm_is_read_as_stored_and_empty_when_absent():
    def make(mi, hr, day, mon, yr, wd=0):
        return mi | hr << 6 | day << 11 | mon << 16 | (yr - 1900) << 20 | wd << 29

    assert dttm(make(30, 9, 12, 11, 2026, 4)) == "2026-11-12 09:30"
    assert dttm(0x84610800) == "1970-01-01 00:00"  # tracked_changes.doc の値
    assert dttm(0) == "" and dttm(make(0, 0, 0, 5, 2026)) == "" and dttm(make(0, 0, 5, 0, 2026)) == ""  # 日・月が 0：日時なし
    assert dttm(make(60, 0, 1, 1, 2026)) == "" and dttm(make(0, 24, 1, 1, 2026)) == "" and dttm(make(0, 0, 1, 13, 2026)) == ""


# ====================================================================== .docx
def test_docx_inserted_text_is_marked_with_who_and_when(tmp_path):
    body = (p(r("提出先："), ins(A1, r("市役所")), gone(A1, "町役場"), r("の窓口。"))
            + p(ins(A1, r("まるごと足した"), r("段落。")))
            + p(r("ふつうの段落。")))
    path = write(tmp_path, docx(body))
    lines = _lines(path)
    assert lines == [
        f"提出先：市役所の窓口。〔うち変更履歴で挿入された字（{NEW}。{W1}）：「市役所」〕〔変更履歴で削除された字（いまの本文ではない。{W1}）: 町役場〕",
        f"〔変更履歴で挿入された段落（{NEW}。{W1}）〕まるごと足した段落。",
        "ふつうの段落。",
    ]
    win = api.window(path)[1]
    assert win["tracked_insertions"] == ["段落 1「市役所」", "段落 2「まるごと足した段落。」"]
    assert win["tracked_deletions"] == ["段落 1「町役場」"]
    d = _detail(path)
    assert any("変更履歴で挿入された字" in x and "2 段落" in x for x in d) and any("変更履歴で削除された字が 1 段落" in x for x in d)
    assert api.audit(path)["ok"]


def test_docx_changes_by_different_people_are_told_apart(tmp_path):
    body = (p(r("甲"), ins(A1, r("乙")), r("丙"), ins(A2, r("丁")), gone(A1, "戊"), gone(A2, "己"))
            + p(ins(A1, r("前の人の字")), ins(A2, r("後の人の字")))  # みな挿入でも、人が違えば段落の印にしない
            + p(gone(A1, "庚"), r("辛"), gone(A1, "壬")))  # 同じ人・日時の削除は、1 つの印にまとめる
    lines = _plain(_lines(write(tmp_path, docx(body))))
    assert lines[0] == (f"甲乙丙丁〔うち変更履歴で挿入された字（{NEW}。{W1}）：「乙」〕〔うち変更履歴で挿入された字（{NEW}。{W2}）：「丁」〕"
                        f"〔変更履歴で削除された字（いまの本文ではない。{W1}）: 戊〕〔変更履歴で削除された字（いまの本文ではない。{W2}）: 己〕")
    assert lines[1] == (f"前の人の字後の人の字〔うち変更履歴で挿入された字（{NEW}。{W1}）：「前の人の字」〕"
                        f"〔うち変更履歴で挿入された字（{NEW}。{W2}）：「後の人の字」〕")
    assert lines[2] == f"辛〔変更履歴で削除された字（いまの本文ではない。{W1}）: 庚壬〕"


def test_docx_change_without_author_or_date_is_marked_without_them(tmp_path):
    body = (p(r("甲"), ins("", r("乙")), gone("", "丙"))
            + p(r("丁"), ins('w:author="係"', r("戊")))
            + p(r("己"), ins('w:date="2026-03-03T00:00:00Z"', r("庚"))))
    lines = _plain(_lines(write(tmp_path, docx(body))))
    assert lines == [
        f"甲乙〔うち変更履歴で挿入された字（{NEW}）：「乙」〕〔変更履歴で削除された字（いまの本文ではない）: 丙〕",
        f"丁戊〔うち変更履歴で挿入された字（{NEW}。変更した人：係）：「戊」〕",
        f"己庚〔うち変更履歴で挿入された字（{NEW}。日時：2026-03-03T00:00:00Z）：「庚」〕",
    ]


def test_docx_moved_text_is_marked_like_inserted_and_deleted_text(tmp_path, monkeypatch):
    # 移る前の字（w:moveFrom）の中の字は、ふつうの w:t で書かれる
    body = p(r("先頭。"), ins(A1, r("移ってきた字"), tag="moveTo")) + p(f'<w:moveFrom w:id="3" {A1}>{r("移る前の字")}</w:moveFrom>', r("残る字。"))
    path = write(tmp_path, docx(body))
    lines = _lines(path)
    assert lines == [
        f"先頭。移ってきた字〔うち変更履歴で挿入された字（{NEW}。{W1}）：「移ってきた字」〕",
        f"残る字。〔変更履歴で削除された字（いまの本文ではない。{W1}）: 移る前の字〕",
    ]
    assert not any("〔非表示の文字" in ln for ln in lines)  # 非表示の字とは分けて示す
    win = api.window(path)[1]
    assert win["tracked_insertions"] == ["段落 1「移ってきた字」"] and win["tracked_deletions"] == ["段落 2「移る前の字」"]
    assert api.audit(path)["ok"]
    _large(monkeypatch, True)
    d = _detail(path)
    assert any("変更履歴で挿入された字" in x for x in d) and any("変更履歴で削除された字が 1 " in x for x in d), d


def test_docx_without_tracked_changes_is_unchanged(tmp_path):
    path = write(tmp_path, docx(p(r("提出先：市役所の窓口。"))))
    win = api.window(path)[1]
    assert _lines(path) == ["提出先：市役所の窓口。"] and win["tracked_insertions"] == [] and win["tracked_deletions"] == []
    assert ("B", "変更履歴") not in _reasons(path)


def test_docx_place_in_an_inserted_paragraph_is_written_at_the_same_id(tmp_path):
    # 段落の前に印がついても、所の id（段落の中の字の位置）と、書き込みの証明・読み直しは変わらない
    plain = write(tmp_path, docx(p(r("氏名：　　　　。"))), "plain.docx")
    path = write(tmp_path, docx(p(ins(A1, r("氏名：　　　　。")))))
    want = [f["id"] for f in api.fields(plain)]
    got = [f["id"] for f in api.fields(path)]
    assert got == want and len(got) == 1
    assert _lines(path)[0].startswith(f"〔変更履歴で挿入された段落（{NEW}。{W1}）〕氏名：")
    res = api.fill(path, {got[0]: "佐藤"}, str(tmp_path / "o.docx"))
    assert res["proof_ok"] and res["verified"] and res["read_back_ok"], res
    assert any("佐藤" in ln and "変更履歴で挿入された" in ln for ln in _lines(tmp_path / "o.docx"))


@pytest.mark.parametrize("chunked", [True, False])
def test_docx_large_road_tells_inserted_text(tmp_path, monkeypatch, chunked):
    body = (p(r("提出先："), ins(A1, r("市役所")), gone(A1, "町役場"), r("の窓口。")) + p(ins(A2, r("まるごと足した段落。")))
            + "".join(p(r(f"ふつうの段落 {i}。")) for i in range(30)))
    path = write(tmp_path, docx(body))
    normal = _lines(path)
    assert f"〔変更履歴で挿入された段落（{NEW}。{W2}）〕まるごと足した段落。" in normal
    _large(monkeypatch, chunked)
    d = _detail(path)
    assert any("変更履歴で挿入された字" in x and "が 2 " in x for x in d), d
    assert any("変更履歴で削除された字が 1 " in x for x in d), d
    assert _range_lines(path, "1-40") == normal


def test_docx_inserted_paragraph_marks_and_rows_are_not_inserted_text(tmp_path, monkeypatch):
    # 段落の終わりの印・表の行の挿入（書式の中の w:ins）は、字の挿入として数えない：段落の区切り・表の行の変更として、別に伝える
    ppr = f'<w:pPr><w:rPr><w:ins w:id="5" {A1}/></w:rPr></w:pPr>'
    row = (f'<w:tbl><w:tr><w:trPr><w:ins w:id="6" {A1}/></w:trPr><w:tc><w:p><w:r><w:t>マス</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
    path = write(tmp_path, docx(p(r("段落。"), ppr=ppr) + row + p(r("終わり。"))))
    win = api.window(path)[1]
    assert win["tracked_insertions"] == [] and win["tracked_marks"] == ["段落 1：挿入された段落の区切り", "段落 2：挿入された表の行"]
    normal = _detail(path)
    assert len(normal) == 1 and "段落の区切り・表の行" in normal[0] and "が 2 か所" in normal[0]
    _large(monkeypatch, True)
    assert _detail(path) == normal


# ====================================================================== .doc（Word 97〜2003）
def _doc_with_who() -> bytes:
    """tracked_changes.doc の作者の表（SttbfRMark）と、変更の日時（DTTM）を、同じ大きさのまま書き換える。
    元：["Unknown", ""]・人の番号 1・1970-01-01 00:00 → ["", "総務課の田中係"]・挿入 2026-11-12 09:30・削除 2026-11-13 18:05。"""
    b = open(DOC_REV, "rb").read()
    old = b"\xff\xff\x02\x00\x00\x00\x07\x00" + "Unknown".encode("utf-16-le") + b"\x00\x00"
    new = b"\xff\xff\x02\x00\x00\x00\x00\x00\x07\x00" + "総務課の田中係".encode("utf-16-le")
    assert b.count(old) == 1 and len(old) == len(new)
    b = b.replace(old, new)

    def make(mi, hr, day, mon, yr, wd):
        return struct.pack("<I", mi | hr << 6 | day << 11 | mon << 16 | (yr - 1900) << 20 | wd << 29)

    for sprm, when in ((b"\x05\x68", make(30, 9, 12, 11, 2026, 4)), (b"\x64\x68", make(5, 18, 13, 11, 2026, 5))):
        a = sprm + b"\x00\x08\x61\x84"
        assert b.count(a) == 1
        b = b.replace(a, sprm + when)
    return b


def test_doc_tells_who_and_when_from_the_author_table(tmp_path):
    path = tmp_path / "a.doc"
    path.write_bytes(_doc_with_who())
    lines = _plain(_lines(path))
    assert "提出先：市役所の窓口。〔変更履歴で削除された字（いまの本文ではない。変更した人：総務課の田中係・日時：2026-11-13 18:05）: 町役場〕" in lines
    assert any(ln.startswith("連絡先：電話") and ln.endswith(f"〔うち変更履歴で挿入された字（{NEW}。変更した人：総務課の田中係・日時：2026-11-12 09:30）：「電話」〕")
               for ln in lines), lines
    win = api.window(str(path))[1]
    assert win["tracked_insertions"] == ["段落 2「電話」"] and win["tracked_deletions"] == ["段落 1「町役場」"]
    assert any("変更履歴で挿入された字" in x for x in _detail(path))


def test_doc_unknown_author_is_not_told_as_a_name():
    # 人の番号が指す名前が空（LibreOffice が書いた .doc）：人は出さず、書いてある日時だけ
    lines = _plain(_lines(DOC_REV))
    assert any(ln.endswith(f"〔うち変更履歴で挿入された字（{NEW}。日時：1970-01-01 00:00）：「電話」〕") for ln in lines), lines
    from todoku.doc.book import DocFile

    assert DocFile(open(DOC_REV, "rb").read(), "a.doc").rev_authors() == ["Unknown", ""]


@pytest.mark.parametrize("chunked", [True, False])
def test_doc_large_road_reads_the_same_lines(tmp_path, monkeypatch, chunked):
    path = tmp_path / "a.doc"
    path.write_bytes(_doc_with_who())
    normal = _lines(path)
    _large(monkeypatch, chunked)
    assert any("変更履歴で挿入された字" in x for x in _detail(path))
    assert _range_lines(path, "1-20") == normal


def test_word6_marks_inserted_text_without_who(tmp_path):
    # Word 6.0/95：挿入・削除の印だけを読む（人・日時は読まない）
    path = tmp_path / "a.doc"
    path.write_bytes(_word6_rev())
    lines = _plain(_lines(path))
    assert any(ln.startswith("連絡先：電話") and ln.endswith(f"〔うち変更履歴で挿入された字（{NEW}）：「電話」〕") for ln in lines), lines
    assert api.window(str(path))[1]["tracked_insertions"] == ["段落 2「電話」"]


# ====================================================================== RTF
PARD = B + "pard" + B + "plain" + B + "f1 "
D1 = 30 | 9 << 6 | 12 << 11 | 11 << 16 | 126 << 20 | 4 << 29  # 2026-11-12 09:30
D2 = 5 | 18 << 6 | 13 << 11 | 11 << 16 | 126 << 20 | 5 << 29  # 2026-11-13 18:05
REVTBL = "{" + B + "*" + B + "revtbl{Unknown;}{" + sj("田中 一郎") + ";}{Sato;}}"


def _rtf(*paras: str, tbl: str = REVTBL) -> bytes:
    return make_rtf(tbl, *paras)


def test_rtf_tells_who_and_when_from_the_revision_table(tmp_path):
    path = tmp_path / "a.rtf"
    path.write_bytes(_rtf(
        PARD + sj("提出先：市役所") + "{" + B + "deleted" + B + "revauthdel2" + B + f"revdttmdel{D2} " + sj("町役場") + "}" + sj("の窓口。") + B + "par",
        PARD + sj("連絡先：") + "{" + B + "revised" + B + "revauth1" + B + f"revdttm{D1} " + sj("電話") + "}" + sj("。") + B + "par",
        PARD + "{" + B + "revised" + B + "revauth2" + B + f"revdttm{D2} " + sj("まるごと足した段落。") + "}" + B + "par",
        PARD + sj("ふつうの段落。") + B + "par"))
    lines = _lines(path)
    assert lines == [
        "提出先：市役所の窓口。〔変更履歴で削除された字（いまの本文ではない。変更した人：Sato・日時：2026-11-13 18:05）: 町役場〕",
        f"連絡先：電話。〔うち変更履歴で挿入された字（{NEW}。変更した人：田中 一郎・日時：2026-11-12 09:30）：「電話」〕",
        f"〔変更履歴で挿入された段落（{NEW}。変更した人：Sato・日時：2026-11-13 18:05）〕まるごと足した段落。",
        "ふつうの段落。",
    ]
    assert "Unknown" not in "".join(lines) and "田中 一郎;" not in "".join(lines)  # 作者の表は、本文として読まない
    win = api.window(str(path))[1]
    assert win["tracked_insertions"] == ["段落 2「電話」", "段落 3「まるごと足した段落。」"]
    assert any("変更履歴で挿入された字" in x and "2 段落" in x for x in _detail(path))


def test_rtf_revision_without_table_or_numbers_is_marked_without_who(tmp_path):
    path = tmp_path / "a.rtf"
    path.write_bytes(_rtf(
        PARD + sj("甲") + "{" + B + "revised " + sj("乙") + "}" + B + "par",
        PARD + sj("丙") + "{" + B + "revised" + B + "revauth7" + B + "revdttm0 " + sj("丁") + "}" + B + "par",  # 表にない番号・日時なし
        tbl=""))
    assert _plain(_lines(path)) == [f"甲乙〔うち変更履歴で挿入された字（{NEW}）：「乙」〕", f"丙丁〔うち変更履歴で挿入された字（{NEW}）：「丁」〕"]


def test_rtf_revision_authors_table():
    from todoku.rtf.document import revision_authors

    blob = ("{" + B + "rtf1{" + B + "*" + B + "revtbl {Unknown;}{" + sj("係") + ";}{A {" + B + "b B};}}" + "{" + B + "*" + B + "revtblx{no;}}}").encode("latin-1")
    assert revision_authors(blob, "cp932")[:2] == ["Unknown", "係"]
    assert revision_authors(b"{\\rtf1 no table}", "cp932") == []


@pytest.mark.parametrize("chunked", [True, False])
def test_rtf_large_road_reads_the_same_lines(tmp_path, monkeypatch, chunked):
    path = tmp_path / "a.rtf"
    path.write_bytes(_rtf(
        PARD + sj("連絡先：") + "{" + B + "revised" + B + "revauth1" + B + f"revdttm{D1} " + sj("電話") + "}" + sj("。") + B + "par",
        *[PARD + sj(f"ふつうの段落 {i}。") + B + "par" for i in range(30)]))
    normal = _lines(path)
    assert normal[0].endswith(f"（{NEW}。変更した人：田中 一郎・日時：2026-11-12 09:30）：「電話」〕")
    _large(monkeypatch, chunked)
    assert any("変更履歴で挿入された字" in x for x in _detail(path))
    assert _range_lines(path, "1-40") == normal


# ====================================================================== OpenDocument の文書
DC = 'xmlns:dc="http://purl.org/dc/elements/1.1/"'


def _region(rid: str, kind: str, who: str | None, when: str | None, inner: str = "") -> str:
    info = "".join([f"<dc:creator {DC}>{who}</dc:creator>" if who is not None else "", f"<dc:date {DC}>{when}</dc:date>" if when is not None else ""])
    return (f'<text:changed-region xml:id="{rid}" text:id="{rid}"><text:{kind}><office:change-info>{info}</office:change-info>{inner}'
            f"</text:{kind}></text:changed-region>")


def _odt(body: str, regions: str) -> bytes:
    return _package("application/vnd.oasis.opendocument.text",
                    f'<office:text><text:tracked-changes text:track-changes="false">{regions}</text:tracked-changes>{body}</office:text>')


def _start(i: str) -> str:
    return f'<text:change-start text:change-id="{i}"/>'


def _end(i: str) -> str:
    return f'<text:change-end text:change-id="{i}"/>'


O1 = "変更した人：田中 一郎・日時：2026-01-05T10:30:00"
O2 = "変更した人：Sato・日時：2026-02-01T09:00:00.123456789"
ODT_REGIONS = (_region("i1", "insertion", "田中 一郎", "2026-01-05T10:30:00") + _region("i2", "insertion", "Sato", "2026-02-01T09:00:00.123456789")
               + _region("i3", "insertion", None, None) + _region("d1", "deletion", "Sato", "2026-02-01T09:00:00.123456789", "<text:p>町役場</text:p>"))


def test_odt_inserted_text_is_marked_with_who_and_when(tmp_path):
    body = (f'<text:p>提出先：{_start("i1")}市役所{_end("i1")}<text:change text:change-id="d1"/>の窓口。</text:p>'
            f'<text:p>前の字。{_start("i2")}足した終わり</text:p><text:p>まるごと足した段落。</text:p><text:p>足した頭{_end("i2")}。後の字</text:p>'
            f'<text:p>甲{_start("i3")}乙{_end("i3")}</text:p>'
            "<text:p>ふつうの段落。</text:p>")
    path = tmp_path / "a.odt"
    path.write_bytes(_odt(body, ODT_REGIONS))
    assert _plain(_lines(path)) == [
        f"提出先：市役所の窓口。〔うち変更履歴で挿入された字（{NEW}。{O1}）：「市役所」〕〔変更履歴で削除された字（いまの本文ではない。{O2}）: 町役場〕",
        f"前の字。足した終わり〔うち変更履歴で挿入された字（{NEW}。{O2}）：「足した終わり」〕",
        f"〔変更履歴で挿入された段落（{NEW}。{O2}）〕まるごと足した段落。",
        f"足した頭。後の字〔うち変更履歴で挿入された字（{NEW}。{O2}）：「足した頭」〕",
        f"甲乙〔うち変更履歴で挿入された字（{NEW}）：「乙」〕",
        "ふつうの段落。",
    ]
    win = api.window(str(path))[1]
    assert win["tracked_insertions"] == ["段落 1「市役所」", "段落 2「足した終わり」", "段落 3「まるごと足した段落。」", "段落 4「足した頭」", "段落 5「乙」"]
    assert any("変更履歴で挿入された字" in x and "5 段落" in x for x in _detail(path))
    assert api.audit(str(path))["ok"]


def test_odt_insertion_in_a_table_cell_and_nested_insertions(tmp_path):
    body = ('<table:table table:name="t1"><table:table-column table:number-columns-repeated="2"/>'
            f'<table:table-row><table:table-cell><text:p>氏名</text:p></table:table-cell>'
            f'<table:table-cell><text:p>{_start("i1")}山田{_end("i1")}</text:p></table:table-cell></table:table-row></table:table>'
            # 挿入の中の挿入：字を直に囲む、内側の挿入の人・日時を示す
            f'<text:p>{_start("i1")}外の字{_start("i2")}中の字{_end("i2")}外の続き{_end("i1")}</text:p>')
    path = tmp_path / "a.odt"
    path.write_bytes(_odt(body, ODT_REGIONS))
    lines = _plain(_lines(path))
    assert any("山田" in ln and f"〔変更履歴で挿入された段落（{NEW}。{O1}）〕" in ln for ln in lines), lines
    assert lines[-1] == (f"外の字中の字外の続き〔うち変更履歴で挿入された字（{NEW}。{O1}）：「外の字」〕〔うち変更履歴で挿入された字（{NEW}。{O2}）：「中の字」〕"
                         f"〔うち変更履歴で挿入された字（{NEW}。{O1}）：「外の続き」〕")


def test_odt_without_insertions_is_unchanged(tmp_path):
    path = tmp_path / "a.odt"
    path.write_bytes(_odt("<text:p>提出先：市役所の窓口。</text:p>", ""))
    assert _lines(path) == ["提出先：市役所の窓口。"] and api.window(str(path))[1]["tracked_insertions"] == []


@pytest.mark.parametrize("chunked", [True, False])
def test_odt_large_road_carries_an_open_insertion_across_chunks(tmp_path, monkeypatch, chunked):
    # まとまり（かたまり）をまたぐ挿入：始まりの印が前のかたまりにある段落も、同じ人・日時の挿入として読む
    body = (f'<text:p>先頭の段落。{_start("i2")}足した終わり</text:p>'
            + "".join(f"<text:p>まるごと足した段落 {i}。</text:p>" for i in range(40))
            + f'<text:p>足した頭{_end("i2")}。後の字</text:p>'
            + "".join(f"<text:p>ふつうの段落 {i}。</text:p>" for i in range(20)))
    path = tmp_path / "a.odt"
    path.write_bytes(_odt(body, ODT_REGIONS))
    normal = _lines(path)
    assert normal[1] == f"〔変更履歴で挿入された段落（{NEW}。{O2}）〕まるごと足した段落 0。" and normal[40].startswith("〔変更履歴で挿入された段落")
    assert normal[41] == f"足した頭。後の字〔うち変更履歴で挿入された字（{NEW}。{O2}）：「足した頭」〕" and normal[42] == "ふつうの段落 0。"
    from todoku.lazy import seqdocs

    monkeypatch.setattr(seqdocs, "UNIT_CHARS", 200)
    monkeypatch.setattr(seqdocs, "UNIT_BYTES", 800)
    _large(monkeypatch, chunked)
    doc = lazy.open_doc(str(path), force=True)
    if chunked:
        assert type(doc).__name__ == "OdtLazy" and len(doc.chunks) > 3
    assert any("変更履歴で挿入された字" in x for x in _detail(path))
    got: list[str] = []
    for a in range(1, 63, 6):
        got += _range_lines(path, f"{a}-{min(a + 5, 62)}")
    assert got == normal


# ====================================================================== 公開の見本（3 つの形式で同じ事実）
@pytest.mark.parametrize("ext, when", [("docx", "2026-01-01T00:00:00Z"), ("odt", "2026-01-01T00:00:00"), ("rtf", "2026-01-01 00:00")])
def test_public_forms_tell_the_same_change_in_each_format(ext, when):
    path = os.path.join(os.path.dirname(__file__), "fixtures", "public", f"form.{ext}")
    who = f"変更した人：todoku・日時：{when}"
    line = next(ln for ln in _lines(path) if ln.startswith("提出先："))
    assert re.sub(r"［[^］]*］", "", line) == (f"提出先：市役所〔うち変更履歴で挿入された字（{NEW}。{who}）：「市役所」〕"
                                             f"〔変更履歴で削除された字（いまの本文ではない。{who}）: 町役場〕")
