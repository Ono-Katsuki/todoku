"""フィールドの結果の字（目次・ページ番号・日付・EQ など）：Word などがフィールドを更新すると作り直す字なので、
その中を書く所として示さず、字を目印にした書き込みも断る。結果が何段落にもなるフィールド（目次）も、
Word（.docx・.doc）・RTF・OpenDocument の文書（目次・索引の本文）で同じ。
フィールドの結果の字のある段落でも、その外の空きは所として示す（EQ で均等に割り付けた見出しの後ろの空欄）。"""

from __future__ import annotations

import os

import pytest
from test_docx import make_docx
from test_openformats import B, make_odt, make_rtf, sj

from todoku import api, lazy, limits
from todoku.lazy import docxdoc

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
TABS = '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>'
EQ = " EQ \\o\\ad(ﾌﾘｶﾞﾅ,　　　　　) "


@pytest.fixture(autouse=True)
def clean():
    api._WINDOWS.clear()
    lazy.clear()
    yield
    api._WINDOWS.clear()
    lazy.clear()


def r(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def fc(kind: str) -> str:
    return f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'


def instr(text: str) -> str:
    return f'<w:r><w:instrText xml:space="preserve">{text}</w:instrText></w:r>'


def fld(code: str, result: str) -> str:
    return fc("begin") + instr(code) + fc("separate") + r(result) + fc("end")


def entry(title: str, page: str, first: str = "") -> str:
    """目次の 1 行：題・タブ（点線のタブ位置）・ページ番号（PAGEREF）。"""
    return f'<w:p>{TABS}{first}{r(title)}<w:r><w:tab/></w:r>{fld(" PAGEREF _Toc1 " + chr(92) + "h ", page)}</w:p>'


def toc_docx(n_fill: int = 0, close: bool = True) -> bytes:
    """目次（TOC フィールドの結果が、段落をまたぐ）と、その後ろの本文。"""
    body = "<w:p>" + r("目次") + "</w:p>"
    body += entry("1　はじめに", "3", first=fc("begin") + instr(' TOC \\o "1-3" ') + fc("separate"))
    body += entry("2　申請の方法", "5")
    for k in range(n_fill):
        body += entry(f"{k + 3}　章の名前 {k}　　　あ", str(k + 7))
    body += entry("9　問い合わせ", "9")
    body += "<w:p>" + (fc("end") if close else "") + r("氏名　　　　　　印") + "</w:p>"
    body += f"<w:p>{TABS}" + r("Name:") + "<w:r><w:tab/></w:r></w:p>"
    return make_docx(body)


def _fields(path) -> list[dict]:
    api._WINDOWS.clear()
    lazy.clear()
    return api.fields(str(path))


def _ids(path) -> list[str]:
    return [f["id"] for f in _fields(path)]


def _lines(path) -> list[str]:
    api._WINDOWS.clear()
    win = api.window(str(path))[1]
    return [ln for st in win["stories"] for ln in st["reading"]]


# ====================================================================== Word（.docx）：段落をまたぐフィールドの結果
def test_docx_contents_lines_give_no_places_and_refuse_writes(tmp_path):
    src = tmp_path / "toc.docx"
    src.write_bytes(toc_docx())
    fl = {f["id"]: f for f in _fields(src)}
    # 目次の行（段落 2〜4）には、空き・点線のタブがあっても所がない。目次の外の空き・点線のタブは所
    assert list(fl) == ["P1.c2", "P5.c2", "P6.c5"]
    assert fl["P5.c2"]["why_code"] == "gap" and fl["P6.c5"]["why_code"] == "leader_tab"
    for old in ("はじめに", "申請の方法", "問い合わせ"):
        with pytest.raises(api.ApiError, match="フィールド（TOC.*の結果の字"):
            api.fill(str(src), {"@1": {"old": old, "new": "書いた字"}}, str(tmp_path / "o.docx"))
    res = api.fill(str(src), {"@1": {"old": "氏名", "new": "氏名（自署）"}, "P6.c5": "Yamada"}, str(tmp_path / "o.docx"))
    assert res["proof_ok"] and res["read_back_ok"], res


def test_docx_field_with_no_end_mark_is_ordinary_text(tmp_path):
    """終わりの印のないフィールド：結果の字と決められないので、ふつうの字として読む（空き・点線のタブは所）。"""
    src = tmp_path / "open.docx"
    src.write_bytes(toc_docx(close=False))
    got = {f["id"]: f["why_code"] for f in _fields(src)}
    assert got["P2.c5"] == "leader_tab" and got["P4.c6"] == "leader_tab" and got["P5.c2"] == "gap"
    res = api.fill(str(src), {"@1": {"old": "はじめに", "new": "序"}}, str(tmp_path / "o.docx"))
    assert res["proof_ok"] and res["read_back_ok"], res


def test_docx_carried_field_reads_the_same_in_ranges(tmp_path, monkeypatch):
    """分けて読む道：まとまりの境をまたぐフィールドも、全体を読んだときと同じ所になる。"""
    src = tmp_path / "big.docx"
    src.write_bytes(toc_docx(n_fill=60))
    whole = _fields(src)
    assert [f["id"] for f in whole] == ["P1.c2", "P65.c2", "P66.c5"]
    monkeypatch.setattr(limits, "LAZY_MIN_BYTES", 0)
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2000)
    monkeypatch.setattr(docxdoc, "CHUNK_BYTES", 2000)
    api._WINDOWS.clear()
    lazy.clear()
    got, lo, total = [], 1, None
    while total is None or lo <= total:
        res = api.fields_range(str(src), {"paragraphs": [lo, lo + 19]}, limit=10000)
        got += res["fields"]
        total = res["total_paragraphs"]
        lo += 20
    assert got == whole
    with pytest.raises(api.ApiError, match="フィールド（TOC.*の結果の字"):
        api.fill(str(src), {"@1": {"old": "章の名前 37", "new": "書いた字"}}, str(tmp_path / "o.docx"))


# ====================================================================== フィールドの結果の字のある段落の、所の決まり
def _eq_docx() -> bytes:
    return make_docx(
        f'<w:p>{fld(EQ, "ﾌﾘｶﾞﾅ")}{r("　　　　　　　　")}</w:p>'
        f'<w:p>{fld(" PAGE ", "3")}</w:p>'
        f'<w:p>{r("申請日")}{fld(" DATE ", "2026/11/11")}</w:p>'
        f'<w:p>{r("申請日")}{fld(" DATE ", "2026/11/11")}{r("　　　　氏名　　　　　印")}</w:p>'
        f'<w:tbl><w:tr><w:tc><w:p>{fld(EQ, "ﾌﾘｶﾞﾅ")}</w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>')


def test_blank_beside_an_eq_label_is_a_place(tmp_path):
    """EQ フィールドで均等に割り付けた見出し：結果の字は見出しの字として読み、その後ろの空きを所にする。"""
    src = tmp_path / "eq.docx"
    src.write_bytes(_eq_docx())
    fl = {f["id"]: f for f in _fields(src)}
    assert fl["P1.c5"]["why_code"] == "gap" and fl["P1.c5"]["facts"]["before"].startswith("ﾌﾘｶﾞﾅ")
    assert fl["T1.R1.C2"]["label"] == "ﾌﾘｶﾞﾅ"  # 左のマスの見出しが、フィールドの結果の字でも見出し
    assert fl["P5.c5"]["why_code"] == "after_label"
    res = api.fill(str(src), {"P1.c5": "ヤマダ", "T1.R1.C2": "ヤマダ"}, str(tmp_path / "o.docx"))
    assert res["proof_ok"] and res["read_back_ok"], res
    with pytest.raises(api.ApiError, match="フィールド（EQ.*の結果の字"):
        api.fill(str(src), {"@1": {"old": "ﾌﾘｶﾞﾅ", "new": "フリガナ", "occurrence": 1}}, str(tmp_path / "o2.docx"))


def test_no_place_inside_or_right_after_a_computed_result(tmp_path):
    """ほかから作る値（ページ番号・日付）：その中と、すぐ後ろの幅のない所は示さない。離れた空きは所。"""
    src = tmp_path / "eq.docx"
    src.write_bytes(_eq_docx())
    ids = _ids(src)
    assert not [i for i in ids if i.startswith(("P2.", "P3."))]
    fl = {f["id"]: f for f in _fields(src)}
    assert fl["P4.c13"]["why_code"] == "gap" and fl["P4.c15"]["why_code"] == "gap"
    assert _lines(src)[1] == "3" and _lines(src)[2] == "申請日2026/11/11"


# ====================================================================== Word 97-2003（.doc）
def test_doc_contents_lines_give_no_places_and_refuse_writes(tmp_path):
    """LibreOffice で .doc にした、上と同じ目次の文書（TOC フィールドの結果が段落をまたぐ）。"""
    src = os.path.join(FIXTURES, "toc_field.doc")
    fl = {f["id"]: f for f in _fields(src)}
    assert list(fl) == ["P1.c2", "P5.c2", "P6.c5"]  # .docx と同じ：目次の外の点線のタブは所（.doc のタブ位置を読む）
    assert fl["P6.c5"]["why_code"] == "leader_tab" and fl["P6.c5"]["facts"]["leaders"] == ["dot"]
    with pytest.raises(api.ApiError, match="フィールド（TOC.*の結果の字"):
        api.fill(src, {"@1": {"old": "申請の方法", "new": "書いた字"}}, str(tmp_path / "o.doc"))
    res = api.fill(src, {"@1": {"old": "氏名", "new": "氏名（自署）"}}, str(tmp_path / "o.doc"))
    assert res["proof_ok"] and res["read_back_ok"], res


# ====================================================================== OpenDocument の文書：目次・索引の本文
def test_odt_index_body_is_generated_text(tmp_path):
    body = ('<text:table-of-content text:name="toc"><text:table-of-content-source text:outline-level="3"/><text:index-body>'
            '<text:index-title text:name="t"><text:p>目次</text:p></text:index-title>'
            '<text:p>1　はじめに<text:tab/>3</text:p><text:p>2　申請の方法<text:tab/>5</text:p></text:index-body></text:table-of-content>'
            '<text:p>氏名　　　　　　印</text:p>')
    src = tmp_path / "idx.odt"
    src.write_bytes(make_odt(body))
    assert _ids(src) == ["P4.c2"]
    assert _lines(src)[1] == "1　はじめに 3"
    with pytest.raises(api.ApiError, match="フィールド（目次）の結果の字"):
        api.fill(str(src), {"@1": {"old": "はじめに", "new": "序"}}, str(tmp_path / "o.odt"))
    res = api.fill(str(src), {"@1": {"old": "氏名", "new": "氏名（自署）"}}, str(tmp_path / "o.odt"))
    assert res["proof_ok"] and res["read_back_ok"], res


# ====================================================================== RTF：結果の中のタブ・段落
def test_rtf_tabs_inside_a_field_result_are_not_places(tmp_path):
    def w(word: str) -> str:
        return B + word

    blob = make_rtf(
        w("pard") + w("f1 ") + sj("目次") + w("par"),
        w("pard") + w("tldot") + w("tqr") + w("tx9000") + w("f1 ") + "{" + w("field") + "{" + w("*") + w("fldinst TOC}{") + w("fldrslt ")
        + sj("1　はじめに") + w("tab 3") + w("par ") + sj("2　申請の方法") + w("tab 5") + "}}" + w("par"),
        w("pard") + w("tldot") + w("tqr") + w("tx9000") + w("f1 ") + "Name:" + w("tab") + w("par"))
    src = tmp_path / "toc.rtf"
    src.write_bytes(blob)
    got = {f["id"]: f["why_code"] for f in _fields(src)}
    assert got == {"P1.c2": "after_label", "P4.c5": "leader_tab"}
    with pytest.raises(api.ApiError, match="フィールド（TOC.*の結果の字"):
        api.fill(str(src), {"@1": {"old": "申請の方法", "new": "書いた字"}}, str(tmp_path / "o.rtf"))


# ====================================================================== 「更新しない」に固定したフィールド
LOCKED = "「更新しない」に固定されていますが"
OPEN = "Word などで開いてフィールドを更新すると"


def _why_refused(path, old: str) -> str:
    """フィールドの結果の字への書き込みを断る理由：固定したフィールドとして（locked）か、ふつうのフィールドとして（open）か。"""
    with pytest.raises(api.ApiError, match="の結果の字です") as e:
        api.fill(str(path), {"@1": {"old": old, "new": "書いた字"}}, str(path) + ".out")
    assert (LOCKED in str(e.value)) != (OPEN in str(e.value)), str(e.value)
    return "locked" if LOCKED in str(e.value) else "open"


def test_docx_locked_field_result_is_refused_and_the_lock_is_told(tmp_path):
    def lfld(code: str, result: str) -> str:
        return (f'<w:r><w:fldChar w:fldCharType="begin" w:fldLock="1"/></w:r>' + instr(code) + fc("separate") + result + fc("end"))

    body = ("<w:p>" + r("日付 ") + lfld(" DATE ", r("1999/01/01")) + r(" と ") + fld(" DATE ", "2000/02/02") + "</w:p>"
            # 固定していないフィールドの結果の中の、固定したフィールド：外の更新で作り直す
            "<w:p>" + r("外 ") + fc("begin") + instr(" IF 1 = 1 ") + fc("separate") + r("外の字 ") + lfld(" AUTHOR ", r("中の固定"))
            + fc("end") + "</w:p>"
            # 固定したフィールドの結果の中：中も固定していれば固定、中が固定していなければ作り直す
            "<w:p>" + r("外 ") + lfld(" QUOTE ", r("固定の外 ") + lfld(" AUTHOR ", r("中も固定")) + fld(" TITLE ", "中は開")) + "</w:p>"
            '<w:p><w:fldSimple w:instr=" PAGE " w:fldLock="1">' + r("77") + "</w:fldSimple></w:p>")
    src = tmp_path / "lock.docx"
    src.write_bytes(make_docx(body))
    assert [_why_refused(src, x) for x in ("1999/01/01", "2000/02/02", "中の固定", "外の字", "中も固定", "固定の外", "中は開", "77")] \
        == ["locked", "open", "open", "open", "locked", "locked", "open", "locked"]
    # 固定していても、結果の字は所にしない（固定していても作り直すアプリがある）
    assert _ids(src) == []


def test_rtf_locked_field_result_is_refused_and_the_lock_is_told(tmp_path):
    def rf(code: str, result: str, lock: bool = False) -> str:
        return "{" + B + "field" + (B + "fldlock" if lock else "") + "{" + B + "*" + B + "fldinst " + code + "}{" + B + "fldrslt " + result + "}}"

    blob = make_rtf("date " + rf("DATE", "1999/01/01", True) + " and " + rf("DATE", "2000/02/02") + B + "par",
                    "o " + rf("IF 1 = 1", "outA " + rf("AUTHOR", "inlocked", True)) + B + "par",
                    "k " + rf("QUOTE", "outC " + rf("AUTHOR", "bothlocked", True) + " " + rf("TITLE", "inopen"), True) + B + "par")
    src = tmp_path / "lock.rtf"
    src.write_bytes(blob)
    assert [_why_refused(src, x) for x in ("1999/01/01", "2000/02/02", "inlocked", "outA", "bothlocked", "outC", "inopen")] \
        == ["locked", "open", "open", "open", "locked", "locked", "open"]


def test_doc_locked_fields_written_by_libreoffice(tmp_path, monkeypatch):
    # .docx の固定した DATE・TIME（w:fldLock）を LibreOffice が .doc に書き出したファイル：終わりの印の fLocked が立つ
    import shutil

    from todoku.doc.book import DocFile

    src = tmp_path / "locked_fields.doc"
    shutil.copy(os.path.join(os.path.dirname(__file__), "fixtures", "locked_fields.doc"), src)
    doc = DocFile(src.read_bytes(), "locked_fields.doc")
    ends = doc._locked_ends()
    assert len(ends) == 2 and all(doc.text[cp] == "\x15" for cp in ends)
    assert sorted(doc.text[doc.text.rindex("\x14", 0, cp) + 1:cp] for cp in ends) == ["03:33", "1999/01/01"]
    assert [_why_refused(src, x) for x in ("1999/01/01", "03:33", "2026/11/12", "frozenauthor")] == ["locked", "locked", "open", "open"]
    # 大きな書類の道（段落の範囲で読む）でも、同じ理由で断る
    monkeypatch.setattr(limits, "LAZY_MIN_BYTES", 0)
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 300)
    api._WINDOWS.clear()
    lazy.clear()
    assert api.fields_range(str(src), {"paragraphs": [1, 5]}, limit=10000)["total_paragraphs"] > 5
    assert [_why_refused(src, x) for x in ("1999/01/01", "03:33", "2026/11/12")] == ["locked", "locked", "open"]


def _word6_with_fields(text: str, locked: tuple, shift: int = 0) -> bytes:
    """Word 6.0 の文書と、フィールドの印の表（PlcffldMom）：結果の字が locked のどれかで始まるフィールドの終わりの印に fLocked。"""
    import struct

    from test_legacy import make_word6

    marks, stack = [], []
    for cp, ch in enumerate(text):
        if ch == "\x13":
            stack.append(cp)
            marks.append((cp, 0x13, 0))
        elif ch == "\x14":
            marks.append((cp, 0x14, 0xFF))
        elif ch == "\x15":
            b = stack.pop()
            lock = text[text.index("\x14", b) + 1:cp].startswith(locked)
            marks.append((cp, 0x15, 0x80 | (0x10 if lock else 0) | (0x40 if stack else 0)))
    plc = struct.pack(f"<{len(marks) + 1}I", *[m[0] + shift for m in marks], len(text)) + b"".join(bytes([m[1], m[2]]) for m in marks)
    return make_word6(text, enc="cp1252", lid=0x0409, more={"PlcfFldMom": plc})


W6_FIELDS = ("date \x13 DATE \x141999/01/01\x15 and \x13 DATE \x142000/02/02\x15\r"
             "o \x13 IF 1 = 1 \x14outA \x13 AUTHOR \x14inlocked\x15\x15\r"
             "k \x13 QUOTE \x14outC \x13 AUTHOR \x14bothlocked\x15 \x13 TITLE \x14inopen\x15\x15\r")


def test_word6_locked_field_result_is_refused_and_the_lock_is_told(tmp_path):
    src = tmp_path / "lock6.doc"
    src.write_bytes(_word6_with_fields(W6_FIELDS, ("1999/01/01", "inlocked", "bothlocked", "outC")))
    assert [_why_refused(src, x) for x in ("1999/01/01", "2000/02/02", "inlocked", "outA", "bothlocked", "outC", "inopen")] \
        == ["locked", "open", "open", "open", "locked", "locked", "open"]


def test_doc_field_table_that_does_not_point_at_the_marks_is_not_used(tmp_path):
    # 表の位置が、字の中の終わりの印（0x15）と合わないときは、固定として読まない
    src = tmp_path / "off6.doc"
    src.write_bytes(_word6_with_fields(W6_FIELDS, ("1999/01/01",), shift=1))
    assert _why_refused(src, "1999/01/01") == "open"
