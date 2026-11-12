"""変更履歴で削除された字・取り消し線・字の色は、形式や道によらず伝える（Word の .docx と同じ印）。
OpenDocument の文書・RTF の変更履歴（読む・段階・書くときの警告）、取り消し線と字の色の事実（OpenDocument・RTF・PowerPoint・
HWPX・HWP 5）、ふつうの道と大きな書類の道を、手で組んだ文書で確かめる。"""

from __future__ import annotations

import io
import json
import os
import re
import struct
import zipfile

import pptx_maker as pm
import pytest
from test_hwpx import HEADER, make_hwp5, make_hwpx, para, run_, t
from test_hwpx import NS as HWPX_NS
from test_hwpx import XML as HWPX_XML
from test_openformats import NS as ODF_NS
from test_openformats import _package, make_rtf, sj

from todoku import api, lazy, limits, mcp_server

GONE = "〔変更履歴で削除された字（いまの本文ではない）: "


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _lines(path) -> list[str]:
    api._WINDOWS.clear()
    win = api.window(str(path))[1]
    return [ln for st in win["stories"] for ln in st["reading"]]


def _plain(lines: list[str]) -> list[str]:
    """「？」の所の印を除いた行（印の位置でなく、字と事実を確かめる）。"""
    return [re.sub(r"［P\d+\.c\d+(?:\.\d+)?？[^］]*］", "", ln) for ln in lines]


def _reasons(path) -> list[tuple[str, str]]:
    g = api.grade(str(path))
    return [(r["grade"], r["title"]) for r in g["reasons"]]


# ====================================================================== OpenDocument の文書：変更履歴
DC = 'xmlns:dc="http://purl.org/dc/elements/1.1/"'
INFO = f"<office:change-info><dc:creator {DC}>係</dc:creator><dc:date {DC}>2026-01-01T00:00:00</dc:date></office:change-info>"


def _region(rid: str, kind: str, inner: str = "") -> str:
    return f'<text:changed-region xml:id="{rid}" text:id="{rid}"><text:{kind}>{INFO}{inner}</text:{kind}></text:changed-region>'


def _odt(body: str, regions: str, track: str = "false", auto: str = "") -> bytes:
    return _package("application/vnd.oasis.opendocument.text",
                    f'<office:text><text:tracked-changes text:track-changes="{track}">{regions}</text:tracked-changes>{body}</office:text>', auto)


ODT_REGIONS = (_region("d1", "deletion", "<text:p>町役場</text:p>")
               + _region("d2", "deletion", "<text:p>消えた終わり</text:p><text:p>まるごと消えた段落</text:p><text:p>消えた頭</text:p>")
               + _region("d3", "deletion", "<text:p>山田</text:p>")
               + _region("d4", "deletion", "<text:p>段落の外で消えた字</text:p>")
               + _region("i1", "insertion"))
ODT_BODY = ('<text:p>提出先：市役所<text:change text:change-id="d1"/>の窓口。</text:p>'
            '<text:p>一行目の残る字<text:change text:change-id="d2"/>三行目の残る字。</text:p>'
            '<table:table table:name="t1"><table:table-column table:number-columns-repeated="2"/>'
            '<table:table-row><table:table-cell><text:p>氏名</text:p></table:table-cell>'
            '<table:table-cell><text:p><text:change text:change-id="d3"/></text:p></table:table-cell></table:table-row></table:table>'
            '<text:change text:change-id="d4"/>'
            '<text:p>連絡先：<text:change-start text:change-id="i1"/>電話　　　　<text:change-end text:change-id="i1"/>。</text:p>')


def test_odt_tracked_deletions_are_told_and_not_read_as_text(tmp_path):
    p = tmp_path / "a.odt"
    p.write_bytes(_odt(ODT_BODY, ODT_REGIONS))
    lines = _plain(_lines(p))
    assert f"提出先：市役所の窓口。{GONE}町役場〕" in lines
    # 段落をまたぐ削除：消えた段落の切れ目を ⏎ で示す
    assert "一行目の残る字三行目の残る字。〔変更履歴で削除された字（いまの本文ではない。⏎ は段落の切れ目）: 消えた終わり⏎まるごと消えた段落⏎消えた頭〕" in lines
    # 消えた字だけのマス：欄の印も、消えた字も出す
    assert f"氏名 ｜ ［T1.R1.C2］ {GONE}山田〕" in lines
    # 段落の外（段落のあいだ）の削除は、次の段落の頭に付ける
    assert any(ln.startswith("連絡先：") and f"{GONE}段落の外で消えた字〕" in ln for ln in lines)
    fl = {f["id"]: f for f in api.fields(str(p))}
    assert fl["T1.R1.C2"]["tracked_deleted"] == "山田" and not fl["T1.R1.C2"].get("value")
    assert ("B", "変更履歴") in _reasons(p)
    au = api.audit(str(p))
    assert au["ok"], [c for c in au["checks"] if not c["ok"]]


def test_odt_without_tracked_changes_is_unchanged(tmp_path):
    p = tmp_path / "a.odt"
    p.write_bytes(_odt("<text:p>提出先：市役所の窓口。</text:p>", ""))
    assert _lines(p) == ["提出先：市役所の窓口。"]
    assert ("B", "変更履歴") not in _reasons(p)


def test_odt_fill_tells_unresolved_changes_insertion_and_tracking(tmp_path):
    p = tmp_path / "a.odt"
    p.write_bytes(_odt(ODT_BODY, ODT_REGIONS, track="true"))
    spot = next(f["id"] for f in api.fields(str(p)) if f["id"].startswith("P") and "電話" in json.dumps(f, ensure_ascii=False))
    r = api.fill(str(p), {"T1.R1.C2": "佐藤", spot: "03-0000-0000"}, str(tmp_path / "o.odt"))
    assert r["proof_ok"] and r["verified"] and r["read_back_ok"], r
    w = r["warnings"]
    assert "まだ承諾も却下もされていない変更履歴（挿入 1・削除 4）" in w[0]
    assert any(x.startswith(f"{spot}: 書いた所は、変更履歴の挿入の中にあります") for x in w)
    assert not any(x.startswith("T1.R1.C2: 書いた所は、変更履歴の挿入") for x in w)
    assert any("「変更履歴の記録」がオン" in x for x in w)
    out = _plain(_lines(tmp_path / "o.odt"))
    assert any("［T1.R1.C2＝佐藤］" in ln and f"{GONE}山田〕" in ln for ln in out)  # 消えた字は、書いた後も消えた字のまま


def test_odt_fill_without_changes_has_no_revision_warning(tmp_path):
    p = tmp_path / "a.odt"
    p.write_bytes(_odt('<table:table table:name="t1"><table:table-column table:number-columns-repeated="2"/><table:table-row>'
                       '<table:table-cell><text:p>氏名</text:p></table:table-cell><table:table-cell><text:p/></table:table-cell>'
                       '</table:table-row></table:table>', ""))
    r = api.fill(str(p), {"T1.R1.C2": "佐藤"}, str(tmp_path / "o.odt"))
    assert r["proof_ok"] and not [x for x in r["warnings"] if "変更履歴" in x]


# ====================================================================== RTF：変更履歴
B = "\\"
PARD = B + "pard" + B + "f1 "
RTF_BODY = ("{" + B + "colortbl;" + B + "red0" + B + "green0" + B + "blue0;" + B + "red255" + B + "green0" + B + "blue0;}",
            PARD + sj("提出先：市役所") + "{" + B + "deleted " + sj("町役場") + "}" + sj("の窓口。") + B + "par",
            PARD + sj("連絡先：") + "{" + B + "revised " + sj("電話　　　　") + "}" + sj("。") + B + "par",
            PARD + sj("氏名：　　　　。") + B + "par")


def test_rtf_revisions_are_counted_and_told_at_fill(tmp_path):
    from todoku.rtf.document import RtfDocument

    p = tmp_path / "a.rtf"
    p.write_bytes(make_rtf(B + "revisions", *RTF_BODY))
    doc = RtfDocument(p.read_bytes(), "a.rtf")
    assert doc.revisions == {"挿入": 1, "削除": 1, "書式": 0} and doc.tracking
    lines = _plain(_lines(p))
    assert f"提出先：市役所の窓口。{GONE}町役場〕" in lines
    assert ("B", "変更履歴") in _reasons(p)
    spot = next(f["id"] for f in api.fields(str(p)) if "電話" in json.dumps(f, ensure_ascii=False))
    r = api.fill(str(p), {spot: "03-0000-0000"}, str(tmp_path / "o.rtf"))
    assert r["proof_ok"] and r["verified"] and r["read_back_ok"], r
    assert "まだ承諾も却下もされていない変更履歴（挿入 1・削除 1）" in r["warnings"][0]
    assert any("「変更履歴の記録」がオン" in x for x in r["warnings"])
    # 書いた字は、挿入の印を打ち消して書く（変更履歴の外の、ふつうの字）
    assert any(x.startswith(f"{spot}: 書く所に残っていた書式（変更履歴の挿入の印）") for x in r["warnings"])
    assert re.search(rb"\{\\revised0 [^{}]*\}", (tmp_path / "o.rtf").read_bytes())
    after = RtfDocument((tmp_path / "o.rtf").read_bytes(), "o.rtf")
    assert after.revisions["削除"] == 1


def test_rtf_without_revisions_has_no_revision_warning(tmp_path):
    p = tmp_path / "a.rtf"
    p.write_bytes(make_rtf(PARD + sj("氏名：　　　　。") + B + "par"))
    spot = api.fields(str(p))[0]["id"]
    r = api.fill(str(p), {spot: "佐藤"}, str(tmp_path / "o.rtf"))
    assert r["proof_ok"] and not [x for x in r["warnings"] if "変更履歴" in x]


# ====================================================================== 古い Word（.doc）：変更履歴
# tests/fixtures/tracked_changes.doc：上の RTF_BODY（「変更履歴の記録」がオン）の RTF を、LibreOffice で .doc（Word 97）にしたもの。
DOC_REV = os.path.join(os.path.dirname(__file__), "fixtures", "tracked_changes.doc")


def test_doc_tracked_deletion_is_told_as_a_revision_and_revisions_are_told_at_fill(tmp_path):
    from todoku.doc.book import DocFile

    doc = DocFile(open(DOC_REV, "rb").read(), "a.doc")
    assert doc.revisions == {"挿入": 1, "削除": 1} and doc.tracking
    raw = _lines(DOC_REV)
    lines = _plain(raw)
    assert f"提出先：市役所の窓口。{GONE}町役場〕" in lines
    assert not any("〔非表示の文字" in ln for ln in raw)  # 非表示の字とは分けて示す
    assert ("B", "変更履歴") in _reasons(DOC_REV)
    fl = api.fields(DOC_REV)
    spot = next(f["id"] for f in fl if "電話" in json.dumps(f, ensure_ascii=False))
    name = next(f["id"] for f in fl if "氏名" in json.dumps(f, ensure_ascii=False))
    r = api.fill(DOC_REV, {spot: "03-0000-0000", name: "佐藤"}, str(tmp_path / "o.doc"))
    assert r["proof_ok"] and r["verified"] and r["read_back_ok"], r
    assert "まだ承諾も却下もされていない変更履歴（挿入 1・削除 1）" in r["warnings"][0]
    assert any("「変更履歴の記録」がオン" in x for x in r["warnings"])
    # 書いた字は、挿入の印を引き継がない（変更履歴の外の、ふつうの字）
    assert any(x.startswith(f"{spot}: 書く所に残っていた書式（変更履歴の挿入の印）") for x in r["warnings"])
    after = DocFile((tmp_path / "o.doc").read_bytes(), "o.doc")
    assert after.revisions["削除"] == 1 and after.tracking
    out = _plain(_lines(tmp_path / "o.doc"))
    assert any("提出先：市役所の窓口。" in ln and f"{GONE}町役場〕" in ln for ln in out)  # 消えた字は、書いた後も消えた字のまま
    assert any("03-0000-0000" in ln for ln in out) and any("佐藤" in ln for ln in out)


def test_doc_without_revisions_has_no_revision_warning(tmp_path):
    from todoku.doc.book import DocFile

    src = os.path.join(os.path.dirname(__file__), "fixtures", "symbol_fonts.doc")
    doc = DocFile(open(src, "rb").read(), "a.doc")
    assert doc.revisions == {"挿入": 0, "削除": 0} and not doc.tracking
    assert ("B", "変更履歴") not in _reasons(src)


def _word6_rev(track: bool = True, marks: bool = True) -> bytes:
    from test_legacy import make_word6

    t = "提出先：市役所町役場の窓口。\r連絡先：電話　　　　。\r氏名：　　　　。\r"
    cp = lambda s: len(t[:t.index(s)].encode("cp932"))  # noqa: E731
    dop = bytearray(16)
    dop[5] = 0x80 if track else 0  # fRevMarking
    # Word 6.0/95 の命令：65 変更履歴の削除（sprmCFStrikeRM）、66 変更履歴の挿入（sprmCFRMark）
    chars = [(cp("町役場"), cp("町役場") + 6, bytes([65, 1])), (cp("電話"), cp("電話") + 12, bytes([66, 1]))] if marks else []
    return make_word6(t, {}, chars, dop=bytes(dop))


def test_word6_tracked_deletion_is_told_as_a_revision_and_revisions_are_told_at_fill(tmp_path):
    from todoku.doc.word6 import Word6File

    p = tmp_path / "a.doc"
    p.write_bytes(_word6_rev())
    doc = Word6File(p.read_bytes(), "a.doc")
    assert doc.revisions == {"挿入": 1, "削除": 1} and doc.tracking
    raw = _lines(p)
    assert f"提出先：市役所の窓口。{GONE}町役場〕" in _plain(raw)
    assert not any("〔非表示の文字" in ln for ln in raw)
    assert ("B", "変更履歴") in _reasons(p)
    fl = api.fields(str(p))
    spot = next(f["id"] for f in fl if "電話" in json.dumps(f, ensure_ascii=False))
    name = next(f["id"] for f in fl if "氏名" in json.dumps(f, ensure_ascii=False))
    r = api.fill(str(p), {spot: "0312", name: "佐藤"}, str(tmp_path / "o.doc"))
    assert r["proof_ok"] and r["verified"] and r["read_back_ok"], r
    w = r["warnings"]
    assert "まだ承諾も却下もされていない変更履歴（挿入 1・削除 1）" in w[0] and "上書きして書きました" in w[0]
    assert any(x.startswith(f"{spot}: 書いた所は、変更履歴の挿入の中にあります") for x in w)
    assert not any(x.startswith(f"{name}: 書いた所は、変更履歴の挿入") for x in w)
    assert any("「変更履歴の記録」がオン" in x for x in w)
    after = Word6File((tmp_path / "o.doc").read_bytes(), "o.doc")
    assert after.revisions == {"挿入": 1, "削除": 1}  # 字のバイトを上書きするので、印は書く所のまま


def test_word6_without_revisions_has_no_revision_warning(tmp_path):
    from todoku.doc.word6 import Word6File

    p = tmp_path / "a.doc"
    p.write_bytes(_word6_rev(track=False, marks=False))
    doc = Word6File(p.read_bytes(), "a.doc")
    assert doc.revisions == {"挿入": 0, "削除": 0} and not doc.tracking
    assert ("B", "変更履歴") not in _reasons(p)
    name = next(f["id"] for f in api.fields(str(p)) if "氏名" in json.dumps(f, ensure_ascii=False))
    r = api.fill(str(p), {name: "佐藤"}, str(tmp_path / "o.doc"))
    assert r["proof_ok"] and not [x for x in r["warnings"] if "変更履歴" in x]


@pytest.mark.parametrize("chunked", [True, False])
@pytest.mark.parametrize("kind", ["word97", "word6"])
def test_large_road_tells_tracked_deletions_of_doc(tmp_path, monkeypatch, kind, chunked):
    p = tmp_path / "a.doc"
    p.write_bytes(open(DOC_REV, "rb").read() if kind == "word97" else _word6_rev())
    normal = _lines(p)
    assert any(GONE + "町役場〕" in ln for ln in normal)
    _large(monkeypatch, chunked)
    assert ("B", "変更履歴") in _reasons(p)
    assert _range_lines(p, "1-20") == normal  # ふつうの道の窓と、同じ行


# ====================================================================== 大きな書類の道
def _large(monkeypatch, chunked: bool) -> None:
    """大きな書類の道にする。chunked：まとまりごとに読む道（小さなファイルは、窓を作ってから分ける道）。"""
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 300)
    if chunked:
        monkeypatch.setattr(limits, "LAZY_MIN_BYTES", 0)
    lazy.clear()
    api._WINDOWS.clear()


def _range_lines(path, rng: str) -> list[str]:
    got = mcp_server._read_window({"path": str(path), "paragraphs": rng})
    return [re.sub(r"^〔段落 \d+(?:〜\d+)?〕", "", ln) for ln in got["content"][0]["text"].split("\n")]


@pytest.mark.parametrize("chunked", [True, False])
@pytest.mark.parametrize("ext", ["odt", "rtf"])
def test_large_road_tells_tracked_deletions_and_reads_paragraph_ranges(tmp_path, monkeypatch, ext, chunked):
    p = tmp_path / f"a.{ext}"
    p.write_bytes(_odt(ODT_BODY, ODT_REGIONS) if ext == "odt" else make_rtf(*RTF_BODY))
    normal = _lines(p)
    assert any(GONE + "町役場〕" in ln for ln in normal)
    _large(monkeypatch, chunked)
    assert type(lazy.open_doc(str(p), force=True)).__name__ == ({"odt": "OdtLazy", "rtf": "ViewLazy"}[ext] if chunked else "BuiltDoc")
    assert ("B", "変更履歴") in _reasons(p)
    assert _range_lines(p, "1-20") == normal  # ふつうの道の窓と、同じ行


# ====================================================================== 取り消し線・字の色
def test_odt_strike_and_color_from_styles(tmp_path):
    auto = ('<style:style style:name="P1" style:family="paragraph"><style:text-properties fo:color="#ff0000"/></style:style>'
            '<style:style style:name="T1" style:family="text"><style:text-properties style:text-line-through-style="solid"/></style:style>'
            '<style:style style:name="T2" style:family="text" style:parent-style-name="T1"><style:text-properties fo:color="#0000ff"/></style:style>'
            '<style:style style:name="T3" style:family="text"><style:text-properties style:use-window-font-color="true" '
            'style:text-line-through-style="none"/></style:style>')
    body = ('<text:p>ふつうの字がいちばん多い段落です。ふつうの字がいちばん多い段落です。</text:p>'
            '<text:p>旧：<text:span text:style-name="T1">千円</text:span>。新：二千円。</text:p>'
            '<text:p text:style-name="P1">赤い段落の字。</text:p>'
            '<text:p text:style-name="P1">赤の中の<text:span text:style-name="T2">青い取り消し</text:span>と<text:span text:style-name="T3">自動の色</text:span>。</text:p>')
    p = tmp_path / "a.odt"
    p.write_bytes(_package("application/vnd.oasis.opendocument.text", f"<office:text>{body}</office:text>", auto))
    lines = _plain(_lines(p))
    assert "旧：千円。新：二千円。〔うち取り消し線：「千円」〕" in lines
    assert "〔字の色 #FF0000〕赤い段落の字。" in lines
    assert "赤の中の青い取り消しと自動の色。〔うち字の色 #FF0000：「赤の中の」〕〔うち取り消し線・字の色 #0000FF：「青い取り消し」〕〔うち字の色 #FF0000：「と」〕〔うち字の色 #FF0000：「。」〕" in lines


def test_rtf_strike_and_color(tmp_path):
    p = tmp_path / "a.rtf"
    p.write_bytes(make_rtf(RTF_BODY[0],
                           PARD + sj("ふつうの字がいちばん多い段落です。") + B + "par",
                           PARD + sj("旧：") + "{" + B + "strike " + sj("千円") + "}" + sj("。新：") + "{" + B + "cf2 " + sj("二千円") + "}"
                           + sj("。") + B + "par"))
    assert "旧：千円。新：二千円。〔うち取り消し線：「千円」〕〔うち字の色 #FF0000：「二千円」〕" in _plain(_lines(p))


THEME_ACCENT = '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'


def _rezip(blob: bytes, change) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, change(info.filename, zin.read(info.filename)))
    return out.getvalue()


def test_pptx_strike_and_color_of_runs(tmp_path):
    def r(rpr: str, text: str) -> str:
        return f"<a:r>{rpr}<a:t>{text}</a:t></a:r>"

    def para_(*runs: str) -> str:
        return "<a:p>" + "".join(runs) + "</a:p>"

    paras = (para_(r('<a:rPr lang="ja-JP"/>', "ふつうの字がいちばん多い段落です。ふつうの字がいちばん多い段落です。"))
             + para_(r('<a:rPr lang="ja-JP"/>', "旧："), r('<a:rPr lang="ja-JP" strike="sngStrike"/>', "千円"),
                     r('<a:rPr lang="ja-JP"/>', "。新："),
                     r('<a:rPr lang="ja-JP"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:rPr>', "二千円"),
                     r('<a:rPr lang="ja-JP"/>', "。"))
             + para_(r('<a:rPr lang="ja-JP"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:rPr>', "テーマの色。"))
             + para_(r('<a:rPr lang="ja-JP"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr>', "テーマの字の色（黒）。"))
             + para_(r('<a:rPr lang="ja-JP"><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:rPr>',
                       "変えた色。"))
             + para_(r('<a:rPr lang="ja-JP"><a:noFill/></a:rPr>', "塗りのない字。")))
    blob = _rezip(pm.make([pm.sp(paras)]), lambda name, data: data.replace(
        b'<a:accent2><a:srgbClr val="000000"/></a:accent2>', THEME_ACCENT.encode()) if name.endswith("theme1.xml") else data)
    p = tmp_path / "a.pptx"
    p.write_bytes(blob)
    lines = _plain(_lines(p))
    assert "旧：千円。新：二千円。〔うち取り消し線：「千円」〕〔うち字の色 #FF0000：「二千円」〕" in lines
    assert "〔字の色 #ED7D31〕テーマの色。" in lines
    assert "テーマの字の色（黒）。" in lines
    assert "〔字の色（#ED7D31 を変えた色）〕変えた色。" in lines
    assert "〔字の色（塗りなし）〕塗りのない字。" in lines


def test_hwpx_strike_and_color_of_char_shapes(tmp_path):
    header = HEADER.replace('<hh:charProperties itemCnt="2">', '<hh:charProperties itemCnt="4">').replace(
        "</hh:charProperties>",
        '<hh:charPr id="2" height="1000" textColor="#FF0000" borderFillIDRef="1"><hh:underline type="NONE" shape="SOLID" color="#000000"/>'
        '<hh:strikeout shape="NONE" color="#000000"/></hh:charPr>'
        '<hh:charPr id="3" height="1000" textColor="#000000" borderFillIDRef="1"><hh:underline type="NONE" shape="SOLID" color="#000000"/>'
        '<hh:strikeout shape="SOLID" color="#000000"/></hh:charPr></hh:charProperties>')
    assert header != HEADER
    sec = (HWPX_XML + f"<hs:sec {HWPX_NS}>"
           + para(run_(t("보통 글자가 가장 많은 문단입니다. 보통 글자가 가장 많은 문단입니다.")))
           + para(run_(t("구: ")) + run_(t("천 원"), "3") + run_(t(". 신: ")) + run_(t("이천 원"), "2") + run_(t(".")))
           + "</hs:sec>")
    p = tmp_path / "a.hwpx"
    p.write_bytes(_rezip(make_hwpx(sec), lambda name, data: header.encode() if name == "Contents/header.xml" else data))
    assert "구: 천 원. 신: 이천 원.〔うち取り消し線：「천 원」〕〔うち字の色 #FF0000：「이천 원」〕" in _plain(_lines(p))


def test_hwp5_strike_and_color_of_char_shapes(tmp_path):
    from test_hwpx import _para, _rec, _w
    from test_officecrypto import build_cfb

    def shape(strike: bool, rgb: tuple[int, int, int]) -> bytes:
        attr = (1 << 18) if strike else 0
        color = rgb[0] | (rgb[1] << 8) | (rgb[2] << 16)
        return _rec(0x15, 0, b"\0" * 42 + struct.pack("<iI", 1000, attr) + b"\0\0" + struct.pack("<IIII", color, 0, 0xFFFFFFFF, 0))

    docinfo = shape(False, (0, 0, 0)) + shape(True, (0, 0, 0)) + shape(False, (255, 0, 0))
    long_ = "보통 글자가 가장 많은 문단입니다. 보통 글자가 가장 많은 문단입니다."
    text = "구: 천 원. 신: 이천 원."
    a, b_ = text.index("천"), text.index("이천")
    runs = struct.pack("<10I", 0, 0, a, 1, a + 3, 0, b_, 2, b_ + 4, 0)
    body = _para(_w(long_) + [13]) + _para(_w(text) + [13], ctrls=_rec(0x44, 1, runs))
    head = (b"HWP Document File".ljust(32, b"\0") + struct.pack("<II", 0x05010000, 0)).ljust(256, b"\0")
    p = tmp_path / "a.hwp"
    p.write_bytes(build_cfb({"FileHeader": head, "DocInfo": docinfo, "BodyText/Section0": body}))
    assert make_hwp5()  # 同じ作り方（test_hwpx）の文書が読めること
    assert "구: 천 원. 신: 이천 원.〔うち取り消し線：「천 원」〕〔うち字の色 #FF0000：「이천 원」〕" in _plain(_lines(p))


@pytest.mark.parametrize("chunked", [True, False])
def test_large_road_uses_the_normal_color_of_the_whole_document(tmp_path, monkeypatch, chunked):
    # 大きな書類の道も、「ふつうの色」をまとまりごとには決めない：文書の全体で数えた色（ふつうの道の窓と同じ行）
    auto = '<style:style style:name="P1" style:family="paragraph"><style:text-properties fo:color="#333333"/></style:style>'
    body = "".join(f'<text:p text:style-name="P1">濃い灰色の段落 {i}。</text:p>' for i in range(40)) + "<text:p>黒い段落。</text:p>"
    p = tmp_path / "a.odt"
    p.write_bytes(_package("application/vnd.oasis.opendocument.text", f"<office:text>{body}</office:text>", auto))
    win = api.window(str(p))[1]
    normal = _lines(p)
    assert win["normal_font_color"] == "333333" and normal[0] == "濃い灰色の段落 0。" and normal[-1] == "〔字の色 #000000〕黒い段落。"
    from todoku.lazy import seqdocs

    monkeypatch.setattr(seqdocs, "UNIT_CHARS", 200)  # 黒い段落だけのまとまりができる大きさ
    monkeypatch.setattr(seqdocs, "UNIT_BYTES", 800)
    _large(monkeypatch, chunked)
    doc = lazy.open_doc(str(p), force=True)
    if chunked:
        assert len(doc.chunks) > 3 and doc.overview()["normal_font_color"] == "333333"
    assert "ふつうの字の色" in doc.overview_text() and "#333333" in doc.overview_text()
    assert _range_lines(p, "1-41") == normal


def test_normal_color_does_not_depend_on_the_order_of_counting():
    from collections import Counter

    from todoku.docx.fontfacts import most_common_color

    assert most_common_color(Counter()) == "000000"
    assert most_common_color(Counter({"FF0000": 5, "000000": 5})) == most_common_color(Counter({"000000": 5, "FF0000": 5})) == "000000"
    assert most_common_color(Counter({"FF0000": 5, "0000FF": 5})) == most_common_color(Counter({"0000FF": 5, "FF0000": 5})) == "0000FF"
    assert most_common_color(Counter({"FF0000": 6, "000000": 5})) == "FF0000"
