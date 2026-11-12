"""段階の直し（続き）：文字コードの候補とエージェントの選択・窓の読み落とし・画像だけのシート・中身が画像の文書。
どれも、テストの中で小さなファイルを組み立てて確かめる。"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku import api
from todoku.datafile.textenc import decode, declared
from todoku.grade import grade_blob

JA = "氏名,住所\n山田太郎,\n"  # Shift_JIS（CP932）としても GBK・UHC としても、仕様どおりにきれいに読める


# ====================================================================== 文字コード
def test_ambiguous_encoding_is_read_b_with_candidates(tmp_path):
    p = tmp_path / "a.csv"
    p.write_bytes(JA.encode("cp932"))
    g = api.grade(str(p))
    assert g["grade"] == "B" and g["write"]["grade"] == "B"
    r = next(r for r in g["reasons"] if r["title"] == "文字コードを規則では決められない")
    assert {"cp932", "gbk"} <= set(r["where"])
    _k, doc, parts = api.window(str(p))
    assert "gbk" in {c["encoding"] for c in doc["format"]["encoding_candidates"]}
    md = "\n".join(t for _n, t in parts)
    assert 'encoding: "gbk"' in md and 'encoding: "cp932"' in md  # 選んだら渡す名前を、窓に示す


def test_agent_chosen_encoding_flows_through_every_tool(tmp_path):
    p = tmp_path / "a.csv"
    p.write_bytes(JA.encode("cp932"))
    g = api.grade(str(p), encoding="cp932")
    assert g["grade"] == "S" and g["write"]["grade"] == "A", g
    chosen = [r for r in g["reasons"] if r["by"] == "agent_choice"]
    assert [r["where"] for r in chosen] == [["cp932"]]
    # 名前は、別名（Shift_JIS）でも、窓の候補の名前でも
    _k, doc, _p = api.window(str(p), encoding="Shift_JIS")
    fm = doc["format"]
    assert fm["encoding"] == "cp932" and fm["encoding_chosen_by"] == "agent" and "encoding_candidates" not in fm
    assert api.fields(str(p), encoding="cp932")
    out = tmp_path / "b.csv"
    res = api.fill(str(p), {"R2.C2": "東京都中央区"}, str(out), encoding="cp932")
    assert res["proof_ok"] and res["read_back_ok"] and res["encoding"]["chosen_by"] == "agent", res
    assert out.read_bytes() == "氏名,住所\n山田太郎,東京都中央区\n".encode("cp932")
    # GBK を選べば GBK で読み、書いた所だけを GBK のバイトにする（ほかのバイトは元のまま）
    out2 = tmp_path / "c.csv"
    res = api.fill(str(p), {"R2.C2": "东京"}, str(out2), encoding="gbk")
    assert res["proof_ok"] and res["read_back_ok"], res
    assert out2.read_bytes() == JA.encode("cp932")[:-1] + "东京".encode("gbk") + b"\n"
    au = api.audit(str(p), encoding="cp932")
    assert au["ok"] and au["grade"]["grade"] == "S"
    # 段階にエージェントの判断を合わせても、選んだ文字コードの理由は残る
    g2 = api.grade(str(p), [{"check": "ui_only", "detail": "電子申請の画面から"}], encoding="cp932")
    assert any(r["by"] == "agent_choice" for r in g2["reasons"])


def test_agent_chosen_encoding_for_txt_and_json(tmp_path):
    t = tmp_path / "a.txt"
    t.write_bytes("申請書\n氏名：山田太郎\n".encode("cp932"))
    assert api.grade(str(t))["grade"] == "B"
    assert api.grade(str(t), encoding="cp932")["grade"] == "A"
    j = tmp_path / "a.json"
    j.write_bytes('{"氏名": "山田太郎", "住所": ""}'.encode("cp932"))
    assert api.grade(str(j))["grade"] == "B"  # JSON も、文字コードを決められなければ B
    assert api.grade(str(j), encoding="cp932")["grade"] == "S"
    out = tmp_path / "b.json"
    res = api.fill(str(j), {"/住所": "東京都"}, str(out), encoding="cp932")
    assert res["proof_ok"] and res["read_back_ok"], res
    assert out.read_bytes() == '{"氏名": "山田太郎", "住所": "東京都"}'.encode("cp932")


def test_chosen_encoding_must_decode_and_is_only_for_text_files(tmp_path):
    p = tmp_path / "a.csv"
    p.write_bytes(JA.encode("cp932"))
    with pytest.raises(api.ApiError, match="厳密に読めません"):
        api.grade(str(p), encoding="utf-8")
    with pytest.raises(api.ApiError, match="知りません"):
        api.window(str(p), encoding="no-such-code")
    r = tmp_path / "a.rtf"
    r.write_bytes(b"{\\rtf1\\ansi hello\\par}")
    with pytest.raises(api.ApiError, match="文字だけのファイル"):
        api.grade(str(r), encoding="cp932")
    with pytest.raises(api.ApiError, match="向き"):
        api.window(str(p), encoding="utf-16")  # BOM がなく、向きの決まらない名前


def test_mcp_tools_take_encoding(tmp_path):
    from todoku import mcp_server

    names = {t["name"] for t in mcp_server.TOOLS if "encoding" in t["inputSchema"]["properties"]}
    assert names == {"open_document", "read_window", "list_fields", "fill_document", "audit_document", "grade_document",
                     "read_table", "explain_calculation", "try_calculation", "extract_processing"}  # 後ろの 4 つは、8 ビットの字の表のため
    p = tmp_path / "a.csv"
    p.write_bytes(JA.encode("cp932"))
    got = mcp_server.call_tool("grade_document", {"path": str(p), "encoding": "cp932"})
    assert got["structuredContent"]["grade"] == "S"
    assert "エージェントの選択を規則で確かめた" in got["content"][0]["text"]
    got = mcp_server.call_tool("open_document", {"path": str(p), "encoding": "gbk"})
    assert "encoding_candidates" not in got["structuredContent"]["format"]
    res = mcp_server.call_tool("fill_document", {"path": str(p), "out_path": str(tmp_path / "b.csv"),
                                                 "values": {"R2.C2": "東京"}, "encoding": "cp932"})
    assert not res["isError"] and res["structuredContent"]["encoding"]["encoding"] == "cp932"


def test_declared_charset_decides_before_ambiguity(tmp_path):
    body = "申請書\n氏名：山田太郎\n"
    for head, enc in ((b'<meta charset="Shift_JIS">\n', "cp932"),
                      (b'<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">\n', "cp932"),
                      (b'<?xml version="1.0" encoding="EUC-JP"?>\n', "euc_jp")):
        blob = head + body.encode(enc)
        d = decode(blob)
        assert (d.encoding, d.basis, d.candidates) == (enc, "declared", []), head
        p = tmp_path / "d.txt"
        p.write_bytes(blob)
        g = api.grade(str(p))
        assert not any("文字コード" in r["title"] for r in g["reasons"]), g["reasons"]
        _k, doc, _p = api.window(str(p))
        assert doc["format"]["encoding_basis"] == "declared" and doc["format"]["encoding_declared_by"]
    assert declared(b"{\\rtf1\\ansi\\ansicpg932\\deff0 x}") == ("cp932", "RTF の \\ansicpg")
    # 宣言と中身が合わない（宣言の文字コードでは読めない）：宣言は使わず、候補を示す
    d = decode(b'<meta charset="EUC-JP">\n' + body.encode("cp932"))
    assert d.basis == "order" and d.encoding == "cp932" and d.candidates
    # BOM は宣言より先（UTF-8 の BOM と、食い違う宣言）
    d = decode(b"\xef\xbb\xbf" + b'<meta charset="Shift_JIS">' + body.encode("utf-8"))
    assert (d.encoding, d.basis) == ("utf-8", "bom")


# ====================================================================== 窓の読み落とし
def test_excel_markdown_omission_makes_read_c(monkeypatch):
    from test_example import make_xlsx

    from todoku.xlsx import audit as xa

    blob = make_xlsx({"申請": ({"A1": "氏名", "B2": "山田太郎"}, [])})
    base = grade_blob(blob, "a.xlsx")
    assert base.grade in ("A", "B") and not any(r.title == "窓の読み落とし" for r in base.reasons)
    real = xa.check_markdown

    def drop_row_2(doc, links=None):  # Markdown の窓が、値のある2行目を落とした
        for sh in doc["sheets"]:
            sh["reading"] = [ln for ln in sh["reading"] if ln.get("row") != 2]
        return real(doc, links)

    monkeypatch.setattr(xa, "check_markdown", drop_row_2)
    g = grade_blob(blob, "a.xlsx")
    assert g.grade == "C"
    r = next(r for r in g.reasons if r.title == "窓の読み落とし")
    assert any(w.startswith("申請!B2") for w in r.where), r.where


def test_word_omission_makes_read_c(monkeypatch):
    from test_docx import make_docx, p

    from todoku.docx import audit as da

    blob = make_docx(p("申請書") + p("氏名を書いてください"))
    assert grade_blob(blob, "a.docx").grade == "A"
    real = da.raw_texts
    monkeypatch.setattr(da, "raw_texts", lambda doc: real(doc) + ["欠"])  # 窓に出ていない字がある
    g = grade_blob(blob, "a.docx")
    assert g.grade == "C"
    r = next(r for r in g.reasons if r.title == "窓の読み落とし")
    assert any("「欠」" in w for w in r.where)


def test_pdf_omission_makes_read_c(monkeypatch):
    from test_pdf_alt import IMG, make

    from todoku.pdf import audit as pa

    blob = make(IMG)
    assert not any(r.title == "窓の読み落とし" for r in grade_blob(blob, "t.pdf").reasons)
    monkeypatch.setattr(pa, "raw_chars", lambda doc, page: ["欠"])
    g = grade_blob(blob, "t.pdf")
    r = next(r for r in g.reasons if r.title == "窓の読み落とし")
    assert r.grade == "C" and any("ページ 1" in w and "「欠」" in w for w in r.where)


# ====================================================================== 中身が画像の文書
_DRAWING = ('<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:docPr id="{n}" name="図 {n}">{deco}</wp:docPr>'
            '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
_DECO = ('<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative '
         'xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>')
_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_STYLES = (f'<?xml version="1.0" encoding="UTF-8"?><w:styles {_W}>'
           '<w:style w:type="paragraph" w:styleId="a1"><w:name w:val="Title"/></w:style>'
           '<w:style w:type="paragraph" w:styleId="a2"><w:name w:val="caption"/></w:style>'
           '<w:style w:type="paragraph" w:styleId="a3"><w:name w:val="見出しの型"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
           '<w:style w:type="paragraph" w:styleId="a4"><w:name w:val="本文の型"/><w:basedOn w:val="a3"/>'
           '<w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>'
           '</w:styles>')


def _pic(n: int, deco: bool = False) -> str:
    return _DRAWING.replace("{n}", str(n)).replace("{deco}", _DECO if deco else "")


def _sp(style: str, text: str) -> str:
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def _docx(body: str, header: str = "") -> bytes:
    """test_docx の文書に、スタイル・画像の部品（と、あればヘッダー）を足す。"""
    from test_docx import make_docx

    src = zipfile.ZipFile(io.BytesIO(make_docx(body)))
    buf = io.BytesIO()
    rels = '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
    if header:
        rels += '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            z.writestr(n, src.read(n))
        z.writestr("word/styles.xml", _STYLES)
        z.writestr("word/media/image1.png", b"\x89PNG\r\n\x1a\n")
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + '</Relationships>')
        if header:
            z.writestr("word/header1.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:hdr {_W}>{header}</w:hdr>')
            z.writestr("word/_rels/header1.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                       'Target="media/image1.png"/></Relationships>')
    return buf.getvalue()


def test_word_document_of_a_title_and_images_is_e():
    """1行の表題と、説明のない画像だけ（中身は画像）：読む E。画像の所と部品の名前を where に。"""
    blob = _docx(_sp("a1", "申請の手引き") + "".join(_pic(n) for n in range(1, 17)) + _sp("a2", "図1 記入例") + _sp("a3", "第1節"))
    g = grade_blob(blob, "a.docx")
    assert g.grade == "E", g.reasons
    r = next(r for r in g.reasons if r.grade == "E")
    assert r.title == "中身が画像の文書" and "16 個" in r.detail
    assert r.where[0].startswith("図1（") and "word/media/image1.png" in r.where[0]
    assert [a.title for a in g.alerts] == ["説明のない画像"]  # 注意（alerts）も別に残す


def test_word_body_text_and_images_is_alert_only():
    # 見出しの型を基にした本文の型（outlineLvl 9）：本文の段落として数える
    blob = _docx(_sp("a1", "申請の手引き") + _sp("a4", "下の図のとおりに記入してください。") + _pic(1) + _pic(2))
    g = grade_blob(blob, "a.docx")
    # 本文に字がある：説明のない画像は段階を変えず、アクセシビリティの注意に位置を示す
    assert g.grade == "A", g.reasons
    assert [a.title for a in g.alerts] == ["説明のない画像"] and len(g.alerts[0].where) == 2


def test_word_header_or_decorative_images_do_not_lower_the_read_grade():
    body = _sp("a1", "申請書") + "<w:p><w:r><w:t>氏名を書いてください。</w:t></w:r></w:p>"
    g = grade_blob(_docx(body + _pic(1, deco=True)), "a.docx")  # 飾りの印のある画像
    assert g.grade == "A" and g.alerts == [], g.reasons
    g = grade_blob(_docx(body, header=_pic(1)), "a.docx")  # ヘッダーの画像（本文の中身ではない）
    assert g.grade == "A", g.reasons
    assert [a.title for a in g.alerts] == ["説明のない画像"]  # 注意は示す（段階は変えない）
    g = grade_blob(_docx(body + _pic(1).replace('name="図 1">', 'name="図 1" descr="記入例の図">')), "a.docx")  # 説明がある
    assert g.grade == "A" and g.alerts == [], g.reasons


def test_pptx_slide_images():
    from pptx_maker import make, para, pic, sp

    only = make([sp(para("事業の概要"), "タイトル", ph="title") + pic(None, 9) + pic(None, 10)])
    g = grade_blob(only, "a.pptx")
    assert g.grade == "E" and any(r.title == "中身が画像の文書" for r in g.reasons), g.reasons
    some = make([sp(para("事業の概要"), "タイトル", ph="title") + sp(para("対象は区内の事業所です"), "本文", sid=3) + pic(None, 9)])
    g = grade_blob(some, "a.pptx")
    assert g.grade == "A" and not any(r.title == "説明のない本文の画像" for r in g.reasons), g.reasons
    ok = make([sp(para("事業の概要"), "タイトル", ph="title") + pic("事業所の地図", 9) + pic(None, 10, decorative=True)])
    assert grade_blob(ok, "a.pptx").grade == "A"


def test_pdf_figure_without_alt_on_a_text_page_is_alert_only(monkeypatch):
    from test_pdf_alt import IMG, make

    from todoku.pdf import window as pw

    real = pw.build

    def with_text(doc):  # 字のあるページとして確かめる（フォントを組み立てずに、窓の image_only だけを外す）
        win = real(doc)
        for pg in win["pages"]:
            pg["image_only"] = False
        return win

    monkeypatch.setattr(pw, "build", with_text)
    g = grade_blob(make(b"/Zu <</MCID 0>> BDC " + IMG + b" EMC", None, tagged=True), "t.pdf")
    # 字のあるページの、Alt のない図：段階は変えず、注意に位置（と大きさ）を示す
    assert not any(r.title.startswith("説明のない図") for r in g.reasons), g.reasons
    assert [a.title for a in g.alerts] == ["説明のない画像"] and g.alerts[0].where[0].startswith("p1.i1（"), g.alerts
    g = grade_blob(make(IMG), "t.pdf")  # タグのない PDF：中身の図か飾りかが作りから決まらない。注意だけ
    assert not any(r.title.startswith("説明のない図") for r in g.reasons)
    assert [a.title for a in g.alerts] == ["説明のない画像"]
