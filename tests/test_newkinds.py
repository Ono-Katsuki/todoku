"""文字だけのファイル（.txt）・CSV の字のない行・PowerPoint（.pptx）：窓・入力欄・書き込み（証明と読み戻し）・検査・段階。"""

from __future__ import annotations

import io
import zipfile

import pytest

from pptx_maker import group, make, para, pic, sp, table
from todoku import api
from todoku.book import kind_of

FORM = "様式第１号\r\n\r\n　令和　　年　　月　　日\r\n\r\n申請者　住所\r\n\t氏名　　　　　　　\r\n□ 同意する\r\n電話番号：\r\n以上"


# ====================================================================== .txt
def test_txt_kind_by_name_and_window_keeps_every_line(tmp_path):
    assert kind_of(b"hello", "a.txt") == "txt" and kind_of(b"hello", "a.dat") == "unknown"
    p = tmp_path / "a.txt"
    p.write_bytes(FORM.encode("cp932"))
    k, doc, parts = api.window(str(p))
    md = parts[0][1]
    assert k == "txt" and doc["format"]["encoding"] == "cp932" and doc["format"]["newline"] == "CRLF"
    # 「令和　年　月　日」・「氏名　　　」の空きは、字の並びからの推定なので「？」（規則では決めず、エージェントが判断する）
    assert "3｜　令和［P3.1？］年［P3.2？］月［P3.3？］日" in md  # 行の頭の全角の空白もそのまま
    assert "2｜" in md.split("\n") and "6｜\t氏名［P6.1？］" in md  # 空の行・行の頭のタブ
    ids = {f["id"]: f for f in api.fields(str(p))}
    assert {"P3.1", "P3.2", "P3.3", "P6.1", "P7.1"} <= set(ids)
    assert ids["P7.1"]["kind"] == "checkbox" and ids["P8.a1"].get("agent")  # 「：」の後ろは「？」の所


@pytest.mark.parametrize("enc", ["cp932", "utf-8", "utf-8-sig", "utf-16", "euc_jp"])
def test_txt_fill_changes_only_written_bytes(tmp_path, enc):
    p = tmp_path / "a.txt"
    raw = FORM.encode(enc)
    p.write_bytes(raw)
    out = tmp_path / "b.txt"
    r = api.fill(str(p), {"P3.1": "8", "P3.2": "4", "P6.1": "山田　太郎", "P7.1": True,
                          "@1": {"old": "以上", "new": "以上です"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    want = FORM.replace("令和　　年　　月", "令和8年4月").replace("氏名　　　　　　　", "氏名山田　太郎").replace("□", "■") \
        .replace("以上", "以上です")
    assert out.read_bytes() == want.encode(enc)  # 改行（CRLF）・BOM・文字コードはそのまま
    # P3.3（「日」の前の空き）は規則では決めない所（？）なので、まだ空いている欄の一覧ではなく、窓の「？」の所に残る（元の id のまま）
    assert "P3.3" not in {d["id"] for d in r["still_empty"]}
    assert {f["id"]: f for f in api.fields(str(out))}["P3.1"].get("agent")
    assert "令和［8］年" in r["say_to_user"]


def test_txt_checked_marks_can_be_unchecked(tmp_path):
    # ☑・■ はチェック済みの候補（？）。外すと印のない箱（☐・□）にし、証明と読み戻しが通る
    p = tmp_path / "c.txt"
    p.write_bytes("☑ 同意する\n■ 見出し\n".encode())
    ids = {f["id"]: f for f in api.fields(str(p))}
    assert ids["P1.1"]["value"] is True and ids["P1.1"].get("agent") and ids["P2.1"].get("agent")
    out = tmp_path / "d.txt"
    r = api.fill(str(p), {"P1.1": False, "P2.1": True}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert out.read_text() == "☐ 同意する\n■ 見出し\n"


def test_txt_proof_catches_extra_change(tmp_path):
    from todoku.txt.fill import prove

    raw = FORM.encode("utf-8")
    bad = raw.replace("以上".encode(), "以下".encode())  # 書いていない行の字
    pr = prove(raw, bad.replace("令和　　年".encode(), "令和8年".encode()), {"P3.1": "8"})
    assert not pr.ok and pr.unexpected
    lf = FORM.replace("\r\n", "\n").encode("utf-8")
    pr = prove(raw, lf, {})
    assert not pr.ok  # 改行が変わった


def test_txt_rejects_newline_and_unencodable(tmp_path):
    p = tmp_path / "a.txt"
    p.write_bytes(FORM.encode("cp932"))
    with pytest.raises(api.ApiError, match="改行"):
        api.fill(str(p), {"P3.1": "8\n9"}, str(tmp_path / "b.txt"))
    with pytest.raises(api.ApiError, match="書けない字"):
        api.fill(str(p), {"P3.1": "😀"}, str(tmp_path / "b.txt"))


def test_txt_audit_and_grade(tmp_path):
    p = tmp_path / "a.txt"
    p.write_bytes(FORM.encode("cp932"))
    au = api.audit(str(p))
    assert au["ok"], [c for c in au["checks"] if not c["ok"]]
    names = {c["name"] for c in au["checks"]}
    assert {"読み落とし（行の字）", "全入力欄への書き込み", "読み戻し", "字を目印にした書き込み"} <= names
    g = api.grade(str(p))
    # 推定の欄・「？」の所は書く段階の B。この字の並びは Shift_JIS としても GBK としても仕様どおりに読めるので、
    # 文字コードは規則では決められない（B：中身は窓にすべて出ていて、候補を窓に示し、エージェントが字を読んで決める）
    assert g["grade"] == "B" and g["write"]["grade"] == "B"
    assert [r["title"] for r in g["reasons"] if r["grade"] == "B"] == ["文字コードを規則では決められない"]
    # エージェントが選んだ文字コードを渡せば、その文字コードで読み直し、段階は下がらない（選んだことは理由に残す）
    g = api.grade(str(p), encoding="cp932")
    assert g["grade"] == "A", g
    assert [(r["title"], r["by"]) for r in g["reasons"] if r["title"].startswith("エージェントが選んだ")] == \
        [("エージェントが選んだ文字コード", "agent_choice")]
    plain = tmp_path / "b.txt"
    plain.write_bytes("ただの文です。\n二行目です。\n".encode())
    g = api.grade(str(plain))
    assert g["grade"] == "A", g
    binary = tmp_path / "c.txt"
    binary.write_bytes(b"\x00\x01\x02abc")
    assert api.grade(str(binary))["grade"] == "F"


def test_txt_reading_check_catches_omission(monkeypatch):
    from todoku.txt import audit as ta
    from todoku.txt import window as tw

    blob = FORM.encode("utf-8")
    real = tw.build

    def drop_last_line(doc):
        win = real(doc)
        win["lines"] = win["lines"][:-1]
        return win

    monkeypatch.setattr(ta, "build", drop_last_line)
    au = ta.run(blob, "a.txt")
    c = next(c for c in au.checks if c.name == "読み落とし（行の字）")
    assert not c.ok


def test_txt_big_window_is_split_at_line_breaks():
    from todoku.txt.document import TxtDocument
    from todoku.txt.window import build, markdown_parts

    text = "".join(f"{i}番目の項目　氏名：\n" for i in range(400))
    parts = markdown_parts(build(TxtDocument(text.encode(), "a.txt")), budget=6000)
    assert len(parts) > 3 and all(len(t) <= 6000 for _n, t in parts[1:])
    joined = "\n".join(t for _n, t in parts[1:])
    assert all(f"{i + 1}｜{i}番目の項目" in joined for i in range(400))


def test_txt_inside_zip(tmp_path):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("readme.txt", FORM.encode("cp932"))
        z.writestr("other.csv", "a\n1\n")
    src = tmp_path / "a.zip"
    src.write_bytes(buf.getvalue())
    k, _doc, parts = api.window(f"{src}!/readme.txt")
    assert k == "txt" and "令和［P3.1？］年" in parts[0][1]
    out = tmp_path / "b.zip"
    r = api.fill(f"{src}!/readme.txt", {"P3.1": "8"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    assert zipfile.ZipFile(out).read("readme.txt") == FORM.replace("令和　　年", "令和8年").encode("cp932")
    assert zipfile.ZipFile(out).read("other.csv") == b"a\n1\n"
    au = api.audit(str(src))
    assert any(c["name"].startswith("readme.txt：") for c in au["checks"]) and au["ok"]


# ====================================================================== CSV の字のない行
@pytest.mark.parametrize("name,blob", [("a.tsv", b"id\tname\n1\ta\n\n"), ("a.csv", b"id,name\r\n1,a\r\n\r\n"),
                                       ("m.csv", b"id,name\n\n1,a\n")])
def test_csv_empty_line_is_one_empty_field(tmp_path, name, blob):
    """RFC 4180 の文法では、字のない行は「空のマスが1つの行」：窓と、別の読み方（csv モジュール）が同じに読む。"""
    p = tmp_path / name
    p.write_bytes(blob)
    au = api.audit(str(p))
    assert au["ok"], [c for c in au["checks"] if not c["ok"]]
    _k, doc, parts = api.window(str(p))
    r = doc["format"]["empty_lines"][0]
    assert doc["rows"][r - 1] == [""] and "字のない行" in parts[0][1]
    d = "\t" if name.endswith(".tsv") else ","
    nl = "\r\n" if b"\r\n" in blob else "\n"
    n = doc["row_count"]
    out = tmp_path / ("o_" + name)
    res = api.fill(str(p), {f"R{r}.C2": "b"}, str(out))  # 字のない行の右に書く：区切りを足す
    assert res["proof_ok"] and res["read_back_ok"]
    lines = out.read_bytes().decode().split(nl)
    assert lines[r - 1] == f"{d}b"
    res = api.fill(str(p), {f"R{n + 1}.C1": "3"}, str(out), overwrite=True)  # 字のない行の後ろに行を足す（同じ名前に書き直す）
    assert res["proof_ok"] and res["read_back_ok"]
    assert out.read_bytes().startswith(blob if blob.endswith(nl.encode()) else blob + nl.encode())


def test_csv_empty_line_check_still_catches_omission():
    from todoku.datafile.audit import csv_reading, independent_rows
    from todoku.datafile.csvfile import CsvFile
    from todoku.datafile.window import build_csv, parts_csv

    assert independent_rows('a\n""\n \n\n', ",") == [["a"], [""], [" "], [""]]
    f = CsvFile(b"id,name\n1,a\n\n", "a.csv")
    doc = build_csv(f)
    md = "\n".join(t for _n, t in parts_csv(doc))
    assert csv_reading(f, doc, md).ok
    f.records[2].fields[0].value = "x"  # 窓の読み方が、字のない行を読み違えたとき
    assert not csv_reading(f, doc, md).ok
    f2 = CsvFile(b"id,name\n1,a\n\n", "a.csv")
    f2.records[2].b = f2.records[2].a + 1  # 字のない行を、字のある行として読んだとき
    assert not csv_reading(f2, doc, md).ok


# ====================================================================== PowerPoint
def _deck() -> bytes:
    return make([
        sp(para("申請書"), "タイトル", "title") + sp(para("令和　　年　　月　　日") + para("□ 同意する"))
        + table([["氏名", ""], ["住所", " "]]) + pic("会社のロゴ") + pic(None, 10),
        group(sp(para("グループの中の字 & <記号>"), sid=30), sp(para("A", "B"), sid=31)) + pic(None, 11, decorative=True),
    ], notes={1: sp(para("ノートの字です"), "ノート", "body")}, master_text="株式会社テスト", hidden={2})


def test_pptx_window_reads_slides_tables_groups_notes_and_template(tmp_path):
    p = tmp_path / "a.pptx"
    p.write_bytes(_deck())
    k, doc, parts = api.window(str(p))
    md = parts[0][1]
    assert k == "pptx" and api.summary(str(p))["slides"] == 2
    # 「令和　年　月　日」の空きは字の並びからの推定（？）。マスの見出し「氏名」の後ろも、右のマスの代わりの所として示す
    for s in ("## スライド 1", "令和［P2.1？］年［P2.2？］月［P2.3？］日", "氏名［P4.a1？］ ｜ ［T1.R1.C2］", "ノートの字です",
              "スライド 2（非表示のスライド）", "グループの中の字 & <記号>", "株式会社テスト", "〔図: 会社のロゴ〕", "〔図2（説明なし）〕"):
        assert s in md, s
    assert "マスターのタイトルの書式" not in md and "タイトルを入力" not in md  # 型のプレースホルダーの見本の字は出さない
    ids = {f["id"] for f in api.fields(str(p))}
    assert {"T1.R1.C2", "T1.R2.C2", "P2.1", "P3.1"} <= ids
    g = api.grade(str(p))
    assert g["write"]["grade"] == "B"
    alerts = g.get("alerts") or []
    assert len(alerts) == 1 and "図2" in alerts[0]["where"][0]  # 飾りの印の絵は除く


def test_pptx_fill_proof_and_read_back(tmp_path):
    p = tmp_path / "a.pptx"
    p.write_bytes(_deck())
    out = tmp_path / "b.pptx"
    r = api.fill(str(p), {"T1.R1.C2": "山田 & <太郎>", "T1.R2.C2": "東京都", "P2.1": "8", "P3.1": True,
                          "@1": {"old": "グループの中の字", "new": "グループの中の字（改）"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    z0, z1 = zipfile.ZipFile(p), zipfile.ZipFile(out)
    changed = [n for n in z0.namelist() if z0.read(n) != z1.read(n)]
    assert sorted(changed) == ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]
    x = z1.read("ppt/slides/slide1.xml").decode()
    assert '<a:r><a:rPr lang="ja-JP" sz="1400"/><a:t>山田 &amp; &lt;太郎&gt;</a:t></a:r><a:endParaRPr lang="ja-JP" sz="1400"/>' in x
    md = api.window(str(out))[2][0][1]
    # 書いた後の ■ は、チェック済みの印らしい字として「？」で示す（値は True）。残りの「月」「日」の空きも「？」の所
    assert "令和8年［P2.1？］月［P2.2？］日" in md and "■［P3.1？］ 同意する" in md and "グループの中の字（改）" in md
    after = {f["id"]: f for f in api.fields(str(out))}
    assert after["P3.1"]["value"] is True and after["P2.2"].get("agent")


def test_pptx_proof_catches_change_outside_written_paragraph():
    from todoku.pptx.fill import prove
    from todoku.zipraw import RawZip

    blob = _deck()
    z = RawZip(blob)
    z.replace("ppt/slides/slide1.xml", z.read("ppt/slides/slide1.xml").replace("申請書".encode(), "届出書".encode()))
    pr = prove(blob, z.build(), {})
    assert not pr.ok
    z = RawZip(blob)
    z.replace("ppt/slideMasters/slideMaster1.xml", z.read("ppt/slideMasters/slideMaster1.xml") + b" ")
    assert not prove(blob, z.build(), {}).ok


def test_pptx_audit_and_reading_check_catches_omission(tmp_path, monkeypatch):
    p = tmp_path / "a.pptx"
    p.write_bytes(_deck())
    au = api.audit(str(p))
    assert au["ok"], [c for c in au["checks"] if not c["ok"]]
    from todoku.pptx import audit as pa
    from todoku.pptx import document as pd

    real = pd.PptxDocument._para

    def drop_second_run(self, node, part, cell):
        out = real(self, node, part, cell)
        out.segs = out.segs[:1]
        return out

    monkeypatch.setattr(pd.PptxDocument, "_para", drop_second_run)
    c = next(c for c in pa.run(_deck(), "a.pptx").checks if c.name == "読み落とし（スライドの字）")
    assert not c.ok and any("「B」" in x for x in c.items)


def test_pptx_refuses_writing_over_slide_number_and_newline(tmp_path):
    fld = ('<a:p><a:r><a:rPr lang="ja-JP"/><a:t>ページ　</a:t></a:r><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
           '<a:rPr lang="ja-JP"/><a:t>1</a:t></a:fld><a:endParaRPr lang="ja-JP"/></a:p>')
    p = tmp_path / "a.pptx"
    p.write_bytes(make([sp(fld) + sp(para("令和　　年"))]))
    with pytest.raises(api.ApiError):
        api.fill(str(p), {"@1": {"old": "1", "new": "2"}}, str(tmp_path / "b.pptx"))
    with pytest.raises(api.ApiError, match="改行"):
        api.fill(str(p), {"P2.1": "8\n9"}, str(tmp_path / "b.pptx"))
    au = api.audit(str(p))
    assert au["ok"], [c for c in au["checks"] if not c["ok"]]
