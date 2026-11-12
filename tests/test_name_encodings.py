"""文字コードの印のない 8 ビットの字（書庫の名前・メールの字・OLE の表示名・絵の字）：日本語と置かずに、規則で読む（namecode）。

- 同じファイルの宣言（メールのコードページ・複合ファイルの要約の情報のコードページ・字体の文字セット）があれば、それ
- UTF-8 の形なら UTF-8
- ほかは、バイトからは決まらない：決まった順のいちばん前で示し、ほかの読み方を候補として並べ、決められないと伝える
- エージェントが選んだ文字コード（encoding）は、厳密に読めるときだけ使う。正しい読み方の名前でそのまま指せる
"""

from __future__ import annotations

import gzip
import io
import struct
import tarfile
import zipfile

import pytest

from test_archives import _lzh_level0, _p
from test_embeds import compobj_stream, make_cfb, package_native
from todoku import api, namecode, ziparc
from todoku.cfb import CFB, SubStorage
from todoku.embeds.ole import compobj, describe, ole10native, summary_codepage

CSV = b"a,b\n1,2\n"


def _zip(names, enc: str, body: bytes = CSV) -> bytes:
    """UTF-8 の印のない、enc の名前の ZIP（zipfile は ASCII でない名前に UTF-8 の印を付けるので、同じ長さの ASCII の名前で作って置き換える）。"""
    buf, swaps = io.BytesIO(), []
    with zipfile.ZipFile(buf, "w") as z:
        for i, n in enumerate(names):
            raw = n.encode(enc) if isinstance(n, str) else n
            ph = (f"@{i:03d}" + "_" * len(raw))[:len(raw)].encode()
            swaps.append((ph, raw))
            z.writestr(ph.decode(), body)
    out = buf.getvalue()
    for ph, raw in swaps:
        assert out.count(ph) == 2
        out = out.replace(ph, raw)
    return out


def _text(path, **kw):
    k, doc, parts = api.window(path, **kw)
    return doc, "\n".join(x for _n, x in parts)


# ====================================================================== 決め方（namecode）
def test_codepage_numbers_map_to_codecs_and_unknown_numbers_to_none():
    c = namecode.codec_of_codepage
    assert (c(932), c(949), c(1251), c(28592), c(50220), c(10001), c(65001)) == (
        "cp932", "cp949", "cp1251", "iso8859_2", "iso2022_jp_ms", "shift_jis", "utf-8")
    # Windows のコードページ 50220〜50222 は、2 バイトの字を CP932 の表で読む ISO-2022-JP（①・髙 がある）
    assert (c(50221), c(50222)) == ("iso2022_jp_ms", "iso2022_jp_ms")
    assert b"\x1b$B-!|b\x1b(B".decode(c(50221)) == "①髙"
    assert c(99999) is None and c(0) is None


def test_ascii_and_utf8_shaped_strings_are_decided_and_the_rest_is_not():
    r = namecode.decide([b"plain.txt", "申請書.csv".encode()])
    assert r.basis == "" and not r.undetermined and r.info() == {} and r.note_lines() == []
    assert r.text(b"plain.txt") == ("plain.txt", "ASCII")
    assert r.text("申請書.csv".encode()) == ("申請書.csv", "UTF-8（バイトの形）")
    for enc, name in (("cp932", "申請書.csv"), ("cp949", "신청서.csv"), ("gbk", "申请书.csv"), ("cp950", "申請書.csv"),
                      ("cp850", "Müller.csv"), ("cp866", "заявление.csv"), ("cp1251", "заявление.csv"), ("cp874", "คำร้อง.csv"),
                      ("cp1253", "αίτηση.csv"), ("cp1256", "طلب.csv")):
        raw = name.encode(enc)
        r = namecode.decide([raw])
        assert r.basis == "order" and r.undetermined, enc
        shown = [r.own_sample] + [c["sample"] for c in r.candidates]
        assert name in shown, (enc, shown)  # 正しい読み方は、いまの読み方か、候補のどれかに必ずある
        assert name in namecode.readings(raw)


def test_declared_code_page_wins_only_when_every_string_reads_strictly():
    ru = "Документ".encode("cp1251")
    r = namecode.decide([ru], declared=[("cp1251", "宣言")])
    assert r.basis == "declared" and r.text(ru) == ("Документ", "Windows-1251（キリル）（宣言）") and not r.undetermined
    # 宣言の文字コードで読めない字があるときは、宣言を使わない（次の宣言 → 決められない）
    ja = "申請書".encode("cp932")
    assert namecode.decide([ja], declared=[("ascii", "a"), ("cp932", "b")]).declared_by == "b"
    assert namecode.decide([b"\x81\x40\x81"], declared=[("cp932", "b")]).basis != "declared"


def test_agent_choice_is_used_only_when_it_reads_strictly():
    ko = "신청서.csv".encode("cp949")
    r = namecode.decide([ko, b"a.txt"], encoding="euc-kr")
    assert r.basis == "given" and r.encoding == "cp949" and r.text(ko)[0] == "신청서.csv" and r.info()["chosen_by"] == "agent"
    with pytest.raises(namecode.NameCodeError, match="知りません"):
        namecode.decide([ko], encoding="no-such-code")
    with pytest.raises(namecode.NameCodeError, match="厳密に読めません"):
        namecode.decide([ko], encoding="utf-8")
    with pytest.raises(namecode.NameCodeError, match="選べる字がありません"):
        namecode.decide([b"a.txt"], encoding="cp949")


# ====================================================================== ZIP
@pytest.mark.parametrize("enc,names", [("cp932", ["申請書.csv", "統計/一覧.csv"]), ("cp949", ["신청서.csv", "제출서류.csv"]),
                                       ("gbk", ["申请书.csv", "统计.csv"]), ("cp850", ["Müller.csv", "Añadir.csv"]),
                                       ("cp866", ["заявление.csv", "отчёт.csv"])])
def test_zip_names_without_a_label_are_shown_with_every_reading(tmp_path, enc, names):
    p = _p(tmp_path, "a.zip", _zip(names, enc))
    doc, t = _text(p)
    ne = doc["name_encoding"]
    assert ne["basis"] == "order" and ne["count"] == 2
    assert "名前の文字コードを規則では決められません" in t and "（いまの読み方）" in t
    # 正しい読み方は、いまの読み方か候補のどれかにある（同じ字になる読み方は、初めの 1 つだけ並べる）
    assert names[0] in [ne["sample"]] + [c["sample"] for c in ne["candidates"]] and f"：{names[0]}" in t
    # 正しい読み方の名前で、encoding を渡さなくても指せる
    assert _text(f"{p}!/{names[1]}")[1].count("| 1 | 2 |") == 1
    # 選んだ文字コードを渡すと、一覧をその読み方で出し直す
    doc2, t2 = _text(p, encoding=enc)
    assert [m["name"] for m in doc2["members"]] == names and doc2["name_encoding"]["basis"] == "given"
    assert "エージェントが選んで指定した" in t2 and "決められません" not in t2
    s = api.summary(p, encoding=enc)
    assert [m["name"] for m in s["members"]] == names and s["name_encoding"]["chosen_by"] == "agent"


def test_zip_summary_and_grade_tell_that_names_are_undetermined(tmp_path):
    p = _p(tmp_path, "a.zip", _zip(["신청서.csv"], "cp949"))
    s = api.summary(p)
    assert s["name_encoding"]["basis"] == "order"
    assert {"encoding": "cp949", "sample": "신청서.csv"} in s["name_encoding"]["candidates"]
    g = api.grade(p)
    assert g["grade"] == "B" and g["reasons"][0]["title"] == "名前の文字コードを規則では決められない"
    assert "cp949" in g["reasons"][0]["where"]
    g2 = api.grade(p, encoding="cp949")  # エージェントが選んだ：段階を下げない
    assert g2["grade"] == "S" and any(r["title"] == "エージェントが選んだ文字コード" and r["by"] == "agent_choice" for r in g2["reasons"])


def test_zip_encoding_that_cannot_read_the_names_is_refused(tmp_path):
    p = _p(tmp_path, "a.zip", _zip(["申請書.csv"], "cp932"))
    with pytest.raises(api.ApiError, match="厳密に読めません"):
        api.window(p, encoding="utf-8")
    with pytest.raises(api.ApiError, match="知りません"):
        api.window(p, encoding="shift-jisx")
    q = _p(tmp_path, "b.zip", _zip(["plain.csv"], "ascii"))
    with pytest.raises(api.ApiError, match="選べる字がありません"):
        api.window(q, encoding="cp932")
    assert "name_encoding" not in api.summary(q)


def test_zip_utf8_shaped_names_without_the_flag_are_utf8(tmp_path):
    p = _p(tmp_path, "a.zip", _zip(["申請書.csv", "신청서.csv"], "utf-8"))
    doc, t = _text(p)
    assert [m["name"] for m in doc["members"]] == ["申請書.csv", "신청서.csv"]
    assert "name_encoding" not in doc and "決められません" not in t
    assert all("UTF-8（バイトの形）" in m["named"] for m in doc["members"])
    assert api.grade(p)["grade"] == "S"


def test_zip_name_that_matches_two_members_under_different_readings_is_refused():
    # 同じ字「重短辞.csv」に、EUC-JP で読める名前と、別のバイトで UTF-8 として読める名前
    a = "重短辞.csv".encode("euc_jp")
    blob = _zip([a, "신청서2.csv".encode("cp949")], "ascii")
    ms, reading = ziparc.listing(blob)
    assert reading.basis == "order"
    hit = [m for m in ms if m.rawname == a][0]
    assert ziparc.find_by_reading(ms, "신청서.csv", lambda m: m.rawname) is hit  # cp949 の読み方で 1 つに決まる
    # 一覧の名前そのもの（UTF-8 の形の名前）があれば、ほかの名前の別の読み方より先
    two = _zip([a, "신청서.csv".encode("utf-8")], "ascii", body=b"")
    two = two  # 中身は空：名前だけを見る
    ms2, _r = ziparc.listing(two)
    assert "신청서.csv" in [m.name for m in ms2]
    assert ziparc.find_by_reading([m for m in ms2 if m.name != "신청서.csv"], "신청서.csv", lambda m: m.rawname) is not None
    assert [m.rawname for m in ms2 if m.name == "신청서.csv"] == ["신청서.csv".encode("utf-8")]


# ====================================================================== tar・gzip・LZH
def _tar(files: dict, enc: str, fmt=tarfile.USTAR_FORMAT) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w", format=fmt, encoding=enc) as t:
        for n, b in files.items():
            ti = tarfile.TarInfo(n)
            ti.size = len(b)
            t.addfile(ti, io.BytesIO(b))
    return buf.getvalue()


def test_tar_names_in_a_legacy_encoding_and_pax_names(tmp_path):
    p = _p(tmp_path, "a.tar", _tar({"신청서.csv": CSV, "b.csv": CSV}, "cp949"))
    doc, t = _text(p)
    assert doc["name_encoding"]["basis"] == "order" and "  - cp949：신청서.csv" in t
    assert "| 1 | 2 |" in _text(f"{p}!/신청서.csv")[1]
    assert [m["name"] for m in _text(p, encoding="cp949")[0]["members"]] == ["신청서.csv", "b.csv"]
    # pax の path は UTF-8 と決まっている：決められない、とは言わない
    q = _p(tmp_path, "b.tar", _tar({"신청서.csv": CSV}, "utf-8", tarfile.PAX_FORMAT))
    doc, t = _text(q)
    assert [m["name"] for m in doc["members"]] == ["신청서.csv"] and "name_encoding" not in doc
    assert "pax" in doc["members"][0]["named"]
    with pytest.raises(api.ApiError, match="選べる字がありません"):
        api.window(q, encoding="cp949")


def test_writing_into_a_member_keeps_the_name_bytes(tmp_path):
    raw = "신청서.csv".encode("cp949")
    for name, blob, n in (("a.zip", _zip(["신청서.csv", "b.csv"], "cp949"), 2), ("a.tar", _tar({"신청서.csv": CSV}, "cp949"), 1)):
        p, out = _p(tmp_path, name, blob), str(tmp_path / ("out_" + name))
        r = api.fill(f"{p}!/신청서.csv", {"R2.C2": "4"}, out)  # 正しい読み方の名前で指して書く
        assert r["read_back_ok"] and open(out, "rb").read().count(raw) == n
        assert "| 1 | 4 |" in _text(f"{out}!/신청서.csv")[1]
        assert _text(out)[0]["name_encoding"]["basis"] == "order"  # 書いた後も、名前のバイトは同じ


def test_gzip_name_and_lzh_names(tmp_path):
    buf = io.BytesIO()
    with gzip.GzipFile(filename="统计.csv".encode("gbk").decode("latin-1"), mode="wb", fileobj=buf, mtime=7) as g:
        g.write(CSV)
    p = _p(tmp_path, "x.gz", buf.getvalue())
    doc, t = _text(p)
    assert doc["name_encoding"]["basis"] == "order" and "  - gbk：统计.csv" in t
    assert [m["name"] for m in _text(p, encoding="gbk")[0]["members"]] == ["统计.csv"]
    assert "| 1 | 2 |" in _text(f"{p}!/统计.csv")[1]
    lz = _p(tmp_path, "a.lzh", _lzh_level0("申请书.csv".encode("gbk"), CSV) + b"\x00")
    doc, t = _text(lz)
    assert doc["name_encoding"]["basis"] == "order" and "  - gbk：申请书.csv" in t
    assert [m["name"] for m in _text(lz, encoding="gbk")[0]["members"]] == ["申请书.csv"]
    assert "| 1 | 2 |" in _text(f"{lz}!/申请书.csv")[1]
    with pytest.raises(api.ApiError, match="にだけ指定できます"):
        api.window(_p(tmp_path, "y.bz2", __import__("bz2").compress(CSV)), encoding="gbk")


# ====================================================================== メール（.msg）
def _props(**cp) -> bytes:
    """__properties_version1.0：見出し 32 バイト ＋ コードページのプロパティ（PR_MESSAGE_CODEPAGE 0x3FFD・PR_INTERNET_CPID 0x3FDE）。"""
    out = bytes(32)
    for pid, v in (("message", 0x3FFD), ("internet", 0x3FDE)):
        if pid in cp:
            out += struct.pack("<HHII", 3, v, 6, cp[pid]) + bytes(4)
    return out


def _msg8(enc: str, subject: str, sender: str, body: str, att: str, **cp) -> bytes:
    e = lambda s: s.encode(enc)  # noqa: E731
    return make_cfb({"__properties_version1.0": _props(**cp),
                     "__substg1.0_0037001E": e(subject), "__substg1.0_0C1A001E": e(sender), "__substg1.0_1000001E": e(body),
                     "__attach_version1.0_#00000000": {"__substg1.0_3707001E": e(att), "__substg1.0_37010102": CSV}})


def test_msg_8bit_strings_follow_the_message_code_page(tmp_path):
    p = _p(tmp_path, "ru.msg", _msg8("cp1251", "Заявление", "Иванов", "Текст письма", "форма.csv", message=1251))
    doc, t = _text(p)
    assert "Заявление" in t and "Иванов" in t and "Текст письма" in t and "форма.csv" in t and "決められ" not in t
    assert doc["name_encoding"]["basis"] == "declared"
    assert "| 1 | 2 |" in _text(f"{p}!/форма.csv")[1]
    # PR_INTERNET_CPID が ISO-2022-JP（送るときの文字コード）でも、8 ビットの字は PR_MESSAGE_CODEPAGE で読む
    q = _p(tmp_path, "ja.msg", _msg8("cp932", "申請について", "山田", "本文の字", "様式.csv", message=932, internet=50220))
    doc, t = _text(q)
    assert "申請について" in t and "山田" in t and "本文の字" in t and "様式.csv" in t and "決められ" not in t
    # 7 ビットの字（ISO-2022-JP の印つき）は PR_INTERNET_CPID で読む：コードページ 50221 は、Windows の字（①・髙・～）も読む
    w = _p(tmp_path, "win.msg", _msg8("iso2022_jp_ms", "①髙橋の申請", "㈱髙橋", "本文～ｶﾅ", "様式①.csv", message=932, internet=50221))
    doc, t = _text(w)
    assert "①髙橋の申請" in t and "㈱髙橋" in t and "本文～ｶﾅ" in t and "様式①.csv" in t and "決められ" not in t


def test_msg_8bit_strings_without_a_code_page_are_undetermined(tmp_path):
    p = _p(tmp_path, "ko.msg", _msg8("cp949", "신청서 제출", "김철수", "본문입니다", "양식.csv"))
    doc, t = _text(p)
    assert doc["name_encoding"]["basis"] == "order" and "文字コードを規則では決められません" in t
    assert "  - cp949：" in t
    g = api.grade(p)
    assert g["grade"] == "B" and any(r["title"] == "字の文字コードを規則では決められない" and r["grade"] == "B" for r in g["reasons"])
    g2 = api.grade(p, encoding="cp949")  # 選んだ後は、文字コードの理由で下げない（メールの本文の B は、もとからの決まり）
    assert [r["by"] for r in g2["reasons"] if "文字コード" in r["title"]] == ["agent_choice"]
    doc2, t2 = _text(p, encoding="cp949")
    assert "신청서 제출" in t2 and "김철수" in t2 and "본문입니다" in t2 and "양식.csv" in t2 and "決められません" not in t2
    assert "| 1 | 2 |" in _text(f"{p}!/양식.csv")[1]  # 正しい読み方の名前で指せる
    with pytest.raises(api.ApiError, match="厳密に読めません"):
        api.window(p, encoding="utf-8")


def test_eml_does_not_take_an_encoding(tmp_path):
    from test_archives import _eml

    with pytest.raises(api.ApiError, match="にだけ指定できます"):
        api.window(_p(tmp_path, "a.eml", _eml()), encoding="cp932")


# ====================================================================== OLE の表示名・パッケージのファイル名
def _summary(cp: int) -> bytes:
    """\\x05SummaryInformation：コードページ（PIDSI_CODEPAGE、VT_I2）だけのプロパティの集まり（MS-OLEPS）。"""
    sect = struct.pack("<II", 24, 1) + struct.pack("<II", 1, 16) + struct.pack("<HHHH", 2, 0, cp, 0)
    return struct.pack("<HHI16sI", 0xFFFE, 0, 0x00020006, bytes(16), 1) + bytes(16) + struct.pack("<I", 48) + sect


def _compobj(user_type: bytes, prog: bytes = b"Excel.Sheet.8", unicode_type: str = "") -> bytes:
    lp = lambda b: struct.pack("<I", len(b) + 1) + b + b"\0"  # noqa: E731
    uni = (unicode_type + "\0").encode("utf-16-le") if unicode_type else b""
    return bytes(28) + lp(user_type) + struct.pack("<I", 0) + lp(prog) + struct.pack("<II", 0x71B239F4, len(uni) // 2) + uni


def test_summary_code_page_is_read_from_the_same_storage_then_the_outer_file():
    c = CFB(make_cfb({"\x05SummaryInformation": _summary(932), "ObjectPool": {
        "_1": {"\x01CompObj": _compobj(b"x"), "\x05SummaryInformation": _summary(1253)}, "_2": {"\x01CompObj": _compobj(b"x")}}}))
    assert summary_codepage(SubStorage(c, "ObjectPool/_1")) == [("cp1253", "同じ入れ物の要約の情報のコードページ"),
                                                              ("cp932", "外側のファイルの要約の情報のコードページ")]
    assert summary_codepage(SubStorage(c, "ObjectPool/_2")) == [("cp932", "外側のファイルの要約の情報のコードページ")]
    assert summary_codepage(CFB(make_cfb({"\x05SummaryInformation": _summary(1200)}))) == []  # Unicode：8 ビットの字の宣言ではない
    assert summary_codepage(CFB(make_cfb({"a": b"x"}))) == []


@pytest.mark.parametrize("cp,enc,name", [(1251, "cp1251", "Лист Microsoft Office Excel"), (1253, "cp1253", "Γράφημα του Microsoft Office Excel"),
                                         (932, "cp932", "Microsoft Excel ワークシート"), (949, "cp949", "Microsoft Excel 워크시트")])
def test_compobj_display_name_follows_the_declared_code_page(cp, enc, name):
    c = CFB(make_cfb({"\x05SummaryInformation": _summary(cp), "\x01CompObj": _compobj(name.encode(enc)), "Workbook": b"x" * 10}))
    info = compobj(c.read("\x01CompObj"), summary_codepage(c))
    assert info["user_type"] == name and "user_type_undetermined" not in info and "要約の情報のコードページ" in info["user_type_read"]
    d = describe(c)
    assert name in d["class"] and not any("決められません" in x for x in d["markdown"])


def test_compobj_unicode_display_name_wins_and_undeclared_names_are_told():
    raw = "Лист Microsoft Office Excel".encode("cp1251")
    assert compobj(_compobj(raw, unicode_type="Лист Microsoft Office Excel")) == {
        "user_type": "Лист Microsoft Office Excel", "prog": "Excel.Sheet.8"}
    info = compobj(_compobj(raw))
    assert info["user_type_undetermined"] and "決められない" in info["user_type_read"]
    assert compobj(_compobj(b"Microsoft Excel Worksheet")) == {"user_type": "Microsoft Excel Worksheet", "prog": "Excel.Sheet.8"}
    # 中身を読めない種類の部品：種類の名前の下に、決められないことを書く
    d = describe(CFB(make_cfb({"\x01CompObj": _compobj("Чертёж".encode("cp1251"), b"Some.Drawing.1"), "Contents": b"x" * 10})))
    assert d["kind"] == "other" and any("種類の名前の文字コードは" in x and "決められません" in x for x in d["markdown"])
    assert d["class_read"].endswith("（決められない）")


def test_package_file_name_follows_the_declared_code_page():
    native = package_native("a.txt", b"hello").replace(b"a.txt", "отчёт".encode("cp1251"))
    how: dict = {}
    assert ole10native(native, [("cp1251", "宣言")], how) == ("отчёт", b"hello") and how == {"file_read": "Windows-1251（キリル）（宣言）"}
    how = {}
    name, body = ole10native(native, (), how)
    assert body == b"hello" and how["file_undetermined"] and "отчёт" in namecode.readings("отчёт".encode("cp1251"))
    d = describe(CFB(make_cfb({"\x01Ole10Native": native, "\x01CompObj": compobj_stream("Package", "Package")})))
    assert any("ファイル名の文字コードは" in x for x in d["markdown"])
    d = describe(CFB(make_cfb({"\x01Ole10Native": native, "\x01CompObj": compobj_stream("Package", "Package"),
                               "\x05SummaryInformation": _summary(1251)})))
    assert d["file"] == "отчёт" and not any("決められません" in x for x in d["markdown"])


# ====================================================================== 絵（WMF・EMF）の 8 ビットの字・リンクの先・古い Word
def test_metafile_text_follows_the_font_charset():
    assert namecode.read_ansi("신청".encode("cp949"), 129) == ("신청", "")
    assert namecode.read_ansi("Заявление".encode("cp1251"), 204) == ("Заявление", "")
    assert namecode.read_ansi("申請".encode("cp932"), 128) == ("申請", "")
    assert namecode.read_ansi(b"abc", None) == ("abc", "")
    assert namecode.read_ansi(b"\xa7\xb7", 2) == ("§·", "")  # SYMBOL の字体：字形の番号のまま
    text, how = namecode.read_ansi("申請".encode("cp932"), 1)  # DEFAULT：文字コードを決めない
    assert text == "申請" and "決められない" in how


@pytest.mark.parametrize("text,charset", [("山田太郎", 128), ("김철수", 129), ("张伟", 134), ("Иванов", 128), ("Їжак", 204), ("Müller", 0),
                                          ("Γιώργος", 161), ("山田 Иванов", 128), ("สมชาย", 222)])
def test_control_picture_in_wmf_is_written_in_a_charset_that_has_the_letters(text, charset):
    from test_activex import morph
    from todoku import activex as ax
    from todoku import axdraw

    # 部品の字体は ＭＳ ゴシック（文字セット 128）：その文字セットで書ける字はそのまま、書けない字は、書ける文字セットで
    c = ax.parse(morph(1, text), "TextBox")
    assert c.font["charset"] == 128
    wmf = axdraw.to_wmf(axdraw.render(c, 120, 24, 96.0), (0, 0, 1800, 360), 1440)
    fonts = [body[13] for func, body in axdraw.wmf_records(wmf) if func == 0x02FB]
    assert charset in fonts
    assert axdraw.wmf_texts(wmf) == [text] and axdraw.check_picture(wmf) == []


def test_charset_for_a_font_without_a_code_page_prefers_single_byte_charsets():
    from todoku.axdraw import _charset_for

    assert _charset_for("abc", 1) == (1, "ascii")
    assert _charset_for("Müller", 1) == (0, "cp1252") and _charset_for("Иванов", 1) == (204, "cp1251")
    assert _charset_for("Иванов", 128) == (128, "cp932") and _charset_for("山田 Иванов", 1) == (128, "cp932")
    assert _charset_for("김철수", 0) == (129, "cp949") and _charset_for("שלום", 0) == (177, "cp1255")
    assert _charset_for("山田 김철수 สมชาย", 128) is None


def test_filling_a_control_redraws_its_wmf_picture_in_any_one_script(tmp_path):
    from test_activex import ax_docx, word_controls
    from todoku import axdraw

    p = _p(tmp_path, "ax.docx", ax_docx(word_controls()))
    for value in ("김철수", "Γιώργος", "Їжак"):
        out = str(tmp_path / "out.docx")
        assert api.fill(p, {"TextBox1": value}, out, overwrite=True)["read_back_ok"]
        with zipfile.ZipFile(out) as z:  # 閉じてから次を書く（Windows では、開いているファイルは置き換えられない）
            assert [value] in [axdraw.wmf_texts(z.read(n)) for n in z.namelist() if n.endswith(".wmf")]
    with pytest.raises(api.ApiError, match="1 つの文字セット"):  # 2 つの文字の混ざった字は、字を落とさずに断る
        api.fill(p, {"TextBox1": "山田 김철수 สมชาย"}, str(tmp_path / "no.docx"))


def test_metafile_text_in_a_font_without_a_code_page_is_told():
    from test_activex import morph
    from todoku import activex as ax
    from todoku import axdraw
    from todoku.imaging import metaraster

    wmf = axdraw.to_wmf(axdraw.render(ax.parse(morph(1, "山田太郎"), "TextBox"), 120, 24, 96.0), (0, 0, 1800, 360), 1440)
    # 文字セット 128（Shift_JIS）の字体：決まる（この機械に字の形がない、という知らせは別のこと）
    assert not any("規則では決められない字" in x for x in metaraster.wmf(wmf).problems)
    i = wmf.index(b"\xfb\x02")  # META_CREATEFONTINDIRECT の文字セットを DEFAULT（1）にする
    assert wmf[i + 15] == 128
    unknown = wmf[:i + 15] + b"\x01" + wmf[i + 16:]
    assert any("規則では決められない字" in x for x in metaraster.wmf(unknown).problems)
    assert axdraw.wmf_texts(unknown) == ["山田太郎"]


def test_jet3_password_code_page():
    from todoku.legacy.jet import _codepage_codec

    assert (_codepage_codec(932), _codepage_codec(1251), _codepage_codec(949), _codepage_codec(4242)) == ("cp932", "cp1251", "cp949", "cp1252")


def test_xls_file_link_with_only_an_8bit_path_is_told():
    from todoku.xls.book import _FILE_MONIKER, _hyperlink_object

    def link(path: bytes, wide: str = "") -> bytes:
        w = wide.encode("utf-16-le")
        tail = struct.pack("<I", len(w) + 6) + struct.pack("<IH", len(w), 3) + w if wide else struct.pack("<I", 0)
        return (struct.pack("<II", 2, 0x01) + _FILE_MONIKER + struct.pack("<HI", 0, len(path) + 1) + path + b"\0"
                + bytes(24) + tail)

    how: dict = {}
    assert _hyperlink_object(link(b"a\\b.xls"), 0, how) == "a\\b.xls" and how == {}
    how = {}
    shown = _hyperlink_object(link("отчёт.xls".encode("cp1251")), 0, how)
    assert "決められない" in how["undetermined"] and shown.endswith(".xls")
    how = {}
    assert _hyperlink_object(link(b"??.xls", "отчёт.xls"), 0, how) == "отчёт.xls" and how == {}  # Unicode の名前があれば、それ


def test_old_word_language_decides_the_code_page():
    from todoku.doc.word6 import lid_codec

    assert (lid_codec(0x0411), lid_codec(0x0412), lid_codec(0x0804), lid_codec(0x0404), lid_codec(0x0419), lid_codec(0x0409)) == (
        "cp932", "cp949", "cp936", "cp950", "cp1251", "cp1252")
    assert lid_codec(0x03FF) is None
