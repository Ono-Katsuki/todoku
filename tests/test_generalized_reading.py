"""読み方の決まりは、国や形式によらず同じ：規則で決められないことは、決めずに事実を出す。
BOM も宣言もない 8 ビットの字のファイルの文字コード（多バイトで読めても、1バイトの文字コードを候補に出す）と、
CSV の区切り（どの行にも同じ数ある区切り）。ふつうの道と大きな書類の道で同じ答え。"""

from __future__ import annotations

import random

import pytest

from todoku import api, lazy, limits, mcp_server
from todoku.datafile import textenc
from todoku.datafile.csvfile import delimiters
from todoku.datafile.textenc import decode
from todoku.lazy import textstream

UNDECIDED = "文字コードを規則では決められません"

# その国の字を、その国でふつうの文字コードで書いたもの（BOM なし・宣言なし）
TEXTS = {
    "cp1252": "Name;Ort\nMüller;Köln\nFrançois;Besançon\n",
    "cp1250": "Jméno,Město\nJiří Dvořák,Plzeň\n",
    "cp1251": "Имя,Город\nИван,Москва\n",
    "cp1253": "Όνομα,Πόλη\nΓιώργος,Αθήνα\n",
    "cp1254": "Ad,Şehir\nGökçe,İstanbul\n",
    "cp1255": "שם,עיר\nדוד,תל אביב\n",
    "cp1256": "الاسم,المدينة\nمحمد,القاهرة\n",
    "cp874": "ชื่อ,เมือง\nสมชาย,กรุงเทพ\n",
    "cp932": "氏名,住所\n山田太郎,東京都新宿区\nｶﾀｶﾅ,半角\n",
    "euc_jp": "氏名,住所\n山田太郎,東京都新宿区\n",
    "cp950": "姓名,地址\n王小明,臺北市\n",
    "gbk": "姓名,地址\n王小明,北京市朝阳区\n",
    "cp949": "이름,주소\n홍길동,서울특별시\n",
}


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _stream(tmp_path, blob: bytes):
    p = tmp_path / "s.bin"
    p.write_bytes(blob)
    src = textstream.Source(str(p))
    try:
        return textstream.detect(src)
    finally:
        src.close()


def _names(d) -> list[str]:
    return [d.encoding] + [c["encoding"] for c in d.candidates]


def _large(monkeypatch, path):
    """ここから先は、大きな書類の道（窓の全体が予算に入らない）。"""
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 300)
    lazy.clear()
    api._WINDOWS.clear()
    assert api.lazy_open(str(path)) is not None


# ====================================================================== 文字コード
@pytest.mark.parametrize("enc", sorted(TEXTS))
def test_the_right_encoding_is_always_among_the_readings(enc, tmp_path):
    """どの国の文字コードで書いたファイルも、正しい読み方が窓か候補にあり、候補の見本は、その文字コードで読んだ字。"""
    blob = TEXTS[enc].encode(enc)
    d = decode(blob)
    assert d.basis == "order" and d.undetermined
    right = next((c for c in d.candidates if c["encoding"] == enc), None)
    assert d.encoding == enc or right is not None
    for c in d.candidates:
        line = blob.decode(c["encoding"]).splitlines()[c["line"]]
        assert c["sample"].strip("…") in line and d.text.splitlines()[c["line"]] != line
    if right is not None:
        assert right["sample"] == TEXTS[enc].splitlines()[right["line"]]
        assert decode(blob, enc).text == TEXTS[enc]  # 選んで渡すと、その字で読む
    s = _stream(tmp_path, blob)  # 大きな書類の道（ファイルを少しずつ読む）も、同じ答え・同じ注意
    assert (s.encoding, s.basis, s.info(), s.note_lines()) == (d.encoding, d.basis, d.info(), d.note_lines())


def test_a_multibyte_reading_alone_does_not_decide():
    """ドイツ語の Windows-1252 は GBK として、タイ語の Windows-874 は Shift_JIS の半角カナとして、厳密に読める。
    多バイトで読めたことでは決めない（前は、候補を出さずに GBK・Shift_JIS に決めていた）。"""
    d = decode("Müller;Köln\n".encode("cp1252"))
    assert d.encoding == "gbk" and d.basis == "order"
    assert {"encoding": "cp1252", "label": "Windows-1252（西欧）", "line": 0, "sample": "Müller;Köln"} in d.info()["encoding_candidates"]
    d = decode(TEXTS["cp874"].encode("cp874"))
    assert d.encoding == "cp932" and "cp874" in _names(d)
    # 半角カナだけの Shift_JIS も、同じ決まり（Windows-874 などとしても、きれいに読める）
    d = decode("ｶﾅ,1\r\nｱｲｳ,2\r\n".encode("cp932"))
    assert d.encoding == "cp932" and d.basis == "order" and "cp874" in _names(d)


def test_single_byte_readings_equal_decoding_the_whole_file_with_each():
    """1バイトの文字コードは、ファイルにあるバイトの種類の表で見る：文字コードの数だけ読み直したときと同じ答え。"""
    rnd = random.Random(11)
    for _ in range(300):
        n = rnd.choice([1, 3, 20, 200])
        blob = bytes(rnd.choice([rnd.randrange(32, 127), rnd.randrange(128, 256), 10, 13]) for _ in range(n * rnd.randrange(1, 6)))
        want: list[tuple[str, str]] = []
        for enc in textenc.SINGLE:
            t = textenc._strict(blob, enc)
            if t is not None and textenc.clean(t) and all(t != u for _e, u in want):
                want.append((enc, t))
        high = textenc.high_bytes((blob,))
        assert high == sorted({b for b in blob if b >= 128})
        assert [e for e, _k in textenc.single_readings(high)] == [e for e, _t in want]
        d = decode(blob)
        if d.basis == "order":
            assert _names(d) == [e for e, _t in textenc._readings(blob, textenc.MULTI)] + [e for e, _t in want]


def test_sample_is_cut_around_the_first_different_place(tmp_path):
    """見本は、字の違う最初の所の少し前から（長い行をそのまま並べない）。行の番号は、その行。"""
    long_line = "x" * 500 + ";Müller;" + "y" * 500
    blob = ("id;name\n" + long_line + "\n").encode("cp1252")
    d = decode(blob)
    c = next(c for c in d.candidates if c["encoding"] == "cp1252")
    assert c["line"] == 1 and c["sample"].startswith("…") and c["sample"].endswith("…") and "Müller" in c["sample"]
    assert len(c["sample"]) <= textenc.SNIP + 2
    own = d.sample_of(d.text, c)
    assert own.startswith("…") and "M黮ler" in own and len(own) <= textenc.SNIP + 2
    assert all("own_sample" not in x for x in d.info()["encoding_candidates"])  # 結果には入れない（窓の注意にだけ使う）
    s = _stream(tmp_path, blob)
    assert s.note_lines() == d.note_lines()
    # 行の頭にある所は、頭を切らない。CRLF・CR でも、行の番号は同じ
    for nl in ("\n", "\r\n", "\r"):
        d = decode(("a,b" + nl + "c,d" + nl + "Köln,1" + nl).encode("cp1252"))
        c = next(c for c in d.candidates if c["encoding"] == "cp1252")
        assert (c["line"], c["sample"]) == (2, "Köln,1")


def test_undecided_encoding_reaches_window_grade_and_tools(tmp_path, monkeypatch):
    p = tmp_path / "de.csv"
    p.write_bytes((TEXTS["cp1252"] + "Meier;Bonn\n" * 60).encode("cp1252"))
    for large in (False, True):
        if large:
            _large(monkeypatch, p)
        g = api.grade(str(p))
        assert g["grade"] == "B" and g["write"]["grade"] == "B"
        r = next(r for r in g["reasons"] if r["title"] == "文字コードを規則では決められない")
        assert r["where"][0] == "gbk" and "cp1252" in r["where"]
        got = mcp_server.call_tool("open_document", {"path": str(p)})
        text = "\n".join(c["text"] for c in got["content"])
        assert UNDECIDED in text and 'encoding: "cp1252"）：Müller;Köln' in text
        fm = got["structuredContent"]["format"]
        assert "cp1252" in {c["encoding"] for c in fm["encoding_candidates"]}
        assert "encoding_notes" not in fm  # 注意の字は、窓の字に 1 度だけ
        # 選んで渡すと、その字で読み、段階は下がらない
        g = api.grade(str(p), encoding="cp1252")
        assert g["grade"] == "S" and g["write"]["grade"] == "A"
        got = mcp_server.call_tool("read_window", {"path": str(p), "encoding": "cp1252", **({"rows": "1-3"} if large else {})})
        text = "\n".join(c["text"] for c in got["content"])
        assert "Müller" in text and "François" in text and UNDECIDED not in text


def test_write_with_the_chosen_single_byte_encoding(tmp_path):
    p = tmp_path / "de.csv"
    p.write_bytes(TEXTS["cp1252"].encode("cp1252"))
    out = tmp_path / "out.csv"
    r = api.fill(str(p), {"R2.C2": "Düsseldorf"}, str(out), encoding="cp1252", overwrite=True)
    assert r["read_back_ok"]
    assert out.read_bytes() == TEXTS["cp1252"].replace("Köln", "Düsseldorf").encode("cp1252")


# ====================================================================== CSV の区切り
@pytest.mark.parametrize("text, want", [
    ("Name;Preis;Ort\nMeier;1,50;Bonn\nSchmidt;2;Köln\n", (";", [], False)),  # 小数点がカンマ：カンマの数は行によって違う
    ("Name\tPreis\nMeier\t1,50\nKurz\t2\n", ("\t", [], False)),
    ("氏名,住所\n山田,東京\n", (",", [], False)),
    ("name,note\nx,\"a;b\"\ny,c\n", (",", [], False)),  # 引用符の中は数えない
    ("a,b\n\"x\ny\",2\n", (",", [], False)),  # 引用符の中の改行
    ("title\na,b,c\nd,e\n", (",", [], False)),  # どれもそろっていない：カンマ
    ("abc\ndef\n", (",", [], False)),  # 1列だけ
    ("a\tb\nc\td\n", ("\t", [], False)),
    ("Meier;1,50;Bonn\nSchmidt;2,75;Köln\n", (",", [(";", 3)], False)),  # 見出しがなく、どちらもそろう：決めない
    ("a\tb;c\nd\te;f\n", ("\t", [(";", 2)], False)),
    ("sep=;\nName;Preis\nMeier;1,50\n", (";", [], True)),  # Excel の、区切りの宣言の行
    ("sep=,\na,b\n", (",", [], True)),
    ("", (",", [], False)),
])
def test_delimiter_is_the_one_every_line_has_the_same_number_of(text, want):
    assert delimiters(text, "csv") == want
    assert delimiters(text, "tsv") == ("\t", [], False)


def test_decimal_comma_file_is_split_at_semicolons_on_both_roads(tmp_path, monkeypatch):
    p = tmp_path / "preise.csv"
    text = "Name;Preis;Ort\nMeier;1,50;Bonn\nSchmidt;2;Köln\n" + "Kurz;3;Kiel\n" * 60
    p.write_bytes(text.encode("utf-8"))
    _k, doc, _parts = api.window(str(p))
    assert doc["rows"][1] == ["Meier", "1,50", "Bonn"] and doc["format"]["delimiter"] == ";"
    assert api.grade(str(p))["grade"] == "S"
    out = tmp_path / "out.csv"
    r = api.fill(str(p), {"R3.Preis": "2,75"}, str(out), overwrite=True)
    assert r["read_back_ok"] and out.read_text(encoding="utf-8") == text.replace("Schmidt;2;", "Schmidt;2,75;")
    _large(monkeypatch, p)
    got = mcp_server.call_tool("read_window", {"path": str(p), "rows": "1-3"})
    text = "\n".join(c["text"] for c in got["content"])
    assert "| Meier | 1,50 | Bonn |" in text
    out2 = tmp_path / "out2.csv"
    r = api.fill(str(p), {"R3.Preis": "2,75"}, str(out2), overwrite=True)
    assert r["read_back_ok"] and out2.read_bytes() == out.read_bytes()


def test_two_even_delimiters_are_told_not_decided(tmp_path, monkeypatch):
    p = tmp_path / "x.csv"
    p.write_bytes(("Meier;1,50;Bonn\nSchmidt;2,75;Köln\n" * 30).encode("utf-8"))
    for large in (False, True):
        if large:
            _large(monkeypatch, p)
        got = mcp_server.call_tool("open_document", {"path": str(p)})
        text = "\n".join(c["text"] for c in got["content"])
        assert "区切りを規則では決められません" in text and "「;」（セミコロン）（3 列）" in text
        fm = got["structuredContent"]["format"]
        assert fm["delimiter"] == "," and fm["delimiter_also"] == [{"delimiter": ";", "delimiter_ja": "「;」（セミコロン）", "columns": 3}]
        g = api.grade(str(p))
        assert g["grade"] == "B" and g["write"]["grade"] == "B"
        assert any(r["title"] == "区切りを規則では決められない" and r["where"] == [";"] for r in g["reasons"])


def test_declared_delimiter_line_is_told(tmp_path):
    p = tmp_path / "s.csv"
    p.write_bytes("sep=;\nName;Preis\nMeier;1,50\n".encode("utf-8"))
    _k, doc, parts = api.window(str(p))
    assert doc["rows"][2] == ["Meier", "1,50"] and doc["format"]["delimiter_declared_by"] == "1行目の sep="
    assert "区切りの宣言の行" in parts[0][1]
    assert api.grade(str(p))["grade"] == "S"
