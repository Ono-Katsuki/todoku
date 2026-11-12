"""切り替えの印で読み方の変わる文字コード（ISO-2022-JP の仲間・ISO-2022-KR・HZ・UTF-7）の、文字だけのファイル。

バイトはどれも ASCII の範囲で、ESC・SO/SI・「~{」・「+」の印の後ろのバイトが、別の文字集合の字になる。
- ISO/IEC 2022 の印のとおりに厳密に読めるファイルは、規則で見分ける（色の指定の ESC・印のない SO/SI は、見分けない）。
- 範囲で読む道（rows・paragraphs・records）は、UTF-8 に直した写しを読むので、全体の窓と同じ字になる。
- 書き込みは、書いた所のバイトだけを変え、書いた後ろの字の読み方を、デコーダで確かめて元と同じに保つ。前後の字と
  ひと続きのバイトになる所は、そのひと続きを同じ字で書き直す。断るのは、その文字コードにない字だけ。
どれも、作ったファイルで確かめる。
"""

from __future__ import annotations

import os
import random
import re

import pytest

from todoku import api, lazy, limits, mcp_server
from todoku.datafile import textenc
from todoku.datafile.textenc import EncodingError, decode, kept_encoding, stateful
from todoku.lazy import textstream
from todoku.lazy.textstream import MemorySource, RecodedSource, Source, detect

JA_CSV = "氏名,住所,備考\n佐藤 花子,ながの,\"あり、が\"\n田中 太郎,大阪府,なし\n山田,東京,\n"
KO_CSV = "성명,주소,비고\n홍 길동,서울시,\"가나，울\"\n김 이박,대판부,다\n산,동경,\n"
ZH_CSV = "氏名,住所,备考\n佐腾 花子,一二三,\"四五、六\"\n田中 太郎,大阪府,七\n山田,东京,\n"
JA_TXT = "申請書\n\n氏名：\n住所：ながの\n\n以上、よろしく。\n"
KO_TXT = "신청서\n\n성명:\n주소:서울시\n\n이상.\n"
JA_JSONL = '{"a": "が", "b": 1}\n{"a": "氏名", "b": 2}\n{"a": "", "b": 3}\n'
JA_JSON = '{"氏名": "", "住所": {"市": "ながの", "番地": null}, "x": [1, "が"]}\n'

# (書くときの文字コード, 規則が決める文字コード)：字は同じ
BY_RULE = [("iso2022_jp", JA_CSV, "iso2022_jp"), ("iso2022_jp_2004", JA_CSV, "iso2022_jp"),
           ("iso2022_jp_ext", "名前,ｶﾅ\n山田,ﾔﾏﾀﾞ\n", "iso2022_jp_ext"), ("iso2022_jp_2004", "名前,𠮟る\n", "iso2022_jp_2004"),
           ("iso2022_jp_2", "名前,한글,ά\n", "iso2022_jp_2"), ("iso2022_kr", KO_CSV, "iso2022_kr"),
           ("iso2022_jp_1", "名前,丂\n", "iso2022_jp_1")]
# ISO/IEC 2022 の印として読めないもの：ASCII の字のまま（ESC・SO・SI は制御文字として数える）
NOT_ISO2022 = [b"\x1b[31mred\x1b[0m,1\n", b"a\x0eb\x0fc,1\n", b"\x1b$B.!\x1b(B,1\n", b"+MEw-,1\n", b"~{JOC{~},1\n",
               b"\x1b[0m\x0e\n"]


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _file(tmp_path, name: str, data: bytes) -> str:
    p = tmp_path / name
    p.write_bytes(data)
    return str(p)


def _stream(tmp_path, data: bytes, encoding=None):
    src = Source(_file(tmp_path, "s.bin", data))
    try:
        return detect(src, encoding)
    finally:
        src.close()


def _same(a: str, b: str) -> bool:
    """ファイルの名前・バイト数・SHA-256・文字コードの名前のほかは同じか。"""
    def norm(t: str) -> str:
        t = re.sub(r"[0-9a-f]{64}", "<sha>", t)
        t = re.sub(r"[\d,]+ バイト", "<n> バイト", t)
        t = re.sub(r"ISO-2022-(JP|KR)(-\w+)?|HZ-GB-2312|UTF-7|UTF-8", "<enc>", t)
        t = re.sub(r"^- 文字コードは、エージェントが選んで指定した .*\n", "", t, flags=re.M)
        t = re.sub(r"^- このファイルの文字コード（.*\n", "", t, flags=re.M)
        return re.sub(r"\b[uxs]\.(csv|txt|jsonl|json)", "f", t)
    return norm(a) == norm(b)


# ---------------------------------------------------------------------- 見分け方
@pytest.mark.parametrize("enc, text, want", BY_RULE)
def test_recognised_by_the_escape_sequences(tmp_path, enc, text, want):
    blob = text.encode(enc)
    assert blob.isascii() and (b"\x1b" in blob)
    d = decode(blob)
    assert (d.encoding, d.basis, d.text, d.candidates) == (want, "escape", text, [])
    assert not d.ascii_only and d.controls == 0
    s = _stream(tmp_path, blob)
    assert (s.encoding, s.basis, s.candidates) == (want, "escape", [])


@pytest.mark.parametrize("blob", NOT_ISO2022)
def test_other_escapes_stay_ascii(tmp_path, blob):
    d = decode(blob)
    assert (d.encoding, d.basis, d.text) == ("utf-8", "ascii", blob.decode("ascii"))
    s = _stream(tmp_path, blob)
    assert (s.encoding, s.basis) == ("utf-8", "ascii")


def test_escape_sequences_must_be_consumed():
    """字の集合の印（ESC $ B）と、色の指定の ESC が混ざったファイル：ISO-2022-JP として読み、残った ESC は制御文字と数える。"""
    blob = "あ".encode("iso2022_jp") + b"\x1b[0m\n"
    d = decode(blob)
    assert (d.encoding, d.text) == ("iso2022_jp", "あ\x1b[0m\n") and d.controls == 1


def test_names_and_given_encodings():
    assert [textenc.canonical(x) for x in ("ISO-2022-JP", "csISO2022JP", "iso-2022-kr", "HZ-GB-2312", "UTF-7")] == \
        ["iso2022_jp", "iso2022_jp", "iso2022_kr", "hz", "utf-7"]
    assert all(stateful(e) for e in (*textenc.ISO2022, "hz", "utf-7"))
    assert not any(stateful(e) for e in ("utf-8", "cp932", "euc_kr", "gb18030", "utf-16-le", "latin-1"))
    assert decode("汉字,一\n".encode("hz"), "hz").text == "汉字,一\n"
    assert decode("氏名,が\n".encode("utf-7"), "UTF-7").text == "氏名,が\n"
    with pytest.raises(EncodingError):
        decode(b"\x1b$B\x21\n", "iso2022_jp")


# ---------------------------------------------------------------------- 窓・段階・範囲
@pytest.mark.parametrize("name, text, enc, given", [
    ("x.csv", JA_CSV, "iso2022_jp", None), ("x.csv", KO_CSV, "iso2022_kr", None), ("x.csv", ZH_CSV, "hz", "hz"),
    ("x.csv", JA_CSV, "utf-7", "utf-7"), ("x.txt", JA_TXT, "iso2022_jp", None), ("x.txt", KO_TXT, "iso2022_kr", None),
    ("x.txt", JA_TXT, "utf-7", "utf-7"), ("x.jsonl", JA_JSONL, "iso2022_jp", None), ("x.json", JA_JSON, "iso2022_jp", None),
    ("x.jsonl", JA_JSONL, "utf-7", "utf-7"),
])
def test_window_grade_and_ranges_match_the_utf8_file(tmp_path, name, text, enc, given):
    u = _file(tmp_path, "u" + name[1:], text.encode("utf-8"))
    p = _file(tmp_path, name, text.encode(enc))
    wu, wp = api.window(u)[2], api.window(p, None, given)[2]
    assert len(wu) == len(wp) and all(_same(a[1], b[1]) for a, b in zip(wu, wp))
    gu, gp = api.grade(u), api.grade(p, encoding=given)
    assert (gu["grade"], gu["write"]["grade"]) == (gp["grade"], gp["write"]["grade"])
    key = {"csv": "rows", "txt": "paragraphs", "jsonl": "records", "json": None}[name.split(".")[1]]
    for rng in ([{key: "2-3"}, {key: "1-1"}, {}] if key else [{}]):
        a, _ = api.read_range(u, rng)
        b, doc = api.read_range(p, rng, None, None, given)
        assert _same(a.text, b.text) and a.next == b.next and (not key or a.covered == b.covered), rng
    fu, fp = api.fields_range(u, {}), api.fields_range(p, {}, encoding=given)
    assert [(f["id"], f.get("label"), f.get("value")) for f in fu["fields"]] == [(f["id"], f.get("label"), f.get("value")) for f in fp["fields"]]


def test_range_road_reads_a_utf8_copy_and_says_so(tmp_path):
    p = _file(tmp_path, "x.csv", JA_CSV.encode("iso2022_jp"))
    page, doc = api.read_range(p, {"rows": "2-2"})
    assert doc.recoded and not doc.streams and doc.told.encoding == "iso2022_jp" and doc.dec.encoding == "utf-8"
    assert isinstance(doc.src, MemorySource)
    assert "| R2 | 佐藤 花子 | ながの | あり、が |" in page.text
    assert "UTF-8 に直した写し" in doc.source_line() and "ISO-2022-JP" in doc.source_line()
    # 書き込み・検査・段階は、全体を読む道で（範囲の道の書類は渡さない）
    assert lazy.open_doc(p, None, None, built_ok=False) is None


def test_bytes_inside_a_shifted_run_are_not_delimiters(tmp_path):
    """JIS の字の 2 バイト目は「,」（0x2C）・「"」（0x22）・改行にはならないが、1 バイト目・2 バイト目は 0x21〜0x7E のどれにもなる：
    「、」は 0x21 0x22、「ぬ」は 0x24 0x4C、「，」は 0x21 0x24。写しを読むので、字のバイトを区切りと読まない。"""
    text = "a,b\n、，,\"ぬ、\"\n1,2\n"
    blob = text.encode("iso2022_jp")
    assert blob.count(b",") > text.count(",") or blob.count(b'"') > text.count('"')
    p = _file(tmp_path, "x.csv", blob)
    page, _doc = api.read_range(p, {"rows": "2-3"})
    assert "| R2 | 、， | ぬ、 |" in page.text and "| R3 | 1 | 2 |" in page.text


def test_large_files_use_a_temporary_copy_that_is_removed(tmp_path, monkeypatch):
    monkeypatch.setattr(textstream, "RECODE_IN_MEMORY", 0)
    made: list[str] = []
    real = RecodedSource.__init__

    def spy(self, src, enc, start):
        real(self, src, enc, start)
        made.append(self._tmp or "")

    monkeypatch.setattr(RecodedSource, "__init__", spy)
    p = _file(tmp_path, "x.txt", JA_TXT.encode("iso2022_jp"))
    page, doc = api.read_range(p, {"paragraphs": "3-4"})
    assert isinstance(doc.src, RecodedSource) and "住所：ながの" in page.text
    assert doc.src.read(0, doc.src.size).decode("utf-8") == JA_TXT
    tmp = doc.src._tmp
    doc.src.close()
    assert made and (tmp is None or not os.path.exists(tmp))  # POSIX は作ってすぐ消す。Windows は閉じるときに消す


def test_audit_of_a_recoded_file_runs_on_the_whole_file(tmp_path):
    p = _file(tmp_path, "x.csv", JA_CSV.encode("iso2022_jp"))
    au = api.audit(p)
    for _ in range(50):
        if au.get("checks"):
            break
        au = api.audit(p)
    assert au["ok"] and all(c["ok"] for c in au["checks"]) and len(au["checks"]) >= 4


# ---------------------------------------------------------------------- 書き込み
def _fill(tmp_path, name, blob, values, encoding=None):
    p = _file(tmp_path, name, blob)
    out = str(tmp_path / ("out_" + name))
    r = api.fill(p, values, out, encoding=encoding)
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"] and not r["unexpected"], r
    return open(out, "rb").read()


def test_fill_csv_iso2022_jp_changes_only_the_written_cells(tmp_path):
    blob = JA_CSV.encode("iso2022_jp")
    new = _fill(tmp_path, "x.csv", blob, {"R2.C2": "松本、市", "R3.備考": "ab", "R5.C1": "鈴木"})
    assert new.decode("iso2022_jp") == JA_CSV.replace("ながの", "松本、市").replace("なし", "ab") + "鈴木,,\n"
    assert decode(new).encoding == "iso2022_jp"
    head = "氏名,住所,備考\n佐藤 花子,".encode("iso2022_jp")
    assert new.startswith(head)  # 書いた所より前のバイトはそのまま


def test_fill_csv_iso2022_kr_keeps_the_shift_state(tmp_path):
    new = _fill(tmp_path, "x.csv", KO_CSV.encode("iso2022_kr"), {"R2.C2": "부산 a", "R3.C3": "x", "R4.C3": "끝"})
    assert new.decode("iso2022_kr") == KO_CSV.replace("서울시", "부산 a").replace(",다\n", ",x\n").replace("동경,\n", "동경,끝\n")
    assert new.count(b"\x1b$)C") == 1  # 行の頭の印は 1 つのまま


def test_fill_txt_inside_and_after_a_shifted_run(tmp_path):
    new = _fill(tmp_path, "x.txt", JA_TXT.encode("iso2022_jp"), {"P3.c3": "山田 太郎", "P1.c3": "abc"})
    assert new.decode("iso2022_jp") == JA_TXT.replace("申請書\n", "申請書abc\n").replace("氏名：\n", "氏名：山田 太郎\n")
    new = _fill(tmp_path, "k.txt", KO_TXT.encode("iso2022_kr"), {"P3.c3": "홍 a길"})
    assert new.decode("iso2022_kr") == KO_TXT.replace("성명:\n", "성명:홍 a길\n")


def test_fill_hz_and_json(tmp_path):
    new = _fill(tmp_path, "x.csv", ZH_CSV.encode("hz"), {"R2.C2": "一a二", "R5.C1": "三"}, "hz")
    assert new.decode("hz") == ZH_CSV.replace("一二三", "一a二") + "三,,\n"
    new = _fill(tmp_path, "x.json", JA_JSON.encode("iso2022_jp"), {"/氏名": "山田", "/住所/市": "x"})
    assert new.decode("iso2022_jp") == JA_JSON.replace('"氏名": ""', '"氏名": "山田"').replace('"ながの"', '"x"')
    new = _fill(tmp_path, "x.jsonl", JA_JSONL.encode("iso2022_jp"), {"L1/a": "穐", "L3/a": "x"})
    assert new.decode("iso2022_jp") == JA_JSONL.replace('"が"', '"穐"').replace('"a": ""', '"a": "x"')


def test_characters_outside_the_encoding_are_refused(tmp_path):
    p = _file(tmp_path, "x.csv", JA_CSV.encode("iso2022_jp"))
    with pytest.raises(api.ApiError, match="ISO-2022-JP）では書けない字"):
        api.fill(p, {"R2.C2": "ｶﾅ"}, str(tmp_path / "o.csv"))
    assert not (tmp_path / "o.csv").exists()


def test_korean_before_a_late_iso2022_kr_header_is_written(tmp_path):
    """ISO-2022-KR の印（ESC $ ) C）より前に韓国語の字を書く：書く字の前に印を置き、後ろの元のバイトは変えない
    （前は断っていた）。印は 2 つになるが、どちらの後ろも同じ読み方。"""
    late = "a,b\n1,홍\n".encode("iso2022_kr")
    assert late == b"a,b\n1,\x1b$)C\x0eH+\x0f\n"  # Python は、初めの韓国語の字の前に印を置く
    kr, tail = b"\x1b$)C", b"\n1,\x1b$)C\x0eH+\x0f\n"
    assert _fill(tmp_path, "k1.csv", late, {"R1.C1": "성"}) == kr + b"\x0e<:\x0f,b" + tail
    assert _fill(tmp_path, "k2.csv", late, {"R1.C1": "성", "R1.C2": "명"}) == kr + b"\x0e<:\x0f,\x0e8m\x0f" + tail
    assert _fill(tmp_path, "k3.csv", late, {"R1.C1": "성", "R2.C2": "길"}) == kr + b"\x0e<:\x0f,b\n1,\x1b$)C\x0e1f\x0f\n"
    assert _fill(tmp_path, "k4.csv", late, {"R1.C1": "z", "R2.C2": "성"}) == b"z,b\n1,\x1b$)C\x0e<:\x0f\n"
    assert _fill(tmp_path, "k5.csv", late, {"R3.C1": "끝"}) == late + b"\x0e3!\x0f,\n"
    for name, want in (("k1", "성,b\n1,홍\n"), ("k2", "성,명\n1,홍\n"), ("k3", "성,b\n1,길\n"), ("k5", "a,b\n1,홍\n끝,\n")):
        got = decode(open(tmp_path / f"out_{name}.csv", "rb").read())
        assert (got.encoding, got.text) == ("iso2022_kr", want)


def test_utf7_inside_and_after_a_base64_run_is_written(tmp_path):
    """UTF-7 の、base64 のひと続きの中・すぐ後ろに書く：デコーダがまとめて字にするひと続きを、同じ字で書き直す
    （前は断っていた）。ひと続きの外のバイトは変えない。"""
    u = JA_CSV.encode("utf-7")
    assert u == (b'+bA9UDQ,+T09iQA,+UJmAAw\n+T1CF5A +grFbUA,+MGowTDBu,"+MEIwijABMEw"\n'
                 b"+dTBOLQ +WSqQzg,+WSeWKl6c,+MGowVw\n+XHF1MA,+Z3FOrA,\n")
    new = _fill(tmp_path, "u1.csv", u, {"R2.C2": "松本"}, "utf-7")
    assert new == u.replace(b"+MGowTDBu", b"+Z35nLA")
    new = _fill(tmp_path, "u2.csv", u, {"R2.C2": "松本", "R2.C1": "a+b"}, "utf-7")
    assert new == u.replace(b"+MGowTDBu", b"+Z35nLA").replace(b"+T1CF5A +grFbUA", b"a+-b")
    new = _fill(tmp_path, "u3.csv", u, {"R4.C3": "x-y", "R3.C2": "", "R5.C1": "終"}, "utf-7")
    assert new == u.replace(b"+WSeWKl6c", b"").replace(b"+Z3FOrA,\n", b"+Z3FOrA,x-y\n+fUI,,\n")
    assert new.decode("utf-7") == JA_CSV.replace("大阪府", "").replace("東京,\n", "東京,x-y\n終,,\n")
    # ひと続きの中の 1 字だけを書き換える：そのひと続きのバイトは変わる（字は同じ）
    one = "氏名,あいう\n".encode("utf-7")
    assert one == b"+bA9UDQ,+MEIwRDBG\n"
    assert _fill(tmp_path, "u4.csv", one, {"R1.C2": "あxう"}, "utf-7") == b"+bA9UDQ,+MEI-x+MEY\n"
    assert _fill(tmp_path, "u5.csv", "あ,い\n".encode("utf-7"), {"R1.C2": "-"}, "utf-7") == b"+MEI,-\n"


def test_a_file_left_with_ascii_only_keeps_its_encoding(tmp_path):
    """ASCII でない字がなくなったファイルには、文字コードの手がかりが残らない：元の文字コードで読んで同じ字なら、変わっていない。"""
    for enc in ("iso2022_jp", "cp932", "euc_kr"):
        ch = "가" if enc == "euc_kr" else "あ"
        new = _fill(tmp_path, f"{enc}.csv", f"a,{ch}\n1,2\n".encode(enc), {"R1.C2": "x"})
        assert new == b"a,x\n1,2\n"
    a = decode("一\n".encode("hz"), "hz")
    assert not kept_encoding(a, decode(b"~{\n"), b"~{\n")  # HZ では「~{」は切り替えの印：ASCII の字のファイルと同じ字にならない
    assert kept_encoding(a, decode(b"x\n"), b"x\n")


POOLS = {"iso2022_jp": "あが、※穐aB,\"1 \n漢字", "iso2022_jp_ext": "あがｶﾅaB,\"\n漢", "iso2022_jp_2004": "あ𠮟が,a\"\n漢",
         "iso2022_jp_2": "あ한é,a\"\nü漢", "iso2022_kr": "한글가,a\" \nb", "iso2022_jp_3": "あ塡a,\n漢", "hz": "汉字一a,\"\nb",
         "utf-7": "あ한aB+-,\"1 \n漢😀", "iso2022_jp_ms": "あ①髙ｶ¥aB,\"1 \n漢～"}
GIVEN = {"hz": "hz", "utf-7": "utf-7", "iso2022_jp_ms": "cp50221"}


@pytest.mark.parametrize("enc", sorted(POOLS))
def test_patches_decode_to_the_expected_text(enc):
    """いろいろな字の並びに、重ならない書き込みをいくつかして、できたバイトを読むと、字の上での置き換えと同じになる。
    断るのは、その文字コードで書けない字だけ（書かずに断る）：切り替えの印の決まりで書けない所はない。"""
    rnd = random.Random(20261112)
    pool, done, refused = POOLS[enc], 0, 0
    other = "".join(sorted(set("".join(POOLS.values())) - set(pool)))
    for _ in range(250):
        t = "".join(rnd.choice(pool) for _ in range(rnd.randint(0, 14)))
        blob = t.encode(enc)
        d = decode(blob, GIVEN.get(enc))
        if d.text != t or not stateful(d.encoding):
            continue  # ASCII の字だけの並び
        cuts = sorted(rnd.sample(range(len(t) + 1), min(len(t) + 1, 4)))
        edits = []
        for i in range(len(cuts) - 1):
            a, e = cuts[i], cuts[i + 1] if rnd.random() < 0.7 else cuts[i]
            s = "".join(rnd.choice(pool) for _ in range(rnd.randint(0, 3)))
            if rnd.random() < 0.05:
                s += rnd.choice(other)
            try:
                d.encode(s)
            except EncodingError as err:
                assert "書けない字" in str(err)
                refused += 1
                continue
            edits.append((a, e, s))
        want, at = "", 0
        for a, e, s in edits:
            want, at = want + t[at:a] + s, e
        want += t[at:]
        out = d.rewrite(blob, edits)
        assert out.decode(d.encoding) == want, (t, edits, out)
        if not edits:
            assert out == blob
        one = [d.rewrite(blob, [x]) for x in edits]  # 1 つずつ書いても、どれも書ける
        assert len(one) == len(edits)
        done += len(edits)
    assert done > 300 and refused < done / 4


def test_rewrite_refuses_overlapping_places_and_checks_its_own_output():
    blob = JA_CSV.encode("iso2022_jp")
    d = decode(blob)
    with pytest.raises(EncodingError, match="重なっています"):
        d.rewrite(blob, [(0, 2, "a"), (1, 3, "b")])
    assert d.rewrite(blob, []) == blob
    assert d.rewrite(blob, [(3, 5, "町"), (0, 2, "名前")]).decode("iso2022_jp") == JA_CSV.replace("氏名,住所", "名前,町")
    # pieces：書き換えるバイトの範囲と新しいバイト（並びは位置の順。範囲は重ならない）
    ps = d.pieces(blob, [(3, 5, "町"), (0, 2, "名前")])
    assert [x for x, _y, _n in ps] == sorted(x for x, _y, _n in ps) and all(ps[i][1] <= ps[i + 1][0] for i in range(len(ps) - 1))
    out, at = bytearray(), 0
    for x, y, new in ps:
        out += blob[at:x] + new
        at = y
    assert bytes(out + blob[at:]) == d.rewrite(blob, [(3, 5, "町"), (0, 2, "名前")])


def test_stateless_patch_is_span_and_encode():
    d = decode("氏名,あ\n".encode("cp932"))
    blob = "氏名,あ\n".encode("cp932")
    assert d.patch(blob, 3, 4, "い") == (5, 7, "い".encode("cp932"))
    d = decode("氏名,あ\n".encode("utf-8"))
    assert d.patch("氏名,あ\n".encode("utf-8"), 0, 2, "x") == (0, 6, b"x")


def test_large_file_is_read_by_range_and_written_through_the_whole_file(tmp_path, monkeypatch):
    """窓の全体が予算に入らないファイル：概要と範囲は、写しを読む道。書き込み・段階は、元のファイルの全体を読む道。"""
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    monkeypatch.setattr(limits, "PAGE_CHARS", 1_000)
    rows = ["番号,氏名,備考"] + [f"{i},山田{i}号、太郎,\"あ,{i}\"" for i in range(2, 801)]
    text = "\n".join(rows) + "\n"
    data = text.encode("iso2022_jp")
    p = _file(tmp_path, "big.csv", data)
    doc = lazy.open_doc(p)
    assert doc is not None and doc.large and doc.recoded
    sc = mcp_server.tool_result("open_document", {"path": p})["structuredContent"]
    assert sc["large"] and sc["format"]["encoding"] == "iso2022_jp" and sc["format"]["encoding_basis"] == "escape"
    assert "UTF-8 に直した写し" in mcp_server.tool_result("open_document", {"path": p})["content"][1]["text"]
    assert lazy.open_doc(p, built_ok=False) is None  # 書き込み・段階には、範囲の道の書類を渡さない
    page, _doc = api.read_range(p, {"rows": [400, 401]})
    assert "| R400 | 400 | 山田400号、太郎 | あ,400 |" in page.text and "| R401 | 401 | 山田401号、太郎 | あ,401 |" in page.text
    fl = api.fields_range(p, {"rows": [400, 400]})
    assert {"R400.C1", "R400.C2", "R400.C3"} <= {f["id"] for f in fl["fields"]}
    out = tmp_path / "big_out.csv"
    res = api.fill(p, {"R400.C2": "鈴木、花子", "R800.C3": "x"}, str(out))
    assert res["proof_ok"] and res["read_back_ok"] is True
    want = text.replace("山田400号、太郎", "鈴木、花子").replace('"あ,800"', '"x"')  # 引用符で囲んだマスは、囲んだまま
    assert out.read_bytes().decode("iso2022_jp") == want
    assert out.read_bytes()[:1000] == data[:1000]
    assert api.grade(p)["grade"] == "S"


@pytest.mark.parametrize("name, body", [
    ("a.html", '<html><head><meta charset="iso-2022-jp"><title>申請</title></head><body><p>氏名：山田</p></body></html>'),
    ("b.html", "<html><body><h1>申請書</h1><p>氏名：山田</p></body></html>"),
    ("c.xml", '<?xml version="1.0" encoding="ISO-2022-JP"?><a><b>氏名：山田</b></a>'),
])
def test_markup_files_are_read_with_the_same_rule(tmp_path, name, body):
    """HTML・XML も、同じ決め方で読む（前は、宣言があっても、ESC の混ざった ASCII の字として読んでいた）。"""
    p = _file(tmp_path, name, body.encode("iso2022_jp"))
    kind, win, parts = api.window(p)
    text = "".join(x[1] for x in parts)
    assert kind == "markup" and win["format"]["encoding"] == "iso2022_jp" and "氏名：山田" in text and "\x1b" not in text
