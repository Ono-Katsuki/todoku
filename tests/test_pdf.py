"""PDFの読み書きのテスト。

外部のライブラリは使わず、テストの中で小さなPDFを組み立てて確かめる（字・罫線・空のマス・□・日付の空欄）。
集めた公開の様式（corpus/wide_pdf）があれば、それでも確かめる。
"""

from __future__ import annotations

import glob
import zlib
from pathlib import Path

import pytest

from todoku.pdf.acrojs import JsError, commit, simple_calculate
from todoku.pdf.audit import run as audit
from todoku.pdf.crypto import aes_cbc_decrypt, aes_cbc_encrypt, aes_ecb_encrypt, rc4
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import FillError, fill, runs_for
from todoku.pdf.fonts import CMap, decode_predefined, glyph_to_unicode, split_predefined, visible_text
from todoku.pdf.layout import gap_blanks
from todoku.pdf.objects import Lexer, Name, Ref, parse_object
from todoku.pdf.proof import prove
from todoku.pdf.window import build

CORPUS = sorted(glob.glob(str(Path(__file__).parent.parent / "corpus" / "wide_pdf" / "**" / "*.pdf"), recursive=True))


# ====================================================================== 小さなPDFを組み立てる
def make_pdf(content: bytes, fonts: dict[str, bytes], extra_objs: list[bytes] | None = None, compress: bool = True) -> bytes:
    """1ページのPDF。fonts は 資源の名前 → 書体の辞書（オブジェクトの本体）。"""
    objs: list[bytes] = []
    objs.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    objs.append(b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
    font_refs = []
    base = 5 + (len(extra_objs) if extra_objs else 0)
    for i, name in enumerate(fonts):
        font_refs.append(f"/{name} {base + i} 0 R".encode())
    objs.append(b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << " + b" ".join(font_refs)
                + b" >> >> /Contents 4 0 R >>")
    data = zlib.compress(content) if compress else content
    objs.append(b"<< /Length " + str(len(data)).encode() + (b" /Filter /FlateDecode" if compress else b"") + b" >>\nstream\n"
                + data + b"\nendstream")
    for o in extra_objs or []:
        objs.append(o)
    for body in fonts.values():
        objs.append(body)
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for i, body in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for o in offs:
        out += f"{o:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(out)


def tounicode(mapping: dict[int, str]) -> bytes:
    lines = [f"<{c:04X}> <{u.encode('utf-16-be').hex().upper()}>" for c, u in mapping.items()]
    body = ("/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CMapName /T def 1 begincodespacerange <0000> <FFFF> "
            f"endcodespacerange {len(lines)} beginbfchar\n" + "\n".join(lines) + "\nendbfchar endcmap end end").encode()
    return body


def jp_font_objs(chars: str, first_obj: int) -> tuple[bytes, list[bytes], dict[str, int]]:
    """日本語の Type0 書体（埋め込みなし・ToUnicode あり）。(書体の辞書, 追加のオブジェクト, 字 → 符号)。"""
    codes = {ch: i + 1 for i, ch in enumerate(dict.fromkeys(chars))}
    cmap = tounicode({c: ch for ch, c in codes.items()})
    tu = b"<< /Length " + str(len(cmap)).encode() + b" >>\nstream\n" + cmap + b"\nendstream"
    desc = (b"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /MS-Gothic /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity)"
            b" /Supplement 0 >> /DW 1000 /FontDescriptor << /Type /FontDescriptor /FontName /MS-Gothic /Flags 4 /Ascent 880"
            b" /Descent -120 /FontBBox [0 -120 1000 880] /ItalicAngle 0 /CapHeight 700 /StemV 80 >> >>")
    font = (b"<< /Type /Font /Subtype /Type0 /BaseFont /MS-Gothic /Encoding /Identity-H /DescendantFonts ["
            + f"{first_obj} 0 R".encode() + b"] /ToUnicode " + f"{first_obj + 1} 0 R".encode() + b" >>")
    return font, [desc, tu], codes


def jp_text(codes: dict[str, int], s: str) -> bytes:
    return b"<" + "".join(f"{codes[ch]:04X}" for ch in s).encode() + b">"


def sample_form() -> bytes:
    """氏名（空のマス）・郵便番号（7桁の枠）・令和　年　月　日・□ 新規 □ 変更 のある申請書。"""
    chars = "申請書氏名郵便番号令和年月日新規変更□"
    font, extra, codes = jp_font_objs(chars, 5)
    t = lambda s: jp_text(codes, s)  # noqa: E731
    c = bytearray()
    c += b"BT /F1 16 Tf 250 780 Td " + t("申請書") + b" Tj ET\n"
    # 表：氏名のマス
    c += b"0.8 w 60 700 m 540 700 l S 60 670 m 540 670 l S 60 640 m 540 640 l S\n"
    c += b"60 700 m 60 640 l S 160 700 m 160 640 l S 540 700 m 540 640 l S\n"
    c += b"BT /F1 10 Tf 70 682 Td " + t("氏名") + b" Tj ET\n"
    c += b"BT /F1 10 Tf 70 652 Td " + t("郵便番号") + b" Tj ET\n"
    # 7桁の枠（160〜300 を 7 つに）
    for i in range(8):
        x = 160 + i * 20
        c += f"{x} 670 m {x} 640 l S\n".encode()
    # 日付の行（マスの外）
    c += b"BT /F1 10 Tf 380 600 Td " + t("令和") + b" Tj 40 0 Td " + t("年") + b" Tj 40 0 Td " + t("月") + b" Tj 40 0 Td " + t("日") + b" Tj ET\n"
    # □ の字
    c += b"BT /F1 10 Tf 70 560 Td " + t("□新規") + b" Tj 80 0 Td " + t("□変更") + b" Tj ET\n"
    return make_pdf(bytes(c), {"F1": font}, extra)


# ====================================================================== 部品
def test_objects_parse():
    lx = Lexer(b"<< /A (x\\(y\\)\\101) /B <48 49> /C [1 2.5 -3 /N#20m 4 0 R] /D true /E null >>")
    d = parse_object(lx)
    assert bytes(d["A"]) == b"x(y)A"
    assert bytes(d["B"]) == b"HI"
    assert d["C"][:3] == [1, 2.5, -3] and d["C"][3] == Name("N m") and d["C"][4] == Ref(4, 0)
    assert d["D"] is True and d["E"] is None


def test_crypto_vectors():
    k = bytes(range(16))
    assert aes_ecb_encrypt(k, bytes.fromhex("00112233445566778899aabbccddeeff")).hex() == "69c4e0d86a7b0430d8cdb78070b4c55a"
    assert rc4(b"Key", b"Plaintext").hex() == "bbf316e8d940af0ad3"
    iv = bytes(16)
    msg = b"todoku " * 5
    pad = 16 - len(msg) % 16
    assert aes_cbc_decrypt(k, iv + aes_cbc_encrypt(k, msg + bytes([pad]) * pad, iv)) == msg


def test_fonts_helpers():
    cm = CMap.parse(b"1 begincodespacerange <0000> <FFFF> endcodespacerange 1 beginbfrange <0010> <0012> <6C0F> endbfrange"
                    b" 1 beginbfchar <0020> <D83DDD0D> endbfchar")
    assert cm.to_unicode(b"\x00\x11") == "氐" and cm.to_unicode(b"\x00\x20") == "🔍"
    codes = split_predefined("申請A".encode("cp932"), "sjis")
    assert [decode_predefined(c, "cp932") for c in codes] == ["申", "請", "A"]
    assert glyph_to_unicode("Aacute") == "Á" and glyph_to_unicode("uni6C0F") == "氏" and glyph_to_unicode("f_i") == "fi"
    assert visible_text("最終更新⽇︓") == "最終更新日："


@pytest.mark.parametrize("text, n", [("令和　　年　　月　　日", 3), ("（郵便番号　　　　－　　　　）", 2),
                                     # 全角の空白 1 つ（PDF の窓では、測った 1 em 以上の空き）も空き
                                     ("　　　都　道　　　　　市　区　　　", 5), ("　　第　号", 2), ("令和　年　月　日", 3),
                                     # 字間を空けた見出しの形（「第　号」「氏　　名」）：落とさず spaced の「？」
                                     # 中央にそろえた行の両側の空きは、行の中に空きのないときだけ、そろえ方の空き（行の中に字間の
                                     # 空きがあれば、両側も 1 字ぶん以上の空きとして示す：見出しかどうかはエージェントが見る）
                                     ("第　号", 1), ("氏　　名", 1), ("　　　氏　　名　　　", 3), ("　　　申請書　　　", 0),
                                     ("金　　　　円", 1),
                                     # マスの縁までの空き（行の終わりの後ろ）：語は見ないので、見出しの後ろも ？ の候補
                                     ("付表第一号（一）　　　", 1),
                                     # 中華民國・韓国・中国の日付と住所、点線・下線・かっこ・チェックの箱（語の表なしで）
                                     ("中華民國　　年　　月　　日", 3), ("　　년　　월　　일", 3), ("　　省　　市　　县", 3),
                                     ("Name ..............................", 1), ("Nom : …………………………", 1),
                                     ("नाम ...................", 1), ("Signature ______", 1), ("[ ] Yes [ ] No", 2),
                                     ("( ) Male ( ) Female", 2), ("الاسم: ____", 1), ("──────────────", 0)])
def test_pdf_blanks(text, n):
    lead = len(text) - len(text.lstrip("　")) if text.startswith("　") else 0
    trail = len(text) - len(text.rstrip("　"))
    assert len(gap_blanks(text, lead, trail)) == n


def test_acrojs_formats():
    assert commit({"F": 'AFDate_FormatEx("dd-mmm-yyyy");'}, "令和8年4月1日") == "01-Apr-2026"
    assert commit({"F": 'AFDate_FormatEx("mm/dd/yyyy");'}, "2026-11-02") == "11/02/2026"
    assert commit({"F": 'AFNumber_Format(2, 0, 0, 0, "$", true);'}, "1234.5") == "$1,234.50"
    assert commit({"F": "AFSpecial_Format(0);"}, "12345") == "12345"
    with pytest.raises(JsError):
        commit({"F": 'AFDate_FormatEx("dd-mmm-yyyy");'}, "あした")
    assert simple_calculate('AFSimple_Calculate("SUM", new Array ("a", "b"));', {"a": "1,000", "b": 2}.get) == 1002


def test_font_runs():
    assert runs_for("Tel 03（1234）") == [("Helv", "Tel 03"), ("TdkJ", "（"), ("Helv", "1234"), ("TdkJ", "）")]
    assert runs_for("김") == [("TdkK", "김")]
    with pytest.raises(FillError):
        runs_for("𠮷")  # 基本多言語面の外の字は、描けないとはっきり断る


# ====================================================================== 組み立てたPDFで、読む・書く・証明・読み戻し
def test_sample_form_window():
    doc = build(PdfDocument(sample_form(), "sample.pdf"))
    pg = doc["pages"][0]
    kinds = {f["label"]: f["kind"] for f in pg["fields"]}
    assert kinds.get("氏名") == "text"
    comb = [f for f in pg["fields"] if f["kind"] == "comb"]
    assert comb and comb[0]["boxes"] == 7 and comb[0]["label"].startswith("郵便番号")
    assert [f["kind"] for f in pg["fields"]].count("checkbox") == 2
    assert [f for f in pg["fields"] if f["kind"] == "inline"], "令和　年　月　日 の空欄"
    text = "\n".join(pg["reading"])
    assert "申請書" in text and "☐［" in text


def test_sample_form_fill_prove_readback():
    blob = sample_form()
    doc = PdfDocument(blob, "sample.pdf")
    win = build(doc)
    fields = win["pages"][0]["fields"]
    vals = {}
    for f in fields:
        if f.get("alternative_to") or f.get("readonly"):
            continue  # 同じ所の別の読み方（両方には書けない）・書けない所
        if f.get("agent") and f["kind"] in ("space", "inline", "text"):
            vals[f["id"]] = "8"  # 規則では決められない所（狭い余白もある）：短い値
        elif f["kind"] == "comb":
            vals[f["id"]] = "1234567"
        elif f["kind"] == "checkbox":
            vals[f["id"]] = f["label"] == "新規"
        elif f["kind"] == "inline":
            vals[f["id"]] = "8"
        else:
            vals[f["id"]] = "とどく 花子"
    res = fill(doc, vals)
    assert res.blob[:len(blob)] == blob  # 元のバイト列はそのまま
    pr = prove(blob, res.blob, set(vals))
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    after = build(PdfDocument(res.blob, "sample.pdf"))
    got = {f["id"]: f["value"] for f in after["pages"][0]["fields"]}
    for k, v in vals.items():
        assert got[k] == v, (k, v, got[k])
    with pytest.raises(FillError):
        fill(doc, {next(f["id"] for f in fields if f["kind"] == "comb"): "123456789"})  # 7字の枠に9字は入らない


def test_sample_form_audit():
    au = audit(sample_form(), "sample.pdf")
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]


# ====================================================================== 公開の様式
@pytest.mark.corpus
@pytest.mark.parametrize("name", ["r9kojinnbanngoukinyuuyousi.pdf", "400747.pdf", "FUNDING-REQUEST-FORM.pdf",
                                  "tamen_siharai-92.pdf"])
def test_corpus_audit(name):
    path = next((p for p in CORPUS if Path(p).name == name), None)
    if path is None:
        pytest.skip(name)
    au = audit(Path(path).read_bytes(), name)
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]


def test_anchored_write_beside_a_label():
    # 窓に欄として示していない所：見出しの字を目印に、その右（または下）の空いている所に書く。証明は、同じ決め方で位置を照らす
    from todoku.pdf.fill import FillError

    blob = sample_form()
    doc = PdfDocument(blob, "s.pdf")
    spec = {"after": "申請書", "value": "（写し）"}
    res = fill(doc, {"@1": spec})
    assert [(p.id, p.value) for p in res.patches] == [("p1.w1", "（写し）")]
    pr = prove(blob, res.blob, {"p1.w1"}, {"p1.w1": spec})
    assert pr.ok, pr.unexpected
    # 位置が目印から決まる所と違えば、証明が見つける
    bad = prove(blob, res.blob, {"p1.w1"}, {"p1.w1": {"after": "氏名", "value": "（写し）"}})
    assert not bad.ok
    with pytest.raises(FillError, match="見つかりません"):
        fill(PdfDocument(blob, "s.pdf"), {"@1": {"after": "存在しない見出し", "value": "x"}})
    # 1行だけのマスの見出しの下には空きがない：書かずに理由を返す
    with pytest.raises(FillError, match="書く空きがありません"):
        fill(PdfDocument(blob, "s.pdf"), {"@1": {"after": "郵便番号", "value": "100-0001", "where": "below"}})


def test_cover_write_over_a_placeholder():
    # 伏せ字（「○○市長　殿」の○○）：規則では決められない所として窓に出し、書くときは白い下地の欄で覆う（元の字は消さない）
    from todoku.pdf.fill import FillError
    from todoku.pdf.window import build

    font, extra, codes = jp_font_objs("○市長殿", 5)
    c = b"BT /F1 12 Tf 60 700 Td " + jp_text(codes, "○○市長") + b" Tj 80 0 Td " + jp_text(codes, "殿") + b" Tj ET\n"
    blob = make_pdf(c, {"F1": font}, extra)
    win = build(PdfDocument(blob, "c.pdf"))
    ph = [f for pg in win["pages"] for f in pg["fields"] if f.get("cover")]
    assert len(ph) == 1 and ph[0]["covered"] == "○○" and ph[0]["agent"], ph
    assert any(f"［{ph[0]['id']}？＝○○］" in ln for pg in win["pages"] for ln in pg["reading"])
    for values in ({ph[0]["id"]: "横浜"}, {"@1": {"cover": "○○", "value": "横浜"}}):
        res = fill(PdfDocument(blob, "c.pdf"), values)
        assert any("白い下地の欄で覆いました" in w for w in res.warnings), res.warnings
        anchored = {x.id: x.anchor for x in res.patches if x.anchor}
        pr = prove(blob, res.blob, {x.id for x in res.patches}, anchored)
        assert pr.ok, pr.unexpected
    with pytest.raises(FillError, match="見つかりません"):
        fill(PdfDocument(blob, "c.pdf"), {"@1": {"cover": "△△", "value": "x"}})


# ====================================================================== 規則では決められない所（エージェントが判断する）
def test_pdf_blanks_tell_structural_from_keyword():
    # 字の形だけで決まる空欄（下線の字・かっこの中の空白）と、語（年・月・市・区長）を手がかりにした空欄を分ける。
    # 2つの数の組としてはこれまでどおり（Word の呼び出しはそのまま）
    (a, b), = gap_blanks("電話（　　　　）")
    assert (a, b) == (3, 7) and gap_blanks("電話（　　　　）")[0].structural
    assert gap_blanks("氏名＿＿＿＿")[0].structural
    date = gap_blanks("令和　　年　　月　　日")
    assert len(date) == 3 and not any(x.structural for x in date)
    assert set(date) == {(2, 4), (5, 7), (8, 10)}


def test_ambiguous_kinds_and_alternatives_are_offered_not_decided():
    # 7つの同じ大きさのマス：1字ずつ書く欄か1つの欄かは形の大きさでは決めない（並び全体の所と、マス 1 つずつの所の両方を示す）。
    # 並びとマスは重なる（alternative_to）：両方には書けない。「氏名」の右の空き（行末の後）と、すぐ右の空のマスは重ならないので、
    # どちらに書くかの見込み（これまでの「値はそちらに書くことが多い」）は付けない
    blob = sample_form()
    doc = PdfDocument(blob, "s.pdf")
    fields = build(doc)["pages"][0]["fields"]
    comb = next(f for f in fields if f["kind"] == "comb")
    assert comb["agent"] and comb["kind_candidates"] == ["comb", "text"] and comb["facts"]["cells"] == 7
    assert len(comb["cell_ids"]) == 7 and comb["facts"]["left"]["text"] == "郵便番号"
    name = next(f for f in fields if f["label"] == "氏名" and f.get("shape") == "cell")
    assert not name.get("agent") and name["facts"]["left"] == {"text": "氏名", "mm": 0.0}
    trail = next(f for f in fields if f.get("shape") == "trail" and f["facts"]["line"] == "氏名")
    assert trail["agent"] and not trail.get("alternative_to") and not name.get("alternative_to")
    cell = next(f for f in fields if f["id"] == comb["cell_ids"][0])
    assert cell["alternative_to"] == comb["id"]
    with pytest.raises(FillError, match="どちらか1つ"):
        fill(PdfDocument(blob, "s.pdf"), {comb["id"]: "1000001", cell["id"]: "1"})
    # 1字ずつ（既定）
    res = fill(PdfDocument(blob, "s.pdf"), {comb["id"]: "1000001"})
    assert prove(blob, res.blob, {comb["id"]}).ok
    got = {f["id"]: f["value"] for f in build(PdfDocument(res.blob, "s.pdf"))["pages"][0]["fields"]}
    assert got[comb["id"]] == "1000001"
    # 1つの欄として（{"kind": "text"}）：マスの仕切りの線が字にかかるので、位置を添えて知らせる（placement。断らない）
    res = fill(PdfDocument(blob, "s.pdf"), {comb["id"]: {"kind": "text", "value": "100-0001"}})
    assert any("罫線がかかる見込み" in w for w in res.warnings), res.warnings
    with pytest.raises(FillError, match="としては書けません"):
        fill(PdfDocument(blob, "s.pdf"), {comb["id"]: {"kind": "checkbox", "value": True}})
    # 日付の空欄は字の間の空き：描かれた空欄ではないので「？」
    assert all(f.get("agent") for f in fields if f["kind"] == "inline")
    # 行の終わりの後ろの 1 字ぶん以上の空き（ページの字の端まで）も所。字を目印にしても書ける
    win = build(doc)
    end = next(f for f in fields if f.get("shape") == "trail" and f["facts"]["line"].startswith("令和"))
    assert end["facts"]["bound"]["kind"] == "page" and end["facts"]["width_em"] >= 1
    assert any(ln.endswith(f"日［{end['id']}？］") for ln in win["pages"][0]["reading"]), win["pages"][0]["reading"]
    # 字のあるマスの、字の下の 1 字ぶん以上の空きは、欄にせず mm の印で示す
    assert "［↓余白 12.0 mm］" in win["notation"] and win["pages"][0]["room_marks"] >= 1
    res = fill(PdfDocument(blob, "s.pdf"), {"@1": {"after": "日", "value": "8", "page": 1}})
    assert prove(blob, res.blob, {"p1.w1"}, {"p1.w1": {"after": "日", "value": "8", "page": 1}}).ok


def test_small_square_cell_can_be_checked_or_written():
    # 小さな正方形のマス（右に選択肢の字）：チェックの印か、1字を書く欄か。true/false は印、字は文字として書く
    font, extra, codes = jp_font_objs("有無", 5)
    c = b"0.5 w 60 700 m 72 700 l S 60 688 m 72 688 l S 60 700 m 60 688 l S 72 700 m 72 688 l S\n"
    c += b"BT /F1 10 Tf 76 690 Td " + jp_text(codes, "有") + b" Tj ET\n"
    blob = make_pdf(c, {"F1": font}, extra)
    # 空のマスはどれも同じ種類の所（字か印を書く）：大きさで checkbox と決めない。右の字は facts に（見出しは右の字）
    f = next(x for x in build(PdfDocument(blob, "b.pdf"))["pages"][0]["fields"] if x.get("shape") == "cell")
    assert f["kind_candidates"] == ["text", "checkbox"] and f["label"] == "有" and f["facts"]["right"]["text"] == "有"
    assert f["facts"]["size_mm"] == [4.2, 4.2]
    for value, want in ((True, True), ("1", "1")):
        res = fill(PdfDocument(blob, "b.pdf"), {f["id"]: value})
        got = {x["id"]: x["value"] for x in build(PdfDocument(res.blob, "b.pdf"))["pages"][0]["fields"]}
        assert got[f["id"]] == want
    assert audit(blob, "b.pdf").ok


def test_printed_checked_mark_is_shown_as_a_candidate():
    # 印刷された記入済みの印（■）：選んだ印か箇条の印かは決めない。消せないので書けない所として、元の字のまま示す
    font, extra, codes = jp_font_objs("■有", 5)
    c = b"BT /F1 10 Tf 60 700 Td " + jp_text(codes, "■有") + b" Tj ET\n"
    blob = make_pdf(c, {"F1": font}, extra)
    pg = build(PdfDocument(blob, "m.pdf"))["pages"][0]
    f = next(x for x in pg["fields"] if x["kind"] == "checkbox")
    assert f["agent"] and f["readonly"] and f["value"] is True and f["facts"]["char"] == "■"
    assert any(f"■［{f['id']}？］有" in ln for ln in pg["reading"]), pg["reading"]
    assert audit(blob, "m.pdf").ok


def test_special_format_with_a_mask_string_is_not_a_number():
    """AFSpecial_* の引数は番号とは限らない（'+65 9999 9999'）：開けて、形をそのまま伝える。"""
    from todoku.pdf import acrojs

    assert acrojs.special_kind([]) == 0 and acrojs.special_kind([2]) == 2 and acrojs.special_kind(["3"]) == 3
    assert acrojs.special_kind(["+65 9999 9999"]) is None and acrojs.special_kind([7]) is None and acrojs.special_kind([1.5]) is None
    assert acrojs.commit_pair({"K": "AFSpecial_Keystroke('+65 9999 9999');"}, "+65 1234 5678")[1] == "+65 1234 5678"


def test_window_opens_field_with_a_mask_string_format():
    import pdfmini
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.window import build

    for call, want in (("AFSpecial_Format('+65 9999 9999');", "書類の決めた形（+65 9999 9999）"),
                       ("AFSpecial_Keystroke(7);", "書類の決めた形（7）"), ("AFSpecial_Format(2);", "電話番号")):
        w = ("<< /Type /Annot /Subtype /Widget /FT /Tx /T (tel) /Rect [100 700 300 716] /F 4 /DA (/Helv 12 Tf 0 g) "
             f"/AA << /F << /S /JavaScript /JS ({call}) >> >> >>")
        blob = pdfmini.build("", "<< >>", [w], annots="/Annots [2 0 R]", acroform="/AcroForm << /Fields [2 0 R] >>")
        f = build(PdfDocument(blob))["pages"][0]["fields"][0]
        assert f["format"] == want
