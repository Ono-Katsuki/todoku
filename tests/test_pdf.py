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
from todoku.pdf.layout import pdf_blanks
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
                                     ("　　　都　道　　　　　市　区　　　", 3), ("付表第一号（一）　　　", 0)])
def test_pdf_blanks(text, n):
    lead = len(text) - len(text.lstrip("　")) if text.startswith("　") else 0
    trail = len(text) - len(text.rstrip("　"))
    assert len(pdf_blanks(text, lead, trail)) == n


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
        if f["kind"] == "comb":
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
@pytest.mark.parametrize("name", ["r9kojinnbanngoukinyuuyousi.pdf", "400747.pdf", "FUNDING-REQUEST-FORM.pdf",
                                  "tamen_siharai-92.pdf"])
def test_corpus_audit(name):
    path = next((p for p in CORPUS if Path(p).name == name), None)
    if path is None:
        pytest.skip(name)
    au = audit(Path(path).read_bytes(), name)
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]
