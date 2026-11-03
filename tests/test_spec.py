"""とどくの中の表を、ファイル形式の公式の仕様書（原本）と照らす。仕様書がなければ飛ばす（scripts/fetch_specs.py で取得）。

仕様書はとどく自身で文字に直したもの（corpus/specs/MS-DOC.txt）を使う。
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from todoku.doc.book import FC_NAMES
from todoku.doc.fill import CACHES, SHIFT, SUB

SPEC = Path(__file__).parent.parent / "corpus" / "specs" / "MS-DOC.txt"


@pytest.fixture(scope="module")
def spec() -> str:
    if not SPEC.exists():
        pytest.skip("MS-DOC の仕様書がありません（uv run python scripts/fetch_specs.py）")
    return SPEC.read_text(encoding="utf-8")


@pytest.mark.corpus
def test_fib_field_order(spec):
    names = []
    for m in re.finditer(r"\n2\.5\.\d+ FibRgFcLcb\w+", spec):
        end = spec.find("\n2.5.", m.start() + 10)
        names += re.findall(r"\n(fc\w+|dwLowDateTime|dwHighDateTime) \(4 bytes\):", spec[m.start():end])
    norm = ["dwLowDateTime_dwHighDateTime" if n == "dwLowDateTime" else n[2:] for n in names if n != "dwHighDateTime"]
    assert [n.lower() for n in norm] == [n.split("_2")[0].lower() if not n.startswith("dw") else n.lower() for n in FC_NAMES]


@pytest.mark.corpus
def test_dop_protection_bits(spec):
    i = spec.find("\n2.7.2 DopBase")
    seg = spec[i:spec.find("\n2.7.3 ", i)]
    pos, bits = 0, {}
    for name, n, unit in re.findall(r"\n(?:[A-Za-z] - )?(\w+) \((\d+) (bits?|bytes?)\):", seg):
        bits[name] = pos
        pos += int(n) * (8 if unit.startswith("byte") else 1)
    # doc/book.py の _dop：fProtEnabled = 7バイト目の 0x02、fLockRev = 7バイト目の 0x40、fLockAtn = 6バイト目の 0x10
    assert (bits["fProtEnabled"] // 8, 1 << bits["fProtEnabled"] % 8) == (7, 0x02)
    assert (bits["fLockRev"] // 8, 1 << bits["fLockRev"] % 8) == (7, 0x40)
    assert (bits["fLockAtn"] // 8, 1 << bits["fLockAtn"] % 8) == (6, 0x10)


@pytest.mark.corpus
def test_sprm_codes(spec):
    lines = spec.split("\n")
    vals = {}
    for i, ln in enumerate(lines):
        if re.fullmatch(r"sprm\w+", ln.strip()):
            for k in range(i + 1, min(i + 8, len(lines))):
                m = re.search(r"\((0x[0-9A-Fa-f]{4})\)", lines[k])
                if m:
                    vals.setdefault(ln.strip(), int(m[1], 16))
                    break
    used = {"sprmPFInTable": 0x2416, "sprmPFTtp": 0x2417, "sprmPItap": 0x6649, "sprmPFInnerTableCell": 0x244B,
            "sprmPFInnerTtp": 0x244C, "sprmCFVanish": 0x083C, "sprmCFRMarkDel": 0x0800, "sprmCKul": 0x2A3E,
            "sprmCPicLocation": 0x6A03, "sprmCSymbol": 0x6A09, "sprmCFSpec": 0x0855, "sprmCFData": 0x0806,
            "sprmTDefTable": 0xD608, "sprmPChgTabs": 0xC615}
    assert {k: vals.get(k) for k in used} == used


@pytest.mark.corpus
def test_shifted_tables_sizes(spec):
    # 各 PLC の要素の大きさ（原本の 2.8 節）と、とどくがずらす表の大きさ
    def plc(name):
        m = re.search(r"\n2\.8\.\d+ " + re.escape(name) + r"\n", spec)
        return re.sub(r"\s+", " ", spec[m.end():m.end() + 1500])
    assert "FBKFD" in plc("Plcfbkfd") and "(6 bytes each)" in plc("Plcfbkfd") and SHIFT["PlcfBkfFcc"] == 6
    assert "(10 bytes each)" in plc("Plcbkfd") and SHIFT["PlcfBkfSdt"] == 10
    assert "SPAs (26 bytes each)" in plc("PlcfSpa") and SHIFT["PlcSpaMom"] == 26
    assert "12-byte Sed" in plc("PlcfSed") and SHIFT["PlcfSed"] == 12
    assert "ATRDPre10 structures (30 bytes each)" in plc("PlcfandRef") and SHIFT["PlcfandRef"] == 30
    assert "(22-bytes each)" in plc("PlcftxbxTxt") and dict(SUB["txbx"][1])["PlcftxbxTxt"] == 22
    assert "6-byte Tbkd" in plc("PlcfTxbxBkd") and dict(SUB["txbx"][1])["PlcfTxbxBkd"] == 6
    for name, size in (("Plcfspl", 2), ("Plcfgram", 2), ("Plcflad", 2), ("PlcfTch", 4), ("Plcfcookie", 10), ("PlcfWKB", 12)):
        assert f"({size} bytes" in plc(name) or f"{size} bytes each" in plc(name), name


@pytest.mark.corpus
def test_cleared_tables_are_deprecated_caches(spec):
    # 空にするのは、原本で「出力しない・読まない」とされた控えだけ
    for n in CACHES:
        m = re.search(r"\nfc" + re.escape(n) + r" \(4 bytes\):(.{0,500})", spec, re.S)
        assert m, n
        desc = re.sub(r"\s+", " ", m.group(1))
        assert "SHOULD NOT" in desc or "MUST be 0" in desc or "ignored" in desc, n
