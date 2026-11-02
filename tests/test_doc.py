"""古い Word（.doc）の読み書きのテスト。公開の様式（corpus/wide_doc、scripts/collect_forms.py --ext doc で取得）で確かめる。"""

from __future__ import annotations

import glob
from pathlib import Path

import pytest

from todoku.book import kind_of
from todoku.doc.audit import run as audit
from todoku.doc.book import DocFile, sprms
from todoku.doc.fill import fill
from todoku.docx.window import build

CORPUS = sorted(glob.glob(str(Path(__file__).parent.parent / "corpus" / "wide_doc" / "**" / "*.doc"), recursive=True))


def test_sprm_sizes():
    # sprmPFInTable（1バイト）、sprmPItap（4バイト）、可変長
    g = bytes.fromhex("1624" "01") + bytes.fromhex("4966" "02000000") + bytes.fromhex("08d6" "0300" "aabb")
    out = list(sprms(g))
    assert out[0] == (0x2416, b"\x01") and out[1] == (0x6649, b"\x02\x00\x00\x00") and out[2][0] == 0xD608


@pytest.mark.parametrize("path", CORPUS, ids=lambda p: Path(p).name)
def test_doc_corpus_audit(path):
    blob = Path(path).read_bytes()
    assert kind_of(blob) == "doc"
    au = audit(blob, Path(path).name)
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]


def test_doc_header_field_and_bookmarks():
    path = next((p for p in CORPUS if Path(p).name == "mynumber.doc"), None)
    if path is None:
        pytest.skip("mynumber.doc")
    doc = DocFile(Path(path).read_bytes(), "m.doc")
    win = build(doc)
    header = [f["id"] for f in win["fields"] if f["context"].startswith("様式第１６号")]
    body = next(f["id"] for f in win["fields"] if f["id"].startswith("T"))
    res = fill(doc, {header[0]: "8", body: "とどく"})
    after = DocFile(res.blob, "m.doc")
    assert any(n.startswith("_todoku_") for n, _a, _b in after.bookmarks)  # 本文の欄はしおりで囲む
    got = {f["id"]: f["value"] for f in build(after)["fields"]}
    assert got[body] == "とどく"
    assert any(p.text.startswith("様式第１６号") and "8年" in p.text for _k, _p, bl in after.stories for p in bl if hasattr(p, "text"))
