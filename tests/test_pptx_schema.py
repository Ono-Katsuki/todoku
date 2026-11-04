"""PowerPoint の書き込みが作りを壊さないか：OOXML の形の確かめ（スキーマの求める並びを手で書いた小さな検査。Office は使わない）。

第 8 回で、手作りの試しのファイル（書く前）が PowerPoint に修復を求められた。部品を比べると、書いて変わるのはスライドの部品の
a:p の中だけで、足りなかったのは試しのファイルの方（内容の型・notesSz・clrMap・テーマ）だった。"""

from __future__ import annotations

import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import pptx_maker as pm  # noqa: E402

from todoku import api  # noqa: E402

CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def package_problems(path) -> list[str]:
    """OPC と PresentationML の、修復を招く形の欠け。"""
    z = zipfile.ZipFile(path)
    names = z.namelist()
    out: list[str] = []
    types = ET.fromstring(z.read("[Content_Types].xml"))
    defaults = {t.get("Extension").lower() for t in types.findall(CT + "Default")}
    over = {t.get("PartName"): t.get("ContentType") for t in types.findall(CT + "Override")}
    for n in names:
        if n == "[Content_Types].xml":
            continue
        if "/" + n not in over and n.rsplit(".", 1)[-1].lower() not in defaults:
            out.append(f"内容の型がない：{n}")
    for part, want in (("ppt/slideMasters/", "slideMaster+xml"), ("ppt/slideLayouts/", "slideLayout+xml"), ("ppt/slides/", "slide+xml"),
                       ("ppt/notesSlides/", "notesSlide+xml"), ("ppt/notesMasters/", "notesMaster+xml"), ("ppt/theme/", "theme+xml")):
        for n in names:
            if n.startswith(part) and n.endswith(".xml") and not (over.get("/" + n) or "").endswith(want):
                out.append(f"内容の型が違う：{n}")
    for n in names:
        if not n.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(n))
        for r in ET.fromstring(z.read(n)).findall(PR + "Relationship"):
            if r.get("TargetMode") == "External":
                continue
            t = r.get("Target")
            t = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base, t))
            if t not in names:
                out.append(f"関係の先がない：{n} → {t}")
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    order = [c.tag.replace(P, "") for c in pres]
    for need in ("sldSz", "notesSz"):
        if need not in order:
            out.append(f"presentation に {need} がない")
    for n in names:
        if n.startswith("ppt/slideMasters/") and n.endswith(".xml"):
            kids = [c.tag.replace(P, "") for c in ET.fromstring(z.read(n))]
            if kids[:2] != ["cSld", "clrMap"]:
                out.append(f"{n}：cSld・clrMap の順でない（{kids}）")
            rels = z.read(n.replace("slideMasters/", "slideMasters/_rels/") + ".rels").decode()
            if "/theme" not in rels:
                out.append(f"{n}：テーマがない")
    for n in names:
        if n.startswith(("ppt/slides/", "ppt/notesSlides/")) and n.endswith(".xml"):
            for p in ET.fromstring(z.read(n)).iter(A + "p"):
                kids = [c.tag.replace(A, "") for c in p]
                body = [k for k in kids if k not in ("pPr", "endParaRPr")]
                if kids != (["pPr"] if "pPr" in kids else []) + body + (["endParaRPr"] if "endParaRPr" in kids else []):
                    out.append(f"{n}：a:p の子の順が違う（{kids}）")
                if any(k not in ("r", "br", "fld") for k in body):
                    out.append(f"{n}：a:p に知らない子（{kids}）")
                for r in p.findall(A + "r"):
                    rk = [c.tag.replace(A, "") for c in r]
                    if rk not in (["t"], ["rPr", "t"]):
                        out.append(f"{n}：a:r の子が違う（{rk}）")
    return out


def _fixture(tmp_path):
    p = tmp_path / "b.pptx"
    p.write_bytes(pm.make([pm.sp(pm.para("氏名：　　　　")) + pm.table([["住所", ""], ["電話", ""]])],
                          notes={1: pm.sp(pm.para("メモ：　　　"))}))
    return p


def test_maker_makes_a_valid_package(tmp_path):
    assert package_problems(_fixture(tmp_path)) == []


def test_write_changes_only_the_slide_and_keeps_the_schema_order(tmp_path):
    src = _fixture(tmp_path)
    fl = api.fields(str(src))
    cell = next(f["id"] for f in fl if ".R1." in f["id"] and not f.get("value"))
    out = tmp_path / "o.pptx"
    r = api.fill(str(src), {cell: "東京都〇〇", "@1": {"after": "氏名：", "value": "山田 花子"}}, str(out), overwrite=True)
    assert r["proof_ok"] and r["read_back_ok"]
    assert package_problems(out) == []
    a, b = zipfile.ZipFile(src), zipfile.ZipFile(out)
    assert a.namelist() == b.namelist()
    assert [n for n in a.namelist() if a.read(n) != b.read(n)] == ["ppt/slides/slide1.xml"]
