"""テストの共通の設定：印（marker）と、飛ばしたテストの数の報告。

印：
  corpus  集めた公開の様式・仕様書（corpus/。再配布しないので git に入れない）を使うテスト。
          corpus/ がないとき（CI・ほかのパソコン）は、ここでまとめて飛ばす。TODOKU_NO_CORPUS=1 でも飛ばす
          （corpus のあるパソコンで、CI と同じ条件を試すとき）。
  fonts   パソコンにある書体（組み立て（シェーピング）・埋め込み）や hb-shape（HarfBuzz）を使うテスト。
          なければ、そのテストの中で理由をつけて飛ばす。

最後に「飛ばしたテスト」を印ごとに数えて出す（GITHUB_STEP_SUMMARY があれば、そこにも表で書く）。
"""

from __future__ import annotations

import os
from collections import Counter
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
CORPUS = ROOT / "corpus"


def corpus_available() -> bool:
    if os.environ.get("TODOKU_NO_CORPUS"):
        return False
    try:
        return CORPUS.is_dir() and any(CORPUS.iterdir())
    except OSError:
        return False


def pytest_configure(config):
    config.addinivalue_line("markers", "corpus: 集めた公開の様式・仕様書（corpus/、git に入れない）を使う。なければ飛ばす")
    config.addinivalue_line("markers", "fonts: パソコンの書体・hb-shape を使う。なければ理由をつけて飛ばす")


def pytest_collection_modifyitems(config, items):
    if corpus_available():
        return
    why = ("corpus/ がありません（集めた公開の様式・仕様書は再配布しないので git に入れていません。"
           "scripts/fetch_*.py・collect_*.py で取得できます）")
    skip = pytest.mark.skip(reason=why)
    for item in items:
        if item.get_closest_marker("corpus") is not None:
            item.add_marker(skip)


@pytest.fixture
def embedding_font():
    """PDF に字形を埋め込める書体（Arial・DejaVu Sans など。TODOKU_PDF_FONT で指せる）。なければ、理由をつけて飛ばす。"""
    from todoku.pdf.embedfont import system_font

    f = system_font()
    if f is None:
        pytest.skip("このパソコンに、PDF に埋め込める TrueType の書体（Arial・DejaVu Sans など）がない")
    return f


def _category(rep) -> str:
    kw = getattr(rep, "keywords", {}) or {}
    if "corpus" in kw:
        return "corpus"
    if "fonts" in kw:
        return "fonts"
    return "other"


def pytest_terminal_summary(terminalreporter, exitstatus, config):
    stats = terminalreporter.stats
    skipped = Counter(_category(r) for r in stats.get("skipped", []))
    counts = {k: len(stats.get(k, [])) for k in ("passed", "failed", "error", "xfailed", "xpassed")}
    line = (f"とどく：通った {counts['passed']}・失敗 {counts['failed']}・エラー {counts['error']}・"
            f"予期した失敗 {counts['xfailed']}・"
            f"飛ばした {sum(skipped.values())}（corpus {skipped['corpus']}・書体 {skipped['fonts']}・そのほか {skipped['other']}）"
            f"・corpus/ {'あり' if corpus_available() else 'なし'}")
    terminalreporter.write_sep("-", line)
    out = os.environ.get("GITHUB_STEP_SUMMARY")
    if out:
        import platform
        import sys

        rows = [("通った", counts["passed"]), ("失敗", counts["failed"]), ("エラー", counts["error"]),
                ("予期した失敗（xfail）", counts["xfailed"]), ("飛ばした：corpus（再配布しない様式）", skipped["corpus"]),
                ("飛ばした：書体・hb-shape がない", skipped["fonts"]), ("飛ばした：そのほか", skipped["other"])]
        md = [f"### pytest（{platform.system()}・Python {sys.version.split()[0]}）", "", "| | 数 |", "|---|---:|"]
        md += [f"| {k} | {v} |" for k, v in rows]
        fails = [r.nodeid for r in stats.get("failed", []) + stats.get("error", [])][:30]
        if fails:
            md += ["", "失敗したテスト：", ""] + [f"- `{n}`" for n in fails]
        try:
            with open(out, "a", encoding="utf-8") as fh:
                fh.write("\n".join(md) + "\n\n")
        except OSError:
            pass
