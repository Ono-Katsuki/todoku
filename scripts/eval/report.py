#!/usr/bin/env python3
"""全部の結果を 1 つにまとめる：開発の書類ととっておいた書類を横に並べ、偏りと限りを書く。使い方: report.py --out 出力 [--dry-run]"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import load_json, md_table, pct  # noqa: E402
from read_compare import CAT_JA, CATS  # noqa: E402

SET_JA = {"dev": "開発に使った書類", "tama": "とっておいた書類（都・多摩）", "world": "とっておいた書類（世界）"}


def read_side_by_side(out: Path, sets: list[str]) -> str:
    """特徴ごとに、とどくと、形式ごとにいちばんよい比べる道具（数えた全部の形式の合計）を、書類の組ごとに横に並べる。"""
    data = {s: load_json(out / "read" / s / "summary.json") for s in sets}
    sets = [s for s in sets if data[s]]
    if not sets:
        return "（読む比べの結果がない）\n"
    rows = []

    def agg(s, cat):
        t_ok = t_n = b_ok = b_n = 0
        for kd in data[s]["by_kind"].values():
            tools = kd["tools"]
            t = tools.get("todoku", {}).get("cat", {}).get(cat)
            base = [v["cat"][cat] for k, v in tools.items() if k != "todoku" and cat in v["cat"]]
            n = None
            if t:
                t_ok += t["ok"]
                n = t["n"]
                t_n += n
            if base:
                best = max(base, key=lambda x: x["ok"])
                b_ok += best["ok"]
                b_n += n if n is not None else best["n"]
            elif n is not None:
                b_n += n
        return t_ok, t_n, b_ok, b_n

    for c in CATS:
        row = [CAT_JA[c]]
        any_n = False
        for s in sets:
            t_ok, t_n, b_ok, b_n = agg(s, c)
            any_n = any_n or t_n or b_n
            row += [pct(t_ok, t_n), pct(b_ok, b_n)]
        if any_n:
            rows.append(row)
    ex = ["正確に読めた（字がある書類）"]
    op = ["開けた"]
    for s in sets:
        te = tn = be = 0
        to = tt = bo = 0
        for kd in data[s]["by_kind"].values():
            tools = kd["tools"]
            x = tools.get("todoku")
            if x:
                te += x["exact_read"]
                tn += x["docs_with_text"]
                to += x["opened"]
                tt += x["docs"]
            others = [v for k, v in tools.items() if k != "todoku"]
            if others:
                be += max(v["exact_read"] for v in others)
                bo += max(v["opened"] for v in others)
        ex += [pct(te, tn), pct(be, tn)]
        op += [pct(to, tt), pct(bo, tt)]
    head = [""]
    for s in sets:
        head += [f"{SET_JA.get(s, s)}：とどく", "比べる道具（形式ごとの最良）"]
    return md_table(head, [op, ex, *rows]) + "\n"


def write_side_by_side(out: Path, sets: list[str]) -> str:
    data = {s: load_json(out / "write" / s / "summary.json") for s in sets}
    sets = [s for s in sets if data[s]]
    groups = sorted({g for s in sets for g in data[s]["groups"]})
    rows = []
    for g in groups:
        row = [g]
        for s in sets:
            d = data[s]["groups"].get(g)
            if not d:
                row += ["—", "—"]
                continue
            t = d["tools"].get("todoku:ja") or {}
            bases = {k: v for k, v in d["tools"].items() if k.endswith(":ja") and not k.startswith("todoku")}
            best = max(bases.items(), key=lambda kv: kv[1]["write_ok"]) if bases else None
            row += [pct(t.get("write_ok", 0), t.get("tried", 0)), f"{best[0].split(':')[0]} {pct(best[1]['write_ok'], best[1]['tried'])}" if best else "書けない"]
        rows.append(row)
    head = ["書けた＋何も消えない（日本語の値）"]
    for s in sets:
        head += [f"{SET_JA.get(s, s)}：とどく", "比べる道具（最良）"]
    return md_table(head, rows) + "\n" if rows else "（書く比べの結果がない）\n"


def processing_side_by_side(out: Path, sets: list[str]) -> str:
    data = {s: load_json(out / "processing" / s / "summary.json") for s in sets}
    sets = [s for s in sets if data[s]]
    if not sets:
        return "（処理の結果がない）\n"
    rows = [["数式のある書類"] + [f"{data[s]['docs_with_formulas']} / {data[s]['docs']}" for s in sets],
            ["数式のセル"] + [f"{data[s]['formula_cells']:,}" for s in sets],
            ["計算した値が保存された値と合う"] + [pct(data[s]["stored_match"], data[s]["stored_compared"], 2) for s in sets],
            ["確かめられた規則"] + [pct(data[s]["rules_proven"], data[s]["rules"]) for s in sets],
            ["計算できないと言った関数"] + [", ".join(data[s]["unknown_functions"]) or "なし" for s in sets],
            ["処理の取り出しが失敗した書類"] + [len(data[s]["failed"]) for s in sets]]
    return md_table(["", *[SET_JA.get(s, s) for s in sets]], rows) + "\n比べる道具（openpyxl）は数式の字を読むだけで、計算も、規則へのまとめもしない。\n"


LIMITS = """## 偏りと限り

偏りはどちらの向きにもある。読むときは、数の向きを確かめること。

**とどくに有利なもの**
- 「変えた保存値」の試験は、とどくが得意な形を作ったもの（比べる道具は計算しないので、見分けようがない）。自然に古い値はほとんどない。
- PDF の「とどくの窓を正解にした」字の再現率は、とどく自身が基準なので、とどくに有利（計画の指定。主の数は pdftotext）。
- 開くだけの比べ：比べる道具は Python のライブラリだけ。LibreOffice や Word に変えてから読む人の手順は入れていない（.doc・.odt・.rtf・.ods は MarkItDown が読まないので 0）。
- 図形の字・チェックボックスは、XML を直に読めば誰でも読める。比べたのは、ライブラリのふつうの使い方。

**比べる道具に有利なもの**
- 書く欄はとどくが選び、比べる道具に教えた。欄を見つける力は比べていない。
- PDF の正解の pdftotext は、pypdf・pdfminer と同じ字の層を読むので、字の層の誤りは見つからない（文字認識の列は、Poppler で描いた絵から読むので、字の層によらないが、それ自体が誤る）。
- 比べる道具の読み方は、使える人が書くように書いた（結合を広げる・非表示の印・字の書式など）。ふだんの使い方はこれより弱い。
- 書く比べの英数字の値は、Helvetica の注釈でも書ける形。様式の実際は日本語。
- PDF の 2 字の並びは pdftotext の空白で切って数える（離れた列の字を 1 つの空白でつなぐため。どの道具にも同じ）。
- 窓の説明の字を除くとき、とどくの「［？　］」は元の空白に戻して比べている。とどくの字の比べには、「規則では決められない所」の表の「いまの値」の列も入れる。

**測り方の限り**
- 値の照合は、出力の中に、値の境目つきで同じ字があるかで見る。短い数（「1」など）は偶然に当たりうる。すべての道具に同じように働く。
- 結合が各行に届くかは、同じ行に字が出るかで見る。とどくの窓が別の形で伝えている場合（表の見出しなど）は数えない。読み手のエージェントがどこまで読み取れるかは測っていない。
- とどくの窓が「同じ形の行が続く」とまとめた所は、窓の言明のとおりに行へ広げて数える（字の再現率も同じ。言明が違えば字が合わず、落ち度に出る）。広げない数は「広げない」の行に残す。
- 図形・チェックボックスの .xls は LibreOffice で .xlsx に変えてから調べるので、変換で落ちるものは見えない。
- Excel が動かないときの正解は LibreOffice で、表示の形式が Excel と違うことがある（正解の出どころは書類ごとに記録した）。Word は動かしていない：Word の正解は XML と LibreOffice。
- 記入の「記入済み」は、エージェントの分類・証明・確かめ役の 3 つが揃ったもの。確かめ役も AI で、人の二重の確かめではない。Office の書類の絵は LibreOffice の描画で、Word・Excel の実物ではない。
- 申請者の情報は固定で、そこにない欄は「聞くこと」として空く。記入済みでも、書類の全部が埋まるわけではない。
- 記入の「確かめ役」は厳しめに見る：様式に書かれていない判断（等級の語に○を付けるべきかなど）も「答えられるのに空いていた」に数えることがあり、記入済みの数を下げる向きに働く。
- Office の絵の日本語は Noto Sans JP で描く（MS の書体はない）。字の幅が実物と違うので、はみ出しの判断は参考。
- 世界の数は Common Crawl の中の推定で、下限。枠の推定の誤差は区間に入っていない。「新しく届く」は層ごとに少数の標本で、国ごとの区間は広い。
- 調査の一部の比べは Excel を使わない（LibreOffice）。書類が大きいと、とどくの段階の時間切れ・メモリの上限で「段階なし」になり、分母から除く（数は調査の結果にある）。
"""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    out = Path(a.out)
    prov = load_json(out / "provenance.json") or {}
    sets = [s for s in ("dev", "tama", "world") if (out / "read" / s).exists() or (out / "fill" / s).exists()]
    L = ["# とどくの最終の評価" + ("（試運転：開発の書類だけ。とっておいた書類には触れていない）" if a.dry_run else ""), "",
         f"版：{prov.get('tag')}（{prov.get('commit')}）。動かした日：{prov.get('started')}。道具の版は provenance.json。", ""]
    if a.dry_run:
        L += ["**ここの数はすべて開発の書類（dev）。最終の結果ではない。**", ""]
    if not prov.get("pdfkit_tool"):
        L += ["**注意：PDFKit の道具（pkr）なしで動いた。PDF の「値が入った」は文字認識でなく画素の変化で数え（pypdf の日本語は描けないものとして数えた）、確かめ役は PDFKit の絵を見ていない。**", ""]
    L += ["## 読む（開発の書類ととっておいた書類を並べて）", "", read_side_by_side(out, sets),
          "形式ごとの細かい表：read/<組>/summary.md。「比べる道具」の列は、特徴ごと・形式ごとに、いちばんよい道具を選んで足したもの（1 つの道具の成績ではない。比べる道具に有利）。", "",
          "## 書く", "", write_side_by_side(out, sets), "消えたものの内訳：write/<組>/summary.md。", "",
          "## 処理の取り出し", "", processing_side_by_side(out, sets), ""]
    for s in sets:
        e = out / "explain" / s / "summary.md"
        if e.exists():
            L += [f"（{SET_JA.get(s, s)}）", e.read_text(encoding="utf-8"), ""]
    for s in sets:
        n = load_json(out / "read" / s / "summary.json")
        if n and n.get("excel_notice"):
            L += [f"注意（{s}）：{n['excel_notice']}", ""]
    fill = out / "fill" / "summary.md"
    L += [fill.read_text(encoding="utf-8") if fill.exists() else "## 記入\n\n（結果がない）\n", ""]
    cnt = out / "world" / "counts.md"
    L += [cnt.read_text(encoding="utf-8") if cnt.exists() else "## 世界地図の数\n\n（結果がない）\n", "", LIMITS]
    (out / "report.md").write_text("\n".join(str(x) for x in L), encoding="utf-8")
    print(f"→ {out / 'report.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
