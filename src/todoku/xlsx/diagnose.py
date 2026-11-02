"""書類の作り手向けの診断：4段階のラダー（届く・読める・操作できる・検証できる）で壁を示す。

変換と検証の過程で「自動でうまくいかなかった箇所」を、そのまま診断として返す。
ここにあるのはすべて決定的な処理による判定。
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import asdict, dataclass, field

from ..formula.parser import addr
from ..formula.values import is_number
from .calc import VerifyReport, used_functions, verify
from .window import input_cells
from .workbook import Workbook

LEVELS = {0: "届く", 1: "読める", 2: "操作できる", 3: "検証できる"}
SEVERITY = {"blocker": "壁", "agent": "規則では越えられない（エージェントなら越えられる）", "warning": "注意", "info": "参考"}


@dataclass
class Finding:
    level: int
    severity: str  # blocker / warning / info
    code: str
    title: str
    detail: str
    fix: str
    where: list[str] = field(default_factory=list)
    count: int = 1

    def to_dict(self) -> dict:
        d = asdict(self)
        d["level_name"] = LEVELS[self.level]
        d["severity_name"] = SEVERITY[self.severity]
        return d


@dataclass
class Diagnosis:
    file: str
    findings: list[Finding]
    verify: dict
    stats: dict

    def score(self) -> dict:
        """段階ごとに、壁があるかどうか。最初に壁がある段階が、その書類の到達点。
        規則だけ（AIを使わない決定的な処理）で越えられない所は、別に数える（rules_reached）。"""
        def first(sev: set) -> int:
            for lv in range(4):
                if any(f.level == lv and f.severity in sev for f in self.findings):
                    return lv
            return 4

        reached = first({"blocker"})
        rules = first({"blocker", "agent"})
        name = lambda x: "すべて" if x == 4 else f"{LEVELS[x]}の手前"  # noqa: E731
        return {"reached": reached, "reached_name": name(reached),
                "rules_reached": rules, "rules_reached_name": name(rules),
                "by_level": {LEVELS[lv]: {s: sum(1 for f in self.findings if f.level == lv and f.severity == s)
                                         for s in SEVERITY} for lv in range(4)}}

    def to_dict(self) -> dict:
        return {"file": self.file, "score": self.score(), "stats": self.stats, "verify": self.verify,
                "findings": [f.to_dict() for f in self.findings]}


def _cells(sheet: str, keys, limit: int = 12) -> list[str]:
    out = [f"{sheet}!{addr(r, c)}" for r, c in list(keys)[:limit]]
    return out


_SPACED = re.compile(r"^(?:\S[ 　]+){2,}\S$")


def diagnose(wb: Workbook, rep: VerifyReport | None = None) -> Diagnosis:
    F: list[Finding] = []
    rep = rep or verify(wb)
    total_cells = sum(len(s.cells) for s in wb.sheets)

    # ---------------------------------------------------- 0. 届く（ファイル単体で分かること）
    if wb.has_macros:
        F.append(Finding(0, "warning", "macro_file", "マクロ付きのファイル（xlsm）",
                         "多くの組織や支援技術・AIの環境では、マクロ付きファイルは開く前に止められるか、マクロが無効になります。",
                         "マクロでしかできない処理をなくし、関数と入力規則だけで動く xlsx で配布する。"))
    if wb.external_links:
        F.append(Finding(0, "warning", "external_link", "ほかのファイルへのリンク",
                         f"{len(wb.external_links)} 件の外部ブック参照があります。受け取った人の環境では参照先がなく、値が更新されません。",
                         "参照先の値をこのブックの中に取り込む。", count=len(wb.external_links)))

    for sh in wb.sheets:
        if sh.state.startswith("non-worksheet"):
            continue
        n = sh.name
        # ---------------------------------------------------- 1. 読める
        if sh.state in ("hidden", "veryHidden"):
            F.append(Finding(1, "info", "hidden_sheet", f"非表示のシート「{n}」",
                             "画面では見えないシートです。計算の元データなどが入っていることがあります。",
                             "利用者が読む必要のある情報なら表示する。シート名で役割が分かるようにする。", [n]))
        narrow = [c for c, w in sh.col_widths.items() if w < 4 and c <= max(sh.max_col, 1)]
        big_merges = [m for m in sh.merges if (m[3] - m[1]) >= 3]
        if len(narrow) >= 15 and len(sh.merges) >= 10:
            F.append(Finding(1, "blocker", "grid_paper", "方眼紙状のレイアウト（ネ申エクセル）",
                             f"幅の狭い列が {len(narrow)} 本あり、セル結合が {len(sh.merges)} か所あります。"
                             "1つの項目が複数の列にまたがるため、表の行と列の関係を機械が取り出せません。",
                             "見た目の位置合わせは列幅で行い、1項目＝1セルにする。印刷用の様式と入力用の表を分ける。",
                             [n], count=len(sh.merges)))
        elif len(sh.merges) >= 20 or len(big_merges) >= 10:
            F.append(Finding(1, "warning", "many_merges", "セル結合が多い",
                             f"セル結合が {len(sh.merges)} か所あります。並べ替え・コピー・読み上げの順番が崩れやすくなります。",
                             "結合をやめ「選択範囲内で中央」の配置を使う。", [n], count=len(sh.merges)))
        spaced = [(k, c.value) for k, c in sh.cells.items() if isinstance(c.value, str) and _SPACED.match(c.value.strip()) and not c.has_formula]
        if spaced:
            F.append(Finding(1, "warning", "spaced_label", "字間を空白で調整した見出し",
                             f"「{spaced[0][1]}」のように文字の間に空白を入れています。検索や読み上げで別の言葉になります"
                             "（この窓では空白を詰めて見出しとして扱いました）。",
                             "空白を入れず、配置の「均等割り付け」を使う。", _cells(n, [k for k, _ in spaced]), len(spaced)))
        num_text = [k for k, c in sh.cells.items() if isinstance(c.value, str) and not c.has_formula and c.t in ("s", "inlineStr", "str")
                    and re.fullmatch(r"[-+]?[\d,，]+(\.\d+)?", unicodedata.normalize("NFKC", c.value.strip()) or "x")
                    and not re.fullmatch(r"0\d+", c.value.strip()) and len(c.value.strip()) < 12]
        if len(num_text) >= 3:
            F.append(Finding(1, "warning", "number_as_text", "文字列として入っている数値",
                             f"{len(num_text)} 個のセルで、数値が文字列として入っています。合計や比較で正しく扱われません。",
                             "数値として入力し、桁区切りなどは表示形式で付ける。", _cells(n, num_text), len(num_text)))
        text_shapes = [s for s in sh.shapes if s.kind == "text" and s.text]
        if text_shapes:
            F.append(Finding(1, "warning", "text_in_shapes", "図形（テキストボックス）の中の文章",
                             f"{len(text_shapes)} 個の図形に文章があります。セルの外にあるため、読む順番が決まらず、見落とされます。"
                             f"例：「{text_shapes[0].text[:40]}」",
                             "注意書きや記入要領はセルに書くか、別シートにまとめる。",
                             [f"{n}!{addr(*s.anchor)}" if s.anchor else n for s in text_shapes[:12]], len(text_shapes)))
        pics = [s for s in sh.shapes if s.kind in ("picture", "chart")]
        no_alt = [s for s in pics if not s.alt.strip()]
        if no_alt:
            F.append(Finding(1, "blocker" if len(no_alt) == len(pics) and len(pics) >= 1 and total_cells < 5 else "warning",
                             "no_alt", "代替テキストのない画像・グラフ",
                             f"{len(no_alt)} 個の画像・グラフに代替テキストがありません。スクリーンリーダーやAIには中身が伝わりません。",
                             "画像を右クリック→「代替テキストを編集」で、何が書いてあるかを文章で入れる。",
                             [f"{n}!{addr(*s.anchor)}" if s.anchor else n for s in no_alt[:12]], len(no_alt)))
        if sh.hidden_rows or sh.hidden_cols:
            F.append(Finding(1, "info", "hidden_rowcol", "非表示の行・列",
                             f"非表示の行が {len(sh.hidden_rows)}、列が {len(sh.hidden_cols)} あります。読み上げでは読まれることも、読まれないこともあります。",
                             "計算用の補助列は別シートに移す。", [n]))

        # ---------------------------------------------------- 2. 操作できる
        if sh.has_vml_controls or any(s.kind == "control" for s in sh.shapes):
            F.append(Finding(2, "blocker", "form_controls", "ボタンやチェックボックスなどのフォームコントロール",
                             "マウスで押す前提の部品があります。キーボード・スイッチ・AIからは操作できないことが多く、押した結果も確かめられません。",
                             "チェックはセルの入力規則（リスト：はい／いいえ）に置き換える。ボタンで動く処理は関数にする。", [n]))
        stamp = [k for k, c in sh.cells.items() if isinstance(c.value, str) and (c.value.strip() in ("印", "㊞", "(印)", "（印）") or "押印" in c.value)]
        if stamp:
            F.append(Finding(2, "blocker", "stamp", "押印欄",
                             "押印を求める欄があります。手を動かせない人は自分で押せず、電子での提出もできません。",
                             "押印を廃止するか、署名・電子申請・本人確認の代替手段を明記する。", _cells(n, stamp), len(stamp)))
        visit = [k for k, c in sh.cells.items() if isinstance(c.value, str) and re.search(r"来庁|窓口(まで|へ|に)(お越し|持参|提出)|持参(して|の上)|郵送(のみ|に限)|FAX", c.value)]
        if visit:
            F.append(Finding(2, "warning", "in_person", "来庁・持参・郵送・FAXを求める記述",
                             "本人が出向く、紙で送る前提の記述があります。AIが代わりに提出できません。",
                             "電子での提出方法を用意し、書類の中に案内する。", _cells(n, visit), len(visit)))
        handwrite = [k for k, c in sh.cells.items() if isinstance(c.value, str) and re.search(r"自署|手書き|直筆", c.value)]
        if handwrite:
            F.append(Finding(2, "blocker", "handwrite", "自署・手書きを求める記述",
                             "手で書くことを求めています。手を動かせない人には越えられない壁です。",
                             "自署の代わりに、本人の意思を確認できる電子的な方法を認める。", _cells(n, handwrite), len(handwrite)))
        inputs = input_cells(wb, sh)
        from .window import undecided_cells

        und = undecided_cells(wb, sh, inputs)
        if und:
            cols = sorted({c for c in und.values() if c})
            F.append(Finding(2, "agent", "undecided_inputs", "入力欄の印のない表の列",
                             f"表の行に入力欄がありますが、{len(und)} か所の空のセル（列「{'」「'.join(cols[:5]) or '見出しなし'}」）には"
                             "入力欄の印（入力規則・枠・ロックの解除）がありません。書く欄かどうかを規則では決められず、"
                             "AIエージェントが見出しから判断する必要があります。",
                             "書いてほしいセルには、入力規則（文字数の制限だけでもよい）か枠線を付け、保護するならロックを外す。",
                             _cells(n, sorted(und)), len(und)))
        if sh.protected and not any(not wb.style(c.style).locked for c in sh.cells.values()):
            expects_input = bool(sh.validations)
            F.append(Finding(2, "blocker" if expects_input else "info", "locked", "保護されていて入力できるセルがない",
                             "シート全体が保護されています。" + ("入力規則があるのに、入力できません。" if expects_input else "読むだけのシートと思われます。"),
                             "入力欄だけ「ロック」を外してから保護する。", [n]))

        # ---------------------------------------------------- 3. 検証できる
        if inputs:
            no_rule = [k for k in inputs if not any(v.covers(*k) for v in sh.validations)]
            ratio = len(no_rule) / len(inputs)
            if ratio > 0.5 and len(no_rule) >= 3:
                F.append(Finding(3, "warning", "no_validation", "入力のきまりが書かれていない入力欄",
                                 f"入力欄の候補 {len(inputs)} か所のうち {len(no_rule)} か所に入力規則がありません。"
                                 "何を書けばよいか、書いた値が正しいかを、記入する前に確かめられません。",
                                 "データの入力規則で、選択肢・数値の範囲・文字数を決め、「入力時メッセージ」に説明を書く。",
                                 _cells(n, sorted(no_rule)), len(no_rule)))
        if not sh.comments and not any(v.prompt for v in sh.validations) and inputs and len(inputs) >= 5:
            F.append(Finding(3, "info", "no_guidance", "入力欄の説明がない",
                             "メモ（コメント）も入力時メッセージもありません。", "記入例と入力ルールを書類の中に置く。", [n]))

    # 数式の検証結果
    s = rep.summary()
    funcs = used_functions(wb)
    rounding = {k: v for k, v in funcs.items() if k in ("ROUND", "ROUNDDOWN", "ROUNDUP", "INT", "TRUNC", "FLOOR", "CEILING", "MROUND")}
    unsup = [r for r in rep.results if r.status == "unsupported"]
    if unsup:
        reasons: dict[str, int] = {}
        for r in unsup:
            reasons[r.reason] = reasons.get(r.reason, 0) + 1
        top = "、".join(f"{k}（{v}件）" for k, v in sorted(reasons.items(), key=lambda x: -x[1])[:4])
        F.append(Finding(3, "warning", "unverifiable_formula", "再計算で確かめられない数式",
                         f"{len(unsup)} 個の数式は、Excelと同じ規則で計算し直して確かめることができませんでした：{top}",
                         "揮発性の関数（INDIRECT・OFFSET・TODAY）や外部参照を避け、単純な参照と関数で組む。",
                         [f"{r.sheet}!{r.address}" for r in unsup[:12]], len(unsup)))
    mism = [r for r in rep.results if r.status == "mismatch"]
    if mism:
        F.append(Finding(3, "blocker", "stale_values", "保存されている計算結果が、数式と一致しない",
                         f"{len(mism)} 個のセルで、ファイルに保存された値と、数式を計算し直した値が違います。"
                         "手動計算の設定や、別のソフトでの保存が原因のことがあります。読む人によって見える金額が変わります。",
                         "計算方法を「自動」にして保存し直す。",
                         [f"{r.sheet}!{r.address}" for r in mism[:12]], len(mism)))
    nocache = [r for r in rep.results if r.status == "no_cache"]
    if nocache:
        F.append(Finding(3, "warning", "no_cached_values", "計算結果が保存されていない",
                         f"{len(nocache)} 個の数式セルに計算結果が保存されていません。Excel以外では値が見えません。",
                         "Excelで開いて保存し直す。", [f"{r.sheet}!{r.address}" for r in nocache[:12]], len(nocache)))
    vol = [r for r in rep.results if r.status == "volatile"]
    if vol:
        F.append(Finding(3, "info", "volatile", "開くたびに値が変わる関数",
                         f"TODAY() などを使うセルが {len(vol)} 個あります。提出した日と確認した日で値が変わります。",
                         "日付は入力欄にし、計算で今日の日付を使わない。", [f"{r.sheet}!{r.address}" for r in vol[:12]], len(vol)))
    if rounding:
        F.append(Finding(3, "info", "rounding_in_formulas", "端数処理が数式の中にある",
                         "端数処理（" + "、".join(f"{k} {v}件" for k, v in rounding.items()) + "）が数式の中に埋め込まれています。"
                         "手で計算する人やほかのシステムでは、同じ規則を知らないと金額がずれます。",
                         "端数処理の規則（切り捨て・四捨五入、どの段階で）を記入要領に文章でも書く。", count=sum(rounding.values())))
    if wb.has_macros:
        F.append(Finding(3, "blocker", "macro_logic", "マクロの中の計算",
                         "マクロで行う計算は、ファイルの中に答えが残らず、結果が正しいかを確かめる手段がありません。",
                         "計算は関数で行い、マクロは使わない。"))

    multi = sum(1 for sh in wb.sheets if not sh.state.startswith("non-worksheet")) > 1
    if multi:
        for f in F:
            if f.where and f.code not in ("hidden_sheet",):
                sheet = f.where[0].split("!")[0]
                if all(w.split("!")[0] == sheet for w in f.where):
                    f.title = f"{f.title}（{sheet}）"
    order = {"blocker": 0, "agent": 1, "warning": 2, "info": 3}
    F.sort(key=lambda f: (f.level, order[f.severity]))
    stats = {
        "sheets": len(wb.sheets),
        "cells": total_cells,
        "formulas": s["total"],
        "merges": sum(len(sh.merges) for sh in wb.sheets),
        "validations": sum(len(sh.validations) for sh in wb.sheets),
        "functions": funcs,
        "macros": wb.has_macros,
    }
    return Diagnosis(wb.filename, F, s, stats)
