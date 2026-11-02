"""診断と検証の結果を、作り手が読むための1枚のHTMLにする（外部ファイルに依存しない）。"""

from __future__ import annotations

import html
from datetime import datetime

LEVEL_NAMES = ["届く", "読める", "操作できる", "検証できる"]
LEVEL_Q = [
    "エージェントが書類にたどり着き、取得できるか",
    "文字と構造を正しく取り出せるか",
    "AIが本人の指示で記入・実行できるか",
    "結果が正しいと確かめられるか",
]

CSS = """
:root{--bg:#fbfaf7;--fg:#1d1d1b;--muted:#5f5e58;--line:#d9d6cc;--card:#fff;--block:#b3261e;--blockbg:#fdecea;
--warn:#8a5a00;--warnbg:#fff4d6;--info:#2f5d8a;--infobg:#e8f0f8;--ok:#1e6b3a;--okbg:#e5f4ea;--accent:#1d1d1b}
@media (prefers-color-scheme:dark){:root{--bg:#171715;--fg:#ecebe6;--muted:#a8a69d;--line:#3a3934;--card:#211f1c;
--block:#ff8a80;--blockbg:#3a1a17;--warn:#ffcc66;--warnbg:#33290f;--info:#9cc3ea;--infobg:#15263a;--ok:#8fd6a5;--okbg:#14301e;--accent:#ecebe6}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.75 "Hiragino Sans","Noto Sans JP",system-ui,sans-serif}
main{max-width:960px;margin:0 auto;padding:32px 16px 80px}h1{font-size:1.6rem;line-height:1.4;margin:0 0 4px}
h2{font-size:1.2rem;margin:40px 0 12px;padding-bottom:6px;border-bottom:2px solid var(--line)}
.meta{color:var(--muted);font-size:.9rem;word-break:break-all}
.ladder{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:24px 0}
.step{border:2px solid var(--line);border-radius:10px;padding:12px;background:var(--card)}
.step b{display:block;font-size:1.05rem}.step small{color:var(--muted);display:block;line-height:1.5}
.step .st{margin-top:8px;font-weight:700}.step.block{border-color:var(--block)}.step.block .st{color:var(--block)}
.step.warn .st{color:var(--warn)}.step.ok .st{color:var(--ok)}
@media (max-width:640px){.ladder{grid-template-columns:1fr 1fr}}
.headline{background:var(--okbg);border-left:6px solid var(--ok);padding:16px;border-radius:6px;font-size:1.05rem}
.headline.ng{background:var(--blockbg);border-color:var(--block)}
.f{background:var(--card);border:1px solid var(--line);border-left:6px solid var(--info);border-radius:8px;padding:14px 16px;margin:12px 0}
.f.blocker{border-left-color:var(--block)}.f.warning,.f.agent{border-left-color:var(--warn)}
.f h3{margin:0 0 4px;font-size:1.05rem}.tag{display:inline-block;font-size:.8rem;font-weight:700;padding:1px 8px;border-radius:99px;margin-right:6px}
.tag.blocker{background:var(--blockbg);color:var(--block)}.tag.warning,.tag.agent{background:var(--warnbg);color:var(--warn)}.tag.info{background:var(--infobg);color:var(--info)}
.fix{margin:8px 0 0;padding:8px 12px;background:var(--okbg);border-radius:6px}.fix::before{content:"直し方：";font-weight:700}
.where{color:var(--muted);font-size:.85rem;word-break:break-all}
table{border-collapse:collapse;width:100%;font-size:.92rem;background:var(--card)}th,td{border:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
th{background:var(--bg)}td.num{text-align:right;font-variant-numeric:tabular-nums}.scroll{overflow-x:auto}
code{font-family:ui-monospace,Menlo,monospace;font-size:.88em;word-break:break-all}
"""


def _e(s) -> str:
    return html.escape(str(s))


def _v(x) -> str:
    if x is None:
        return "（空）"
    if isinstance(x, float) and x == int(x) and abs(x) < 1e15:
        return f"{int(x):,}"
    return str(x)


def _grade_section(g: dict) -> list[str]:
    w = g.get("write") or {"grade": "A", "reasons": []}
    out = ["<section aria-labelledby='grade'><h2 id='grade'>段階</h2>",
           f"<p class='headline{' ng' if g['grade'] in 'DEF' or w['grade'] == 'D' else ''}'>"
           f"読む <b style='font-size:1.6em'>{_e(g['grade'])}</b>　書く <b style='font-size:1.6em'>{_e(w['grade'])}</b>　{_e(g['headline'])}</p>",
           "<ul>"]
    for label, reasons, top, plain in (("読む", g["reasons"], g["grade"], "SA"), ("書く", w["reasons"], w["grade"], "A")):
        for r in reasons:
            if r["grade"] in plain and top not in plain:
                continue
            where = f"（{_e('、'.join(r['where'][:6]))}）" if r["where"] else ""
            by = "（エージェントの判断）" if r.get("by") == "agent" else "（規則）"
            out.append(f"<li><b>{label}[{_e(r['grade'])}] {_e(r['title'])}</b>{by}：{_e(r['detail'])}{where}</li>")
    for c in g.get("agent_checks", []):
        label = "書く" if c.get("score") == "write" else "読む"
        out.append(f"<li><b>？ エージェントが確かめること（あれば{label}段階 {_e(c['grade'])}）</b>：{_e(c['question'])}</li>")
    out.append("</ul><p class='meta'>読む段階　" + "　".join(f"{_e(k)}：{_e(v)}" for k, v in g["scale"].items()) + "</p>")
    out.append("<p class='meta'>書く段階　" + "　".join(f"{_e(k)}：{_e(v)}" for k, v in (g.get("write_scale") or {}).items())
               + "</p></section>")
    return out


def render(diag: dict, proof: dict | None = None, title: str | None = None, grade: dict | None = None) -> str:
    score = diag["score"]
    by = score["by_level"]
    out = ["<!doctype html><html lang='ja'><head><meta charset='utf-8'>",
           "<meta name='viewport' content='width=device-width,initial-scale=1'>",
           f"<title>{_e(title or '書類の診断')}</title><style>{CSS}</style></head><body><main>"]
    out.append(f"<h1>{_e(title or '書類の診断')}</h1>")
    out.append(f"<p class='meta'>{_e(diag['file'])}　·　{datetime.now():%Y-%m-%d %H:%M}　·　todoku</p>")
    if grade:
        out += _grade_section(grade)

    out.append("<section aria-labelledby='ladder'><h2 id='ladder'>4段階のラダー</h2>")
    out.append("<p>上の段階ほど難しく、最初に「壁」がある段階で止まります。")
    if score.get("rules_reached", score["reached"]) < score["reached"]:
        out.append(f"AIエージェントが判断すれば「{_e(score['reached_name'])}」まで届きますが、"
                   f"AIを使わない規則だけでは「{_e(score['rules_reached_name'])}」で止まります。")
    out.append("</p><ol class='ladder'>")
    for i, name in enumerate(LEVEL_NAMES):
        b = by[name]["blocker"]
        a = by[name].get("agent", 0)
        w = by[name]["warning"]
        cls, st = ("block", f"壁 {b} 件") if b else (("warn", f"規則では越えられない {a} 件") if a else
                                                    (("warn", f"注意 {w} 件") if w else ("ok", "問題なし")))
        out.append(f"<li class='step {cls}'><b>{i}. {name}</b><small>{_e(LEVEL_Q[i])}</small><div class='st'>{st}</div></li>")
    out.append("</ol></section>")

    v = diag["verify"]
    ok = v["mismatch"] == 0
    out.append("<section aria-labelledby='calc'><h2 id='calc'>関数の全数照合</h2>")
    out.append(f"<p class='headline{'' if ok else ' ng'}'>数式 {v['total']:,} 個のうち <b>{v['match']:,} 個</b> が、Excelの保存した計算結果と一致"
               f"（不一致 {v['mismatch']}・再計算できない {v['unsupported']}・揮発性 {v['volatile']}・結果なし {v['no_cache']}）。</p>")
    out.append("<p>AIを使わない決定的な処理で、Excelと同じ規則（有効数字15桁、四捨五入は0から遠い方へ）で計算し直しています。</p></section>")

    if proof:
        out.append("<section aria-labelledby='proof'><h2 id='proof'>書き込みの証明</h2>")
        cls = "" if proof["ok"] else " ng"
        out.append(f"<p class='headline{cls}'>{_e(proof['headline'])}</p>")
        out.append(f"<p class='meta'>元: <code>{proof['original_sha256']}</code><br>後: <code>{proof['new_sha256']}</code></p>")
        if proof["unexpected"]:
            out.append("<p><b>想定外の変更：</b></p><ul>" + "".join(f"<li>{_e(u)}</li>" for u in proof["unexpected"]) + "</ul>")
        out.append("<div class='scroll'><table><thead><tr><th>シート</th><th>セル</th><th>種類</th><th>前</th><th>後</th></tr></thead><tbody>")
        for c in proof["cell_changes"]:
            kind = "入力" if c["kind"] == "input" else "計算結果"
            out.append(f"<tr><td>{_e(c['sheet'])}</td><td>{_e(c['cell'])}</td><td>{kind}</td>"
                       f"<td class='num'>{_e(_v(c['before']))}</td><td class='num'>{_e(_v(c['after']))}</td></tr>")
        out.append("</tbody></table></div></section>")

    out.append("<section aria-labelledby='find'><h2 id='find'>見つかった壁と直し方</h2>")
    if not diag["findings"]:
        out.append("<p>機械的に分かる壁は見つかりませんでした。</p>")
    for i, name in enumerate(LEVEL_NAMES):
        fs = [f for f in diag["findings"] if f["level"] == i]
        if not fs:
            continue
        out.append(f"<h3 style='margin-top:24px'>{i}. {name}</h3>")
        for f in fs:
            out.append(f"<article class='f {f['severity']}'><h3><span class='tag {f['severity']}'>{f['severity_name']}</span>{_e(f['title'])}</h3>")
            out.append(f"<p style='margin:0'>{_e(f['detail'])}</p><p class='fix'>{_e(f['fix'])}</p>")
            if f["where"]:
                more = f"ほか {f['count'] - len(f['where'])} か所" if f["count"] > len(f["where"]) else ""
                out.append(f"<p class='where'>場所：{_e('、'.join(f['where']))} {more}</p>")
            out.append("</article>")
    out.append("</section>")

    st = diag["stats"]
    out.append("<section aria-labelledby='stats'><h2 id='stats'>書類の概要</h2><div class='scroll'><table><tbody>")
    rows = [("シート", st["sheets"]), ("セル", f"{st['cells']:,}"), ("数式", f"{st['formulas']:,}"),
            ("セル結合", st["merges"]), ("入力規則", st["validations"]), ("マクロ", "あり" if st["macros"] else "なし"),
            ("使われている関数", "、".join(f"{k} {v}" for k, v in sorted(st["functions"].items(), key=lambda x: -x[1])) or "なし")]
    for k, v in rows:
        out.append(f"<tr><th scope='row'>{k}</th><td>{_e(v)}</td></tr>")
    out.append("</tbody></table></div></section>")
    out.append("<p class='meta' style='margin-top:40px'>この診断はすべてAIを使わない決定的な処理で行いました。同じファイルからは必ず同じ結果が出ます。</p>")
    out.append("</main></body></html>")
    return "".join(out)
