"""書類の段階（S〜F）：誰が扱えば、正確に読んで入れられるか。いちばん悪い条件で決める。

S 素で問題ない（表のデータそのもの：CSV・JSON など）
A 規則だけで正確に読み書きできる（入力欄がすべて作り手の印で決まり、数式もすべて確かめられる）
B エージェントが読めば正確に入れられる（一部の欄が配置からの推定か、規則では決められない）
C エージェントが工夫して変換できる（マクロなど、規則では再現しない処理がある）
D エージェントの解釈が必要（書く値がおおよそでよい・人によって解釈が分かれる書き方がある）
E エージェントが得意でない形式（文字の情報がないページ：画像だけの PDF・スキャン）
F アクセスできない（開けない・パスワードが分からない・読めない部分がある、または押印・自署のように、エージェントにはできない記入がある）
  ※郵送・持参・来庁などの提出のしかたは、書類の外のことなので、段階の対象にしない

判定は「規則」か「エージェント」のどちらかで行う。言葉の一致（「押印」「約」など）は、打ち消しや書類の名前の中の語も
拾ってしまい判定の根拠として弱いので、使わない。
- 規則：ファイルの構造から確実に分かること（フォームの欄・入力規則などの印、配置からの推定か、画像だけのページ、
  マクロの命令（意味の決まったプログラムの文法で調べる）、開けるか）
- エージェント：書かれた内容の意味が要ること（エージェントにはできない記入、書く値のぼかし）。todoku は問い（agent_checks）を
  示し、エージェントが窓を読んで答えた判断（combine）を、出どころ「エージェントの判断」として段階に合わせる
"""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field

ORDER = "SABCDEF"
NAMES = {
    "S": "素で問題ない（データそのもの）",
    "A": "規則だけで正確に読み書きできる",
    "B": "エージェントが読めば正確に入れられる",
    "C": "エージェントが工夫して変換できる",
    "D": "エージェントの解釈が必要",
    "E": "エージェントが得意でない形式",
    "F": "アクセスできない",
}

# エージェントが窓を読んで確かめること（規則では決めない）。見つかれば、その段階
AGENT_CHECKS = [
    {"id": "cannot_do", "grade": "F", "title": "エージェントにはできない記入",
     "question": "押印・自署（手書き）のように、書類の中で、エージェントには代わりにできない記入を求めていますか。"
                 "求めていれば、その場所と書き方を答えてください（「押印不要」のように求めていないものは除く。"
                 "郵送・持参・来庁などの提出のしかたは、書類の外のことなので、ここでは数えません）。"},
    {"id": "vague_value", "grade": "D", "title": "書く値をぼかす書き方",
     "question": "書く値そのものがおおよそでよい、または人によって解釈が分かれる書き方の入力欄（「約［　］人」「月［　］日程度」"
                 "「必要に応じて記入」など）がありますか。あれば、その欄の名前と書き方を答えてください。"},
]

# 印刷・表示だけのマクロ：この命令だけでできた手続きは、セルの値を変えない（VBA の文の意味で決める）
_SAFE_STMT = re.compile(
    r"^(?:(?:Private|Public)\s+)?(?:Sub|Function)\b|^End\s+(?:Sub|Function|With|If)\b|^Exit\s+(?:Sub|Function)\b|^With\b|^Else\b|"
    r"^(?:Dim|Const)\b|^'|^Rem\b|^Option\b|^Attribute\b|^On\s+Error\b|"
    r"\.(?:Select|Activate|PrintOut|PrintPreview)\b[^=]*$|^(?:ActiveWindow|Application)\.(?:SelectedSheets\.)?(?:PrintOut|PrintPreview)\b|"
    r"^Application\.(?:ScreenUpdating|DisplayAlerts|MoveAfterReturnDirection|StatusBar|Cursor|EnableEvents)\s*=|"
    r"^MsgBox\b|^(?:Sheets|Worksheets|Range|Cells)\(.*\)\.(?:Select|Activate)\b", re.I)


@dataclass
class Reason:
    grade: str
    title: str
    detail: str
    where: list[str] = field(default_factory=list)
    by: str = "rules"  # rules（規則）/ agent（エージェントの判断）


@dataclass
class Grade:
    grade: str
    reasons: list[Reason] = field(default_factory=list)
    agent_checks: list[dict] = field(default_factory=list)  # エージェントが確かめること（まだ答えていない問い）
    without_barriers: str | None = None  # エージェントにはできない手続き（F）がなければ、どの段階か

    @property
    def name(self) -> str:
        return NAMES[self.grade]

    @property
    def barriers(self) -> list[Reason]:
        """エージェントの判断で見つかった、エージェントにはできない手続き。"""
        return [r for r in self.reasons if r.by == "agent" and r.grade == "F"]

    def headline(self) -> str:
        s = f"段階 {self.grade}：{self.name}"
        worst = [r for r in self.reasons if r.grade == self.grade]
        if worst and self.grade not in ("S", "A"):
            s += "（" + "・".join(dict.fromkeys(r.title + ("：エージェントの判断" if r.by == "agent" else "") for r in worst)) + "）"
        if self.without_barriers:
            s += f"。それがなければ段階 {self.without_barriers}：{NAMES[self.without_barriers]}"
        if self.agent_checks:
            s += "。エージェントが確かめること：" + "・".join(f"{c['title']}（あれば {c['grade']}）" for c in self.agent_checks)
        return s

    def to_dict(self) -> dict:
        return {"grade": self.grade, "name": self.name, "headline": self.headline(),
                "reasons": [asdict(r) for r in self.reasons], "agent_checks": self.agent_checks,
                "without_barriers": self.without_barriers, "scale": NAMES}


def _final(reasons: list[Reason], base: str = "A") -> Grade:
    g = max([base] + [r.grade for r in reasons], key=ORDER.index)
    reasons = sorted(reasons, key=lambda r: -ORDER.index(r.grade))
    checks = [dict(c) for c in AGENT_CHECKS] if g not in ("S", "F") else []
    return Grade(g, reasons, checks)


def combine(rule: "Grade | dict", findings: list[dict], answered: list[str] | None = None) -> Grade:
    """規則で決めた段階に、エージェントの判断（findings：{grade, title, detail, where}）を合わせる。
    answered：エージェントが確かめ終えた問いの id（見つからなかったものも含む）。答えていない問いは残す。"""
    if isinstance(rule, dict):
        rule = Grade(rule["grade"], [Reason(**r) for r in rule["reasons"]], list(rule.get("agent_checks", [])))
    reasons = [r for r in rule.reasons if r.by == "rules"]
    for f in findings:
        g = str(f.get("grade", "")).upper()
        if len(g) != 1 or g not in ORDER:
            raise ValueError(f"段階は {', '.join(ORDER)} のどれかです：{g!r}")
        where = f.get("where") or []
        reasons.append(Reason(g, str(f.get("title") or "エージェントの判断"), str(f.get("detail") or ""),
                              [where] if isinstance(where, str) else [str(w) for w in where], by="agent"))
    rest = max([rule.grade] + [r.grade for r in reasons if not (r.by == "agent" and r.grade == "F")], key=ORDER.index)
    g = max([rest] + [r.grade for r in reasons], key=ORDER.index)
    found = {c["id"] for c in AGENT_CHECKS if any(r.by == "agent" and r.grade == c["grade"] for r in reasons)}
    done = set(answered or []) | found
    out = Grade(g, sorted(reasons, key=lambda r: -ORDER.index(r.grade)), [c for c in rule.agent_checks if c["id"] not in done])
    if any(r.by == "agent" and r.grade == "F" for r in reasons) and rest != "F":
        out.without_barriers = rest
    return out


def macro_is_display_only(source: str) -> bool:
    """手続きの命令が、すべて「選ぶ・切り替える・印刷する・画面の設定」だけなら True（セルの値を変えない）。"""
    lines = []
    for raw in source.replace(" _\n", " ").split("\n"):
        ln = raw.strip()
        if ln:
            lines.append(ln)
    return all(_SAFE_STMT.search(ln) for ln in lines)


def _macro_reasons(info) -> list[Reason]:
    if not isinstance(info, dict):
        return []
    if not info.get("readable"):
        return [Reason("F", "読めないマクロ", "マクロの入れ物を読めませんでした。読めないものには、エージェントもアクセスできません")]
    busy = [m for m in info["modules"] if m["procedures"] and not macro_is_display_only(m["source"])]
    if not busy:
        if info["modules"]:
            return [Reason("A", "印刷・表示だけのマクロ", "マクロの命令は、シートを選ぶ・印刷するなどだけで、セルの値を変えません",
                           [m["name"] for m in info["modules"]])]
        return []
    on_change = [f"{e['module']}.{e['procedure']}" for e in info["events"] if e["on_change"]]
    detail = (f"マクロの処理（{len(busy)} 個のモジュール）は、規則では再現しません。ボタンや自動で動く処理が要るかは、"
              "ソースを読んで判断します")
    if on_change:
        detail += f"。セルの値を変えると動くマクロ（{'、'.join(on_change[:5])}）があります"
    return [Reason("C", "マクロの処理", detail, [m["name"] for m in busy][:12])]


# ====================================================================== 形式ごと
def grade_blob(blob: bytes, filename: str = "") -> Grade:
    from .book import kind_of

    name = filename.lower()
    if name.endswith((".csv", ".tsv", ".json", ".jsonl", ".xml")):
        return Grade("S", [Reason("S", "データそのもの", "表のデータ（CSV・JSON など）で、変換しなくても機械が読めます")])
    try:
        k = kind_of(blob)
    except Exception as e:  # noqa: BLE001
        return Grade("F", [Reason("F", "開けない", f"ファイルの形式を判別できませんでした（{type(e).__name__}）")])
    if k == "encrypted":
        return Grade("F", [Reason("F", "パスワード", "パスワードで暗号化されています。パスワードがなければ開けません")])
    try:
        if k == "pdf":
            return _grade_pdf(blob, filename)
        if k in ("docx", "doc"):
            return _grade_word(blob, filename, k)
        if k in ("xlsx", "xls"):
            return _grade_excel(blob, filename)
    except Exception as e:  # noqa: BLE001
        return Grade("F", [Reason("F", "開けない", f"ファイルを読めませんでした（{type(e).__name__}: {str(e)[:80]}）")])
    return Grade("F", [Reason("F", "対応していない形式", f"この形式（{k}）は、まだ読めません")])


def _grade_pdf(blob: bytes, filename: str) -> Grade:
    from .pdf.doc import PdfDocument
    from .pdf.window import build

    doc = PdfDocument(blob, filename)
    win = build(doc)
    reasons: list[Reason] = []
    unmapped = [f"{p['number']}ページ（{p['unmapped']} 字）" for p in win["pages"] if p.get("unmapped")]
    if unmapped:
        reasons.append(Reason("F", "文字に直せない字", "フォントに文字の対応がなく、文字に直せない字があります（窓では �）。"
                              "読めないものには、エージェントもアクセスできません", unmapped))
    img = [f"{p['number']}ページ" for p in win["pages"] if p.get("image_only")]
    if img:
        reasons.append(Reason("E", "画像だけのページ", f"{len(img)} ページは文字の情報がなく、画像（スキャン）です。読むには文字の認識が要ります", img))
    if win["form"]["xfa"]:
        reasons.append(Reason("C", "XFA のフォーム", "Adobe 独自の動くフォーム（XFA）です。開くソフトによって中身が変わります"))
    agent = [f for p in win["pages"] for f in p["fields"] if f.get("agent")]
    fields = [f for p in win["pages"] for f in p["fields"] if not f.get("agent")]
    layout = [f for f in fields if f.get("source") == "layout"]
    if agent:
        reasons.append(_agent_reason(agent))
    if layout:
        reasons.append(Reason("B", "配置から推定した入力欄",
                              f"入力欄 {len(fields)} か所のうち {len(layout)} か所は、フォームの欄ではなく、罫線や空きから推定しました"
                              "（窓には見出しと前後の文字を添えています）", [f["id"] for f in layout[:12]]))
    elif fields:
        reasons.append(Reason("A", "フォームの欄", f"入力欄 {len(fields)} か所は、すべて PDF のフォームの欄です"))
    return _final(reasons)


def _agent_reason(agent: list[dict]) -> Reason:
    """規則では決められない所（窓の［…？］）：エージェントが窓を読んで判断し、書く所なら id を指定して書く（C）。"""
    kinds: dict[str, int] = {}
    for f in agent:
        kinds[f.get("why", "")] = kinds.get(f.get("why", ""), 0) + 1
    return Reason("C", "規則では決められない所",
                  f"書き込む所かもしれない所 {len(agent)} か所は、規則では欄と決められません（"
                  + "・".join(f"{k} {n}" for k, n in kinds.items())
                  + "）。エージェントが窓を読んで判断し、書く所なら id を指定して書きます",
                  [f["id"] for f in agent[:12]])


def _grade_word(blob: bytes, filename: str, kind: str) -> Grade:
    from .docx.window import build

    if kind == "doc":
        from .doc.book import DocFile

        doc = DocFile(blob, filename)
    else:
        from .docx.document import DocxDocument

        doc = DocxDocument(blob, filename)
    win = build(doc)
    reasons: list[Reason] = []
    agent = [f for f in win["fields"] if f.get("agent")]
    fields = [f for f in win["fields"] if not f.get("agent")]
    layout = [f for f in fields if f.get("source") == "layout"]
    if agent:
        reasons.append(_agent_reason(agent))
    if layout:
        reasons.append(Reason("B", "配置から推定した入力欄",
                              f"入力欄 {len(fields)} か所のうち {len(layout)} か所は、フォームの欄ではなく、空のマスや空白・下線から推定しました",
                              [f["id"] for f in layout[:12]]))
    elif fields:
        reasons.append(Reason("A", "フォームの欄", f"入力欄 {len(fields)} か所は、すべて Word のフォームの欄・コンテンツコントロールです"))
    reasons += _macro_reasons(win.get("macros"))
    return _final(reasons)


def _grade_excel(blob: bytes, filename: str) -> Grade:
    from .book import open_book
    from .xlsx.calc import verify
    from .xlsx.window import build

    wb = open_book(blob, filename)
    doc = build(wb)
    reasons: list[Reason] = []
    explicit = {"入力規則があるセル", "シート保護の中で編集できるセル", "枠のある空欄", "色の付いた空欄"}
    inferred, inline, overlap, undecided, n_fields = [], [], [], [], 0
    for sh in doc["sheets"]:
        for f in sh.get("fields", []):
            if f.get("unlabeled"):
                continue
            n_fields += 1
            if f["kind"] == "inline":
                inline.append(f["id"])
            elif f["kind"] != "checkbox" and not (set(f.get("reason", [])) & explicit):
                inferred.append(f["id"])
            if f.get("rules_overlap"):
                overlap.append(f["id"])
        undecided += [f"{sh['name']}!{u['cell']}" for u in sh.get("undecided", [])]
    if inline:
        reasons.append(Reason("B", "文字の中の空欄", f"{len(inline)} か所は、「令和　年　月　日」のような文字の中の空白から推定した欄です", inline[:12]))
    if inferred:
        reasons.append(Reason("B", "印のない入力欄", f"{len(inferred)} か所は、入力規則も枠もなく、数式の入力になっていることから推定した欄です",
                              inferred[:12]))
    if undecided:
        reasons.append(Reason("C", "入力欄か決められないセル",
                              f"表の {len(undecided)} か所の空のセルは、入力欄の印がなく、規則では決められません。"
                              "エージェントが窓を読んで判断し、書くならセルの番地を指定して書きます", undecided[:12]))
    if overlap:
        reasons.append(Reason("C", "重なった入力規則",
                              f"{len(overlap)} か所に入力規則が重なっていて、Excelがどれを使うかはファイルから決まりません。"
                              "窓に出したそれぞれの規則を、エージェントが読んで判断します", overlap[:12]))
    if n_fields and not (inline or inferred or undecided or overlap):
        reasons.append(Reason("A", "印のある入力欄", f"入力欄 {n_fields} か所は、すべて入力規則・シート保護・枠などの印で決まります"))
    rep = verify(wb).summary()
    if rep["unsupported"]:
        reasons.append(Reason("C", "確かめられない数式", f"{rep['unsupported']} 個の数式は、規則どおりに計算し直して確かめられません"))
    if rep["mismatch"]:
        reasons.append(Reason("C", "保存された値と合わない数式", f"{rep['mismatch']} 個のセルで、保存された計算結果が数式と一致しません"))
    reasons += _macro_reasons(doc.get("macros"))
    for sh_obj in wb.sheets:
        pics = [s for s in getattr(sh_obj, "shapes", []) if s.kind in ("picture", "chart")]
        if pics and len(sh_obj.cells) < 5:
            reasons.append(Reason("E", "画像だけのシート", f"シート「{sh_obj.name}」の中身は画像で、セルの文字がありません", [sh_obj.name]))
    return _final(reasons)
