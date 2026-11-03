"""書類の段階：読む段階（S〜F）と書く段階（A〜D）。どちらも、いちばん悪い条件で決める。

読む段階（S〜F）：誰が扱えば、中身（字・数・表の並び）を正確に読めるか。読むことだけで決める：入力欄（書く所）が
作り手の印で決まるか・配置からの推定か・規則では決められない所（窓の「？」）かは、書く段階で決め、読む段階には入れない
（様式でない統計表・報告の見出しや空のマスが「？」の候補になっても、中身が正確に読めれば読む段階は下げない）。

S 素で問題ない（表のデータそのもの：CSV・JSON など）
A 規則だけで正確に読める（中身の字・並び・表示の値がファイルの作りから決まり、数式もすべて確かめられる）
B エージェントが読めば正確（中身は窓にすべて出ているが、一部の値の読み方が規則では決まらず、窓に示した根拠で
  エージェントが決める：Excel 以外のアプリが保存した計算結果、いくつもの文字コードで仕様どおりに読める文字のファイル
  （候補と見本を窓に並べる。エージェントが選んで encoding で渡せば、その文字コードで読み直し、段階は下げない）、
  本文の中身である説明のない画像（字の中身はあるが、画像の所はエージェントが見て確かめる。タグ付きの PDF の Alt のない図も）など）
C エージェントが工夫して変換できる（マクロ・確かめられない数式・自動で動くスクリプトなど、規則では再現しない処理がある。
  どの文字コードとしても読めない。窓の読み落とし：窓とは別の読み方で拾った中身が窓に出ていない（Excel の Markdown の窓・
  PDF のページの字・Word の文字）ので、エージェントは窓だけを信じてはいけない。落ちた所を where に示す）
D エージェントの解釈が必要（書く値がおおよそでよい・人によって解釈が分かれる書き方がある：エージェントの判断）
E エージェントが得意でない形式（字として読めない中身があり、エージェントがページを画像として見て読む：画像だけの
  ページ・スキャン（todoku が開けない圧縮の画像も含む）、字に直せない本文の字、画像だけのシート・文書、本文に見出し・
  図表番号のほかに字の段落がなく中身が説明のない画像の文書・スライド）
F アクセスできない（開けない・パスワードが分からない・中身を取り出せない・ページが1つもない）
  ※字に直せない字が、行の頭の印（箇条の印など）だけなら、語は欠けていないので段階は変えず、注意（alerts）に位置を示す

書く段階（A〜D）：エージェントが、どうすれば書けるか。
A エージェントがそのまま書ける（入力欄がすべて作り手の印で決まる：フォームの欄・コンテンツコントロール・入力規則など）
B 解析しないと書けない（書く所を、罫線・空き・空のセルなどの配置から推定するか、規則では決められない所（窓の「？」）を
  エージェントが窓を読んで決める。スキャンのページ（PDF）は、エージェントがページの画像（view_page）を見て枠で書く）
C 画面（UI）から入力する（ボタンで動くマクロ、または書類が電子申請などの画面からの入力を求めている）
D 書けない（アクセスできない・字の情報がなく、ページの画像も作れない紙・押印や自署のような手書き）
  ※様式が認める所には、本人が同意したときだけ、本人の印影・署名の画像を置ける（fill の seal・signature）。
    それで書けるかは書類の決まり（意味）によるので、段階は変えず、問い（cannot_do）でエージェントが確かめる
  ※郵送・持参・来庁などの提出のしかたは、書類の外のことなので、段階の対象にしない

段階とは別に、アクセシビリティの注意（alerts）を示す：エージェントが読めるかとは別に、画面の読み上げなどを使う人に
届かない作りのこと（説明（代わりの字）のない画像：WCAG 1.1.1・PDF/UA、字に直せない行の頭の印：PDF/UA 7.21.7）。段階は変えない。

判定は「規則」か「エージェント」のどちらかで行う。言葉の一致（「押印」「約」など）は、打ち消しや書類の名前の中の語も
拾ってしまい判定の根拠として弱いので、使わない。
- 規則：ファイルの構造から確実に分かること（フォームの欄・入力規則などの印、配置からの推定か、画像だけのページ、
  字に直せない字が行の頭の印か本文の字か（字の並びの形）、マクロの命令（意味の決まったプログラムの文法で調べる）、開けるか）
- エージェント：書かれた内容の意味が要ること（エージェントにはできない記入、書く値のぼかし）。todoku は問い（agent_checks）を
  示し、エージェントが窓を読んで答えた判断（combine）を、出どころ「エージェントの判断」として段階に合わせる
"""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field

ORDER = "SABCDEF"
WRITE_ORDER = "ABCD"
WRITE_NAMES = {
    "A": "エージェントがそのまま書ける",
    "B": "解析しないと書けない",
    "C": "画面（UI）から入力する",
    "D": "書けない（アクセスできない・手書き）",
}
NAMES = {
    "S": "素で問題ない（データそのもの）",
    "A": "規則だけで正確に読める",
    "B": "エージェントが読めば正確",
    "C": "エージェントが工夫して変換できる",
    "D": "エージェントの解釈が必要",
    "E": "エージェントが得意でない形式",
    "F": "アクセスできない",
}

# エージェントが窓を読んで確かめること（規則では決めない）。見つかれば、その段階（score：read 読む・write 書く）
AGENT_CHECKS = [
    {"id": "cannot_do", "score": "write", "grade": "D", "title": "エージェントにはできない記入",
     "question": "押印・自署（手書き）のように、書類の中で、エージェントには代わりにできない記入を求めていますか。"
                 "求めていれば、その場所と書き方を答えてください（「押印不要」のように求めていないものは除く。"
                 "郵送・持参・来庁などの提出のしかたは、書類の外のことなので、ここでは数えません）。"
                 "なお、様式が印影・署名の画像を認めている所（注意書きに、画像や電子の印影でよいとある所など）には、"
                 "本人が同意したときだけ、fill_document で本人の印影・署名の画像を置けます"
                 "（{\"seal\": 画像のパス, \"consent\": true, …}）。実印の押印・自署（手書き）が求められる所や、"
                 "画像でよいかが書類から分からない所は、ここに答えて、本人に伝えてください。"},
    {"id": "ui_only", "score": "write", "grade": "C", "title": "画面からの入力",
     "question": "この書類のファイルに書くのではなく、電子申請の画面など、別の画面（UI）から入力するよう求めていますか。"
                 "求めていれば、その画面の名前と、書類のどこにそう書いてあるかを答えてください"
                 "（郵送・持参・来庁などの提出のしかたは、ここでは数えません）。"},
    {"id": "vague_value", "score": "read", "grade": "D", "title": "書く値をぼかす書き方",
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
    # rules（規則）/ agent（エージェントの判断：combine の findings）/ agent_choice（エージェントが道具に渡した選択で、
    # 規則がそれを確かめたもの：文字だけのファイルの encoding など。規則の理由と同じく、combine で残す）
    by: str = "rules"


@dataclass
class Grade:
    grade: str  # 読む段階（S〜F）
    reasons: list[Reason] = field(default_factory=list)
    agent_checks: list[dict] = field(default_factory=list)  # エージェントが確かめること（まだ答えていない問い）
    without_barriers: str | None = None  # エージェントにはできない記入（書く D）がなければ、書く段階はどれか
    alerts: list[Reason] = field(default_factory=list)  # 段階とは別の、アクセシビリティの注意
    write: str = "A"  # 書く段階（A〜D）
    write_reasons: list[Reason] = field(default_factory=list)

    @property
    def name(self) -> str:
        return NAMES[self.grade]

    @property
    def write_name(self) -> str:
        return WRITE_NAMES[self.write]

    @property
    def barriers(self) -> list[Reason]:
        """エージェントの判断で見つかった、エージェントにはできない記入。"""
        return [r for r in self.write_reasons if r.by == "agent" and r.grade == "D"]

    def headline(self) -> str:
        def why(reasons, g, plain):
            worst = [r for r in reasons if r.grade == g]
            if not worst or g in plain:
                return ""
            return "（" + "・".join(dict.fromkeys(r.title + ("：エージェントの判断" if r.by == "agent" else "") for r in worst)) + "）"

        s = f"読む段階 {self.grade}：{self.name}{why(self.reasons, self.grade, 'SA')}"
        s += f"。書く段階 {self.write}：{self.write_name}{why(self.write_reasons, self.write, 'A')}"
        if self.without_barriers:
            s += f"。それがなければ書く段階 {self.without_barriers}：{WRITE_NAMES[self.without_barriers]}"
        if self.agent_checks:
            s += "。エージェントが確かめること：" + "・".join(
                f"{c['title']}（あれば{'書く' if c['score'] == 'write' else '読む'}段階 {c['grade']}）" for c in self.agent_checks)
        if self.alerts:
            s += "。段階とは別の、アクセシビリティの注意：" + "・".join(a.title for a in self.alerts)
        return s

    def to_dict(self) -> dict:
        return {"grade": self.grade, "name": self.name, "headline": self.headline(),
                "reasons": [asdict(r) for r in self.reasons], "agent_checks": self.agent_checks,
                "write": {"grade": self.write, "name": self.write_name, "reasons": [asdict(r) for r in self.write_reasons]},
                "without_barriers": self.without_barriers, "alerts": [asdict(a) for a in self.alerts],
                "scale": NAMES, "write_scale": WRITE_NAMES}


def _worst(order: str, base: str, reasons: list[Reason]) -> str:
    return max([base] + [r.grade for r in reasons], key=order.index)


def _final(reasons: list[Reason], write: list[Reason] | None = None, base: str = "A") -> Grade:
    write = write or []
    g = _worst(ORDER, base, reasons)
    w = _worst(WRITE_ORDER, "A", write)
    checks = [dict(c) for c in AGENT_CHECKS] if g not in ("S", "F") else []
    return Grade(g, sorted(reasons, key=lambda r: -ORDER.index(r.grade)), checks, write=w,
                 write_reasons=sorted(write, key=lambda r: -WRITE_ORDER.index(r.grade)))


def _locked(title: str, detail: str, write_detail: str) -> Grade:
    """開けない書類：読む F・書く D。"""
    return Grade("F", [Reason("F", title, detail)], write="D", write_reasons=[Reason("D", title, write_detail)])


def combine(rule: "Grade | dict", findings: list[dict], answered: list[str] | None = None) -> Grade:
    """規則で決めた段階に、エージェントの判断（findings）を合わせる。
    findings：{check: 問いの id, title, detail, where}（段階は問いで決まる）か、{score: "read"|"write", grade, title, detail, where}。
    answered：エージェントが確かめ終えた問いの id（見つからなかったものも含む）。答えていない問いは残す。"""
    if isinstance(rule, dict):
        w = rule.get("write") or {}
        rule = Grade(rule["grade"], [Reason(**r) for r in rule["reasons"]], list(rule.get("agent_checks", [])),
                     alerts=[Reason(**a) for a in rule.get("alerts", [])], write=w.get("grade", "A"),
                     write_reasons=[Reason(**r) for r in w.get("reasons", [])])
    checks = {c["id"]: c for c in AGENT_CHECKS}
    reasons = [r for r in rule.reasons if r.by != "agent"]
    write = [r for r in rule.write_reasons if r.by != "agent"]
    found: set[str] = set()
    for f in findings:
        c = checks.get(str(f.get("check") or ""))
        if f.get("check") and c is None:
            raise ValueError(f"問いの id は {', '.join(checks)} のどれかです：{f.get('check')!r}")
        score = c["score"] if c else str(f.get("score") or "read")
        if score not in ("read", "write"):
            raise ValueError(f"score は read（読む）か write（書く）です：{score!r}")
        order = ORDER if score == "read" else WRITE_ORDER
        g = c["grade"] if c else str(f.get("grade", "")).upper()
        if len(g) != 1 or g not in order:
            raise ValueError(f"{'読む' if score == 'read' else '書く'}段階は {', '.join(order)} のどれかです：{g!r}")
        if c:
            found.add(c["id"])
        else:
            found |= {k for k, x in checks.items() if x["score"] == score and x["grade"] == g}
        where = f.get("where") or []
        r = Reason(g, str(f.get("title") or (c["title"] if c else "エージェントの判断")), str(f.get("detail") or ""),
                   [where] if isinstance(where, str) else [str(x) for x in where], by="agent")
        (reasons if score == "read" else write).append(r)
    rest = _worst(WRITE_ORDER, rule.write, [r for r in write if not (r.by == "agent" and r.grade == "D")])
    done = set(answered or []) | found
    out = Grade(_worst(ORDER, rule.grade, reasons), sorted(reasons, key=lambda r: -ORDER.index(r.grade)),
                [c for c in rule.agent_checks if c["id"] not in done], alerts=list(rule.alerts),
                write=_worst(WRITE_ORDER, rule.write, write), write_reasons=sorted(write, key=lambda r: -WRITE_ORDER.index(r.grade)))
    if any(r.by == "agent" and r.grade == "D" for r in write) and rest != "D":
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
def grade_blob(blob: bytes, filename: str = "", password: str | None = None, encoding: str | None = None) -> Grade:
    """password：パスワードで保護された PDF を開くとき（Office の書類は、api で解いた中身を渡す）。
    encoding：文字だけのファイル（CSV・TSV・JSON・.txt）の文字コードを、エージェントが候補から選んで指定したとき。"""
    from .book import kind_of

    # ---- 追加の形式（CSV・TSV・JSON・JSON Lines・RTF・OpenDocument）：形式ごとの段階の決め方で（openkinds）
    from .openkinds import KINDS as _OPEN_KINDS
    from .openkinds import grade as _open_grade

    try:
        k_open = kind_of(blob, filename)
    except Exception:  # noqa: BLE001
        k_open = None
    if k_open in _OPEN_KINDS:
        return _open_grade(k_open, blob, filename, encoding)
    # ----
    try:
        k = kind_of(blob)
    except Exception as e:  # noqa: BLE001
        return _locked("開けない", f"ファイルの形式を判別できませんでした（{type(e).__name__}）", "開けないので、書けません")
    if k == "zip":
        return _grade_zip(blob, password)
    locked = k == "encrypted"
    if not locked and k in ("xls", "doc", "pdf") and not password:
        try:
            from .unlock import needs_password

            locked = needs_password(blob)
        except Exception:  # noqa: BLE001
            locked = False
    if locked:
        return _locked("パスワード", "パスワードで保護されています。パスワードがなければ開けません"
                       "（本人にパスワードを聞けば、解いた中身で段階を決めます）", "パスワードがなければ、書けません")
    try:
        if k == "pdf":
            return _grade_pdf(blob, filename, password)
        if k in ("docx", "doc"):
            return _grade_word(blob, filename, k)
        if k in ("xlsx", "xls"):
            return _grade_excel(blob, filename)
        if k == "jtd":
            return _grade_jtd(blob, filename)
    except Exception as e:  # noqa: BLE001
        return _locked("開けない", f"ファイルを読めませんでした（{type(e).__name__}: {str(e)[:80]}）", "読めないので、書けません")
    return _locked("対応していない形式", f"この形式（{k}）は、まだ読めません", "読めないので、書けません")


def _grade_zip(blob: bytes, password: str | None = None, depth: int = 0) -> Grade:
    """ZIP：中のファイルごとに段階を決め、いちばん低いものを ZIP の段階にする（理由の頭に、中の名前を付ける）。
    中の、扱えない形式のファイル・取り出せないファイルは、そのファイルが F（開けない）。"""
    from .ziparc import ArchiveError, extract, files

    try:
        ms = files(blob)
    except ArchiveError as e:
        return _locked("開けない", str(e), "開けないので、書けません")
    if not ms:
        return _locked("空の ZIP", "ZIP の中にファイルがありません", "書くファイルがありません")
    if depth >= 4:
        return _locked("深い入れ子の ZIP", "ZIP の中の ZIP が深すぎて、中を読みません", "読めないので、書けません")
    reasons: list[Reason] = []
    write: list[Reason] = []
    alerts: list[Reason] = []
    reads, writes = [], []
    for m in ms:
        try:
            sub = extract(blob, m.name)
            g = _grade_zip(sub, password, depth + 1) if sub[:2] == b"PK" and _is_plain_zip(sub) else grade_blob(sub, m.name, password)
        except ArchiveError as e:
            g = _locked("取り出せない", str(e), "取り出せないので、書けません")
        reads.append(g.grade)
        writes.append(g.write)
        reasons += [Reason(r.grade, f"{m.name}：{r.title}", r.detail, r.where, r.by) for r in g.reasons]
        write += [Reason(r.grade, f"{m.name}：{r.title}", r.detail, r.where, r.by) for r in g.write_reasons]
        alerts += [Reason(a.grade, a.title, f"{m.name}：{a.detail}", a.where, a.by) for a in g.alerts]
    gr = max(reads, key=ORDER.index)
    wr = max(writes, key=WRITE_ORDER.index)
    checks = [dict(c) for c in AGENT_CHECKS] if gr not in ("S", "F") else []
    return Grade(gr, sorted(reasons, key=lambda r: -ORDER.index(r.grade)), checks, alerts=alerts, write=wr,
                 write_reasons=sorted(write, key=lambda r: -WRITE_ORDER.index(r.grade)))


def _is_plain_zip(blob: bytes) -> bool:
    from .book import kind_of

    try:
        return kind_of(blob) == "zip"
    except Exception:  # noqa: BLE001
        return False


def _grade_pdf(blob: bytes, filename: str, password: str | None = None) -> Grade:
    from .pdf.doc import PdfDocument
    from .pdf.window import build

    doc = PdfDocument(blob, filename, password=password) if password else PdfDocument(blob, filename)
    win = build(doc)
    if not win["pages"]:
        return _locked("ページがない", "ページの木（Catalog の Pages）をたどっても、ファイルを頭から読み直しても、ページが1つも"
                       "見つかりません。窓は空なので、中身を読めません", "ページがないので、書けません")
    reasons: list[Reason] = []
    write: list[Reason] = []
    marks, body = _unmapped(win["pages"])
    if body:
        n = sum(k for _w, k in body)
        reasons.append(Reason("E", "字に直せない本文の字",
                              f"{n} 字は、フォントに文字の対応がなく、字に直せません（窓では �）。その所の中身は、ページを画像として"
                              "見て読みます（todoku の view_page は画像だけを描き、字は描かないので、PDF を描ける別の道具で見ます）。"
                              "中身を推測で埋めないでください", [w for w, _k in body][:12]))
    img = [f"{p['number']}ページ" for p in win["pages"] if p.get("image_only")]
    if img:
        from .pdf.pageimage import unreadable_images

        pages = doc.pages()
        bad = {f"{p['number']}ページ": unreadable_images(doc, pages[p["number"] - 1]) for p in win["pages"] if p.get("image_only")}
        bad = {k: v for k, v in bad.items() if v}
        ok = [x for x in img if x not in bad]
        if ok:
            reasons.append(Reason("E", "画像だけのページ", f"{len(ok)} ページは文字の情報がなく、画像（スキャン）です。"
                                  "エージェントがページの画像（view_page）を見て読みます", ok))
            write.append(Reason("B", "画像だけのページ", f"{len(ok)} ページは字の情報のない紙（スキャン）です。エージェントがページの画像"
                                "（view_page）を見て、書く所を枠（mm）で決めて書きます", ok))
        if bad:
            # 書類はスキャン（画像だけ）で、中身は画像として見て読むもの（E）。開けないのは todoku の圧縮の対応で、書類のせいではない
            codecs = "・".join(sorted({x.split('：')[-1] for v in bad.values() for x in v}))
            reasons.append(Reason("E", "画像だけのページ（todoku が開けない圧縮）",
                                  f"{len(bad)} ページは文字の情報がなく、画像（スキャン）です。その画像の圧縮（{codecs}）を todoku は"
                                  "開けないので、view_page ではページの画像を出せません。PDF を描ける別の道具でページを見て読みます",
                                  list(bad)))
            write.append(Reason("D", "開けない画像だけのページ", f"{len(bad)} ページは、画像の圧縮（{codecs}）を todoku が開けず、"
                                "ページの画像で書く所を確かめられないので、書けません", list(bad)))
    reasons += _pdf_lost(doc, win)
    reasons += _pdf_figures(win)
    if win["form"]["xfa"]:
        xr, xw_ = _xfa_reasons(win.get("xfa"))
        reasons += xr
        write += xw_
    agent = [f for p in win["pages"] for f in p["fields"] if f.get("agent")]
    fields = [f for p in win["pages"] for f in p["fields"] if not f.get("agent")]
    layout = [f for f in fields if f.get("source") == "layout"]
    write += _write_fields(fields, layout, agent, "フォームの欄",
                           "見出しの右や下の空いている所を、見出しの字を手がかりに指定して書きます",
                           "罫線や空きから推定しました（窓には見出しと前後の文字を添えています）")
    g = _final(reasons, write)
    g.alerts = picture_alerts([p for pg in win["pages"] for p in pg.get("pictures", [])]) + _mark_alerts(marks)
    return g


def _pdf_lost(doc, win: dict) -> list[Reason]:
    """PDF の窓の読み落とし：ページの描画の命令を順に読んだ字と、窓に出た字を照らす（audit の check_glyphs。落ちた字・
    重ね書きでないのに1字にした字・元にない字は、ページと字で where に）。字に直せない字・画像だけのページは、別の理由で決める。"""
    from .pdf.audit import Check, check_glyphs

    try:
        c = check_glyphs(doc, win)
    except Exception as e:  # noqa: BLE001  確かめられないものは、通ったことにしない
        c = Check("読み落とし（確かめられない）", False, f"窓の読み落としを確かめられませんでした（{type(e).__name__}: {str(e)[:80]}）")
    return lost_reason([c], "ページ")


def _pdf_figures(win: dict) -> list[Reason]:
    """タグ付きの PDF の、字のあるページの、説明（Alt）のない図（構造の木の要素に結びつき、Artifact でない画像）：中身の図
    （PDF/UA 7.3）なので、エージェントがページを見て確かめる（読む B。where に位置）。画像だけのページは E（別の理由）。
    タグのない PDF の画像は、中身の図か飾りかがファイルの作りから決まらない（大きさの目安で決めることはしない）ので、
    段階は変えず、注意（alerts）に位置と大きさを示す。"""
    figs = [f"{p['id']}（{p['where']}・{p['size_mm'][0]}×{p['size_mm'][1]} mm" + (f"・近くの字「{p['near']}」" if p.get("near") else "") + "）"
            for pg in win["pages"] if not pg.get("image_only")
            for p in pg.get("pictures", []) if p.get("tagged") and not p.get("alt") and not p.get("decorative")]
    if not figs:
        return []
    return [Reason("B", "説明のない図（タグ付きの PDF）",
                   f"字のあるページの図 {len(figs)} 個は、構造の木（タグ）で中身の図とされていますが、説明（Alt）がなく、中身を字として"
                   "読めません。図の所は、エージェントがページを見て読み（PDF を描ける道具で）、中身を推測で埋めないでください", figs[:12])]


def _unmapped(pages: list[dict]) -> tuple[list[tuple[str, int]], list[tuple[str, int]]]:
    """字に直せない字を、行の頭の印（箇条の印）と本文の字に分ける：(所, 字の数) の一覧。行に入らなかった字（窓の作り方で
    1字にした重ね書き・見えない大きさの字を除いて、数の合わないもの）は、本文の字として数える（読めない中身かもしれない）。"""
    marks, body = [], []
    for p in pages:
        if not p.get("unmapped"):
            continue
        n0 = p["number"]
        for m in p.get("unmapped_marks") or []:
            marks.append((f"{n0}ページ・{m['where']}の行「{m['near'][:20]}」の頭", m["count"]))
        for m in p.get("unmapped_text") or []:
            body.append((f"{n0}ページ・{m['where']}の行「{m['near'][:20]}」", m["count"]))
        placed = sum(m["count"] for m in (p.get("unmapped_marks") or []) + (p.get("unmapped_text") or []))
        dropped = sum(t.count("\ufffd") for _k, t in p.get("_dropped") or [])
        rest = p["unmapped"] - placed - dropped
        if rest > 0:
            body.append((f"{n0}ページ（行に入らない字）", rest))
    return marks, body


def _mark_alerts(marks: list[tuple[str, int]]) -> list[Reason]:
    """字に直せない行の頭の印：語は欠けていないので段階は変えない。画面の読み上げでは、印が読めない（PDF/UA 7.21.7）。"""
    if not marks:
        return []
    return [Reason("F", "字に直せない行の頭の印",
                   f"行の頭の印（箇条の印など）{sum(k for _w, k in marks)} 字は、フォントに文字の対応がなく、字に直せません（窓では �）。"
                   "後ろの文の字は欠けていないので、段階は変えません。印の形（・や □ など）を知りたいときは、ページを見て確かめます。"
                   "画面の読み上げを使う人には、印が届きません（PDF/UA 7.21.7）",
                   [w for w, _k in marks][:12])]


def picture_alerts(pictures: list[dict]) -> list[Reason]:
    """説明（代わりの字）のない画像：段階とは別の、アクセシビリティの注意（飾りと印の付いたものは除く）。"""
    bare = [p for p in pictures if not p.get("alt") and not p.get("decorative")]
    if not bare:
        return []
    return [Reason("F", "説明のない画像",
                   f"画像 {len(bare)} 個に、説明（代わりの字）がありません。画面の読み上げを使う人には中身が届かず、"
                   "エージェントも画像の中身は字として読めません（WCAG 1.1.1）。段階は変えません",
                   [f"{p['id']}（{p.get('where', '')}）" for p in bare[:12]])]


def _write_fields(fields: list[dict], layout: list[dict], agent: list[dict], marked: str, free: str,
                  how: str = "空のマスや空白・下線から推定しました") -> list[Reason]:
    """入力欄の書く段階：作り手の印の欄だけなら A、配置からの推定（B）・規則では決められない所（窓の「？」。B）があれば B。
    入力欄の見つけ方は書く段階だけで決める（読む段階は、中身が正確に読めるかで決める）。"""
    out = []
    if layout:
        out.append(Reason("B", "配置から推定した入力欄",
                          f"入力欄 {len(fields)} か所のうち {len(layout)} か所は、{marked}ではなく、{how}", [f["id"] for f in layout[:12]]))
    if agent:
        out.append(_agent_reason(agent))
    if out:
        return out
    if fields:
        return [Reason("A", marked, f"入力欄 {len(fields)} か所は、すべて{marked}で、id を指定してそのまま書けます")]
    return [Reason("B", "印のある欄がない", f"作り手の印のある欄がありません。{free}")]


def word_reasons(win: dict) -> tuple[list[Reason], list[Reason]]:
    """Word と同じ窓の文書（Word・RTF・OpenDocument の文書）の、読む・書くの理由：変更履歴で削除された字が残っている（読む B）、
    書けない入力欄がある（書く D）。"""
    read, write = [], []
    gone = win.get("tracked_deletions") or []
    if gone:
        read.append(Reason("B", "変更履歴", f"変更履歴で削除された字が {len(gone)} 段落に残っています。窓では本文に入れず"
                           "〔変更履歴で削除された字〕と示しています。印刷される字は、ソフトの変更履歴の表示の設定で変わります", gone[:12]))
    bad = [f for f in win.get("fields") or [] if f.get("unwritable")]
    if bad:
        write.append(Reason("D", "書けない入力欄", f"入力欄 {len(bad)} か所は、ファイルの作りのために書けません：{bad[0].get('why', '')}",
                            [f["id"] for f in bad[:12]]))
    return read, write


def _write_embeds(embeds: list[dict]) -> list[Reason]:
    """文書に埋め込んだ Excel の表：表のセルと、表示・印刷の絵を描き直して書く（B）。描き直せない形なら書けない（D）。"""
    out = []
    ok = [e for e in embeds if e.get("fields") and e.get("writable")]
    ng = [e for e in embeds if e.get("fields") and not e.get("writable")]
    if ok:
        out.append(Reason("B", "埋め込んだ表", f"埋め込んだ Excel の表 {len(ok)} 個は、表のセルと、Word が表示・印刷に使う絵を解析して書き直します",
                          [f"E{e['n']}" for e in ok]))
    if ng:
        out.append(Reason("D", "書けない埋め込みの表", "埋め込んだ Excel の表に書けません："
                          + "・".join(f"E{e['n']}（{e.get('why') or '描き直せない形'}）" for e in ng), [f"E{e['n']}" for e in ng]))
    return out


def _write_macros(info) -> list[Reason]:
    """セルの値を変えるマクロ：ボタンなど、画面（UI）から動かす処理（C）。run_macro で、本人に確かめながら動かします。"""
    if not isinstance(info, dict):
        return []
    if not info.get("readable"):
        return [Reason("D", "読めないマクロ", "マクロを読めないので、マクロが書く所を確かめられません")]
    busy = [m for m in info["modules"] if m["procedures"] and not macro_is_display_only(m["source"])]
    if not busy:
        return []
    return [Reason("C", "画面から動かすマクロ", "ボタンや、値を変えると動く処理（マクロ）が、セルや欄を書き換えます。"
                   "Excel のマクロは、問い（はい・いいえ、入力）を本人に確かめながら、run_macro で動かします", [m["name"] for m in busy][:12])]


def _xfa_reasons(xw: dict | None) -> tuple[list[Reason], list[Reason]]:
    """XFA のフォーム：読めれば、XFA であることだけでは段階を下げない。スクリプトと、データに結びつかない欄で決める。"""
    if not xw or not xw.get("readable"):
        why = f"（{xw['error']}）" if xw and xw.get("error") else ""
        return ([Reason("C", "XFA のフォーム", f"Adobe 独自の動くフォーム（XFA）を読めませんでした{why}。開くソフトによって中身が変わります")],
                [Reason("B", "XFA のフォーム", "動くフォーム（XFA）の欄は、フォームの作りを解析して書きます")])
    read, write = [], []
    kind = "動的" if xw["kind"] == "dynamic" else "静的"
    fields = xw["fields"]
    ok = [f for f in fields if not f.get("why_not")]
    if fields:
        read.append(Reason("A", "XFA のフォームの欄", f"XFA の{kind}なフォームを読みました。入力欄 {len(fields)} か所は、すべて XFA の欄です"))
    auto = [s_ for s_ in xw["scripts"] if s_["automatic"]]
    if auto:
        read.append(Reason("C", "XFA のスクリプト",
                           f"自動で動くスクリプト {len(auto)} 種類（開いた時・値を変えた時・計算・確かめなど）は、規則では再現しません。"
                           "値や欄の出し入れを変えるかは、窓のソースを読んで判断します", [" / ".join(s_["where"][:1]) for s_ in auto[:12]]))
    write.append(Reason("B", "XFA のフォーム", f"XFA の欄（書ける欄 {len(ok)} か所）は、フォームの作りを解析して、データ（datasets）に書きます",
                        [f["id"] for f in ok[:12]]))
    ui = [f for f in fields if f.get("why_not") and not f.get("hidden") and f["kind"] in ("text", "number", "date", "checkbox", "radio",
                                                                                             "combo", "list")
          and f.get("no_data")]
    if ui:
        write.append(Reason("C", "データに結びつかない XFA の欄",
                            f"{len(ui)} か所は、データに結びつかない欄（bind が none など）で、書いた値が Adobe で残る書き方が仕様で決まりません。"
                            "Adobe Acrobat・Reader の画面から入力します", [f["id"] for f in ui[:12]]))
    return read, write


def _agent_reason(agent: list[dict]) -> Reason:
    """規則では決められない所（窓の［…？］）：エージェントが窓を読んで判断し、書く所なら id を指定して書く。
    書く所の見つけ方なので、書く段階の B（解析しないと書けない）。読む段階は変えない。"""
    kinds: dict[str, int] = {}
    for f in agent:
        kinds[f.get("why", "")] = kinds.get(f.get("why", ""), 0) + 1
    return Reason("B", "規則では決められない所",
                  f"書き込む所かもしれない所 {len(agent)} か所は、規則では欄と決められません（"
                  + "・".join(f"{k} {n}" for k, n in kinds.items())
                  + "）。エージェントが窓を読んで判断し、書く所なら id を指定して書きます",
                  [f["id"] for f in agent[:12]])


def _grade_word(blob: bytes, filename: str, kind: str) -> Grade:
    from .docx.window import build

    if kind == "doc":
        from .doc.book import DocFile

        doc = DocFile(blob, filename)
        if getattr(doc, "word6", False):
            return _grade_word6(doc)
    else:
        from .docx.document import DocxDocument

        doc = DocxDocument(blob, filename)
    win = build(doc)
    reasons: list[Reason] = []
    write: list[Reason] = []
    agent = [f for f in win["fields"] if f.get("agent")]
    fields = [f for f in win["fields"] if not f.get("agent")]
    layout = [f for f in fields if f.get("source") == "layout"]
    reasons += _macro_reasons(win.get("macros"))
    pics = [p for p in win.get("pictures") or [] if not p.get("decorative")]
    has_text = any(str(t).strip() for st in win.get("stories") or [] for t in st.get("_plain") or [])
    if pics and not has_text and not fields:
        # 字が1字もなく、画像だけの文書（スキャンを貼ったものなど）：中身は画像として見て読む
        reasons.append(Reason("E", "画像だけの文書", f"本文に字がなく、中身は画像 {len(pics)} 個だけです。中身は画像を見て読みます",
                              [p["id"] for p in pics[:12]]))
        write.append(Reason("D", "画像だけの文書", "字の情報がなく、書く所を字で示せないので、書けません（手で書く紙として扱います）"))
    else:
        reasons += content_image_reasons(doc, win, doc.body, lambda p: word_heading(doc, p), "本文")
    reasons += _word_lost(doc, win, kind)
    write += _write_fields(fields, layout, agent, "フォームの欄・コンテンツコントロール",
                           "欄のない所は、前後の字を指定して書き換えます")
    write += _write_embeds(win.get("embeds") or [])
    write += _write_macros(win.get("macros"))
    wr, ww = word_reasons(win)
    reasons += wr
    write += ww
    g = _final(reasons, write)
    g.alerts = picture_alerts(win.get("pictures") or [])
    return g


# 見出し・図表番号の、組み込みのスタイルの名前（ECMA-376 第1部 17.7.4.9 の name。大文字・小文字は区別しない）
_HEADING_STYLES = {f"heading {k}" for k in range(1, 10)} | {"title", "subtitle", "caption"}


def word_heading(doc, p) -> bool:
    """Word の段落が見出し・図表番号か（ファイルの作りで決める：段落か、スタイル（basedOn をたどる）の w:outlineLvl が 0〜8、
    またはスタイルが組み込みの見出し（heading 1〜9）・表題（Title・Subtitle）・図表番号（caption））。字の中身は見ない。"""
    node = getattr(p, "node", None)
    ppr = node.find("pPr") if node is not None and hasattr(node, "find") else None
    ol = ppr.find("outlineLvl") if ppr is not None else None
    if ol is not None and str(ol.get("val", "")).isdigit():
        return int(ol.get("val")) < 9
    sid, seen = getattr(p, "style", "") or "", 0
    names = getattr(doc, "style_names", {}) or {}
    while sid and seen < 20:
        if str(names.get(sid, sid)).strip().lower() in _HEADING_STYLES:
            return True
        lvl = (getattr(doc, "style_outline", {}) or {}).get(sid)
        if lvl is not None:
            return lvl < 9
        sid = (getattr(doc, "style_based", {}) or {}).get(sid, "")
        seen += 1
    return False


def _words(p) -> str:
    """段落の、ファイルの文字の要素から来た字（非表示の字・図の印〔図N…〕などの、todoku が足した印を除く）。"""
    return "".join(s.text for s in getattr(p, "segs", []) if not s.hidden and s.kind in ("t", "sym", "math", "ruby")).strip()


def content_image_reasons(doc, win: dict, body_blocks, heading, body_ja: str) -> list[Reason]:
    """本文の中身である、説明（代わりの字）のない画像：エージェントは画像の中身を字として読めないので、見て確かめる。
    本文の中身か（ヘッダー・フッターの飾りでないか）は、画像のある段落が本文の段落か（表のマス・テキストボックスの中も含む）で、
    飾りかは作り手の印（decorative）で決める（ファイルの作り。大きさや字の数の目安は使わない）。
    - 本文に、見出し・図表番号・空の段落のほかに字のある段落が1つもない：読む E（中身は画像。画像を見て読む）
    - ほかに本文の字がある：読む B（字の中身は窓にあるが、画像の所はエージェントが見て確かめる。where に位置）"""
    from .docx.document import iter_paras

    paras = list(iter_paras(body_blocks))
    ids = {id(p) for p in paras}
    bare = [p for p in getattr(doc, "pictures", []) or []
            if id(p.get("para")) in ids and not p.get("alt") and not p.get("decorative")]
    if not bare:
        return []
    where_of = {p["id"]: p.get("where", "") for p in win.get("pictures") or []}
    where = []
    for p in bare:
        pid = f"図{p['n']}"
        where.append(f"{pid}（{where_of.get(pid, '')}" + (f"・{p['media']}" if p.get("media") else "") + "）")
    text = [p for p in paras if _words(p) and not heading(p)]
    see = ("画像の中身は、todoku は字にしません。エージェントが画像そのもの（ファイルの中の画像の部品。where に名前）を見て読み、"
           "中身を推測で埋めないでください")
    if not text:
        return [Reason("E", "中身が画像の文書",
                       f"{body_ja}には、見出し・図表番号のほかに字の段落がなく、中身は説明（代わりの字）のない画像 {len(bare)} 個です。{see}",
                       where[:12])]
    return [Reason("B", "説明のない本文の画像",
                   f"{body_ja}の画像 {len(bare)} 個に説明（代わりの字）がなく、中身を字として読めません。{see}", where[:12])]


def _word_lost(doc, win: dict, kind: str) -> list[Reason]:
    """Word の窓の読み落とし：XML の文字の要素（.doc はピースの表）から直接拾った字が、窓に出ているか（audit の check_reading）。"""
    from .xlsx.audit import Check

    try:
        if kind == "doc":
            from .doc.audit import check_reading as doc_check

            checks = [doc_check(doc, win)]
        else:
            from .docx.audit import check_reading

            checks = check_reading(doc, win)
    except Exception as e:  # noqa: BLE001  確かめられないものは、通ったことにしない
        checks = [Check("読み落とし（確かめられない）", False, f"窓の読み落としを確かめられませんでした（{type(e).__name__}: {str(e)[:80]}）")]
    return lost_reason(checks, "文書")


def _saved_by(wb) -> str | None:
    """xlsx を保存したアプリ（docProps/app.xml の Application。記録がなければ「記録なし」）。xls は None（分からない）。"""
    import re as _re

    z = getattr(wb, "zip", None)
    if z is None:
        return None
    try:
        xml = z.read("docProps/app.xml").decode("utf-8", "replace")
    except Exception:  # noqa: BLE001
        return "保存したアプリの記録なし"
    m = _re.search(r"<(?:\w+:)?Application>([^<]*)</", xml)
    return m[1].strip() if m and m[1].strip() else "保存したアプリの記録なし"


def _cell_text(v) -> str:
    from .formula.values import ExcelError

    if isinstance(v, ExcelError):
        return v.code
    if v is None or v == "":
        return "（空）"
    return repr(v) if isinstance(v, str) else str(v)


def _unverifiable_rules(wb) -> list[str]:
    """数式を使う入力規則のうち、読めない数式か、計算できない関数を使うもの（シート!範囲：数式）。"""
    from dataclasses import fields as dc_fields

    from .formula.functions import EAGER, LAZY
    from .formula.parser import Func, Node, parse

    def unknown(n) -> list[str]:
        """計算できない関数の名前（なければ空）。"""
        out = []
        if isinstance(n, Func) and n.name not in EAGER and n.name not in LAZY:
            out.append(n.name)
        if isinstance(n, Node):
            for fd in dc_fields(n):
                x = getattr(n, fd.name)
                for y in (x if isinstance(x, list) else [x]):
                    out += unknown(y)
        return out

    out = []
    for sh in wb.sheets:
        for v in sh.validations:
            for f in (v.formula1, v.formula2):
                if not f or (v.type == "list" and f.strip().startswith('"')):
                    continue
                try:
                    bad = unknown(parse(f.lstrip("=")))
                    why = f"（計算できない関数：{'・'.join(dict.fromkeys(bad))}）" if bad else ""
                except Exception:  # noqa: BLE001
                    bad, why = True, "（数式を読めない）"
                if bad:
                    out.append(f"{sh.name}!{v.sqref_text or ''}：={f}{why}")
                    break
    return out


_SHEET_MARKS = {"入力規則があるセル", "シート保護の中で編集できるセル"}  # ファイルの作りとして入力欄の印になっているもの
_SHEET_BOXED = {"枠のある空欄", "色の付いた空欄"}  # 罫線・塗りと、近くの見出しから推定したもの（印ではない）
_CONTROL_KINDS = ("checkbox", "radio", "select", "multiselect", "spin", "option", "combo")


def sheet_input_reasons(doc: dict, reasons: list, write: list, protect: str = "シート保護") -> None:
    """表計算（Excel・OpenDocument）の入力欄の、書く段階（入力欄の見つけ方は読む段階に入れない。reasons は変えない）：
    - A：すべての欄が、ファイルの作りとしての印（入力規則・保護の中のロックの解除・フォームの部品）で決まる
    - B：罫線・塗りと近くの見出し・文字の中の空欄の字の形・数式の入力から推定した欄（印ではない）、規則では決められない所
      （窓の［B5？］・文字の中の［？］。エージェントが判断する）、重なった入力規則（どれが使われるかをエージェントが読む）"""
    marked, boxed, inferred, inline, overlap, agent = [], [], [], [], [], []
    for sh in doc["sheets"]:
        for f in sh.get("fields", []):
            if f.get("agent"):
                agent.append(f)
                continue
            if f.get("unlabeled"):
                continue
            why = set(f.get("reason", []))
            if f["kind"] == "inline":
                inline.append(f["id"])
            elif why & _SHEET_MARKS or ((f["kind"] in _CONTROL_KINDS or f.get("activex")) and not f.get("undecided")):
                marked.append(f["id"])
            elif why & _SHEET_BOXED:
                boxed.append(f["id"])
            else:
                inferred.append(f["id"])
            if f.get("rules_overlap"):
                overlap.append(f["id"])
    n_fields = len(marked) + len(boxed) + len(inferred) + len(inline)
    guessed = inline + boxed + inferred
    if boxed:
        write.append(Reason("B", "枠・塗りから推定した入力欄",
                            f"{len(boxed)} か所は、入力規則も{protect}の印もなく、枠や塗りのある空のマスで、そばに見出しがあることから推定した欄です",
                            boxed[:12]))
    if inline:
        write.append(Reason("B", "文字の中の空欄", f"{len(inline)} か所は、「（　　）」「＿＿＿」のような、文字の中の空欄の形から推定した欄です", inline[:12]))
    if inferred:
        write.append(Reason("B", "印のない入力欄", f"{len(inferred)} か所は、入力規則も枠もなく、数式の入力になっていることから推定した欄です",
                            inferred[:12]))
    if agent:
        write.append(_agent_reason(agent))
    if overlap:
        write.append(Reason("B", "重なった入力規則",
                            f"{len(overlap)} か所に入力規則が重なっていて、どの規則が使われるかはファイルから決まりません。"
                            "窓に出したそれぞれの規則を、エージェントが読んで、合う値を書きます", overlap[:12]))
    if not (guessed or agent or overlap):
        if n_fields:
            write.append(Reason("A", "印のある入力欄", f"入力欄 {n_fields} か所は、すべて入力規則・{protect}・フォームの部品の印で決まり、そのまま書けます"))
        else:
            write.append(Reason("B", "印のある欄がない", f"入力欄の印（入力規則・{protect}など）がありません。表の作りを読んで、書くセルを番地で指定します"))


def image_only_sheets(wb) -> list[str]:
    """画像だけのシート：ファイルの作りだけで決める。絵（picture。グラフは除く：グラフはセルの値から描くもので、中身は値として
    窓に出る）が1つ以上あり、シートのどこにも値・数式のあるセルがなく、字のある図形・部品もないシート。
    印刷範囲は見ない（印刷範囲の外のセルの値も、窓に出る中身なので）。"""
    out = []
    for sh in wb.sheets:
        shapes = getattr(sh, "shapes", [])
        if not any(s.kind == "picture" for s in shapes):
            continue
        if any(c.has_formula or c.value not in (None, "") for c in sh.cells.values()):
            continue
        if any((getattr(s, "text", "") or "").strip() for s in shapes) or \
                any((getattr(c, "text", "") or "").strip() for c in getattr(sh, "controls", [])):
            continue
        out.append(sh.name)
    return out


def _grade_excel(blob: bytes, filename: str) -> Grade:
    from .book import open_book
    from .xlsx.calc import verify
    from .xlsx.window import build

    wb = open_book(blob, filename)
    doc = build(wb)
    reasons: list[Reason] = []
    write: list[Reason] = []
    sheet_input_reasons(doc, reasons, write)
    vr = verify(wb)
    rep = vr.summary()
    if rep["unsupported"]:
        reasons.append(Reason("C", "確かめられない数式", f"{rep['unsupported']} 個の数式は、規則どおりに計算し直して確かめられません"))
    if rep["mismatch"]:
        app = _saved_by(wb)
        where = [f"{r.sheet}!{r.address}：保存 {_cell_text(r.expected)} → Excel の規則で {_cell_text(r.actual)}"
                 for r in vr.results if r.status == "mismatch"][:12]
        if app is None or app.startswith("Microsoft Excel"):
            reasons.append(Reason("C", "保存された値と合わない数式", f"{rep['mismatch']} 個のセルで、保存された計算結果が数式と一致しません", where))
        else:
            # Excel 以外のアプリが保存した計算結果（Google スプレッドシートの書き出しなど）：Excel で開くと計算し直される
            reasons.append(Reason("B", "Excel 以外で保存された計算結果",
                                  f"このファイルは Excel 以外（{app}）で保存されていて、{rep['mismatch']} 個のセルの保存された計算結果が、"
                                  "Excel の規則で計算した値と違います。Excel で開くと、計算し直した値になります（窓に出ているのは保存された値）。"
                                  "読むときは、計算し直した値を使ってください", where))
    unchecked = _unverifiable_rules(wb)
    if unchecked:
        write.append(Reason("B", "確かめられない入力規則",
                            f"{len(unchecked)} 個の入力規則は、todoku が計算できない関数を使っているので、書いた値が規則に合うかを"
                            "確かめられません。窓に出した規則を、エージェントが読んで合う値を書きます", unchecked[:12]))
    if rep["unsupported"]:
        write.append(Reason("B", "確かめられない数式", f"{rep['unsupported']} 個の数式は計算し直せないので、書いた値から計算される所を確かめられません"))
    reasons += _macro_reasons(doc.get("macros"))
    write += _write_macros(doc.get("macros"))
    reasons += _excel_lost(wb, doc)
    for name in image_only_sheets(wb):
        reasons.append(Reason("E", "画像だけのシート", f"シート「{name}」の中身は画像で、セルの文字がありません", [name]))
        write.append(Reason("D", "画像だけのシート", f"シート「{name}」は画像だけで、書くセルがありません。手で書く紙として扱います",
                            [name]))
    pictures = []
    for sh_obj in wb.sheets:
        from .formula.parser import addr

        for n, s in enumerate((s for s in getattr(sh_obj, "shapes", []) if s.kind == "picture"), 1):
            at = f"の {addr(*s.anchor)} 付近" if s.anchor else ""
            pictures.append({"id": f"{sh_obj.name}!{s.name or f'図{n}'}", "alt": (s.alt or "").strip() or None,
                             "decorative": getattr(s, "decorative", False), "where": f"シート「{sh_obj.name}」{at}"})
    # Excel 5.0/95（BIFF5）のブックも、BIFF8 と同じく書けて証明できる（xls/write.py・xls/proof.py）。段階は、ほかの理由だけで決まる
    g = _final(reasons, write)
    g.alerts = picture_alerts(pictures)
    return g


def lost_reason(checks: list, what: str) -> list[Reason]:
    """読み落としの検査（audit の、窓とは別の読み方で拾った中身と窓を照らすもの）が落ちたら、読む C：中身が窓にすべては
    出ていないので、エージェントは窓だけを信じてはいけない（落ちた所は where に）。検査が通れば、何も足さない。"""
    bad = [c for c in checks if not c.ok]
    if not bad:
        return []
    items = [x for c in bad for x in (c.items or [c.detail])]
    return [Reason("C", "窓の読み落とし",
                   f"{what}の中身のうち {len(items)} 件が、窓に出ていないか、元と違って出ています（{'・'.join(c.name for c in bad)}）。"
                   "窓だけを信じず、落ちた所（where）はファイルの別の読み方で確かめます", items[:12])]


def _excel_lost(wb, doc: dict) -> list[Reason]:
    """Excel の窓の読み落とし：XML（.xls はレコード）から直接拾ったセル・図形・メモなどが JSON の窓にあるか（check_reading）と、
    値のあるすべてのセルが、エージェントの読む Markdown の窓に出ているか（check_markdown）。"""
    from .xlsx.audit import Check, check_markdown, check_reading

    links = {(c.link_cell[0], c.link_cell[1], c.link_cell[2]) for s in wb.sheets for c in s.controls if c.link_cell}
    try:
        checks = check_reading(wb, doc) + [check_markdown(doc, links)]
    except Exception as e:  # noqa: BLE001  確かめられないものは、通ったことにしない
        checks = [Check("読み落とし（確かめられない）", False, f"窓の読み落としを確かめられませんでした（{type(e).__name__}: {str(e)[:80]}）")]
    return lost_reason(checks, "ブック")


def _grade_word6(doc) -> Grade:
    """Word 6.0/95：読む段階は Word と同じ決め方。書く段階は、字のバイト数を変えない上書きなので、空白・伏せ字・□ の上に
    書ける欄は B（解析して書く）、字のない欄（空のマスなど）は D（todoku では書けない）。"""
    from .doc.word6 import build

    win = build(doc)
    reasons: list[Reason] = []
    write: list[Reason] = []
    agent = [f for f in win["fields"] if f.get("agent")]
    fields = [f for f in win["fields"] if not f.get("agent")]
    if agent:
        write.append(_agent_reason(agent))
    layout = [f for f in fields if f.get("source") == "layout"]
    if layout:
        write.append(Reason("B", "配置から推定した入力欄",
                            f"入力欄 {len(fields)} か所のうち {len(layout)} か所は、フォームの欄ではなく、空のマスや空白・下線から推定しました",
                            [f["id"] for f in layout[:12]]))
    unread = sum(t.count("状態は読めません") for st in win["stories"] for t in st["reading"])
    if unread:
        reasons.append(Reason("C", "状態の読めないフォームの欄",
                              f"チェックボックス・ドロップダウンのフォームの欄 {unread} か所は、Word 6.0/95 の欄のデータの形が仕様書にないので、"
                              "状態（チェックの有無・選んだ項目）を読んでいません"))
    reasons += _macro_reasons(win.get("macros"))
    ok = [f for f in fields if not f.get("readonly")]
    blocked = [f for f in fields if f.get("readonly")]
    if ok or agent:
        write.append(Reason("B", "同じバイト数の上書き",
                            "Word 6.0/95 の文書は、空白・伏せ字（○○）・□ の上に、そのバイト数（全角の字は 2 バイト）までの値を、文書の文字コードで"
                            f"上書きします（位置・書式の表は変えない）。規則で見つけた欄 {len(ok)} か所に書けます", [f["id"] for f in ok[:12]]))
    if blocked:
        write.append(Reason("D", "字のない欄",
                            f"{len(blocked)} か所は字のない所（空のマスなど）です。字を足すと位置の表・書式の表・ピースの表をずらす必要があり、"
                            "その書き込みと証明をまだ用意していないので、書けません", [f["id"] for f in blocked[:12]]))
    if unread:
        write.append(Reason("D", "状態の読めないフォームの欄", "状態の読めないチェックボックス・ドロップダウンには書けません"))
    return _final(reasons, write)


# ====================================================================== 一太郎（.jtd）
def _grade_jtd(blob: bytes, filename: str) -> Grade:
    """一太郎：読む段階は Word と同じ決め方（欄はすべて配置からの推定）。書く段階は、書ける欄（空白・伏せ字・□ の上書きか、
    本文の空のマス・段落に字を足す・長い値に替える、字数を変える書き込み）は B（解析して書く）、書けない欄（枠・柱の中で
    字が足りない、文書の構造の意味を確かめられないなど）は D（todoku では書けない）。書式の切り替わる所（前の記録と段落の
    終わりの印の書式が違う、空のマスなど）も、足す字を終わりの印の書式で言い直して書くので、書ける欄。
    圧縮して保存した文書も、中の文書に書いて縮め直すので同じ。"""
    from .jtd.book import JtdFile
    from .jtd.edit import policy
    from .jtd.window import build

    doc = JtdFile(blob, filename)
    win = build(doc)
    reasons: list[Reason] = []
    write: list[Reason] = []
    agent = [f for f in win["fields"] if f.get("agent")]
    fields = [f for f in win["fields"] if not f.get("agent")]
    if agent:
        write.append(_agent_reason(agent))
    if fields:
        write.append(Reason("B", "配置から推定した入力欄",
                            f"一太郎の文書には入力欄の印がないので、入力欄 {len(fields)} か所は、空のマスや空白・□ から推定しました",
                            [f["id"] for f in fields[:12]]))
    ok = [f for f in fields if not f.get("readonly")]
    blocked = [f for f in fields if f.get("readonly")]
    grow = [f for f in ok if f["kind"] != "checkbox" and not f.get("maxlen")]
    fixed = [f for f in ok if f.get("maxlen")]
    pol = policy(doc)
    packed = "。圧縮して保存した文書なので、中の文書に書いてから -lh5- で縮め直し、解き直して確かめます" if doc.compressed else ""
    if ok or agent:
        detail = f"規則で見つけた欄 {len(ok)} か所に書けます"
        if grow:
            detail += (f"（うち {len(grow)} か所は、空のマスに字を足す・空白や伏せ字より長い値に替える、字数を変える書き込みが"
                       "できます。文字の書式の区切り・表の位置の表・ストリームの一覧などを合わせて直し、行・ページの割り付けの控えを"
                       "取り除きます）")
        if fixed:
            detail += f"。{len(fixed)} か所は、空白・伏せ字（○○）の字数までの上書きだけです"
        write.append(Reason("B", "解析して書く（字数を変える書き込みを含む）" if grow else "同じ字数の上書き",
                            detail + packed, [f["id"] for f in ok[:12]]))
    if blocked:
        whys = sorted({f.get("why_fixed", "") for f in blocked} - {""})
        write.append(Reason("D", "書けない欄",
                            f"{len(blocked)} か所は書けません（上書きできる字がなく、字数を変える書き込みもできない所）。理由："
                            + "／".join(whys[:4]), [f["id"] for f in blocked[:12]]))
    if not (ok or agent or blocked):
        write.append(Reason("B", "印のある欄がない", "欄のない所は、前後の字を目印にして書きます"
                            + ("（本文は字数が違ってもよい）" if pol.ok else "（同じ字数で上書き）")))
    return _final(reasons, write)
