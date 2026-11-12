"""大きな .docx の道の変更履歴：ふつうの道と同じ段落の一覧（削除された字の残る段落・挿入された字のある段落）と、同じ段階の理由。

前は、大きな .docx の道は、本文の w:delText・w:ins の要素を数えて「n か所」と伝えていた：ふつうの道（段落の数と、段落の番号・字の頭）
と数え方が違い、移ってきた字（w:moveTo）と、ヘッダー・脚注などの変更履歴は数えていなかった。
いまは、変更履歴の要素のあるまとまりだけを、初めて要るときに 1 回、小さな文書にして、ふつうの道と同じ数え方で数える。
"""
from test_docx_large_normal_color import PLAIN, _docx, _p, large  # noqa: F401（large は fixture）
from todoku.docx.document import DocxDocument
from todoku.docx.window import build
from todoku.grade import word_reasons

A1 = 'w:author="田中 一郎" w:date="2026-01-05T10:30:00Z"'


def _r(t: str, rpr: str = "") -> str:
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'


def _ins(*runs: str, tag: str = "ins") -> str:
    return f'<w:{tag} w:id="1" {A1}>{"".join(runs)}</w:{tag}>'


def _del(t: str) -> str:
    return f'<w:del w:id="2" {A1}><w:r><w:delText xml:space="preserve">{t}</w:delText></w:r></w:del>'


def _para(*inner: str, ppr: str = "") -> str:
    return f'<w:p><w:pPr>{ppr}</w:pPr>{"".join(inner)}</w:p>'


def _plain(n: int, k: int = 0) -> str:
    return "".join(_p(f"ふつうの段落 {i}。") for i in range(k, k + n))


def _normal(blob: bytes) -> tuple[dict, list]:
    """ふつうの道：窓と、変更履歴の理由。"""
    win = build(DocxDocument(blob, "a.docx"))
    return win, _reasons(word_reasons(win)[0])


def _reasons(rs) -> list:
    return [(r.grade, r.title, r.detail, list(r.where)) for r in rs if r.title == "変更履歴"]


def _no_chunks(doc, monkeypatch) -> list:
    calls = []
    real = doc.chunk_doc
    monkeypatch.setattr(doc, "chunk_doc", lambda i, real_=(): calls.append(i) or real(i, real_))
    return calls


def test_revised_paragraphs_are_listed_like_the_normal_path(large):
    body = (_plain(12) + _para(_r("提出先："), _ins(_r("市役所")), _r("の"), _ins(_r("窓口")), _r("。"))  # 1 つの段落に 2 つの挿入
            + _plain(12, 12) + _para(_r("電話は"), _del("町役場"), _r("まで。")) + _para(_del("消した段落。"))
            + _plain(12, 24) + _para(_ins(_r("足した段落。"))) + _para(_r("両方："), _ins(_r("足す")), _del("消す")))
    blob = _docx(body, PLAIN)
    win, want = _normal(blob)
    assert win["tracked_insertions"] == ["段落 13「市役所窓口」", "段落 40「足した段落。」", "段落 41「足す」"]
    assert win["tracked_deletions"] == ["段落 26「町役場」", "段落 27「消した段落。」", "段落 41「消す」"]
    doc = large(blob)
    assert len(doc.chunks) > 5
    assert doc.revisions() == (win["tracked_deletions"], win["tracked_insertions"], [])
    got = _reasons(doc.grade().reasons)
    assert got == want and len(got) == 2
    assert "3 段落" in got[0][2] and "3 段落" in got[1][2] and not any("か所" in g[2] for g in got)


def test_moved_text_is_counted(large):
    # 移ってきた字（w:moveTo）は挿入、移る前の字（w:moveFrom：字は w:t）は削除
    body = (_plain(10) + _para(_r("先頭。"), _ins(_r("移ってきた字"), tag="moveTo"))
            + _plain(10, 10) + _para(f'<w:moveFrom w:id="3" {A1}>{_r("移る前の字")}</w:moveFrom>', _r("残る字。")))
    blob = _docx(body, PLAIN)
    win, want = _normal(blob)
    assert win["tracked_insertions"] == ["段落 11「移ってきた字」"] and win["tracked_deletions"] == ["段落 22「移る前の字」"]
    doc = large(blob)
    assert doc.revisions() == (win["tracked_deletions"], win["tracked_insertions"], win["tracked_marks"])
    assert _reasons(doc.grade().reasons) == want


def test_revisions_in_a_table_and_hidden_inserted_text(large):
    cell = lambda inner: f"<w:tc>{inner}</w:tc>"  # noqa: E731
    table = "<w:tbl><w:tr>" + cell(_p("見出し")) + cell(_para(_r("値："), _ins(_r("足した値")))) + "</w:tr></w:tbl>"
    body = (_plain(10) + table + _plain(10, 10) + _para(_r("見える字。"), _ins(_r("隠した字", "<w:vanish/>")))
            + _para(_r("空の挿入。"), _ins()) + _para(_ins(_r("　"))))
    blob = _docx(body, PLAIN)
    win, want = _normal(blob)
    assert win["tracked_insertions"][0] == "段落 12「足した値」" and not any("隠した字" in x for x in win["tracked_insertions"])
    doc = large(blob)
    assert doc.revisions() == (win["tracked_deletions"], win["tracked_insertions"], win["tracked_marks"])
    assert _reasons(doc.grade().reasons) == want


def test_revisions_outside_the_body_are_counted(large):
    # ヘッダーの変更履歴も、ふつうの道と同じ段落の番号で数える（前は、本文だけを数えていた）
    header = _para(_r("様式第 1 号"), _ins(_r("（改）"))) + _para(_del("古い見出し"), _r("新しい見出し"))
    blob = _docx(_plain(15), PLAIN, header)
    win, want = _normal(blob)
    assert win["tracked_insertions"] == ["段落 16「（改）」"] and win["tracked_deletions"] == ["段落 17「古い見出し」"]
    doc = large(blob)
    assert not doc._rev_chunks  # 本文には、変更履歴の要素がない
    assert doc.revisions() == (win["tracked_deletions"], win["tracked_insertions"], win["tracked_marks"])
    assert _reasons(doc.grade().reasons) == want and len(want) == 2


def _mark(kind: str) -> str:
    return f'<w:rPr><w:{kind} w:id="5" {A1}/></w:rPr>'


def _row(kind: str, *cells: str) -> str:
    return f'<w:tr><w:trPr><w:{kind} w:id="6" {A1}/></w:trPr>' + "".join(f"<w:tc>{c}</w:tc>" for c in cells) + "</w:tr>"


WHO = "変更した人：田中 一郎・日時：2026-01-05T10:30:00Z"


def test_inserted_and_deleted_paragraph_breaks_are_shown(large):
    """段落の区切り（段落の終わりの印）だけの挿入・削除：字は変わらないが、段落の分かれ方が変わる。前は、示さなかった。"""
    body = (_para(_r("前の半分。"), ppr=_mark("ins")) + _para(_r("後ろの半分。"))  # 段落の途中で改行した
            + _para(_r("つながる前。"), ppr=_mark("del")) + _para(_r("つながる後ろ。"))  # 段落の区切りを消した
            + _para(_ins(_r("まるごと足した段落。")), ppr=_mark("ins"))  # 段落ごとの挿入：字の印が伝えている
            + _para(_del("まるごと消した段落。"), ppr=_mark("del"))  # 段落ごとの削除：字の印が伝えている
            + _para(_r("移った区切り。"), ppr=_mark("moveTo")) + _para(ppr=_mark("ins")) + _plain(20))
    blob = _docx(body, PLAIN)
    win, want = _normal(blob)
    lines = win["stories"][0]["reading"]
    assert lines[:7] == [
        f"前の半分。〔変更履歴で挿入された段落の区切り（まだ承諾も却下もされていない。{WHO}）〕",
        "後ろの半分。",
        f"つながる前。〔変更履歴で削除された段落の区切り（いまの本文では、次の段落に続く。{WHO}）〕",
        "つながる後ろ。",
        f"〔変更履歴で挿入された段落（まだ承諾も却下もされていない。{WHO}）〕まるごと足した段落。",
        f"〔変更履歴で削除された字（いまの本文ではない。{WHO}）: まるごと消した段落。〕",
        f"移った区切り。〔変更履歴で挿入された段落の区切り（まだ承諾も却下もされていない。{WHO}）〕",
    ]
    assert win["tracked_marks"] == ["段落 1：挿入された段落の区切り", "段落 3：削除された段落の区切り", "段落 7：挿入された段落の区切り",
                                    "段落 8：挿入された段落の区切り"]
    assert any("段落の区切り・表の行" in d and "が 4 か所" in d for _g, _t, d, _w in want)
    doc = large(blob)
    assert doc.revisions() == (win["tracked_deletions"], win["tracked_insertions"], win["tracked_marks"])
    assert _reasons(doc.grade().reasons) == want and len(want) == 3
    got = [ln for b in doc.read({}, 4000).body for ln in b.split("\n")]
    assert [ln.split("〕", 1)[1] for ln in got[:7]] == lines[:7]  # 頭の〔段落 n〕を除く


def test_inserted_and_deleted_table_rows_are_shown(large):
    """表の行の挿入・削除（w:trPr の中の印）：行の頭の段落の前に示す。前は、示さなかった。"""
    table = ("<w:tbl>" + f"<w:tr><w:tc>{_p('氏名')}</w:tc><w:tc>{_p('住所')}</w:tc></w:tr>"
             + _row("ins", _para(_ins(_r("足した行"))), _para()) + _row("del", _para(_del("消した行")), _para())
             + _row("ins", "<w:tbl>" + f"<w:tr><w:tc>{_p('入れ子')}</w:tc></w:tr>" + "</w:tbl>" + _para(), _p("右")) + "</w:tbl>")
    blob = _docx(_plain(10) + table + _plain(10, 10), PLAIN)
    win, want = _normal(blob)
    text = "\n".join(win["stories"][0]["reading"])
    assert f"〔変更履歴で挿入された表の行（まだ承諾も却下もされていない。{WHO}）〕〔変更履歴で挿入された段落" in text
    assert f"〔変更履歴で削除された表の行（いまの本文ではない。{WHO}）〕〔変更履歴で削除された字" in text
    assert f"〔変更履歴で挿入された表の行（まだ承諾も却下もされていない。{WHO}）〕入れ子" in text  # 行の頭の段落（入れ子の表の中）
    assert text.count("表の行（") == 3
    assert win["tracked_marks"] == ["段落 13：挿入された表の行", "段落 15：削除された表の行", "段落 17：挿入された表の行"]
    doc = large(blob)
    assert doc.revisions() == (win["tracked_deletions"], win["tracked_insertions"], win["tracked_marks"])
    assert _reasons(doc.grade().reasons) == want and len(want) == 3


def test_a_document_without_tracked_marks_reads_as_before(large):
    blob = _docx(_plain(8) + f"<w:tbl><w:tr><w:tc>{_p('マス')}</w:tc></w:tr></w:tbl>", PLAIN)
    win, want = _normal(blob)
    assert win["tracked_marks"] == [] and want == [] and "変更履歴" not in "\n".join(win["stories"][0]["reading"])


def test_a_document_without_revisions_is_not_counted(large, monkeypatch):
    doc = large(_docx(_plain(40), PLAIN))
    calls = _no_chunks(doc, monkeypatch)
    assert doc.revisions() == ([], [], []) and _reasons(doc.grade().reasons) == [] and calls == []


def test_only_the_chunks_with_revisions_are_read_and_only_once(large, monkeypatch):
    blob = _docx(_plain(30) + _para(_r("ここだけ："), _ins(_r("足した字"))) + _plain(30, 30), PLAIN)
    doc = large(blob)
    body = [i for i, ch in enumerate(doc.chunks) if "story" not in ch]
    assert len(body) > 6 and len(doc._rev_chunks) == 1
    calls = _no_chunks(doc, monkeypatch)
    assert doc.revisions()[1] == ["段落 31「足した字」"]
    assert calls == sorted(doc._rev_chunks)
    doc.grade()
    doc.revisions()
    assert len(calls) == 1
