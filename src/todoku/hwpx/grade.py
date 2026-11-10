"""한글の文書の段階：Word・OpenDocument と同じ決め方（grade.py の定義）に、한글に特有の条件を足す。

- HWPX：누름틀・フォームの部品は作り手の印の欄（書く A）、空のマス・空欄は配置からの推定（書く B）。한글のマクロ（JScript）は
  規則では再現しない処理（読む C・書く C）
- HWP 5：読むだけ（書く D：HWP 5 には書けません。HWPX で保存し直せば書ける）。配布用の文書（본문を暗号化した ViewText）と、
  パスワードで保護した文書は、中身を読めない（読む F・書く D）
"""

from __future__ import annotations


def _worst(order: str, reasons) -> str:
    return max((r.grade for r in reasons), key=order.index, default=order[0])


def grade(k: str, blob: bytes, name: str, textdoc):
    from ..grade import WRITE_ORDER, Reason, _locked

    if k == "hwpx":
        from .document import HwpxDocument
        from .window import build

        doc = HwpxDocument(blob, name)
        win = build(doc)
        extra = []
        if doc.scripts:
            extra.append(Reason("C", "한글のマクロ", f"文書のマクロ（JScript）{len(doc.scripts)} 個は、規則では再現しません。"
                                "文書を開く・閉じるときに動く処理が要るかは、エージェントがソースを読んで判断します",
                                [n for n, _s in doc.scripts]))
        g = textdoc(win, "한글の文書", "누름틀・フォームの部品", extra, doc)
        if doc.scripts:
            g.write_reasons.append(Reason("C", "한글のマクロ", "文書のマクロ（JScript）が動いて欄を書き換えるかは、規則では分かりません。"
                                          "要る処理なら、本人に한글で開いて動かしてもらいます", [n for n, _s in doc.scripts]))
            g.write = _worst(WRITE_ORDER, g.write_reasons)
            g.write_reasons.sort(key=lambda r: -WRITE_ORDER.index(r.grade))
        return g
    from .hwp5 import Hwp5Document, Hwp5Locked
    from .window import build5

    try:
        doc = Hwp5Document(blob, name)
    except Hwp5Locked as e:
        return _locked(e.title, str(e), "読めないので、書けません")
    win = build5(doc)
    extra = []
    if doc.numbered:
        extra.append(Reason("B", "段落の番号", f"番号・箇条書きの段落 {doc.numbered} 個の番号の字（문단 번호・개요 번호）は、HWP 5 では窓に"
                            "出していません。番号を読むときは、HWPX で保存し直した文書を読みます"))
    g = textdoc(win, "한글の文書（HWP 5）", "누름틀・フォームの部品", extra, doc)
    g.write_reasons = [Reason("D", "HWP 5 には書けない", "HWP 5（.hwp）には書けません。HWPX で保存し直せば書ける"
                              "（한글の「다른 이름으로 저장」で .hwpx を選ぶ）")]
    g.write = "D"
    return g
