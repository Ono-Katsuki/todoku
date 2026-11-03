"""形式を問わない入口（api・grade）から呼ぶ、追加の形式：CSV・TSV・JSON・JSON Lines・RTF・OpenDocument（.odt・.ods）・
文字だけのファイル（.txt）・PowerPoint（.pptx）。

api.py と grade.py には、この部品を呼ぶ短い区切りだけを足す（ほかの形式の処理には手を加えない）。
- 窓：window（窓の JSON と Markdown の部分）
- 入力欄：field_rows（api.fields が、欄ごとに _field を通す）
- 書き込み：fill（書いたファイル・証明・読み戻しに使う id）と read_back（書いた後のファイルから、書いた所の値を読む）
- 検査：audit_runner、段階：grade
"""

from __future__ import annotations

KIND_JA = {"csv": "CSV", "tsv": "TSV（タブ区切り）", "json": "JSON", "jsonl": "JSON Lines", "rtf": "RTF（リッチテキスト）",
           "odt": "OpenDocument の文書（.odt）", "ods": "OpenDocument の表計算（.ods）",
           "fodt": "OpenDocument の文書（.fodt・1つの XML）", "fods": "OpenDocument の表計算（.fods・1つの XML）",
           "txt": "テキスト（.txt）", "pptx": "PowerPoint（.pptx）",
           "hwpx": "한글の文書（.hwpx）", "hwp": "한글の文書（.hwp・HWP 5）",
           "markup": "HTML・XML（字だけを読む）"}
KINDS = frozenset(KIND_JA)
DATA = frozenset({"csv", "tsv", "json", "jsonl"})
TEXT_DOCS = frozenset({"odt", "fodt"})
SHEETS = frozenset({"ods", "fods"})
HANGUL = frozenset({"hwpx", "hwp"})  # 한글：HWPX は読む・書く、HWP 5 は読むだけ（hwpx/）


class OpenFormatError(ValueError):
    pass


def _err(e: Exception) -> OpenFormatError:
    return OpenFormatError(str(e))


# ====================================================================== 窓
TEXT_FILES = DATA | {"txt", "markup"}  # 文字だけのファイル：文字コードを textenc で決める（encoding を渡せる）


def window(k: str, blob: bytes, name: str, encoding: str | None = None) -> tuple[dict, list[tuple[str, str]]]:
    """encoding：文字だけのファイル（TEXT_FILES）の文字コードを、エージェントが選んで指定したとき（ほかの形式では使わない）。"""
    try:
        if k in DATA:
            from .datafile import window as dw

            doc = dw.build(blob, name, k, encoding)
            return doc, dw.markdown_parts(doc)
        if k in TEXT_DOCS:
            from .odf.text import OdtDocument
            from .odf.textwindow import build, markdown_parts

            win = build(OdtDocument(blob, name))
            return win, markdown_parts(win)
        if k in SHEETS:
            from .odf.sheetwindow import build, markdown_parts

            win = build(blob, name)
            return win, markdown_parts(win)
        if k == "rtf":
            from .rtf.document import RtfDocument
            from .rtf.window import build, markdown_parts

            win = build(RtfDocument(blob, name))
            return win, markdown_parts(win)
        if k == "txt":
            from .txt.document import TxtDocument
            from .txt.window import build, markdown_parts

            win = build(TxtDocument(blob, name, encoding))
            return win, markdown_parts(win)
        if k == "pptx":
            from .pptx.document import PptxDocument
            from .pptx.window import build, markdown_parts

            win = build(PptxDocument(blob, name))
            return win, markdown_parts(win)
        if k == "hwpx":
            from .hwpx.document import HwpxDocument
            from .hwpx.window import build, markdown_parts

            win = build(HwpxDocument(blob, name))
            return win, markdown_parts(win)
        if k == "hwp":
            from .hwpx.hwp5 import Hwp5Document
            from .hwpx.window import build5, markdown_parts

            win = build5(Hwp5Document(blob, name))
            return win, markdown_parts(win)
        if k == "markup":
            from . import markup

            win = markup.build(blob, name, encoding)
            return win, markup.markdown_parts(win)
    except (ValueError, KeyError) as e:
        raise _err(e) from None
    raise OpenFormatError(f"この形式（{k}）は扱えません")


def field_rows(k: str, doc: dict):
    """(欄, どこか)。表計算はシートの名前を添える。"""
    if k in SHEETS:
        for sh in doc["sheets"]:
            for f in sh.get("fields", []):
                yield f, {"sheet": sh["name"]}
        return
    for f in doc.get("fields", []):
        yield f, {}


def summary(k: str, doc: dict) -> dict:
    d: dict = {}
    if k in SHEETS:
        d["sheets"] = [s["name"] for s in doc["sheets"]]
        m = doc.get("macros")
        d["macros"] = bool(m)
    elif k in DATA:
        d["format"] = doc.get("format")
        if k in ("csv", "tsv"):
            d["rows"] = doc.get("row_count")
            d["columns"] = doc.get("width")
    elif k == "txt":
        d["format"] = doc.get("format")
    elif k == "pptx":
        d["slides"] = len(doc.get("slides") or [])
    elif k == "markup":
        d["format"] = doc.get("format")
        d["read_only"] = doc.get("note")
    else:
        d["protection"] = doc.get("protection")
    return d


# ====================================================================== 書き込み
def fill(k: str, blob: bytes, name: str, values: dict, force: bool = False, encoding: str | None = None) -> dict:
    """{"blob": 書いた中身, "result": api の結果に足すもの, "readback": 読み戻しで照らす id → 値, "labels": id → 見出し}。"""
    if k in ("csv", "tsv"):
        return _fill_csv(k, blob, name, values, encoding)
    if k in ("json", "jsonl"):
        return _fill_json(k, blob, name, values, encoding)
    if k in TEXT_DOCS:
        return _fill_odt(blob, name, values, force)
    if k in SHEETS:
        from .odf.sheetfill import api_fill

        return api_fill(blob, name, values, force)
    if k == "rtf":
        from .rtf.fill import api_fill

        return api_fill(blob, name, values, force)
    if k == "txt":
        from .txt.fill import api_fill

        return api_fill(blob, name, values, force, encoding)
    if k == "pptx":
        from .pptx.fill import api_fill

        return api_fill(blob, name, values, force)
    if k == "hwpx":
        return _fill_hwpx(blob, name, values, force)
    if k == "markup":
        raise OpenFormatError("HTML・XML には書きません（とどくは字だけを読みます。タグのある字への書き込みを証明する方法を"
                              "作っていません）。書く必要があれば、本人に、Word・Excel・PDF など書ける形式の様式があるかを確かめてください")
    if k == "hwp":
        raise OpenFormatError("HWP 5（.hwp）には書けません。HWPX で保存し直せば書ける（한글の「다른 이름으로 저장」で .hwpx を選ぶ。"
                              "本人に、HWPX で保存し直したファイルを頼んでください）")
    raise OpenFormatError(f"この形式（{k}）には書けません")


def _fill_csv(k: str, blob: bytes, name: str, values: dict, encoding: str | None = None) -> dict:
    from .datafile.csvfile import CsvError, CsvFile, fill, prove

    f = CsvFile(blob, name, k, encoding)
    try:
        res = fill(f, values)
    except CsvError as e:
        raise _err(e) from None
    written = {(p.row, p.col): p.value for p in res.patches}
    pr = prove(blob, res.blob, written, name, k, encoding)
    head = f.header
    labels = {}
    for p in res.patches:
        col = head[p.col - 1].strip() if p.col <= len(head) and head[p.col - 1].strip() else f"C{p.col}"
        labels[p.key] = f"{col}（{p.row}行目" + ("・足した行" if p.row > len(f.records) else "") + "）"
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20],
              "cells": {p.key: p.id for p in res.patches if p.key != p.id}}
    if res.appended_rows:
        result["appended_rows"] = res.appended_rows
    return {"blob": res.blob, "result": result, "readback": {p.key: p.value for p in res.patches}, "labels": labels}


def _fill_json(k: str, blob: bytes, name: str, values: dict, encoding: str | None = None) -> dict:
    from .datafile.jsonfile import JsonError, JsonFile, fill, prove

    f = JsonFile(blob, name, k, encoding)
    try:
        res = fill(f, values)
    except JsonError as e:
        raise _err(e) from None
    pr = prove(blob, res.blob, {p.id: p.value for p in res.patches}, name, k, encoding)
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20], "written_as": {p.key: p.text for p in res.patches}}
    added = {p.key: p.id for p in res.patches if p.added}
    if added:
        result["added"] = added  # 足した名前・値：書いた所の id
    labels = {p.key: (p.id.rsplit("/", 1)[-1] or p.id) for p in res.patches}
    return {"blob": res.blob, "result": result, "readback": {p.key: p.value for p in res.patches}, "labels": labels,
            "ids": {p.key: p.id for p in res.patches}}


def _fill_odt(blob: bytes, name: str, values: dict, force: bool) -> dict:
    from .docx.fill import FillError
    from .odf.text import OdtDocument
    from .odf.textfill import fill, written_of
    from .odf.textproof import prove

    try:
        res = fill(OdtDocument(blob, name), values, force=force)
    except FillError as e:
        raise _err(e) from None
    pr = prove(blob, res.blob, written_of(res.patches))
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20]}
    readback = {key: v for key, v in values.items() if not key.startswith("@")}
    anchored = {p.key: p.id for p in res.patches if p.anchor}
    if anchored:
        result["anchored"] = anchored
        readback.update({p.id: p.value for p in res.patches if p.anchor})
    # 欄の書式で書いた値（日付の欄の書式「yyyy年M月d日」など）：読み戻しは、書式を通した字で比べる
    shaped = {p.id: p.value for p in res.patches if not p.anchor and p.kind == "input" and p.id in values
              and not isinstance(values[p.id], bool) and str(p.value) != str(values[p.id])}
    if shaped:
        result["written_as"] = shaped
        readback.update(shaped)
    return {"blob": res.blob, "result": result, "readback": readback, "labels": {}}


def _fill_hwpx(blob: bytes, name: str, values: dict, force: bool) -> dict:
    from .docx.fill import FillError
    from .hwpx.document import HwpxDocument
    from .hwpx.fill import fill, written_of
    from .hwpx.proof import prove

    circ = {k: v for k, v in values.items() if isinstance(v, dict) and "circle" in v}  # 字を囲む（字の囲み線）
    plain = {k: v for k, v in values.items() if k not in circ}
    try:
        res = fill(HwpxDocument(blob, name), plain, force=force)
    except FillError as e:
        raise _err(e) from None
    pr = prove(blob, res.blob, written_of(res.patches))
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20],
              "relayout": "字を変えた段落は、한글の行の配置の控え（hp:linesegarray）を除きました。한글が開くときに、その段落の行を配置し直します"}
    readback = {key: v for key, v in plain.items() if not key.startswith("@")}
    anchored = {p.key: p.id for p in res.patches if p.anchor}
    if anchored:
        result["anchored"] = anchored
        readback.update({p.id: p.value for p in res.patches if p.anchor})
    out = res.blob
    if circ:
        from .docx.circle import CircleError
        from .hwpx import circle as hc

        doc = HwpxDocument(out, name)
        used: dict = {}
        try:
            sds = [(k, hc.resolve(doc, spec, used)) for k, spec in circ.items()]
            boxed = hc.apply(doc, [sd for _k, sd in sds])
        except CircleError as e:
            raise _err(e) from None
        bad = hc.prove(out, boxed, name, [sd for _k, sd in sds], list(circ.values()))
        out = boxed
        result["unexpected"] = result["unexpected"] + bad
        result["proof_ok"] = result["proof_ok"] and not bad
        result["proof"] += (f"。囲み {len(sds)} つは、字を字の並びに分けて囲み線の形（header.xml に足した charPr・borderFill）を"
                            "付けただけ（段落の字は同じ）" if not bad else f"。囲みで想定外の変更 {len(bad)} 件")
        result["circled"] = [{"key": k, "id": sd["id"], "text": sd["text"], "label": sd["label"], "context": sd["context"],
                              "style": "border", "how": "字に四角い囲み線を付けた（○ではなく□）", "uncertain": hc.UNCERTAIN}
                             for k, sd in sds]
        result["warnings"] = result["warnings"] + [hc.UNCERTAIN]
        result["anchored"] = {**result.get("anchored", {}), **{k: sd["id"] for k, sd in sds}}
    return {"blob": out, "result": result, "readback": readback, "labels": {}}


def read_back(k: str, blob: bytes, name: str, keys: dict, encoding: str | None = None) -> dict:
    """書いた後のファイルから、書いた所（keys の id）の値を読む（欄の一覧に出ないマス・値のため）。"""
    out: dict = {}
    try:
        if k in ("csv", "tsv"):
            from .datafile.csvfile import CsvFile

            f = CsvFile(blob, name, k, encoding)
            for key in keys:
                try:
                    r, c = f.resolve(key)
                except ValueError:
                    continue
                out[key] = f.value(r, c)
        elif k in ("json", "jsonl"):
            from .datafile.jsonfile import JsonFile, _plain

            f = JsonFile(blob, name, k, encoding)
            for key in keys:
                try:
                    node, _parent, _last, _d = f.find(key)
                except ValueError:
                    continue
                if node is not None:
                    v = _plain(node)
                    out[key] = v if not isinstance(v, (dict, list)) else keys[key] if v == keys[key] else v
        elif k in SHEETS:
            from .odf.sheetfill import read_cells

            out.update(read_cells(blob, name, keys))
    except ValueError:
        return out
    return out


# ====================================================================== 検査・段階
def audit_runner(k: str, encoding: str | None = None):
    if k in DATA:
        from .datafile.audit import run

        return lambda blob, name: run(blob, name, k, encoding)
    if k in TEXT_DOCS:
        from .odf.textaudit import run

        return run
    if k in SHEETS:
        from .odf.sheetaudit import run

        return run
    if k == "rtf":
        from .rtf.audit import run

        return run
    if k == "txt":
        from .txt.audit import run

        return lambda blob, name: run(blob, name, encoding)
    if k == "pptx":
        from .pptx.audit import run

        return run
    if k == "hwpx":
        from .hwpx.audit import run

        return run
    if k == "hwp":
        from .hwpx.audit import run5

        return run5
    if k == "markup":
        from .markup import audit

        return audit
    raise OpenFormatError(f"この形式（{k}）は検査できません")


def grade(k: str, blob: bytes, name: str, encoding: str | None = None):
    from .grade import _locked

    try:
        if k in DATA:
            return _grade_data(k, blob, name, encoding)
        if k in TEXT_DOCS:
            from .odf.text import OdtDocument
            from .odf.textwindow import build

            return _grade_textdoc(build(OdtDocument(blob, name)), "OpenDocument の文書", "フォームの部品・入力の欄")
        if k in SHEETS:
            from .odf.sheetwindow import grade_sheets

            return grade_sheets(blob, name)
        if k == "rtf":
            from .rtf.document import RtfDocument
            from .rtf.window import build

            return _grade_textdoc(build(RtfDocument(blob, name)), "RTF", "フォームの欄")
        if k == "txt":
            return _grade_txt(blob, name, encoding)
        if k == "markup":
            from .markup import grade as mgrade

            return mgrade(blob, name, encoding)
        if k == "pptx":
            from .pptx.document import PptxDocument
            from .pptx.window import build

            from .grade import content_image_reasons

            doc = PptxDocument(blob, name)
            win = build(doc)
            # スライドの画像は、スライドの中身（スライドの型の絵は数えない）。見出しは、タイトルのプレースホルダーの段落
            slides = [b for sl in doc.slides for b in sl["blocks"]]
            extra = content_image_reasons(doc, win, slides, lambda p: p.style == "Title", "スライド",
                                          pages=[sl["blocks"] for sl in doc.slides])
            return _grade_textdoc(win, "PowerPoint", "作り手の印のある欄", extra)
        if k in HANGUL:
            from .hwpx.grade import grade as hgrade

            return hgrade(k, blob, name, _grade_textdoc)
    except Exception as e:  # noqa: BLE001
        from .hwpx.package import HwpxEncrypted
        from .odf.package import OdfEncrypted

        if isinstance(e, HwpxEncrypted):
            return _locked("暗号化", str(e), "暗号化された部品は読めないので、書けません")
        if isinstance(e, OdfEncrypted):
            return _locked("パスワード", str(e), "パスワードがなければ、書けません")
        return _locked("開けない", f"ファイルを読めませんでした（{type(e).__name__}: {str(e)[:80]}）", "読めないので、書けません")
    return _locked("対応していない形式", f"この形式（{k}）は、まだ読めません", "読めないので、書けません")


def encoding_reason(dec, base: str = "S"):
    """文字コードの、読む段階と書く段階の理由：(読む, 書く)。規則で決まっていれば (None, None)。
    - どの文字コードとしても読めない（Latin-1 で1字ずつ読んだ）：読む C（窓の字を信じられない。エージェントが確かめる）
    - いくつもの文字コードで、仕様どおりにきれいに読める：読む B（中身はすべて窓に出ていて、候補ごとの読み方の見本も窓に
      並べた。どれが正しいかを、窓に示した根拠でエージェントが決める）・書く B（文字コードを決めてから書く）
    - エージェントが候補から選んで encoding で指定した：その文字コードで厳密に読めることを確かめたので、段階は下げない
      （base：S か A）。理由に、選んだ文字コードと、選んだのがエージェントであることを残す（出どころ agent_choice）"""
    from .datafile.textenc import NAME_JA
    from .grade import Reason

    name = NAME_JA.get(dec.encoding, dec.encoding)
    if dec.basis == "given":
        r = Reason(base, "エージェントが選んだ文字コード",
                   f"文字コードは、エージェントが選んで指定した {name}（encoding: \"{dec.encoding}\"）で読みました。"
                   "このファイルが、その文字コードとして厳密に読めることを確かめました（どれが正しいかは、規則ではなくエージェントの判断）",
                   [dec.encoding], by="agent_choice")
        w = Reason("A", "エージェントが選んだ文字コード",
                   f"書く値は、選んだ文字コード（{name}）のバイトにして、書いた所だけを置き換えます（ほかのバイトは元のまま）",
                   [dec.encoding], by="agent_choice")
        return r, w
    if dec.basis == "none":
        return (Reason("C", "文字コードを決められない",
                       "BOM がなく、UTF-8・UTF-16・多バイト（Shift_JIS・EUC-JP・Big5・GBK・EUC-KR）・1バイトの文字コードのどれとしても"
                       "きれいに読めないので、バイトを1字ずつ Latin-1 の字にして読みました。窓の字は正しいとは限らないので、"
                       "エージェントが確かめます（正しい文字コードが分かれば encoding で指定します）"),
                Reason("B", "文字コードを決められない", "書く値の文字コードは、エージェントが正しい文字コードを決めてから書きます"))
    if dec.candidates:
        names = [name] + [c["label"] for c in dec.candidates]
        codes = [dec.encoding] + [c["encoding"] for c in dec.candidates]
        return (Reason("B", "文字コードを規則では決められない",
                       f"{len(names)} つの文字コード（{'・'.join(names)}）のどれとしても、仕様どおりにきれいに読めます"
                       "（BOM・ファイルの中の宣言もありません）。"
                       f"窓は {names[0]} で読み、ほかの読み方の見本（字の違う最初の行）を窓に並べました。"
                       "どれが正しいかは、エージェントが字を読んで決め、その名前を encoding で渡します"
                       f"（{'・'.join(codes)}）", codes),
                Reason("B", "文字コードを規則では決められない",
                       "書く値の文字コードは、エージェントが字を読んで正しい文字コードを決め、encoding で指定してから書きます", codes))
    return None, None


def _grade_data(k: str, blob: bytes, name: str, encoding: str | None = None):
    from .grade import Grade, Reason, _final, _locked

    direct = [Reason("A", "データそのもの", "行と列（か、名前と値）のデータに、そのまま書けます")]
    if k in ("csv", "tsv"):
        from .datafile.csvfile import CsvFile

        f = CsvFile(blob, name, k, encoding)
        nul = f.text.count("\x00")
        if f.dec.binary:
            # NUL があり、文字のファイルである手がかり（BOM・UTF-16/32 の形・UTF-8 の多バイトの並び）がない
            return _locked("読めない", "文字のファイルではありません（どの文字コードとしても読めず、NUL のバイトがあります）",
                           "読めないので、書けません")
        reasons = [Reason("S", "データそのもの", f"表のデータ（{KIND_JA[k]}・{f.dec.label}・{len(f.records):,} 行）で、変換しなくても機械が読めます")]
        if nul:
            # 文字コードとして正しく読めたファイルの中の NUL（U+0000）：値の中の制御文字。中身は読めるので、段階は変えない
            rows = [str(i) for i, ln in enumerate(f.text.split("\n"), 1) if "\x00" in ln][:12]
            reasons.append(Reason("S", "値の中の NUL", f"値の中に NUL（U+0000、制御文字）が {nul:,} 個あります（{f.dec.label}として読めるファイル）。"
                                  "表のデータとしては読めます。NUL を値の一部とするかは、使う側で決めます", [f"{r}行目" for r in rows]))
        if f.problems:
            reasons.append(Reason("S", "書き方の問題", "RFC 4180 に合わない書き方があります（どう読んだかは窓に書いてあります）", f.problems[:5]))
        enc, wenc = encoding_reason(f.dec)
        if enc is not None:
            return _final(reasons + [enc], [wenc] + direct, base="S")
        return Grade("S", reasons, write="A", write_reasons=direct)
    from .datafile.jsonfile import JsonFile

    f = JsonFile(blob, name, k, encoding)
    reasons = [Reason("S", "データそのもの", f"構造のあるデータ（{KIND_JA[k]}・{f.dec.label}）で、変換しなくても機械が読めます")]
    if f.duplicates:
        reasons.append(Reason("C", "同じ名前が2つ以上ある", "オブジェクトに同じ名前が2つ以上あり、どの値を使うかは読むソフトによって違います。"
                              "JSON Pointer では1つに決まらないので、エージェントが読んで判断します", f.duplicates[:12]))
        return Grade("C", reasons, write="B", write_reasons=[Reason("B", "同じ名前が2つ以上ある",
                                                                  "どの値に書くかを、エージェントが読んで決めます", f.duplicates[:12])])
    enc, wenc = encoding_reason(f.dec)
    if enc is not None:
        return _final(reasons + [enc], [wenc] + direct, base="S")
    return Grade("S", reasons, write="A", write_reasons=direct)


def _grade_txt(blob: bytes, name: str, encoding: str | None = None):
    """文字だけのファイル：RTF と同じ決め方（作り手の印の欄はないので、欄があれば推定の B・規則で決められない所は書く B）。
    文字コードは encoding_reason のとおり（候補がいくつもあれば読む B、どれでも読めなければ読む C）。"""
    from .grade import Reason, _final, _locked
    from .txt.document import TxtDocument, TxtError
    from .txt.window import build

    try:
        doc = TxtDocument(blob, name, encoding)
    except TxtError as e:
        return _locked("読めない", str(e), "読めないので、書けません")
    g = _grade_textdoc(build(doc), "テキスト", "入力欄の印")
    enc, wenc = encoding_reason(doc.dec, "A")
    if enc is None:
        return g
    out = _final(g.reasons + [enc], g.write_reasons + [wenc])
    out.alerts = g.alerts
    return out


def _grade_textdoc(win: dict, kind_ja: str, explicit: str, extra: list | None = None):
    """Word と同じ決め方：入力欄（作り手の印・配置からの推定・規則では決められない所）は書く段階で決め、読む段階は
    中身が読めるか（マクロ・読めない部分）だけで決める。"""
    from .grade import Reason, _final, _macro_reasons, _write_fields, _write_macros

    reasons: list = []
    agent = [f for f in win["fields"] if f.get("agent")]
    fields = [f for f in win["fields"] if not f.get("agent")]
    layout = [f for f in fields if f.get("source") == "layout"]
    reasons += _macro_reasons(win.get("macros"))
    for e in win.get("unread") or []:
        reasons.append(Reason("F", "読めない部分", e))
    reasons += extra or []
    write = _write_fields(fields, layout, agent, f"{kind_ja}の{explicit}", "欄のない所は、前後の字を指定して書き換えます")
    write += _write_macros(win.get("macros"))
    from .grade import picture_alerts, word_reasons

    wr, ww = word_reasons(win)  # 変更履歴で削除された字（読む B）・書けない入力欄（書く D）
    reasons += wr
    write += ww

    g = _final(reasons, write)
    g.alerts = picture_alerts(win.get("pictures") or [])
    return g
