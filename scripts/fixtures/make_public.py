"""公開してよいテスト用の書類（tests/fixtures/public/）を作り直す。

ここにあるのは、すべてこのスクリプトが作った架空の様式です（個人情報・実在の様式の写しはありません。書体のデータも埋め込みません）。
テストは、ここで作ってコミットしたファイルを使います（テストの側では作り直さないので、作る道具がなくてもテストは動きます）。

使い方:  uv run python scripts/fixtures/make_public.py
  doc・odt・rtf・xls・ods は LibreOffice（soffice）で変換して作ります。変換の出力は毎回バイトが変わるので、
  すでにあれば残し、--reconvert のときだけ作り直します（docx・xlsx を変えたときは --reconvert をつける）。

MANIFEST.json に、各ファイルの形式・含む仕掛け・窓に必ず出る字（作った側が入れた字）・書き込みの試し（欄の名前と値）を書きます。
これは作った書類の中身から決めたもので、とどくの出力に合わせて決めたものではありません。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import make_office as mo  # noqa: E402
import make_pdfs as mp  # noqa: E402

OUT = HERE.parents[1] / "tests" / "fixtures" / "public"
LIMIT = 100 * 1024


def manifest() -> list[dict]:
    """各ファイルの、作った側が知っている正解。

    texts：窓（Markdown）に、そのままの字で出るはずの字。order：その順に出るはずの字（読み順）。
    fields：書類の中にある欄の名前（PDF の欄の名前など）。
    fill：{欄の名前: 値}。欄の名前は書類の中の名前（PDF の欄の名前・FORMTEXT の名前・セルの番地など）。
    fill_by_label：{見出し: 値}。書類に欄の名前がない形式で、見出しの字が同じ欄（規則で決めた欄）に書く。
    needs：{欄の名前: 文字の種類}。その文字を組み立てられる書体がなければ、書き込みの試しを飛ばす。
    encoding：文字だけのファイルの文字コード（作った側が知っている正解）。
    members・member_fill：ZIP の中のファイルと、その中の欄への書き込み。
    """
    pdf_form = ["補助金交付申請書", "氏名", "住所", "生年月日", "電話番号", "区分", "上記の内容に相違ありません。"]
    return [
        {"file": "acroform.pdf", "kind": "pdf", "features": ["AcroForm", "text", "combo", "checkbox", "UniJIS-UTF16-H"],
         "texts": pdf_form, "fields": ["name", "address", "birth", "tel", "kind", "agree"],
         "fill": {"name": "Taro Todoku", "kind": "B", "agree": True}},
        {"file": "flat.pdf", "kind": "pdf", "features": ["no form fields", "ruled table", "2 pages"],
         "texts": pdf_form + ["事業計画書", "事業の名称", "経費の総額", "（注）金額は円単位で記入してください。"],
         "pages": 2},
        {"file": "rotated.pdf", "kind": "pdf", "features": ["/Rotate 90", "AcroForm"],
         "texts": ["受付票", "受付番号", "担当者名", "Reference No."], "order": ["受付票", "受付番号", "担当者名", "Reference No."],
         "fields": ["receipt_no", "staff"], "fill": {"receipt_no": "A-0123"}},
        {"file": "vertical.pdf", "kind": "pdf", "features": ["vertical writing", "UniJIS-UTF16-V"],
         "texts": ["届出書", "私は次のとおり届け出ます。", "氏名", "本籍"],
         "order": ["届出書", "私は次のとおり届け出ます。", "氏名", "本籍"]},
        {"file": "rtl.pdf", "kind": "pdf", "features": ["Hebrew", "Arabic", "lam-alef ligature", "visual order", "ToUnicode"],
         "texts": ["טופס בקשה", "שם מלא", "الاسم"], "fields": ["full_name", "ism"],
         "fill": {"full_name": "דנה לוי"}, "needs": {"full_name": "Hebrew"}},
        {"file": "indic.pdf", "kind": "pdf", "features": ["Devanagari", "conjuncts", "pre-base matra", "ToUnicode"],
         "texts": ["आवेदन पत्र", "नाम", "पता", "हिन्दी"], "fields": ["naam", "pata"],
         "fill": {"naam": "सीता"}, "needs": {"naam": "Devanagari"}},
        {"file": "scanned.pdf", "kind": "pdf", "features": ["image only (scan)", "no text"], "texts": [], "image_only": True},
        {"file": "encrypted.pdf", "kind": "pdf", "features": ["RC4 128-bit (R3)", "user password"], "password": "todoku",
         "texts": pdf_form, "fields": ["name", "address", "birth", "tel", "kind", "agree"], "fill": {"name": "Taro Todoku"}},
        {"file": "form.docx", "kind": "docx",
         "features": ["content controls (text, drop-down, checkbox)", "legacy FORMTEXT/FORMCHECKBOX", "text box", "tracked changes", "table"],
         "texts": mo.DOCX_TEXTS + ["内容に同意します", "結果を郵送で受け取る"], "fill": {"Tel": "03-0000-0000", "氏名": "とどく 花子", "区分": "変更", "同意": True}},
        {"file": "form.doc", "kind": "doc", "features": ["Word 97-2003 (LibreOffice)", "legacy form fields", "table"],
         "texts": ["届出書", "氏名", "区分", "電話番号", "結果を郵送で受け取る"], "converted_from": "form.docx",
         "fill_by_label": {"電話番号": "03-0000-0000"}},
        {"file": "form.odt", "kind": "odt", "features": ["ODF text (LibreOffice)", "table"],
         "texts": ["届出書", "氏名", "区分", "電話番号"], "converted_from": "form.docx",
         "fill": {"Tel": "03-0000-0000", "氏名": "とどく 花子"}},
        {"file": "form.rtf", "kind": "rtf", "features": ["RTF (LibreOffice)", "table", "tracked changes"],
         "texts": ["届出書", "氏名", "区分", "電話番号"], "converted_from": "form.docx", "fill": {"Tel": "03-0000-0000"}},
        {"file": "sheet.xlsx", "kind": "xlsx",
         "features": ["merged cells", "list validation", "formulas with cached values", "rich text", "1904 date system", "sheet protection"],
         "texts": ["利用申請書", "氏名", "区分", "利用開始日", "2026/4/1", "太枠の中を記入してください", "金額", "合計", str(mo.XLSX_TOTAL)],
         "fill": {"申請書!B3": "とどく 太郎"}, "options": {"申請書!B4": ["個人", "法人"]}},
        {"file": "sheet.xls", "kind": "xls", "features": ["BIFF8 (LibreOffice)", "merged cells", "validation", "formulas", "1904"],
         "texts": ["利用申請書", "氏名", "利用開始日", "2026/4/1", "合計", str(mo.XLSX_TOTAL)], "converted_from": "sheet.xlsx",
         "fill": {"申請書!B3": "とどく 太郎"}},
        {"file": "sheet.ods", "kind": "ods", "features": ["ODF spreadsheet (LibreOffice)", "merged cells", "formulas"],
         "texts": ["利用申請書", "氏名", "利用開始日", "合計", str(mo.XLSX_TOTAL)], "converted_from": "sheet.xlsx",
         "fill": {"申請書!B3": "とどく 太郎"}},
        {"file": "slides.pptx", "kind": "pptx", "features": ["title", "table", "notes"],
         "texts": ["説明会の参加申込", "氏名", "所属", "参加日"], "fill_by_label": {"氏名": "とどく 太郎"}},
        {"file": "form.hwpx", "kind": "hwpx", "features": ["OWPML", "table", "click-here fields", "form controls", "tracked deletion"],
         "texts": ["보조금 지급 신청서", "성명", "주소", "전화번호"], "fill": {"전화번호": "010-0000-0000"}},
        {"file": "form.jtd", "kind": "jtd", "features": ["Ichitaro (synthetic CFB)", "table"],
         "texts": ["申請書", "氏名", "住所"], "fill_by_label": {"氏名": "とどく 太郎"}},
        {"file": "notes.txt", "kind": "txt", "features": ["UTF-8 text"], "texts": ["申請書の書き方", "氏名を記入する"]},
        {"file": "items.csv", "kind": "csv", "features": ["UTF-8 CSV", "quoted comma"], "texts": ["品目", "封筒", "窓あき, 長形3号"],
         "fill": {"R2.C4": "至急"}},
        {"file": "items_sjis.csv", "kind": "csv", "features": ["Shift_JIS (cp932) CSV"], "texts": ["品目", "封筒", "窓あき, 長形3号"],
         "encoding": "cp932", "fill": {"R2.C4": "至急"}},
        {"file": "form.json", "kind": "json", "features": ["JSON"], "texts": ["参加申込", "電話番号"],
         "fill": {"/項目/0/値": "とどく 太郎"}},
        {"file": "bundle.zip", "kind": "zip", "features": ["ZIP of several forms"],
         "members": ["申請書/acroform.pdf", "申請書/form.docx", "申請書/items.csv"],
         "member_fill": {"申請書/form.docx": {"Tel": "03-0000-0000"}}},
    ]


def build(reconvert: bool = False) -> dict[str, bytes]:
    out: dict[str, bytes] = {name: fn() for name, fn in mp.ALL.items()}
    out["form.docx"] = mo.docx()
    out["sheet.xlsx"] = mo.xlsx()
    out["slides.pptx"] = mo.pptx()
    out["form.hwpx"] = mo.hwpx()
    out["form.jtd"] = mo.jtd()
    out.update(mo.texts())
    for name, src, ext, flt in (("form.doc", "form.docx", "doc", "MS Word 97"), ("form.odt", "form.docx", "odt", None),
                                ("form.rtf", "form.docx", "rtf", None), ("sheet.xls", "sheet.xlsx", "xls", "MS Excel 97"),
                                ("sheet.ods", "sheet.xlsx", "ods", None)):
        if (OUT / name).exists() and not reconvert:
            continue  # LibreOffice の書き出しは毎回バイトが変わるので、--reconvert のときだけ作り直す
        blob = mo.soffice_convert(src, out[src], ext, flt)
        if blob is not None:
            out[name] = blob
        else:
            print(f"{name}: soffice がないので作れません", file=sys.stderr)
    out["bundle.zip"] = mo.bundle({"申請書/" + m: out[m] for m in ("acroform.pdf", "form.docx", "items.csv")})
    return out


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    files = build(reconvert="--reconvert" in sys.argv)
    for name, blob in files.items():
        assert len(blob) < LIMIT, f"{name} は {len(blob)} バイト（{LIMIT} 未満にする）"
        (OUT / name).write_bytes(blob)
    entries = manifest()
    missing = [e["file"] for e in entries if not (OUT / e["file"]).exists()]
    assert not missing, missing
    (OUT / "MANIFEST.json").write_text(json.dumps({
        "about": "とどくのテスト用の架空の書類（scripts/fixtures/make_public.py が作った）。個人情報なし。書体のデータは埋め込んでいない。",
        "license": "CC0-1.0",
        "files": entries}, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    for e in entries:
        print(f"{e['file']:18} {(OUT / e['file']).stat().st_size:7,d}  {', '.join(e['features'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
