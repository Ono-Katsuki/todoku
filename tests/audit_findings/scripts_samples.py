"""世界の文字の監査に使う、見本の文（論理順の正しい字）と、それを入れた平らな ODT（.fodt）を作る。

使い方: python scripts_samples.py OUTDIR   → OUTDIR/scripts.fodt（LibreOffice で PDF・docx にする）
"""

from __future__ import annotations

import sys
import unicodedata
from pathlib import Path
from xml.sax.saxutils import escape

# 名前 → (文, 書く向き)。文は NFC。数字の混じる右から左の文も入れる
SAMPLES = {
    "arabic": ("الاسم الكامل: محمد عبد الله", "rl"),
    "arabic_num": ("رقم الهاتف 0123 456 والمبلغ 1500 ريال", "rl"),
    "arabic_lig": ("لا إله الله", "rl"),  # لا の合字
    "persian": ("نام و نام خانوادگی", "rl"),
    "urdu": ("درخواست فارم", "rl"),
    "hebrew": ("שם מלא: ישראל כהן 2024", "rl"),
    "hindi": ("आवेदक का नाम किशोर क्षत्रिय श्री", "lr"),  # ि（前に出る母音）、क्ष（結合子音）
    "bengali": ("আবেদনকারীর নাম কিশোর", "lr"),  # ি・ী、কে（前後に分かれる母音）
    "bengali_split": ("কোথায় যেতে চাও", "lr"),  # ো＝ে＋া（前後に分かれる）
    "tamil": ("விண்ணப்பதாரர் பெயர் கொடு", "lr"),  # ொ（前後に分かれる）
    "thai": ("ชื่อผู้ยื่นคำขอ", "lr"),
    "lao": ("ຊື່ຜູ້ຍື່ນຄຳຮ້ອງ", "lr"),
    "khmer": ("ឈ្មោះអ្នកដាក់ពាក្យ", "lr"),
    "myanmar": ("လျှောက်ထားသူအမည်", "lr"),
    "ethiopic": ("የአመልካች ስም", "lr"),
    "georgian": ("განმცხადებლის სახელი", "lr"),
    "armenian": ("Դիմողի անունը", "lr"),
    "cyrillic": ("Фамилия, имя, отчество заявителя", "lr"),
    "greek": ("Ονοματεπώνυμο αιτούντος", "lr"),
    "vietnamese": ("Họ và tên người nộp đơn: Nguyễn Văn Đức", "lr"),
    "korean": ("신청인 성명 홍길동", "lr"),
    "chinese": ("申请人姓名 张伟", "lr"),
    "japanese": ("申請者の氏名 山田太郎 ｶﾅ ＡＢＣ", "lr"),
}


FONTS = {"arabic": "Geeza Pro", "arabic_num": "Geeza Pro", "arabic_lig": "Geeza Pro", "persian": "Geeza Pro",
         "urdu": "Geeza Pro", "hebrew": "Arial Unicode MS", "hindi": "Kohinoor Devanagari", "bengali": "Kohinoor Bangla",
         "bengali_split": "Kohinoor Bangla", "tamil": "Tamil Sangam MN", "thai": "Thonburi", "lao": "Lao Sangam MN",
         "khmer": "Khmer Sangam MN", "myanmar": "Myanmar Sangam MN", "ethiopic": "Kefa III", "korean": "Apple SD Gothic Neo",
         "chinese": "Songti SC", "japanese": "BIZ UDMincho"}


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def fodt() -> str:
    paras = []
    spans = "".join(f'<style:style style:name="F{k}" style:family="text"><style:text-properties '
                    f'fo:font-family="&apos;{f}&apos;" style:font-family-asian="&apos;{f}&apos;" '
                    f'style:font-family-complex="&apos;{f}&apos;"/></style:style>' for k, f in FONTS.items())
    for k, (t, d) in SAMPLES.items():
        style = "R" if d == "rl" else "L"
        f = FONTS.get(k)
        span = f'<text:span text:style-name="F{k}">{escape(t)}</text:span>' if f else escape(t)
        paras.append(f'<text:p text:style-name="{style}">{escape(k)} | {span}</text:p>')
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
<office:automatic-styles>
 <style:style style:name="L" style:family="paragraph"><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
 <style:style style:name="R" style:family="paragraph"><style:paragraph-properties style:writing-mode="rl-tb" fo:text-align="end"/></style:style>
{spans}
</office:automatic-styles>
<office:body><office:text>
{chr(10).join(paras)}
</office:text></office:body></office:document>
"""


def html() -> str:
    rows = []
    for k, (t, d) in SAMPLES.items():
        f = FONTS.get(k, "Times New Roman")
        rows.append(f'<p dir="{"rtl" if d == "rl" else "ltr"}"><span dir="ltr">{escape(k)} |</span> '
                    f'<span style="font-family:\'{f}\'">{escape(t)}</span></p>')
    return ('<!doctype html><html><head><meta charset="utf-8"><style>p{margin:4px;font-size:14px}</style></head><body>'
            + "\n".join(rows) + "</body></html>")


if __name__ == "__main__":
    out = Path(sys.argv[1])
    out.mkdir(parents=True, exist_ok=True)
    (out / "scripts.fodt").write_text(fodt(), encoding="utf-8")
    (out / "scripts.html").write_text(html(), encoding="utf-8")
    (out / "scripts.tsv").write_text("".join(f"{k}\t{FONTS.get(k, 'Times New Roman')}\t{t}\n" for k, (t, _d) in SAMPLES.items()),
                                     encoding="utf-8")
    print(out / "scripts.fodt", out / "scripts.html")
