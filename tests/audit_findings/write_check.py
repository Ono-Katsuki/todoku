"""世界の文字を、様式（form_samples.py）の欄に書き、結果（書けたか・読み戻し）を出す。
PDF は書いた後に pdftoppm で画像にする（目で見て、字の形・向き・つながりを確かめる）。

使い方: PYTHONPATH=src python write_check.py FORM_FILE OUTDIR [--fields]
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path

from todoku import api

VALUES = {
    "Arabic name": "محمد عبد الله ١٢٣",
    "Hebrew name": "ישראל כהן",
    "Hindi name": "किशोर क्षत्रिय",
    "Bengali name": "কিশোর বেগম",
    "Tamil name": "கொடு பெயர்",
    "Thai name": "สมชาย ใจดี",
    "Khmer name": "សុខ ចាន់",
    "Myanmar name": "အောင်ဆန်း",
    "Russian name": "Иванов Иван",
    "Greek name": "Παπαδόπουλος",
    "Vietnamese name": "Nguyễn Văn Đức",
    "Chinese name": "张伟",
    "Korean name": "홍길동",
    "Japanese name": "髙橋 葛\U000E0100飾",
    "Georgian name": "გიორგი",
    "Amharic name": "አበበ በቀለ",
}


def field_for(fields: list[dict], label: str) -> str | None:
    for f in fields:
        blob = json.dumps(f, ensure_ascii=False)
        if label in blob:
            return f.get("id") or f.get("name")
    return None


def main(argv):
    src, outdir = Path(argv[0]), Path(argv[1])
    outdir.mkdir(parents=True, exist_ok=True)
    fields = api.fields(str(src))
    if "--fields" in argv:
        for f in fields:
            print(json.dumps(f, ensure_ascii=False)[:300])
        return
    if "--all" in argv:  # 全部を1つのファイルにまとめて書く（描けない字があれば、その欄は外す）
        vals = {field_for(fields, lab): v for lab, v in VALUES.items() if field_for(fields, lab)}
        out = outdir / f"{src.stem}_all{src.suffix}"
        r = api.fill(str(src), vals, str(out))
        print("proof_ok", r.get("proof_ok"), "read_back_ok", r.get("read_back_ok"), r.get("read_back_differences"),
              r.get("unexpected"), r.get("warnings"))
        k, doc, parts = api.window(str(out))
        text = "\n".join(md for _t, md in parts)
        for lab, v in VALUES.items():
            print(f"{lab:16} back={'OK' if v in text else 'MISSING'}")
        print(out)
        return
    results = []
    for label, value in VALUES.items():
        fid = field_for(fields, label)
        key = fid if fid else "@1"
        v = value if fid else {"after": label, "value": value}
        out = outdir / f"{src.stem}_{label.split()[0]}{src.suffix}"
        try:
            r = api.fill(str(src), {key: v}, str(out))
            ok = r.get("ok", True)
            errs = r.get("errors") or r.get("problems") or r.get("refused")
            k, doc, parts = api.window(str(out))
            text = "\n".join(md for _t, md in parts)
            back = value in text
            back_nfc = unicodedata.normalize("NFC", value) in unicodedata.normalize("NFC", text)
            results.append((label, key, "written" if ok else "not-ok", back, back_nfc, str(errs)[:160]))
        except Exception as e:  # noqa: BLE001
            results.append((label, key, f"ERROR {type(e).__name__}: {str(e)[:160]}", None, None, ""))
            continue
        if src.suffix == ".pdf" and shutil.which("pdftoppm"):
            subprocess.run(["pdftoppm", "-r", "60", "-png", "-singlefile", str(out), str(out.with_suffix(""))], check=False,
                           capture_output=True)
    for r in results:
        print(" | ".join(str(x) for x in r))


if __name__ == "__main__":
    main(sys.argv[1:])
