"""世界の文字の監査（2026-11-03）の再現：外部の道具（LibreOffice・Chrome）なしで動く分だけを、まとめて確かめる。

各行の「期待」は、正しい（論理順の、元の）字。いまの todoku で違うものが、見つかった問題。
使い方: PYTHONPATH=src python tests/audit_findings/run_world_scripts.py OUTDIR
（LibreOffice・Chrome・Quartz で作る見本は scripts_samples.py・cjk_samples.py・form_samples.py・quartz_pdf.swift）
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import handmade_pdf  # noqa: E402

from todoku import api  # noqa: E402

EXPECT = {
    "arabic_pf.pdf": ["محمد", "الاسم"],  # 表示形（U+FExx）は基の字に、並びは論理順に
    "deva_visual.pdf": ["किशोर", "कि"],  # 前に出る母音 ि は子音の後ろ（क्षि も）
    "pua.pdf": ["山太郎"],  # 字は正しいが、外字（私用領域）として知らせるべき
    "ivs.pdf": ["葛\U000E0100飾区"],
    "notu_identity.pdf": [],  # 読めない字として F になるのが正しい
    "notu_aj1.pdf": ["亜一丁"],
    "vertical.pdf": [("order", "申請者", "（申）")],  # 右の列（申請者）から読む
    "nfd_marks.pdf": [],
    "thai_marks.pdf": ["ชื่อ"],
    "narrow_double.pdf": ["Castillo Williams"],  # l・i の重なりを太字の重ね書きと取り違えない
}


def body(parts) -> str:
    md = "\n".join(m for _t, m in parts)
    return md.split("### 読み順（上から）", 1)[-1]


def main(outdir: Path) -> int:
    outdir.mkdir(parents=True, exist_ok=True)
    bad = 0
    for name, blob in handmade_pdf.samples().items():
        p = outdir / name
        p.write_bytes(blob)
        k, doc, parts = api.window(str(p))
        text = body(parts)
        g = api.grade(str(p))
        miss = [e for e in EXPECT.get(name, []) if (e not in text if isinstance(e, str) else not (e[1] in text and e[2] in text and text.index(e[1]) < text.index(e[2])))]
        flag = g.get("grade")
        st = "OK" if not miss else "NG"
        bad += bool(miss)
        lines = [ln for ln in text.splitlines() if ln.startswith("- ")]
        print(f"{st} {name:20} grade={flag} window={lines[:3]} missing={miss}")
    # 書く：PDF に描けない字
    from todoku.pdf.fill import FillError, runs_for

    for s in ["محمد", "किशोर", "สมชาย", "Nguyễn", "Ţară", "Țară", "Παπαδόπουλος", "Łódź"]:
        try:
            runs = runs_for(s)
            print(f"   pdf-write {s!r}: runs={[(k, r) for k, r in runs]}（Helv 以外は埋め込まない CJK の書体で描く）")
        except FillError as e:
            print(f"   pdf-write {s!r}: FillError {e}")
    return bad


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(tempfile.mkdtemp())
    sys.exit(1 if main(out) else 0)
