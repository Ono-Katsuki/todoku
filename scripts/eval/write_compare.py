#!/usr/bin/env python3
"""書く比べ：同じ欄に同じ値を、比べる道具（openpyxl・python-docx・pypdf・pypdf＋reportlab の重ね描き）ととどくで書き、とどくを使わずに結果を比べる。
- 欄はとどくが選ぶ（比べる道具にも同じ欄を渡す。欄を見つける力は問わないので、比べる道具に有利）
- 値は日本語（様式の実際）と英数字（Helvetica の注釈でも書ける形：比べる道具に有利）の両方
- 確かめること：値が入ったか（PDF は Poppler で描いて、欄の四角の中の画素が変わったか）・何が黙って消えたか・元のバイト列・電子署名
- 署名は、ここで作った署名つきの PDF で必ず 1 件確かめる（公開の書類にはほとんどない）

書けた（write_ok。world_counts.py の「同じ定義」もこれ）：値が入り、何も消えず、署名があれば署名が有効のまま。

使い方: write_compare.py --files 一覧 --out 出力 --baseline-python PY --todoku-python PY [--label dev] [--no-signed] [--purge]
途中で止めても、もう一度動かせば続きから。"""

from __future__ import annotations

import argparse
import collections
import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import make_signed  # noqa: E402
from common import ROOT, atomic_json, doc_id, eprint, guard_path, kind_of, last_json, load_json, md_table, norm, read_list, run_watched  # noqa: E402

TOOLS = {"xlsx": ("openpyxl",), "docx": ("python-docx",), "pdf": ("pypdf", "pypdf_overlay"), "xls": (), "doc": ()}
VALUES = {"ja": "山田花子", "ascii": "TEST-12345"}


def one(path: str, a, base: Path) -> dict:
    guard_path(path)
    k = kind_of(path)
    did = doc_id(path)
    w = base / "work" / did
    w.mkdir(parents=True, exist_ok=True)
    src = w / ("doc" + Path(path).suffix.lower())
    if not src.exists():
        shutil.copy(path, src)
    tw = [a.todoku_python, "-I", HERE / "todoku_write.py", ROOT / "src", src]
    pick = last_json(run_watched([*tw, "pick"], limit_s=300))
    rec = {"id": did, "path": path, "kind": k, "target": pick, "runs": {}}
    if not pick.get("id"):
        rec["skip"] = pick.get("error") or "書ける欄がない"
        return rec
    tgt = json.dumps(pick, ensure_ascii=False)
    sig_before = make_signed.verify(src, w) if k == "pdf" else None
    for vname, val in VALUES.items():
        if pick.get("kind") == "number":
            val = "12345"
        for tool in ("todoku", *TOOLS.get(k, ())):
            out = w / f"{vname}.{tool}{src.suffix}"
            out.unlink(missing_ok=True)
            if tool == "todoku":
                res = last_json(run_watched([*tw, "fill", out, pick["id"], val], limit_s=300))
            else:
                res = last_json(run_watched([a.baseline_python, "-I", HERE / "baseline_write.py", "write", tool, src, out, k, tgt, val], limit_s=300))
            r = {"tool": tool, "value": vname, **res}
            if out.exists() and not res.get("error"):
                r["check"] = last_json(run_watched([a.baseline_python, "-I", HERE / "baseline_write.py", "check", src, out, k, tgt, val], limit_s=300))
                c = r["check"]
                if k == "pdf" and sig_before and sig_before.get("valid"):
                    r["signature_after"] = make_signed.verify(out, w)
                    if not r["signature_after"].get("valid"):
                        c.setdefault("lost", {})["signature"] = "無効になった"
                if k == "pdf" and a.pkr and c.get("visible") and not c.get("xfa"):
                    # 描いた絵を文字認識で読む：FreeText の枠だけが描かれて字がない、字化けする、を見分ける（Vision はとどくと関係がない）
                    png = w / "crop.png"
                    run_watched([a.baseline_python, "-I", HERE / "baseline_write.py", "crop", out, pick.get("page", 1), json.dumps(pick.get("rect")), png], limit_s=180)
                    got = run_watched([a.pkr, "ocr", png], limit_s=120)["stdout"]
                    try:
                        text = "".join(json.loads(got.splitlines()[-1]).get("text", []))
                    except (IndexError, json.JSONDecodeError):
                        text = None
                    png.unlink(missing_ok=True)
                    if text is not None:
                        c["ocr_text"] = text[:80]
                        c["landed"] = norm(val) in norm(text)
                        c["checked_by"] = "ocr"
                r["write_ok"] = bool(c.get("landed")) and not c.get("lost") and not c.get("error")
            else:
                r["write_ok"] = False
            rec["runs"][f"{tool}:{vname}"] = r
            out.unlink(missing_ok=True)
    rec["signed_before"] = sig_before
    if a.purge:
        shutil.rmtree(w, ignore_errors=True)
    return rec


def form_class(rec: dict) -> str:
    if rec["kind"] != "pdf":
        return rec["kind"]
    if (rec.get("signed_before") or {}).get("valid"):
        return "pdf（署名あり）"
    if any((r.get("check") or {}).get("xfa") for r in rec["runs"].values()):
        return "pdf（XFA）"
    return "pdf（フォームの欄）" if (rec.get("target") or {}).get("name") else "pdf（欄のない PDF）"


def summarize(recs: list[dict]) -> dict:
    out = {"docs": len(recs), "skipped": sum(1 for r in recs if r.get("skip")), "groups": {}}
    groups = collections.defaultdict(list)
    for r in recs:
        if not r.get("skip") and not r.get("failed"):
            groups[form_class(r)].append(r)
    for g, rs in sorted(groups.items()):
        tools = sorted({key for r in rs for key in r["runs"]}, key=lambda x: (not x.startswith("todoku"), x))
        d = {"docs": len(rs), "tools": {}}
        for key in tools:
            runs = [r["runs"][key] for r in rs if key in r["runs"]]
            checks = [x["check"] for x in runs if x.get("check")]
            lost = collections.Counter()
            for c in checks:
                for k, v in (c.get("lost") or {}).items():
                    lost[k] += v if isinstance(v, int) else 1
            d["tools"][key] = {"tried": len(runs), "wrote": len(checks), "landed": sum(1 for c in checks if c.get("landed")),
                               "no_loss": sum(1 for c in checks if not c.get("lost") and not c.get("error")), "write_ok": sum(1 for x in runs if x.get("write_ok")),
                               "parts_same": sum(c.get("parts_same") or 0 for c in checks), "parts_total": sum(c.get("parts_total") or 0 for c in checks),
                               "bytes_kept": sum(1 for c in checks if c.get("original_bytes_kept")), "proof_ok": sum(1 for x in runs if x.get("proof_ok")),
                               "read_back_ok": sum(1 for x in runs if x.get("read_back_ok")), "lost": dict(lost.most_common(14)),
                               "signature_valid": sum(1 for x in runs if (x.get("signature_after") or {}).get("valid")),
                               "signature_runs": sum(1 for x in runs if x.get("signature_after")),
                               "errors": collections.Counter((x.get("error") or "").split(":")[0] for x in runs if x.get("error")).most_common(4)}
        out["groups"][g] = d
    return out


def markdown(s: dict, label: str) -> str:
    L = [f"# 書く比べ（{label}）", "", f"書類 {s['docs']} 件（書ける欄がなく除いた {s['skipped']} 件）。「道具:ja」は日本語の値、「道具:ascii」は英数字の値。", ""]
    for g, d in s["groups"].items():
        tools = list(d["tools"])
        rows = [["書けた"] + [f"{d['tools'][t]['wrote']}/{d['tools'][t]['tried']}" for t in tools],
                ["値が入った（PDF は描いて確かめた）"] + [d["tools"][t]["landed"] for t in tools],
                ["何も消えなかった"] + [d["tools"][t]["no_loss"] for t in tools],
                ["書けた＋何も消えない（write_ok）"] + [d["tools"][t]["write_ok"] for t in tools]]
        if any(d["tools"][t]["parts_total"] for t in tools):
            rows.append(["元とバイト一致の部品"] + [f"{d['tools'][t]['parts_same']}/{d['tools'][t]['parts_total']}" for t in tools])
        if g.startswith("pdf"):
            rows.append(["元のバイト列をそのまま残す"] + [d["tools"][t]["bytes_kept"] for t in tools])
        if any(d["tools"][t]["signature_runs"] for t in tools):
            rows.append(["署名が有効のまま"] + [f"{d['tools'][t]['signature_valid']}/{d['tools'][t]['signature_runs']}" for t in tools])
        rows.append(["証明・読み戻しが通った（とどくだけ）"] + [f"{d['tools'][t]['proof_ok']}・{d['tools'][t]['read_back_ok']}" if t.startswith("todoku") else "—" for t in tools])
        L += [f"## {g}（{d['docs']} 件）", "", md_table(["", *tools], rows), ""]
        for t in tools:
            if d["tools"][t]["lost"]:
                L.append(f"- {t} で消えたもの：" + "・".join(f"{k} {v}" for k, v in d["tools"][t]["lost"].items()))
            if d["tools"][t]["errors"]:
                L.append(f"- {t} の失敗：" + "・".join(f"{k} {v}" for k, v in d["tools"][t]["errors"]))
        L.append("")
    return "\n".join(L)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--files", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--baseline-python", required=True)
    ap.add_argument("--todoku-python", required=True)
    ap.add_argument("--label", default="")
    ap.add_argument("--no-signed", action="store_true")
    ap.add_argument("--pkr", default=None, help="Vision の文字認識（pkr）。あれば、PDF の「値が入った」を描いた絵を読んで確かめる。なければ画素の変化だけ")
    ap.add_argument("--purge", action="store_true")
    a = ap.parse_args()
    base = Path(a.out)
    (base / "scores").mkdir(parents=True, exist_ok=True)
    files = [p for p in read_list(a.files) if kind_of(p) in TOOLS]
    if not a.no_signed:
        files.append(str(make_signed.make(base / "signed")) if not (base / "signed" / "signed.pdf").exists() else str(base / "signed" / "signed.pdf"))
    for i, p in enumerate(files, 1):
        sp = base / "scores" / f"{doc_id(p)}.json"
        if sp.exists():
            continue
        try:
            rec = one(p, a, base)
        except Exception as e:  # noqa: BLE001
            eprint(f"失敗 {p}: {type(e).__name__}: {e}")
            rec = {"id": doc_id(p), "path": p, "kind": kind_of(p), "failed": f"{type(e).__name__}: {str(e)[:300]}", "runs": {}}
        atomic_json(sp, rec)
        print(f"[{i}/{len(files)}] {Path(p).name} " + " ".join(f"{k}:{'OK' if r.get('write_ok') else 'NG'}" for k, r in rec["runs"].items()) + (f" 除いた：{rec['skip'][:40]}" if rec.get("skip") else ""), flush=True)
    recs = [r for r in (load_json(base / "scores" / f"{doc_id(p)}.json") for p in files) if r]
    s = summarize(recs)
    atomic_json(base / "summary.json", s)
    (base / "summary.md").write_text(markdown(s, a.label), encoding="utf-8")
    print(markdown(s, a.label))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
