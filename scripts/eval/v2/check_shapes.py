"""出力が、動画の地図とポータルの読む形と同じかを確かめる。本物のファイルは読むだけ。
使い方: check_shapes.py <publish_candidate のフォルダ> <world/portal/data.json> <film/src/v3/maps.py>"""

import ast
import json
import os
import sys
from pathlib import Path


def same_keys(real, mock, path=""):
    """本物の形（辞書の鍵と、値の型）に合うか。国の鍵のように数が変わる所は呼ぶ側で切る。"""
    errs = []
    if isinstance(real, dict):
        if not isinstance(mock, dict):
            return [f"{path}：辞書でない"]
        for k in real:
            if k not in mock:
                errs.append(f"{path}.{k}：ない")
            else:
                errs += same_keys(real[k], mock[k], f"{path}.{k}")
    elif isinstance(real, list):
        if not isinstance(mock, list) or len(real) != len(mock):
            errs.append(f"{path}：並びが違う")
    elif isinstance(real, (int, float)) and not isinstance(mock, (int, float)):
        errs.append(f"{path}：数でない（{mock!r}）")
    return errs


def film_check(maps_py: str, film_json: str) -> dict:
    """maps.py の world_counts() をそのまま取り出して動かす（numpy などを読まない）。"""
    src = Path(maps_py).read_text(encoding="utf-8")
    fn = next(n for n in ast.parse(src).body if isinstance(n, ast.FunctionDef) and n.name == "world_counts")
    g = {"json": json, "WORLD_COUNTS": film_json}
    exec(compile(ast.Module([fn], []), maps_py, "exec"), g)  # noqa: S102
    return g["world_counts"]()


def main():
    pub, data_json, maps_py = sys.argv[1:4]
    errs = []
    real = json.loads(Path(data_json).read_text(encoding="utf-8"))
    frag = json.loads((Path(pub) / "portal_fragment.json").read_text(encoding="utf-8"))
    errs += same_keys(real["final"], frag["final"], "final")
    one_real = next(c["counts"] for c in real["countries"].values() if c.get("counts"))
    if not frag["countries"]:
        errs.append("countries が空")
    for cc, c in frag["countries"].items():
        errs += same_keys(one_real, c["counts"], f"countries.{cc}.counts")
    if os.path.exists(maps_py):
        w = film_check(maps_py, str(Path(pub) / "film_world_counts.json"))
        print("動画の形：status", w["status"], "countries_surveyed", w["countries_surveyed"], "total", w["total"], "国", len(w["countries"]))
    else:
        print("maps.py がないので動画の形は確かめない", maps_py)
    if errs:
        print("形が違う：\n  " + "\n  ".join(errs[:30]), file=sys.stderr)
        return 1
    print("形：ポータルの final と countries[cc].counts に合う")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
