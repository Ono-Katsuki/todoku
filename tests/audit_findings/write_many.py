"""いくつかの欄にまとめて書き、結果（say_to_user の一部）と、読み戻しの窓の行を出す。

使い方: PYTHONPATH=src python write_many.py FORM_FILE OUT_FILE 'JSON の値'
"""

from __future__ import annotations

import json
import sys

from todoku import api


def main(src, out, values):
    vals = json.loads(open(values[1:], encoding="utf-8").read() if values.startswith("@") else values)
    r = api.fill(src, vals, out)
    print(json.dumps({k: r[k] for k in r if k not in ("say_to_user",)}, ensure_ascii=False, default=str)[:1500])
    k, doc, parts = api.window(out)
    for _t, md in parts:
        for ln in md.splitlines():
            for v in vals.values():
                v = v["value"] if isinstance(v, dict) else v
                if isinstance(v, str) and v[:2] in ln:
                    print("  BACK:", ln)
                    break


if __name__ == "__main__":
    main(*sys.argv[1:4])
