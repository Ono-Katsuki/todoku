"""pyproject.toml の requires-python が、指定した Python の版（例 3.12）を許すかを答える（標準ライブラリだけ）。

CI の行列で、許していない版は「飛ばす」にする（版の制限を緩めたら、そのまま動き出す）。
使い方:  python .github/scripts/python_allowed.py 3.12   → 標準出力に allowed=true / allowed=false
"""

from __future__ import annotations

import operator
import re
import sys
import tomllib
from pathlib import Path

OPS = {">=": operator.ge, "<=": operator.le, ">": operator.gt, "<": operator.lt, "==": operator.eq, "!=": operator.ne}


def version(s: str) -> tuple[int, ...]:
    return tuple(int(x) for x in s.split("."))


def allowed(spec: str, ver: str) -> bool:
    v = version(ver)
    for part in filter(None, (p.strip() for p in spec.split(","))):
        m = re.fullmatch(r"(>=|<=|==|!=|>|<|~=)\s*([0-9.]+)(\.\*)?", part)
        if not m:
            raise SystemExit(f"読めない requires-python: {part!r}")
        op, want, star = m.group(1), version(m.group(2)), m.group(3)
        if op == "~=":
            if not (v >= want and v[:len(want) - 1] == want[:-1]):
                return False
            continue
        n = len(want) if star or op in ("==", "!=") else max(len(want), len(v))
        a, b = (v + (0,) * n)[:n], (want + (0,) * n)[:n]
        if not OPS[op](a, b):
            return False
    return True


def main() -> int:
    spec = tomllib.loads((Path(__file__).resolve().parents[2] / "pyproject.toml").read_text(encoding="utf-8"))["project"]["requires-python"]
    ok = allowed(spec, sys.argv[1])
    print(f"allowed={'true' if ok else 'false'}")
    print(f"requires-python {spec!r} → {sys.argv[1]}: {'許す' if ok else '許さない'}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
