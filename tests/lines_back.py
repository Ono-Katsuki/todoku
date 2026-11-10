"""list_fields の結果の context_same_as を、道具の説明に書いた決まりだけで context に戻す（テスト用。とどくの実装は使わない）。"""

from __future__ import annotations

import re


def restore(fields: list[dict]) -> list[dict]:
    """行を持つ欄は、同じ結果の、前の欄。▶ のない行はそのまま。▶ のある行は、▶ を外して、その欄の番地の所（［番地）に ▶ を付ける。"""
    lines: dict[str, str] = {}
    out = []
    for f in fields:
        if "context_same_as" in f:
            assert "context" not in f
            head = lines[f["context_same_as"]]
            line = head.replace("▶", "")
            if "▶" in head:
                at = [m.start() for m in re.finditer(re.escape("［" + f["id"].rsplit("!", 1)[-1]) + r"(?![0-9A-Za-z])", line)]
                assert len(at) == 1, (f["id"], head)
                line = line[:at[0] + 1] + "▶" + line[at[0] + 1:]
            f = {("context" if k == "context_same_as" else k): (line if k == "context_same_as" else v) for k, v in f.items()}
        elif isinstance(f.get("context"), str):
            lines[f["id"]] = f["context"]
        out.append(f)
    return out
