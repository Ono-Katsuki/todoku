"""記入例を「Excelが出した答え」として使い、記入と再計算の正しさを確かめる（todoku example と同じ）。

使い方: uv run python scripts/check_with_example.py 空の様式.xlsx 記入例.xlsx
"""

from __future__ import annotations

import sys

from todoku.cli import main

if __name__ == "__main__":
    sys.exit(main(["example", *sys.argv[1:]]))
