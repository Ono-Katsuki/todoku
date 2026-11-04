#!/bin/zsh
# 比べる道具の venv を作る。プロジェクトの依存には入れない。
# 使い方: setup_baselines.sh <venv のフォルダ>
set -eu
venv="$1"
here="${0:A:h}"
[[ -x "$venv/bin/python" ]] || uv venv --python 3.11 "$venv"
uv pip install --python "$venv/bin/python" -r "$here/baseline_requirements.txt"
"$venv/bin/python" -I -c "import markitdown, openpyxl, xlrd, pdfplumber, pypdf, docx, PIL, reportlab; print('baselines ok')"
