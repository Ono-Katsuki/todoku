"""PDF に書くとき、字ごとにどの書体（Helv・TdkJ・TdkK・TdkG・TdkC）で描くかを出す。

使い方: PYTHONPATH=src python pdf_runs.py
"""

from __future__ import annotations

from todoku.pdf.fill import FillError, runs_for

for s in ["Иванов", "Παπαδόπουλος", "Nguyễn Đức", "Đ", "Ð", "ş ğ İ ı", "Łódź", "Ő ű", "€", "髙﨑", "张伟", "①", "ｱ"]:
    try:
        print(f"{s!r:20}", [(k, r) for k, r in runs_for(s)])
    except FillError as e:
        print(f"{s!r:20}", "FillError", e)
