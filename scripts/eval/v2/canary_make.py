"""試し（preflight）の小さな書類を作る。比べる道具の venv で動かす（openpyxl・python-docx・reportlab）。公的な書類も個人の情報も使わない。
使い方: canary_make.py <出力のフォルダ>"""

import os
import sys

import docx
import openpyxl
from openpyxl.styles import Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

d = sys.argv[1]
os.makedirs(d, exist_ok=True)

# CSV：引用符・区切り・UTF-8 の字
with open(os.path.join(d, "canary.csv"), "w", encoding="utf-8", newline="") as f:
    f.write('name;city;note\r\n"Müller, A";Zürich;"He said ""ok"""\r\nBeta;Genève;plain\r\n')

# Excel：字・数の書式・罫線のある空のマス（書く所）
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Form"
ws["A1"] = "Application form"
ws["A2"] = "Name"
ws["A3"] = "Amount"
ws["B3"] = 1234  # 書式なし：比べる道具が元の値を出しても同じ字（比べる道具が動くことを確かめる）
thin = Side(style="thin")
for c in ("A2", "B2", "A3", "B3"):
    ws[c].border = Border(left=thin, right=thin, top=thin, bottom=thin)
wb.save(os.path.join(d, "canary.xlsx"))

# Word：段落と表（空のマスに書く）
w = docx.Document()
w.add_paragraph("Application form for the canary check")
t = w.add_table(rows=2, cols=2)
t.cell(0, 0).text = "Name"
t.cell(1, 0).text = "Address"
w.save(os.path.join(d, "canary.docx"))

# PDF：字と、見出しの右の空の罫線（書く所）
c = canvas.Canvas(os.path.join(d, "canary.pdf"), pagesize=A4)
c.setFont("Helvetica", 14)
c.drawString(72, 780, "Application form for the canary check")
c.setFont("Helvetica", 11)
for i, label in enumerate(("Name", "Address", "Telephone")):
    y = 720 - i * 40
    c.drawString(72, y, label + ":")
    c.line(160, y - 2, 500, y - 2)
c.save()
