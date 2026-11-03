"""第6回の見直しで見つかった、証明・読み戻しを通ってしまう誤った書き出しと、誤った断りの確かめ（合成の文書）。

- Excel：数式セルの計算結果を差し替えるとき、属性のある <v xml:space="preserve"> も取り除き、<v> を一つにする。
  <v>・<f>・<is> が二つ以上あるセル（CT_Cell に合わない）は、証明で断り、読み戻しでも合うとしない
"""

from __future__ import annotations

import io
import re
import zipfile

from todoku import api
from test_round5_office import _xlsx


# ---------------------------------------------------------------- Excel：<v> の二重化
def _vcount(blob: bytes, ref: str) -> int:
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        s = z.read("xl/worksheets/sheet1.xml").decode()
    m = re.search(r'<c r="' + ref + r'"[^>]*>(.*?)</c>', s, re.S)
    return len(re.findall(r"<v[ >]", m[1]))


def test_formula_value_with_space_attr_is_replaced_not_duplicated(tmp_path):
    rows = ('<row r="1"><c r="A1"/><c r="B1" t="str"><f>"区 "&amp;A1&amp;" 号 "</f>'
            '<v xml:space="preserve">区  号 </v></c></row>')
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx(rows))
    out = tmp_path / "o.xlsx"
    r = api.fill(str(src), {"S!A1": "業平"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    blob = out.read_bytes()
    assert _vcount(blob, "B1") == 1
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        s = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<v xml:space="preserve">区 業平 号 </v>' in s


def test_proof_rejects_cell_with_two_values():
    from todoku.xlsx.proof import prove

    a = _xlsx('<row r="1"><c r="A1"><v>1</v></c></row>')
    b = _xlsx('<row r="1"><c r="A1"><v>2</v><v>1</v></c></row>')
    pr = prove(a, b, allowed_inputs={("S", "A1")})
    assert not pr.ok
    assert any("CT_Cell" in u for u in pr.unexpected)


def test_readback_rejects_cell_with_two_values():
    from todoku.xlsx.readback import check
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(_xlsx('<row r="1"><c r="A1"><v>2</v><v>1</v></c></row>'))
    assert wb.sheet("S").cells[(1, 1)].schema_error
    assert check(wb, {"S!A1": 2})["S!A1"][0] is False
