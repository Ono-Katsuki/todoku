"""表示形式の経過時間（[h]:mm など）：24 時間で折り返さず、窓の字にする（元の値のまま出していた）。"""

from __future__ import annotations

import pytest

from todoku.formula.numfmt import format_value


@pytest.mark.parametrize(("fmt", "value", "want"), [
    ("[h]:mm", 1.5, "36:00"),
    ("[h]:mm", 0.0, "0:00"),
    ("[h]:mm:ss", 0.0416666666666667, "1:00:00"),
    ("[hh]:mm", 0.0416666666666667, "01:00"),
    ("[m]:ss", 0.5, "720:00"),
    ("[s]", 0.001, "86"),
    (r"\ [h]:\ mm;\ @", 0.291666666666667, " 7: 00"),
])
def test_elapsed(fmt, value, want):
    assert format_value(value, fmt) == want


def test_clock_hours_still_wrap():
    assert format_value(1.5, "h:mm") == "12:00"
