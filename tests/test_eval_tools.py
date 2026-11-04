"""最終評価の道具（scripts/eval）の、数の部品の試験。"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts" / "eval"))
from common import betainc, clopper_pearson, doc_id, guard_path, korn_graubard  # noqa: E402


def test_betainc_known_values():
    assert betainc(1, 1, 0.3) == pytest.approx(0.3)
    assert betainc(2, 2, 0.5) == pytest.approx(0.5)
    assert betainc(5, 2, 0.5) == pytest.approx(7 / 64)


def test_clopper_pearson_known_values():
    # 0/10：上限は 1-(0.025)^(1/10)。10/10：下限は (0.025)^(1/10)
    assert clopper_pearson(0, 10) == pytest.approx([0.0, 1 - 0.025 ** 0.1], abs=1e-4)
    assert clopper_pearson(10, 10) == pytest.approx([0.025 ** 0.1, 1.0], abs=1e-4)
    lo, hi = clopper_pearson(5, 10)
    assert lo == pytest.approx(0.1871, abs=1e-3)
    assert hi == pytest.approx(0.8129, abs=1e-3)


def test_korn_graubard_uses_effective_size():
    # 分散が大きいと有効な大きさが小さくなり、区間は広がる。n を超えない
    wide = korn_graubard(0.5, 0.01, 1000)
    assert wide["n_eff"] == pytest.approx(25)
    narrow = korn_graubard(0.5, 0.0001, 1000)
    assert narrow["n_eff"] == pytest.approx(1000)
    assert wide["ci"][1] - wide["ci"][0] > narrow["ci"][1] - narrow["ci"][0]
    # どの層も 100%：分散 0 でも幅 0 にならない
    allok = korn_graubard(1.0, 0.0, 20)
    assert allok["ci"][0] < 1.0 and allok["ci"][1] == 1.0


def test_guard_refuses_held_out_in_dry_run(monkeypatch):
    monkeypatch.setenv("EVAL_DEV_ONLY", "1")
    with pytest.raises(SystemExit):
        guard_path("/x/forms_tama/a.pdf")
    with pytest.raises(SystemExit):
        guard_path("/x/forms_world/a.pdf")
    guard_path("/x/forms23/files/a.pdf")


def test_doc_id_is_stable_and_distinct():
    assert doc_id("/a/x.pdf") == doc_id("/a/x.pdf")
    assert doc_id("/a/x.pdf") != doc_id("/b/x.pdf")


def test_clean_todoku_restores_document_text():
    from read_compare import clean_todoku

    assert clean_todoku("- 6: 〔P6〕申請日　［p1.g143-24？　　］年［？　］月") == "申請日　　　年　月"
    assert clean_todoku("［L15＝山田］ ｜ 〈AF35＝1,200〉") == "山田   1,200"
    assert clean_todoku("対象者［↓余白 18.1 mm］") == "対象者 "
