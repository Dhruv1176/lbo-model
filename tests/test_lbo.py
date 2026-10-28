import numpy as np
import pytest
from openpyxl import load_workbook

from lbo import analysis as an
from lbo import build_excel as bx
from lbo.engine import N, Params, run
from lbo.recalc import recalc, soffice_path
from lbo.reconcile import reconcile

P = Params()


def test_sources_equal_uses_and_equity_plug():
    r = run(P)
    assert np.isclose(r["uses"], r["ev"] + r["txn_fees"] + r["fin_fees"] + P.min_cash)
    assert np.isclose(r["equity"] + r["tlb0"] + r["notes0"] + P.rev_drawn, r["uses"])


@pytest.mark.parametrize("scn", [1, 2, 3])
def test_balance_sheet_balances_every_year(scn):
    r = run(P.with_(scenario=scn))
    fees_net = r["fin_fees"] - r["fee_amort"] * np.arange(N + 1)
    assets = r["cash"] + r["nwc"] + r["ppe"] + r["goodwill"] + fees_net
    assert np.allclose(assets, r["debt"] + r["equity"] + r["ret_earn"], atol=1e-9)


def test_cash_never_below_minimum_and_revolver_within_commitment_under_stress():
    stress = P.with_(scenario=3, tlb_x=5.0, notes_x=2.0, entry_mult=11.0)
    r = run(stress, margin=[0.10] * N)         # brutal margin collapse forces revolver use
    assert r["rev_bal"].max() > 0               # the revolver actually draws
    assert r["rev_bal"].max() <= P.rev_commit + 1e-9
    assert (r["cash"][1:] >= P.min_cash - 1e-9).all() or r["rev_bal"].max() >= P.rev_commit - 1e-9


def test_sweep_and_amortisation_logic():
    r = run(P)
    assert (np.diff(r["tlb"]) <= 1e-9).all()                       # TLB only goes down
    assert np.allclose(r["notes"], r["notes"][0])                  # bullet
    assert np.isclose(r["mand"][1], 0.01 * r["tlb0"])              # 1% mandatory
    zero_sweep = run(P.with_(sweep_pct=0.0))
    assert zero_sweep["cash"][N] > r["cash"][N]                    # no sweep -> cash builds instead
    assert zero_sweep["tlb"][N] > r["tlb"][N]


def test_returns_math():
    r = run(P)
    X = P.exit_year
    eq_exit = P.exit_mult * r["ebitda"][X] - r["net_debt"][X]
    proceeds = eq_exit - P.mip_pct * max(0, eq_exit - r["equity"])
    assert np.isclose(r["sel_proc"] if "sel_proc" in r else r["sel_proceeds"], proceeds)
    assert np.isclose(r["sel_irr"], (proceeds / r["equity"]) ** (1 / X) - 1)
    assert np.isclose(r["bridge"]["total"], proceeds - r["equity"])  # value bridge reconciles


def test_interest_cap_raises_taxes():
    assert run(P)["taxes"][1:].sum() >= run(P.with_(int_cap_on=0))["taxes"][1:].sum()


def test_scenario_ordering_and_monotonicity():
    irr = [run(P.with_(scenario=s))["sel_irr"] for s in (3, 1, 2)]
    assert irr[0] < irr[1] < irr[2]
    assert run(P.with_(entry_mult=10))["sel_irr"] > run(P)["sel_irr"] > run(P.with_(entry_mult=12))["sel_irr"]
    assert run(P.with_(exit_mult=12))["sel_irr"] > run(P)["sel_irr"]


def test_analysis_outputs():
    be = an.breakevens(P)
    assert 9 < be["max_entry_mult_for_20pct_irr"] < P.entry_mult          # must pay less than 11x to earn 20%
    assert np.isclose(an.irr_of(P.with_(entry_mult=be["max_entry_mult_for_20pct_irr"])), 0.20, atol=1e-6)
    assert be["exit_mult_for_20pct_irr"] > P.exit_mult
    mc = an.monte_carlo(P, 400)
    assert 0 <= mc["p_loss"] <= 1 and mc["p5"] < mc["median_irr"] < mc["p95"]
    tor = an.tornado(P)
    assert tor["swing"].is_monotonic_decreasing


@pytest.mark.skipif(soffice_path() is None, reason="LibreOffice not installed")
@pytest.mark.parametrize("changes", [{}, dict(scenario=3, exit_year=4, entry_mult=12.0, tlb_x=4.5, sweep_pct=1.0),
                                     dict(scenario=2, exit_year=3, exit_mult=13.0, int_cap_on=0)])
def test_excel_matches_engine_after_recalc(tmp_path, changes):
    """The audit: change inputs INSIDE the workbook, let LibreOffice recalculate, compare every key line to the engine."""
    path = str(tmp_path / "m.xlsx")
    bx.build(Params(), path)
    names = {"int_cap_on": "int_cap_toggle"}
    if changes:
        wb = load_workbook(path)
        for k, v in changes.items():
            wb["Assumptions"][f"C{bx.R['Assumptions'][names.get(k, k)]}"].value = v
        wb.save(path)
    assert recalc(path)
    res = reconcile(path, Params(**changes))
    assert res["master_check"] == "ALL OK"
    assert max(res["max_abs_diff"].values()) < 1e-6
