"""LBO engine. Mirrors the Excel workbook cell-for-cell (tests reconcile the two).

Conventions
- $ in millions; Year 0 = close / LTM, Years 1..5 = projection.
- Cash-free / debt-free purchase. Interest is charged on BEGINNING balances (no circularity).
- Cash sweep: a % of excess cash (after mandatory amortisation and minimum cash) prepays the Term Loan B.
- Revolver draws automatically to cover any shortfall below minimum cash, up to the commitment.
- Tax: US-style 163(j) cap on net interest deductibility (toggle); no NOL carry-forward (conservative).
- Management incentive plan (MIP) = % of the equity gain at exit.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace

import numpy as np

N = 5


@dataclass
class Params:
    company: str = "Project Atlas (fictional industrial-services business)"
    scenario: int = 1                       # 1 Base, 2 Upside, 3 Downside
    # LTM financials and opening balance sheet
    ltm_rev: float = 500.0
    ltm_ebitda: float = 100.0
    ppe_open: float = 200.0
    # transaction
    entry_mult: float = 11.0
    txn_fee_pct: float = 0.02
    fin_fee_pct: float = 0.025
    fee_amort_yrs: int = 7
    min_cash: float = 20.0
    # exit
    exit_year: int = 5
    exit_mult: float = 11.0
    mip_pct: float = 0.08
    # financing
    rev_commit: float = 75.0
    rev_drawn: float = 0.0
    rev_spread: float = 0.0350
    rev_undrawn: float = 0.0050
    tlb_x: float = 4.0
    tlb_spread: float = 0.0375
    tlb_amort: float = 0.01
    sweep_pct: float = 0.75
    notes_x: float = 1.5
    notes_rate: float = 0.0925
    cash_rate: float = 0.02
    sofr: tuple = (0.0400, 0.0375, 0.0350, 0.0350, 0.0350)
    cov_max: tuple = (7.00, 6.75, 6.50, 6.25, 6.00)   # max net debt / EBITDA covenant
    # operating
    tax: float = 0.25
    da_pct: float = 0.035
    capex_pct: float = 0.040
    nwc_pct: float = 0.12
    int_cap_on: int = 1
    int_cap_pct: float = 0.30
    growth: dict = field(default_factory=lambda: {
        1: (0.07, 0.07, 0.06, 0.06, 0.05),
        2: (0.09, 0.09, 0.08, 0.08, 0.07),
        3: (0.02, -0.04, 0.01, 0.03, 0.03)})
    margin: dict = field(default_factory=lambda: {
        1: (0.205, 0.210, 0.215, 0.220, 0.220),
        2: (0.210, 0.220, 0.230, 0.235, 0.240),
        3: (0.195, 0.175, 0.180, 0.185, 0.190)})

    def with_(self, **kw) -> "Params":
        return replace(self, **kw)


SCENARIOS = {1: "Base", 2: "Upside", 3: "Downside"}


def run(p: Params, growth=None, margin=None, sofr=None) -> dict:
    """Run the model. Optional growth / margin / sofr override the selected scenario (used by Monte Carlo)."""
    g = np.asarray(growth if growth is not None else p.growth[p.scenario], float)
    m = np.asarray(margin if margin is not None else p.margin[p.scenario], float)
    sofr = np.asarray(sofr if sofr is not None else p.sofr, float)

    z = lambda: np.zeros(N + 1)
    rev, ebitda, da, capex, nwc = z(), z(), z(), z(), z()
    rev[0], ebitda[0] = p.ltm_rev, p.ltm_ebitda
    nwc[0] = p.nwc_pct * p.ltm_rev
    for t in range(1, N + 1):
        rev[t] = rev[t - 1] * (1 + g[t - 1])
        ebitda[t] = rev[t] * m[t - 1]
        da[t], capex[t], nwc[t] = rev[t] * p.da_pct, rev[t] * p.capex_pct, rev[t] * p.nwc_pct

    # transaction
    tlb0, notes0 = p.tlb_x * p.ltm_ebitda, p.notes_x * p.ltm_ebitda
    ev = p.entry_mult * p.ltm_ebitda
    txn_fees = p.txn_fee_pct * ev
    fin_fees = (tlb0 + notes0 + p.rev_commit) * p.fin_fee_pct
    uses = ev + txn_fees + fin_fees + p.min_cash
    equity = uses - p.rev_drawn - tlb0 - notes0
    fee_amort = fin_fees / p.fee_amort_yrs
    goodwill = ev + txn_fees - nwc[0] - p.ppe_open

    keys = ("cash rev_bal tlb notes cash_int int_inc ebt net_int disallowed taxes ni fcf mand sweep rev_flow "
            "cash_avail ppe ret_earn unlev_fcf").split()
    o = {k: z() for k in keys}
    o["cash"][0], o["rev_bal"][0], o["tlb"][0], o["notes"][0] = p.min_cash, p.rev_drawn, tlb0, notes0
    o["ppe"][0] = p.ppe_open
    for t in range(1, N + 1):
        cash_b, rev_b, tlb_b, notes_b = o["cash"][t - 1], o["rev_bal"][t - 1], o["tlb"][t - 1], o["notes"][t - 1]
        cash_int = (rev_b * (sofr[t - 1] + p.rev_spread) + (p.rev_commit - rev_b) * p.rev_undrawn
                    + tlb_b * (sofr[t - 1] + p.tlb_spread) + notes_b * p.notes_rate)
        int_inc = cash_b * p.cash_rate
        ebit = ebitda[t] - da[t]
        ebt = ebit - cash_int + int_inc - fee_amort
        net_int = cash_int + fee_amort - int_inc
        allowed = min(net_int, p.int_cap_pct * ebitda[t]) if p.int_cap_on else net_int
        disallowed = net_int - allowed
        taxes = max(0.0, ebt + disallowed) * p.tax
        ni = ebt - taxes
        fcf = ni + da[t] + fee_amort - (nwc[t] - nwc[t - 1]) - capex[t]
        mand = min(tlb_b, p.tlb_amort * tlb0)
        avail = cash_b + fcf - mand - p.min_cash
        flow = min(-avail, p.rev_commit - rev_b) if avail < 0 else -min(rev_b, avail)
        sweep = min(tlb_b - mand, p.sweep_pct * max(0.0, avail + flow))
        o["cash_int"][t], o["int_inc"][t], o["ebt"][t], o["net_int"][t] = cash_int, int_inc, ebt, net_int
        o["disallowed"][t], o["taxes"][t], o["ni"][t], o["fcf"][t] = disallowed, taxes, ni, fcf
        o["mand"][t], o["sweep"][t], o["rev_flow"][t], o["cash_avail"][t] = mand, sweep, flow, avail
        o["cash"][t] = cash_b + fcf - mand + flow - sweep
        o["rev_bal"][t] = rev_b + flow
        o["tlb"][t] = tlb_b - mand - sweep
        o["notes"][t] = notes_b
        o["ppe"][t] = o["ppe"][t - 1] + capex[t] - da[t]
        o["ret_earn"][t] = o["ret_earn"][t - 1] + ni
        o["unlev_fcf"][t] = ebitda[t] - (ebitda[t] - da[t]) * p.tax - capex[t] - (nwc[t] - nwc[t - 1])

    debt = o["rev_bal"] + o["tlb"] + o["notes"]
    net_debt = debt - o["cash"]
    out = dict(rev=rev, ebitda=ebitda, da=da, capex=capex, nwc=nwc, debt=debt, net_debt=net_debt, **o)
    out.update(ev=ev, txn_fees=txn_fees, fin_fees=fin_fees, uses=uses, equity=equity, goodwill=goodwill,
               tlb0=tlb0, notes0=notes0, fee_amort=fee_amort)

    # returns by exit year (index 1..N)
    ev_x = p.exit_mult * ebitda
    eq_x = ev_x - net_debt
    gain = eq_x - equity
    mip = p.mip_pct * np.maximum(0, gain)
    proceeds = eq_x - mip
    n = np.arange(N + 1)
    with np.errstate(divide="ignore", invalid="ignore"):
        moic = proceeds / equity
        irr = np.where(proceeds > 0, np.power(np.maximum(moic, 1e-12), 1 / np.maximum(n, 1)) - 1, -1.0)
    out.update(ev_x=ev_x, eq_x=eq_x, gain=gain, mip=mip, proceeds=proceeds, moic=moic, irr=irr)

    # credit statistics (years 1..N)
    with np.errstate(divide="ignore", invalid="ignore"):
        out["tot_lev"] = debt / ebitda
        out["net_lev"] = net_debt / ebitda
        out["sr_lev"] = (o["rev_bal"] + o["tlb"]) / ebitda
        out["int_cov"] = ebitda / o["cash_int"]
        out["fccr"] = (ebitda - capex - o["taxes"]) / (o["cash_int"] + o["mand"])
    cov = np.r_[np.nan, np.asarray(p.cov_max, float)]
    out["cov_max"] = cov
    out["breach"] = np.where(np.isnan(cov), False, out["net_lev"] > cov)

    # exit-year summary and value-creation bridge
    X = p.exit_year
    out["X"] = X
    out["sel_irr"], out["sel_moic"], out["sel_proceeds"] = irr[X], moic[X], proceeds[X]
    m0, m1 = p.ltm_ebitda / p.ltm_rev, ebitda[X] / rev[X]
    nd0 = debt[0] - o["cash"][0]
    out["bridge"] = dict(
        revenue_growth=(rev[X] - rev[0]) * m0 * p.entry_mult,
        margin_expansion=(m1 - m0) * rev[X] * p.entry_mult,
        multiple_change=(p.exit_mult - p.entry_mult) * ebitda[X],
        deleveraging=nd0 - net_debt[X],
        fees=-(txn_fees + fin_fees),
        mip=-mip[X])
    out["bridge"]["total"] = sum(out["bridge"].values())
    return out


def irr_table(p: Params, **axes):
    """Convenience: IRR for a list of Params variants."""
    return [run(v)["sel_irr"] for v in axes.get("variants", [])]
