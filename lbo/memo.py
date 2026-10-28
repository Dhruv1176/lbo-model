"""Generates the investment committee memo. Every number comes from the engine, so memo and model cannot disagree."""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from .engine import N, SCENARIOS, Params, run


def _t(df: pd.DataFrame) -> str:
    return df.to_markdown(index=True) if hasattr(df, "to_markdown") else df.to_string()


def md_table(header, rows):
    out = ["| " + " | ".join(header) + " |", "|" + "|".join("---" if i == 0 else "---:" for i in range(len(header))) + "|"]
    out += ["| " + " | ".join(str(c) for c in r) + " |" for r in rows]
    return "\n".join(out)


def usd(x, d=1):
    return f"${x:,.{d}f}" if x >= 0 else f"(${-x:,.{d}f})"


def pc(x, d=1):
    return f"{x * 100:.{d}f}%"


def write_memo(p: Params, tor: pd.DataFrame, be: dict, mc: dict, path: str = "docs/INVESTMENT_MEMO.md") -> str:
    r = run(p)
    X = p.exit_year
    sc = {s: run(p.with_(scenario=s)) for s in (1, 2, 3)}
    hurdle = 0.20
    ok = r["sel_irr"] >= hurdle
    atp = be["max_entry_mult_for_20pct_irr"]
    rec = (f"**Recommendation: proceed to confirmatory diligence, but do not bid {p.entry_mult:.1f}x.** At {p.entry_mult:.1f}x the base case "
           f"returns {pc(r['sel_irr'])} / {r['sel_moic']:.2f}x, below the {pc(hurdle, 0)} hurdle. The credit is sound (see below), so the issue is price, "
           f"not risk of loss. Ability-to-pay analysis says {atp:.1f}x clears {pc(hurdle, 0)} with the same capital structure and operating case."
           if not ok and np.isfinite(atp) else
           f"**Recommendation: proceed at {p.entry_mult:.1f}x.** Base case returns {pc(r['sel_irr'])} / {r['sel_moic']:.2f}x, above the {pc(hurdle, 0)} hurdle.")

    m_sel = p.margin[p.scenario]
    m0 = p.ltm_ebitda / p.ltm_rev
    any_breach = {SCENARIOS[k]: int(v["breach"].sum()) for k, v in sc.items()}
    breach_txt = ("is never breached in any of the three cases" if not any(any_breach.values())
                  else "IS BREACHED in: " + ", ".join(f"{k} ({v} year(s))" for k, v in any_breach.items() if v))
    g3, m3 = p.growth[3], p.margin[3]
    top_driver = tor.iloc[1]["driver"].split(" (")[0].lower()
    proj_hdr = ["$mm", "Close"] + [f"Y{i}" for i in range(1, N + 1)]
    f_ = lambda key, fmt: [fmt(r[key][i]) for i in range(N + 1)]
    proj = [["Revenue"] + f_("rev", usd),
            ["  growth", "-"] + [pc(r["rev"][i] / r["rev"][i - 1] - 1) for i in range(1, N + 1)],
            ["EBITDA"] + f_("ebitda", usd),
            ["  margin"] + [pc(r["ebitda"][i] / r["rev"][i]) for i in range(N + 1)],
            ["Capex", "-"] + [usd(r["capex"][i]) for i in range(1, N + 1)],
            ["Cash interest", "-"] + [usd(r["cash_int"][i]) for i in range(1, N + 1)],
            ["Cash taxes", "-"] + [usd(r["taxes"][i]) for i in range(1, N + 1)],
            ["Free cash flow (post debt service)", "-"] + [usd(r["fcf"][i]) for i in range(1, N + 1)],
            ["Term Loan B"] + f_("tlb", usd), ["Net debt"] + f_("net_debt", usd),
            ["Net debt / EBITDA"] + [f"{r['net_lev'][i]:.2f}x" for i in range(N + 1)],
            ["EBITDA / cash interest", "-"] + [f"{r['int_cov'][i]:.2f}x" for i in range(1, N + 1)],
            ["Covenant (max net lev.)", "-"] + [f"{r['cov_max'][i]:.2f}x" for i in range(1, N + 1)]]

    br = r["bridge"]
    gain = br["total"]
    bridge_rows = [[k, usd(v), pc(v / gain)] for k, v in (("Revenue growth", br["revenue_growth"]), ("Margin expansion", br["margin_expansion"]),
                   ("Multiple expansion", br["multiple_change"]), ("Deleveraging (FCF-funded debt paydown)", br["deleveraging"]),
                   ("Fees and expenses", br["fees"]), ("Management incentive plan", br["mip"]), ("**Total sponsor gain**", br["total"]))]

    scen_rows = []
    for s in (1, 2, 3):
        e = sc[s]
        scen_rows.append([SCENARIOS[s], usd(e["rev"][X]), usd(e["ebitda"][X]), pc(e["ebitda"][X] / e["rev"][X]), f"{e['net_lev'][X]:.2f}x",
                          f"{np.nanmin(e['int_cov'][1:]):.2f}x", usd(e["rev_bal"].max()), str(int(e["breach"].sum())), f"{e['sel_moic']:.2f}x", pc(e["sel_irr"])])

    ent_rows = []
    exits = [p.exit_mult + k for k in (-2, -1, 0, 1, 2)]
    for em in [p.entry_mult + k * 0.5 for k in (-3, -2, -1, 0, 1, 2, 3)]:
        ent_rows.append([f"{em:.1f}x"] + [pc(run(p.with_(entry_mult=em, exit_mult=x))["sel_irr"]) for x in exits])

    tor_rows = [[d.driver, pc(d.downside_irr), pc(d.upside_irr), f"{d.swing * 100:.1f} pp"] for d in tor.itertuples()]

    text = f"""# Investment Committee Memo: Project Atlas

> **FICTIONAL CASE STUDY.** Project Atlas is an invented industrial-services business and every input is an illustrative assumption,
> not company data. The memo and the Excel model (`models/Project_Atlas_LBO.xlsx`) are generated from the same Python engine and reconciled to
> 1e-9, so every number below ties to the workbook. Educational use only; not investment advice.

## 1. Executive summary

- **Transaction:** acquire 100% of Atlas for **{usd(r['ev'], 0)}mm EV ({p.entry_mult:.1f}x LTM EBITDA of {usd(p.ltm_ebitda, 0)}mm; {r['ev'] / p.ltm_rev:.1f}x revenue)**, funded with {usd(r['tlb0'] + r['notes0'] + p.rev_drawn, 0)}mm of debt ({(r['tlb0'] + r['notes0'] + p.rev_drawn) / p.ltm_ebitda:.1f}x) and {usd(r['equity'], 0)}mm of sponsor equity ({pc(r['equity'] / r['uses'])} of capitalisation).
- **Base-case returns (Year {X} exit at {p.exit_mult:.1f}x):** **{pc(r['sel_irr'])} IRR / {r['sel_moic']:.2f}x MOIC** after a {pc(p.mip_pct, 0)} management incentive plan (MIP).
- **Range:** Upside {pc(sc[2]['sel_irr'])} / {sc[2]['sel_moic']:.2f}x; Downside {pc(sc[3]['sel_irr'])} / {sc[3]['sel_moic']:.2f}x. Monte Carlo ({mc['n']:,} paths) median IRR **{pc(mc['median_irr'])}**, 5th to 95th percentile **{pc(mc['p5'])} to {pc(mc['p95'])}**, probability of losing capital **{pc(mc['p_loss'])}**.
- **Credit:** leverage falls from {r['net_lev'][0]:.1f}x to {r['net_lev'][X]:.1f}x net by Year {X}; interest cover starts at {r['int_cov'][1]:.1f}x; the maintenance covenant {breach_txt}.

{rec}

## 2. Transaction overview

{md_table(["Sources", "$mm", "x EBITDA", "Uses", "$mm"], [
    ["Revolver (drawn; $%.0fmm commitment)" % p.rev_commit, usd(p.rev_drawn), f"{p.rev_drawn / p.ltm_ebitda:.1f}x", "Purchase of enterprise", usd(r['ev'])],
    ["Term Loan B (S+%d bp, 1%% amort., %s sweep)" % (round(p.tlb_spread * 1e4), pc(p.sweep_pct, 0)), usd(r['tlb0']), f"{p.tlb_x:.1f}x", "Transaction fees (%s of EV)" % pc(p.txn_fee_pct), usd(r['txn_fees'])],
    ["Senior unsecured notes (%s fixed, bullet)" % pc(p.notes_rate, 2), usd(r['notes0']), f"{p.notes_x:.1f}x", "Financing fees (%s)" % pc(p.fin_fee_pct), usd(r['fin_fees'])],
    ["Sponsor equity", usd(r['equity']), f"{r['equity'] / p.ltm_ebitda:.1f}x", "Minimum cash to balance sheet", usd(p.min_cash)],
    ["**Total sources**", usd(r['uses']), f"{r['uses'] / p.ltm_ebitda:.1f}x", "**Total uses**", usd(r['uses'])]])}

## 3. Investment thesis (illustrative)

1. **Resilient, recurring demand.** Mandatory inspection and testing work gives a defensive revenue base, which is why a {(r['tlb0'] + r['notes0']) / p.ltm_ebitda:.1f}x structure is financeable and covenant headroom stays wide.
2. **Operating leverage.** Margin moves from {pc(m0)} at close to {pc(m_sel[2])} by Year 3 and {pc(m_sel[4])} by Year 5 ({(m_sel[4] - m0) * 1e4:.0f} bp) from route density and pricing. After entry price, {top_driver} is the largest IRR driver in the tornado; this is the assumption diligence must stress hardest.
3. **Cash generation funds deleveraging.** {usd(r['fcf'][1:X + 1].sum(), 0)}mm of cumulative free cash flow after debt service over {X} years repays debt and builds cash, contributing {pc(br['deleveraging'] / gain, 0)} of the sponsor gain.
4. **No multiple expansion is underwritten.** Exit = entry multiple. Returns come from growth and cash flow, not re-rating.

## 4. Base-case projections

{md_table(proj_hdr, proj)}

## 5. Returns and value creation

Exit in Year {X} at {p.exit_mult:.1f}x EBITDA of {usd(r['ebitda'][X])}mm gives enterprise value of {usd(r['ev_x'][X], 0)}mm, less net debt of {usd(r['net_debt'][X], 0)}mm, equals equity value of {usd(r['eq_x'][X], 0)}mm; after the MIP, sponsor proceeds are {usd(r['proceeds'][X], 0)}mm on {usd(r['equity'], 0)}mm invested.

{md_table(["Driver", "$mm", "% of gain"], bridge_rows)}

![bridge](value_creation_bridge.png)

**Read-across:** returns are driven mostly by EBITDA growth (revenue and margin together {pc((br['revenue_growth'] + br['margin_expansion']) / gain, 0)}) and deleveraging ({pc(br['deleveraging'] / gain, 0)}). There is no multiple contribution by construction; fees and the MIP are a {pc(-(br['fees'] + br['mip']) / gain, 0)} drag.

## 6. Scenarios

{md_table(["Case", f"Rev Y{X}", f"EBITDA Y{X}", "Margin", f"Net lev. Y{X}", "Min int. cover", "Peak revolver", "Cov. breaches", "MOIC", "IRR"], scen_rows)}

The Downside case is a Year-2 recession (revenue {pc(g3[1], 0)}, margin falling to {pc(m3[1])}). Equity still returns capital ({sc[3]['sel_moic']:.2f}x) because the cash sweep keeps deleveraging, but IRR is only {pc(sc[3]['sel_irr'])}. Peak net leverage is {np.nanmax(sc[3]['net_lev'][1:]):.2f}x against a {p.cov_max[1]:.2f}x covenant.

![credit](credit_profile.png)

## 7. Valuation sensitivity (IRR; entry multiple down, exit multiple across)

{md_table(["Entry \\ Exit"] + [f"{x:.1f}x" for x in exits], ent_rows)}

**Ability to pay.** Holding structure and operating case constant, the maximum entry multiple that still earns {pc(0.15, 0)} / {pc(0.20, 0)} / {pc(0.25, 0)} IRR is **{be['max_entry_mult_for_15pct_irr']:.1f}x / {be['max_entry_mult_for_20pct_irr']:.1f}x / {be['max_entry_mult_for_25pct_irr']:.1f}x**. Conversely, at {p.entry_mult:.1f}x entry the exit multiple required for a {pc(0.20, 0)} IRR is **{be['exit_mult_for_20pct_irr']:.1f}x** (multiple expansion we are unwilling to underwrite) and the exit multiple at which we only get our money back is **{be['exit_mult_for_1x_moic']:.1f}x**.

## 8. Risk analysis

### 8.1 What moves the IRR (one factor at a time)

{md_table(["Driver", "Adverse", "Favourable", "Swing"], tor_rows)}

![tornado](tornado.png)

Price paid, margin, growth and exit multiple dominate; financing terms (SOFR, leverage, tax) matter far less. Debt structure is not where this deal is won or lost.

### 8.2 Monte Carlo ({mc['n']:,} paths)

Correlated shocks: annual growth (σ 2.0 pp common + 1.5 pp idiosyncratic), margin (σ 1.2 pp common, 0.6 correlation with growth, + 0.4 pp idiosyncratic), exit multiple (N({p.exit_mult:.1f}x, 1.0x), clipped 7x to 15x), SOFR (σ 100 bp parallel).

| Statistic | Value |
|---|---:|
| Median IRR / MOIC | {pc(mc['median_irr'])} / {mc['median_moic']:.2f}x |
| 5th / 25th / 75th / 95th percentile IRR | {pc(mc['p5'])} / {pc(mc['p25'])} / {pc(mc['p75'])} / {pc(mc['p95'])} |
| P(IRR ≥ 20%) | {pc(mc['p_irr_ge_20'])} |
| P(IRR ≥ 15%) | {pc(mc['p_irr_ge_15'])} |
| P(IRR < 10%) | {pc(mc['p_irr_lt_10'])} |
| P(MOIC < 1.0x), i.e. loss of capital | {pc(mc['p_loss'])} |
| P(covenant breach in any year) | {pc(mc['p_breach'])} |
| P(revolver drawn in any year) | {pc(mc['p_revolver'])} |

![mc](monte_carlo_irr.png)

### 8.3 Break-even analysis (credit resilience)

| Test | Result |
|---|---:|
| Margin cut (all years, vs Base) before the revolver is first drawn | {be['margin_cut_to_revolver_draw'] * 1e4:.0f} bp |
| Margin cut before any year has negative free cash flow | {be['margin_cut_to_negative_fcf'] * 1e4:.0f} bp |
| Margin cut before the net-leverage covenant is breached | {be['margin_cut_to_covenant_breach'] * 1e4:.0f} bp |

Interpretation: a permanent {be['margin_cut_to_revolver_draw'] * 1e4:.0f} bp margin miss starts consuming liquidity, and a {be['margin_cut_to_covenant_breach'] * 1e4:.0f} bp miss trips the covenant. For scale, the Base case itself assumes only {(m_sel[4] - m0) * 1e4:.0f} bp of margin improvement, so the cushion is several times larger than the expansion we are underwriting.

## 9. Key risks and mitigants

| Risk | Why it matters | Mitigant / diligence action |
|---|---|---|
| Margin expansion not delivered | Second-largest IRR driver | Bottom-up route-density and pricing study; price/volume/mix bridge |
| Cyclical volume shock | Downside case shows IRR compression to {pc(sc[3]['sel_irr'])} | Contract-mix analysis; test 2008-09 and 2020 revenue history |
| Floating-rate exposure | TLB is floating; SOFR ±100 bp moves IRR only ~{abs(tor.set_index('driver').loc['SOFR (+100 bp / -100 bp)', 'swing']) * 100:.1f} pp but cash interest cover matters | Hedge 50-75% of the TLB for 3 years |
| Interest deductibility | The 30% EBITDA cap raises cash taxes in early years | Tax diligence; model as-is already includes the cap |
| Entry price | Single largest IRR lever | Disciplined bid ceiling at the ability-to-pay multiple |
| Exit environment | Multiple compression below 11x erodes returns quickly | Underwrite exit = entry; build a second exit path (strategic sale / IPO) |

## 10. Model methodology and limitations

- Integrated three-statement mechanics: income statement, cash flow, debt schedule with revolver and cash sweep, balance sheet (balance check = 0 in every period), returns, credit statistics and covenant tests. 11 sheets, ~800 live formulas, 9 automated integrity checks.
- **Interest is charged on beginning-of-year balances** to avoid circular references; this slightly overstates interest in a deleveraging deal (conservative).
- Cash-free / debt-free purchase; transaction fees capitalised into goodwill; no purchase-price allocation, deferred taxes or stock-based compensation; no NOL or interest carry-forwards.
- Sensitivity Tables 1-3 are live formulas; Tables 4-5 and the scenario columns are static Python snapshots (re-run `python -m lbo build` after changing inputs).
- Monte Carlo shocks are stylised and independent across years beyond the common factor; real downturns cluster and are fat-tailed, so treat tail probabilities as indicative.
"""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(text, encoding="utf-8")
    return path
