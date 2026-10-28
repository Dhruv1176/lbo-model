"""Deal analytics on top of the engine: tornado, Monte Carlo, ability-to-pay, break-evens, charts."""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from .engine import N, SCENARIOS, Params, run


def irr_of(p: Params, **kw) -> float:
    return float(run(p, **kw)["sel_irr"])


def bisect(fn, lo, hi, target, tol=1e-10, it=100):
    """Root of fn(x) = target on [lo, hi] for a monotone fn (either direction). NaN if not bracketed."""
    flo, fhi = fn(lo) - target, fn(hi) - target
    if not (np.isfinite(flo) and np.isfinite(fhi)) or flo * fhi > 0:
        return float("nan")
    for _ in range(it):
        mid = (lo + hi) / 2
        fm = fn(mid) - target
        if abs(fm) < tol:
            return mid
        if flo * fm <= 0:
            hi = mid
        else:
            lo, flo = mid, fm
    return (lo + hi) / 2


def tornado(p: Params) -> pd.DataFrame:
    """One-at-a-time IRR impact of a realistic low/high move in each driver."""
    base = irr_of(p)
    g, m = p.growth[p.scenario], p.margin[p.scenario]
    shifts = [
        ("Exit multiple (±1.0x)", p.with_(exit_mult=p.exit_mult - 1), p.with_(exit_mult=p.exit_mult + 1)),
        ("Entry multiple (+1.0x / -1.0x)", p.with_(entry_mult=p.entry_mult + 1), p.with_(entry_mult=p.entry_mult - 1)),
        ("EBITDA margin (±200 bp)", None, None, dict(margin=[x - .02 for x in m]), dict(margin=[x + .02 for x in m])),
        ("Revenue growth (±2 pp p.a.)", None, None, dict(growth=[x - .02 for x in g]), dict(growth=[x + .02 for x in g])),
        ("Senior leverage (TLB -0.5x / +0.5x)", p.with_(tlb_x=p.tlb_x - .5), p.with_(tlb_x=p.tlb_x + .5)),
        ("SOFR (+100 bp / -100 bp)", None, None, dict(sofr=[x + .01 for x in p.sofr]), dict(sofr=[x - .01 for x in p.sofr])),
        ("Capex (+1 pp / -1 pp of revenue)", p.with_(capex_pct=p.capex_pct + .01), p.with_(capex_pct=p.capex_pct - .01)),
        ("Cash tax rate (+3 pp / -3 pp)", p.with_(tax=p.tax + .03), p.with_(tax=p.tax - .03)),
    ]
    rows = []
    for s in shifts:
        if s[1] is not None:
            lo, hi = irr_of(s[1]), irr_of(s[2])
        else:
            lo, hi = irr_of(p, **s[3]), irr_of(p, **s[4])
        rows.append({"driver": s[0], "downside_irr": lo, "upside_irr": hi, "swing": abs(hi - lo)})
    df = pd.DataFrame(rows).sort_values("swing", ascending=False).reset_index(drop=True)
    df["base_irr"] = base
    return df


def breakevens(p: Params) -> dict:
    out = {}
    X = p.exit_year
    out["exit_mult_for_20pct_irr"] = bisect(lambda x: irr_of(p.with_(exit_mult=x)), 5, 25, 0.20)
    out["exit_mult_for_15pct_irr"] = bisect(lambda x: irr_of(p.with_(exit_mult=x)), 5, 25, 0.15)
    out["exit_mult_for_1x_moic"] = bisect(lambda x: run(p.with_(exit_mult=x))["sel_moic"], 2, 25, 1.0)
    out["max_entry_mult_for_20pct_irr"] = bisect(lambda x: irr_of(p.with_(entry_mult=x)), 7, 20, 0.20)
    out["max_entry_mult_for_25pct_irr"] = bisect(lambda x: irr_of(p.with_(entry_mult=x)), 7, 20, 0.25)
    out["max_entry_mult_for_15pct_irr"] = bisect(lambda x: irr_of(p.with_(entry_mult=x)), 7, 20, 0.15)
    m = np.array(p.margin[p.scenario])
    g = np.array(p.growth[p.scenario])

    def breach(shift):
        r = run(p, margin=m - shift)
        return float(r["breach"].any())

    def revolver(shift):
        return float(run(p, margin=m - shift)["rev_bal"].max() > 1e-6)

    def fcf_neg(shift):
        return float(run(p, margin=m - shift)["fcf"][1:].min() < 0)

    for name, fn in (("margin_cut_to_covenant_breach", breach), ("margin_cut_to_revolver_draw", revolver), ("margin_cut_to_negative_fcf", fcf_neg)):
        lo = None
        for s in np.arange(0, 0.2001, 0.0025):
            if fn(s) >= 1.0:
                lo = s
                break
        out[name] = lo if lo is not None else float("nan")
    # Year-2-style recession: how deep a one-off revenue shock in Year 2 breaches the covenant?
    def shock(drop):
        gg = g.copy(); gg[1] = g[1] - drop
        return run(p, growth=gg, margin=m - 0.02 * (drop > 0))
    out["rev_shock_to_breach"] = next((float(d) for d in np.arange(0, 0.40, 0.005) if shock(d)["breach"].any()), float("nan"))
    return out


def monte_carlo(p: Params, n: int = 5000, seed: int = 11) -> dict:
    """Correlated shocks to growth, margin, exit multiple and rates; reports IRR / credit outcome distribution."""
    rng = np.random.default_rng(seed)
    g, m = np.array(p.growth[p.scenario]), np.array(p.margin[p.scenario])
    cov = np.array([[0.020**2, 0.6 * 0.020 * 0.012], [0.6 * 0.020 * 0.012, 0.012**2]])   # growth, margin (common, corr 0.6)
    common = rng.multivariate_normal([0, 0], cov, size=n)
    irr, moic, breach, rev_draw, minlev_head = (np.empty(n) for _ in range(5))
    for i in range(n):
        gi = g + common[i, 0] + rng.normal(0, 0.015, N)
        mi = m + common[i, 1] + rng.normal(0, 0.004, N)
        xm = float(np.clip(rng.normal(p.exit_mult, 1.0), 7.0, 15.0))
        sf = np.asarray(p.sofr) + rng.normal(0, 0.01)
        r = run(p.with_(exit_mult=xm), growth=gi, margin=mi, sofr=sf)
        irr[i], moic[i] = r["sel_irr"], r["sel_moic"]
        breach[i] = r["breach"].any()
        rev_draw[i] = r["rev_bal"].max() > 1e-6
        minlev_head[i] = np.nanmin(1 - r["net_lev"][1:] / r["cov_max"][1:])
    q = lambda a, x: float(np.percentile(a, x))
    return dict(n=n, irr=irr, moic=moic, breach=breach, rev_draw=rev_draw,
                mean_irr=float(irr.mean()), median_irr=q(irr, 50), p5=q(irr, 5), p25=q(irr, 25), p75=q(irr, 75), p95=q(irr, 95),
                p_irr_ge_20=float((irr >= .20).mean()), p_irr_ge_15=float((irr >= .15).mean()), p_irr_lt_10=float((irr < .10).mean()),
                p_loss=float((moic < 1).mean()), p_breach=float(breach.mean()), p_revolver=float(rev_draw.mean()),
                median_moic=q(moic, 50), p5_moic=q(moic, 5))


def charts(p: Params, tor: pd.DataFrame, mc: dict, out_dir: str = "docs") -> list[str]:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    files = []
    base = run(p)
    plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False, "axes.spines.right": False})

    # 1 tornado
    fig, ax = plt.subplots(figsize=(9, 5))
    t = tor.iloc[::-1]
    b = base["sel_irr"]
    ax.barh(t["driver"], (t["downside_irr"] - b) * 100, left=b * 100, color="#C00000", label="adverse move")
    ax.barh(t["driver"], (t["upside_irr"] - b) * 100, left=b * 100, color="#2E7D32", label="favourable move")
    ax.axvline(b * 100, color="k", lw=1)
    ax.set(xlabel="Sponsor IRR (%)", title=f"IRR tornado, {SCENARIOS[p.scenario]} case (base {b:.1%})")
    ax.legend(loc="lower right")
    fig.tight_layout(); fig.savefig(out / "tornado.png", dpi=150); plt.close(fig); files.append("tornado.png")

    # 2 Monte Carlo histogram
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.hist(mc["irr"] * 100, bins=50, color="#2E75B6", edgecolor="white")
    for v, c, lab in ((mc["p5"], "#C00000", "P5"), (mc["median_irr"], "k", "Median"), (mc["p95"], "#2E7D32", "P95")):
        ax.axvline(v * 100, color=c, ls="--"); ax.text(v * 100, ax.get_ylim()[1] * .92, f" {lab} {v:.1%}", color=c, fontsize=8)
    ax.axvline(20, color="orange", lw=2); ax.text(20, ax.get_ylim()[1] * .8, " 20% hurdle", color="orange", fontsize=8)
    ax.set(xlabel="Sponsor IRR (%)", ylabel="Simulations", title=f"Monte Carlo IRR distribution ({mc['n']:,} paths)")
    fig.tight_layout(); fig.savefig(out / "monte_carlo_irr.png", dpi=150); plt.close(fig); files.append("monte_carlo_irr.png")

    # 3 value-creation waterfall
    br = base["bridge"]
    labels = ["Revenue\ngrowth", "Margin\nexpansion", "Multiple\nchange", "De-\nleveraging", "Fees", "MIP", "Sponsor\ngain"]
    vals = [br["revenue_growth"], br["margin_expansion"], br["multiple_change"], br["deleveraging"], br["fees"], br["mip"]]
    fig, ax = plt.subplots(figsize=(9, 4.5))
    cum = 0
    for i, v in enumerate(vals):
        ax.bar(i, v, bottom=cum if v >= 0 else cum + v, color="#2E7D32" if v >= 0 else "#C00000")
        ax.text(i, cum + max(v, 0) + 8, f"{v:+,.0f}", ha="center", fontsize=8)
        cum += v
    ax.bar(len(vals), cum, color="#1F3864"); ax.text(len(vals), cum + 8, f"{cum:,.0f}", ha="center", fontsize=8)
    ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels, fontsize=8)
    ax.set(ylabel="$ millions", title=f"Value creation bridge, exit Year {p.exit_year}")
    fig.tight_layout(); fig.savefig(out / "value_creation_bridge.png", dpi=150); plt.close(fig); files.append("value_creation_bridge.png")

    # 4 credit profile by scenario
    fig, ax = plt.subplots(1, 2, figsize=(11, 4.2))
    yrs = np.arange(0, N + 1)
    for s, colr in ((1, "#2E75B6"), (2, "#2E7D32"), (3, "#C00000")):
        r = run(p.with_(scenario=s))
        ax[0].plot(yrs, r["net_lev"], marker="o", color=colr, label=SCENARIOS[s])
        ax[1].plot(yrs[1:], r["int_cov"][1:], marker="o", color=colr, label=SCENARIOS[s])
    ax[0].plot(yrs[1:], p.cov_max, "k--", label="Covenant max")
    ax[0].set(title="Net debt / EBITDA", xlabel="Year"); ax[1].set(title="EBITDA / cash interest", xlabel="Year")
    ax[0].legend(); fig.tight_layout(); fig.savefig(out / "credit_profile.png", dpi=150); plt.close(fig); files.append("credit_profile.png")
    return files
