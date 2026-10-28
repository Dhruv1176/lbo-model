"""Compare a recalculated workbook against the Python engine (the model's independent audit)."""
from __future__ import annotations

import numpy as np
from openpyxl import load_workbook

from . import build_excel as bx
from .engine import Params, run

TOL = 1e-6


def _row(ws, key, sheet, cols):
    r = bx.R[sheet][key]
    return np.array([ws[f"{c}{r}"].value for c in cols], dtype=float)


def reconcile(path: str, p: Params) -> dict:
    """Returns {metric: max abs diff}. Requires bx.R to be populated (call build() first in this process)."""
    wb = load_workbook(path, data_only=True)
    e = run(p)
    om, ds, bs, rt, cs = wb["Operating Model"], wb["Debt Schedule"], wb["Balance Sheet"], wb["Returns"], wb["Credit Stats"]
    Y, A6 = bx.YC, bx.AC
    diffs = {
        "revenue": _row(om, "rev", "Operating Model", A6) - e["rev"],
        "ebitda": _row(om, "ebitda", "Operating Model", A6) - e["ebitda"],
        "taxes": _row(om, "taxes", "Operating Model", Y) - e["taxes"][1:],
        "net_income": _row(om, "ni", "Operating Model", Y) - e["ni"][1:],
        "fcf": _row(om, "fcf", "Operating Model", Y) - e["fcf"][1:],
        "unlevered_fcf": _row(om, "ufcf", "Operating Model", Y) - e["unlev_fcf"][1:],
        "cash": _row(ds, "end_cash", "Debt Schedule", A6) - e["cash"],
        "revolver": _row(ds, "rev_end", "Debt Schedule", A6) - e["rev_bal"],
        "tlb": _row(ds, "tlb_end", "Debt Schedule", A6) - e["tlb"],
        "cash_interest": _row(ds, "total_int", "Debt Schedule", Y) - e["cash_int"][1:],
        "net_debt": _row(ds, "net_debt", "Debt Schedule", A6) - e["net_debt"],
        "bs_check": _row(bs, "check", "Balance Sheet", A6),
        "moic": _row(rt, "moic", "Returns", Y) - e["moic"][1:],
        "irr": _row(rt, "irr", "Returns", Y) - e["irr"][1:],
        "net_leverage": _row(cs, "net_lev", "Credit Stats", A6) - e["net_lev"],
        "int_coverage": _row(cs, "int_cov", "Credit Stats", Y) - np.nan_to_num(e["int_cov"][1:]),
        "fccr": _row(cs, "fccr", "Credit Stats", Y) - e["fccr"][1:],
    }
    b = e["bridge"]
    for k, key in (("revenue_growth", "b_rev"), ("margin_expansion", "b_mar"), ("multiple_change", "b_mult"),
                   ("deleveraging", "b_del"), ("fees", "b_fee"), ("mip", "b_mip"), ("total", "b_tot")):
        diffs[f"bridge_{k}"] = np.array([rt[f"C{bx.R['Returns'][key]}"].value - b[k]])
    diffs["sponsor_equity"] = np.array([wb["Transaction"][f"C{bx.R['Transaction']['s_eq']}"].value - e["equity"]])
    master = wb["Checks"][f"C{bx.R['Checks']['master']}"].value
    return {"max_abs_diff": {k: float(np.max(np.abs(v))) for k, v in diffs.items()}, "master_check": master}
