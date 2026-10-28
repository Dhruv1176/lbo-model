"""python -m lbo build      -> Excel model + charts + IC memo (+ reconciliation if LibreOffice is installed)
python -m lbo analyze    -> print tornado / break-evens / Monte Carlo
python -m lbo verify     -> recalculate the workbook and reconcile it to the engine"""
import argparse

from . import analysis as an
from . import build_excel as bx
from .engine import Params, run
from .memo import write_memo
from .reconcile import reconcile
from .recalc import recalc, soffice_path


def main(argv=None):
    ap = argparse.ArgumentParser(prog="lbo", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["build", "analyze", "verify"])
    ap.add_argument("--xlsx", default="models/Project_Atlas_LBO.xlsx")
    ap.add_argument("--sims", type=int, default=10000)
    a = ap.parse_args(argv)
    p = Params()
    if a.command in ("build", "verify"):
        bx.build(p, a.xlsx)
        print(f"Built {a.xlsx}")
        if recalc(a.xlsx):
            res = reconcile(a.xlsx, p)
            worst = max(res["max_abs_diff"].items(), key=lambda kv: kv[1])
            print(f"Recalculated with LibreOffice | master check: {res['master_check']} | worst Excel-vs-engine diff: {worst[1]:.2e} ({worst[0]})")
        else:
            print("LibreOffice not found: workbook has formulas only (Excel calculates on open). Install LibreOffice to reconcile.")
    if a.command in ("build", "analyze"):
        tor, be, mc = an.tornado(p), an.breakevens(p), an.monte_carlo(p, a.sims)
        print("\nIRR tornado:\n", tor.round(4).to_string(index=False))
        print("\nBreak-evens:", {k: round(float(v), 4) for k, v in be.items()})
        print("\nMonte Carlo:", {k: round(v, 3) for k, v in mc.items() if not hasattr(v, "shape")})
        if a.command == "build":
            an.charts(p, tor, mc, "docs")
            print("Memo:", write_memo(p, tor, be, mc, "docs/INVESTMENT_MEMO.md"))


if __name__ == "__main__":
    main()
