"""Builds the formula-driven Excel LBO workbook. Every calculation is a live Excel formula; only inputs
(blue) and the clearly-labelled static snapshot tables (Python engine output) are values."""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .engine import N, SCENARIOS, Params, run

FONT = "Arial"
USD = '$#,##0.0;($#,##0.0);"-"'
NUM = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.0x;(0.0x);"-"'
MULT2 = '0.00x;(0.00x);"-"'
INT = "0"
YC = ["D", "E", "F", "G", "H"]          # Year 1..5
AC = ["C"] + YC                         # Year 0..5
NAVY, GREY, YELLOW = "1F3864", "D9D9D9", "FFFF00"

SHEETS = ["Summary", "Assumptions", "Transaction", "Operating Model", "Debt Schedule", "Balance Sheet",
          "Returns", "Credit Stats", "Sensitivity", "Scenarios", "Checks"]
R: dict[str, dict[str, int]] = {s: {} for s in SHEETS}


# ---------------------------------------------------------------- layout registry
def layout(sheet: str, items, start: int = 6):
    """items: list of (key, label, fmt, style). key None -> spacer/section. Assigns row numbers."""
    r = start
    out = []
    for it in items:
        key, label, fmt, style = (list(it) + [None, None, None])[:4]
        if key:
            R[sheet][key] = r
        out.append((r, key, label, fmt, style))
        r += 1
    return out


_DRY = {"on": False}


def ref(sheet: str, key: str, col: str, absolute: bool = False) -> str:
    if _DRY["on"] and key not in R[sheet]:
        row = 1          # first pass only registers row numbers; formulas are rewritten on pass 2
    else:
        row = R[sheet][key]
    return f"'{sheet}'!${col}${row}" if absolute else f"'{sheet}'!{col}{row}"


def A(key: str) -> str:   # absolute link to an Assumptions input (column C)
    return ref("Assumptions", key, "C", True)


def T(key: str, col: str = "C") -> str:   # absolute link to Transaction
    return ref("Transaction", key, col, True)


# ---------------------------------------------------------------- styling helpers
def f(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


BLUE, GREEN, BLACK, WHITE = "0000FF", "008000", "000000", "FFFFFF"
LINK = re.compile(r"^='?[A-Za-z ]+'?!\$?[A-Z]+\$?\d+$")


def put(ws, addr, value, fmt=None, bold=False, input_=False, key_input=False, italic=False, align=None):
    c = ws[addr]
    c.value = value
    color = BLACK
    if input_:
        color = BLUE
    elif isinstance(value, str) and value.startswith("=") and LINK.match(value):
        color = GREEN
    c.font = f(bold=bold, color=color, italic=italic)
    if fmt:
        c.number_format = fmt
    if key_input:
        c.fill = PatternFill("solid", fgColor=YELLOW)
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def header(ws, title, subtitle="($ in millions unless stated)  |  Project Atlas is a FICTIONAL company; all inputs are illustrative",
           years=True, ncols=9):
    ws.sheet_view.showGridLines = False
    ws["A1"].value, ws["A1"].font = title, f(True, NAVY, 14)
    ws["A2"].value, ws["A2"].font = subtitle, f(italic=True, color="595959", size=9)
    if years:
        for col, lab in zip(["A", "B"] + AC, ["", "Unit", "Close / LTM"] + [f"Year {i}" for i in range(1, N + 1)]):
            c = ws[f"{col}4"]
            c.value, c.font = lab, f(True, WHITE)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center" if col not in "AB" else "left")
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 11
    for col in AC:
        ws.column_dimensions[col].width = 13
    ws.freeze_panes = "C5" if years else None


def label_rows(ws, items, ncols=8):
    for r, key, label, fmt, style in items:
        c = ws[f"A{r}"]
        c.value = label
        if style == "section":
            for col in range(1, ncols + 1):
                ws.cell(r, col).fill = PatternFill("solid", fgColor=GREY)
            c.font = f(True)
        elif style in ("total", "bold"):
            c.font = f(True)
        else:
            c.font = f()


def fill(ws, sheet, key, y0=None, yt=None, fmt=USD, bold=False, unit="$mm", border=False):
    r = R[sheet][key]
    ws[f"B{r}"].value, ws[f"B{r}"].font = unit, f(color="7F7F7F", size=9)
    if y0 is not None:
        put(ws, f"C{r}", y0, fmt, bold)
    if yt is not None:
        for i, col in enumerate(YC):
            prev = AC[i]
            put(ws, f"{col}{r}", yt(col, prev), fmt, bold)
    if border:
        for col in ["A", "B"] + AC:
            ws[f"{col}{r}"].border = Border(top=Side(style="thin"))


# ---------------------------------------------------------------- the workbook
def build(p: Params | None = None, path: str = "models/Project_Atlas_LBO.xlsx") -> str:
    """Two passes: pass 1 registers every row number (so sheets can reference each other), pass 2 writes the file."""
    _DRY["on"] = True
    _build(p or Params(), path)
    _DRY["on"] = False
    return _build(p or Params(), path)


def _build(p: Params, path: str) -> str:
    wb = Workbook()
    wb.remove(wb.active)
    ws = {s: wb.create_sheet(s) for s in SHEETS}

    # ======================= Assumptions =======================
    sh = "Assumptions"
    items = [
        (None, "GENERAL", None, "section"),
        ("scenario", "Operating scenario (1 = Base, 2 = Upside, 3 = Downside)", INT, ""),
        ("ltm_rev", "LTM revenue", USD, ""), ("ltm_ebitda", "LTM EBITDA", USD, ""),
        ("ppe_open", "Net PP&E at close", USD, ""), (None, None, None, None),
        (None, "TRANSACTION", None, "section"),
        ("entry_mult", "Entry EV / LTM EBITDA", MULT, ""), ("txn_fee_pct", "Transaction fees (% of EV)", PCT, ""),
        ("fin_fee_pct", "Financing fees (% of debt raised + revolver commitment)", PCT, ""),
        ("fee_amort_yrs", "Financing fee amortisation period (years)", INT, ""),
        ("min_cash", "Minimum cash (funded at close)", USD, ""), (None, None, None, None),
        (None, "EXIT", None, "section"),
        ("exit_year", "Exit year (1 to 5)", INT, ""), ("exit_mult", "Exit EV / EBITDA", MULT, ""),
        ("mip_pct", "Management incentive plan (% of equity gain)", PCT, ""), (None, None, None, None),
        (None, "FINANCING", None, "section"),
        ("rev_commit", "Revolver commitment", USD, ""), ("rev_drawn", "Revolver drawn at close", USD, ""),
        ("rev_spread", "Revolver spread over SOFR", PCT, ""), ("rev_undrawn", "Revolver undrawn commitment fee", PCT, ""),
        ("tlb_x", "Term Loan B (x LTM EBITDA)", MULT, ""), ("tlb_spread", "Term Loan B spread over SOFR", PCT, ""),
        ("tlb_amort", "Term Loan B mandatory amortisation (% of original, p.a.)", PCT, ""),
        ("sweep_pct", "Excess cash flow sweep to Term Loan B", PCT, ""),
        ("notes_x", "Senior unsecured notes (x LTM EBITDA)", MULT, ""), ("notes_rate", "Senior notes fixed coupon", PCT, ""),
        ("cash_rate", "Interest rate earned on cash", PCT, ""), (None, None, None, None),
        (None, "TAX AND OPERATING", None, "section"),
        ("tax_rate", "Cash tax rate", PCT, ""), ("int_cap_toggle", "Apply 163(j)-style interest cap (1 = yes)", INT, ""),
        ("int_cap_pct", "Interest deduction cap (% of EBITDA)", PCT, ""),
        ("da_pct", "D&A (% of revenue)", PCT, ""), ("capex_pct", "Capex (% of revenue)", PCT, ""),
        ("nwc_pct", "Net working capital (% of revenue)", PCT, ""), (None, None, None, None),
        (None, "TIME-SERIES INPUTS", None, "section"),
        ("sofr", "SOFR forward curve", PCT, ""), ("cov_max", "Maximum Net Debt / EBITDA covenant", MULT2, ""),
        (None, None, None, None),
        (None, "SCENARIO LEVERS", None, "section"),
        ("g1", "Revenue growth: Base", PCT, ""), ("g2", "Revenue growth: Upside", PCT, ""), ("g3", "Revenue growth: Downside", PCT, ""),
        ("m1", "EBITDA margin: Base", PCT, ""), ("m2", "EBITDA margin: Upside", PCT, ""), ("m3", "EBITDA margin: Downside", PCT, ""),
        ("g_sel", "Revenue growth: SELECTED", PCT, "bold"), ("m_sel", "EBITDA margin: SELECTED", PCT, "bold"),
    ]
    L = layout(sh, items)
    w = ws[sh]
    header(w, "Assumptions & Inputs")
    label_rows(w, L)
    w["I4"].value, w["I4"].font = "Basis / source", f(True, WHITE)
    w["I4"].fill = PatternFill("solid", fgColor=NAVY)
    w.column_dimensions["I"].width = 78
    scalar = {
        "scenario": (p.scenario, "Scenario switch: drives the whole workbook live", True),
        "ltm_rev": (p.ltm_rev, "Illustrative assumption (fictional company)", False),
        "ltm_ebitda": (p.ltm_ebitda, "Illustrative assumption: 20% LTM EBITDA margin", False),
        "ppe_open": (p.ppe_open, "Illustrative assumption; used for the opening balance sheet", False),
        "entry_mult": (p.entry_mult, "Illustrative; mid-market industrial services deals typically price ~9x to 13x", True),
        "txn_fee_pct": (p.txn_fee_pct, "Advisory, legal, diligence; capitalised into goodwill (simplification)", False),
        "fin_fee_pct": (p.fin_fee_pct, "Underwriting / OID proxy; capitalised and amortised straight-line, tax-deductible", False),
        "fee_amort_yrs": (p.fee_amort_yrs, "Matches Term Loan B tenor", False),
        "min_cash": (p.min_cash, "Operating cash buffer; revolver funds any shortfall below this level", False),
        "exit_year": (p.exit_year, "Hold period in years (1 to 5)", True),
        "exit_mult": (p.exit_mult, "Base case assumes NO multiple expansion (exit = entry)", True),
        "mip_pct": (p.mip_pct, "Illustrative; applied to the equity gain at exit", False),
        "rev_commit": (p.rev_commit, "Illustrative revolving credit facility size", False),
        "rev_drawn": (p.rev_drawn, "Undrawn at close", False),
        "rev_spread": (p.rev_spread, "Illustrative", False), "rev_undrawn": (p.rev_undrawn, "Illustrative", False),
        "tlb_x": (p.tlb_x, "Illustrative senior secured leverage", True),
        "tlb_spread": (p.tlb_spread, "Illustrative spread for a single-B credit", False),
        "tlb_amort": (p.tlb_amort, "Standard 1% p.a. mandatory amortisation", False),
        "sweep_pct": (p.sweep_pct, "Excess cash flow sweep; remainder builds on balance sheet", False),
        "notes_x": (p.notes_x, "Illustrative junior leverage; bullet maturity, non-callable in model", True),
        "notes_rate": (p.notes_rate, "Illustrative fixed coupon", False),
        "cash_rate": (p.cash_rate, "Interest on beginning-of-year cash", False),
        "tax_rate": (p.tax, "Blended federal + state, illustrative", False),
        "int_cap_toggle": (p.int_cap_on, "1 = deduct at most cap% x EBITDA of net interest; disallowed amount is lost (no carry-forward, conservative)", False),
        "int_cap_pct": (p.int_cap_pct, "Section 163(j) uses ~30% of adjusted taxable income", False),
        "da_pct": (p.da_pct, "Illustrative", False), "capex_pct": (p.capex_pct, "Illustrative; above D&A = growth capex", False),
        "nwc_pct": (p.nwc_pct, "Illustrative; cash is absorbed as revenue grows", False),
    }
    for k, (v, note, key_in) in scalar.items():
        r = R[sh][k]
        fmt = [it[3] for it in items if it[0] == k][0]
        put(w, f"C{r}", v, fmt, input_=True, key_input=key_in)
        w[f"I{r}"].value, w[f"I{r}"].font = note, f(color="595959", size=9)
    ts = {"sofr": (p.sofr, "Illustrative forward curve; Term Loan B and revolver float over this"),
          "cov_max": (p.cov_max, "Illustrative maintenance covenant with annual step-downs"),
          "g1": (p.growth[1], "Base: steady mid-single-digit organic growth"),
          "g2": (p.growth[2], "Upside: share gains and pricing"),
          "g3": (p.growth[3], "Downside: mild recession in Year 2, slow recovery"),
          "m1": (p.margin[1], "Base: modest operating leverage"), "m2": (p.margin[2], "Upside: stronger operating leverage"),
          "m3": (p.margin[3], "Downside: margin compression in the recession")}
    for k, (vals, note) in ts.items():
        r = R[sh][k]
        fmt = [it[3] for it in items if it[0] == k][0]
        for col, v in zip(YC, vals):
            put(w, f"{col}{r}", v, fmt, input_=True)
        w[f"I{r}"].value, w[f"I{r}"].font = note, f(color="595959", size=9)
    for key, a, b, c3 in (("g_sel", "g1", "g2", "g3"), ("m_sel", "m1", "m2", "m3")):
        r = R[sh][key]
        for col in YC:
            put(w, f"{col}{r}", f"=CHOOSE({A('scenario')},{col}{R[sh][a]},{col}{R[sh][b]},{col}{R[sh][c3]})", PCT, True)
    w[f"I{R[sh]['g_sel']}"].value = "Formula: picks the scenario chosen above"
    w[f"I{R[sh]['g_sel']}"].font = f(color="595959", size=9)
    w["A3"].value = "Legend: blue = input you can change · black = formula · green = link to another sheet · yellow fill = key deal lever"
    w["A3"].font = f(color="595959", size=9, italic=True)

    # ======================= Transaction =======================
    sh = "Transaction"
    items = [
        (None, "PURCHASE PRICE", None, "section"),
        ("ltm_e", "LTM EBITDA", USD, ""), ("mult", "Entry multiple (EV / LTM EBITDA)", MULT, ""),
        ("ev", "Enterprise value (cash-free, debt-free)", USD, "bold"), ("ev_rev", "EV / LTM revenue", MULT2, ""),
        (None, None, None, None),
        (None, "SOURCES OF FUNDS", None, "section"),
        ("s_rev", "Revolver (drawn)", USD, ""), ("s_tlb", "Term Loan B", USD, ""), ("s_notes", "Senior unsecured notes", USD, ""),
        ("s_eq", "Sponsor equity (plug)", USD, ""), ("s_tot", "Total sources", USD, "total"), (None, None, None, None),
        (None, "USES OF FUNDS", None, "section"),
        ("u_ev", "Purchase of enterprise", USD, ""), ("u_txn", "Transaction fees", USD, ""), ("u_fin", "Financing fees (capitalised)", USD, ""),
        ("u_cash", "Cash to balance sheet (minimum cash)", USD, ""), ("u_tot", "Total uses", USD, "total"), (None, None, None, None),
        (None, "DERIVED", None, "section"),
        ("eq_pct", "Sponsor equity as % of total capitalisation", PCT, ""), ("nwc0", "Opening net working capital", USD, ""),
        ("gw", "Goodwill created (EV + fees - NWC - net PP&E)", USD, ""), ("fee_am", "Annual financing fee amortisation", USD, ""),
    ]
    L = layout(sh, items, start=6)
    w = ws[sh]
    header(w, "Transaction: Sources & Uses", years=False)
    label_rows(w, L, 5)
    for col, lab in zip("ABCDE", ["", "", "$mm", "% of total", "x LTM EBITDA"]):
        c = w[f"{col}4"]; c.value, c.font, c.fill = lab, f(True, WHITE), PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center")
    w.column_dimensions["A"].width = 52
    for col in "BCDE":
        w.column_dimensions[col].width = 15
    rr = lambda k: R[sh][k]
    vals = {"ltm_e": (f"={A('ltm_ebitda')}", USD), "mult": (f"={A('entry_mult')}", MULT),
            "ev": (f"=C{rr('ltm_e')}*C{rr('mult')}", USD), "ev_rev": (f"=C{rr('ev')}/{A('ltm_rev')}", MULT2),
            "s_rev": (f"={A('rev_drawn')}", USD), "s_tlb": (f"={A('tlb_x')}*C{rr('ltm_e')}", USD),
            "s_notes": (f"={A('notes_x')}*C{rr('ltm_e')}", USD),
            "s_eq": (f"=C{rr('u_tot')}-C{rr('s_rev')}-C{rr('s_tlb')}-C{rr('s_notes')}", USD),
            "s_tot": (f"=SUM(C{rr('s_rev')}:C{rr('s_eq')})", USD),
            "u_ev": (f"=C{rr('ev')}", USD), "u_txn": (f"={A('txn_fee_pct')}*C{rr('ev')}", USD),
            "u_fin": (f"=({A('tlb_x')}*C{rr('ltm_e')}+{A('notes_x')}*C{rr('ltm_e')}+{A('rev_commit')})*{A('fin_fee_pct')}", USD),
            "u_cash": (f"={A('min_cash')}", USD), "u_tot": (f"=SUM(C{rr('u_ev')}:C{rr('u_cash')})", USD),
            "eq_pct": (f"=C{rr('s_eq')}/C{rr('s_tot')}", PCT), "nwc0": (f"={A('nwc_pct')}*{A('ltm_rev')}", USD),
            "gw": (f"=C{rr('u_ev')}+C{rr('u_txn')}-C{rr('nwc0')}-{A('ppe_open')}", USD),
            "fee_am": (f"=C{rr('u_fin')}/{A('fee_amort_yrs')}", USD)}
    for k, (fx, fm) in vals.items():
        put(w, f"C{rr(k)}", fx, fm, bold=k in ("ev", "s_tot", "u_tot"))
    for k in ("s_rev", "s_tlb", "s_notes", "s_eq", "s_tot"):
        put(w, f"D{rr(k)}", f"=C{rr(k)}/C${rr('s_tot')}", PCT)
        put(w, f"E{rr(k)}", f"=C{rr(k)}/C${rr('ltm_e')}", MULT)
    for k in ("u_ev", "u_txn", "u_fin", "u_cash", "u_tot"):
        put(w, f"D{rr(k)}", f"=C{rr(k)}/C${rr('u_tot')}", PCT)
        put(w, f"E{rr(k)}", f"=C{rr(k)}/C${rr('ltm_e')}", MULT)
    for k in ("s_tot", "u_tot"):
        for col in "ABCDE":
            w[f"{col}{rr(k)}"].border = Border(top=Side(style="thin"))
    w["G6"].value = "Note: sponsor equity is the balancing item. Transaction fees are capitalised into goodwill for simplicity."
    w["G6"].font = f(italic=True, color="595959", size=9)

    # ======================= Operating Model =======================
    sh = "Operating Model"
    items = [
        (None, "INCOME STATEMENT", None, "section"),
        ("rev", "Revenue", USD, "bold"), ("growth", "  % growth", PCT, ""), ("ebitda", "EBITDA", USD, "bold"), ("margin", "  % margin", PCT, ""),
        ("da", "Less: depreciation & amortisation", USD, ""), ("ebit", "EBIT", USD, "total"),
        ("cash_int", "Less: cash interest expense (incl. undrawn fees)", USD, ""), ("int_inc", "Plus: interest income", USD, ""),
        ("fee_am", "Less: amortisation of financing fees", USD, ""), ("ebt", "Pre-tax income", USD, "total"),
        (None, None, None, None),
        (None, "TAX CALCULATION", None, "section"),
        ("net_int", "Net interest expense (incl. fee amortisation)", USD, ""), ("cap", "Deductible interest cap (cap % x EBITDA)", USD, ""),
        ("disallowed", "Disallowed interest (added back for tax)", USD, ""), ("taxable", "Taxable income", USD, ""),
        ("taxes", "Cash taxes", USD, "bold"), ("ni", "Net income", USD, "total"), (None, None, None, None),
        (None, "CASH FLOW", None, "section"),
        ("c_ni", "Net income", USD, ""), ("c_da", "Plus: D&A", USD, ""), ("c_fee", "Plus: financing fee amortisation (non-cash)", USD, ""),
        ("nwc", "Net working capital (balance)", USD, ""), ("d_nwc", "Less: increase in NWC", USD, ""), ("capex", "Less: capital expenditures", USD, ""),
        ("fcf", "Free cash flow available for debt repayment", USD, "total"), ("fcf_conv", "  FCF / EBITDA", PCT, ""),
        (None, None, None, None),
        (None, "UNLEVERED FREE CASH FLOW (memo)", None, "section"),
        ("u_tax", "Unlevered taxes (EBIT x tax rate)", USD, ""), ("ufcf", "Unlevered free cash flow", USD, "total"),
    ]
    L = layout(sh, items)
    w = ws[sh]
    header(w, "Operating Model: Income Statement & Cash Flow")
    label_rows(w, L)
    S = sh
    rw = lambda k, S="Operating Model": R[S][k]
    fill(w, S, "rev", f"={A('ltm_rev')}", lambda c, pv: f"={pv}{rw('rev')}*(1+{ref('Assumptions', 'g_sel', c)})", bold=True)
    fill(w, S, "growth", None, lambda c, pv: f"={c}{rw('rev')}/{pv}{rw('rev')}-1", PCT, unit="%")
    fill(w, S, "ebitda", f"={A('ltm_ebitda')}", lambda c, pv: f"={c}{rw('rev')}*{ref('Assumptions', 'm_sel', c)}", bold=True)
    for c in AC:
        put(w, f"{c}{rw('margin')}", f"={c}{rw('ebitda')}/{c}{rw('rev')}", PCT)
    w[f"B{rw('margin')}"].value = "%"
    fill(w, S, "da", None, lambda c, pv: f"={c}{rw('rev')}*{A('da_pct')}")
    fill(w, S, "ebit", None, lambda c, pv: f"={c}{rw('ebitda')}-{c}{rw('da')}", bold=True, border=True)
    fill(w, S, "cash_int", None, lambda c, pv: f"={ref('Debt Schedule', 'total_int', c)}")
    fill(w, S, "int_inc", None, lambda c, pv: f"={ref('Debt Schedule', 'int_income', c)}")
    fill(w, S, "fee_am", None, lambda c, pv: f"={T('fee_am')}")
    fill(w, S, "ebt", None, lambda c, pv: f"={c}{rw('ebit')}-{c}{rw('cash_int')}+{c}{rw('int_inc')}-{c}{rw('fee_am')}", bold=True, border=True)
    fill(w, S, "net_int", None, lambda c, pv: f"={c}{rw('cash_int')}+{c}{rw('fee_am')}-{c}{rw('int_inc')}")
    fill(w, S, "cap", None, lambda c, pv: f"=IF({A('int_cap_toggle')}=1,MIN({c}{rw('net_int')},{A('int_cap_pct')}*{c}{rw('ebitda')}),{c}{rw('net_int')})")
    fill(w, S, "disallowed", None, lambda c, pv: f"={c}{rw('net_int')}-{c}{rw('cap')}")
    fill(w, S, "taxable", None, lambda c, pv: f"={c}{rw('ebt')}+{c}{rw('disallowed')}")
    fill(w, S, "taxes", None, lambda c, pv: f"=MAX(0,{c}{rw('taxable')})*{A('tax_rate')}", bold=True)
    fill(w, S, "ni", None, lambda c, pv: f"={c}{rw('ebt')}-{c}{rw('taxes')}", bold=True, border=True)
    fill(w, S, "c_ni", None, lambda c, pv: f"={c}{rw('ni')}")
    fill(w, S, "c_da", None, lambda c, pv: f"={c}{rw('da')}")
    fill(w, S, "c_fee", None, lambda c, pv: f"={c}{rw('fee_am')}")
    fill(w, S, "nwc", f"={A('nwc_pct')}*C{rw('rev')}", lambda c, pv: f"={A('nwc_pct')}*{c}{rw('rev')}")
    fill(w, S, "d_nwc", None, lambda c, pv: f"={c}{rw('nwc')}-{pv}{rw('nwc')}")
    fill(w, S, "capex", None, lambda c, pv: f"={c}{rw('rev')}*{A('capex_pct')}")
    fill(w, S, "fcf", None, lambda c, pv: f"={c}{rw('c_ni')}+{c}{rw('c_da')}+{c}{rw('c_fee')}-{c}{rw('d_nwc')}-{c}{rw('capex')}", bold=True, border=True)
    fill(w, S, "fcf_conv", None, lambda c, pv: f"={c}{rw('fcf')}/{c}{rw('ebitda')}", PCT, unit="%")
    fill(w, S, "u_tax", None, lambda c, pv: f"={c}{rw('ebit')}*{A('tax_rate')}")
    fill(w, S, "ufcf", None, lambda c, pv: f"={c}{rw('ebitda')}-{c}{rw('u_tax')}-{c}{rw('capex')}-{c}{rw('d_nwc')}", bold=True, border=True)

    # ======================= Debt Schedule =======================
    sh = "Debt Schedule"
    items = [
        (None, "CASH WATERFALL", None, "section"),
        ("beg_cash", "Beginning cash", USD, ""), ("fcf", "Plus: free cash flow (after interest and taxes)", USD, ""),
        ("mand_w", "Less: mandatory Term Loan B amortisation", USD, ""), ("min_c", "Less: minimum cash", USD, ""),
        ("avail", "Cash available for revolver / sweep (negative = shortfall)", USD, "total"),
        ("rev_flow", "Revolver draw / (repayment)", USD, ""), ("sw_av", "Cash available for sweep", USD, ""),
        ("sweep_w", "Term Loan B cash sweep", USD, ""), ("end_cash", "Ending cash", USD, "total"), (None, None, None, None),
        (None, "REVOLVER", None, "section"),
        ("rev_beg", "Beginning balance", USD, ""), ("rev_chg", "Draw / (repayment)", USD, ""), ("rev_end", "Ending balance", USD, "total"),
        ("rev_room", "Undrawn availability", USD, ""), (None, None, None, None),
        (None, "TERM LOAN B", None, "section"),
        ("tlb_beg", "Beginning balance", USD, ""), ("mand", "Mandatory amortisation", USD, ""), ("sweep", "Cash sweep", USD, ""),
        ("tlb_end", "Ending balance", USD, "total"), (None, None, None, None),
        (None, "SENIOR UNSECURED NOTES (bullet)", None, "section"),
        ("n_beg", "Beginning balance", USD, ""), ("n_end", "Ending balance", USD, "total"), (None, None, None, None),
        (None, "TOTALS", None, "section"),
        ("tot_debt", "Total debt", USD, "bold"), ("net_debt", "Net debt (total debt less cash)", USD, "bold"), (None, None, None, None),
        (None, "INTEREST (on beginning balances, no circularity)", None, "section"),
        ("sofr", "SOFR", PCT, ""), ("rev_rate", "Revolver all-in rate", PCT, ""), ("tlb_rate", "Term Loan B all-in rate", PCT, ""),
        ("rev_int", "Revolver interest", USD, ""), ("undrawn", "Revolver undrawn fee", USD, ""), ("tlb_int", "Term Loan B interest", USD, ""),
        ("n_int", "Senior notes interest", USD, ""), ("total_int", "Total cash interest expense", USD, "total"),
        ("int_income", "Interest income on beginning cash", USD, ""),
    ]
    L = layout(sh, items)
    w = ws[sh]
    header(w, "Debt Schedule & Cash Sweep")
    label_rows(w, L)
    S = sh
    d = lambda k, S="Debt Schedule": R[S][k]
    O = lambda k, c: ref("Operating Model", k, c)
    fill(w, S, "beg_cash", None, lambda c, pv: f"={pv}{d('end_cash')}")
    fill(w, S, "fcf", None, lambda c, pv: f"={O('fcf', c)}")
    fill(w, S, "mand_w", None, lambda c, pv: f"={c}{d('mand')}")
    fill(w, S, "min_c", None, lambda c, pv: f"={A('min_cash')}")
    fill(w, S, "avail", None, lambda c, pv: f"={c}{d('beg_cash')}+{c}{d('fcf')}-{c}{d('mand_w')}-{c}{d('min_c')}", bold=True, border=True)
    fill(w, S, "rev_flow", None, lambda c, pv: f"=IF({c}{d('avail')}<0,MIN(-{c}{d('avail')},{A('rev_commit')}-{c}{d('rev_beg')}),-MIN({c}{d('rev_beg')},{c}{d('avail')}))")
    fill(w, S, "sw_av", None, lambda c, pv: f"=MAX(0,{c}{d('avail')}+{c}{d('rev_flow')})")
    fill(w, S, "sweep_w", None, lambda c, pv: f"={c}{d('sweep')}")
    fill(w, S, "end_cash", f"={T('u_cash')}", lambda c, pv: f"={c}{d('beg_cash')}+{c}{d('fcf')}-{c}{d('mand_w')}+{c}{d('rev_flow')}-{c}{d('sweep_w')}", bold=True, border=True)
    fill(w, S, "rev_beg", None, lambda c, pv: f"={pv}{d('rev_end')}")
    fill(w, S, "rev_chg", None, lambda c, pv: f"={c}{d('rev_flow')}")
    fill(w, S, "rev_end", f"={T('s_rev')}", lambda c, pv: f"={c}{d('rev_beg')}+{c}{d('rev_chg')}", bold=True, border=True)
    fill(w, S, "rev_room", f"={A('rev_commit')}-C{d('rev_end')}", lambda c, pv: f"={A('rev_commit')}-{c}{d('rev_end')}")
    fill(w, S, "tlb_beg", None, lambda c, pv: f"={pv}{d('tlb_end')}")
    fill(w, S, "mand", None, lambda c, pv: f"=MIN({c}{d('tlb_beg')},{A('tlb_amort')}*$C{d('tlb_end')})")
    fill(w, S, "sweep", None, lambda c, pv: f"=MIN({c}{d('tlb_beg')}-{c}{d('mand')},{A('sweep_pct')}*{c}{d('sw_av')})")
    fill(w, S, "tlb_end", f"={T('s_tlb')}", lambda c, pv: f"={c}{d('tlb_beg')}-{c}{d('mand')}-{c}{d('sweep')}", bold=True, border=True)
    fill(w, S, "n_beg", None, lambda c, pv: f"={pv}{d('n_end')}")
    fill(w, S, "n_end", f"={T('s_notes')}", lambda c, pv: f"={c}{d('n_beg')}", bold=True, border=True)
    for c in AC:
        put(w, f"{c}{d('tot_debt')}", f"={c}{d('rev_end')}+{c}{d('tlb_end')}+{c}{d('n_end')}", USD, True)
        put(w, f"{c}{d('net_debt')}", f"={c}{d('tot_debt')}-{c}{d('end_cash')}", USD, True)
    fill(w, S, "sofr", None, lambda c, pv: f"={ref('Assumptions', 'sofr', c)}", PCT, unit="%")
    fill(w, S, "rev_rate", None, lambda c, pv: f"={c}{d('sofr')}+{A('rev_spread')}", PCT, unit="%")
    fill(w, S, "tlb_rate", None, lambda c, pv: f"={c}{d('sofr')}+{A('tlb_spread')}", PCT, unit="%")
    fill(w, S, "rev_int", None, lambda c, pv: f"={c}{d('rev_beg')}*{c}{d('rev_rate')}")
    fill(w, S, "undrawn", None, lambda c, pv: f"=({A('rev_commit')}-{c}{d('rev_beg')})*{A('rev_undrawn')}")
    fill(w, S, "tlb_int", None, lambda c, pv: f"={c}{d('tlb_beg')}*{c}{d('tlb_rate')}")
    fill(w, S, "n_int", None, lambda c, pv: f"={c}{d('n_beg')}*{A('notes_rate')}")
    fill(w, S, "total_int", None, lambda c, pv: f"=SUM({c}{d('rev_int')}:{c}{d('n_int')})", bold=True, border=True)
    fill(w, S, "int_income", None, lambda c, pv: f"={c}{d('beg_cash')}*{A('cash_rate')}")

    # ======================= Balance Sheet =======================
    sh = "Balance Sheet"
    items = [
        (None, "ASSETS", None, "section"),
        ("cash", "Cash", USD, ""), ("nwc", "Net working capital", USD, ""), ("ppe", "Net PP&E", USD, ""), ("gw", "Goodwill & intangibles", USD, ""),
        ("fees", "Capitalised financing fees", USD, ""), ("ta", "Total assets", USD, "total"), (None, None, None, None),
        (None, "LIABILITIES & EQUITY", None, "section"),
        ("rev", "Revolver", USD, ""), ("tlb", "Term Loan B", USD, ""), ("notes", "Senior unsecured notes", USD, ""),
        ("debt", "Total debt", USD, "bold"), ("sp_eq", "Sponsor equity contributed", USD, ""), ("re", "Retained earnings", USD, ""),
        ("te", "Total equity", USD, "bold"), ("tle", "Total liabilities & equity", USD, "total"), (None, None, None, None),
        ("check", "Balance check (assets - liabilities & equity)", NUM, "bold"),
    ]
    L = layout(sh, items)
    w = ws[sh]
    header(w, "Pro-Forma Balance Sheet")
    label_rows(w, L)
    S = sh
    b = lambda k, S="Balance Sheet": R[S][k]
    D = lambda k, c: ref("Debt Schedule", k, c)
    fill(w, S, "cash", f"={T('u_cash')}", lambda c, pv: f"={D('end_cash', c)}")
    fill(w, S, "nwc", f"={O('nwc', 'C')}", lambda c, pv: f"={O('nwc', c)}")
    fill(w, S, "ppe", f"={A('ppe_open')}", lambda c, pv: f"={pv}{b('ppe')}+{O('capex', c)}-{O('da', c)}")
    fill(w, S, "gw", f"={T('gw')}", lambda c, pv: f"={pv}{b('gw')}")
    fill(w, S, "fees", f"={T('u_fin')}", lambda c, pv: f"={pv}{b('fees')}-{O('fee_am', c)}")
    for c in AC:
        put(w, f"{c}{b('ta')}", f"=SUM({c}{b('cash')}:{c}{b('fees')})", USD, True)
        put(w, f"{c}{b('debt')}", f"=SUM({c}{b('rev')}:{c}{b('notes')})", USD, True)
        put(w, f"{c}{b('te')}", f"={c}{b('sp_eq')}+{c}{b('re')}", USD, True)
        put(w, f"{c}{b('tle')}", f"={c}{b('debt')}+{c}{b('te')}", USD, True)
        put(w, f"{c}{b('check')}", f"=ROUND({c}{b('ta')}-{c}{b('tle')},6)", NUM, True)
    fill(w, S, "rev", f"={D('rev_end', 'C')}", lambda c, pv: f"={D('rev_end', c)}")
    fill(w, S, "tlb", f"={D('tlb_end', 'C')}", lambda c, pv: f"={D('tlb_end', c)}")
    fill(w, S, "notes", f"={D('n_end', 'C')}", lambda c, pv: f"={D('n_end', c)}")
    fill(w, S, "sp_eq", f"={T('s_eq')}", lambda c, pv: f"={pv}{b('sp_eq')}")
    fill(w, S, "re", 0, lambda c, pv: f"={pv}{b('re')}+{O('ni', c)}")
    for k in ("ta", "tle"):
        for col in ["A", "B"] + AC:
            w[f"{col}{b(k)}"].border = Border(top=Side(style="thin"))

    # ======================= Returns =======================
    sh = "Returns"
    items = [
        (None, "RETURNS BY EXIT YEAR (exit at the Exit Multiple on that year's EBITDA)", None, "section"),
        ("n", "Exit year", INT, ""), ("ebitda_x", "EBITDA at exit", USD, ""), ("ev_x", "Enterprise value at exit", USD, ""),
        ("nd_x", "Less: net debt at exit", USD, ""), ("eq_x", "Equity value at exit (100%)", USD, "bold"),
        ("eq_in", "Sponsor equity invested at close", USD, ""), ("gain", "Equity gain", USD, ""),
        ("mip", "Less: management incentive plan", USD, ""), ("proceeds", "Sponsor proceeds", USD, "total"),
        ("moic", "MOIC", MULT2, "bold"), ("irr", "IRR", PCT, "bold"), (None, None, None, None),
        (None, "SELECTED EXIT YEAR", None, "section"),
        ("x", "Exit year selected", INT, ""), ("sel_proc", "Sponsor proceeds", USD, ""), ("sel_moic", "MOIC", MULT2, "bold"), ("sel_irr", "IRR", PCT, "bold"),
        ("cf", "Sponsor cash flows (for IRR function)", USD, ""), ("irr_fn", "IRR via Excel IRR() function (cross-check)", PCT, ""),
        (None, None, None, None),
        (None, "VALUE CREATION BRIDGE (selected exit year)", None, "section"),
        ("b_rev", "Revenue growth (at entry margin and multiple)", USD, ""), ("b_mar", "Margin expansion (at entry multiple)", USD, ""),
        ("b_mult", "Multiple expansion / (contraction)", USD, ""), ("b_del", "Deleveraging (net debt reduction)", USD, ""),
        ("b_fee", "Fees and expenses", USD, ""), ("b_mip", "Management incentive plan", USD, ""),
        ("b_tot", "Total sponsor gain", USD, "total"), ("b_chk", "Check vs. proceeds less equity invested", NUM, ""),
    ]
    L = layout(sh, items)
    w = ws[sh]
    header(w, "Returns Analysis")
    label_rows(w, L)
    S = sh
    q = lambda k, S="Returns": R[S][k]
    for i, c in enumerate(YC):
        put(w, f"{c}{q('n')}", i + 1, INT)
    fill(w, S, "ebitda_x", None, lambda c, pv: f"={O('ebitda', c)}")
    fill(w, S, "ev_x", None, lambda c, pv: f"={A('exit_mult')}*{c}{q('ebitda_x')}")
    fill(w, S, "nd_x", None, lambda c, pv: f"={D('net_debt', c)}")
    fill(w, S, "eq_x", None, lambda c, pv: f"={c}{q('ev_x')}-{c}{q('nd_x')}", bold=True)
    fill(w, S, "eq_in", None, lambda c, pv: f"={T('s_eq')}")
    fill(w, S, "gain", None, lambda c, pv: f"={c}{q('eq_x')}-{c}{q('eq_in')}")
    fill(w, S, "mip", None, lambda c, pv: f"={A('mip_pct')}*MAX(0,{c}{q('gain')})")
    fill(w, S, "proceeds", None, lambda c, pv: f"={c}{q('eq_x')}-{c}{q('mip')}", bold=True, border=True)
    fill(w, S, "moic", None, lambda c, pv: f"={c}{q('proceeds')}/{c}{q('eq_in')}", MULT2, True, unit="x")
    fill(w, S, "irr", None, lambda c, pv: f"=IF({c}{q('proceeds')}>0,({c}{q('proceeds')}/{c}{q('eq_in')})^(1/{c}{q('n')})-1,-1)", PCT, True, unit="%")
    rng = lambda k: f"$D{q(k)}:$H{q(k)}"
    put(w, f"C{q('x')}", f"={A('exit_year')}", INT)
    put(w, f"C{q('sel_proc')}", f"=INDEX({rng('proceeds')},C{q('x')})", USD)
    put(w, f"C{q('sel_moic')}", f"=INDEX({rng('moic')},C{q('x')})", MULT2, True)
    put(w, f"C{q('sel_irr')}", f"=INDEX({rng('irr')},C{q('x')})", PCT, True)
    put(w, f"C{q('cf')}", f"=-{T('s_eq')}", USD)
    for c in YC:
        put(w, f"{c}{q('cf')}", f"=IF({c}{q('n')}=$C{q('x')},{c}{q('proceeds')},0)", USD)
    put(w, f"C{q('irr_fn')}", f"=IRR(C{q('cf')}:H{q('cf')},0.1)", PCT)
    ex = lambda sheet, k: f"INDEX('{sheet}'!$D{R[sheet][k]}:$H{R[sheet][k]},$C{q('x')})"
    m0 = f"({A('ltm_ebitda')}/{A('ltm_rev')})"
    m1 = f"({ex('Operating Model', 'ebitda')}/{ex('Operating Model', 'rev')})"
    put(w, f"C{q('b_rev')}", f"=({ex('Operating Model', 'rev')}-{A('ltm_rev')})*{m0}*{A('entry_mult')}", USD)
    put(w, f"C{q('b_mar')}", f"=({m1}-{m0})*{ex('Operating Model', 'rev')}*{A('entry_mult')}", USD)
    put(w, f"C{q('b_mult')}", f"=({A('exit_mult')}-{A('entry_mult')})*{ex('Operating Model', 'ebitda')}", USD)
    put(w, f"C{q('b_del')}", f"={D('net_debt', 'C')}-{ex('Debt Schedule', 'net_debt')}", USD)
    put(w, f"C{q('b_fee')}", f"=-({T('u_txn')}+{T('u_fin')})", USD)
    put(w, f"C{q('b_mip')}", f"=-INDEX({rng('mip')},C{q('x')})", USD)
    put(w, f"C{q('b_tot')}", f"=SUM(C{q('b_rev')}:C{q('b_mip')})", USD, True)
    put(w, f"C{q('b_chk')}", f"=ROUND(C{q('b_tot')}-(C{q('sel_proc')}-{T('s_eq')}),6)", NUM)
    for k in ("b_rev", "b_mar", "b_mult", "b_del", "b_fee", "b_mip", "b_tot"):
        put(w, f"D{q(k)}", f"=IF($C{q('b_tot')}=0,0,C{q(k)}/$C{q('b_tot')})", PCT)
    w[f"D{q('b_rev') - 1}"].value, w[f"D{q('b_rev') - 1}"].font = "% of gain", f(True)
    for col in ["A", "B"] + AC:
        w[f"{col}{q('proceeds')}"].border = Border(top=Side(style="thin"))

    # ======================= Credit Stats =======================
    sh = "Credit Stats"
    items = [
        (None, "LEVERAGE", None, "section"),
        ("tot_lev", "Total debt / EBITDA", MULT2, ""), ("net_lev", "Net debt / EBITDA", MULT2, "bold"), ("sr_lev", "Senior secured debt / EBITDA", MULT2, ""),
        (None, None, None, None), (None, "COVERAGE", None, "section"),
        ("int_cov", "EBITDA / cash interest", MULT2, ""), ("capex_cov", "(EBITDA - capex) / cash interest", MULT2, ""),
        ("fccr", "Fixed-charge coverage: (EBITDA - capex - taxes) / (interest + mandatory amortisation)", MULT2, ""),
        (None, None, None, None), (None, "DELEVERAGING", None, "section"),
        ("fcf_debt", "FCF / total debt", PCT, ""), ("paydown", "Cumulative debt paydown (% of debt at close)", PCT, ""),
        (None, None, None, None), (None, "COVENANT (maximum net leverage)", None, "section"),
        ("cov", "Covenant level", MULT2, ""), ("headroom", "Leverage headroom (% below covenant)", PCT, ""),
        ("cushion", "EBITDA cushion (% EBITDA can fall before breach)", PCT, ""), ("flag", "Covenant status", None, "bold"),
    ]
    L = layout(sh, items)
    w = ws[sh]
    header(w, "Credit Statistics & Covenant Headroom")
    label_rows(w, L)
    S = sh
    k_ = lambda k, S="Credit Stats": R[S][k]
    G = lambda k, c: f"IFERROR({k},0)"
    for c in AC:
        put(w, f"{c}{k_('tot_lev')}", f"=IFERROR({D('tot_debt', c)}/{O('ebitda', c)},0)", MULT2)
        put(w, f"{c}{k_('net_lev')}", f"=IFERROR({D('net_debt', c)}/{O('ebitda', c)},0)", MULT2, True)
        put(w, f"{c}{k_('sr_lev')}", f"=IFERROR(({D('rev_end', c)}+{D('tlb_end', c)})/{O('ebitda', c)},0)", MULT2)
    fill(w, S, "int_cov", None, lambda c, pv: f"=IFERROR({O('ebitda', c)}/{D('total_int', c)},0)", MULT2, unit="x")
    fill(w, S, "capex_cov", None, lambda c, pv: f"=IFERROR(({O('ebitda', c)}-{O('capex', c)})/{D('total_int', c)},0)", MULT2, unit="x")
    fill(w, S, "fccr", None, lambda c, pv: f"=IFERROR(({O('ebitda', c)}-{O('capex', c)}-{O('taxes', c)})/({D('total_int', c)}+{D('mand', c)}),0)", MULT2, unit="x")
    fill(w, S, "fcf_debt", None, lambda c, pv: f"=IFERROR({O('fcf', c)}/{D('tot_debt', c)},0)", PCT, unit="%")
    fill(w, S, "paydown", None, lambda c, pv: f"=1-{D('tot_debt', c)}/{D('tot_debt', 'C')}", PCT, unit="%")
    fill(w, S, "cov", None, lambda c, pv: f"={ref('Assumptions', 'cov_max', c)}", MULT2, unit="x")
    fill(w, S, "headroom", None, lambda c, pv: f"=1-{c}{k_('net_lev')}/{c}{k_('cov')}", PCT, unit="%")
    fill(w, S, "cushion", None, lambda c, pv: f"=1-({D('net_debt', c)}/{c}{k_('cov')})/{O('ebitda', c)}", PCT, unit="%")
    fill(w, S, "flag", None, lambda c, pv: f'=IF({c}{k_("net_lev")}>{c}{k_("cov")},"BREACH","OK")', None, True, unit="")
    for c in YC:
        w[f"{c}{k_('flag')}"].alignment = Alignment(horizontal="right")
    w.conditional_formatting.add(f"D{k_('flag')}:H{k_('flag')}", FormulaRule(formula=[f'D{k_("flag")}="BREACH"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", bgColor="C00000")))

    # ======================= Sensitivity =======================
    sh = "Sensitivity"
    w = ws[sh]
    w.sheet_view.showGridLines = False
    w["A1"].value, w["A1"].font = "Sensitivity Analysis", f(True, NAVY, 14)
    w["A2"].value = "Tables 1-3 are LIVE closed-form formulas (debt schedule is independent of entry/exit multiple). Tables 4-5 are STATIC snapshots from the Python engine."
    w["A2"].font = f(italic=True, color="595959", size=9)
    w.column_dimensions["A"].width = 4
    w.column_dimensions["B"].width = 36
    for col in "CDEFGHI":
        w.column_dimensions[col].width = 12
    helpers = [("Exit year (N)", f"={A('exit_year')}", INT), ("EBITDA at exit", f"={ex('Operating Model', 'ebitda').replace('$C' + str(q('x')), A('exit_year'))}", USD),
               ("Net debt at exit", f"={ex('Debt Schedule', 'net_debt').replace('$C' + str(q('x')), A('exit_year'))}", USD),
               ("LTM EBITDA", f"={A('ltm_ebitda')}", USD), ("Funded debt at close", f"=SUM({T('s_rev')}:{T('s_notes')})", USD),
               ("Financing fees", f"={T('u_fin')}", USD), ("Minimum cash funded", f"={T('u_cash')}", USD),
               ("Transaction fee %", f"={A('txn_fee_pct')}", PCT), ("MIP %", f"={A('mip_pct')}", PCT),
               ("Base entry multiple", f"={A('entry_mult')}", MULT), ("Base exit multiple", f"={A('exit_mult')}", MULT),
               ("Sponsor equity at close", f"={T('s_eq')}", USD)]
    w["B4"].value, w["B4"].font = "Live helper cells", f(True)
    H = {}
    for i, (lab, fx, fm) in enumerate(helpers):
        r = 5 + i
        w[f"B{r}"].value, w[f"B{r}"].font = lab, f()
        put(w, f"C{r}", fx, fm)
        H[lab] = f"$C${r}"
    w["B17"].value, w["B17"].font = "Entry multiple step (x)", f()
    put(w, "C17", 0.5, MULT, input_=True)
    w["B18"].value, w["B18"].font = "Exit multiple step (x)", f()
    put(w, "C18", 1.0, MULT, input_=True)
    Nn, EB, ND, LT, FD, FF, MC, TX, MP, BE, BX, EQ = [H[k] for k in ("Exit year (N)", "EBITDA at exit", "Net debt at exit", "LTM EBITDA",
                                                                    "Funded debt at close", "Financing fees", "Minimum cash funded",
                                                                    "Transaction fee %", "MIP %", "Base entry multiple", "Base exit multiple",
                                                                    "Sponsor equity at close")]

    def table(top, title, row_vals, col_vals, cell_fn, row_fmt, col_fmt, cell_fmt, corner):
        w[f"B{top}"].value, w[f"B{top}"].font = title, f(True, NAVY, 11)
        w[f"B{top + 1}"].value, w[f"B{top + 1}"].font = corner, f(True, size=9)
        for j, cv in enumerate(col_vals):
            put(w, f"{get_column_letter(3 + j)}{top + 1}", cv, col_fmt, True, align="center")
            w[f"{get_column_letter(3 + j)}{top + 1}"].fill = PatternFill("solid", fgColor=GREY)
        for i, rv in enumerate(row_vals):
            r = top + 2 + i
            put(w, f"B{r}", rv, row_fmt, True, align="right")
            w[f"B{r}"].fill = PatternFill("solid", fgColor=GREY)
            for j in range(len(col_vals)):
                col = get_column_letter(3 + j)
                put(w, f"{col}{r}", cell_fn(f"$B{r}", f"{col}${top + 1}"), cell_fmt)
        last = f"{get_column_letter(2 + len(col_vals))}{top + 1 + len(row_vals)}"
        w.conditional_formatting.add(f"C{top + 2}:{last}", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
        return top + 2 + len(row_vals) + 2

    def eq_in(entry):
        return f"({entry}*{LT}*(1+{TX})+{FF}+{MC}-{FD})"

    def proceeds(exit_m, ebitda, nd, eqin):
        ee = f"({exit_m}*{ebitda}-{nd})"
        return f"({ee}-{MP}*MAX(0,{ee}-{eqin}))"

    ent_rows = [f"={BE}+{k}*$C$17" for k in (-3, -2, -1, 0, 1, 2, 3)]
    exit_cols = [f"={BX}+{k}*$C$18" for k in (-2, -1, 0, 1, 2)]

    def irr_cell(r, c):
        e = eq_in(r)
        pr = proceeds(c, EB, ND, e)
        return f"=IF(AND({pr}>0,{e}>0),({pr}/{e})^(1/{Nn})-1,-1)"

    def moic_cell(r, c):
        e = eq_in(r)
        return f"=IFERROR({proceeds(c, EB, ND, e)}/{e},0)"

    nxt = table(21, "Table 1: Sponsor IRR, entry multiple (rows) vs exit multiple (columns), selected exit year", ent_rows, exit_cols, irr_cell, MULT, MULT, PCT, "Entry \\ Exit")
    nxt = table(nxt, "Table 2: Sponsor MOIC, entry multiple vs exit multiple", ent_rows, exit_cols, moic_cell, MULT, MULT, MULT2, "Entry \\ Exit")

    def yr_cell(r, c):
        eb = f"INDEX('Operating Model'!$D${R['Operating Model']['ebitda']}:$H${R['Operating Model']['ebitda']},{r})"
        nd = f"INDEX('Debt Schedule'!$D${R['Debt Schedule']['net_debt']}:$H${R['Debt Schedule']['net_debt']},{r})"
        pr = proceeds(c, eb, nd, EQ)
        return f"=IF({pr}>0,({pr}/{EQ})^(1/{r})-1,-1)"

    nxt = table(nxt, "Table 3: Sponsor IRR, exit year (rows) vs exit multiple (columns), base entry", [3, 4, 5], exit_cols, yr_cell, INT, MULT, PCT, "Year \\ Exit")

    # static tables from engine
    lev_vals, ex_vals = [3.0, 3.5, 4.0, 4.5, 5.0], [p.exit_mult + k for k in (-2, -1, 0, 1, 2)]
    w[f"B{nxt}"].value, w[f"B{nxt}"].font = f"Table 4 (STATIC, {SCENARIOS[p.scenario]} case): IRR, Term Loan B leverage (rows, x EBITDA) vs exit multiple", f(True, NAVY, 11)
    w[f"B{nxt + 1}"].value, w[f"B{nxt + 1}"].font = "TLB x \\ Exit", f(True, size=9)
    for j, xv in enumerate(ex_vals):
        put(w, f"{get_column_letter(3 + j)}{nxt + 1}", xv, MULT, True, align="center")
    for i, lv in enumerate(lev_vals):
        put(w, f"B{nxt + 2 + i}", lv, MULT, True, align="right")
        for j, xv in enumerate(ex_vals):
            put(w, f"{get_column_letter(3 + j)}{nxt + 2 + i}", float(run(p.with_(tlb_x=lv, exit_mult=xv))["sel_irr"]), PCT)
    w.conditional_formatting.add(f"C{nxt + 2}:G{nxt + 6}", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    nxt += 9
    g_d, m_d = [-0.03, -0.015, 0.0, 0.015, 0.03], [-0.03, -0.015, 0.0, 0.015, 0.03]
    w[f"B{nxt}"].value, w[f"B{nxt}"].font = f"Table 5 (STATIC, {SCENARIOS[p.scenario]} case): IRR, EBITDA-margin shift (rows) vs revenue-growth shift (columns), every year", f(True, NAVY, 11)
    w[f"B{nxt + 1}"].value, w[f"B{nxt + 1}"].font = "Margin Δ \\ Growth Δ", f(True, size=9)
    base_g, base_m = p.growth[p.scenario], p.margin[p.scenario]
    for j, gd in enumerate(g_d):
        put(w, f"{get_column_letter(3 + j)}{nxt + 1}", gd, '+0.0%;-0.0%;0.0%', True, align="center")
    for i, md in enumerate(m_d):
        put(w, f"B{nxt + 2 + i}", md, '+0.0%;-0.0%;0.0%', True, align="right")
        for j, gd in enumerate(g_d):
            r_ = run(p, growth=[x + gd for x in base_g], margin=[x + md for x in base_m])
            put(w, f"{get_column_letter(3 + j)}{nxt + 2 + i}", float(r_["sel_irr"]), PCT)
    w.conditional_formatting.add(f"C{nxt + 2}:G{nxt + 6}", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    R["Sensitivity"]["t1_top"] = 21

    # ======================= Scenarios =======================
    sh = "Scenarios"
    w = ws[sh]
    w.sheet_view.showGridLines = False
    w["A1"].value, w["A1"].font = "Scenario Comparison", f(True, NAVY, 14)
    w["A2"].value = "Base / Upside / Downside columns are STATIC snapshots from the Python engine (all other inputs as in Assumptions). The LIVE column follows the scenario switch."
    w["A2"].font = f(italic=True, color="595959", size=9)
    w.column_dimensions["A"].width = 52
    for col in "BCDE":
        w.column_dimensions[col].width = 15
    for col, lab in zip("ABCDE", ["Metric (exit year per Assumptions)", "Base", "Upside", "Downside", "LIVE (selected)"]):
        c = w[f"{col}4"]; c.value, c.font, c.fill = lab, f(True, WHITE), PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center" if col != "A" else "left")
    runs = {s: run(p.with_(scenario=s)) for s in (1, 2, 3)}
    X = p.exit_year
    n_rng = f"'Returns'!$D${R['Returns']['n']}:$H${R['Returns']['n']}"
    metrics = [
        ("Revenue at exit", USD, lambda r: r["rev"][X], f"={ex('Operating Model', 'rev').replace('$C' + str(q('x')), A('exit_year'))}"),
        ("EBITDA at exit", USD, lambda r: r["ebitda"][X], f"={ex('Operating Model', 'ebitda').replace('$C' + str(q('x')), A('exit_year'))}"),
        ("EBITDA margin at exit", PCT, lambda r: r["ebitda"][X] / r["rev"][X], None),
        ("Cumulative free cash flow to exit", USD, lambda r: r["fcf"][1:X + 1].sum(), f"=SUMIF({n_rng},\"<=\"&{A('exit_year')},'Operating Model'!$D${R['Operating Model']['fcf']}:$H${R['Operating Model']['fcf']})"),
        ("Net debt / EBITDA at exit", MULT2, lambda r: r["net_lev"][X], f"=INDEX('Credit Stats'!$D${R['Credit Stats']['net_lev']}:$H${R['Credit Stats']['net_lev']},{A('exit_year')})"),
        ("Minimum interest coverage (5-yr)", MULT2, lambda r: r["int_cov"][1:].min(), f"=MIN('Credit Stats'!$D${R['Credit Stats']['int_cov']}:$H${R['Credit Stats']['int_cov']})"),
        ("Peak revolver draw (5-yr)", USD, lambda r: r["rev_bal"].max(), f"=MAX('Debt Schedule'!$C${R['Debt Schedule']['rev_end']}:$H${R['Debt Schedule']['rev_end']})"),
        ("Covenant breaches (count, 5-yr)", INT, lambda r: int(r["breach"].sum()), f"=COUNTIF('Credit Stats'!$D${R['Credit Stats']['flag']}:$H${R['Credit Stats']['flag']},\"BREACH\")"),
        ("Sponsor equity value at exit (100%)", USD, lambda r: r["eq_x"][X], f"=INDEX('Returns'!$D${R['Returns']['eq_x']}:$H${R['Returns']['eq_x']},{A('exit_year')})"),
        ("Sponsor MOIC", MULT2, lambda r: r["moic"][X], f"={ref('Returns', 'sel_moic', 'C')}"),
        ("Sponsor IRR", PCT, lambda r: r["irr"][X], f"={ref('Returns', 'sel_irr', 'C')}"),
    ]
    for i, (lab, fm, fn, live) in enumerate(metrics):
        r = 5 + i
        w[f"A{r}"].value, w[f"A{r}"].font = lab, f(bold=lab.startswith("Sponsor I") or lab.startswith("Sponsor M"))
        for col, s in zip("BCD", (1, 2, 3)):
            put(w, f"{col}{r}", float(fn(runs[s])), fm)
        if live:
            put(w, f"E{r}", live, fm)
        elif lab.startswith("EBITDA margin"):
            put(w, f"E{r}", f"=E{r - 1}/E{r - 2}", fm)
    R["Scenarios"]["irr_row"] = 5 + len(metrics) - 1

    # ======================= Checks =======================
    sh = "Checks"
    w = ws[sh]
    w.sheet_view.showGridLines = False
    w["A1"].value, w["A1"].font = "Model Integrity Checks", f(True, NAVY, 14)
    w.column_dimensions["A"].width = 64
    w.column_dimensions["B"].width = 16
    w.column_dimensions["C"].width = 12
    for col, lab in zip("ABC", ["Check", "Value", "Status"]):
        c = w[f"{col}4"]; c.value, c.font, c.fill = lab, f(True, WHITE), PatternFill("solid", fgColor=NAVY)
    BSr, DS = R["Balance Sheet"], R["Debt Schedule"]
    checks = [
        ("Sources equal uses", f"=ROUND({T('s_tot')}-{T('u_tot')},6)", "abs"),
        ("Balance sheet balances in every period (max abs difference)", f"=MAX(ABS(MIN('Balance Sheet'!C{BSr['check']}:H{BSr['check']})),ABS(MAX('Balance Sheet'!C{BSr['check']}:H{BSr['check']})))", "abs"),
        ("Cash never below minimum cash (shortfall, 0 = fine)", f"=MAX(0,{A('min_cash')}-MIN('Debt Schedule'!D{DS['end_cash']}:H{DS['end_cash']}))", "abs"),
        ("Revolver within commitment (excess draw, 0 = fine)", f"=MAX(0,MAX('Debt Schedule'!C{DS['rev_end']}:H{DS['rev_end']})-{A('rev_commit')})", "abs"),
        ("No negative debt balances (most negative, 0 = fine)", f"=MAX(0,-MIN('Debt Schedule'!C{DS['tlb_end']}:H{DS['tlb_end']},'Debt Schedule'!C{DS['rev_end']}:H{DS['rev_end']}))", "abs"),
        ("Closed-form IRR equals Excel IRR() function", f"=ROUND({ref('Returns', 'sel_irr', 'C')}-{ref('Returns', 'irr_fn', 'C')},6)", "abs"),
        ("Value-creation bridge reconciles to sponsor gain", f"={ref('Returns', 'b_chk', 'C')}", "abs"),
        ("Exit year is an integer between 1 and 5", f"=IF(AND({A('exit_year')}>=1,{A('exit_year')}<=5,{A('exit_year')}=INT({A('exit_year')})),0,1)", "abs"),
        ("Scenario selector is 1, 2 or 3", f"=IF(OR({A('scenario')}=1,{A('scenario')}=2,{A('scenario')}=3),0,1)", "abs"),
    ]
    for i, (lab, fx, _) in enumerate(checks):
        r = 5 + i
        w[f"A{r}"].value, w[f"A{r}"].font = lab, f()
        put(w, f"B{r}", fx, '0.000000;-0.000000;0.000000')
        put(w, f"C{r}", f'=IF(ABS(B{r})<0.0001,"OK","ERROR")', None, True, align="center")
    last = 4 + len(checks)
    mr = last + 2
    w[f"A{mr}"].value, w[f"A{mr}"].font = "MASTER CHECK", f(True, size=12)
    put(w, f"C{mr}", f'=IF(COUNTIF(C5:C{last},"OK")={len(checks)},"ALL OK","CHECK MODEL")', None, True, align="center")
    R["Checks"]["master"] = mr
    w.conditional_formatting.add(f"C5:C{mr}", FormulaRule(formula=['OR(C5="ERROR",C5="CHECK MODEL")'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", bgColor="C00000")))
    w.conditional_formatting.add(f"C5:C{mr}", FormulaRule(formula=['OR(C5="OK",C5="ALL OK")'], font=Font(name=FONT, bold=True, color="006100"), fill=PatternFill("solid", bgColor="C6EFCE")))

    # ======================= Summary =======================
    sh = "Summary"
    w = ws[sh]
    w.sheet_view.showGridLines = False
    w["A1"].value, w["A1"].font = "Project Atlas: Leveraged Buyout Model", f(True, NAVY, 16)
    w["A2"].value = "($ in millions) | FICTIONAL company with illustrative assumptions | educational model, not investment advice"
    w["A2"].font = f(italic=True, color="595959", size=9)
    w.column_dimensions["A"].width = 38
    for col in "BCDEFGH":
        w.column_dimensions[col].width = 13

    def block(r, title):
        for col in range(1, 9):
            w.cell(r, col).fill = PatternFill("solid", fgColor=NAVY)
        w[f"A{r}"].value, w[f"A{r}"].font = title, f(True, WHITE)

    block(4, "TRANSACTION")
    rows = [("Scenario (1 Base / 2 Up / 3 Down)", f"={A('scenario')}", INT), ("Entry EV", f"={T('ev')}", USD),
            ("Entry EV / LTM EBITDA", f"={T('mult')}", MULT), ("Entry EV / LTM revenue", f"={T('ev_rev')}", MULT2),
            ("Total leverage at close (x EBITDA)", f"=({T('s_rev')}+{T('s_tlb')}+{T('s_notes')})/{T('ltm_e')}", MULT),
            ("Sponsor equity", f"={T('s_eq')}", USD), ("Equity % of capitalisation", f"={T('eq_pct')}", PCT)]
    for i, (lab, fx, fm) in enumerate(rows):
        w[f"A{5 + i}"].value, w[f"A{5 + i}"].font = lab, f()
        put(w, f"B{5 + i}", fx, fm)
    block(4, "TRANSACTION")
    for col in range(4, 9):
        w.cell(4, col).fill = PatternFill("solid", fgColor=NAVY)
    w["D4"].value, w["D4"].font = "RETURNS (selected exit year)", f(True, WHITE)
    ret = [("Exit year", f"={A('exit_year')}", INT), ("Exit multiple", f"={A('exit_mult')}", MULT),
           ("Sponsor IRR", f"={ref('Returns', 'sel_irr', 'C')}", PCT), ("Sponsor MOIC", f"={ref('Returns', 'sel_moic', 'C')}", MULT2),
           ("Sponsor proceeds", f"={ref('Returns', 'sel_proc', 'C')}", USD),
           ("Model integrity", f"='Checks'!C{R['Checks']['master']}", None)]
    for i, (lab, fx, fm) in enumerate(ret):
        w[f"D{5 + i}"].value, w[f"D{5 + i}"].font = lab, f(bold=lab in ("Sponsor IRR", "Sponsor MOIC"))
        w[f"D{5 + i}"].alignment = Alignment(indent=2)
        put(w, f"G{5 + i}", fx, fm, bold=lab in ("Sponsor IRR", "Sponsor MOIC"), align="right")

    block(13, "SOURCES & USES")
    su_s = [("Revolver", "s_rev"), ("Term Loan B", "s_tlb"), ("Senior notes", "s_notes"), ("Sponsor equity", "s_eq"), ("Total sources", "s_tot")]
    su_u = [("Purchase of enterprise", "u_ev"), ("Transaction fees", "u_txn"), ("Financing fees", "u_fin"), ("Minimum cash", "u_cash"), ("Total uses", "u_tot")]
    for i, ((ls, ks), (lu, ku)) in enumerate(zip(su_s, su_u)):
        r = 14 + i
        bold = ks.endswith("tot")
        w[f"A{r}"].value, w[f"A{r}"].font = ls, f(bold)
        put(w, f"B{r}", f"={T(ks)}", USD, bold)
        put(w, f"C{r}", f"={T(ks, 'E')}", MULT, bold)
        w[f"D{r}"].value, w[f"D{r}"].font = lu, f(bold)
        w[f"D{r}"].alignment = Alignment(indent=2)
        put(w, f"G{r}", f"={T(ku)}", USD, bold)
        put(w, f"H{r}", f"={T(ku, 'E')}", MULT, bold)
    block(20, "FINANCIAL SUMMARY")
    for col, lab in zip("ABCDEFGH", ["", "Close", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", ""]):
        c = w[f"{col}21"]; c.value, c.font = lab, f(True)
        c.alignment = Alignment(horizontal="right" if col != "A" else "left")
        c.border = Border(bottom=Side(style="thin"))
    summ = [("Revenue", "Operating Model", "rev", USD, AC), ("  % growth", "Operating Model", "growth", PCT, YC),
            ("EBITDA", "Operating Model", "ebitda", USD, AC), ("  % margin", "Operating Model", "margin", PCT, AC),
            ("Free cash flow (after debt service)", "Operating Model", "fcf", USD, YC), ("Net debt", "Debt Schedule", "net_debt", USD, AC),
            ("Net debt / EBITDA", "Credit Stats", "net_lev", MULT2, AC), ("EBITDA / cash interest", "Credit Stats", "int_cov", MULT2, YC),
            ("Covenant status", "Credit Stats", "flag", None, YC)]
    for i, (lab, sheet, key, fm, cols) in enumerate(summ):
        r = 22 + i
        w[f"A{r}"].value, w[f"A{r}"].font = lab, f(bold=not lab.startswith("  "))
        for col in cols:
            tgt = "BCDEFG"[AC.index(col)]
            put(w, f"{tgt}{r}", f"={ref(sheet, key, col)}", fm, align="right")
    R["Summary"].update(rev=22, ebitda=24, fcf=26, netlev=28)

    # charts (anchored below the tables)
    cs = ws["Credit Stats"]
    ch = LineChart()
    ch.title, ch.height, ch.width = "Net debt / EBITDA vs covenant", 7.5, 13
    ch.add_data(Reference(cs, min_col=3, max_col=8, min_row=R["Credit Stats"]["net_lev"]), from_rows=True, titles_from_data=False)
    ch.add_data(Reference(cs, min_col=3, max_col=8, min_row=R["Credit Stats"]["cov"]), from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(cs, min_col=3, max_col=8, min_row=4))
    ch.series[0].tx = None
    from openpyxl.chart.series import SeriesLabel
    ch.series[0].tx, ch.series[1].tx = SeriesLabel(v="Net debt / EBITDA"), SeriesLabel(v="Covenant max")
    w.add_chart(ch, "A33")
    bc = BarChart()
    bc.title, bc.height, bc.width = "EBITDA and free cash flow ($mm)", 7.5, 13
    om = ws["Operating Model"]
    bc.add_data(Reference(om, min_col=4, max_col=8, min_row=R["Operating Model"]["ebitda"]), from_rows=True, titles_from_data=False)
    bc.add_data(Reference(om, min_col=4, max_col=8, min_row=R["Operating Model"]["fcf"]), from_rows=True, titles_from_data=False)
    bc.set_categories(Reference(om, min_col=4, max_col=8, min_row=4))
    bc.series[0].tx, bc.series[1].tx = SeriesLabel(v="EBITDA"), SeriesLabel(v="Free cash flow")
    w.add_chart(bc, "E33")
    rt = ws["Returns"]
    vb = BarChart()
    vb.title, vb.height, vb.width = "Value creation bridge ($mm)", 7.5, 13
    vb.add_data(Reference(rt, min_col=3, min_row=R["Returns"]["b_rev"], max_row=R["Returns"]["b_mip"]), titles_from_data=False)
    vb.set_categories(Reference(rt, min_col=1, min_row=R["Returns"]["b_rev"], max_row=R["Returns"]["b_mip"]))
    vb.legend = None
    w.add_chart(vb, "A49")

    # sheet tab colours & default calc
    for s, colr in zip(SHEETS, ["1F3864", "FFC000", "2E75B6", "2E75B6", "2E75B6", "2E75B6", "70AD47", "70AD47", "ED7D31", "ED7D31", "C00000"]):
        ws[s].sheet_properties.tabColor = colr
    for sheet in wb.worksheets:   # print-friendly: landscape, one page wide
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    print(build())
