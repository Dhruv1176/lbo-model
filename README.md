# LBO Model: Excel + Python, Investment Memo, Monte Carlo

![CI](https://github.com/YOUR_USERNAME/lbo-model/actions/workflows/ci.yml/badge.svg)

A complete leveraged-buyout analysis of a **fictional** company ("Project Atlas"), delivered the way a PE associate would:

1. **`models/Project_Atlas_LBO.xlsx`**: an 11-sheet, ~800-formula Excel model (live formulas, scenario switch, sensitivity tables, integrity checks)
2. **A Python engine** that mirrors the workbook cell-for-cell, with tests that **change inputs inside Excel, recalculate with LibreOffice, and reconcile every key line to the engine within 1e-6** (currently ~1e-13)
3. **`docs/INVESTMENT_MEMO.md`**: an investment-committee memo whose numbers are generated from the engine, so they cannot disagree with the model
4. Risk analytics: IRR tornado, Monte Carlo, ability-to-pay, break-evens

> Project Atlas is invented and all inputs are illustrative. This is an educational model, not investment advice.

## The deal in one table
| | |
|---|---|
| Entry | 11.0x LTM EBITDA of $100mm = $1,100mm EV |
| Capital structure | $75mm revolver (undrawn), $400mm Term Loan B (4.0x), $150mm senior notes (1.5x), $608mm sponsor equity (52.5%) |
| Base case | **15.9% IRR / 2.09x MOIC** (Year-5 exit at 11.0x, no multiple expansion, 8% MIP) |
| Downside / Upside | 1.7% / 21.6% IRR (1.09x / 2.66x MOIC) |
| Monte Carlo (10k paths) | median IRR 15.6%, P5 to P95 = 6.6% to 23.8%, P(loss) 0.4% |
| Key finding | Base return is below a 20% hurdle at 11.0x; **ability-to-pay for 20% is ~10.0x** with the same structure |

## Workbook map
| Sheet | What it does |
|---|---|
| **Summary** | Dashboard: transaction, S&U, returns, projections, covenant status, 3 charts |
| **Assumptions** | Every input (blue), a source/rationale note for each, scenario levers (Base/Upside/Downside) and a live scenario switch |
| **Transaction** | Purchase price, sources & uses, goodwill, fee amortisation |
| **Operating Model** | Income statement, tax calculation with 163(j)-style interest cap, cash flow, unlevered FCF |
| **Debt Schedule** | Cash waterfall, auto-drawing revolver, mandatory amortisation, % cash sweep, bullet notes, interest on beginning balances |
| **Balance Sheet** | Full pro-forma balance sheet with a zero-tolerance balance check each year |
| **Returns** | MOIC/IRR by exit year, IRR() cross-check, **value-creation bridge** (revenue, margin, multiple, deleveraging, fees, MIP) |
| **Credit Stats** | Leverage, coverage, FCCR, paydown, covenant headroom and EBITDA cushion, breach flags |
| **Sensitivity** | Live IRR/MOIC grids (entry × exit multiple, exit year × exit multiple); static grids (leverage × exit multiple; margin × growth) |
| **Scenarios** | Side-by-side Base / Upside / Downside plus a LIVE column |
| **Checks** | 9 integrity checks and a master flag |

Colour code: blue = input, black = formula, green = cross-sheet link, yellow fill = key deal lever.

## Run it
```bash
pip install -r requirements.txt
python -m lbo build            # rebuilds the workbook, charts and memo (reconciles if LibreOffice is installed)
python -m lbo analyze          # prints tornado, break-evens, Monte Carlo
python -m pytest -q            # 13 tests, including the Excel-vs-engine reconciliation
```
Open `models/Project_Atlas_LBO.xlsx` and try: set **Assumptions → scenario** to 3, change **exit year** to 4, push **entry multiple** to 12x. Everything recalculates, and the Checks sheet stays green.

## Modelling choices (and why)
- **Interest on beginning balances.** No circular references, no iterative calc. Slightly conservative in a deleveraging deal.
- **Cash sweep and revolver.** Shortfalls draw the revolver automatically (capped at the commitment); excess cash repays the revolver first, then sweeps a % to the TLB.
- **163(j) interest cap** (30% of EBITDA, toggleable) with no carry-forward: realistic for post-2022 US deals and a common thing candidates forget.
- **Returns attribution** that reconciles exactly to sponsor gain, so you can say *where* returns come from.
- **Static vs live.** Tables that need a full re-run of the debt schedule are labelled **STATIC** (Python snapshot). Everything else recalculates.

## Limitations
No purchase-price allocation or deferred taxes, no PIK / dividend recap / add-on acquisitions, one operating segment, stylised Monte Carlo shocks, illustrative assumptions throughout.

## Ideas to extend (good interview talking points)
Average-balance interest with an iterative-calc toggle · PIK toggle notes · dividend recap and exit-timing optimisation · add-on M&A · management rollover and sweet equity waterfall · DCF cross-check using the Operating Model's unlevered FCF · quarterly model

```
lbo/  engine.py build_excel.py reconcile.py recalc.py analysis.py memo.py cli.py
models/ Project_Atlas_LBO.xlsx     docs/ memo + charts     tests/ test_lbo.py
```
MIT License.

## Author

**Dhruv Singh**

GitHub: [@Dhruv1176](https://github.com/Dhruv1176)
