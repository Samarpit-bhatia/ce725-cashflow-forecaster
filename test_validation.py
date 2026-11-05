"""
Cross-check: Excel tool (values as calculated by Microsoft Excel) vs independent Python engine.
Also runs edge-case tests on the engine.
Run after the workbook has been opened/recalculated and saved in Excel:
    python test_validation.py
"""
import os, sys
from dataclasses import replace
import openpyxl
sys.path.insert(0, os.path.dirname(__file__))
import cashflow_engine as E

XL = os.path.join(os.path.dirname(__file__), "..", "Team1_Cash_Flow_Forecaster_v2.xlsx")
ROWMAP = {"gross_bill": 9, "retention": 10, "advance_recovery": 11, "tds": 12, "net_bill": 13, "receipts": 14,
          "advance_in": 15, "retention_release": 16, "inflow": 17, "cost_incurred": 18, "outflow": 23,
          "opening": 24, "interest": 25, "closing": 27, "funding": 28}
SUMMAP = {"peak_funding": 37, "peak_month": 38, "total_interest": 39, "months_in_overdraft": 40,
          "months_over_limit": 41, "advance_recovered": 43, "retention_held": 44, "tds_deducted": 45,
          "total_cost": 46, "margin_after_interest": 48}

def compare(ws, res, label):
    worst, n = 0.0, 0
    for k, r in ROWMAP.items():
        for i in range(E.HORIZON):
            xv = ws.cell(row=r, column=3 + i).value or 0
            d = abs(xv - res[k][i]); worst = max(worst, d); n += 1
            if d > 0.5:
                print(f"  MISMATCH {label} {k} M{i+1}: excel {xv} python {res[k][i]}")
    for k, r in SUMMAP.items():
        xv = ws.cell(row=r, column=3).value or 0
        d = abs(xv - res["summary"][k]); worst = max(worst, d); n += 1
        if d > 0.5:
            print(f"  MISMATCH {label} summary {k}: excel {xv} python {res['summary'][k]}")
    return n, worst

def main():
    wb = openpyxl.load_workbook(XL, data_only=True)
    if wb["Calc_Base"]["C37"].value is None:
        sys.exit("Workbook has no calculated values - open and save it in Excel first.")
    total, allworst = 0, 0.0
    sheets = {"Base": "Calc_Base", "Payment delay": "Calc_Delay", "Stress": "Calc_Stress", "Custom": "Calc_Custom"}
    print("Excel vs Python engine")
    for name, inp in E.scenarios().items():
        n, w = compare(wb[sheets[name]], E.run(inp), name); total += n; allworst = max(allworst, w)
        s = E.run(inp)["summary"]
        print(f"  {name:14s} {n} values compared, max abs diff INR {w:.4f} | peak {s['peak_funding']:,.0f} M{s['peak_month']} interest {s['total_interest']:,.0f}")
    for k in range(5):
        inp = replace(E.Inputs(), payment_lag=k)
        n, w = compare(wb[f"Sens_Lag{k}"], E.run(inp), f"lag{k}"); total += n; allworst = max(allworst, w)
        s = E.run(inp)["summary"]
        print(f"  Sensitivity lag {k}: {n} values, max diff {w:.4f} | peak {s['peak_funding']:,.0f} interest {s['total_interest']:,.0f}")
    print(f"TOTAL {total} values compared, max abs difference INR {allworst:.4f}")
    print("Checks sheet overall:", wb["Checks"]["C4"].value, "| Validation sheet:", wb["Validation"]["C30"].value)

    print("\nEdge-case tests on engine")
    base = E.Inputs()
    tests = []
    r = E.run(replace(base, advance=0.0)); tests.append(("No advance -> no recovery", sum(r["advance_recovery"]) == 0))
    r = E.run(replace(base, payment_lag=0)); tests.append(("Zero lag -> receipt in bill month", r["receipts"][0] == r["net_bill"][0]))
    r = E.run(replace(base, retention=0.0)); tests.append(("No retention -> no release", sum(r["retention_release"]) == 0))
    r = E.run(replace(base, recovery_rate=0.9)); tests.append(("High recovery rate is capped at advance", abs(sum(r["advance_recovery"]) - 600000) < 1e-6))
    r = E.run(replace(base, billing_cycle=2)); tests.append(("2-month billing: bills only in M2, M4, M6", [i+1 for i,g in enumerate(r["gross_bill"]) if g>0] == [2,4,6]))
    pk = [E.run(replace(base, payment_lag=k))["summary"]["peak_funding"] for k in range(5)]
    tests.append(("Peak funding never falls as lag rises", all(pk[i] <= pk[i+1] for i in range(4))))
    r = E.run(replace(base, opening_cash=10_000_000)); tests.append(("Large own funds -> no interest", r["summary"]["total_interest"] == 0))
    r = E.run(base); s = r["summary"]
    tests.append(("Cash reconciles: end cash = CV - cost - TDS - interest", abs(s["closing_cash_end"] - (6e6 - s["total_cost"] - s["tds_deducted"] - s["total_interest"])) < 0.01))
    acts = [E.Activity("A", 1, 12, 1.0)]
    r = E.run(replace(base, activities=acts)); tests.append(("12-month single activity spreads 1/12 per month", all(abs(x - 1/12) < 1e-12 for x in r["work_share"][:12])))
    for t, ok in tests:
        print(f"  {'PASS' if ok else 'FAIL'}  {t}")
    return all(ok for _, ok in tests) and allworst < 0.5

if __name__ == "__main__":
    sys.exit(0 if main() else 1)
