"""
Edge-case tests run through Microsoft Excel itself (not just the Python engine).
For each case the 'Custom' scenario column of a COPY of the workbook is changed, Excel recalculates it
(macOS AppleScript), and the Calc_Custom results are compared with the Python engine.
    python test_excel_edge_cases.py <scratch_dir>
"""
import os, sys, shutil, subprocess
from dataclasses import replace
import openpyxl
sys.path.insert(0, os.path.dirname(__file__))
import cashflow_engine as E
from test_validation import compare

SRC = os.path.join(os.path.dirname(__file__), "..", "Team1_Cash_Flow_Forecaster_v2.xlsx")
ROW = {"cycle": 7, "lag": 8, "ret": 9, "dlp": 10, "tds": 11, "adv": 12, "rate": 14, "cost": 16, "lM": 22, "open": 27}
CASES = [
    ("E1 No mobilisation advance", {"adv": 0}, dict(advance=0.0)),
    ("E2 Zero payment lag", {"lag": 0}, dict(payment_lag=0)),
    ("E3 Bi-monthly billing", {"cycle": 2}, dict(billing_cycle=2)),
    ("E4 Recovery rate 90% (cap test)", {"rate": 0.9}, dict(recovery_rate=0.9)),
    ("E5 No retention, no TDS", {"ret": 0, "tds": 0}, dict(retention=0.0, tds=0.0)),
    ("E6 Own funds INR 1 crore", {"open": 10_000_000}, dict(opening_cash=10_000_000)),
    ("E7 Lag 4 + 2-month billing + materials paid in cash", {"lag": 4, "cycle": 2, "lM": 0},
     dict(payment_lag=4, billing_cycle=2, cost_lag={"Labour": 0, "Materials": 0, "Plant and equipment": 0, "Site overheads": 0})),
]

def recalc(path):
    script = f'''tell application "Microsoft Excel"
set display alerts to false
open POSIX file "{path}"
set wb to active workbook
calculate full
save wb
close wb saving no
end tell'''
    subprocess.run(["osascript", "-e", script], check=True)

def main(tmp):
    ok_all = True
    for name, xl, py in CASES:
        p = os.path.join(tmp, name.split()[0] + ".xlsx"); shutil.copy(SRC, p)
        wb = openpyxl.load_workbook(p)
        for k, v in xl.items(): wb["Assumptions"][f"F{ROW[k]}"] = v
        wb.save(p); recalc(p)
        ws = openpyxl.load_workbook(p, data_only=True)
        res = E.run(replace(E.Inputs(), **py))
        n, w = compare(ws["Calc_Custom"], res, name)
        chk = ws["Checks"]["C4"].value
        s = res["summary"]
        ok = w < 0.5
        ok_all &= ok
        print(f"{'PASS' if ok else 'FAIL'} {name}: {n} values, max diff {w:.4f}; peak {s['peak_funding']:,.0f}, "
              f"interest {s['total_interest']:,.0f}, advance recovered {s['advance_recovered']:,.0f}; Checks sheet: {chk}")
    return ok_all

if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1]) else 1)
