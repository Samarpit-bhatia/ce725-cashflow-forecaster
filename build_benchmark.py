"""
Builds 03_Validation/Manual_Benchmark_Hand_Calculation.xlsx: a plain worksheet with TYPED numbers and the
arithmetic written out in words, so the Validation Engineer can redo every step with a calculator.
Nothing in it links to the tool. Also writes benchmark_values.py (values transcribed into the tool's
Validation sheet).

The arithmetic is laid out step by step in the order of the method note, using simple whole-rupee
numbers. Because it was prepared with AI assistance it MUST be independently recomputed by hand and
signed off by the Validation Engineer before submission (see Sign-off sheet).
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

OUT = os.path.join(os.path.dirname(__file__), "..", "..", "03_Validation", "Manual_Benchmark_Hand_Calculation.xlsx")
os.makedirs(os.path.dirname(OUT), exist_ok=True)

CV = 6_000_000
work = [0.10, 0.15, 0.20, 0.20, 0.15, 0.20]          # from the activity schedule (Step 1)
N = 13                                              # months shown (work M1-M6, last receipt M7, retention M12)
f = lambda x: f"{x:,.0f}"

# ---- hand calculation, written as plain sequential arithmetic ----
value = [CV * w for w in work] + [0] * (N - 6)
gross = value[:]                                    # monthly billing: bill = work done that month
ret = [g * 0.05 for g in gross]
tds = [g * 0.02 for g in gross]
rec, outst = [], 600_000
for m in range(N):
    r = 0 if m == 0 else min(gross[m] * 0.15, outst)
    rec.append(r); outst -= r
net = [gross[m] - ret[m] - rec[m] - tds[m] for m in range(N)]
receipt = [0] + net[:-1]                            # 1-month lag
adv = [600_000] + [0] * (N - 1)
rel = [0] * N; rel[11] = sum(ret)                   # completion M6 + DLP 6 = M12
cin = [receipt[m] + adv[m] + rel[m] for m in range(N)]
cost = [v * 0.85 for v in value]
lab = [c * 0.30 for c in cost]; plant = [c * 0.15 for c in cost]; ovh = [c * 0.15 for c in cost]
mat = [0] + [c * 0.40 for c in cost[:-1]]          # 1-month supplier credit
cout = [lab[m] + mat[m] + plant[m] + ovh[m] for m in range(N)]
opening, intr, closing = [], [], []
bal = 0
for m in range(N):
    opening.append(bal)
    i = round(max(0, -bal) * 0.12 / 12); intr.append(i)
    bal = bal + cin[m] - cout[m] - i; closing.append(bal)

wb = Workbook()
ws = wb.active; ws.title = "Inputs"
ws["A1"] = "Manual benchmark - Base case (Pune G+1, INR 60 lakh, 6 months)"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Plain worksheet: typed numbers only, no formulas, no links to the tool."
inputs = [("Contract value", "6,000,000"), ("Monthly work done (from schedule)", "10%, 15%, 20%, 20%, 15%, 20%"),
          ("Billing", "Monthly; bill = value of work done that month"), ("Payment lag", "1 month"),
          ("Retention", "5% of gross bill; released in month 12 (completion M6 + 6-month DLP)"),
          ("TDS", "2% of gross bill"), ("Mobilisation advance", "10% = 600,000 received in month 1"),
          ("Advance recovery", "15% of gross bill from month 2, until 600,000 is recovered"),
          ("Cost", "85% of value of work; labour 30%, materials 40%, plant 15%, overheads 15%"),
          ("Cost timing", "Materials paid one month later; labour, plant and overheads paid in the work month"),
          ("Overdraft", "12% p.a.; interest = opening negative balance x 1% per month, rounded to the rupee")]
for i, (a, b) in enumerate(inputs):
    ws.cell(row=4 + i, column=1, value=a).font = Font(bold=True); ws.cell(row=4 + i, column=2, value=b)
ws.column_dimensions["A"].width = 36; ws.column_dimensions["B"].width = 90

# Step-by-step worksheet
s = wb.create_sheet("Step_by_step")
s["A1"] = "Step-by-step hand calculation (each line can be checked with a calculator)"; s["A1"].font = Font(bold=True, size=13)
hdr = ["Month", "Step", "Working", "Result (INR)"]
for j, h in enumerate(hdr):
    c = s.cell(row=3, column=1 + j, value=h); c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864")
r = 4
for m in range(N):
    M = m + 1
    lines = []
    if value[m]:
        lines += [("Value of work", f"6,000,000 x {work[m]:.0%}", value[m]),
                  ("Retention", f"{f(gross[m])} x 5%", ret[m]),
                  ("Advance recovery", "nil in month 1" if M == 1 else
                   (f"{f(gross[m])} x 15% = {f(gross[m]*0.15)}" + ("" if rec[m] == gross[m]*0.15 else f", capped at balance outstanding {f(rec[m])}")), rec[m]),
                  ("TDS", f"{f(gross[m])} x 2%", tds[m]),
                  ("Net bill payable", f"{f(gross[m])} - {f(ret[m])} - {f(rec[m])} - {f(tds[m])}", net[m])]
    parts = []
    if receipt[m]: parts.append(f"bill of M{M-1} {f(receipt[m])}")
    if adv[m]: parts.append(f"advance {f(adv[m])}")
    if rel[m]: parts.append(f"retention release {f(rel[m])}")
    lines.append(("Cash in", " + ".join(parts) if parts else "nil", cin[m]))
    cparts = []
    if cost[m]: cparts.append(f"labour+plant+overheads {f(cost[m])} x 60% = {f(cost[m]*0.6)}")
    if mat[m]: cparts.append(f"materials of M{M-1} {f(cost[m-1])} x 40% = {f(mat[m])}")
    lines.append(("Cash out", "; ".join(cparts) if cparts else "nil", cout[m]))
    lines.append(("Overdraft interest", f"opening {f(opening[m])} -> " + ("no overdraft" if opening[m] >= 0 else f"{f(-opening[m])} x 1% = {-opening[m]*0.01:,.2f}, rounded"), intr[m]))
    lines.append(("Closing balance", f"{f(opening[m])} + {f(cin[m])} - {f(cout[m])} - {f(intr[m])}", closing[m]))
    for k, (st, wk, res) in enumerate(lines):
        s.cell(row=r, column=1, value=M if k == 0 else None)
        s.cell(row=r, column=2, value=st); s.cell(row=r, column=3, value=wk)
        c = s.cell(row=r, column=4, value=round(res, 2)); c.number_format = '#,##0;[Red]-#,##0'
        if st == "Closing balance":
            for j in range(4): s.cell(row=r, column=1 + j).font = Font(bold=True)
        r += 1
    r += 1
for col, w in zip("ABCD", [8, 22, 80, 16]): s.column_dimensions[col].width = w

# Summary table
t = wb.create_sheet("Monthly_table")
t["A1"] = "Benchmark results by month (typed values)"; t["A1"].font = Font(bold=True, size=13)
rows = [("Gross running bill", gross), ("Retention", ret), ("Advance recovery", rec), ("TDS", tds),
        ("Net amount payable", net), ("Client payment received", receipt), ("Total cash in", cin),
        ("Total cash out", cout), ("Overdraft interest", intr), ("Closing cash balance", closing)]
t.cell(row=3, column=1, value="Line item").font = Font(bold=True)
for m in range(N): t.cell(row=3, column=2 + m, value=f"M{m+1}").font = Font(bold=True)
for i, (lab_, vals) in enumerate(rows):
    t.cell(row=4 + i, column=1, value=lab_)
    for m in range(N):
        c = t.cell(row=4 + i, column=2 + m, value=round(vals[m], 2)); c.number_format = '#,##0;[Red]-#,##0'
peak = max(-c for c in closing)
res = [("Peak funding requirement", f(peak), f"month {[-c for c in closing].index(peak)+1}"),
       ("Total overdraft interest", f(sum(intr)), "sum of monthly interest"),
       ("Closing cash after retention release", f(closing[-1]), f"Check: 6,000,000 contract - 5,100,000 cost - 120,000 TDS - {sum(intr):,} interest = {closing[-1]:,.0f}")]
for i, (a, b, c_) in enumerate(res):
    t.cell(row=16 + i, column=1, value=a).font = Font(bold=True); t.cell(row=16 + i, column=2, value=b); t.cell(row=16 + i, column=4, value=c_)
t.column_dimensions["A"].width = 34
for m in range(N): t.column_dimensions[chr(66 + m)].width = 11

so = wb.create_sheet("Sign_off")
so["A1"] = "Validation Engineer sign-off"; so["A1"].font = Font(bold=True, size=13)
txt = ["This worksheet was laid out with AI assistance (see AI usage log). The course requires the benchmark to be",
       "solved by hand. Before submission the Validation Engineer must recompute every 'Result' with a calculator or",
       "on paper, without looking at the tool, and record any difference below.",
       "", "Recomputed by: ______________________   Date: ____________   Signature: ____________",
       "", "Differences found (line, worksheet value, hand value, resolution):", "", "", ""]
for i, x in enumerate(txt): so.cell(row=3 + i, column=1, value=x)
so.column_dimensions["A"].width = 110
wb.save(OUT)

with open(os.path.join(os.path.dirname(__file__), "benchmark_values.py"), "w") as fh:
    fh.write("# Values typed from 03_Validation/Manual_Benchmark_Hand_Calculation.xlsx (Monthly_table), months 1-13\nBENCH = {\n")
    for k, v in [("Gross running bill", gross), ("Net amount payable", net), ("Total cash in", cin),
                 ("Total cash out", cout), ("Overdraft interest", intr), ("Closing cash balance", closing)]:
        fh.write(f"    {k!r}: {[round(x, 2) for x in v]},\n")
    fh.write("}\n")
print("peak", peak, "interest", sum(intr), "end", closing[-1])
