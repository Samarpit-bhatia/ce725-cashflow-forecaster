"""
Builds Team1_Cash_Flow_Forecaster_v2.xlsx (the Excel tool) with live formulas.
Run:  python build_workbook.py
The workbook is formula-driven; open it in Excel and it recalculates.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.series import SeriesLabel
import os, sys

OUT = os.path.join(os.path.dirname(__file__), "..", "Team1_Cash_Flow_Forecaster_v2.xlsx")
H = 24
M0 = 3                                     # first month column (C)
MC = [L(M0 + i) for i in range(H)]         # C..Z
TOT = L(M0 + H)                            # AA

NAVY = "1F3864"; YELLOW = "FFF2CC"; GREY = "F2F2F2"; GREEN = "E2EFDA"; RED = "FCE4D6"
hdr_font = Font(bold=True, color="FFFFFF"); hdr_fill = PatternFill("solid", fgColor=NAVY)
in_fill = PatternFill("solid", fgColor=YELLOW); sub_fill = PatternFill("solid", fgColor=GREY)
bold = Font(bold=True); title_font = Font(bold=True, size=14, color=NAVY)
thin = Side(style="thin", color="BFBFBF"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
INR = '#,##0;[Red]-#,##0'; PCT = '0.0%'

wb = Workbook()

def header(ws, row, values, col=1):
    for j, v in enumerate(values):
        c = ws.cell(row=row, column=col + j, value=v)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True, vertical="center")

def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = title_font
    if sub: ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="595959")

# ------------------------------------------------------------------ Guide
g = wb.active; g.title = "Guide"
title(g, "Contractor Cash Flow and Working Capital Forecaster (v2)",
      "CE725 Construction Economics and Finance - Team 36 - Topic 3")
rows = [
    ("Purpose", "Converts a project schedule and contract payment terms into monthly cash flows, S-curves, cash-in vs cash-out curves, the peak funding requirement and the overdraft interest cost."),
    ("Sample case", "INR 60 lakh, six-month G+1 site-office and staff-accommodation building, Pune (Base case)."),
    ("How to use - 1", "Schedule: enter activities (name, start month, duration, share of contract value). Shares must total 100%."),
    ("How to use - 2", "Assumptions: edit the yellow cells. Each scenario (Base, Payment delay, Stress, Custom) has its own column."),
    ("How to use - 3", "Dashboard: choose the scenario to display in cell C4. Read the key results, S-curves and cash charts."),
    ("How to use - 4", "Checks: every line must read PASS before results are used. The Dashboard shows the overall status."),
    ("How to use - 5", "Sensitivity: peak funding and interest for payment delays of 0 to 4 months (Base case inputs)."),
    ("How to use - 6", "Validation: tool results against the independent hand-calculated benchmark (Base case)."),
    ("Sheet colours", "Yellow = input. White = formula (do not overwrite). Grey = headings and notes."),
    ("Time step", "Monthly. Horizon 24 months. Work in each activity is spread evenly over its duration."),
    ("Billing", "A running bill is raised for all work done since the last bill, every 'billing cycle' months and in the completion month."),
    ("Deductions", "Retention, mobilisation-advance recovery and TDS are deducted from each gross bill. Advance recovery is a % of the bill, capped so the total recovered never exceeds the advance."),
    ("Receipts", "Net bill is received 'payment lag' months after the bill month. Retention is released in (completion month + defects liability period)."),
    ("Costs", "Cost = value of work x cost ratio, split into labour, materials, plant and overheads. Each category is paid after its own credit period (months)."),
    ("Interest", "Overdraft interest = opening negative balance x annual rate / 12, rounded to the rupee, charged monthly and added to the balance (so it compounds monthly). Positive balances earn nothing."),
    ("Not modelled", "GST and input-tax-credit timing, interest on mobilisation advance, price escalation, bank processing fees, daily interest, disputed or rejected bill items."),
    ("Source code", "An independent Python engine and Streamlit app with the same method are in 1_Tool/source_code (also on GitHub: Samarpit-bhatia/ce725-cashflow-forecaster). Used for cross-checking."),
    ("Accountability", "Planning tool only. It does not replace contract review, lender approval or certified accounts."),
]
header(g, 4, ["Item", "Detail"])
for i, (a, b) in enumerate(rows):
    g.cell(row=5 + i, column=1, value=a).font = bold
    c = g.cell(row=5 + i, column=2, value=b); c.alignment = Alignment(wrap_text=True, vertical="top")
g.column_dimensions["A"].width = 18; g.column_dimensions["B"].width = 120

# ------------------------------------------------------------------ Assumptions
A = wb.create_sheet("Assumptions")
title(A, "Assumptions - contract terms, costs and finance", "Yellow cells are inputs. One column per scenario.")
SC = ["Base", "Payment delay", "Stress", "Custom"]
SCOL = {"Base": "C", "Payment delay": "D", "Stress": "E", "Custom": "F"}
header(A, 5, ["Driver", "Unit"] + SC + ["Basis / source (verify before external use)"])
P = [  # key, label, unit, base, delay, stress, custom, fmt, note
    ("cv", "Contract value", "INR", 6000000, 6000000, 6000000, 6000000, INR, "Sample case: compact six-month low-rise package"),
    ("cycle", "Billing cycle", "months", 1, 1, 1, 1, "0", "Monthly running-account bills are standard; 2 = bi-monthly"),
    ("lag", "Payment lag (bill to receipt)", "months", 1, 3, 3, 1, "0", "Base 30 days; delay/stress 90 days (charter objective: payment-delay impact)"),
    ("ret", "Retention / security deposit", "% of bill", 0.05, 0.05, 0.05, 0.05, PCT, "Typical 5% of each running bill"),
    ("dlp", "Defects liability period", "months", 6, 6, 6, 6, "0", "Retention released this many months after completion"),
    ("tds", "TDS deducted from bills", "% of bill", 0.02, 0.02, 0.02, 0.02, PCT, "Income-tax Act s.194C: 2% (company/firm), 1% (individual/HUF)"),
    ("adv", "Mobilisation advance", "% of contract", 0.10, 0.10, 0.10, 0.10, PCT, "Up to 10% common in public works contracts"),
    ("advm", "Month advance received", "month", 1, 1, 1, 1, "0", "Received at start of work"),
    ("rate", "Advance recovery rate", "% of bill", 0.15, 0.15, 0.15, 0.15, PCT, "Recovered from each bill, capped at the advance"),
    ("start", "Recovery starts in month", "month", 2, 2, 2, 2, "0", "Recovery from second bill onward"),
    ("cost", "Total cash cost", "% of contract", 0.85, 0.85, 0.90, 0.85, PCT, "Leaves ~15% margin before finance cost; stress assumes overrun"),
    ("sL", "Cost share - Labour", "% of cost", 0.30, 0.30, 0.30, 0.30, PCT, "Cost split for a low-rise RCC building"),
    ("sM", "Cost share - Materials", "% of cost", 0.40, 0.40, 0.40, 0.40, PCT, ""),
    ("sP", "Cost share - Plant and equipment", "% of cost", 0.15, 0.15, 0.15, 0.15, PCT, ""),
    ("sO", "Cost share - Site overheads", "% of cost", 0.15, 0.15, 0.15, 0.15, PCT, ""),
    ("lL", "Payment delay to labour", "months", 0, 0, 0, 0, "0", "Wages paid in the work month"),
    ("lM", "Supplier credit - Materials", "months", 1, 1, 0, 1, "0", "30-day supplier credit; stress: suppliers demand cash"),
    ("lP", "Payment delay - Plant hire", "months", 0, 0, 0, 0, "0", ""),
    ("lO", "Payment delay - Site overheads", "months", 0, 0, 0, 0, "0", ""),
    ("od", "Overdraft interest rate", "% p.a.", 0.12, 0.12, 0.14, 0.12, PCT, "Indicative MSME working-capital rate; confirm with bank"),
    ("lim", "Sanctioned overdraft limit", "INR", 1000000, 1000000, 1000000, 1000000, INR, "Used to flag months where funding need exceeds the limit"),
    ("open", "Opening cash (own funds)", "INR", 0, 0, 0, 0, INR, "Cash the contractor puts in at start"),
]
PROW = {}
for i, p in enumerate(P):
    r = 6 + i; PROW[p[0]] = r
    A.cell(row=r, column=1, value=p[1]); A.cell(row=r, column=2, value=p[2])
    for j, v in enumerate(p[3:7]):
        c = A.cell(row=r, column=3 + j, value=v); c.fill = in_fill; c.number_format = p[7]; c.border = box
    A.cell(row=r, column=7, value=p[8]).alignment = Alignment(wrap_text=True)
r = 6 + len(P) + 1
A.cell(row=r, column=1, value="Cost shares total").font = bold
for j in range(4):
    col = L(3 + j)
    c = A.cell(row=r, column=3 + j, value=f"=SUM({col}{PROW['sL']}:{col}{PROW['sO']})"); c.number_format = PCT
PROW["sharetot"] = r
for col, w in zip("ABCDEFG", [34, 14, 14, 14, 14, 14, 70]): A.column_dimensions[col].width = w
dv = DataValidation(type="decimal", operator="between", formula1=0, formula2=1, showErrorMessage=True,
                    error="Enter a percentage between 0% and 100%")
A.add_data_validation(dv)
for k in ["ret", "tds", "adv", "rate", "cost", "sL", "sM", "sP", "sO", "od"]:
    dv.add(f"C{PROW[k]}:F{PROW[k]}")
dvi = DataValidation(type="whole", operator="between", formula1=0, formula2=12, showErrorMessage=True,
                     error="Enter a whole number of months (0-12)")
A.add_data_validation(dvi)
for k in ["lag", "dlp", "lL", "lM", "lP", "lO"]:
    dvi.add(f"C{PROW[k]}:F{PROW[k]}")
dvc = DataValidation(type="whole", operator="between", formula1=1, formula2=12, showErrorMessage=True,
                     error="Enter a whole number of months (1-12)")
A.add_data_validation(dvc)
for k in ["cycle", "advm", "start"]:
    dvc.add(f"C{PROW[k]}:F{PROW[k]}")

# ------------------------------------------------------------------ Schedule
S = wb.create_sheet("Schedule")
title(S, "Project schedule - activities converted to monthly work done",
      "Enter activities in yellow cells. Work is spread evenly over each activity's duration. Shared by all scenarios.")
SM0 = 7  # month columns G..AD
SMC = [L(SM0 + i) for i in range(H)]
header(S, 5, ["Activity", "Start month", "Duration (months)", "Share of contract value", "End month", ""] + [f"M{i+1}" for i in range(H)])
acts = [("Site clearance and excavation", 1, 1, 0.04), ("Foundations and plinth", 1, 2, 0.12),
        ("RCC frame (columns, beams, slabs)", 2, 3, 0.27), ("Brick masonry", 3, 2, 0.14),
        ("Plumbing and electrical (MEP)", 3, 4, 0.16), ("Plaster and flooring", 5, 1, 0.11),
        ("Doors, windows, painting, handover", 6, 1, 0.16)]
NA = 12
for i in range(NA):
    r = 6 + i
    a = acts[i] if i < len(acts) else ("", None, None, None)
    for j, v in enumerate(a):
        c = S.cell(row=r, column=1 + j, value=v); c.fill = in_fill; c.border = box
    S.cell(row=r, column=4).number_format = PCT
    S.cell(row=r, column=5, value=f'=IF(N(C{r})>0,B{r}+C{r}-1,"")')
    for m, col in enumerate(SMC):
        c = S.cell(row=r, column=SM0 + m,
                   value=f"=IF(AND(N($C{r})>0,{m+1}>=N($B{r}),{m+1}<=N($B{r})+N($C{r})-1),N($D{r})/N($C{r}),0)")
        c.number_format = PCT
RT = 6 + NA            # total row
S.cell(row=RT, column=1, value="Total share / monthly work done").font = bold
c = S.cell(row=RT, column=4, value=f"=SUM(D6:D{RT-1})"); c.number_format = PCT; c.font = bold
for m, col in enumerate(SMC):
    c = S.cell(row=RT, column=SM0 + m, value=f"=SUM({col}6:{col}{RT-1})"); c.number_format = PCT; c.font = bold
S.cell(row=RT + 1, column=1, value="Cumulative work done (planned S-curve)").font = bold
for m, col in enumerate(SMC):
    c = S.cell(row=RT + 1, column=SM0 + m, value=f"=SUM(${SMC[0]}{RT}:{col}{RT})"); c.number_format = PCT
S.cell(row=RT + 2, column=1, value="Month number")
for m in range(H): S.cell(row=RT + 2, column=SM0 + m, value=m + 1)
S.cell(row=RT + 3, column=1, value="Completion month (last month with work)").font = bold
S.cell(row=RT + 3, column=4, value=f"=SUMPRODUCT(MAX(({SMC[0]}{RT}:{SMC[-1]}{RT}>0.0000001)*{SMC[0]}{RT+2}:{SMC[-1]}{RT+2}))")
S.cell(row=RT + 4, column=1, value="Work scheduled within 24-month horizon")
c = S.cell(row=RT + 4, column=4, value=f"=SUM({SMC[0]}{RT}:{SMC[-1]}{RT})"); c.number_format = PCT
SCHED_ROW = RT; COMPL = f"Schedule!$D${RT+3}"
S.column_dimensions["A"].width = 38
for col in "BCDE": S.column_dimensions[col].width = 12
S.column_dimensions["F"].width = 2
for col in SMC: S.column_dimensions[col].width = 7
S.freeze_panes = "B6"

# ------------------------------------------------------------------ Calculation sheets
ROWS = {}
def calc_sheet(name, col, lag_override=None, hidden=False, subtitle=""):
    ws = wb.create_sheet(name)
    a = lambda k: f"Assumptions!${col}${PROW[k]}"
    lag = lag_override if lag_override is not None else a("lag")
    title(ws, f"Monthly cash flow calculation - {subtitle or name}",
          "All cells are formulas. Do not overwrite. Inputs come from Assumptions and Schedule.")
    labels = [
        (4, "Month", ""), (5, "Work done this month", "% of contract"), (6, "Value of work done", "INR"),
        (7, "Cumulative value of work done", "INR"), (8, "Running bill raised this month?", "1 = yes"),
        (9, "Gross running bill", "INR"), (10, "Less: retention", "INR"), (11, "Less: advance recovery", "INR"),
        (12, "Less: TDS", "INR"), (13, "Net amount payable on bill", "INR"),
        (14, "Client payment received", "INR"), (15, "Mobilisation advance received", "INR"),
        (16, "Retention released", "INR"), (17, "TOTAL CASH IN", "INR"),
        (18, "Cost incurred (value x cost ratio)", "INR"), (19, "Labour paid", "INR"), (20, "Materials paid", "INR"),
        (21, "Plant and equipment paid", "INR"), (22, "Site overheads paid", "INR"), (23, "TOTAL CASH OUT (before interest)", "INR"),
        (24, "Opening cash balance", "INR"), (25, "Overdraft interest", "INR"), (26, "Net cash flow after interest", "INR"),
        (27, "CLOSING CASH BALANCE", "INR"), (28, "Funding requirement (overdraft used)", "INR"),
        (29, "Over sanctioned limit?", "1 = yes"), (30, "Cumulative cash in", "INR"),
        (31, "Cumulative cash out incl. interest", "INR"), (32, "Cumulative cost incurred", "INR"),
    ]
    for r, t, u in labels:
        ws.cell(row=r, column=1, value=t); ws.cell(row=r, column=2, value=u)
        if t.isupper() or r == 4:
            ws.cell(row=r, column=1).font = bold
    header(ws, 4, ["Month", "Unit"] + list(range(1, H + 1)) + ["Total"])
    for i, c in enumerate(MC):
        p = L(M0 + i - 1)  # previous column (B for first)
        f = {
            5: f"=Schedule!{SMC[i]}${SCHED_ROW}",
            6: f"={a('cv')}*{c}5",
            7: f"=SUM($C6:{c}6)",
            8: f"=IF(AND(OR(MOD({c}$4,{a('cycle')})=0,{c}$4={COMPL}),{c}7-SUM($B9:{p}9)>0.005),1,0)",
            9: f"=IF({c}8=1,{c}7-SUM($B9:{p}9),0)",
            10: f"={c}9*{a('ret')}",
            11: f"=IF({c}$4>={a('start')},MAX(0,MIN({c}9*{a('rate')},{a('cv')}*{a('adv')}-SUM($B11:{p}11))),0)",
            12: f"={c}9*{a('tds')}",
            13: f"={c}9-{c}10-{c}11-{c}12",
            14: f"=IF({c}$4>{lag},INDEX($C13:${MC[-1]}13,{c}$4-{lag}),0)",
            15: f"=IF({c}$4={a('advm')},{a('cv')}*{a('adv')},0)",
            16: f"=IF({c}$4={COMPL}+{a('dlp')},SUM($C10:${MC[-1]}10),0)",
            17: f"=SUM({c}14:{c}16)",
            18: f"={c}6*{a('cost')}",
            19: f"=IF({c}$4>{a('lL')},INDEX($C18:${MC[-1]}18,{c}$4-{a('lL')})*{a('sL')},0)",
            20: f"=IF({c}$4>{a('lM')},INDEX($C18:${MC[-1]}18,{c}$4-{a('lM')})*{a('sM')},0)",
            21: f"=IF({c}$4>{a('lP')},INDEX($C18:${MC[-1]}18,{c}$4-{a('lP')})*{a('sP')},0)",
            22: f"=IF({c}$4>{a('lO')},INDEX($C18:${MC[-1]}18,{c}$4-{a('lO')})*{a('sO')},0)",
            23: f"=SUM({c}19:{c}22)",
            24: f"={a('open')}" if i == 0 else f"={p}27",
            25: f"=ROUND(MAX(0,-{c}24)*{a('od')}/12,0)",
            26: f"={c}17-{c}23-{c}25",
            27: f"={c}24+{c}26",
            28: f"=MAX(0,-{c}27)",
            29: f"=IF({c}28>{a('lim')}+0.005,1,0)",
            30: f"=SUM($C17:{c}17)",
            31: f"=SUM($C23:{c}23)+SUM($C25:{c}25)",
            32: f"=SUM($C18:{c}18)",
        }
        for r, v in f.items():
            cell = ws.cell(row=r, column=M0 + i, value=v)
            cell.number_format = PCT if r == 5 else ("0" if r in (8, 29) else INR)
    for r in [5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 25, 26]:
        c = ws.cell(row=r, column=M0 + H, value=f"=SUM(C{r}:{MC[-1]}{r})")
        c.number_format = PCT if r == 5 else INR; c.font = bold
    for r in (17, 23, 27):
        for i in range(H + 2):
            ws.cell(row=r, column=1 + i).fill = sub_fill
    # summary
    ws.cell(row=34, column=1, value="KEY RESULTS").font = bold
    summ = [
        (35, "Completion month", f"={COMPL}", "0"),
        (36, "Retention release month", f"={COMPL}+{a('dlp')}", "0"),
        (37, "Peak funding requirement", f"=MAX(C28:{MC[-1]}28)", INR),
        (38, "Month of peak funding", f"=IF(C37>0,MATCH(C37,C28:{MC[-1]}28,0),0)", "0"),
        (39, "Total overdraft interest", f"=SUM(C25:{MC[-1]}25)", INR),
        (40, "Months in overdraft", f'=COUNTIF(C27:{MC[-1]}27,"<-0.005")', "0"),
        (41, "Months above sanctioned limit", f"=SUM(C29:{MC[-1]}29)", "0"),
        (42, "Mobilisation advance received", f"=SUM(C15:{MC[-1]}15)", INR),
        (43, "Mobilisation advance recovered", f"=SUM(C11:{MC[-1]}11)", INR),
        (44, "Retention held (and released)", f"=SUM(C10:{MC[-1]}10)", INR),
        (45, "TDS deducted (claimable against tax)", f"=SUM(C12:{MC[-1]}12)", INR),
        (46, "Total project cost", f"=SUM(C18:{MC[-1]}18)", INR),
        (47, "Gross margin before finance cost", f"={a('cv')}-C46", INR),
        (48, "Margin after overdraft interest", f"=C47-C39", INR),
        (49, "Interest as % of gross margin", f"=IF(C47<>0,C39/C47,0)", PCT),
        (50, "Closing cash at month 24", f"={MC[-1]}27", INR),
        (51, "Reconciliation difference (should be 0)", f"=ROUND(C50-({a('open')}+{a('cv')}-C46-C45-C39),2)", INR),
        (52, "Sanctioned overdraft limit", f"={a('lim')}", INR),
    ]
    for r, t, f, fmt in summ:
        ws.cell(row=r, column=1, value=t)
        c = ws.cell(row=r, column=3, value=f); c.number_format = fmt; c.font = bold
    ws.cell(row=51, column=4, value="Closing cash should equal opening cash + contract value - cost - TDS - interest once all receipts fall inside the horizon.")
    ws.column_dimensions["A"].width = 38; ws.column_dimensions["B"].width = 12
    for c in MC + [TOT]: ws.column_dimensions[c].width = 12
    ws.freeze_panes = "C5"
    if hidden: ws.sheet_state = "hidden"
    return ws

CALC = {"Base": "Calc_Base", "Payment delay": "Calc_Delay", "Stress": "Calc_Stress", "Custom": "Calc_Custom"}
for sc, nm in CALC.items():
    calc_sheet(nm, SCOL[sc], subtitle=f"{sc} scenario")

# ------------------------------------------------------------------ Dashboard (insert at position 1)
D = wb.create_sheet("Dashboard", 1)
title(D, "Dashboard - Contractor Cash Flow and Working Capital",
      "Choose the scenario in C4. All figures in INR. Charts update automatically.")
D["A4"] = "Scenario shown"; D["A4"].font = bold
D["C4"] = "Base"; D["C4"].fill = in_fill; D["C4"].font = Font(bold=True, size=12); D["C4"].border = box
dvs = DataValidation(type="list", formula1='"Base,Payment delay,Stress,Custom"', allow_blank=False)
D.add_data_validation(dvs); dvs.add("C4")
D["E4"] = '=MATCH(C4,{"Base","Payment delay","Stress","Custom"},0)'; D["E4"].font = Font(color="FFFFFF")
D["A5"] = "Input checks"; D["A5"].font = bold
D["C5"] = '=Checks!C4'; D["C5"].font = bold
D.conditional_formatting.add("C5", CellIsRule(operator="equal", formula=['"ALL PASS"'], fill=PatternFill("solid", fgColor=GREEN)))
D.conditional_formatting.add("C5", CellIsRule(operator="notEqual", formula=['"ALL PASS"'], fill=PatternFill("solid", fgColor=RED)))

def pick(r):
    return "=CHOOSE($E$4," + ",".join(f"{CALC[s]}!$C${r}" for s in SC) + ")"
header(D, 7, ["Key result (selected scenario)", "", "Value"])
kpis = [("Peak funding requirement", 37, INR), ("Month of peak funding", 38, "0"),
        ("Total overdraft interest", 39, INR), ("Interest as % of gross margin", 49, PCT),
        ("Months in overdraft", 40, "0"), ("Months above sanctioned limit", 41, "0"),
        ("Sanctioned overdraft limit", 52, INR), ("Gross margin before finance cost", 47, INR),
        ("Margin after overdraft interest", 48, INR), ("Completion month", 35, "0"),
        ("Retention release month", 36, "0")]
for i, (t, r, fmt) in enumerate(kpis):
    D.cell(row=8 + i, column=1, value=t)
    c = D.cell(row=8 + i, column=3, value=pick(r)); c.number_format = fmt; c.font = bold; c.border = box
D["A20"] = "What this means"; D["A20"].font = bold
D["C20"] = ('=IF(C8=0,"No overdraft is needed in this scenario.",'
            '"Arrange a working-capital facility of at least INR "&TEXT(C8,"#,##0")&" by month "&C9&". '
            'Overdraft interest of INR "&TEXT(C10,"#,##0")&" uses "&TEXT(C11,"0.0%")&" of the margin."'
            '&IF(C13>0," WARNING: the need exceeds the sanctioned limit in "&C13&" month(s).",""))')
D["C20"].alignment = Alignment(wrap_text=True, vertical="top"); D.merge_cells("C20:H22")

header(D, 24, ["Scenario comparison", ""] + SC)
comp = [("Payment lag (months)", None, "0"), ("Peak funding requirement", 37, INR), ("Month of peak", 38, "0"),
        ("Total overdraft interest", 39, INR), ("Interest % of margin", 49, PCT),
        ("Months above limit", 41, "0"), ("Margin after interest", 48, INR)]
for i, (t, r, fmt) in enumerate(comp):
    D.cell(row=25 + i, column=1, value=t)
    for j, s in enumerate(SC):
        f = f"=Assumptions!{SCOL[s]}{PROW['lag']}" if r is None else f"={CALC[s]}!$C${r}"
        c = D.cell(row=25 + i, column=3 + j, value=f); c.number_format = fmt; c.border = box

# chart data block
CD = 60
D.cell(row=CD - 1, column=1, value="Chart data (selected scenario) - formulas, do not edit").font = bold
cd_rows = [("Month", None), ("Cumulative value of work (S-curve)", 7), ("Cumulative cost incurred", 32),
           ("Cumulative cash in", 30), ("Cumulative cash out incl. interest", 31), ("Cash in", 17),
           ("Cash out incl. interest", None), ("Closing cash balance", 27), ("Overdraft limit (negative)", None)]
for k, (t, r) in enumerate(cd_rows):
    rr = CD + k
    D.cell(row=rr, column=1, value=t)
    for i, c in enumerate(MC):
        col = M0 + i
        if t == "Month":
            v = i + 1
        elif t.startswith("Cash out"):
            v = "=CHOOSE($E$4," + ",".join(f"{CALC[s]}!{c}23+{CALC[s]}!{c}25" for s in SC) + ")"
        elif t.startswith("Overdraft limit"):
            v = f"=-C14"
        else:
            v = "=CHOOSE($E$4," + ",".join(f"{CALC[s]}!{c}{r}" for s in SC) + ")"
        cell = D.cell(row=rr, column=col, value=v); cell.number_format = INR
D.cell(row=CD + 10, column=1, value="Closing balance by scenario").font = bold
for j, s in enumerate(SC):
    rr = CD + 11 + j
    D.cell(row=rr, column=1, value=s)
    for i, c in enumerate(MC):
        D.cell(row=rr, column=M0 + i, value=f"={CALC[s]}!{c}27").number_format = INR

months = Reference(D, min_col=M0, max_col=M0 + H - 1, min_row=CD)
def line_chart(title_, rows_, ytitle="INR", w=17, h=8.5):
    ch = LineChart(); ch.title = title_; ch.y_axis.title = ytitle; ch.x_axis.title = "Month"
    ch.width = w; ch.height = h
    for rr in rows_:
        ref = Reference(D, min_col=M0, max_col=M0 + H - 1, min_row=rr)
        ch.add_data(ref, from_rows=True, titles_from_data=False)
        ch.series[-1].tx = SeriesLabel(v=D.cell(row=rr, column=1).value)
    ch.set_categories(months); ch.y_axis.number_format = '#,##0'
    ch.x_axis.delete = False; ch.y_axis.delete = False
    return ch

D.add_chart(line_chart("S-curves: cumulative value, cost and cash", [CD + 1, CD + 2, CD + 3, CD + 4]), "J4")
bc = BarChart(); bc.type = "col"; bc.title = "Monthly cash in vs cash out"; bc.width = 17; bc.height = 8.5
for rr in (CD + 5, CD + 6):
    bc.add_data(Reference(D, min_col=M0, max_col=M0 + H - 1, min_row=rr), from_rows=True, titles_from_data=False)
    bc.series[-1].tx = SeriesLabel(v=D.cell(row=rr, column=1).value)
bc.set_categories(months); bc.y_axis.number_format = '#,##0'; bc.x_axis.title = "Month"
bc.x_axis.delete = False; bc.y_axis.delete = False
D.add_chart(bc, "J22")
D.add_chart(line_chart("Closing cash balance vs overdraft limit", [CD + 7, CD + 8]), "J40")
D.add_chart(line_chart("Closing cash balance - all scenarios", [CD + 11 + j for j in range(4)]), "A34")
D.column_dimensions["A"].width = 34; D.column_dimensions["B"].width = 2
for c in MC: D.column_dimensions[c].width = 12   # wide enough for -99,99,999
for c in "CDEF": D.column_dimensions[c].width = 15

# ------------------------------------------------------------------ Sensitivity
SE = wb.create_sheet("Sensitivity")
title(SE, "Payment-delay sensitivity (Base case inputs, payment lag varied)",
      "Answers the charter objective: how much does late payment by the client cost the contractor?")
header(SE, 4, ["Payment lag (months)", "Peak funding requirement", "Month of peak", "Total overdraft interest",
               "Interest % of margin", "Months above limit", "Extra peak funding vs 0-month lag"])
for k in range(5):
    nm = f"Sens_Lag{k}"
    calc_sheet(nm, SCOL["Base"], lag_override=f"Sensitivity!$A${5+k}", hidden=True, subtitle=f"Base with {k}-month lag")
    r = 5 + k
    SE.cell(row=r, column=1, value=k).fill = in_fill
    for j, (src, fmt) in enumerate([(37, INR), (38, "0"), (39, INR), (49, PCT), (41, "0")]):
        c = SE.cell(row=r, column=2 + j, value=f"={nm}!$C${src}"); c.number_format = fmt; c.border = box
    c = SE.cell(row=r, column=7, value=f"=B{r}-$B$5"); c.number_format = INR
SE["A11"] = "Note: the lag values in column A can be changed (whole months). Each row has its own hidden calculation sheet."
for c, w in zip("ABCDEFG", [20, 24, 14, 22, 18, 18, 30]): SE.column_dimensions[c].width = w
sb = BarChart(); sb.type = "col"; sb.title = "Peak funding requirement vs payment lag"
sb.add_data(Reference(SE, min_col=2, min_row=4, max_row=9), titles_from_data=True)
sb.set_categories(Reference(SE, min_col=1, min_row=5, max_row=9))
sb.x_axis.title = "Payment lag (months)"; sb.y_axis.title = "INR"; sb.y_axis.number_format = '#,##0'
sb.x_axis.delete = False; sb.y_axis.delete = False; sb.width = 16; sb.height = 8
SE.add_chart(sb, "A13")

# ------------------------------------------------------------------ Checks
CK = wb.create_sheet("Checks")
title(CK, "Input and model checks", "Every line must read PASS before results are used.")
CK["A4"] = "Overall status"; CK["A4"].font = bold
header(CK, 6, ["Check", "Scenario", "Result", "Value", "Expected"])
checks = [("Activity shares total 100%", "All", f"=ABS(Schedule!D{SCHED_ROW}-1)<0.0001", f"=Schedule!D{SCHED_ROW}", 1),
          ("All work falls inside the 24-month horizon", "All", f"=ABS(Schedule!D{SCHED_ROW+4}-Schedule!D{SCHED_ROW})<0.0001", f"=Schedule!D{SCHED_ROW+4}", f"=Schedule!D{SCHED_ROW}"),
          ("Activity start months and durations are whole numbers >= 1", "All",
           f"=SUMPRODUCT((N(+Schedule!C6:C{SCHED_ROW-1})>0)*((Schedule!B6:B{SCHED_ROW-1}<1)+(Schedule!B6:B{SCHED_ROW-1}<>INT(Schedule!B6:B{SCHED_ROW-1}))+(Schedule!C6:C{SCHED_ROW-1}<>INT(Schedule!C6:C{SCHED_ROW-1}))))=0", "", "0 errors")]
for s in SC:
    col = SCOL[s]; cs = CALC[s]
    checks += [
        ("Cost shares total 100%", s, f"=ABS(Assumptions!{col}{PROW['sharetot']}-1)<0.0001", f"=Assumptions!{col}{PROW['sharetot']}", 1),
        ("Advance fully recovered (INR recovered = INR received)", s, f"=ABS({cs}!C43-{cs}!C42)<1", f"={cs}!C43", f"={cs}!C42"),
        ("Retention release month within horizon", s, f"={cs}!C36<=24", f"={cs}!C36", "<= 24"),
        ("Last bill received within horizon", s, f"={cs}!C35+Assumptions!{col}{PROW['lag']}<=24", f"={cs}!C35+Assumptions!{col}{PROW['lag']}", "<= 24"),
        ("Cash reconciles (closing cash = CV - cost - TDS - interest)", s, f"=ABS({cs}!C51)<1", f"={cs}!C51", 0),
        ("Deductions do not exceed the bill (retention + TDS + recovery rate < 100%)", s,
         f"=Assumptions!{col}{PROW['ret']}+Assumptions!{col}{PROW['tds']}+Assumptions!{col}{PROW['rate']}<1", "", "< 100%"),
    ]
for i, (t, s, f, v, e) in enumerate(checks):
    r = 7 + i
    CK.cell(row=r, column=1, value=t); CK.cell(row=r, column=2, value=s)
    CK.cell(row=r, column=3, value=f'=IF({f[1:]},"PASS","CHECK")').font = bold
    is_pct = "100%" in t or "horizon" in t          # share checks are percentages
    fmt = PCT if is_pct else ('0' if "month" in t else '#,##0')
    c = CK.cell(row=r, column=4, value=v if v else None); c.number_format = fmt
    e_cell = CK.cell(row=r, column=5, value=e)
    if not isinstance(e, str) or e.startswith("="):
        e_cell.number_format = fmt
    e_cell.alignment = Alignment(horizontal="right")
    c.alignment = Alignment(horizontal="right")
last = 7 + len(checks) - 1
CK["C4"] = f'=IF(COUNTIF(C7:C{last},"PASS")={len(checks)},"ALL PASS",COUNTIF(C7:C{last},"CHECK")&" CHECK(S) FAILED")'
CK["C4"].font = Font(bold=True, size=12)
CK.conditional_formatting.add(f"C7:C{last}", CellIsRule(operator="equal", formula=['"PASS"'], fill=PatternFill("solid", fgColor=GREEN)))
CK.conditional_formatting.add(f"C7:C{last}", CellIsRule(operator="equal", formula=['"CHECK"'], fill=PatternFill("solid", fgColor=RED)))
for c, w in zip("ABCDE", [70, 16, 10, 16, 16]): CK.column_dimensions[c].width = w

# ------------------------------------------------------------------ Validation
V = wb.create_sheet("Validation")
title(V, "Validation - tool (Calc_Base) vs independent hand-calculated benchmark",
      "Benchmark values are TYPED from 3_Validation/Manual_Benchmark_Hand_Calculation.xlsx (no links to the tool). Base inputs only.")
sys.path.insert(0, os.path.dirname(__file__))
BENCH = {}
bench_path = os.path.join(os.path.dirname(__file__), "benchmark_values.py")
if os.path.exists(bench_path):
    from benchmark_values import BENCH  # typed from the hand worksheet
vr = [("Gross running bill", 9), ("Net amount payable", 13), ("Total cash in", 17), ("Total cash out", 23),
      ("Overdraft interest", 25), ("Closing cash balance", 27)]
NB = 13
header(V, 4, ["Line item", "Source"] + [f"M{i+1}" for i in range(NB)])
r = 5
for t, src in vr:
    V.cell(row=r, column=1, value=t).font = bold; V.cell(row=r, column=2, value="Hand benchmark")
    V.cell(row=r + 1, column=2, value="Tool")
    V.cell(row=r + 2, column=2, value="Difference")
    for i in range(NB):
        c = L(3 + i)
        b = V.cell(row=r, column=3 + i, value=BENCH.get(t, [None] * NB)[i]); b.number_format = INR; b.fill = in_fill
        V.cell(row=r + 1, column=3 + i, value=f"=Calc_Base!{MC[i]}{src}").number_format = INR
        d = V.cell(row=r + 2, column=3 + i, value=f'=IF({c}{r}="","",ROUND({c}{r}-{c}{r+1},2))'); d.number_format = INR
    r += 4
V.cell(row=r, column=1, value="Cells with |difference| >= INR 1").font = bold
V.cell(row=r, column=3, value=f'=SUMPRODUCT((C7:{L(2+NB)}{r-2}<>"")*(ABS(N(+C7:{L(2+NB)}{r-2}))>=1)*(MOD(ROW(C7:{L(2+NB)}{r-2})-7,4)=0))')
V.cell(row=r + 1, column=1, value="Validation result").font = bold
V.cell(row=r + 1, column=3, value=f'=IF(COUNT(C5:{L(2+NB)}5)=0,"BENCHMARK NOT ENTERED",IF(C{r}=0,"PASS","REVIEW"))').font = bold
V.column_dimensions["A"].width = 26; V.column_dimensions["B"].width = 16
for i in range(NB): V.column_dimensions[L(3 + i)].width = 12
V.freeze_panes = "C5"

# order sheets
order = ["Guide", "Dashboard", "Assumptions", "Schedule", "Calc_Base", "Calc_Delay", "Calc_Stress", "Calc_Custom",
         "Sensitivity", "Checks", "Validation"] + [f"Sens_Lag{k}" for k in range(5)]
wb._sheets = [wb[n] for n in order]
wb.active = 1
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", os.path.abspath(OUT))
