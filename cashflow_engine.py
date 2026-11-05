"""
Contractor Cash Flow and Working Capital Forecaster - calculation engine
CE725 Team 1, Topic 3

This engine is written directly from the method note (02_Method_Note_and_Data_Sources.docx),
independently of the Excel workbook. It is used for:
  * the Streamlit app (app.py), and
  * an independent cross-check of the Excel tool (test_validation.py).

All amounts are in INR. Time step = 1 month. Horizon = 24 months.
"""
from __future__ import annotations
from dataclasses import dataclass, field, replace

HORIZON = 24


@dataclass
class Activity:
    name: str
    start_month: int      # 1-based
    duration: int         # months, work spread evenly
    share: float          # share of contract value (0-1)


@dataclass
class Inputs:
    contract_value: float = 6_000_000
    activities: list = field(default_factory=lambda: [
        Activity("Site clearance and excavation", 1, 1, 0.04),
        Activity("Foundations and plinth", 1, 2, 0.12),
        Activity("RCC frame (columns, beams, slabs)", 2, 3, 0.27),
        Activity("Brick masonry", 3, 2, 0.14),
        Activity("Plumbing and electrical (MEP)", 3, 4, 0.16),
        Activity("Plaster and flooring", 5, 1, 0.11),
        Activity("Doors, windows, painting, handover", 6, 1, 0.16),
    ])
    billing_cycle: int = 1            # months between running bills (1 or 2)
    payment_lag: int = 1              # months from bill to receipt
    retention: float = 0.05           # share of each gross bill withheld
    dlp_months: int = 6               # retention released this many months after completion
    tds: float = 0.02                 # statutory deduction (income-tax TDS)
    advance: float = 0.10             # mobilisation advance, share of contract value
    advance_month: int = 1            # month advance is received
    recovery_rate: float = 0.15       # share of each gross bill recovered against the advance
    recovery_start: int = 2           # first month in which recovery is made
    cost_ratio: float = 0.85          # total direct + site cost as share of contract value
    cost_split: dict = field(default_factory=lambda: {
        "Labour": 0.30, "Materials": 0.40, "Plant and equipment": 0.15, "Site overheads": 0.15})
    cost_lag: dict = field(default_factory=lambda: {
        "Labour": 0, "Materials": 1, "Plant and equipment": 0, "Site overheads": 0})
    od_rate: float = 0.12             # annual overdraft interest rate
    od_limit: float = 1_000_000       # sanctioned overdraft / working-capital limit
    opening_cash: float = 0.0


def monthly_work_share(inp: Inputs) -> list[float]:
    w = [0.0] * HORIZON
    for a in inp.activities:
        if a.duration <= 0 or a.share == 0:
            continue
        for k in range(a.duration):
            m = a.start_month - 1 + k
            if 0 <= m < HORIZON:
                w[m] += a.share / a.duration
    return w


def run(inp: Inputs) -> dict:
    H = HORIZON
    cv = inp.contract_value
    work = monthly_work_share(inp)
    value = [cv * x for x in work]
    last_work = max((i + 1 for i, x in enumerate(work) if x > 1e-12), default=0)

    # --- billing ---
    gross = [0.0] * H
    cum_value = cum_billed = 0.0
    for i in range(H):
        m = i + 1
        cum_value += value[i]
        bill_month = (m % inp.billing_cycle == 0) or (m == last_work)
        if bill_month and cum_value - cum_billed > 0.005:
            gross[i] = cum_value - cum_billed
            cum_billed = cum_value

    retention = [g * inp.retention for g in gross]
    tds = [g * inp.tds for g in gross]
    adv_amount = cv * inp.advance
    recovery = [0.0] * H
    outstanding = adv_amount
    for i in range(H):
        if i + 1 >= inp.recovery_start and outstanding > 0:
            r = min(gross[i] * inp.recovery_rate, outstanding)
            recovery[i] = r
            outstanding -= r
    net_bill = [gross[i] - retention[i] - recovery[i] - tds[i] for i in range(H)]

    receipts = [net_bill[i - inp.payment_lag] if i - inp.payment_lag >= 0 else 0.0 for i in range(H)]
    advance_in = [adv_amount if i + 1 == inp.advance_month else 0.0 for i in range(H)]
    release_month = last_work + inp.dlp_months
    total_ret = sum(retention)
    ret_release = [total_ret if i + 1 == release_month else 0.0 for i in range(H)]
    inflow = [receipts[i] + advance_in[i] + ret_release[i] for i in range(H)]

    # --- costs ---
    cost_incurred = [v * inp.cost_ratio for v in value]
    cost_paid_by = {}
    for k, share in inp.cost_split.items():
        lag = inp.cost_lag.get(k, 0)
        cost_paid_by[k] = [cost_incurred[i - lag] * share if i - lag >= 0 else 0.0 for i in range(H)]
    outflow = [sum(cost_paid_by[k][i] for k in cost_paid_by) for i in range(H)]

    # --- cash balance and overdraft interest ---
    opening, interest, closing = [], [], []
    bal = inp.opening_cash
    for i in range(H):
        opening.append(bal)
        it = round(max(0.0, -bal) * inp.od_rate / 12)
        interest.append(it)
        bal = bal + inflow[i] - outflow[i] - it
        closing.append(bal)
    funding = [max(0.0, -c) for c in closing]

    cum = lambda xs: [sum(xs[: i + 1]) for i in range(len(xs))]
    peak = max(funding)
    peak_month = funding.index(peak) + 1 if peak > 0 else 0
    total_cost = sum(cost_incurred)
    total_interest = sum(interest)
    total_tds = sum(tds)
    summary = {
        "contract_value": cv,
        "completion_month": last_work,
        "retention_release_month": release_month,
        "peak_funding": peak,
        "peak_month": peak_month,
        "total_interest": total_interest,
        "months_in_overdraft": sum(1 for c in closing if c < -0.005),
        "months_over_limit": sum(1 for f in funding if f > inp.od_limit + 0.005),
        "advance_received": adv_amount,
        "advance_recovered": sum(recovery),
        "retention_held": total_ret,
        "tds_deducted": total_tds,
        "total_cost": total_cost,
        "gross_margin": cv - total_cost,
        "margin_after_interest": cv - total_cost - total_interest,
        "interest_share_of_margin": total_interest / (cv - total_cost) if cv != total_cost else 0.0,
        "closing_cash_end": closing[-1],
        "work_share_total": sum(work),
    }
    return {
        "work_share": work, "value": value, "cum_value": cum(value), "gross_bill": gross,
        "retention": retention, "advance_recovery": recovery, "tds": tds, "net_bill": net_bill,
        "receipts": receipts, "advance_in": advance_in, "retention_release": ret_release,
        "inflow": inflow, "cost_incurred": cost_incurred, "cum_cost": cum(cost_incurred),
        "cost_paid_by": cost_paid_by, "outflow": outflow, "opening": opening, "interest": interest,
        "closing": closing, "funding": funding, "cum_inflow": cum(inflow), "cum_outflow": cum(outflow),
        "summary": summary,
    }


def scenarios() -> dict[str, Inputs]:
    base = Inputs()
    return {
        "Base": base,
        "Payment delay": replace(base, payment_lag=3),
        "Stress": replace(base, payment_lag=3, cost_ratio=0.90, od_rate=0.14,
                          cost_lag={"Labour": 0, "Materials": 0, "Plant and equipment": 0, "Site overheads": 0}),
        "Custom": base,
    }


if __name__ == "__main__":
    for name, inp in scenarios().items():
        s = run(inp)["summary"]
        print(f"{name:14s} peak {s['peak_funding']:>12,.0f} (M{s['peak_month']})  "
              f"interest {s['total_interest']:>9,.0f}  end cash {s['closing_cash_end']:>12,.0f}")
