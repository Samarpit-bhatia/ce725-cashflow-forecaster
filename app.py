"""
Streamlit app - Contractor Cash Flow and Working Capital Forecaster (CE725 Team 1, Topic 3)
Same method as the Excel tool, using cashflow_engine.py.
Run locally:   streamlit run app.py
Live version: https://ce725-cashflow-forecaster.streamlit.app
Privacy: run locally, nothing leaves your computer. The hosted version processes inputs on
Streamlit Community Cloud for the session only; nothing is saved. Use sample or anonymised figures there.
"""
from dataclasses import replace
import pandas as pd
import altair as alt
import streamlit as st
import cashflow_engine as E

st.set_page_config(page_title="Contractor Cash Flow Forecaster", layout="wide")
st.title("Contractor Cash Flow and Working Capital Forecaster")
st.caption("CE725 Team 1 · Topic 3 · Converts a project schedule and payment terms into monthly cash flows, "
           "S-curves, peak funding requirement and overdraft interest. Inputs are used only for this session and are not saved. Use sample or anonymised figures on the hosted version.")

base = E.Inputs()
presets = E.scenarios()

with st.sidebar:
    st.header("Inputs")
    preset = st.selectbox("Start from scenario", ["Base", "Payment delay", "Stress"])
    p = presets[preset]
    cv = st.number_input("Contract value (INR)", 100_000, 10_000_000_000, int(p.contract_value), step=100_000)
    st.subheader("Billing and payment")
    cycle = st.selectbox("Billing cycle (months)", [1, 2], index=p.billing_cycle - 1)
    lag = st.slider("Payment lag: bill to receipt (months)", 0, 6, p.payment_lag)
    ret = st.slider("Retention (% of bill)", 0.0, 10.0, p.retention * 100, 0.5) / 100
    dlp = st.slider("Defects liability period (months)", 0, 12, p.dlp_months)
    tds = st.slider("TDS (% of bill)", 0.0, 5.0, p.tds * 100, 0.5) / 100
    st.subheader("Mobilisation advance")
    adv = st.slider("Advance (% of contract)", 0.0, 20.0, p.advance * 100, 1.0) / 100
    rate = st.slider("Recovery rate (% of each bill)", 0.0, 50.0, p.recovery_rate * 100, 1.0) / 100
    start = st.slider("Recovery starts in month", 1, 12, p.recovery_start)
    st.subheader("Costs")
    cost = st.slider("Total cost (% of contract)", 50.0, 110.0, p.cost_ratio * 100, 1.0) / 100
    mlag = st.slider("Supplier credit on materials (months)", 0, 3, p.cost_lag["Materials"])
    st.subheader("Finance")
    od = st.slider("Overdraft rate (% p.a.)", 0.0, 24.0, p.od_rate * 100, 0.5) / 100
    lim = st.number_input("Sanctioned overdraft limit (INR)", 0, 10_000_000_000, int(p.od_limit), step=100_000)
    opening = st.number_input("Opening own funds (INR)", 0, 10_000_000_000, int(p.opening_cash), step=100_000)

st.subheader("1. Project schedule")
st.write("Edit activities: start month, duration and share of contract value. Work is spread evenly over the duration.")
act_df = pd.DataFrame([{"Activity": a.name, "Start month": a.start_month, "Duration (months)": a.duration,
                        "Share of contract (%)": round(a.share * 100, 2)} for a in base.activities])
edited = st.data_editor(act_df, num_rows="dynamic", width="stretch", key="acts")
acts = []
for _, r in edited.dropna(subset=["Start month", "Duration (months)", "Share of contract (%)"]).iterrows():
    acts.append(E.Activity(str(r["Activity"]), int(r["Start month"]), int(r["Duration (months)"]),
                           float(r["Share of contract (%)"]) / 100))
share_total = sum(a.share for a in acts)

inp = replace(p, contract_value=cv, activities=acts, billing_cycle=cycle, payment_lag=lag, retention=ret,
              dlp_months=dlp, tds=tds, advance=adv, recovery_rate=rate, recovery_start=start, cost_ratio=cost,
              cost_lag={**p.cost_lag, "Materials": mlag}, od_rate=od, od_limit=lim, opening_cash=opening)

problems = []
if abs(share_total - 1) > 1e-6: problems.append(f"Activity shares total {share_total:.1%}; they must total 100%.")
if ret + tds + rate >= 1: problems.append("Retention + TDS + recovery rate is 100% or more of each bill.")
if any(a.start_month < 1 or a.duration < 1 for a in acts): problems.append("Start month and duration must be at least 1.")
res = E.run(inp); s = res["summary"]
if s["retention_release_month"] > E.HORIZON: problems.append("Retention release falls after month 24; shorten the DLP or schedule.")
if abs(s["advance_recovered"] - s["advance_received"]) > 1: problems.append("The advance is not fully recovered within the bills; raise the recovery rate.")
for pr in problems: st.error(pr)
if not problems: st.success("All input checks pass.")

st.subheader("2. Key results")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Peak funding requirement", f"₹{s['peak_funding']:,.0f}", f"month {s['peak_month']}" if s["peak_month"] else "none", delta_color="off")
c2.metric("Overdraft interest", f"₹{s['total_interest']:,.0f}", f"{s['interest_share_of_margin']:.1%} of margin", delta_color="inverse")
c3.metric("Months in overdraft", s["months_in_overdraft"])
c4.metric("Months above limit", s["months_over_limit"], delta_color="inverse")
if s["peak_funding"] > 0:
    msg = (f"Arrange a working-capital facility of at least **₹{s['peak_funding']:,.0f}** by month {s['peak_month']}. "
           f"Overdraft interest of ₹{s['total_interest']:,.0f} reduces the margin from ₹{s['gross_margin']:,.0f} "
           f"to ₹{s['margin_after_interest']:,.0f}.")
    if s["months_over_limit"]:
        msg += f" **Warning:** the need exceeds the sanctioned limit of ₹{lim:,.0f} in {s['months_over_limit']} month(s)."
    st.info(msg)
else:
    st.info("No overdraft is needed in this scenario.")

H = E.HORIZON
show = min(H, max(s["retention_release_month"], s["completion_month"] + inp.payment_lag) + 1)
months = list(range(1, H + 1))
df = pd.DataFrame({"Month": months, "Cumulative value of work": res["cum_value"], "Cumulative cost incurred": res["cum_cost"],
                   "Cumulative cash in": res["cum_inflow"], "Cumulative cash out": res["cum_outflow"],
                   "Cash in": res["inflow"], "Cash out": [o + i for o, i in zip(res["outflow"], res["interest"])],
                   "Closing balance": res["closing"]}).iloc[:show]

def line(cols, title):
    d = df.melt("Month", cols, "Series", "INR")
    return alt.Chart(d, title=title).mark_line(point=True).encode(
        x="Month:O", y=alt.Y("INR:Q", axis=alt.Axis(format=",.0f")), color="Series:N",
        tooltip=["Month", "Series", alt.Tooltip("INR:Q", format=",.0f")])

st.subheader("3. Charts")
a, b = st.columns(2)
a.altair_chart(line(["Cumulative value of work", "Cumulative cost incurred", "Cumulative cash in", "Cumulative cash out"],
                    "S-curves: cumulative value, cost and cash"), width="stretch")
bars = df.melt("Month", ["Cash in", "Cash out"], "Series", "INR")
b.altair_chart(alt.Chart(bars, title="Monthly cash in vs cash out").mark_bar().encode(
    x="Month:O", xOffset="Series:N", y=alt.Y("INR:Q", axis=alt.Axis(format=",.0f")), color="Series:N",
    tooltip=["Month", "Series", alt.Tooltip("INR:Q", format=",.0f")]), width="stretch")
bal = alt.Chart(df, title="Closing cash balance vs overdraft limit").mark_area(opacity=0.5).encode(
    x="Month:O", y=alt.Y("Closing balance:Q", title="INR (red line = overdraft limit)", axis=alt.Axis(format=",.0f")))
limit = alt.Chart(pd.DataFrame({"y": [-lim]})).mark_rule(color="red", strokeDash=[6, 4]).encode(y="y:Q")
a.altair_chart(bal + limit, width="stretch")

sens = []
for k in range(0, 5):
    r = E.run(replace(inp, payment_lag=k))["summary"]
    sens.append({"Payment lag (months)": k, "Peak funding": r["peak_funding"], "Overdraft interest": r["total_interest"]})
sdf = pd.DataFrame(sens)
b.altair_chart(alt.Chart(sdf, title="Payment-delay sensitivity: peak funding").mark_bar().encode(
    x="Payment lag (months):O", y=alt.Y("Peak funding:Q", axis=alt.Axis(format=",.0f")),
    tooltip=[alt.Tooltip("Peak funding:Q", format=",.0f"), alt.Tooltip("Overdraft interest:Q", format=",.0f")]),
    width="stretch")

st.subheader("4. Scenario comparison (preset scenarios)")
rows = []
for nm in ["Base", "Payment delay", "Stress"]:
    r = E.run(replace(presets[nm], activities=acts))["summary"]
    rows.append({"Scenario": nm, "Payment lag": presets[nm].payment_lag, "Peak funding (INR)": f"{r['peak_funding']:,.0f}",
                 "Peak month": r["peak_month"], "Interest (INR)": f"{r['total_interest']:,.0f}",
                 "Interest % of margin": f"{r['interest_share_of_margin']:.1%}", "Months above limit": r["months_over_limit"]})
st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

st.subheader("5. Monthly cash flow table")
tbl = pd.DataFrame({
    "Work done %": [f"{x:.1%}" for x in res["work_share"]], "Gross bill": res["gross_bill"], "Retention": res["retention"],
    "Advance recovery": res["advance_recovery"], "TDS": res["tds"], "Net bill": res["net_bill"],
    "Client receipt": res["receipts"], "Advance in": res["advance_in"], "Retention released": res["retention_release"],
    "Cash in": res["inflow"], "Cash out": res["outflow"], "Interest": res["interest"], "Closing balance": res["closing"],
}, index=[f"M{m}" for m in months]).iloc[:show]
st.dataframe(tbl.style.format({c: "{:,.0f}" for c in tbl.columns if c != "Work done %"}), width="stretch")
st.download_button("Download monthly table (CSV)", tbl.to_csv().encode(), "cash_flow_forecast.csv", "text/csv")

st.caption("Method: see the user manual and method note. Planning tool only — it does not replace contract review, "
           "lender approval or certified accounts. Not modelled: GST timing, interest on mobilisation advance, escalation, bank fees.")
