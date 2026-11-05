# Contractor Cash Flow Forecaster – Streamlit app (CE725 Team 1, Topic 3)

The same method as the Excel tool (`../Team1_Cash_Flow_Forecaster_v2.xlsx`), as an interactive web app.

**Live app (no installation): https://ce725-cashflow-forecaster.streamlit.app**
If it shows a "This app has gone to sleep" screen, click the wake-up button and wait about 30 seconds.

Run locally (below) to keep all data on your own computer. The hosted version processes inputs on Streamlit Community Cloud for the session only and saves nothing.

## Run it (about 2 minutes)
Requires **Python 3.10 or later** (`python3 --version`).

**macOS / Linux**
```bash
cd ce725-cashflow-forecaster
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

**Windows (PowerShell)**
```powershell
cd ce725-cashflow-forecaster
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

The app opens at http://localhost:8501. Press `Ctrl+C` in the terminal to stop it.

**Optional – host it online:** push this folder to a GitHub repository and deploy it on Streamlit Community Cloud, with `app.py` as the main file. No code changes are needed.

## What to try
1. **Start from scenario** (sidebar): Base, Payment delay or Stress.
2. Move the **Payment lag** slider from 1 to 4 months: peak funding, interest and every chart update.
3. Edit the **Project schedule** table: change start months, durations or shares. Shares must total 100%; the app warns you if they don't.
4. Read **Key results**, the **S-curves**, **cash in vs cash out**, **closing balance vs overdraft limit** and the **payment-delay sensitivity** chart.
5. Download the monthly table as a CSV.

Base case results you should see: peak funding **₹229,817 in month 6**, overdraft interest **₹4,115**.

## Files
| File | Purpose |
|---|---|
| `app.py` | Streamlit user interface |
| `cashflow_engine.py` | Calculation engine (pure Python, no dependencies); run `python cashflow_engine.py` for a text summary |
| `test_validation.py` | Compares every value in the Excel workbook with this engine (the workbook must have been saved by Excel) |
| `test_excel_edge_cases.py` | Runs edge cases through Microsoft Excel (macOS only, uses AppleScript) |
| `build_workbook.py`, `build_benchmark.py`, `benchmark_values.py` | Scripts that generated the Excel tool and the benchmark workbook |
| `requirements.txt` | Python packages |

## Method (short)
Activity schedule → monthly value of work → running bills less retention, capped advance recovery and TDS → receipts after the payment lag → costs paid after each cost type's credit period → monthly cash balance with overdraft interest (opening negative balance × rate ÷ 12, compounding monthly) → peak funding, interest, months over the overdraft limit. The full method is in the user manual and the method note.
