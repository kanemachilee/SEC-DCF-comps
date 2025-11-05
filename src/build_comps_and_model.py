# src/build_comps_and_model.py
import math
import pandas as pd
from pathlib import Path
from openpyxl import Workbook

def coerce_num(s):
    return pd.to_numeric(s, errors="coerce")

# ---- Load inputs (robust) ----
fin = pd.read_csv("data_proc/financials_tidy.csv", encoding="utf-8-sig")
latest_px = pd.read_csv("data_proc/latest_prices.csv", encoding="utf-8-sig")

# Normalize tickers to UPPER and strip spaces
if "ticker" not in latest_px.columns:
    raise SystemExit("[ERR] latest_prices.csv missing 'ticker' column")
latest_px["ticker"] = latest_px["ticker"].astype(str).str.strip().str.upper()
fin["ticker"] = fin["ticker"].astype(str).str.strip().str.upper()

# Prefer last FY per ticker
lastfy = (fin.sort_values(["ticker","fy"])
            .groupby("ticker").tail(1)
            .set_index("ticker"))

# Pull fields (numeric coercion + safe defaults)
for col in ["diluted_shares","cash","debt","revenue","ebit","da"]:
    if col in lastfy.columns:
        lastfy[col] = coerce_num(lastfy[col])
    else:
        lastfy[col] = pd.NA

px = latest_px.set_index("ticker")["last_price"]
px = coerce_num(px)

shares = lastfy.get("diluted_shares")
cash   = lastfy.get("cash").fillna(0)
debt   = lastfy.get("debt").fillna(0)
revenue= lastfy.get("revenue")
ebit   = lastfy.get("ebit")
da     = lastfy.get("da")

# --- Share sanity: if value looks like "in millions", scale to units ---
def scale_shares_if_needed(s):
    s2 = s.copy()
    # If typical is < 10,000, assume it's in millions and scale
    med = coerce_num(s2).median(skipna=True)
    if pd.notna(med) and med < 10_000:
        s2 = s2 * 1_000_000
    return s2

shares = scale_shares_if_needed(shares)

# ---- Build comps ----
equity_value = px * shares
net_debt = debt - cash
ev = equity_value + net_debt

comps = pd.DataFrame({
    "Price": px,
    "DilutedShares": shares,
    "EquityValue": equity_value,
    "NetDebt": net_debt,
    "EV": ev,
    "Revenue (FY)": revenue,
    "EBIT (FY)": ebit,
})

# EV/Revenue (avoid divide-by-zero/NaN)
comps["EV/Revenue"] = comps["EV"] / comps["Revenue (FY)"]

# Rough P/E using EBIT*(1-25% tax)
if ebit is not None:
    comps["P/E (rough)"] = comps["EquityValue"] / (ebit * 0.75)
else:
    comps["P/E (rough)"] = math.nan

# Optional EV/EBITDA using EBIT + D&A if D&A present
if da is not None and ebit is not None:
    comps["EV/EBITDA (rough)"] = comps["EV"] / (ebit + da)
else:
    comps["EV/EBITDA (rough)"] = math.nan

# Clean index/header
comps.index.name = "ticker"
comps = comps.reset_index()
comps["ticker"] = comps["ticker"].astype(str).str.strip().str.upper()
comps = comps.set_index("ticker").sort_index()

# Save comps to CSV
Path("data_proc").mkdir(exist_ok=True)
comps.reset_index().to_csv("data_proc/comps.csv", index=False, encoding="utf-8-sig")

# ---- Minimal Excel model ----
try:
    latest_rf = pd.read_csv("data_proc/latest_rf.csv")
except FileNotFoundError:
    latest_rf = pd.DataFrame([{"date":"", "rf_10y_pct":""}])

wb = Workbook()

# Assumptions sheet
wsA = wb.active
wsA.title = "Assumptions"
rf = "" if latest_rf.empty else float(latest_rf.iloc[-1]["rf_10y_pct"])
rf_date = "" if latest_rf.empty else latest_rf.iloc[-1]["date"]

wsA.append(["Input","Value","Notes"])
wsA.append(["Risk-free (10Y, %)", rf, "From FRED DGS10; update rf_date below"])
wsA.append(["rf_date", rf_date, "Last observation date"])
wsA.append(["ERP (Damodaran, %)", "", "Enter current US implied ERP (Damodaran)"])
wsA.append(["Industry beta (unlevered)", "", "Enter industry unlevered beta (Damodaran)"])
wsA.append(["Tax rate (%)", 25.0, "Base assumption"])
wsA.append(["Terminal growth (%)", 2.5, "Base assumption"])

# Comps sheet
wsC = wb.create_sheet("Comps")
wsC.append(["ticker"] + list(comps.columns))
df_out = comps.reset_index().where(pd.notnull(comps.reset_index()), None)
for _, r in df_out.iterrows():
    wsC.append([r["ticker"]] + [r[col] for col in comps.columns])

# Save workbook
Path("model").mkdir(exist_ok=True)
wb.save("model/valuation_pack.xlsx")
print("Wrote: data_proc/comps.csv and model/valuation_pack.xlsx")
