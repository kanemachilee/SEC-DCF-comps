import pandas as pd, requests
import yfinance as yf
from config import TICKERS, FRED_API_KEY

# Prices (10y, adjusted)
px = yf.download(TICKERS, period="10y", interval="1d", auto_adjust=True, progress=False)
px.to_csv("data_raw/prices_10y.csv")

# Risk-free (FRED DGS10)
if not FRED_API_KEY:
    raise RuntimeError("FRED_API_KEY missing in .env")
url = f"https://api.stlouisfed.org/fred/series/observations?series_id=DGS10&file_type=json&api_key={FRED_API_KEY}"
obs = requests.get(url, timeout=30).json()["observations"]
rf = (pd.DataFrame(obs)[["date","value"]]
      .assign(value=lambda d: pd.to_numeric(d["value"], errors="coerce"))
      .dropna())
rf.to_csv("data_raw/fred_dgs10.csv", index=False)

# Snapshots for Excel assumptions
latest_rf = rf.tail(1).iloc[0]
latest_close = (px["Close"].tail(1).T
                .reset_index()
                .rename(columns={"index":"ticker", px["Close"].tail(1).index[0]:"last_price"}))
latest_close.to_csv("data_proc/latest_prices.csv", index=False)
pd.DataFrame([{"date": latest_rf["date"], "rf_10y_pct": float(latest_rf["value"])}]).to_csv("data_proc/latest_rf.csv", index=False)

print("saved data_raw/prices_10y.csv, data_raw/fred_dgs10.csv, data_proc/latest_prices.csv, data_proc/latest_rf.csv")
