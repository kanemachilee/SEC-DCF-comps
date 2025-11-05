# src/rebuild_latest_prices.py
import pandas as pd
import yfinance as yf
from pathlib import Path
from config import TICKERS

IN = Path("data_raw/prices_10y.csv")
OUT = Path("data_proc/latest_prices.csv")

def from_prices_csv(path: Path):
    """
    Try to extract last prices from data_raw/prices_10y.csv,
    handling both MultiIndex (['Close', ticker]) and wide formats ('TDOC Close', etc).
    Returns a DataFrame with columns: ['ticker','last_price'] or None if not possible.
    """
    try:
        df = pd.read_csv(path, header=[0,1])
        # MultiIndex like ('Close','TDOC'), ('Close','LH'), ...
        if isinstance(df.columns, pd.MultiIndex) and "Close" in df.columns.get_level_values(0):
            last = df["Close"].tail(1).T.reset_index()
            last.columns = ["ticker", "last_price"]
            return last
    except Exception:
        pass

    # Try single header; columns may look like "TDOC Close", "LH Close", ...
    try:
        df = pd.read_csv(path)
        close_cols = [c for c in df.columns if c.endswith(" Close")]
        if close_cols:
            last_row = df.tail(1).iloc[0]
            rows = []
            for c in close_cols:
                tkr = c.replace(" Close","").strip()
                rows.append({"ticker": tkr, "last_price": float(last_row[c])})
            return pd.DataFrame(rows)
    except Exception:
        pass
    return None

def fetch_fallback(tickers):
    """If csv parsing fails, fetch current last close via yfinance per ticker."""
    data = []
    for t in tickers:
        try:
            h = yf.download(t, period="5d", interval="1d", auto_adjust=True, progress=False, threads=False)
            if not h.empty:
                data.append({"ticker": t, "last_price": float(h["Close"].iloc[-1])})
        except Exception as e:
            print(f"[WARN] {t}: fallback fetch failed: {e}")
    return pd.DataFrame(data)

def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    tidy = from_prices_csv(IN)
    if tidy is None or tidy.empty:
        print("[INFO] Could not parse prices_10y.csv reliably; using fallback fetch.")
        tidy = fetch_fallback(TICKERS)
    if tidy.empty or "ticker" not in tidy.columns or "last_price" not in tidy.columns:
        raise SystemExit("[ERR] Could not build latest_prices.csv")
    # Keep only our tickers
    tidy = tidy[tidy["ticker"].isin(TICKERS)].dropna()
    tidy.to_csv(OUT, index=False)
    print("Rebuilt", OUT)

if __name__ == "__main__":
    main()
