import os
from dotenv import load_dotenv
load_dotenv()

USER_AGENT = os.getenv("USER_AGENT")
FRED_API_KEY = os.getenv("FRED_API_KEY")

if not USER_AGENT:
    raise RuntimeError("Missing USER_AGENT. Set it in your .env file.")
if not FRED_API_KEY:
    print("[WARN] FRED_API_KEY missing — required for the FRED script later.")

HEADERS = {"User-Agent": USER_AGENT}

TICKERS = ["TDOC", "LH", "DGX"]

CIK_MAP = {
    "TDOC": "0001477449",
    "LH": "0000920148",
    "DGX": "0001022079",
}
