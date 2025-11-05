import time, json, pathlib, requests
from config import CIK_MAP, HEADERS

BASE = "https://data.sec.gov/api/xbrl/companyfacts/CIK{}.json"
OUTDIR = pathlib.Path("data_raw")
OUTDIR.mkdir(parents=True, exist_ok=True)

def get_companyfacts(cik):
    url = BASE.format(cik)
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()

def main():
    for tkr, cik in CIK_MAP.items():
        data = get_companyfacts(cik)
        with open(OUTDIR / f"{tkr}_companyfacts.json", "w") as f:
            json.dump(data, f)
        print(f"saved: data_raw/{tkr}_companyfacts.json  (facts: {len(data.get('facts',{}))})")
        time.sleep(0.2)

if __name__ == "__main__":
    main() 