import json, pandas as pd
from pathlib import Path

TAG_MAP = {
    # Revenue / Sales
    "revenue": [
        "Revenues",
        "SalesRevenueNet",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "SalesRevenueGoodsNet", "SalesRevenueServicesNet"
    ],

    # Cost of goods / cost of revenue
    "cogs": [
        "CostOfRevenue",
        "CostOfGoodsAndServicesSold",
        "CostOfSales"
    ],

    # EBIT / operating income (more fallbacks)
    "ebit": [
        "OperatingIncomeLoss",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxes",
        # as last resort, GrossProfit - SG&A (we won’t compute here; keep as API-driven)
    ],

    # Depreciation & Amortization (many filers split it)
    "da": [
        "DepreciationAndAmortization",
        "DepreciationDepletionAndAmortization",
        "AmortizationOfIntangibleAssets",
        "Depreciation",
        "AmortizationOfFiniteLivedIntangibleAssets"
    ],

    # Cash from operations
    "cfo": [
        "NetCashProvidedByUsedInOperatingActivities",
        "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations",
    ],

    # CapEx
    "capex": [
        "PaymentsToAcquirePropertyPlantAndEquipment",
        "CapitalExpenditures",
        "PaymentsForProceedsFromProductiveAssets",
        "PurchaseOfPropertyPlantAndEquipment"
    ],

    # Cash
    "cash": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments"
    ],

    # Debt (try combined then components)
    "debt": [
        "LongTermDebtAndCapitalLeaseObligations",
        "LongTermDebt",
        "DebtCurrent",
        "ShortTermBorrowings",
        "LongTermDebtNoncurrent",
        "DebtAndCapitalLeaseObligations"
    ],

    # Shares (prefer diluted average; fallback to point-in-time SO)
    "diluted_shares": [
        "WeightedAverageNumberOfDilutedSharesOutstanding",
        "WeightedAverageNumberOfShareDilutedOutstanding",
        "CommonStockSharesOutstanding"
    ]
}


def _latest_per_fy(values):
    rows=[]
    for v in values:
        fy=v.get("fy"); fp=v.get("fp"); end=v.get("end"); form=v.get("form")
        if fy and form in ("10-K","10-Q"):
            rows.append({"fy":int(fy),"fp":fp,"end":end,"val":v.get("val")})
    if not rows: return pd.DataFrame(columns=["fy","val"])
    df=pd.DataFrame(rows).sort_values(["fy","end"])
    return df.drop_duplicates("fy", keep="last")[["fy","val"]]

def _extract(path):
    j=json.load(open(path))
    facts=j.get("facts",{}).get("us-gaap",{})
    frames=[]
    for line_item,tags in TAG_MAP.items():
        found=None
        for tag in tags:
            if tag in facts:
                units=facts[tag].get("units",{})
                if "USD" in units:
                    found=_latest_per_fy(units["USD"]).rename(columns={"val":line_item})
                    break
                if "shares" in units:
                    found=_latest_per_fy(units["shares"]).rename(columns={"val":line_item})
                    break
        if found is not None and not found.empty:
            frames.append(found)
    if not frames: return pd.DataFrame()
    df=frames[0]
    for f in frames[1:]:
        df=df.merge(f,on="fy",how="outer")
    return df.sort_values("fy")

def main():
    out=[]
    for p in Path("data_raw").glob("*_companyfacts.json"):
        tkr=p.name.split("_")[0]
        df=_extract(p)
        if df.empty: 
            print(f"[WARN] no data extracted for {tkr}")
            continue
        df.insert(0,"ticker",tkr)
        out.append(df)
    if not out:
        print("[ERR] No companyfacts JSON found. Run pull_sec_companyfacts.py first.")
        return
    allf=pd.concat(out, ignore_index=True)
    Path("data_proc").mkdir(exist_ok=True)
    allf.to_csv("data_proc/financials_tidy.csv",index=False)
    print("saved data_proc/financials_tidy.csv")
    print(allf.tail(6))

if __name__=="__main__":
    main()
