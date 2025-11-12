SEC-Powered DCF & Comps Pack (TDOC, LH, DGX)

This project builds a DCF and comparable companies model using official data from the SEC, FRED, and Yahoo Finance  
scripts are written in python on vs, with outputs automatically generated



Project Overview

project goal: automate entire valuation pipeline from raw SEC filings to finished DCF model.

used program (pandas, requests, yfinance, openpyxl), Excel, FRED API, SEC EDGAR XBRL.

for data sourcing:
- SEC EDGAR API: Company financials (`companyfacts` JSON)
- yahoo finance: Market prices
- FRED: 10-Year Treasury yield for risk free rate
- damodaran Online: industry ERP & beta references

to pull financials:
   ```powershell
   python src\pull_sec_companyfacts.py
   python src\normalize_financials.py

py -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python src\build_comps_and_model.py
python src\build_dcf_per_company.py
start model\valuation_pack.xlsx
