# SEC-Powered DCF & Comps Pack (TDOC, LH, DGX)

This project builds a **professional DCF + Comps Excel model** using real data:
- **SEC XBRL (EDGAR)** for fundamentals  
- **FRED** for 10-Year Treasury risk-free rate  
- **Yahoo Finance** for live prices  

It automatically creates:
- Cleaned financials in `/data_proc`
- A valuation Excel file (`model/valuation_pack.xlsx`)
- Per-company DCF tabs (TDOC, LH, DGX) with editable assumptions
- Sensitivity grids for WACC × g
- A Valuation Summary comparing implied vs market price

## How to Run
1. Create a `.env` file with:
USER_AGENT="Kane Lee (kane@example.com)"
2. In PowerShell:
```powershell
py -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python src\build_comps_and_model.py
python src\build_dcf_per_company.py
start model\valuation_pack.xlsx
