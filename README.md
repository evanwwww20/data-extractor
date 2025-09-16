# Rent Roll & T12 Analyzer (No‑code)

A Streamlit app that lets you drop Excel files and automatically:
- Extract per‑unit **Market Rent** and **Rent** from messy rent rolls.
- Parse **T12** financials to derive **Gross Income**, **Operating Expenses**, and **NOI**.
- Compute **Cap Rate** (NOI / Purchase Price) and **Implied Price** (NOI / Cap Rate).
- Do **rent statistics** (mean & median of leased rents).
- Compute **Annual Gross Rent** (sum of monthly rents × 12).
- Compute **Per Unit Expense** (Annual OpEx / Units) and **Price per Unit** (Price / Units).

## Use online (no installs)
Deploy to Streamlit Community Cloud and upload your Excels in the browser.

## Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Notes
- Handles multi-line headers like "Market \\nRent" and inconsistent columns via fuzzy matching.
- T12 finder scans for labels like "Gross Income", "Operating Expenses", and "NOI" and reads numeric neighbors (right/below), then computes missing values where possible.
- You can override **cap rate**, **purchase price**, and **total units** in the UI to see per-unit metrics instantly.
