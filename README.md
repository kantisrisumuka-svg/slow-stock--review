# Company X Slow-Stock Review

A simple Streamlit app for monthly slow-stock review.

## Run locally
1. Install Python 3.10+.
2. In a terminal, run `pip install -r requirements.txt`.
3. Run `streamlit run app.py`.
4. Open the local URL shown in the terminal.

## Uploads
- One workbook with sheets named `Item_Master`, `Stock_Aging`, `Sales_Lines`, `Customers`; or
- Four separate `.xlsx` files whose filenames identify each sheet.

## Important assumptions
- Exact duplicate sales lines are removed and flagged for review.
- SKU whitespace and letter case are normalized.
- Confirmed `SWITCH` pairs are consolidated under the alternate/new SKU.
- Negative stock bucket values are set to zero for the current run and flagged; this is not a warehouse stock adjustment.
- Inactive SKUs are excluded from classifications, with an alert if they hold stock.
- Customer purchase eligibility uses strict `> ₹1,00,000` over the available six-month sales period.
- This is a prototype. Validate results with Supply Chain before production use.

## Free hosting
Push `app.py` and `requirements.txt` to a GitHub repository, then create a Streamlit Community Cloud app connected to that repository. Keep the repository public if required by the hosting setup, and do not include passwords, tokens, or private customer data. Verify the deployed app URL in a private/incognito browser and keep the app active for the requested 30-day period.
