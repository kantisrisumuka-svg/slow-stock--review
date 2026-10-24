import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Company X Slow-Stock Review", page_icon="📦", layout="wide")
st.title("Company X | Slow-Stock Review")
st.caption("Upload the monthly exports, review data-quality flags, identify Push items and export the results.")

ITEM_COLS = ["SKU", "Item Name", "Category", "Master Pack Qty", "Unit Cost (Rs)", "Alternate Item (name)", "Alternate Flag", "Status"]
STOCK_COLS = ["SKU", "Qty 0-30 days", "Qty 31-60 days", "Qty 61-90 days", "Qty 91-180 days", "Qty 180+ days"]
SALES_COLS = ["Month", "CustomerID", "CustomerName", "Industry", "SKU", "Qty", "Value (Rs)"]
CUSTOMER_COLS = ["CustomerID", "CustomerName", "Industry", "Area"]
AGE_COLS = STOCK_COLS[1:]

def clean_text(s):
    return s.astype("string").str.strip()

def read_uploaded_files(files):
    """Accept either one workbook with named sheets or four separate Excel files."""
    frames = {}
    if len(files) == 1:
        xls = pd.ExcelFile(files[0])
        mapping = {
            "Item_Master": ("items", ITEM_COLS),
            "Stock_Aging": ("stock", STOCK_COLS),
            "Sales_Lines": ("sales", SALES_COLS),
            "Customers": ("customers", CUSTOMER_COLS),
        }
        for sheet, (key, _) in mapping.items():
            if sheet not in xls.sheet_names:
                raise ValueError(f"Workbook is missing required sheet: {sheet}")
            frames[key] = pd.read_excel(xls, sheet_name=sheet)
    else:
        for f in files:
            name = f.name.lower()
            if "item" in name and "master" in name:
                frames["items"] = pd.read_excel(f)
            elif "stock" in name or "aging" in name:
                frames["stock"] = pd.read_excel(f)
            elif "sales" in name:
                frames["sales"] = pd.read_excel(f)
            elif "customer" in name:
                frames["customers"] = pd.read_excel(f)
        missing = {"items", "stock", "sales", "customers"} - set(frames)
        if missing:
            raise ValueError("Could not identify these files by filename: " + ", ".join(sorted(missing)))
    for key, required in [("items", ITEM_COLS), ("stock", STOCK_COLS), ("sales", SALES_COLS), ("customers", CUSTOMER_COLS)]:
        missing = set(required) - set(frames[key].columns)
        if missing:
            raise ValueError(f"{key} file is missing columns: {', '.join(sorted(missing))}")
    return frames

def run_review(frames):
    items, stock, sales, customers = [frames[k].copy() for k in ("items", "stock", "sales", "customers")]
    issues = []

    # Normalize key text to reduce false mismatches from leading/trailing spaces.
    for df in (items, stock, sales):
        df["SKU"] = df["SKU"].astype("string").str.strip().str.upper()
    for df in (sales, customers):
        df["CustomerID"] = df["CustomerID"].astype("string").str.strip()
    for df in (items, customers):
        for c in df.select_dtypes(include=["object", "string"]).columns:
            df[c] = df[c].astype("string").str.strip()

    # Duplicate master and stock keys are unsafe to merge automatically.
    if items["SKU"].duplicated().any():
        raise ValueError("Item_Master contains duplicate SKU keys. Resolve these before analysis.")
    if stock["SKU"].duplicated().any():
        raise ValueError("Stock_Aging contains duplicate SKU keys. Resolve these before analysis.")
    if customers["CustomerID"].duplicated().any():
        raise ValueError("Customers contains duplicate CustomerID keys. Resolve these before analysis.")

    # Coerce numeric fields and flag invalid values.
    numeric_item = ["Master Pack Qty", "Unit Cost (Rs)"]
    for c in numeric_item:
        items[c] = pd.to_numeric(items[c], errors="coerce")
    for c in AGE_COLS:
        stock[c] = pd.to_numeric(stock[c], errors="coerce")
    for c in ["Qty", "Value (Rs)"]:
        sales[c] = pd.to_numeric(sales[c], errors="coerce")
    if items[numeric_item].isna().any().any():
        issues.append("Missing/non-numeric pack quantity or unit cost: affected items are excluded from calculation; fix Item_Master at source.")
    if stock[AGE_COLS].isna().any().any():
        issues.append("Missing/non-numeric stock bucket: treated as zero for calculation and flagged; confirm with warehouse before acting.")
        stock[AGE_COLS] = stock[AGE_COLS].fillna(0)
    if sales[["Qty", "Value (Rs)"]].isna().any().any():
        issues.append("Missing/non-numeric sales quantity or value: affected rows are excluded; correct the source export.")
        sales = sales.dropna(subset=["Qty", "Value (Rs)"])

    # Parse the six-month window robustly, including YYYY-MM and Mon-YYYY formats.
    raw_month = sales["Month"].astype("string").str.strip()
    month = pd.to_datetime(raw_month, format="%Y-%m", errors="coerce")
    month = month.fillna(pd.to_datetime(raw_month, format="%b-%Y", errors="coerce"))
    sales["MonthParsed"] = month
    if sales["MonthParsed"].isna().any():
        bad = int(sales["MonthParsed"].isna().sum())
        issues.append(f"{bad} sales rows have unrecognized month formats and were excluded; standardize Month as YYYY-MM at source.")
        sales = sales.dropna(subset=["MonthParsed"])
    months = sorted(sales["MonthParsed"].dt.to_period("M").astype(str).unique())
    if len(months) != 6:
        issues.append(f"Expected 6 distinct sales months but found {len(months)} ({', '.join(months)}); verify the reporting period.")

    # Exact duplicate transaction rows: flag and remove identical records only.
    dedup_cols = ["MonthParsed", "CustomerID", "CustomerName", "Industry", "SKU", "Qty", "Value (Rs)"]
    duplicate_count = int(sales.duplicated(subset=dedup_cols).sum())
    if duplicate_count:
        issues.append(f"Removed {duplicate_count} exact duplicate sales rows. Confirm whether these are duplicate exports or legitimate repeated transactions; add a unique transaction-line ID at source.")
        sales = sales.drop_duplicates(subset=dedup_cols)

    # SKU values present in sales/stock but absent from master are reported.
    known_skus = set(items["SKU"].dropna())
    unknown_sales = sorted(set(sales["SKU"].dropna()) - known_skus)
    unknown_stock = sorted(set(stock["SKU"].dropna()) - known_skus)
    if unknown_sales:
        issues.append("Sales contain SKU(s) not found in Item_Master: " + ", ".join(unknown_sales) + ". Excluded from item-level analysis; correct SKU mapping.")
    if unknown_stock:
        issues.append("Stock contains SKU(s) not found in Item_Master: " + ", ".join(unknown_stock) + ". Excluded from item-level analysis; correct SKU mapping.")

    # Inactive stock alert. Inactive items are not included in classification.
    stock["StockTotalBeforeFix"] = stock[AGE_COLS].fillna(0).sum(axis=1)
    inactive = items[items["Status"].astype("string").str.lower().eq("inactive")]
    inactive_stock = inactive.merge(stock[["SKU", "StockTotalBeforeFix"]], on="SKU", how="left").fillna({"StockTotalBeforeFix": 0})
    inactive_stock = inactive_stock[inactive_stock["StockTotalBeforeFix"] > 0]
    for _, r in inactive_stock.iterrows():
        issues.append(f"Inactive item {r['SKU']} still holds {r['StockTotalBeforeFix']:g} units; excluded from analysis, but warehouse should confirm disposition.")

    # Negative stock values are invalid for straightforward ageing calculations. Clamp only for this run and disclose.
    neg_mask = stock[AGE_COLS] < 0
    if neg_mask.any().any():
        details = []
        for _, r in stock.loc[neg_mask.any(axis=1)].iterrows():
            for c in AGE_COLS:
                if pd.notna(r[c]) and r[c] < 0:
                    details.append(f"{r['SKU']} / {c}: {r[c]:g}")
        issues.append("Negative stock bucket(s) set to zero for this run: " + "; ".join(details) + ". Do not treat this as a stock adjustment; reconcile warehouse transactions and correct source data.")
        stock[AGE_COLS] = stock[AGE_COLS].clip(lower=0)

    # Resolve confirmed SWITCH links using alternate item name, with whitespace normalization.
    name_to_skus = {}
    for _, r in items.iterrows():
        name_to_skus.setdefault(re.sub(r"\s+", " ", str(r["Item Name"]).strip()).casefold(), []).append(r["SKU"])
    switch_map = {}
    switch_source_rows = items[items["Alternate Flag"].astype("string").str.strip().str.upper().eq("SWITCH")]
    for _, r in switch_source_rows.iterrows():
        alt_name = re.sub(r"\s+", " ", str(r["Alternate Item (name)"]).strip()).casefold()
        matches = name_to_skus.get(alt_name, [])
        if len(matches) == 1:
            switch_map[r["SKU"]] = matches[0]
        else:
            issues.append(f"Switch mapping for {r['SKU']} is ambiguous or missing ({r['Alternate Item (name)']}); source item left unmapped pending correction.")
    if switch_map:
        issues.append("Combined confirmed pack-size switches under the new SKU: " + ", ".join(f"{old} → {new}" for old, new in switch_map.items()) + ".")
    # Alternate names with no SWITCH flag are not merged automatically.
    flagged_names = items["Alternate Item (name)"].notna() & items["Alternate Item (name)"].astype("string").str.strip().ne("")
    no_flag = items[flagged_names & ~items["Alternate Flag"].astype("string").str.upper().eq("SWITCH").fillna(False)]
    if len(no_flag):
        issues.append("Alternate item name exists without SWITCH flag or with another flag for: " + ", ".join(no_flag["SKU"].tolist()) + ". No merge was assumed; confirm intended relationship at source.")

    def canonical(sku):
        return switch_map.get(sku, sku)
    stock["CanonicalSKU"] = stock["SKU"].map(canonical)
    sales["CanonicalSKU"] = sales["SKU"].map(canonical)
    stock_agg = stock.groupby("CanonicalSKU")[AGE_COLS].sum()
    sales_agg = sales.groupby("CanonicalSKU").agg(SalesQty=("Qty", "sum"), SalesValue=("Value (Rs)", "sum"))
    item_lookup = items.set_index("SKU")
    active = items[items["Status"].astype("string").str.lower().eq("active")].copy()
    active = active[~active["SKU"].isin(switch_map.keys())].copy()

    # Sales metadata reconciled to customer master; customer master is authoritative for industry.
    sales = sales.merge(customers[["CustomerID", "Industry"]].rename(columns={"Industry": "MasterIndustry"}),
                        on="CustomerID", how="left", suffixes=("", "_master"))
    missing_customer = sales[sales["MasterIndustry"].isna()]
    if len(missing_customer):
        issues.append(f"{missing_customer['CustomerID'].nunique()} customer ID(s) in sales are missing from Customers; customer targeting for those IDs is limited.")
    if customers["Industry"].isna().any() or customers["Industry"].astype("string").str.strip().eq("").any():
        issues.append("One or more customers have no industry. They are not used as an industry-match basis for Bucket 2; complete the Customers master.")
    sales["CanonicalSKU"] = sales["SKU"].map(canonical)
    sales["Category"] = sales["CanonicalSKU"].map(item_lookup["Category"].to_dict())
    sales["IndustryForTarget"] = sales["MasterIndustry"]

    cust_value = sales.groupby("CustomerID")["Value (Rs)"].sum()
    eligible = set(cust_value[cust_value > 100000].index)
    customer_meta = customers.set_index("CustomerID")
    customer_item = sales.groupby(["CustomerID", "CanonicalSKU"])["Qty"].sum()
    customer_category = sales[sales["Category"].notna()].groupby(["CustomerID", "Category"])["Qty"].sum()
    result_rows = []
    for _, item in active.iterrows():
        sku = item["SKU"]
        stock_row = stock_agg.loc[sku] if sku in stock_agg.index else pd.Series(0, index=AGE_COLS)
        sales_row = sales_agg.loc[sku] if sku in sales_agg.index else pd.Series({"SalesQty": 0, "SalesValue": 0})
        total_stock = float(stock_row.sum())
        aged_qty = float(stock_row["Qty 91-180 days"] + stock_row["Qty 180+ days"])
        aged_value = aged_qty * float(item["Unit Cost (Rs)"])
        sales_qty = float(sales_row["SalesQty"])
        avg_monthly = sales_qty / 6
        months_clear = total_stock / avg_monthly if avg_monthly > 0 else float("inf")
        if sales_qty < float(item["Master Pack Qty"]):
            classification = "Red List"
        elif aged_qty == 0 or months_clear < 3:
            classification = "No Action"
        else:
            classification = "Push"
        result_rows.append({
            "SKU": sku, "Item Name": item["Item Name"], "Category": item["Category"],
            "Master Pack Qty": item["Master Pack Qty"], "Unit Cost (Rs)": item["Unit Cost (Rs)"],
            "Total Stock": total_stock, "Aged Qty": aged_qty, "Aged Value (Rs)": aged_value,
            "6-Month Sales Qty": sales_qty, "Average Monthly Sales": avg_monthly,
            "Months to Clear": months_clear, "Class": classification
        })
    results = pd.DataFrame(result_rows)

    # Target list: only for Push items and customers with > ₹1,00,000 total purchase value.
    target_rows = []
    cat_map = item_lookup["Category"].to_dict()
    for _, item in results[results["Class"].eq("Push")].iterrows():
        sku, category = item["SKU"], item["Category"]
        item_buyers = {cid for cid in eligible if customer_item.get((cid, sku), 0) > 0}
        buyer_industries = set()
        for cid in item_buyers:
            ind = customer_meta.loc[cid, "Industry"] if cid in customer_meta.index else pd.NA
            if pd.notna(ind) and str(ind).strip():
                buyer_industries.add(str(ind).strip())
        for cid in eligible:
            if cid in item_buyers:
                bucket = 1
            else:
                ind = customer_meta.loc[cid, "Industry"] if cid in customer_meta.index else pd.NA
                if pd.notna(ind) and str(ind).strip() and str(ind).strip() in buyer_industries:
                    bucket = 2
                elif customer_category.get((cid, category), 0) > 0:
                    bucket = 3
                else:
                    continue
            meta = customer_meta.loc[cid] if cid in customer_meta.index else {}
            target_rows.append({
                "SKU": sku, "CustomerID": cid,
                "CustomerName": meta.get("CustomerName", cid),
                "Industry": meta.get("Industry", ""),
                "Area": meta.get("Area", ""),
                "Bucket": bucket,
                "6-Month Customer Purchase Value (Rs)": float(cust_value.get(cid, 0))
            })
    targets = pd.DataFrame(target_rows)
    if not targets.empty:
        targets = targets.sort_values(["SKU", "Bucket", "6-Month Customer Purchase Value (Rs)"], ascending=[True, True, False])

    return results, targets, issues, inactive_stock, len(sales), len(items), len(active)

st.markdown("### 1. Upload monthly files")
uploaded = st.file_uploader(
    "Upload one combined workbook, or the four separate .xlsx files",
    type=["xlsx"], accept_multiple_files=True,
    help="For separate files, filenames should contain Item_Master, Stock_Aging, Sales_Lines and Customers."
)
if uploaded:
    try:
        frames = read_uploaded_files(uploaded)
        results, targets, issues, inactive_stock, sales_rows_used, master_count, active_count = run_review(frames)
        st.success("Review completed. Check data-quality notes before acting on the list.")
        with st.expander(f"Data-quality checks ({len(issues)} notes)", expanded=True):
            if issues:
                for issue in issues:
                    st.warning(issue)
            else:
                st.write("No issues detected by the current checks.")
        st.markdown("### 2. Summary")
        a, b, c, d = st.columns(4)
        a.metric("Red List", int((results["Class"] == "Red List").sum()))
        b.metric("No Action", int((results["Class"] == "No Action").sum()))
        c.metric("Push", int((results["Class"] == "Push").sum()))
        d.metric("Push aged value", f"₹{results.loc[results['Class']=='Push', 'Aged Value (Rs)'].sum():,.0f}")
        st.markdown("#### Top 10 Push items by aged value")
        st.dataframe(results[results["Class"] == "Push"].sort_values("Aged Value (Rs)", ascending=False).head(10), use_container_width=True, hide_index=True)

        st.markdown("### 3. Item list")
        col1, col2 = st.columns(2)
        classes = ["All"] + sorted(results["Class"].dropna().unique().tolist())
        categories = ["All"] + sorted(results["Category"].dropna().unique().tolist())
        chosen_class = col1.selectbox("Filter by class", classes)
        chosen_category = col2.selectbox("Filter by category", categories)
        filtered = results.copy()
        if chosen_class != "All":
            filtered = filtered[filtered["Class"] == chosen_class]
        if chosen_category != "All":
            filtered = filtered[filtered["Category"] == chosen_category]
        st.dataframe(filtered.sort_values(["Class", "Aged Value (Rs)"], ascending=[True, False]), use_container_width=True, hide_index=True)

        st.markdown("### 4. Customer call list")
        push_skus = sorted(results.loc[results["Class"] == "Push", "SKU"].tolist())
        if push_skus:
            selected_sku = st.selectbox("Choose a Push SKU", push_skus)
            selected_targets = targets[targets["SKU"] == selected_sku].copy() if not targets.empty else pd.DataFrame()
            if selected_targets.empty:
                st.info("No eligible customer targets for this item under the current rules.")
            else:
                selected_targets["Bucket"] = selected_targets["Bucket"].map({1: "1 – Bought before", 2: "2 – Similar business", 3: "3 – Same category"})
                st.dataframe(selected_targets, use_container_width=True, hide_index=True)
                # A mobile-friendly short call list.
                for _, r in selected_targets.head(10).iterrows():
                    st.markdown(f"**{r['CustomerName']}** · {r['Area']} · {r['Bucket']}  \nIndustry: {r['Industry']} · 6-month purchases: ₹{r['6-Month Customer Purchase Value (Rs)']:,.0f}")
        else:
            st.info("There are no Push items in this run.")

        st.markdown("### 5. Export")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            results.to_excel(writer, sheet_name="Item_Review", index=False)
            targets.to_excel(writer, sheet_name="Customer_Targets", index=False)
            pd.DataFrame({"Data Quality Notes": issues or ["No issues detected by current checks."]}).to_excel(writer, sheet_name="Data_Quality", index=False)
            inactive_stock.to_excel(writer, sheet_name="Inactive_Stock", index=False)
        st.download_button("Download review results (.xlsx)", data=output.getvalue(),
                           file_name="CompanyX_Slow_Stock_Review.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Reminder: confirm data-quality warnings with Supply Chain before taking action. This tool is decision support, not a substitute for stock reconciliation.")
    except Exception as e:
        st.error(f"Could not run the review: {e}")
        st.info("Check that you uploaded the combined workbook with the four required sheet names, or all four separate files.")
else:
    st.info("Upload the combined workbook or all four separate files to begin.")
