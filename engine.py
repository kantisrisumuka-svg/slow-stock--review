"""Slow-stock review engine. Pure pandas, no UI, so it can be tested on its own.

Flow: load -> clean (log every problem) -> merge switch pairs (R1) -> drop inactive (R2)
      -> numbers (R3, R4) -> classify (R5) -> customer targets (R6, R7)
"""
import re
import pandas as pd

AGED_COLS = ["Qty 91-180 days", "Qty 180+ days"]
BUCKET_NAMES = {1: "1 - Bought before", 2: "2 - Similar business", 3: "3 - Same category"}
MIN_SPEND = 100000   # R6: strictly more than Rs 1,00,000
MONTHS = 6           # R4: fixed divisor from the business rule

# Each table is recognised by columns it must have, so users can upload the
# workbook OR four separate files in any order.
SIGNATURES = {
    "Item_Master": {"SKU", "Category", "Master Pack Qty"},
    "Stock_Aging": {"SKU", "Qty 91-180 days"},
    "Sales_Lines": {"Month", "CustomerID", "Qty"},
    "Customers": {"CustomerID", "Area"},
}


def load_tables(files):
    """files: list of (name, file-like). Returns dict of the 4 DataFrames."""
    found = {}
    for name, f in files:
        for sheet, df in pd.read_excel(f, sheet_name=None).items():
            df.columns = [str(c).strip() for c in df.columns]
            for table, sig in SIGNATURES.items():
                if sig <= set(df.columns) and table not in found:
                    found[table] = df
    missing = set(SIGNATURES) - set(found)
    if missing:
        raise ValueError(f"Could not find these tables in the upload: {sorted(missing)}")
    return found


def norm_text(s):
    """Lower-case, trim, collapse repeated spaces - used to match names and SKUs."""
    return re.sub(r"\s+", " ", str(s)).strip().casefold()


def clean(tables):
    """Returns cleaned copies plus an `issues` list of dicts (what, count, handling)."""
    im, sa = tables["Item_Master"].copy(), tables["Stock_Aging"].copy()
    sl, cu = tables["Sales_Lines"].copy(), tables["Customers"].copy()
    issues = []

    def log(sheet, problem, count, handling):
        if count:
            issues.append({"Sheet": sheet, "Problem": problem, "Rows": count, "How it was handled": handling})

    # SKU formatting: strip spaces and upper-case everywhere so joins don't silently fail
    for df in (im, sa, sl):
        df["SKU"] = df["SKU"].astype(str).str.strip().str.upper()
    bad_fmt = (tables["Sales_Lines"]["SKU"].astype(str) != sl["SKU"]).sum()
    log("Sales_Lines", "SKU written in lower case / with trailing space", bad_fmt,
        "Trimmed and upper-cased so they match the item master")

    # Month formats
    raw_m = sl["Month"].astype(str)
    sl["Month"] = pd.to_datetime(raw_m, format="mixed").dt.to_period("M")
    log("Sales_Lines", "Month written in a different format (e.g. 2026-06 vs Apr-2026)",
        (~raw_m.str.match(r"^[A-Za-z]{3}-\d{4}$")).sum(), "Converted all to one month format")
    n_months = sl["Month"].nunique()
    if n_months != MONTHS:
        log("Sales_Lines", f"Data covers {n_months} months, rule R4 assumes {MONTHS}", 1,
            f"Still dividing by {MONTHS} as per the rule - please check the file")

    # Exact duplicate sales rows
    dup = sl.duplicated(keep="first")
    log("Sales_Lines", "Exact duplicate rows (same month, customer, SKU, qty, value)", int(dup.sum()),
        "Kept the first copy and removed the rest (assumed double-posted). Confirm with the data owner")
    sl = sl[~dup]

    # SKUs in sales not in master
    unknown = ~sl["SKU"].isin(im["SKU"])
    log("Sales_Lines", "SKU not found in Item Master", int(unknown.sum()), "Rows excluded from the analysis")
    sl = sl[~unknown]

    # Stock: negative buckets
    qcols = [c for c in sa.columns if c.startswith("Qty")]
    neg = (sa[qcols] < 0).sum().sum()
    log("Stock_Aging", "Negative stock quantity in an age bucket", int(neg),
        "Set to 0 for the calculation (cannot hold negative stock). Likely an unposted transfer or return")
    sa[qcols] = sa[qcols].clip(lower=0)

    # Item master flags that do not follow the rules
    flag = im["Alternate Flag"].fillna("").astype(str).str.strip().str.upper()
    im["Alternate Flag"] = flag
    has_alt = im["Alternate Item (name)"].notna()
    log("Item_Master", "Has an alternate item name but the flag is blank (possible missing SWITCH flag)",
        int((has_alt & (flag == "")).sum()),
        "Treated as separate items because R1 applies only when the flag is SWITCH")
    log("Item_Master", "Flag value 'EXCLUDE' is not defined in the rules",
        int((flag == "EXCLUDE").sum()), "Treated as a normal item (not merged). Ask what EXCLUDE should mean")
    odd_name = im["Alternate Item (name)"].dropna().astype(str).map(lambda s: s != re.sub(r"\s+", " ", s).strip())
    log("Item_Master", "Alternate item name has extra spaces (would break an exact match)",
        int(odd_name.sum()), "Names are compared after trimming and collapsing spaces")

    # Customers
    log("Customers", "Customer with no industry recorded", int(cu["Industry"].isna().sum()),
        "Kept; cannot be placed in bucket 2 (similar business), as per R6")
    no_sales = ~cu["CustomerID"].isin(sl["CustomerID"])
    log("Customers", "Customer with no sales in the 6 months", int(no_sales.sum()), "Never targeted (no spend)")
    log("Item_Master", "Item with no sales in the 6 months", int((~im["SKU"].isin(sl["SKU"])).sum()),
        "Goes to Red List if active")

    # R2 - inactive items that still hold stock (reported, then excluded)
    inactive = im[im["Status"].astype(str).str.strip().str.lower() == "inactive"]
    held = sa[sa["SKU"].isin(inactive["SKU"])]
    for _, r in held.iterrows():
        q = int(r[qcols].sum())
        if q:
            log("Item_Master", f"Inactive item {r['SKU']} still holds {q} units of stock", 1,
                "Left out of the analysis (R2). Warehouse should decide whether to liquidate")

    # Item master: negative or zero cost / pack
    log("Item_Master", "Zero or missing pack quantity or unit cost",
        int(((im["Master Pack Qty"] <= 0) | (im["Unit Cost (Rs)"] <= 0)).sum()), "Item will not be classified correctly - fix at source")

    return im, sa, sl, cu, issues


def resolve_switches(im):
    """Map each old SKU to the final SKU it was replaced by (R1). Follows chains A->B->C."""
    name_to_sku = {norm_text(n): s for n, s in zip(im["Item Name"], im["SKU"])}
    parent = {}
    for _, r in im.iterrows():
        if r["Alternate Flag"] == "SWITCH":
            target = name_to_sku.get(norm_text(r["Alternate Item (name)"]))
            if target and target != r["SKU"]:
                parent[r["SKU"]] = target
    def final(s, seen=()):
        return final(parent[s], seen + (s,)) if s in parent and s not in seen else s
    return {s: final(s) for s in parent}, parent


def run_review(tables):
    im, sa, sl, cu, issues = clean(tables)
    qcols = [c for c in sa.columns if c.startswith("Qty")]
    mapping, _ = resolve_switches(im)

    # --- R1: re-label old SKUs to the new SKU, then add stock and sales together
    sa["Item"] = sa["SKU"].replace(mapping)
    sl["Item"] = sl["SKU"].replace(mapping)
    stock = sa.groupby("Item")[qcols].sum()
    sold = sl.groupby("Item")["Qty"].sum()
    merged_from = pd.Series({k: ", ".join(v) for k, v in
                             pd.Series(mapping).groupby(pd.Series(mapping)).groups.items()}) if mapping else pd.Series(dtype=str)

    # --- R2: keep only active items; old switch SKUs disappear into their new SKU
    items = im[(im["Status"].astype(str).str.strip().str.lower() == "active") & ~im["SKU"].isin(mapping)].copy()
    items = items.set_index("SKU")
    items["Merged From"] = merged_from.reindex(items.index).fillna("")
    items = items.join(stock).fillna({c: 0 for c in qcols})
    items["Qty Sold (6m)"] = sold.reindex(items.index).fillna(0)

    # --- R3 and R4
    items["Total Stock"] = items[qcols].sum(axis=1)
    items["Aged Qty"] = items[AGED_COLS].sum(axis=1)
    items["Aged Value (Rs)"] = items["Aged Qty"] * items["Unit Cost (Rs)"]
    items["Avg Monthly Sales"] = items["Qty Sold (6m)"] / MONTHS
    items["Months to Clear"] = items["Total Stock"] / items["Avg Monthly Sales"].where(items["Avg Monthly Sales"] > 0)

    # --- R5: checked in this order
    def classify(r):
        if r["Qty Sold (6m)"] < r["Master Pack Qty"]:
            return "Red List"
        if r["Aged Qty"] == 0 or r["Months to Clear"] < 3:
            return "No Action"
        return "Push"
    items["Class"] = items.apply(classify, axis=1)
    items["Reason"] = items.apply(lambda r: why(r), axis=1)

    targets = build_targets(items, sl, cu)
    return items.reset_index(), targets, pd.DataFrame(issues), sl


def why(r):
    if r["Class"] == "Red List":
        return f"Sold {r['Qty Sold (6m)']:.0f} units in 6 months, less than 1 master pack ({r['Master Pack Qty']})"
    if r["Aged Qty"] == 0:
        return "No stock older than 90 days"
    if r["Months to Clear"] < 3:
        return f"Stock clears in {r['Months to Clear']:.1f} months (under 3)"
    return f"{r['Aged Qty']:.0f} aged units and {r['Months to Clear']:.1f} months to clear"


def build_targets(items, sl, cu):
    """R6 and R7. `items` is indexed by SKU. One row per (Push item, customer);
    each customer lands in the first bucket that applies."""
    cust_total = sl.groupby("CustomerID")["Value (Rs)"].sum()
    eligible = cust_total[cust_total > MIN_SPEND]               # R6: only big customers
    industry = cu.set_index("CustomerID")["Industry"]
    bought = sl.groupby("CustomerID")["Item"].apply(set)        # items each customer bought (switch-merged)
    rows = []
    for sku, it in items[items["Class"] == "Push"].iterrows():
        buyers = {c for c, s in bought.items() if sku in s}
        buyer_industries = {industry[c] for c in buyers if pd.notna(industry.get(c))}
        for cid, spend in eligible.items():
            ind = industry.get(cid)
            if cid in buyers:
                b = 1                                            # bought this item before
            elif pd.notna(ind) and ind in buyer_industries:
                b = 2                                            # same industry as a buyer
            elif any(items["Category"].get(x) == it["Category"] for x in bought.get(cid, set()) if x != sku):
                b = 3                                            # bought another item of same category
            else:
                continue
            rows.append({"SKU": sku, "Item Name": it["Item Name"], "CustomerID": cid, "Industry": ind,
                         "Bucket": b, "Bucket Name": BUCKET_NAMES[b], "6m Purchase (Rs)": spend})
    t = pd.DataFrame(rows)
    if t.empty:
        return t
    # R7: bucket first, then highest 6-month purchase value
    return t.sort_values(["SKU", "Bucket", "6m Purchase (Rs)", "CustomerID"], ascending=[True, True, False, True])
