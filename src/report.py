import sys
from datetime import date
from pathlib import Path
import pandas as pd
from combine import load_all, RAW
from checks import check_files, check_data
from suppliers import add_supplier_clean
from categories import add_category

OUT = Path("output")
OUT.mkdir(exist_ok=True)

# 0. Check the input files before doing anything else
problems = check_files(RAW)
if problems:
    print("Report NOT created. Fix these files first:")
    for p in problems:
        print("  -", p)
    sys.exit(1)

# 1. Load, reconcile, clean, flag
spend, recon = load_all()
spend = add_supplier_clean(spend)
spend = add_category(spend)
key = ["quarter", "supplier_clean", "euro", "description"]
spend["dup_flag"] = spend.duplicated(subset=key, keep=False)

total = spend["euro"].sum()
first, latest = spend["quarter"].min(), spend["quarter"].max()

# 2. Analysis tables
by_quarter = (spend.groupby("quarter")["euro"]
                   .agg(orders="size", spend="sum").reset_index())

top = (spend.groupby("supplier_clean")["euro"]
            .agg(orders="size", spend="sum")
            .sort_values("spend", ascending=False).reset_index())
top["share_%"] = (top["spend"] / total * 100).round(1)
top["cumulative_%"] = top["share_%"].cumsum().round(1)

by_cat = (spend.groupby("category")["euro"]
               .agg(orders="size", spend="sum")
               .sort_values("spend", ascending=False).reset_index())
by_cat["share_%"] = (by_cat["spend"] / total * 100).round(1)

by_desc = (spend.groupby(["category", "description"])["euro"]
                .agg(orders="size", spend="sum")
                .sort_values("spend", ascending=False).reset_index())
by_desc["share_%"] = (by_desc["spend"] / total * 100).round(1)

quarters = sorted(spend["quarter"].unique())
previous = quarters[-2]
year_ago = f"{int(latest[:4]) - 1}{latest[4:]}"
pivot = spend.pivot_table(index="category", columns="quarter", values="euro",
                          aggfunc="sum", fill_value=0)
latest_q = pd.DataFrame({
    "same_quarter_last_year": pivot.get(year_ago, 0),
    "previous_quarter": pivot[previous],
    "latest_quarter": pivot[latest],
})
latest_q["change_vs_last_year"] = latest_q["latest_quarter"] - latest_q["same_quarter_last_year"]
latest_q = latest_q.sort_values("latest_quarter", ascending=False).reset_index()

seen_before = set(spend.loc[spend["quarter"] < latest, "supplier_clean"])
new_suppliers = (spend[(spend["quarter"] == latest) & ~spend["supplier_clean"].isin(seen_before)]
                 .groupby("supplier_clean")["euro"].agg(orders="size", spend="sum")
                 .sort_values("spend", ascending=False).reset_index())

dups = (spend[spend["dup_flag"]]
        .groupby(key).size().reset_index(name="copies")
        .sort_values("euro", ascending=False))
dups["extra_value"] = dups["euro"] * (dups["copies"] - 1)

# 3. Summary and data-quality log
summary = pd.DataFrame([
    ("Period covered", f"{first} to {latest}"),
    ("Orders", len(spend)),
    ("Total spend (EUR)", round(total, 2)),
    ("Suppliers", spend["supplier_clean"].nunique()),
    ("Quarters reconciled to source totals", f"{recon['reconciles'].sum()} of {len(recon)}"),
    ("Latest quarter", latest),
    ("Latest quarter spend (EUR)", round(latest_q["latest_quarter"].sum(), 2)),
    (f"Same quarter last year ({year_ago}) (EUR)", round(latest_q["same_quarter_last_year"].sum(), 2)),
    ("New suppliers in latest quarter", len(new_suppliers)),
    ("Largest supplier", f"{top.loc[0, 'supplier_clean']} ({top.loc[0, 'share_%']}% of spend)"),
    ("Largest category", f"{by_cat.loc[0, 'category']} ({by_cat.loc[0, 'share_%']}% of spend)"),
    ("Suppliers making up half of spend", int((top["cumulative_%"] < 50).sum() + 1)),
    ("Possible duplicate groups to review", len(dups)),
    ("Report generated", date.today().isoformat()),
], columns=["Measure", "Value"])

desc_fixed = spend["description_raw"].fillna("") != spend["description"].fillna("")
supp_fixed = spend["supplier"].str.strip() != spend["supplier_clean"]
dq_log = pd.DataFrame([
    ("Rejected", "TOTAL row at the bottom of each file (used for reconciliation)", int(recon["total_rows"].sum())),
    ("Rejected", "Hidden 'HM old FW' tab in 2024 Q3: internal invoice notes with personal data, removed from the source file", 26),
    ("Fixed", "Description spacing or capital letters standardised", int(desc_fixed.sum())),
    ("Fixed", "Supplier name standardised by rules or alias file", int(supp_fixed.sum())),
    ("Flagged, kept", "Description matched no category rule", int((spend["category"] == "Uncategorised").sum())),
    ("Flagged, kept", "Row belongs to a possible duplicate group", int(spend["dup_flag"].sum())),
    ("Flagged, kept", "Missing description", int(spend["description"].isna().sum())),
], columns=["Action", "Issue", "Rows"])

orders = spend[["quarter", "supplier", "supplier_clean", "category", "description",
                "euro", "dup_flag", "source_file"]]

# 4. Write the workbook
path = OUT / f"wicklow_spend_report_{latest}.xlsx"
with pd.ExcelWriter(path, engine="openpyxl") as xw:
    summary.to_excel(xw, sheet_name="Summary", index=False)
    latest_q.to_excel(xw, sheet_name="Latest Quarter", index=False)
    new_suppliers.to_excel(xw, sheet_name="New Suppliers", index=False)
    by_quarter.to_excel(xw, sheet_name="By Quarter", index=False)
    top.head(20).to_excel(xw, sheet_name="Top Suppliers", index=False)
    by_cat.to_excel(xw, sheet_name="By Category", index=False)
    by_desc.to_excel(xw, sheet_name="By Description", index=False)
    dups.to_excel(xw, sheet_name="Duplicates to Review", index=False)
    dq_log.to_excel(xw, sheet_name="Data Quality Log", index=False)
    recon.to_excel(xw, sheet_name="Reconciliation", index=False)
    orders.to_excel(xw, sheet_name="All Orders", index=False)

print(f"Report written: {path}")
print(summary.to_string(index=False))
print()
print(dq_log.to_string(index=False))

from excel_format import format_workbook
format_workbook(path)
print("Formatting and charts added.")

warnings = check_data(spend, recon)
if warnings:
    print("\nWARNINGS - check before sharing the report:")
    for w in warnings:
        print("  -", w)
else:
    print("\nAll checks passed.")