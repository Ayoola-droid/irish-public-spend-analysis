from datetime import date
from pathlib import Path
import pandas as pd
from combine import load_all
from suppliers import add_supplier_clean

OUT = Path("output")
OUT.mkdir(exist_ok=True)

# 1. Load, reconcile, clean, flag
spend, recon = load_all()
spend = add_supplier_clean(spend)
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

by_desc = (spend.groupby("description")["euro"]
                .agg(orders="size", spend="sum")
                .sort_values("spend", ascending=False).reset_index())
by_desc["share_%"] = (by_desc["spend"] / total * 100).round(1)

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
    ("Largest supplier", f"{top.loc[0, 'supplier_clean']} ({top.loc[0, 'share_%']}% of spend)"),
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
    ("Flagged, kept", "Row belongs to a possible duplicate group", int(spend["dup_flag"].sum())),
    ("Flagged, kept", "Missing description", int(spend["description"].isna().sum())),
], columns=["Action", "Issue", "Rows"])

orders = spend[["quarter", "supplier", "supplier_clean", "description",
                "euro", "dup_flag", "source_file"]]

# 4. Write the workbook
path = OUT / f"wicklow_spend_report_{latest}.xlsx"
with pd.ExcelWriter(path, engine="openpyxl") as xw:
    summary.to_excel(xw, sheet_name="Summary", index=False)
    by_quarter.to_excel(xw, sheet_name="By Quarter", index=False)
    top.head(20).to_excel(xw, sheet_name="Top Suppliers", index=False)
    by_desc.head(20).to_excel(xw, sheet_name="By Category", index=False)
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