import re
import difflib
from pathlib import Path
import pandas as pd
from combine import load_all

ALIASES = Path("data/supplier_aliases.csv")


def clean_supplier(name):
    s = str(name).upper().strip()
    s = s.replace("&", " AND ")
    s = re.sub(r"[.,'()]", " ", s)
    s = re.sub(r"\bLIMITED\b", "LTD", s)
    s = re.sub(r"\bTA\b", "T/A", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def add_supplier_clean(spend):
    spend = spend.copy()
    spend["supplier_clean"] = spend["supplier"].map(clean_supplier)
    aliases = pd.read_csv(ALIASES)
    mapping = dict(zip(aliases["alias"], aliases["canonical"]))
    spend["supplier_clean"] = spend["supplier_clean"].replace(mapping)
    return spend


if __name__ == "__main__":
    spend, recon = load_all()
    spend = add_supplier_clean(spend)

    print(f"Distinct suppliers: {spend['supplier'].nunique()} raw -> "
          f"{spend['supplier_clean'].nunique()} after rules and aliases")

    total = spend["euro"].sum()
    top = (spend.groupby("supplier_clean")["euro"]
                .agg(orders="size", total="sum")
                .sort_values("total", ascending=False))
    top["share_%"] = (top["total"] / total * 100).round(1)
    top["cumulative_%"] = top["share_%"].cumsum().round(1)

    print("\nTop 10 suppliers by spend, Q1 2024 to Q2 2026:")
    print(top.head(10).to_string())

    n_half = (top["cumulative_%"] < 50).sum() + 1
    print(f"\n{n_half} of {len(top)} suppliers account for half of all spend.")

    names = sorted(spend["supplier_clean"].unique())
    print("\nRemaining similar pairs (checked, not merged):")
    for i, name in enumerate(names):
        for other in difflib.get_close_matches(name, names[i + 1:], n=3, cutoff=0.85):
            print(f"  {name}  <->  {other}")