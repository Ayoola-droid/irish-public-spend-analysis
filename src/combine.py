from pathlib import Path
import pandas as pd

RAW = Path("data/raw")


def load_quarter(path):
    df = pd.read_excel(path, sheet_name="Sheet1")
    df.columns = df.columns.str.strip().str.lower()
    df = df[["supplier", "euro", "description"]].copy()
    df["description_raw"] = df["description"]    
    df["description"] = (df["description"]
                         .str.strip()
                         .str.replace(r"\s+", " ", regex=True))

    is_total = df["supplier"].astype(str).str.strip().str.lower() == "total"
    stated_total = df.loc[is_total, "euro"].sum()

    orders = df[~is_total].copy()
    orders["quarter"] = path.stem.replace("_", "-")
    orders["source_file"] = path.name
    return orders, stated_total, is_total.sum()


def load_all():
    frames, checks = [], []
    for f in sorted(RAW.glob("*.xlsx")):
        orders, stated, n_total_rows = load_quarter(f)
        calculated = orders["euro"].sum()
        checks.append({
            "quarter": f.stem.replace("_", "-"),
            "orders": len(orders),
            "total_rows": n_total_rows,
            "calculated": round(calculated, 2),
            "stated_total": round(stated, 2),
            "difference": round(calculated - stated, 2),
        })
        frames.append(orders)
    recon = pd.DataFrame(checks)
    recon["reconciles"] = recon["difference"].abs() < 0.01
    spend = pd.concat(frames, ignore_index=True)

    lower = spend["description"].str.lower()
    most_common = spend.groupby(lower)["description"].agg(lambda s: s.mode().iloc[0])
    spend["description"] = lower.map(most_common)

    return spend, recon


if __name__ == "__main__":
    spend, recon = load_all()
    print("Reconciliation against each file's TOTAL row:")
    print(recon.to_string(index=False))
    print(f"\nOrders: {len(spend)}   Total spend: €{spend['euro'].sum():,.2f}")
    print(f"Quarters that reconcile: {recon['reconciles'].sum()} of {len(recon)}")