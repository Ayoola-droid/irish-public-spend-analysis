from pathlib import Path
import pandas as pd

RAW = Path("data/raw")
frames = []

print("Every sheet in every file:")
for f in sorted(RAW.glob("*.xlsx")):
    sheets = pd.read_excel(f, sheet_name=None)
    for name, sheet in sheets.items():
        print(f"  {f.name:14} {name!r:14} {len(sheet):4} rows")

    df = sheets["Sheet1"].copy()
    df.columns = df.columns.str.strip().str.lower()
    df["quarter"] = f.stem.replace("_", "-")
    df["source_file"] = f.name
    frames.append(df)

spend = pd.concat(frames, ignore_index=True)
spend["amount"] = pd.to_numeric(spend["euro"], errors="coerce")

print("\nColumns:", spend.columns.tolist())
print(f"Total rows: {len(spend)}")

print("\nRows and total spend per quarter:")
print(spend.groupby("quarter").agg(rows=("amount", "size"), total=("amount", "sum")))

print("\nMissing values per column:")
print(spend.isna().sum())

bad = spend[spend["amount"].isna()]
print(f"\nRows where EURO is not a number: {len(bad)}")
print(bad.to_string())