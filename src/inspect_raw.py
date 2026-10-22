from pathlib import Path
import pandas as pd

RAW = Path("data/raw")
files = sorted(RAW.glob("*.xlsx"))
print(f"Found {len(files)} files\n")

for f in files:
    xl = pd.ExcelFile(f)
    print(f"=== {f.name} | sheets: {xl.sheet_names}")
    preview = pd.read_excel(f, sheet_name=0, header=None, nrows=8)
    print(preview.to_string())
    print()