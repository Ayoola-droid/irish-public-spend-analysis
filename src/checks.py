import re
import pandas as pd

FILE_NAME = re.compile(r"^\d{4}_Q[1-4]\.xlsx$")
NEEDED = {"supplier", "euro", "description"}


def check_files(raw):
    problems = []
    for f in sorted(raw.iterdir()):
        if f.name.startswith("~$"):
            problems.append(f"{f.name}: Excel has a file open. Close it and run again.")
        elif not FILE_NAME.match(f.name):
            problems.append(f"{f.name}: rename it like 2026_Q3.xlsx (year_quarter).")
        else:
            try:
                cols = pd.read_excel(f, sheet_name="Sheet1", nrows=0).columns
            except ValueError:
                problems.append(f"{f.name}: no tab called Sheet1.")
                continue
            missing = NEEDED - set(cols.str.strip().str.lower())
            if missing:
                problems.append(f"{f.name}: missing column(s) {sorted(missing)}.")
    return problems


def check_data(spend, recon):
    warnings = []
    for _, r in recon[~recon["reconciles"]].iterrows():
        warnings.append(f"{r['quarter']}: orders do not add up to the file's TOTAL row "
                        f"(difference €{r['difference']:,.2f}).")
    for _, r in recon[recon["total_rows"] != 1].iterrows():
        warnings.append(f"{r['quarter']}: expected 1 TOTAL row, found {r['total_rows']}.")
    unknown = sorted(spend.loc[spend["category"] == "Uncategorised", "description"].dropna().unique())
    if unknown:
        warnings.append("No category rule for: " + ", ".join(unknown)
                        + ". Add keywords to data/category_rules.csv.")
    return warnings