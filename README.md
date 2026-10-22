# Irish Public Spend Analysis: Wicklow County Council

**Business question:** Where does Wicklow County Council spend money on purchase orders over €20,000, who are its biggest suppliers, and how reliable is the published data?

**Data:** Wicklow County Council purchase orders over €20,000, Q1 2024 to Q2 2026 (10 quarterly .xlsx files), from [wicklow.ie](https://www.wicklow.ie/Living/Your-Council/Finance/Procurement/Purchase-Orders-Over-20-000).

**Tools:** Python (pandas, openpyxl), Excel

**Status:** In progress

## Setup

Requires Windows with [Python 3](https://www.python.org/downloads/) and [Git](https://git-scm.com/downloads) installed.

```powershell
git clone https://github.com/Ayoola-droid/irish-public-spend-analysis.git
cd irish-public-spend-analysis
py -m venv .venv
.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

If PowerShell says "running scripts is disabled on this system", run this once and then activate again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Run the scripts from the project folder:

| Script | What it does |
|---|---|
| `py src\inspect_raw.py` | Previews the first rows of each raw file to check headers and layout |
| `py src\combine.py` | Combines all quarters into one table and prints row counts, totals and missing values |

## Data notes

- All 10 files share three columns (Supplier, EURO, Description), but header casing varies between quarters.
- There is no order date or PO number, so the quarter comes from the filename and duplicates can only be flagged, not confirmed.
- Each file ends with a TOTAL row (an Excel SUM). Counting it as an order doubles every quarter, so it is removed and used instead to reconcile the order rows back to the council's own total. Excluding these rows, the 10 files hold 1,686 orders; the published TOTAL rows add up to €149.7m.
- All amounts are numeric.
- The 2024 Q3 file, as published, has a second tab ("HM old FW") of internal invoice notes: negative amounts against a retired housing maintenance framework, plus staff names and tenants' home addresses. It is not part of the purchase order list, so it is excluded from the analysis. For privacy, it has been removed from the copy in this repo, and that file's TOTAL formula was saved as its value.
- The 2026 Q1 file has stray blank columns beyond column C; only Supplier, EURO and Description are used.
- Supplier names are inconsistent ("T/A" vs "TA", "LTD" vs "LIMITED") and include typos.
- Some identical orders recur across quarters (e.g. the same €24,215.63 engineering order in 2024 Q4, 2025 Q2 and 2026 Q2), which suggests recurring contracts.

## Early observations

- Q4 is the largest quarter in both full years (€22.9m in 2024, €23.9m in 2025), roughly three times Q1 (€8.0m and €7.9m), consistent with year-end spending. To be confirmed after cleaning.

## Project structure

- `data/raw/`: council files as published (except the personal-data tab removed from 2024 Q3)
- `src/`: Python scripts
- `output/`: generated reports (not committed)
