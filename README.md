# Irish Public Spend Analysis: Wicklow County Council

**Business question:** Where does Wicklow County Council spend money on purchase orders over €20,000, who are its biggest suppliers, and how reliable is the published data?

**Data:** Wicklow County Council purchase orders over €20,000, Q1 2024 to Q2 2026 (10 quarterly .xlsx files), from [wicklow.ie](https://www.wicklow.ie/Living/Your-Council/Finance/Procurement/Purchase-Orders-Over-20-000).

**Tools:** Python (pandas, openpyxl), Excel

**Status:** In progress. The pipeline runs end to end, from the raw files to a formatted Excel report. Still to do: a monthly run for new quarters, and the final write-up.

## Key findings so far

- **€149.7m across 1,686 orders**, Q1 2024 to Q2 2026, from 276 suppliers.
- **The data is complete.** All 10 quarters reconcile to the council's own published totals to the cent.
- **Spend is concentrated.** 10 of 276 suppliers received half of all spend. The largest, ABM Contractors, received €21.6m (14.4%), all for housing construction.
- **Housing and roads take almost two thirds of spend.** Housing is the largest category at €58.2m (38.9%), mostly new housing construction (€45.8m). Roads & Transport follows at €35.9m (24.0%). ICT, Fire & Emergency and Corporate Services together account for under 4%.
- **Q4 is the busiest quarter.** Q4 spend was €22.9m in 2024 and €23.9m in 2025, roughly three times Q1 (€8.0m and €7.9m), consistent with year-end spending.
- **Most repeated orders are recurring contracts, not errors.** 76 groups of identical orders were flagged (extra copies worth €4.67m). Most follow contract patterns: monthly homeless support payments and quarterly parking and engineering fees. One engineering order of €24,215.63 appears in all 10 quarters.
- **One possible double entry stands out:** €320,854.73 to PFH Technology Group for ICT infrastructure, entered twice in 2024 Q2. It only became visible after supplier names and descriptions were cleaned.

## The report

`py src\report.py` writes `output/wicklow_spend_report_<latest quarter>.xlsx`, a formatted workbook with charts:

| Tab | Contents |
|---|---|
| Summary | Headline measures: period, orders, spend, suppliers, reconciliation status |
| By Quarter | Orders and spend per quarter, with a chart |
| Top Suppliers | Top 20 suppliers with share and cumulative share of spend, with a chart |
| By Category | Spend across 12 categories, with a chart |
| By Description | Every description, its category and spend |
| Duplicates to Review | Groups of identical orders in the same quarter and the value at stake |
| Data Quality Log | Rows rejected, fixed and flagged, and why |
| Reconciliation | Calculated vs published total for every quarter |
| All Orders | Every cleaned order with its original supplier name and source file |

## Method

1. **Load:** read Sheet1 of each quarterly file and keep the three published columns.
2. **Reconcile:** each file ends with a TOTAL row (an Excel SUM). Counting it as an order would double every quarter, so it is removed from the orders and used to check that the order rows add up to the council's total.
3. **Clean descriptions:** strip leading, trailing and double spaces, then standardise capital letters to the most common spelling (e.g. "Legal fees" → "Legal Fees").
4. **Clean supplier names:** rules (`&` → `AND`, `LIMITED` → `LTD`, `TA` → `T/A`, punctuation removed), then a reviewed alias file (`data/supplier_aliases.csv`) for matches that rules cannot catch. Fuzzy matching suggested candidate pairs; each was judged by hand.
5. **Categorise:** each description is matched against keyword rules in `data/category_rules.csv`, top to bottom, and the first match sets one of 12 categories (e.g. any description containing "footpath" → Roads & Transport). Rule order settles overlaps: "fire" comes before "vehicle", so fire engines count as Fire & Emergency. Descriptions that match no rule are labelled Uncategorised and counted in the data-quality log.
6. **Flag duplicates:** there is no PO number or date, so identical rows (same quarter, supplier, amount and description) cannot be proven to be errors. They are flagged for review and kept in the totals.
7. **Report:** pandas writes the tables; openpyxl adds formatting and charts.

## Data quality log

| Action | Issue | Rows |
|---|---|---|
| Rejected | TOTAL row at the bottom of each file (used for reconciliation) | 10 |
| Rejected | Hidden "HM old FW" tab in 2024 Q3 (see data notes) | 26 |
| Fixed | Description spacing or capital letters standardised | 1,197 |
| Fixed | Supplier name standardised by rules or alias file | 545 |
| Flagged, kept | Description matched no category rule | 0 |
| Flagged, kept | Row belongs to a possible duplicate group | 197 |

Most fixes are cosmetic: invisible trailing spaces from hand-typed spreadsheets, and legal-form spellings. Only 2 supplier names were merged with another supplier (ESB Networks, PFH Technology Group). Each individual supplier is otherwise spelled consistently by the council.

## Data notes

- All 10 files share three columns (Supplier, EURO, Description), but header casing varies between quarters.
- There is no order date or PO number, so the quarter comes from the filename and duplicates can only be flagged, not confirmed.
- The 2024 Q3 file, as published, has a second tab ("HM old FW") of internal invoice notes: negative amounts against a retired housing maintenance framework, plus staff names and tenants' home addresses. It is not part of the purchase order list, so it is excluded from the analysis. For privacy, it has been removed from the copy in this repo, and that file's TOTAL formula was saved as its value.
- The 2026 Q1 file has stray blank columns beyond column C; only Supplier, EURO and Description are used.
- The 173 distinct descriptions include different words for the same work (e.g. "Energy Retrofit" and "Energy Retrofit Works") and typos ("Footpath Maintenannce"); the category rules group these together.

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
| `py src\report.py` | **Main script:** builds the formatted Excel report in `output/` |
| `py src\combine.py` | Loads all quarters and prints the reconciliation against each file's TOTAL row |
| `py src\suppliers.py` | Prints supplier counts, the top 10 suppliers and spend concentration |
| `py src\categories.py` | Prints spend by category and lists the descriptions in each category for review |
| `py src\quality.py` | Prints duplicate groups, orders recurring across quarters and description variants |
| `py src\inspect_raw.py` | Previews the first rows of each raw file to check headers and layout |

## Project structure

- `data/raw/`: council files as published (except the personal-data tab removed from 2024 Q3)
- `data/supplier_aliases.csv`: reviewed supplier-name matches
- `data/category_rules.csv`: keyword rules that map descriptions to categories
- `src/combine.py`: load, reconcile and clean descriptions (`load_all()`)
- `src/suppliers.py`: supplier-name standardisation (`add_supplier_clean()`)
- `src/categories.py`: spend categories (`add_category()`)
- `src/quality.py`: duplicate and consistency checks
- `src/report.py`, `src/excel_format.py`: build and format the Excel report
- `output/`: generated reports (not committed)
