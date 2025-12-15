# DeltaMaster Report Automation

Python script to merge two DeltaMaster Excel exports (TopM + Addison), calculate KPIs and “Modifikationen”, and produce a standardized management report in Excel.

## What it does
- Reads TopM export (DB I / contribution margin logic)
- Applies business rules (“Modifikationen”, incl. special 09 & 32 logic)
- Aggregates to cost center level (KSt) without summing percentage columns
- Reads Addison export (Umsatzerlöse / Aufwendungen / Rohergebnis)
- Merges both sources and calculates “Aufwendungen final”
- Exports formatted Excel and highlights key columns

## Data privacy
This repo should contain **only synthetic sample files** (if any).
Do **not** upload real employer exports, names, cost centers, invoices, or account numbers.

## How to run (local)
```bash
pip install -r requirements.txt
python src/deltamaster_merge.py --topm "path/to/topm.xlsx" --addison "path/to/addison.xlsx" --out "outputs/Ergebnis_final_strukturiert.xlsx"
