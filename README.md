# SEU Summary Excel Script

This script reads Athena + Zeus `*_SEU.csv` files and generates a formatted Excel summary with
error counts, cross section per flip-flop, and FIT values for each shift register.

---

## What it does
- Scans a folder for `*_SEU.csv` files
- Sums errors per SR (SR-0 to SR-59)
- Computes:
  - CrossSection/FF  
    `errors / (fluence × #BlackBoxes × #FF per BlackBox)`
  - FIT  
    `CrossSection × 1e15 × 0.001`
- Writes everything into a single Excel file (`summary.xlsx`)
- Each run appears as a colored block in the Summary sheet

---

## Requirements
- Python 3
- `pandas`
- `openpyxl`

Install dependencies:
```bash
pip install pandas openpyxl
