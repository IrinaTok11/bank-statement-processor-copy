# Demo Files

This folder contains sample files for demonstration purposes.

## Sample Output

**File:** `sample_output_consolidated_statements.xlsx`

This is an example of the processor's output workbook containing:

1. **Transactions** — Standardized table with 645 transactions from 3 banks (Global, Northbridge, UK)
2. **Findings & Alerts** — Executive summary and 7 flagged items (P1/P2/P3 priority)
3. **Evidence Index** — Transaction-level drill-down with source references
4. **Analytics Dashboard** — Cash metrics, Benford's Law analysis, vendor concentration
5. **Config & Thresholds** — Complete methodology documentation and QA validation

**Data Notice:** All data in this file is **synthetic** and generated specifically for demonstration. No real client information, bank statements, or transaction details are included.

## How This Output Was Generated

This sample was created by processing three synthetic bank statement files:
- `GLOBAL_BUSINESS_BANK_USD.xlsx` (267 transactions)
- `NORTHBRIDGE_BUSINESS_BANK_USD.xlsx` (234 transactions)  
- `UK_BUSINESS_BANK_USD.xlsx` (144 transactions)

The processor:
1. Auto-detected bank formats and column layouts
2. Standardized transactions into canonical schema
3. Applied transaction categorization (20+ operation types)
4. Ran forensic screening (6 detection algorithms)
5. Generated analytics and validation reports

**Bank Name Extraction:** Note that bank names were extracted from the **first word before underscore** in filenames:
- `GLOBAL_BUSINESS_BANK_USD.xlsx` → Bank = "GLOBAL"
- `NORTHBRIDGE_BUSINESS_BANK_USD.xlsx` → Bank = "NORTHBRIDGE"  
- `UK_BUSINESS_BANK_USD.xlsx` → Bank = "UK"

## Using This Sample

You can open this file in Excel to explore:
- How transactions are standardized across different bank formats
- What forensic findings look like (Findings & Alerts sheet)
- How source traceability works (Evidence Index sheet)
- What analytics are provided (Analytics Dashboard sheet)
- How methodology is documented (Config & Thresholds sheet)

This demonstrates the processor's capabilities without needing to run the script yourself.

## Input Statement Files

Input statement files (Excel workbooks) are **not** included in this repository to keep file size manageable. Screenshots of input examples are available in `docs/assets/screenshots/`:
- `global_input.png`
- `northbridge_input.png`
- `uk_input.png`

These show the raw bank statement formats before processing.
