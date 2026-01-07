[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](#quick-start)

# Bank statement processor (Excel → forensic analysis)

A production‑grade automation for consolidating business bank statements and screening for financial irregularities. Give it multiple statement exports (Excel) and it will merge them into a single **Transactions** table, categorize operations, detect anomalies, and generate a ready‑to‑review **findings report** with full traceability to source rows. No manual cleanup, no guesswork: repeatable forensic screening for fraud detection, compliance reviews, and financial investigations.

**Live page:** https://irinatok11.github.io/bank-statement-processor/

---

## Table of contents
- [Overview](#overview)
  - [Scope](#scope)
- [Who this is for](#who-this-is-for)
- [Quick start](#quick-start)
- [Key features](#key-features)
- [How it works](#how-it-works)
- [What's included](#whats-included)
- [Design decisions](#design-decisions)
- [Methodology](#methodology)
- [Reference articles](#reference-articles)
- [Project structure](#project-structure)
- [Getting started](#getting-started)
- [Input data requirements](#input-data-requirements)
- [Output](#output)
- [Reproducibility](#reproducibility)
- [Audit trail](#audit-trail)
- [Demo: Before → After](#demo-before--after)
- [Limitations](#limitations)
- [Formatting and code style](#formatting-and-code-style)
- [Tests (smoke)](#tests-smoke)
- [Data privacy](#data-privacy)
- [License](#license)
- [Contact](#contact)

---

## Overview

Bank statements vary by institution and export method: headers, date formats, amount conventions, and text fields are inconsistent. This project consolidates multiple Excel exports into one standardized workbook, categorizes transactions, applies rule‑based forensic screening, and produces audit‑grade findings reports with full source traceability.

### Scope

This repository focuses on:
- **Multi‑file consolidation:** Merges statements from different banks/accounts into a single **Transactions** table
- **Format detection:** Auto‑detects column layout (date / description / debit / credit / amount / balance / transaction ID)
- **Standardization:** Normalizes dates, amounts, descriptions into a canonical schema
- **Transaction categorization:** Classifies operations into business categories (payroll, vendor payments, customer receipts, taxes, banking fees, etc.)
- **Forensic screening:** Applies explainable rules to flag duplicates, split payments, vendor concentration spikes, round‑amount clustering, weekend activity, and new high‑value counterparties
- **Traceability:** Every row includes source file, sheet, and row number for audit verification
- **Multi‑sheet output:** Produces **Transactions**, **Findings & Alerts** (executive summary + detail table), **Evidence Index** (transaction‑level drill‑down), **Analytics Dashboard** (cash metrics, Benford's Law, vendor concentration), and **Config & Thresholds** (full methodology documentation)

**Audience.** Financial analysts, fraud examiners, forensic accountants, compliance officers, and forensic accountants who need repeatable preprocessing and systematic anomaly detection.

**Tech stack.** Python · pandas · openpyxl

---

## Who this is for

- **Financial analysts and fraud examiners** conducting suspicious transaction analyses
- **Fraud examiners** screening business accounts for payment irregularities
- **Forensic accountants** investigating financial misconduct or asset dissipation
- **Compliance officers** reviewing vendor payment patterns and cash flow anomalies
- **Financial analysts and BI specialists** building dashboards from multi‑bank data
- **Professionals who need repeatable preprocessing** instead of manual Excel cleanup

---

## Quick start

```bash
cd run
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r ../requirements.txt
python bank_statement_processor.py
```

Result: The script reads all `.xlsx`/`.xls` files in `run/`, consolidates them, applies forensic rules, and creates `consolidated_statements.xlsx` in `run/` with five analysis sheets.

---

## Key features

- **One‑click forensic pipeline:** Merge → standardize → categorize → screen → report
- **Intelligent format detection:** Handles Chase, Santander, ABSA formats (and generic patterns) with auto‑header detection
- **Transaction categorization:** Rule‑based classification into 20+ operation types (Vendor Payments, Customer Receipts, Payroll, Taxes, Banking Fees, Transfers, etc.)
- **Explainable screening rules:** Duplicate detection, split payment patterns, vendor concentration growth, round‑amount clustering, weekend activity alerts, new counterparty flags
- **Priority‑based triage:** Findings tagged as P1 (High), P2 (Medium), P3 (Low) for efficient review
- **Full audit traceability:** Every transaction includes `Source_File`, `Source_Sheet`, `Source_Row` for verification
- **Multi‑sheet outputs:** Executive summary, detailed findings table, transaction‑level evidence index, analytics dashboard, and configuration documentation
- **Benford's Law analysis:** First‑digit distribution testing for debit amounts
- **Vendor concentration tracking:** Identifies top spending patterns and growth anomalies
- **Cash flow metrics:** Period analysis, monthly trends, net cash movement
- **Excel‑first output:** Single workbook for quick review and further analysis in Excel, Power BI, or Python

---

## How it works

High‑level pipeline:

1. **Read all statement files** from the `run/` workspace (Excel `.xlsx`/`.xls`)
2. **Detect bank format** by filename and header analysis (auto‑identifies column positions)
3. **Standardize schema** into canonical fields: `Date`, `Description`, `Debit`, `Credit`, `Bank`, `Source_File`, `Source_Sheet`, `Source_Row`
4. **Normalize text** by removing transaction IDs, cleaning descriptions, extracting vendor names
5. **Categorize transactions** using keyword‑based rules (Vendor Payment, Customer Receipt, Payroll, Tax, Banking Fee, etc.)
6. **Apply forensic screening:**
   - Duplicate Payment Detection (same vendor ±1% amount within 3 days)
   - Split Payment Patterns (2+ same‑day payments to same vendor >$10k total)
   - Vendor Concentration Growth (baseline→latest month spend >2x)
   - Round Amount Clustering (exact $1,000 increments)
   - Weekend Activity Alerts (large debits on Sat/Sun, excludes taxes/payroll)
   - New Counterparty Flags (first‑time vendors with debits >$50k)
7. **Generate findings report** with priority levels and grouped alerts
8. **Build analytics dashboard** with cash metrics, Benford's Law analysis, vendor concentration tables, and monthly trends
9. **Document methodology** in **Config & Thresholds** sheet (all rules, thresholds, rationale)
10. **Export consolidated workbook** with five sheets: **Transactions**, **Findings & Alerts**, **Evidence Index**, **Analytics Dashboard**, **Config & Thresholds**

---

## What's included

### Inputs
- **Bank statement exports:** Excel files (`.xlsx`, `.xls`) with transaction tables
- **Supported formats:** Chase, Santander, ABSA (named formats), plus generic auto‑detection for other banks
- **Minimum required fields per file:** Date column + Description column + Amount(s) (either signed Amount or separate Debit/Credit columns)

> **Note:** The demo uses **synthetic statement data** (generated for portfolio demonstration). Real client statements are never committed to public repositories.

### Outputs
Single consolidated workbook (`consolidated_statements.xlsx`) containing:

1. **Transactions** — Standardized table with full source traceability
2. **Findings & Alerts** — Executive summary + priority‑tagged findings table with vendor context
3. **Evidence Index** — Transaction‑level drill‑down for every alert (source file/sheet/row)
4. **Analytics Dashboard** — Cash flow metrics, Benford's Law analysis, vendor concentration, monthly trends, category breakdowns
5. **Config & Thresholds** — Complete documentation of all detection rules, thresholds, rationale, and reproducibility notes

### Process artifacts
- Console logging with transaction counts, findings summary (P1/P2/P3), opening/ending balances
- Automated QA validation report in **Config & Thresholds** sheet comparing expected vs. actual values for key metrics

---

## Design decisions

- **Excel engine.** Uses `openpyxl` for read/write to preserve formatting, formulas, and conditional styles while editing existing workbooks. `xlsxwriter` is excellent for fresh files but cannot modify existing workbooks without breaking their structure.

- **Format detection strategy.** Filename‑based bank identification (first word before underscore) combined with heuristic header detection (searches for Date + Description + amount columns). This handles both named formats (Chase, Santander, ABSA) and generic bank exports without configuration files.

- **Column matching.** Headers are matched in a **case/space‑insensitive** way (so `Date`, `date`, or `Date ` all resolve to the same target). This reduces edge‑case failures when source files come from different export systems.

- **Categorization over black‑box ML.** Transaction classification uses keyword‑based rules (transparent, auditable, explainable) rather than ML models. This ensures analysts can verify why a transaction was categorized a specific way and adjust rules as needed.

- **Forensic rules as data.** Detection thresholds and parameters are stored in a central `RULES` dictionary with display names, parameters, rationale, and source references. The **Config & Thresholds** sheet reads these values to document methodology rather than hard‑coding narrative text.

- **Priority‑based triage.** Findings are assigned P1 (High), P2 (Medium), P3 (Low) based on risk severity. This allows reviewers to focus on critical issues first without being overwhelmed by lower‑priority observations.

- **Source traceability.** Every transaction includes `Source_File`, `Source_Sheet`, `Source_Row` so analysts can verify findings by looking up the original statement row. This is critical for audit defensibility and legal proceedings.

- **Single workspace folder.** The pipeline expects all input files in `run/`, processes them, and writes the output to the same folder. This constraint keeps the workflow simple and avoids accidental cross‑contamination between analysis runs.

---

## Methodology

A concise, practitioner‑focused write‑up of the pipeline is published in HTML:

- **Website:** https://irinatok11.github.io/bank-statement-processor/methodology.html
- **Repo file:** [docs/methodology.html](docs/methodology.html)

It covers the architecture and data flow, concrete outputs (five‑sheet workbook structure), the forensic questions we answer, detection logic for each rule type, operational constraints, and a reproducibility checklist.

---

## Reference articles

Supporting notes used across the project:

- **Architecture (high‑level):** [docs/reference/architecture.md](docs/reference/architecture.md)
- **Supported formats & header matching:** [docs/reference/supported_formats.md](docs/reference/supported_formats.md)
- **Case study — Suspicious transaction analysis in practice:** [docs/reference/case_study.md](docs/reference/case_study.md)

---

## Project structure

```
bank-statement-processor/
├─ run/                                    # working folder (run the script here)
│  ├─ bank_statement_processor.py          # main script
│  ├─ consolidated_statements.xlsx         # appears after running
│  └─ README.md                            # workspace usage notes
├─ demo/                                   # sample files for demonstration
│  ├─ sample_output_consolidated_statements.xlsx  # example output (synthetic data)
│  └─ README.md                            # explanation of demo files
├─ docs/                                   # GitHub Pages (live site content)
│  ├─ _config.yml                          # Jekyll config for GitHub Pages
│  ├─ index.html                           # landing page for the processor case
│  ├─ methodology.html                     # methodology and reproducibility details
│  ├─ assets/
│  │  ├─ .gitkeep
│  │  ├─ css/
│  │  │  └─ portfolio.css                  # global styles for documentation pages
│  │  └─ screenshots/                      # demo screenshots stored here
│  │     ├─ .gitkeep
│  │     ├─ global_input.png               # Input example (Global Business Bank statement)
│  │     ├─ northbridge_input.png          # Input example (Northbridge statement)
│  │     ├─ uk_input.png                   # Input example (UK bank statement)
│  │     ├─ transactions.png               # Output: consolidated Transactions table
│  │     ├─ findings_exec_summary.png      # Output: Findings executive summary
│  │     ├─ findings_alerts_table.png      # Output: detailed findings table
│  │     ├─ evidence_index.png             # Output: transaction‑level evidence
│  │     ├─ analytics_dashboard.png        # Output: analytics dashboard
│  │     ├─ vendor_concentration.png       # Output: vendor concentration analysis
│  │     ├─ config_thresholds.png          # Output: detection thresholds documentation
│  │     └─ config_run_specific.png        # Output: run‑specific validation report
│  └─ reference/                           # lightweight reference notes (Markdown)
│     ├─ architecture.md                   # system architecture / data flow overview
│     ├─ case_study.md                     # short case narrative
│     └─ supported_formats.md              # supported formats and header detection notes
├─ .gitignore
├─ requirements.txt
├─ LICENSE
└─ README.md                               # this file
```

---

## Getting started

### 1) Create and activate a virtual environment

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
```

### 2) Install dependencies

```bash
pip install -r requirements.txt
```

### 3) Put input files into the run workspace

Copy your bank statement exports (`.xlsx` / `.xls`) into:
- `run/` (next to `bank_statement_processor.py`)

> **Important:** The processor extracts the bank name from the **first word before underscore** in the filename.  
> Examples:  
> - `CHASE_20230515.xlsx` → Bank = "CHASE"  
> - `NORTHBRIDGE_May_June_2023.xlsx` → Bank = "NORTHBRIDGE"  
> - `GLOBAL_BUSINESS_BANK_USD.xlsx` → Bank = "GLOBAL"

> **Tip:** Excel temporary files like `~$...` are ignored automatically.

### 4) Run the processor

```bash
cd run
python bank_statement_processor.py
```

### 5) Open the result

The processor creates an output file in the same folder:
- `run/consolidated_statements.xlsx`

Open it in Excel to review the five analysis sheets: **Transactions**, **Findings & Alerts**, **Evidence Index**, **Analytics Dashboard**, **Config & Thresholds**.

> **Note:** A sample output workbook with synthetic data is available in `demo/sample_output_consolidated_statements.xlsx` for reference — you can explore it without running the processor.

> If you use different statement files, place them into `run/` and ensure they are the **only** Excel files there when running the script (the processor reads all `.xlsx`/`.xls` in the working directory).

---

## Input data requirements

### File formats
- Excel exports: `.xlsx` / `.xls`

### Minimum required fields per file
Each input file must contain:
- A transaction **date** column
- A **description** column

Amounts can be represented as:
- A signed **amount** column, or
- Separate **in/out** (credit/debit) columns (recommended for better categorization)

### Supported header patterns

The processor auto‑detects columns by keyword matching (case‑insensitive, space‑tolerant). Supported patterns and detection notes:

- **Date columns:** `Date`, `Transaction Date`, `Trans Date`, `Posting Date`, `Value Date`, `DATE`
- **Description columns:** `Description`, `Details`, `Transaction`, `Memo`, `Narrative`
- **Debit columns:** `Debit`, `Out`, `Money Out`, `Withdrawal`, `Payment`, `Payments Out`
- **Credit columns:** `Credit`, `In`, `Money In`, `Deposit`, `Receipt`, `Payments In`
- **Amount columns:** `Amount`, `Value`, `Transaction Amount`
- **Balance columns:** `Balance`, `Running Balance`, `Closing Balance`
- **Transaction ID columns:** `ID`, `Reference`, `Transaction ID`, `Ref`, `Check #`

For detailed format specifications, see: [docs/reference/supported_formats.md](docs/reference/supported_formats.md)

---

## Output

### Output file
- `consolidated_statements.xlsx` (saved into `run/`)

### Output sheets

1. **Transactions** — Consolidated standardized table
   - Typical columns: `Date`, `Description`, `Debit`, `Credit`, `Bank`, `Source_File`, `Source_Sheet`, `Source_Row`, `Description_clean`, `Vendor_Normalized`, `Operation_Type`, `Direction`, `Month`, `Weekday`

2. **Findings & Alerts** — Forensic screening results
   - Executive summary (total findings, priority breakdown)
   - Detailed findings table with columns: `Alert_ID`, `Priority`, `Category`, `Vendor`, `Details`, `Related_Txns`, `Total_Amount`, `First_Date`, `Last_Date`, `Review_Status`

3. **Evidence Index** — Transaction‑level drill‑down
   - One row per transaction involved in an alert
   - Columns: `Alert_ID`, `Category`, `Vendor`, `Group`, `Txn_ID`, `Txn_Date`, `Debit`, `Credit`, `Bank`, `Source` (file | sheet | row)

4. **Analytics Dashboard** — Summary metrics and analysis
   - Cash flow metrics (analysis period, total transactions, opening/ending balances, net cash change)
   - Monthly trends (inflows, outflows, net by month)
   - Benford's Law analysis (first‑digit distribution for debit amounts)
   - Vendor concentration (top vendors by spending with percentage of total)
   - Spending by category (outflows and inflows breakdowns)

5. **Config & Thresholds** — Complete methodology documentation
   - Detection thresholds table (all rules, thresholds, rationale, source)
   - Categorization rules (operation type definitions and keyword patterns)
   - Exclusion criteria (what was filtered out and why)
   - Run‑specific parameters (p95 debit, weekend threshold, date ranges)
   - Automated QA validation report (expected vs. actual values for key metrics)

---

## Reproducibility

The processor is designed for **audit‑grade reproducibility**:

1. **Deterministic output:** Same input files → same results (no randomness, no ML model variability)
2. **Version control:** Script includes `__version__` header; commit hash can be added to output metadata
3. **Rule documentation:** All detection thresholds and parameters are stored in central `RULES` dictionary and documented in **Config & Thresholds** sheet
4. **Source traceability:** Every transaction and finding links back to original statement file/sheet/row
5. **Validation report:** **Config & Thresholds** sheet includes automated QA checks comparing expected vs. actual values for cash metrics, vendor totals, and category sums

**Reproducibility checklist:**
- [ ] Input files archived with consistent naming (bank_YYYYMMDD format)
- [ ] Python version and dependencies documented (`requirements.txt` pinned)
- [ ] Script version logged in output metadata
- [ ] Run parameters documented in **Config & Thresholds** sheet
- [ ] Validation report shows zero tolerance violations

---

## Audit trail

The processor maintains full audit traceability:

- **Transaction‑level:** Every row in **Transactions** sheet includes `Source_File`, `Source_Sheet`, `Source_Row`
- **Finding‑level:** Every alert in **Findings & Alerts** references specific transaction IDs
- **Evidence‑level:** **Evidence Index** sheet provides one‑row‑per‑transaction drill‑down for every finding with direct source references
- **Methodology‑level:** **Config & Thresholds** sheet documents all rules, thresholds, and exclusions applied

**Audit verification workflow:**
1. Reviewer opens **Findings & Alerts** and identifies an alert of interest (e.g., Alert A0001)
2. Reviewer navigates to **Evidence Index** and filters by `Alert_ID = A0001`
3. Reviewer sees all related transactions with source references (e.g., `GLOBAL_BUSINESS_BANK_USD.xlsx | GBB_BizChk_USD_2023-07 | R22`)
4. Reviewer opens original statement file, navigates to sheet `GBB_BizChk_USD_2023-07`, row 22, and verifies the transaction details

---

## Demo: Before → After

> **Note:** All screenshots show **synthetic demo data** generated for portfolio demonstration. These are not real bank statements or client transactions.

### Input examples — Raw bank statement exports

The processor handles multiple statement formats automatically. Here are examples of the input files before processing:

| Global Business Bank (USD) | Northbridge Business Bank (USD) | UK Business Bank (USD) |
|---|---|---|
| ![Global input](docs/assets/screenshots/global_input.png) | ![Northbridge input](docs/assets/screenshots/northbridge_input.png) | ![UK input](docs/assets/screenshots/uk_input.png) |

### Output: Consolidated analysis workbook

After processing, the script generates a five‑sheet workbook with standardized data and forensic findings:

#### Transactions — Consolidated table with full traceability
![Transactions table](docs/assets/screenshots/transactions.png)

#### Findings & Alerts — Executive summary
![Findings executive summary](docs/assets/screenshots/findings_exec_summary.png)

#### Findings & Alerts — Detailed findings table
![Findings alerts table](docs/assets/screenshots/findings_alerts_table.png)

#### Evidence Index — Transaction‑level drill‑down
![Evidence index](docs/assets/screenshots/evidence_index.png)

#### Analytics Dashboard — Cash metrics and trends
![Analytics dashboard](docs/assets/screenshots/analytics_dashboard.png)

#### Analytics Dashboard — Vendor concentration analysis
![Vendor concentration](docs/assets/screenshots/vendor_concentration.png)

#### Config & Thresholds — Detection rules documentation
![Config thresholds](docs/assets/screenshots/config_thresholds.png)

#### Config & Thresholds — Run‑specific validation report
![Config run specific](docs/assets/screenshots/config_run_specific.png)

---

## Limitations

- **Language:** Transaction descriptions must be in English (or Latin script) for keyword‑based categorization
- **Date formats:** Auto‑detection handles most formats but may require manual adjustment for ambiguous patterns (e.g., DD/MM/YYYY vs. MM/DD/YYYY)
- **Categorization accuracy:** Keyword‑based rules achieve ~85–95% accuracy; edge cases require manual review
- **Duplicate detection:** Current logic uses vendor name + amount + time window; does not account for legitimate recurring payments (e.g., monthly subscriptions)
- **Vendor normalization:** Text cleaning removes common prefixes/suffixes but may not handle all variations (e.g., "ABC Corp" vs. "ABC Corporation" vs. "ABC Ltd")
- **Performance:** Large datasets (>10,000 transactions) may take several minutes to process; consider splitting analysis by time period
- **File formats:** Currently supports Excel only (`.xlsx`, `.xls`); CSV and PDF exports require conversion

---

## Formatting and code style

This is a production script, so we keep the toolchain light. If you want the same formatting used in this project:

```bash
# Optional, but recommended
pip install black ruff

# Format
black .

# Lint (fast default rule set)
ruff check .
```

---

## Tests (smoke)

A quick "does it run" test is often enough for a utility like this. The suggested structure:

```
tests/
└─ test_smoke.py
```

And a minimal test body (pseudo‑code):

```python
from pathlib import Path
from run.bank_statement_processor import main  # or similar entry point

def test_smoke(tmp_path: Path):
    # Copy demo statement files into a temp run/
    # Call the pipeline
    # Assert that consolidated_statements.xlsx was created
    # Assert that Transactions sheet has expected row count
    # Assert that Findings sheet was generated
    assert (tmp_path / "consolidated_statements.xlsx").exists()
```

---

## Data privacy

**Important:** Do not commit client statements or personally identifiable information to the repository.

- Use synthetic/anonymized examples for demos and screenshots
- The `.gitignore` file excludes common sensitive patterns (`*.xlsx`, `*.xls`, `run/*.xlsx`)
- For portfolio/case study purposes, generate synthetic data with realistic patterns but no real PII
- When sharing analysis results publicly, ensure all company names, account numbers, and personal details are replaced with placeholders

**Demo data disclaimer:** All statement screenshots in this repository are **synthetic** and generated specifically for demonstration purposes. They do not represent real bank accounts, transactions, or client data.

---

## License

This project is available under the **MIT License**. See [LICENSE](LICENSE).

---

## Contact

**IRINA TOKMIANINA** — Financial Analyst / Fraud Detection Specialist  
LinkedIn: [linkedin.com/in/tokmianina](https://www.linkedin.com/in/tokmianina/) · Email: <irinatokmianina@gmail.com>
