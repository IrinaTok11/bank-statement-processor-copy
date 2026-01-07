# Case study: Consolidating statements for review

## Scenario
In investigation and compliance workflows (including financial investigations), you often receive:
- multiple statement exports
- from different banks
- covering different months
- with inconsistent column headers and formats

Before any analysis starts, you need **one table**.

## What the tool solves
- merges many exports into a single workbook
- keeps the original source visible via `Source_Bank`
- removes common non‑transaction rows (totals, opening/closing lines)
- standardises dates so you can sort and filter reliably

## Typical downstream uses
- reconciliation against invoices / contracts
- building a transaction timeline
- spotting duplicates and unusual patterns
- preparing a clean dataset for Power BI or Excel pivots

## Deliverable
A single file: `consolidated_statements.xlsx` with a **Transactions** sheet.

The output is intentionally plain: one row per transaction, ready for filters, pivots, or import into a BI model.
