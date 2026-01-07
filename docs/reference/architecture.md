# Architecture (high‑level)

```
┌───────────────────────────────┐
│   run/ (working folder)       │
│   - bank_statement_processor  │
│   - statement_*.xlsx / *.xls  │
└───────────────┬───────────────┘
                │
                ▼
┌─────────────────────────────────────────────┐
│  Column detection (keyword matching)         │
│  - date / description / amount columns       │
│  - optional: balance, transaction id         │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│  Normalisation                               │
│  - date parsing → ISO-like date column        │
│  - amount parsing (debit/credit or amount)    │
│  - light description cleanup (optional)       │
│  - source tracking (filename → Source_Bank)   │
│  - filter obvious service rows (totals, etc.) │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│  Merge + output                               │
│  - append all rows                             │
│  - sort by Date                                │
│  - write consolidated_statements.xlsx          │
│    (Transactions sheet)                        │
└─────────────────────────────────────────────┘
```

## Design choices
- **One‑folder workflow.** The script is meant to live in the same folder as the statements you’re processing.
- **Deterministic output.** The same inputs produce the same consolidated workbook (no hidden state).
- **Traceability.** Every row carries `Source_Bank` so you can track where it came from.

## Assumptions and limits
- Inputs are Excel transaction tables with headers on the first row.
- If a bank export uses unusual column names, you may need to rename headers to match obvious keywords.
- PDF parsing and multi‑sheet reconciliation are out of scope for this snapshot.
