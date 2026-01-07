# Supported formats

This project is built for **Excel exports** (`.xlsx`, `.xls`) that represent a list of transactions.

## Works best with
Any statement file that has, at minimum:
- a **date** column (e.g., `Date`, `Transaction Date`)
- a **description/payee** column (e.g., `Description`, `Payee`, `Merchant`, `Details`)
- either:
  - **Debit / Credit** columns, or
  - a single **Amount** column

The script detects columns by keyword matching (case‑insensitive). If a bank uses non‑standard headers, renaming the columns in Excel is usually the fastest fix.

## Common header keywords the detector understands

### Date
- `date`, `transaction date`, `posted`, `value date`

### Description / payee
- `description`, `payee`, `merchant`, `counterparty`, `details`, `memo`

### Amounts
- outgoing: `debit`, `withdrawal`, `amount out`, `paid out`, `charge`
- incoming: `credit`, `deposit`, `amount in`, `paid in`, `received`
- single amount: `amount`

### Optional fields
- transaction id: `transaction id`, `reference`, `ref`, `check number`
- balance: `balance`, `running balance`, `account balance`

## Quick self‑check before you run
1) Open one of your exports and confirm the headers exist on the first row.  
2) Make sure dates are real dates (not screenshots or merged cells).  
3) If your file contains subtotal/total lines, it’s fine — the script filters common “service rows”.  
4) Run on a small sample first (one or two files), then scale up.

## Not a fit for
- PDF statements (not parsed here)
- CSVs (easy to add later, but not in this snapshot)
- exports where the “table” is split across multiple sheets or has multi‑row headers
