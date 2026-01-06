# Run workspace

This folder is the **working directory**.

Put your bank statement files (`.xlsx` / `.xls`) in this folder (next to the script) and run the processor.

## Run

```bash
python bank_statement_processor.py
```

## Output

- The merged result is saved **in this same folder** as `consolidated_statements.xlsx`.

## Notes

- Keep only the statement files you want to merge in this folder when you run the script.
- The script ignores temporary Excel files that start with `~$`.
