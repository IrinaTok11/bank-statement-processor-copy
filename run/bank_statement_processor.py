"""
Bank Statement Processor

Merge multiple Excel bank statement exports into a single consolidated workbook.

Workflow:
  1) Put your statement files (.xlsx / .xls) in the same folder as this script
  2) Run: python bank_statement_processor.py
  3) Open: consolidated_statements.xlsx (created in the same folder)

Notes:
  - Column detection is keyword-based (date / description / debit / credit / amount).
  - Temporary Excel files that start with "~$" are ignored.
"""

import pandas as pd
import os
from pathlib import Path
import warnings
import re
from datetime import datetime

# Suppress warnings for cleaner output
warnings.filterwarnings('ignore')

# Output file name
OUTPUT_FILE = 'consolidated_statements.xlsx'


__author__ = "Irina Tokmianina"


def normalize_date_column(df, column_name):
    """
    Normalize date column - handles US (MM/DD/YYYY), EU (DD/MM/YYYY), 
    ISO (YYYY-MM-DD) formats.
    
    Args:
        df: DataFrame
        column_name: name of date column
        
    Returns:
        DataFrame with normalized dates
    """
    if column_name not in df.columns:
        return df
    
    # Try different date formats
    date_formats = [
        '%m/%d/%Y',      # US: MM/DD/YYYY
        '%d/%m/%Y',      # EU: DD/MM/YYYY
        '%Y-%m-%d',      # ISO: YYYY-MM-DD
        '%d.%m.%Y',      # DE/RU: DD.MM.YYYY
        '%m-%d-%Y',      # US alt: MM-DD-YYYY
        '%d-%m-%Y',      # EU alt: DD-MM-YYYY
    ]
    
    # Try each format
    for date_format in date_formats:
        try:
            df[column_name] = pd.to_datetime(df[column_name], format=date_format, errors='coerce')
            # If successful (most dates parsed), keep it
            if df[column_name].notna().sum() > len(df) * 0.8:  # 80% success rate
                df[column_name] = df[column_name].dt.date
                return df
        except:
            continue
    
    # Fallback: let pandas infer
    df[column_name] = pd.to_datetime(df[column_name], errors='coerce')
    df[column_name] = df[column_name].dt.date
    
    return df


def normalize_amount_column(df, column_name):
    """
    Normalize amount columns - handles different number formats.
    Removes currency symbols ($, €, £), normalizes separators.
    
    Args:
        df: DataFrame
        column_name: name of amount column
        
    Returns:
        DataFrame with normalized amounts
    """
    if column_name not in df.columns:
        return df
    
    # Convert to string, remove currency symbols and spaces
    df[column_name] = df[column_name].astype(str)
    df[column_name] = df[column_name].str.replace('$', '', regex=False)
    df[column_name] = df[column_name].str.replace('€', '', regex=False)
    df[column_name] = df[column_name].str.replace('£', '', regex=False)
    df[column_name] = df[column_name].str.replace('¥', '', regex=False)
    df[column_name] = df[column_name].str.replace(' ', '', regex=False)
    
    # Handle European format (1.000,00) vs US format (1,000.00)
    # Strategy: if comma is last separator, it's decimal
    df[column_name] = df[column_name].apply(lambda x: x.replace(',', '.') if ',' in x and x.rfind(',') > x.rfind('.') else x)
    df[column_name] = df[column_name].str.replace(',', '', regex=False)  # Remove thousand separators
    
    # Convert to numeric
    df[column_name] = pd.to_numeric(df[column_name], errors='coerce')
    
    return df


def normalize_counterparty_name(df, column_name):
    """
    Normalize counterparty/description names.
    Handles US (LLC, Inc, Corp) and EU (GmbH, Ltd, SA) entities.
    
    Creates new column with '_normalized' suffix.
    
    Args:
        df: DataFrame
        column_name: name of description/counterparty column
        
    Returns:
        DataFrame with normalized names in new column
    """
    if column_name not in df.columns:
        return df
    
    # Entity type mapping (US/EU/International)
    entity_mapping = {
        # US entities
        r'\bllc\b': 'LLC',
        r'\binc\b': 'Inc',
        r'\bcorp\b': 'Corp',
        r'\bcorporation\b': 'Corp',
        r'\blimited\s+liability\s+company\b': 'LLC',
        r'\bincorporated\b': 'Inc',
        
        # EU entities
        r'\bgmbh\b': 'GmbH',
        r'\bltd\b': 'Ltd',
        r'\blimited\b': 'Ltd',
        r'\bs\.?a\.?\b': 'SA',
        r'\bs\.?r\.?l\.?\b': 'SRL',
        r'\ba\.?g\.?\b': 'AG',
        
        # Common merchant patterns
        r'amazon\.com.*mktplc': 'Amazon',
        r'amzn\s+mktp': 'Amazon',
        r'paypal\s*\*': 'PayPal',
        r'sq\s*\*': 'Square',
    }
    
    def clean_name(text):
        if pd.isna(text) or str(text).strip() == '':
            return text
        
        original = str(text).strip()
        cleaned = original.lower()
        
        # Apply entity normalization
        for pattern, replacement in entity_mapping.items():
            cleaned = re.sub(pattern, replacement, cleaned, flags=re.IGNORECASE)
        
        # Remove special characters
        cleaned = re.sub(r'["\'\(\)\[\]]', '', cleaned)
        cleaned = re.sub(r'\s+', ' ', cleaned)
        cleaned = cleaned.strip()
        
        # Capitalize first letter of each word
        cleaned = ' '.join(word.capitalize() for word in cleaned.split())
        
        return cleaned
    
    new_column_name = f"{column_name}_normalized"
    df[new_column_name] = df[column_name].apply(clean_name)
    
    return df


def detect_columns(df):
    """
    Intelligent column detection - finds Date, Amount, Description columns
    by keyword matching (case-insensitive).
    
    Args:
        df: DataFrame
        
    Returns:
        dict with detected column names or None if not found
    """
    detected = {
        'date': None,
        'amount_out': None,   # Debit/Withdrawal
        'amount_in': None,    # Credit/Deposit
        'description': None,
        'transaction_id': None,
        'balance': None,
    }
    
    # Column name patterns (case-insensitive)
    patterns = {
        'date': ['date', 'transaction date', 'posted date', 'trans date'],
        'amount_out': ['debit', 'withdrawal', 'amount out', 'spent', 'payments'],
        'amount_in': ['credit', 'deposit', 'amount in', 'received', 'deposits'],
        'description': ['description', 'payee', 'merchant', 'counterparty', 'details', 'memo'],
        'transaction_id': ['transaction id', 'reference', 'ref', 'check number', 'trans id'],
        'balance': ['balance', 'running balance', 'account balance'],
    }
    
    # Normalize column names for matching
    columns_lower = {col.lower().strip(): col for col in df.columns}
    
    # Match patterns
    for field, keywords in patterns.items():
        for keyword in keywords:
            for col_lower, col_original in columns_lower.items():
                if keyword in col_lower:
                    detected[field] = col_original
                    break
            if detected[field]:
                break
    
    return detected


def process_statement_file(file_path):
    """
    Process a single bank statement file.
    
    Args:
        file_path: Path to Excel file
        
    Returns:
        DataFrame with processed data or None if failed
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path, sheet_name=0)
        
        # Detect columns
        columns = detect_columns(df)
        
        # Check minimum requirements
        if not columns['date'] or not columns['description']:
            print(f"⚠️  SKIPPED '{file_path.name}': Missing required columns (Date or Description)")
            return None
        
        if not columns['amount_out'] and not columns['amount_in']:
            print(f"⚠️  SKIPPED '{file_path.name}': Missing amount columns")
            return None
        
        # Filter out service rows
        initial_rows = len(df)
        df = df.dropna(how='all')
        
        # Remove header duplicates and summary rows
        service_keywords = ['total', 'subtotal', 'balance', 'opening', 'closing', 'beginning', 'ending']
        if columns['date']:
            df = df[df[columns['date']].notna()]
            mask = ~df[columns['date']].astype(str).str.lower().str.contains('|'.join(service_keywords), na=False)
            df = df[mask]
        
        filtered_rows = initial_rows - len(df)
        
        # Add source tracking
        df['Source_Bank'] = file_path.stem
        
        # Normalize columns
        df = normalize_date_column(df, columns['date'])
        
        if columns['amount_out']:
            df = normalize_amount_column(df, columns['amount_out'])
        if columns['amount_in']:
            df = normalize_amount_column(df, columns['amount_in'])
        
        if columns['description']:
            df = normalize_counterparty_name(df, columns['description'])
        
        # Rename to standard column names
        rename_map = {
            columns['date']: 'Date',
            columns['description']: 'Description',
        }
        
        if columns['amount_out']:
            rename_map[columns['amount_out']] = 'Amount_Out'
        if columns['amount_in']:
            rename_map[columns['amount_in']] = 'Amount_In'
        if columns['transaction_id']:
            rename_map[columns['transaction_id']] = 'Transaction_ID'
        if columns['balance']:
            rename_map[columns['balance']] = 'Balance'
        
        df = df.rename(columns=rename_map)
        
        # Select output columns
        output_columns = ['Date', 'Description', 'Source_Bank']
        if 'Amount_Out' in df.columns:
            output_columns.insert(2, 'Amount_Out')
        if 'Amount_In' in df.columns:
            output_columns.insert(2, 'Amount_In')
        if 'Transaction_ID' in df.columns:
            output_columns.insert(1, 'Transaction_ID')
        if 'Balance' in df.columns:
            output_columns.append('Balance')
        if 'Description_normalized' in df.columns:
            output_columns.append('Description_normalized')
        
        df = df[[col for col in output_columns if col in df.columns]]
        
        print(f"✓ Processed '{file_path.name}': {len(df)} transactions (filtered {filtered_rows} service rows)")
        
        return df
        
    except Exception as e:
        print(f"❌ ERROR processing '{file_path.name}': {e}")
        return None


def main():
    """
    Main function - processes all Excel files in the working folder.
    """
    # Paths
    SCRIPT_DIR = Path(__file__).parent
    # Work in the folder where the script is located.
    # Put your bank statement files (.xlsx / .xls) in the same folder.
    WORK_DIR = SCRIPT_DIR

    print(f"Working folder: {WORK_DIR}")

    excel_files = [p for p in (list(WORK_DIR.glob('*.xlsx')) + list(WORK_DIR.glob('*.xls'))) 
                   if p.name != OUTPUT_FILE and not p.name.startswith('~$')]
    
    print(f"\n{'='*70}")
    print("Bank Statement Processor")
    print(f"{'='*70}\n")
    print(f"Working folder: {WORK_DIR}")
    print(f"")
    print(f"\n{'='*70}\n")
    
    # Find all Excel files
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]
    
    if not excel_files:
        print("⚠️  No Excel files found in the working folder.")
        print(f"   Place your bank statement files in: {WORK_DIR}")
        return
    
    print(f"Found {len(excel_files)} file(s) to process:\n")
    
    # Process each file
    all_dataframes = []
    for file_path in excel_files:
        df = process_statement_file(file_path)
        if df is not None:
            all_dataframes.append(df)
    
    if not all_dataframes:
        print("\n❌ No files were successfully processed.")
        return
    
    # Merge all dataframes
    print(f"\n{'='*70}")
    print("Merging data...")
    result_df = pd.concat(all_dataframes, ignore_index=True)
    
    # Sort by date
    if 'Date' in result_df.columns:
        result_df = result_df.sort_values('Date')
    
    # Save output
    output_path = WORK_DIR / OUTPUT_FILE
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Transactions')
        
        # Format columns
        worksheet = writer.sheets['Transactions']
        
        # Set column widths
        column_widths = {
            'A': 15,   # Date
            'B': 15,   # Transaction_ID
            'C': 15,   # Amount_Out
            'D': 15,   # Amount_In
            'E': 40,   # Description
            'F': 20,   # Source_Bank
            'G': 40,   # Description_normalized
            'H': 15,   # Balance
        }
        
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        # Format date column
        from openpyxl.styles import numbers
        for row in range(2, len(result_df) + 2):
            cell = worksheet[f'A{row}']
            cell.number_format = 'YYYY-MM-DD'
    
    print(f"✓ Done!")
    print(f"{'='*70}")
    print(f"Files processed: {len(all_dataframes)}")
    print(f"Total transactions: {len(result_df)}")
    print(f"Output saved to: {output_path}")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()