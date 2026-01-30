# src/vendor_loader.py
import pandas as pd
import numpy as np


def find_header_row(df: pd.DataFrame) -> int:
    """Scans the first 20 rows to find the actual header."""
    for i, row in df.head(20).iterrows():
        row_str = " ".join([str(x).lower() for x in row.values])
        if "description" in row_str and ("rate" in row_str or "amount" in row_str):
            return i
    # Fallback to 0 if not found, though realistically we might want to warn
    return 0

def load_vendor(path: str, sheet_name: str = None) -> pd.DataFrame:
    read_sheet = sheet_name if sheet_name else 0

    # Read first to find header
    temp_df = pd.read_excel(path, sheet_name=read_sheet, header=None)
    header_idx = find_header_row(temp_df)
    
    df = pd.read_excel(path, sheet_name=read_sheet, header=header_idx)
    # Sanitize "nan" strings
    df.replace(["nan", "NaN", "NAN", "Nan"], np.nan, inplace=True)
    
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    # Drop empty rows
    df = df.dropna(how='all')
    
    return df

