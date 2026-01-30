import pandas as pd
import numpy as np
from schema import QuotationSchema


# -------------------------------------------------
# Semantic master column definitions
# -------------------------------------------------
# Vendors rarely follow strict naming conventions.
# These synonym sets allow us to resolve columns
# semantically instead of relying on exact names.
#
# Example:
#   "item no", "sr no", "sl no" → item_no
#   "qty", "quantity", "qnty" → qty
# -------------------------------------------------

MASTER_COLUMN_SYNONYMS = {
    "item_no": [
        "s.no", "sr no", "sr.no", "serial", "serial no",
        "item no", "item", "line no", "sl no","item nos."
    ],
    "description": [
        "description", "item description", "particulars",
        "details", "material description"
    ],
    "qty": [
        "qty", "quantity", "qnty", "qty.", "no."
    ],
    "unit": [
        "unit", "uom", "units", "measure", "unit of measure"
    ]
}


# -------------------------------------------------
# Header detection logic
# -------------------------------------------------
# Many quotation Excel files have:
# - title rows
# - tender metadata
# - empty spacing
# before the actual table header starts.
#
# We scan the first ~20 rows and look for
# a row that "looks like" a table header
# using simple heuristics.
# -------------------------------------------------

def find_header_row(df: pd.DataFrame) -> int:
    """Scans the first 20 rows to find the actual header."""
    for i, row in df.head(20).iterrows():
        # Convert row to string and normalize
        row_str = " ".join(str(x).lower() for x in row.values if pd.notna(x))

        # Heuristic:
        # A valid header usually contains 'description'
        # and at least one quantity indicator
        if "description" in row_str and ("qty" in row_str or "quantity" in row_str):
            return i

    raise ValueError("Could not detect header row in reference file")


# -------------------------------------------------
# Semantic column resolver
# -------------------------------------------------
# This function maps actual Excel column names
# to semantic roles (item_no, qty, etc.)
#
# It avoids:
# - hardcoded column names
# - positional assumptions
# -------------------------------------------------

def resolve_master_columns(columns: list[str]) -> dict:
    resolved = {}
    
    # Track used columns to prevent double mapping
    used_cols = set()

    # Priority order: item_no -> description -> qty -> unit
    # "item_no" is most specific, so resolve it first to grab "Item Nos"
    # before "qty" grabs it via "Nos" (if we hadn't removed it)
    search_order = ["item_no", "description", "qty", "unit"]

    for role in search_order:
        synonyms = MASTER_COLUMN_SYNONYMS[role]
        for col in columns:
            if col in used_cols:
                continue
                
            # Normalize column name
            col_norm = col.lower().replace("_", " ").strip()

            # Match against known synonyms
            if any(s in col_norm for s in synonyms):
                resolved[role] = col
                used_cols.add(col) # Mark as used
                break

    return resolved


# -------------------------------------------------
# Main reference loader
# -------------------------------------------------
# Responsibilities:
# 1. Read Excel file
# 2. Detect correct header row
# 3. Normalize column names
# 4. Resolve master vs vendor columns
# 5. Build QuotationSchema
#
# NOTE:
# - Sheet selection is handled upstream (UI)
# - This loader assumes one table per sheet
# -------------------------------------------------

def load_reference(
    path: str,
    sheet_name: str | None = None
) -> tuple[pd.DataFrame, QuotationSchema]:

    # If sheet_name is provided, use it.
    # Otherwise default to the first sheet.
    read_sheet = sheet_name if sheet_name else 0

    # -------------------------------------------------
    # First read: load without headers
    # Used only to detect header row
    # -------------------------------------------------
    temp_df = pd.read_excel(path, sheet_name=read_sheet, header=None)
    header_idx = find_header_row(temp_df)

    # -------------------------------------------------
    # Second read: reload with correct header
    # -------------------------------------------------
    df = pd.read_excel(path, sheet_name=read_sheet, header=header_idx)

    # Sanitize "nan" strings that might have polluted the source
    # This ensures they are treated as real NaNs downstream
    df.replace(["nan", "NaN", "NAN", "Nan"], np.nan, inplace=True)

    # Normalize column names:
    # - lowercase
    # - trim whitespace
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Collect all column names
    all_columns = list(df.columns)

    # -------------------------------------------------
    # Resolve semantic master columns
    # -------------------------------------------------
    resolved = resolve_master_columns(all_columns)
    
    # -------------------------------------------------
    # Standardization: Rename columns to internal schema
    # -------------------------------------------------
    # We want consistent column names for the rest of the application
    # mapping: original_col -> standard_col
    rename_map = {}
    
    # Map 'item_no' role -> 's.no' standard
    if "item_no" in resolved:
        rename_map[resolved["item_no"]] = "s.no"
        
    # Map other roles directly if they exist
    # (description -> description, qty -> qty, unit -> unit)
    for role in ["description", "qty", "unit"]:
        if role in resolved:
             rename_map[resolved[role]] = role
             
    # Apply renaming
    df.rename(columns=rename_map, inplace=True)
    
    # Update all_columns based on new names for schema building
    all_columns = list(df.columns)

    # Required master columns for any quotation
    # Unit is optional
    # NOTE: We now look for the STANDARD names, not just role existence
    required_std_cols = ["s.no", "description", "qty"]
    missing = [c for c in required_std_cols if c not in all_columns]

    # Fail fast if core columns are missing
    if missing:
        # Fallback debug: print what we have
        print(f"DEBUG: Missing {missing}. Resolved was: {resolved}. Columns are: {all_columns}")
        raise ValueError(
            f"Could not resolve required master columns: {missing}. "
            f"Available columns: {all_columns}"
        )

    # -------------------------------------------------
    # Build ordered master columns
    # -------------------------------------------------
    # Order matters for downstream merging logic
    master_columns = [c for c in ["s.no", "description", "qty", "unit"] if c in all_columns]

    # Append unit only if present
    if "unit" in resolved:
        master_columns.append(resolved["unit"])

    # Everything else is considered vendor-specific
    # (rates, amounts, taxes, totals, etc.)
    vendor_value_columns = [
        c for c in all_columns if c not in master_columns
    ]

    # -------------------------------------------------
    # Debug visibility (replace with logging later)
    # -------------------------------------------------
    print("Resolved master columns:")
    for role, column in resolved.items():
        print(f"  {role} → {column}")

    # -------------------------------------------------
    # Build schema object
    # -------------------------------------------------
    schema = QuotationSchema(
        all_columns=all_columns,
        master_columns=master_columns,
        vendor_value_columns=vendor_value_columns,
        header_row_index=header_idx
    )

    return df, schema

