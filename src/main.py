import os
import pandas as pd
from config import DATA_DIR
from reference_loader import load_reference
from vendor_loader import load_vendor
from aligner import align_vendor_to_reference


def verify_amounts(final_report: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """
    Verifies that Vendor Amount = Reference Qty * Vendor Rate.
    Overwrites Amount if incorrect.
    Returns: (updated_df, list_of_errors)
    """
    verification_errors = []
    
    # Identify vendor columns
    # We look for pairs of Rate_X and Amount_X
    # We need 'qty' column from report
    if 'qty' in final_report.columns:
        # Convert qty to numeric, coercing errors
        qty_series = pd.to_numeric(final_report['qty'], errors='coerce')
        
        for col in final_report.columns:
            if col.startswith('Rate_'):
                vendor_suffix = col.replace('Rate_', '')
                amt_col = f'Amount_{vendor_suffix}'
                
                if amt_col in final_report.columns:
                    # Calculate Expected
                    rate_series = pd.to_numeric(final_report[col], errors='coerce')
                    current_amt_series = pd.to_numeric(final_report[amt_col], errors='coerce')
                    
                    # Iterate rows to check math
                    for idx in final_report.index:
                        q = qty_series[idx]
                        r = rate_series[idx]
                        curr_a = current_amt_series[idx]
                        
                        # Skip if Qty is valid text (e.g. "L.S.") but not numeric, or invalid/nan
                        # We only verify math if both Qty and Rate are numbers.
                        if pd.isna(q) or pd.isna(r):
                            continue
                            
                        # Double check q is actually a number (though pd.to_numeric handled it, 
                        # just ensure we aren't multiplying by a text residual if something slipped)
                        if not isinstance(q, (int, float)):
                            continue
                            
                        expected_a = q * r
                        
                        # Compare with current amount
                        # If current amount is NaN, we fill it.
                        # If current amount exists, we check tolerance.
                        
                        update_needed = False
                        
                        if pd.isna(curr_a):
                            update_needed = True
                        else:
                            # Tolerance of 1.0 (currency units) to be safe against rounding vs calc
                            if abs(curr_a - expected_a) > 1.0:
                                update_needed = True
                                
                        if update_needed:
                            # Update DataFrame
                            final_report.at[idx, amt_col] = expected_a
                            
                            # Log Error if it wasn't just a simple fill-in of a NaN 
                            # (User wants to know about ERRORS, filling NaN is implicitly fixing incomplete data, but maybe worth noting too?)
                            # Let's log it if it was a mismatch.
                            if not pd.isna(curr_a):
                                desc = final_report.at[idx, 'description']
                                s_no = final_report.at[idx, 's.no']
                                verification_errors.append({
                                    's_no': s_no,
                                    'description': desc,
                                    'vendor': vendor_suffix,
                                    'old': curr_a,
                                    'new': expected_a
                                })
                                
    return final_report, verification_errors


def process_quotations(ref_file, vendor_files, sheet_config=None) -> tuple[pd.DataFrame, dict]:
    """
    Args:
        ref_file: Path (str) or file-like object for reference.
        vendor_files: List of (name, file_object) or (name, path).
        sheet_config: Dict mapping filename (or vendor name) to sheet_name.
    """
    if sheet_config is None:
        sheet_config = {}

    print("Loading Reference...")
    try:
        # Determine Reference Sheet
        # If ref_file is a path, we might use the filename key
        # If it's a file object, we might use its name attribute if available, or a specific key 'reference'
        
        # For simplicity, let's assume sheet_config has a key 'reference' for the reference file
        ref_sheet = sheet_config.get('reference')
        
        ref_df, schema = load_reference(ref_file, sheet_name=ref_sheet)
    except Exception as e:
        raise ValueError(f"Fatal Error loading reference: {e}")

    # ... (init report logic stays same) ...
    # Initialize report
    report_df = ref_df.copy()
    cols_to_keep = [c for c in report_df.columns if c in ['s.no', 'description', 'qty', 'unit']]
    if not cols_to_keep:
        cols_to_keep = report_df.columns.tolist()
        
    final_report = report_df[cols_to_keep].copy()

    print(f"Processing {len(vendor_files)} Vendors...")
    
    for vendor_name, vendor_file in vendor_files:
        print(f"  - {vendor_name}...", end="")
        try:
            # Determine Vendor Sheet
            # Using vendor_name as key is safest if we construct the config that way
            v_sheet = sheet_config.get(vendor_name)
            
            vendor_df = load_vendor(vendor_file, sheet_name=v_sheet)
            aligned_df = align_vendor_to_reference(ref_df, vendor_df, schema)
            
            final_report[f'Rate_{vendor_name}'] = aligned_df['Rate']
            final_report[f'Amount_{vendor_name}'] = aligned_df['Amount']
            print(" Done.")
        except Exception as e:
            print(f" Failed: {e}")

    # Backfill Total Row Logic
    # 1. Identify "Total" row (Strict check to avoid matching "Total Concrete" etc.)
    # We look for "Total", "Grand Total", "Total Amount" (case-insensitive, stripped)
    clean_desc = final_report['description'].astype(str).str.lower().str.strip()
    is_traceable_total = clean_desc.isin(['total', 'grand total', 'total amount'])
    
    if is_traceable_total.any():
        # Assume the LAST match is the main Total row (usually at bottom)
        # using index[-1] instead of [0] just in case there are multiple 'Total's? 
        # Actually usually there's just one Grand Total. Let's stick to strict matching.
        total_idx = final_report[is_traceable_total].index[0]
        
        for col in final_report.columns:
            if col.startswith('Amount_'):
                current_val = final_report.at[total_idx, col]
                
                # Check if missing (None, NaN, or empty string)
                if pd.isna(current_val) or current_val == '':
                    # Calculate sum of all OTHER rows
                    series = pd.to_numeric(final_report.loc[~is_traceable_total, col], errors='coerce')
                    calc_sum = series.sum()
                    
                    # Backfill
                    final_report.at[total_idx, col] = calc_sum

    # Phase 5: Advanced Analytics & Highlighting
    # -------------------------------------------------------------------------
    
    # --- Amount Verification & Correction ---
    final_report, verification_errors = verify_amounts(final_report)

    metadata = {
        'valid_rows': [],
        'min_max': {},       # {row_idx: {'min': val, 'max': val}}
        'incomplete_vendors': [],
        'verification_errors': verification_errors,
        'header_row_index': schema.header_row_index # Propagate detect header location
    }
    
    # 1. Identify Valid Items (Qty is numeric > 0 OR acceptable text like "L.S.")
    # We need to operate on specific columns.
    
    qty_series = pd.to_numeric(final_report['qty'], errors='coerce')
    
    # Check for non-numeric non-empty strings in 'qty'
    # Treat them as valid for analysis (we just can't verify amounts mathematically)
    raw_qty = final_report['qty'].astype(str).str.strip().str.lower()
    
    # Valid if: Numeric > 0 OR (Not Numeric AND Not Empty/Nan)
    is_numeric_valid = (qty_series.notna()) & (qty_series > 0)
    is_text_valid = (qty_series.isna()) & (raw_qty != 'nan') & (raw_qty != '') & (raw_qty != 'none')
    
    valid_mask = is_numeric_valid | is_text_valid
    
    metadata['valid_rows'] = final_report.index[valid_mask].tolist()
    
    # Identify Vendor Rate Columns
    rate_cols = [c for c in final_report.columns if c.startswith("Rate_")]
    vendor_names = [c.replace("Rate_", "") for c in rate_cols]
    
    # 2. Compute Min/Max Rates per Valid Item & Identify Vendors
    # We iterate only over valid rows to determine highlighting stats AND population of new columns
    
    # Initialize new columns
    final_report['Lowest_Vendor'] = ""
    final_report['Highest_Vendor'] = ""

    for idx in metadata['valid_rows']:
        # Extract rates for this row
        row_rates = []     # List of values for stats
        rate_map = []      # List of (val, vendor_name) for identification
        
        for v_col in rate_cols:
            val = final_report.at[idx, v_col]
            v_name = v_col.replace("Rate_", "")
            
            # Check if it's a valid number
            try:
                f_val = float(val)
                if pd.notna(f_val):
                     row_rates.append(f_val)
                     rate_map.append((f_val, v_name))
            except:
                pass
        
        if row_rates:
            min_val = min(row_rates)
            max_val = max(row_rates)
            
            metadata['min_max'][idx] = {
                'min': min_val,
                'max': max_val
            }
            
            # Identify which vendors have these values (Handle ties)
            min_vendors = [v for r, v in rate_map if r == min_val]
            max_vendors = [v for r, v in rate_map if r == max_val]
            
            final_report.at[idx, 'Lowest_Vendor'] = ", ".join(min_vendors)
            final_report.at[idx, 'Highest_Vendor'] = ", ".join(max_vendors)
            
    # 3. Identify Incomplete Vendors
    # A vendor is incomplete if they have ANY missing/zero rate for a VALID item.
    # (Assuming Quote means Rate > 0)
    
    for v_col in rate_cols:
        v_name = v_col.replace("Rate_", "")
        
        # Check rates for all valid rows
        # We extract the subset of this column for valid rows
        # Ensure we treat empty strings or weird chars as NaN
        subset = pd.to_numeric(final_report.loc[valid_mask, v_col], errors='coerce')
        
        # If any is NaN (missing) OR Zero (if we assume Rate must be > 0), tag as incomplete
        # User requested "not null quantity" must have items.
        # We'll treat 0.0 as incomplete too, as free items are rare in this context or usually explicit.
        if subset.isna().any() or (subset <= 0).any():
            metadata['incomplete_vendors'].append(v_name)


    # -------------------------------------------------------------------------
    # Final Cleanup
    # -------------------------------------------------------------------------

    # Ensure S.No is string to prevent PyArrow mixed-type errors in Streamlit
    s_no_col = next((c for c in final_report.columns if 's.no' in str(c).lower()), None)
    if s_no_col:
        final_report[s_no_col] = final_report[s_no_col].astype(str)

    # Ensure Qty is string (mixed int/str like '10' vs 'L.S') matches are causing issues
    qty_col = next((c for c in final_report.columns if 'qty' in str(c).lower()), None)
    if qty_col:
        final_report[qty_col] = final_report[qty_col].astype(str)

    return final_report, metadata