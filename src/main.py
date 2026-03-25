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


def process_quotations(ref_file, vendor_files, sheet_config=None, estimated_file=None) -> tuple[pd.DataFrame, dict]:
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

    if estimated_file:
        print("Processing Estimated Rates...", end="")
        try:
            est_sheet = sheet_config.get('estimated_rate')
            est_df = load_vendor(estimated_file, sheet_name=est_sheet)
            aligned_est = align_vendor_to_reference(ref_df, est_df, schema)
            final_report['Estimated_Rate'] = aligned_est['Rate']
            print(" Done.")
        except Exception as e:
            print(f" Failed: {e}")
            estimated_file = None

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
        'min_max': {},               # {row_idx: {'min': val, 'max': val}}
        'min_max_amounts': {},       # {row_idx: {'min': val, 'max': val}}
        'closest_vendors': {},       # {row_idx: [list of vendor names]}
        'incomplete_vendors': [],
        'verification_errors': verification_errors,
        'header_row_index': schema.header_row_index # Propagate detect header location
    }
    
    # 1. Identify Valid Items
    qty_series = pd.to_numeric(final_report['qty'], errors='coerce')
    raw_qty = final_report['qty'].astype(str).str.strip().str.lower()
    
    is_numeric_valid = (qty_series.notna()) & (qty_series > 0)
    is_text_valid = (qty_series.isna()) & (raw_qty != 'nan') & (raw_qty != '') & (raw_qty != 'none')
    
    valid_mask = is_numeric_valid | is_text_valid
    metadata['valid_rows'] = final_report.index[valid_mask].tolist()
    
    # Identify Columns
    rate_cols = [c for c in final_report.columns if c.startswith("Rate_")]
    amount_cols = [c for c in final_report.columns if c.startswith("Amount_")]
    vendor_names = [c.replace("Rate_", "") for c in rate_cols]
    
    # Initialize calculated metrics
    final_report['Lowest_Vendor'] = ""
    final_report['Highest_Vendor'] = ""
    if estimated_file:
        final_report['Closest_to_Estimate'] = ""

    # Iterate over all rows for exhaustive Min/Max and Estimated Rate analytics
    for idx in final_report.index:
        is_item = idx in metadata['valid_rows']
        
        if is_item:
            # IT IS AN ITEM (Analyze Rates)
            row_rates = []
            rate_map = []
            
            for v_col in rate_cols:
                val = final_report.at[idx, v_col]
                v_name = v_col.replace("Rate_", "")
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
                
                metadata['min_max'][idx] = {'min': min_val, 'max': max_val}
                
                min_vendors = [v for r, v in rate_map if r == min_val]
                max_vendors = [v for r, v in rate_map if r == max_val]
                
                final_report.at[idx, 'Lowest_Vendor'] = ", ".join(min_vendors)
                final_report.at[idx, 'Highest_Vendor'] = ", ".join(max_vendors)
            
            # --- Estimate Comparison Logic ---
            if estimated_file and 'Estimated_Rate' in final_report.columns:
                est_val_raw = final_report.at[idx, 'Estimated_Rate']
                try:
                    est_val = float(est_val_raw)
                    if pd.notna(est_val) and row_rates:
                        # Find closest rate
                        closest_diff = float('inf')
                        closest_vs = []
                        for r, v in rate_map:
                            diff = abs(r - est_val)
                            # Floating point comparison safety margin
                            if abs(diff - closest_diff) < 1e-9:
                                closest_vs.append(v)
                            elif diff < closest_diff:
                                closest_diff = diff
                                closest_vs = [v]
                        
                        metadata['closest_vendors'][idx] = closest_vs
                        final_report.at[idx, 'Closest_to_Estimate'] = ", ".join(closest_vs)
                except Exception:
                    pass  # Missing or invalid estimated rate for this item

        else:
            # NOT AN ITEM: Possibly a Subtotal or Total (Analyze Amounts)
            row_amounts = []
            amt_map = []
            
            for a_col in amount_cols:
                val = final_report.at[idx, a_col]
                v_name = a_col.replace("Amount_", "")
                try:
                    f_val = float(val)
                    if pd.notna(f_val):
                         row_amounts.append(f_val)
                         amt_map.append((f_val, v_name))
                except:
                    pass
            
            # If the row has valid amounts, track min/max for amounts instead of rates
            if row_amounts:
                min_amt = min(row_amounts)
                max_amt = max(row_amounts)
                
                # Check for completely blank rows where sum is 0
                if max_amt > 0 or min_amt < 0:
                    metadata['min_max_amounts'][idx] = {'min': min_amt, 'max': max_amt}
                    
                    min_vendors = [v for a, v in amt_map if a == min_amt]
                    max_vendors = [v for a, v in amt_map if a == max_amt]
                    
                    final_report.at[idx, 'Lowest_Vendor'] = ", ".join(min_vendors)
                    final_report.at[idx, 'Highest_Vendor'] = ", ".join(max_vendors)

    # 3. Identify Incomplete Vendors
    for v_col in rate_cols:
        v_name = v_col.replace("Rate_", "")
        subset = pd.to_numeric(final_report.loc[valid_mask, v_col], errors='coerce')
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
        
    # Ensure Rate and Amount columns are purely numeric to prevent PyArrow mixed-type crashes (e.g. from '-')
    for col in final_report.columns:
        if str(col).startswith('Rate_') or str(col).startswith('Amount_') or col == 'Estimated_Rate':
            final_report[col] = pd.to_numeric(final_report[col], errors='coerce')

    return final_report, metadata