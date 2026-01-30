import pandas as pd
from rapidfuzz import fuzz
from schema import QuotationSchema

def align_vendor_to_reference(reference_df: pd.DataFrame, vendor_df: pd.DataFrame, schema: QuotationSchema) -> pd.DataFrame:
    """
    Aligns vendor dataframe rows to the reference dataframe.
    """
    
    # Initialize aligned dataframe with same index as reference
    # We want columns like Rate, Amount (which are typically in vendor_value_columns)
    aligned_data = []
    
    # Identify target columns in vendor df
    # We look for 'rate' and 'amount' specifically, or use schema.vendor_value_columns if they map directly
    # But usually vendor columns might be named slightly differently, e.g. "Rate (INR)", "Amount"
    
    # Let's find the best column match for "rate" and "amount" in the vendor df
    vendor_cols = [str(c).lower() for c in vendor_df.columns]
    
    rate_col = next((c for c in vendor_cols if 'rate' in c), None)
    amount_col = next((c for c in vendor_cols if 'amount' in c), None)
    
    # If we can't find rate/amount, we can't extract much
    if not rate_col or not amount_col:
        print(f"Warning: Could not identify Rate/Amount columns in vendor file. Found: {vendor_cols}")
        # Return empty df with correct index
        return pd.DataFrame(index=reference_df.index, columns=['Rate', 'Amount'])



    # Initialize set to track used vendor rows
    used_vendor_indices = set()

    # Iterate through reference rows
    for idx, ref_row in reference_df.iterrows():
        ref_desc = str(ref_row.get('description', '')).lower()
        ref_sr = str(ref_row.get('s.no', '')).lower().replace('.0', '').strip()
        
        best_match_idx = None
        best_match_score = 0
        
        # Strategy 1: Exact Sr No Match + Description Sanity Check
        vendor_sr_col = next((c for c in vendor_cols if 's.no' in c or 'sr' in c), None)
        vendor_desc_col = next((c for c in vendor_cols if 'description' in c or 'particular' in c), None)
        
        if vendor_sr_col and ref_sr not in ['nan', '', 'none']:
            # Find all candidates with matching Sr matching AND NOT USED
            v_srs = vendor_df[vendor_sr_col].astype(str).str.lower().str.replace('.0', '').str.strip()
            
            # Simple iteration to check usage
            # Filter candidates that are not in used_vendor_indices
            candidates = vendor_df[v_srs == ref_sr]
            candidates = candidates[~candidates.index.isin(used_vendor_indices)]
            
            if not candidates.empty:
                # If multiple matches (or even one), find the one with best description overlap
                if vendor_desc_col:
                    best_cand_score = -1
                    best_cand_idx = None
                    
                    for c_idx, c_row in candidates.iterrows():
                        v_desc = str(c_row[vendor_desc_col]).lower()
                        score = fuzz.ratio(ref_desc, v_desc)
                        if score > best_cand_score:
                            best_cand_score = score
                            best_cand_idx = c_idx
                    
                    # SANITY CHECK: 
                    if best_cand_score > 40:
                        best_match_idx = best_cand_idx
                        # best_match_score = 100 # Sr match is prioritized
                else:
                    # No description column? Just take first.
                    best_match_idx = candidates.index[0]
                    # best_match_score = 100
        
        # Strategy 2: Fuzzy Description Match (if no valid Sr match found)
        # Only run if we didn't find a good Sr match
        if best_match_idx is None:
            if vendor_desc_col:
                for v_idx, v_row in vendor_df.iterrows():
                    if v_idx in used_vendor_indices:
                        continue
                        
                    v_desc = str(v_row[vendor_desc_col]).lower()
                    score = fuzz.ratio(ref_desc, v_desc)
                    
                    # Higher threshold for pure fuzzy to avoid garbage matching
                    if score > 80 and score > best_match_score: 
                        best_match_score = score
                        best_match_idx = v_idx

        
        # Extract data if match found
        if best_match_idx is not None:
            used_vendor_indices.add(best_match_idx)
            
            rate_val = vendor_df.at[best_match_idx, rate_col]
            amount_val = vendor_df.at[best_match_idx, amount_col]
            
            # Treat 0 as NaN
            if pd.to_numeric(rate_val, errors='coerce') == 0:
                rate_val = None
            if pd.to_numeric(amount_val, errors='coerce') == 0:
                amount_val = None
        else:
            rate_val = None
            amount_val = None
            
        aligned_data.append({
            'Rate': rate_val,
            'Amount': amount_val
        })


        
    return pd.DataFrame(aligned_data, index=reference_df.index)
