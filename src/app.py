import streamlit as st
import pandas as pd
from io import BytesIO
from main import process_quotations
import os

from utils import setup_tender_structure, create_session_workspace, detect_boq_index
import shutil
from excel_utils import apply_formatting, overlay_analysis_on_template, apply_openpyxl_highlighting

st.set_page_config(page_title="Quotalyze", layout="wide")

st.title("📊 Quotalyze")
st.markdown("Manage tenders and generate comparison reports.")

# --- Session Initialization ---
if 'workspace_root' not in st.session_state:
    st.session_state['workspace_root'] = create_session_workspace()

# Handle Project Reset (triggered from bottom)
if st.session_state.get('reset_project_flag'):
    st.session_state['project_name_input'] = ""
    st.session_state['reset_project_flag'] = False

# Handle Restore (triggered from restore button)
if st.session_state.get('restore_pending_name'):
    st.session_state['project_name_input'] = st.session_state['restore_pending_name']
    del st.session_state['restore_pending_name']
    
# Main Output
with st.expander("⚠️ Important Usage Guidelines", expanded=True):
    st.markdown("""
    **APPLICATION CONSTRAINTS:**
    1. **No External Links**: Ensure Reference & Vendor files contain *values only*, not formulas linking to other files.
    2. **Table Headers**: Reference file main table must have correct headers (S.No, Description, Qty, Unit).
    3. **Vendor Naming**: Name your vendor files using the Vendor's Name (e.g., `VendorA.xlsx`), as this name will appear in reports.
    4. **Partial Quotations**: If a Quantity is present (not null) and any vendor has not quoted for that particular item, an alert will be generated stating "All vendors have not quoted for all items".
    """)

tender_name = st.text_input("Project / Tender Name", value="", key="project_name_input")

if not tender_name:
    st.info("👋 Please enter a Project Name to begin.")
    st.stop()

# Ensure structure exists in SESSION WORKSPACE
paths = setup_tender_structure(tender_name, root_path=st.session_state['workspace_root'])

reference_dir = paths['reference']
vendor_dir = paths['vendors']
output_dir = paths['output']

# --- Sidebar: File Manager ---
with st.sidebar:
    st.header(f"📂 Project: {tender_name}")
    
    st.divider()
    
    # 2. Start Over (Clear Workspace)
    st.subheader("⚠️ Actions")

    if st.button("🗑️ Start Over (Clear Session)", type="secondary", key="btn_clear_session"):
        st.session_state['confirm_clear'] = True
        
    if st.session_state.get('confirm_clear'):
        st.warning(f"Are you sure? This will delete all temporary files.", icon="⚠️")
        c_yes, c_no = st.columns(2)
        if c_yes.button("Yes, Clear Everything", type="primary"):
            # Clear Project Folder Only (Keep workspace root for next project?)
            # Or just delete the whole workspace content.
            try:
                if os.path.exists(paths['base']):
                   shutil.rmtree(paths['base'])
            except Exception as e:
                pass
            
            # Reset UI via Rerun Flag
            # We cannot set widget state here directly because it's already rendered.
            st.session_state['reset_project_flag'] = True

            # Clear other state
            for key in list(st.session_state.keys()):
                if key not in ["project_name_input", "reset_project_flag", "workspace_root"]:
                     del st.session_state[key]
            st.rerun()
            
        if c_no.button("Cancel", key="cancel_clear"):
            st.session_state['confirm_clear'] = False
            st.rerun()

reference_dir = paths['reference']
existing_ref_files = [f for f in os.listdir(reference_dir) if not f.startswith('.')] if os.path.exists(reference_dir) else []

# --- Layout: Single Column Vertical Stack ---

# 1. Reference File Section
st.subheader("Reference File")

# Display Current
if existing_ref_files:
    # Should be only 1 if we enforce policy
    current_ref = existing_ref_files[0]
    st.success(f"📄 Using: **{current_ref}**")
    selected_ref_file_path = os.path.join(reference_dir, current_ref)
else:
    st.info("No reference file.")
    selected_ref_file_path = None

# Uploader (Replace) - Use Dynamic Key to allow reset
if 'ref_uploader_key_id' not in st.session_state:
    st.session_state['ref_uploader_key_id'] = 0

ref_key = f"ref_uploader_{st.session_state['ref_uploader_key_id']}"
reference_file = st.file_uploader("Upload Reference (Replaces current)", type=["xlsx", "xls"], key=ref_key)

if reference_file:
    # Delete ALL existing
    for f in existing_ref_files:
        try:
            os.remove(os.path.join(reference_dir, f))
        except: pass
    
    # Save new
    save_path = os.path.join(reference_dir, reference_file.name)
    with open(save_path, "wb") as f:
        f.write(reference_file.getbuffer())
    
    # Reset uploader & analysis
    st.session_state['ref_uploader_key_id'] += 1
    st.session_state['analysis_done'] = False
    st.rerun()

st.divider()

# 2. Vendor Quotations Section
st.subheader("Vendor Quotations")

# Show existing vendors "Ready for Processing"
vendor_dir = paths['vendors']
existing_vendor_files = [f for f in os.listdir(vendor_dir) if not f.startswith('.')] if os.path.exists(vendor_dir) else []

# --- Vendor UI ---

# 1. Display Managed List (Files on Disk)
st.markdown("### 🗂️ Manage Vendor Files")
if existing_vendor_files:
    for v_file in sorted(existing_vendor_files):
        col_f, col_del = st.columns([0.85, 0.15])
        col_f.text(f"📄 {v_file}")
        # Use a unique key for delete button
        if col_del.button("🗑️", key=f"del_v_{v_file}", help=f"Delete {v_file}"):
            try:
                os.remove(os.path.join(vendor_dir, v_file))
            except: pass
            # Clear analysis results
            st.session_state['analysis_done'] = False
            st.session_state['output_generated'] = False
            st.rerun()
else:
    st.info("No vendor files uploaded yet.")

st.divider()

# 2. Uploader (Immediate Upload)
st.markdown("### 📤 Add Vendor Files")
if 'uploader_key_id' not in st.session_state:
    st.session_state['uploader_key_id'] = 0
    
uploader_key = f"vendor_uploader_{st.session_state['uploader_key_id']}"

new_uploads = st.file_uploader("Upload Vendor Excels", type=["xlsx", "xls"], accept_multiple_files=True, key=uploader_key, label_visibility="collapsed")

if new_uploads:
    # Save immediately and reset
    for vf in new_uploads:
            v_path = os.path.join(vendor_dir, vf.name)
            with open(v_path, "wb") as f:
                f.write(vf.getbuffer())
    
    # Reset uploader to clear list
    st.session_state['uploader_key_id'] += 1
    st.rerun()


# --- Configuration (Sheet Selection) ---
# We need to know valid files to show sheet selection.
# Aggregate all sources.

valid_ref_source = reference_file if reference_file else selected_ref_file_path
valid_vendor_sources = [] # list of (name, path_or_filelike)

# Logic to aggregate vendors
# Use a set to track names processing order
seen_vendors = set()

# 1. Existing
for f in existing_vendor_files:
    clean_name = os.path.splitext(f)[0]
    if clean_name not in seen_vendors:
        valid_vendor_sources.append((f, os.path.join(vendor_dir, f)))
        seen_vendors.add(clean_name)



if valid_ref_source and valid_vendor_sources:
    st.divider()
    with st.expander("⚙️  Sheet Configuration", expanded=True):
        st.write("Select the sheet containing the quotation data for each file.")
        
        sheet_config = {}
        


        # 1. Reference Sheet
        try:
            # Reference Sheet Logic
            if isinstance(valid_ref_source, str):
                xl = pd.ExcelFile(valid_ref_source)
            else:
                valid_ref_source.seek(0)
                xl = pd.ExcelFile(valid_ref_source)
                valid_ref_source.seek(0)
            
            ref_idx = detect_boq_index(xl.sheet_names)
            selected_sheet = st.selectbox(f"Reference: {os.path.basename(valid_ref_source) if isinstance(valid_ref_source, str) else valid_ref_source.name}", xl.sheet_names, index=ref_idx, key="ref_sheet")
            sheet_config['reference'] = selected_sheet
        except Exception as e:
            st.error(f"Error reading Reference: {e}")

        st.divider()

        # 2. Vendor Sheets
        st.write("Vendor Files:")
        for v_name, v_source in valid_vendor_sources:
            try:
                # Clean name for key
                clean_v_name = os.path.splitext(v_name)[0]
                
                if isinstance(v_source, str):
                    v_xl = pd.ExcelFile(v_source)
                else:
                    v_source.seek(0)
                    v_xl = pd.ExcelFile(v_source)
                    v_source.seek(0)
                    
                v_idx = detect_boq_index(v_xl.sheet_names)
                selected_v_sheet = st.selectbox(f"{clean_v_name}", v_xl.sheet_names, index=v_idx, key=f"v_sheet_{clean_v_name}")
                sheet_config[clean_v_name] = selected_v_sheet
                
            except Exception as e:
                    st.error(f"Error reading {v_name}: {e}")
    
    st.divider()
    
    # Action Button
    if st.button("Generate Analysis Report", type="primary"):
        # Reset State specific to this new run
        st.session_state['analysis_done'] = False
        st.session_state['output_generated'] = False
        
        with st.status("Processing...", expanded=True) as status:
                
                # Save NEW files to disk -> REMOVED (Handled by Immediate Upload)
                
                # Reference Path
                final_ref_path = ""
                if reference_file:
                     # If uploader has file, save it (Reference still uses old pattern? 
                     # Actually ref uses precedence pattern. Let's keep it, but valid_ref_source has the info)
                     # Wait, if reference_file is set, we use it.
                     save_path = os.path.join(reference_dir, reference_file.name)
                     with open(save_path, "wb") as f:
                        f.write(reference_file.getbuffer())
                     final_ref_path = save_path
                else:
                     final_ref_path = selected_ref_file_path
                
                # Vendor Paths
                # We already have 'valid_vendor_sources' populated from DISK files in the config section above.
                # We just need to convert them to the format process_quotations expectations.
                # process_quotations expects list of (name, path).
                
                final_vendor_list = []
                for v_name, v_source in valid_vendor_sources:
                     # valid_vendor_sources is list of (filename, full_path)
                     clean_name = os.path.splitext(v_name)[0]
                     final_vendor_list.append((clean_name, v_source))

                
                status.write("Files ready. Starting analysis...")
                
                # Cleanup Old Reports (Ensure only the latest exists)
                if os.path.exists(output_dir):
                    for f in os.listdir(output_dir):
                        if f.endswith(".xlsx") or f.endswith(".xls"):
                            try:
                                os.remove(os.path.join(output_dir, f))
                            except Exception as e:
                                print(f"Cleanup error: {e}")
                
                # Run Process
                try:
                    # Note: process_quotations expects 'reference' key for ref file
                    # and vendor_name keys for vendors.
                    
                    final_report, metadata = process_quotations(final_ref_path, final_vendor_list, sheet_config=sheet_config)
                    
                    # Store results in Session State for persistence
                    st.session_state['analysis_done'] = True
                    st.session_state['final_report'] = final_report
                    st.session_state['metadata'] = metadata
                    st.session_state['paths'] = paths
                    st.session_state['ref_save_path'] = final_ref_path
                    st.session_state['sheet_config'] = sheet_config
                    
                    status.write("Analysis Complete!")
                    st.rerun() # Force reload to update sidebar and show results
                    
                except Exception as e:
                    st.error(f"Processing Error: {e}")
                    status.update(label="Failed", state="error")

    # --- Results Display (Persistent) ---
    if st.session_state.get('analysis_done'):
        final_report = st.session_state['final_report']
        metadata = st.session_state['metadata']
        paths = st.session_state['paths']
        ref_save_path = st.session_state['ref_save_path']
        sheet_config = st.session_state['sheet_config']

        st.subheader("Analysis Results")
        
        # --- Verification Alerts ---
        verification_errors = metadata.get('verification_errors', [])
        if verification_errors:
            count = len(verification_errors)
            st.warning(f"⚠️ Verification Found {count} Data Discrepancies")
            st.write("The following Amounts were incorrect in the provided files and have been corrected based on (Qty * Rate):")
            
            # Convert to minimal DF for display
            err_df = pd.DataFrame(verification_errors)
            # columns: s_no, description, vendor, old, new
            st.dataframe(err_df, hide_index=True)
        else:
            st.success("✅ Amount Verification Passed: All Vendor amounts match (Qty * Rate).")

        # --- Incomplete Vendor Check ---
        incomplete = metadata.get('incomplete_vendors', [])
        if incomplete:
            st.warning(f"⚠️ Partial Quotations: The following vendors did not quote for all items: {', '.join(incomplete)}")
        
        # --- Output Generation for Download ---
        # Rename output file to Project Name
        report_filename = f"{tender_name}_report.xlsx"
        output_path = os.path.join(paths['output'], report_filename)
        
        if not st.session_state.get('output_generated'):
            # --- Advanced Styling & Save (One Time) ---
            def highlight_min_max(df):
                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                min_max_data = metadata.get('min_max', {})
                for row_idx, stats in min_max_data.items():
                    r_min = stats['min']
                    r_max = stats['max']
                    for col in df.columns:
                        if col.startswith("Rate_"):
                            val = df.at[row_idx, col]
                            try:
                                f_val = float(val)
                                if f_val == r_min:
                                    styles.at[row_idx, col] = 'background-color: lightgreen; color: black;'
                                elif f_val == r_max:
                                    styles.at[row_idx, col] = 'background-color: lightcoral; color: black;'
                            except:
                                pass
                return styles

            # Apply UI styles for preview
            styled_df = final_report.style.apply(highlight_min_max, axis=None).format(na_rep="", precision=2)

            # Save Styled Excel using Template Overlay
            try:
                start_row = metadata.get('header_row_index', 0)
                ref_sheet_name = sheet_config.get('reference')
                
                # Step 1: Overlay Data
                overlay_analysis_on_template(ref_save_path, output_path, final_report, start_row, ref_sheet_name)
                
                # Step 2: Apply Highlighting
                apply_openpyxl_highlighting(output_path, metadata, start_row, final_report.columns, sheet_name=ref_sheet_name)
                
                # Mark as done
                st.session_state['output_generated'] = True
                
            except Exception as e:
                st.error(f"Export Error: {e}")
                # Fallback
                final_report.to_excel(output_path, index=False)

        else:
             # Re-create styled df for preview if needed, or just show raw
             # Let's recreate style for consistency
            def highlight_min_max(df):
                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                min_max_data = metadata.get('min_max', {})
                for row_idx, stats in min_max_data.items():
                    r_min = stats['min']
                    r_max = stats['max']
                    for col in df.columns:
                        if col.startswith("Rate_"):
                            val = df.at[row_idx, col]
                            try:
                                f_val = float(val)
                                if f_val == r_min:
                                    styles.at[row_idx, col] = 'background-color: lightgreen; color: black;'
                                elif f_val == r_max:
                                    styles.at[row_idx, col] = 'background-color: lightcoral; color: black;'
                            except:
                                pass
                return styles
            styled_df = final_report.style.apply(highlight_min_max, axis=None).format(na_rep="", precision=2)

        
        # Display Preview
        st.dataframe(styled_df)
        
        # Download Button
        if os.path.exists(output_path):
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Analysis Report",
                    data=f,
                    file_name=report_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
        else:
            st.error("Report file missing.")

