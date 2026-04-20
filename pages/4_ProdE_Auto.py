import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
import os
import shutil
import tempfile
from datetime import datetime
from io import BytesIO

# Set page configuration
st.set_page_config(page_title="ProdE Automation", page_icon=None, layout="wide")

# Apply consistent premium design
set_premium_style()


# Check for win32com availability
try:
    import win32com.client
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    st.warning(" pywin32 is not installed. Excel COM automation will not work. Install with: pip install pywin32")

# --- FILE PATH CONFIGURATION ---
TEMPLATE_PATH = r"C:\Users\SPM\Downloads\BPI\Template\MADPL TEMPLATE v3 ETO MUNA v1.xlsm"
SAVE_PATH = r"C:\Users\SPM\Downloads\BPI\Productivity Report_ Every Monday\DRR"

# Create folder if it doesn't exist
if not os.path.exists(SAVE_PATH):
    try:
        os.makedirs(SAVE_PATH, exist_ok=True)
    except Exception:
        pass

def process_template(df_cleaned, tmpl_path, output_path):
    """Copy template and use Excel COM to write data and replicate formulas."""
    if not WIN32COM_AVAILABLE:
        st.error("pywin32 is required but not installed.")
        return False
    
    import win32com.client
    import pythoncom
    import time
    
    try:
        # Check if source file exists
        if not os.path.exists(tmpl_path):
            st.error(f"Template file not found: {tmpl_path}")
            return False
        
        # Copy template
        shutil.copy2(tmpl_path, output_path)
        
        # Use Excel COM automation
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            abs_path = os.path.abspath(output_path)
            wb = excel.Workbooks.Open(abs_path)
            
            # Select target sheet
            try:
                ws = wb.Sheets("VOLARE EXTRACTION")
            except Exception:
                st.error("Sheet 'VOLARE EXTRACTION' not found in the template.")
                return False

            num_rows = len(df_cleaned)
            num_cols = len(df_cleaned.columns)
            
            # 1. Paste cleaned data starting from A3
            # Convert dataframe to values
            data_values = df_cleaned.values.tolist()
            
            # Handle timestamps/dates for Excel COM compatibility
            for r_idx in range(len(data_values)):
                for c_idx in range(len(data_values[r_idx])):
                    val = data_values[r_idx][c_idx]
                    if pd.isna(val):
                        data_values[r_idx][c_idx] = ""
                    elif isinstance(val, (pd.Timestamp, datetime)):
                        data_values[r_idx][c_idx] = val.strftime('%m/%d/%Y')

            st.write(f"Pasting {num_rows} rows into 'VOLARE EXTRACTION' starting at A3...")
            start_cell = ws.Cells(3, 1)
            end_cell = ws.Cells(3 + num_rows - 1, num_cols)
            ws.Range(start_cell, end_cell).Value = data_values

            # 2. Formula Replication (AZ to BL)
            # AZ is column 52 (index starting at 1 for Excel)
            # BL is column 64
            # We copy formulas from Row 2 and apply them to Rows 3 to 3+num_rows-1
            if num_rows > 0:
                st.write("Replicating formulas from Row 1 for columns AZ:BL...")
                last_data_row = 3 + num_rows - 1
                
                # Column AZ to BL
                # AZ in Excel is 52, BL is 64
                # User specified to replicate from Row 1
                formula_range_source = ws.Range(ws.Cells(1, 52), ws.Cells(1, 64))
                formula_range_dest = ws.Range(ws.Cells(3, 52), ws.Cells(last_data_row, 64))
                
                # Copy formulas from Row 1 to the rest
                formula_range_source.Copy()
                formula_range_dest.PasteSpecial(Paste=-4123) # xlPasteFormulas
                excel.Application.CutCopyMode = False

                
                # Optional: Bake formulas if needed, but usually templates prefer keeping them
                # For consistency with Daily Prod's "bake" logic, we can bake if asked, 
                # but "copy the formula" usually implies keeping them active.
                # However, Daily Prod bakes them for performance/stability.
                # Let's keep them as formulas for now unless performance is an issue.

            # Save and close
            wb.Save()
            wb.Close(False)
            return True
            
        except Exception as e:
            st.error(f"Error in Excel processing: {e}")
            return False
        finally:
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
            
    except Exception as e:
        st.error(f"Error in process_template: {e}")
        return False

def main():
    st.title(" ProdE Efforts Automation")
    st.markdown("Automated cleaning and template integration for Production Efforts.")

    uploaded_file = st.file_uploader(" Upload Raw ProdE File (Excel or CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        file_ext = uploaded_file.name.split('.')[-1]
        
        with st.spinner("Reading data..."):
            try:
                if file_ext == 'csv':
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                df.columns = df.columns.str.strip()
            except Exception as e:
                st.error(f"Error loading file: {e}")
                st.stop()

        # --- DATA CLEANING (STRICTLY FROM 4_ProdE_Clean.py) ---
        st.subheader(" Cleaning Data")
        
        orig_count = len(df)
        
        # 1. Filter Status: Remove rows with "BP"
        if 'Status' in df.columns:
            df = df[df['Status'].astype(str) != 'BP']
        
        # 2. Filter Remark By: Remove specific names
        if 'Remark By' in df.columns:
            excluded_remark_by = ["DCCAUNTE", "CMENRIQUEZ", "EDSUMAIT", "JSCANA", "NRPARAYAOAN"]
            mask = ~df['Remark By'].astype(str).str.upper().isin([name.upper() for name in excluded_remark_by])
            df = df[mask]
        
        # 3. Account No. formatting: Add six "0" and force as text (Column E / "Account No.")
        # Logic adapted from 1_Daily_Prod.py
        target_col = "Account No."
        if target_col in df.columns:
            df[target_col] = df[target_col].astype(str).str.replace(r'\.0$', '', regex=True)
            df[target_col] = df[target_col].replace(['nan', 'None'], '')
            # Drop empty rows where Account No. is missing
            df = df[df[target_col] != ""]
            # Prefix with '000000 to force text in Excel with leading zeroes
            df[target_col] = "'000000" + df[target_col]
        
        # 4. Remark cleaning (Column K): Remove leading "=" to prevent Excel formula errors
        col_remark = "Remark"
        if col_remark in df.columns:
            # We only remove the leading "=" sign if it exists
            df[col_remark] = df[col_remark].apply(lambda x: str(x)[1:] if str(x).startswith('=') else x)
        
        # 5. Delete last 2 columns
        if len(df.columns) > 2:


            df = df.iloc[:, :-2]
            
        # 4. Move "Next Call" to index 23
        column_to_move = "Next Call"
        target_position = 23
        
        found_col = None
        if column_to_move in df.columns:
            found_col = column_to_move
        else:
            for col in df.columns:
                if 'next' in str(col).lower() and 'call' in str(col).lower():
                    found_col = col
                    break
        
        if found_col:
            col_data = df.pop(found_col)
            # Ensure target position is within bounds
            actual_target = min(target_position, len(df.columns))
            df.insert(actual_target, found_col, col_data)
        
        cleaned_count = len(df)
        st.success(f"Cleaned! Rows: {orig_count} → {cleaned_count}. Columns: {len(df.columns)}")
        
        with st.expander("Preview Cleaned Data"):
            st.dataframe(df.head(50))

        st.divider()
        
        # --- EXPORT SECTION ---
        st.subheader(" Export to Template")
        
        filename = f"ProdE_Automated_{datetime.now().strftime('%Y%m%d')}.xlsm"
        full_save_path = os.path.join(SAVE_PATH, filename)
        
        st.info(f"Report will be saved to:\n`{full_save_path}`")
        
        if st.button(" PROCESS AND SAVE TO TEMPLATE", type="primary", use_container_width=True):
            if not WIN32COM_AVAILABLE:
                st.error("Excel COM (pywin32) is not available on this system.")
            else:
                with st.spinner("Pasting data into template and replicating formulas..."):
                    success = process_template(df, TEMPLATE_PATH, full_save_path)
                    if success:
                        st.balloons()
                        st.success(f" Successfully saved to: {full_save_path}")
                        
                        # Provide a download button for the local file (copying it to a buffer)
                        with open(full_save_path, "rb") as f:
                            btn = st.download_button(
                                label=" Download Processed Report",
                                data=f,
                                file_name=filename,
                                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                            )

if __name__ == "__main__":
    main()
