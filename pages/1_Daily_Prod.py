import streamlit as st
import pandas as pd
import io
import os
import shutil
import tempfile
from datetime import datetime, timedelta

# Set page configuration
st.set_page_config(page_title="Daily Prod Automation", page_icon="⚡", layout="wide")

# Check for win32com availability
try:
    import win32com.client
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    st.warning("⚠️ pywin32 is not installed. Excel COM automation will not work. Install with: pip install pywin32")

# --- DATE HELPER FOR FILENAMES ---
def get_report_date():
    """Returns (report_date_obj, filename_string) based on user's Monday/Friday logic."""
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        report_date = today - timedelta(days=3) # Use Friday last week
    else:
        report_date = today - timedelta(days=1) # Use yesterday
    return report_date, report_date.strftime("%Y%m%d")

# --- FILE PATH CONFIGURATION ---
SAVE_PATH = r"C:\Users\SPM\Downloads\BPI\PL_Daily\DRR Auto"

# Create folder if it doesn't exist
if not os.path.exists(SAVE_PATH):
    try:
        os.makedirs(SAVE_PATH)
    except Exception:
        pass

# STATUSES TO EXCLUDE (matches anything starting with these strings)
EXCLUDED_STATUSES = [
    'CONFIRMED', 'PTP_FF', 'SMS SENT - T6 NO RESPONSE',
    'LS VIA SOCMED - T6 NO RESPONSE', 'SMS RECEIVED - UNDER NEGO',
    'SMS RECEIVED - WILL SETTLE', 'NEW', 'BP', 'ENCO', 'UNLOCKED',
    'BULK SMS', 'REACTIVE', 'LOCKED', 'PM',
    'SMS RECEIVED - NO SUCH PERSON (NSP)',
    'SERVICE REQUEST (SR) - FOR COLLECTION'
]

rfd_mapping = {
    "AWAITING_FUNDS": [
        "ALLOTMENT", "REMITTANCE", "ALLOWANCE", "LOANS", "PROFIT", 
        "SALARY", "INCENTIVES", "BACKPAY", "BENEFITS", "PENSION", 
        "COLLECTIONS", "COMMISSION"
    ],
    "EMPLOYMENT_STATUS": [
        'AWAITING TO ONBOARD', 'CONTRACTUAL', 'NEWLY-HIRED', 'CONFLICT WITH EMPLOYER', 
        'UNEMPLOYMENT', 'FREELANCE', 'SUSPENSION', 'ON MATERNITY LEAVE'
    ],
    "EMERGENCY": [
        'HOSPITALIZATION', 'VICTIM OF CALAMITY', 'ACCIDENT', 'DEATH OF RELATIVE'
    ],
    "NO_CAPACITY_TO_PAY": [
        "SHORT OF FUNDS", "FINANCIAL DIFFICULTY", "BUSINESS SLOWDOWN"
    ],
    "PRIORITIZE_OTHER_BILLS_EXPENSES": [
        "LOANS/OTHER DUES", "TUITION", "UTILITY BILLS", "MEDICATION"
    ],
    "LOAN_CLARIFICATIONS": [
        "PAST DUE BALANCE", "MONTHLY AMORTIZATION", "MATURITY DATE", "REMAINING TERM"
    ],
    "PAYMENT_CLARIFICATIONS": [
        "FUNDED ADA", "CLAIMING SETTLED ACCOUNT", "LATE CHARGES", "PAYMENT HISTORY"
    ],
    "PERSONAL_REASON": [
        "SCAM VICTIM", "OUT OF COUNTRY/TOWN", "BUSY SCHEDULE", 
        "CHILDBIRTH", "FAMILY MATTER", "ROBBERY VICTIM"
    ],
    "OTHERS": [
        "DROPPED CALL", "ONLINE BANKING CONCERN", "OVERLOOKED DUEDATE", "REFUSE TO DISCLOSED REASON"
    ]
}

def classify_remark(remark):
    if pd.isna(remark) or str(remark).strip() == "":
        return None, None
    
    remark_lower = str(remark).lower()
    
    if any(keyword in remark_lower for keyword in ["got scammed", "scammed", "scammer"]):
        return "PERSONAL_REASON", "SCAM VICTIM"
    if any(keyword in remark_lower for keyword in ['ootc', 'out of the country', 'abroad']):
        return 'PERSONAL_REASON', 'OUT OF COUNTRY/TOWN'
    if any(keyword in remark_lower for keyword in ['unemployed']):
        return 'EMPLOYMENT_STATUS', 'UNEMPLOYMENT'
    if any(keyword in remark_lower for keyword in ['dropped', 'dropped call', 'call drop', 'call dropped']):
        return 'OTHERS', 'DROPPED CALL'
    if any(keyword in remark_lower for keyword in ['overlooked the account', 'overlooked account']):
        return 'OTHERS', 'OVERLOOKED DUEDATE'
    if any(keyword in remark_lower for keyword in ['financial struggle', 'financial problem']):
        return 'NO_CAPACITY_TO_PAY', 'FINANCIAL DIFFICULTY'
    if any(keyword in remark_lower for keyword in ['other loans', 'other loan']):
        return 'PRIORITIZE_OTHER_BILLS_EXPENSES', 'LOANS/OTHER DUES'
    if any(keyword in remark_lower for keyword in ['bankcruptcy']):
        return 'NO_CAPACITY_TO_PAY', 'BUSINESS SLOWDOWN'
    if any(keyword in remark_lower for keyword in ['maternity leave']):
        return 'EMPLOYMENT_STATUS', 'ON MATERNITY LEAVE'
    if any(keyword in remark_lower for keyword in ['give birth', 'gave birth']):
        return 'EMPLOYMENT_STATUS', 'CHILDBIRTH'
    if any(keyword in remark_lower for keyword in ['medication','undermedication','hospitalization']):
        return 'PRIORITIZE_OTHER_BILLS_EXPENSES', 'MEDICATION'

    for rfd, sub_rfds in rfd_mapping.items():
        for sub_rfd in sub_rfds:
            if sub_rfd.lower() in remark_lower:
                return rfd, sub_rfd
    
    return None, None

st.title("Daily Prod Automation")
st.markdown("Removes unwanted statuses and remark sources for cleaner data processing")

# FILE UPLOAD
uploaded_file = st.file_uploader("📂 Drag and drop your file (Excel or CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # --- COLLECTOR SETTINGS (SHOWN AFTER UPLOAD) ---
    st.subheader("👤 Collector Settings")
    replacement_name = st.selectbox(
        "Choose replacement name for 'Mendoza, Joshua':",
        ["Select Name...", "Keep Original", "Bandola, Diana", "Reyes, Berlyn", "Galang, Jayson", "Antonio, Christine"]
    )

    if replacement_name == "Select Name...":
        st.info("💡 Please select a replacement name (or **'Keep Original'**) to proceed with the automation.")
        st.stop()

    # Set replacement_name to None if "Keep Original" is selected
    if replacement_name == "Keep Original":
        replacement_name = None

    st.divider()

    file_ext = uploaded_file.name.split('.')[-1]
    
    with st.spinner("🚀 Reading data..."):
        try:
            if file_ext == 'csv':
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            df.columns = df.columns.str.strip()
            
        except Exception as e:
            st.error(f"Error loading file: {e}")
            st.stop()

    if "Status" in df.columns and "Remark By" in df.columns:
        with st.spinner("🧼 Scrubbing data..."):
            # Truncate Remark to 250 characters (same as dailyprod.py)
            if 'Remark' in df.columns:
                df['Remark'] = df['Remark'].astype(str).str[:250]

            # Filtering Status: remove rows where Status starts with any excluded status
            mask_status = df['Status'].astype(str).str.startswith(tuple(EXCLUDED_STATUSES))

            # Format Date column
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%m/%d/%Y')

            # Filtering Remark By: remove rows matching specific remark sources
            mask_remark = df['Remark By'].astype(str).str.contains('SMGONZALES|CMENRIQUEZ', na=False)
            
            df_removed = df[mask_status | mask_remark]
            df_cleaned = df[~(mask_status | mask_remark)].copy()
            
            # Check if we lost too many rows
            if len(df_cleaned) < 100 and len(df) > 1000:
                st.warning(f"⚠️ Only {len(df_cleaned)} rows remaining after filtering out of {len(df)}. Check filter criteria.")

            # --- RFD / SUB-RFD Classification ---
            if "Remark" in df_cleaned.columns:
                classifications = df_cleaned['Remark'].apply(classify_remark)
                df_cleaned["RFD"], df_cleaned["SUB-RFD"] = zip(*classifications)
                df_cleaned["RFD"] = df_cleaned["RFD"].fillna("UNCATEGORIZED")
                df_cleaned["SUB-RFD"] = df_cleaned["SUB-RFD"].fillna("UNCATEGORIZED")
                
            # --- Account Number formatting (add 6 zeroes and force as text) ---
            if "Account No." in df_cleaned.columns:
                df_cleaned["Account No."] = df_cleaned["Account No."].astype(str).str.replace(r'\.0$', '', regex=True)
                df_cleaned["Account No."] = df_cleaned["Account No."].replace('nan', '')
                df_cleaned["Account No."] = df_cleaned["Account No."].replace('None', '')
                
                # Ensure we drop any trailing empty/nan rows from ruining the dataset before pasting
                df_cleaned = df_cleaned[df_cleaned["Account No."] != ""]
                
                # Prefix with single quote to force win32com to paste it as a text string with zeroes
                df_cleaned["Account No."] = "'000000" + df_cleaned["Account No."]
                
            # --- Time formatting (force as text with AM/PM to bypass Excel fractional floats) ---
            if "Time" in df_cleaned.columns:
                # Convert to standard formatted string and prefix with single quote
                def format_time(t):
                    try:
                        # Convert to datetime to standardize, then output as 12-hour AM/PM string
                        t_str = pd.to_datetime(str(t), errors='coerce').strftime('%I:%M %p')
                        if pd.isna(t_str) or t_str == 'NaT':
                            return str(t) # Fallback to original
                        return f"'{t_str}" # Prefix quote forces Excel to store exactly "hh:mm AM" as text
                    except Exception:
                        return str(t)
                df_cleaned["Time"] = df_cleaned["Time"].apply(format_time)
            
            # Prepare data identically to what will be pasted into the template (S.No to SUB-RFD)
            if 'S.No' in df_cleaned.columns and 'SUB-RFD' in df_cleaned.columns:
                start_idx = df_cleaned.columns.get_loc('S.No')
                end_idx = df_cleaned.columns.get_loc('SUB-RFD') + 1
                df_extracted_display = df_cleaned.iloc[:, start_idx:end_idx]
            else:
                df_extracted_display = df_cleaned.iloc[:, :53]

        # Top-level metrics
        c1, c2, c3 = st.columns(3)
        c1.metric("Original Rows", f"{len(df):,}")
        c2.metric("Kept Rows", f"{len(df_cleaned):,}")
        c3.metric("Removed Rows", f"{len(df_removed):,}", delta=f"-{len(df_removed)}", delta_color="inverse")

        st.divider()

        # Removed Rows Log
        st.subheader("🗑️ Removed Rows")
        st.markdown(f"**{len(df_removed)} rows** were filtered out based on their Status and Remarks.")
        if len(df_removed) > 0:
            st.dataframe(df_removed.head(100), use_container_width=True)
        
        st.divider()

        # Preview of what gets pasted
        st.subheader("📄 Template Paste Preview")
        st.markdown(f"A preview of the **{df_extracted_display.shape[1]} columns** that will be pasted into the template.")
        st.dataframe(df_extracted_display.head(50), use_container_width=True)
        
        st.divider()

        st.subheader("💾 Export Options")
        
        template_path = st.text_input("Template File Path:", value=r"C:\Users\SPM\Downloads\BPI\Template\ONE PROD REPORT TEMPLATEv1.xlsm")
        
        # Preserve the template's file extension since we copy it byte-for-byte
        tmpl_ext = os.path.splitext(template_path)[1] or ".xlsx"
        final_filename = f"Cleaned_Daily_Prod{tmpl_ext}"
        full_save_path = os.path.join(SAVE_PATH, final_filename)

        
        def process_template(df_data, tmpl_path, output_path, replacement_collector=None, password=None, is_bucket_mode=False, start_row=3, return_buckets=False):
                """Copy template and use Excel COM to write data — preserves everything perfectly."""
                if not WIN32COM_AVAILABLE:
                    st.error("pywin32 is required but not installed. Please run: pip install pywin32")
                    return False
                
                import win32com.client
                import pythoncom
                import time
                
                try:
                    # Check if source file exists
                    if not os.path.exists(tmpl_path):
                        st.error(f"Template file not found: {tmpl_path}")
                        return False
                    
                    # Check if destination directory exists
                    dest_dir = os.path.dirname(output_path)
                    if dest_dir and not os.path.exists(dest_dir):
                        os.makedirs(dest_dir, exist_ok=True)
                    
                    # Copy template with retry logic
                    max_retries = 3
                    retry_delay = 1
                    
                    for attempt in range(max_retries):
                        try:
                            shutil.copy2(tmpl_path, output_path)
                            break
                        except PermissionError as e:
                            if attempt < max_retries - 1:
                                st.warning(f"File in use, retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})")
                                time.sleep(retry_delay)
                            else:
                                st.error(f"Cannot copy template file after {max_retries} attempts. Please close Excel and try again.")
                                return False
                    
                    # Extract columns from S.No to SUB-RFD
                    if 'S.No' in df_data.columns and 'SUB-RFD' in df_data.columns:
                        start_idx = df_data.columns.get_loc('S.No')
                        end_idx = df_data.columns.get_loc('SUB-RFD') + 1
                        df_extracted_save = df_data.iloc[:, start_idx:end_idx]
                    else:
                        st.warning("Using fallback: first 53 columns")
                        df_extracted_save = df_data.iloc[:, :53]
                    
                    num_cols = df_extracted_save.shape[1]
                    num_rows = df_extracted_save.shape[0]
                    
                    st.write(f"Preparing to paste {num_rows} rows and {num_cols} columns into template")
                    
                    # Use Excel COM automation to write data
                    pythoncom.CoInitialize()
                    excel = None
                    wb = None
                    
                    try:
                        # DispatchEx strictly forces a brand new, isolated instance of Excel.
                        excel = win32com.client.DispatchEx("Excel.Application")
                        excel.Visible = False
                        excel.DisplayAlerts = False
                        
                        abs_path = os.path.abspath(output_path)
                        
                        # Open workbook with retry logic and password
                        for attempt in range(max_retries):
                            try:
                                wb = excel.Workbooks.Open(abs_path, Password=password)
                                break
                            except Exception as e:
                                if attempt < max_retries - 1:
                                    st.warning(f"Opening file, retrying... (Attempt {attempt + 1}/{max_retries})")
                                    time.sleep(retry_delay)
                                else:
                                    st.error(f"Cannot open workbook after {max_retries} attempts: {e}")
                                    return False
                        
                        ws = wb.Sheets("VOLARE EXTRACTION")
                        
                        # Convert data to a 2D list for instant bulk pasting
                        data_batch = []
                        import datetime as dt
                        
                        for r_idx, row in enumerate(df_data.itertuples(index=False), start=start_row):
                            row_data = []
                            # 1 to num_cols: S.No to SUB-RFD
                            for c_idx, value in enumerate(row, start=1):
                                if c_idx > num_cols:
                                    break
                                
                                if pd.isna(value):
                                    val = ""
                                elif isinstance(value, pd.Timestamp):
                                    val = value.to_pydatetime()
                                    if val.tzinfo is None:
                                        val = val.replace(tzinfo=dt.timezone.utc)
                                elif isinstance(value, datetime):
                                    val = value
                                    if val.tzinfo is None:
                                        val = val.replace(tzinfo=dt.timezone.utc)
                                else:
                                    val = value
                                    
                                if isinstance(val, tuple):
                                    val = str(val)
                                row_data.append(val)
                                
                            # 54 to 83: 30 Requested Formula Columns (BB to CE)
                            formulas = [
                                f'=XLOOKUP($BJ{r_idx},BPI!D:D,BPI!BR:BR,0)',                            # BB (54): TAGGING
                                f'=$J{r_idx}',                                                          # BC (55): STATUS
                                f'=VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[placement]],4,0)', # BD (56): BUCKET
                                f'=VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[endorsement date]],6,0)', # BE (57): ENDORSEMENT DATE
                                f'=VLOOKUP($BD{r_idx},NOTES!H:J,3,0)',                                   # BF (58): BUCKET (2)
                                "",                                                                     # BG (59): REF CODE
                                "MADRID",                                                                # BH (60): AGENCY NAME
                                "",                                                                     # BI (61): CYCLE
                                f'=$E{r_idx}',                                                          # BJ (62): LOAN ACCOUNT NUMBER
                                f'=IF(LEFT($BC{r_idx},12)="LS VIA EMAIL", VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[email]],24,0), VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[email]],22,0))', # BK (63): Mobile/Email
                                f'=VLOOKUP($J{r_idx},NOTES!AA:AB,2,0)',                                  # BL (64): TYPE OF COLLECTION EFFORT
                                "",                                                                     # BM (65): TYPE OF SMS / EMAIL
                                f'=IFERROR(VLOOKUP($L{r_idx},NOTES!V:W,2,0),VLOOKUP(\'VOLARE EXTRACTION\'!$BB{r_idx},NOTES!U:W,3,0))', # BN (66): COLLECTOR
                                f'=VLOOKUP($J{r_idx},NOTES!K:L,2,0)',                                    # BO (67): ACTION
                                f'=VLOOKUP($BO{r_idx},NOTES!B:E,4,0)',                                   # BP (68): REACTION CODE
                                f'=$K{r_idx}',                                                          # BQ (69): REMARKS
                                f'=$W{r_idx}',                                                          # BR (70): DATE OF PTP/TFIP
                                f'=$V{r_idx}',                                                          # BS (71): PTP AMOUNT
                                "",                                                                     # BT (72): PAST DUE AMOUNT
                                f'=VLOOKUP($BJ{r_idx},BPI!D:P,12,0)',                                    # BU (73): PAY OFF AMOUNT (OB)
                                f'=VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[prin]],13,0)',     # BV (74): PRINCIPAL AMOUNT
                                f'=VLOOKUP($BJ{r_idx},E{r_idx}:BA{r_idx},48,0)',                         # BW (75): REASON FOR DELIQUENCY
                                f'=VLOOKUP($BJ{r_idx},E{r_idx}:BA{r_idx},49,0)',                         # BX (76): SUB-CATEGORIES
                                "",                                                                     # BY (77): Due Date
                                f'=IF(COUNTIF(NOTES!Q:Q,\'VOLARE EXTRACTION\'!$J{r_idx}) > 0, "POSITIVE", "NEGATIVE")', # BZ (78): SKIPTRACE PLATFORM
                                "",                                                                     # CA (79): FIELD VISIT REMARKS
                                "",                                                                     # CB (80): GATHERED CONTACT NO.
                                "",                                                                     # CC (81): SOURCE/ RELATIONSHIP
                                f'=$B{r_idx}',                                                          # CD (82): DATE OF ACTION
                                f'=RIGHT($C{r_idx}, 2)'                                                 # CE (83): AM/PM1/PM2
                            ]
                            
                            row_data.extend(formulas)
                            
                            # --- SPECIAL OVERRIDES FOR BUCKET MODE ---
                            if is_bucket_mode:
                                # 1. Copy BE (index 56) to BZ (index 77)
                                row_data[77] = row_data[56]
                                # 2. DATE OF ACTION (CD, index 81) to AA (index 26)
                                row_data[26] = row_data[81]
                                # 3. AM/PM (CE, index 82) to AE (index 30)
                                row_data[30] = row_data[82]
                            
                            if return_buckets:
                                # Temporary track original row index in a far column (e.g. column 702 = ZZ)
                                while len(row_data) < 702:
                                    row_data.append("")
                                row_data[701] = r_idx - start_row # The relative index 0-based
                            
                            data_batch.append(row_data)
                        
                        # Paste everything instantly in one operation
                        if data_batch:
                            start_cell_obj = ws.Cells(start_row, 1)
                            last_row_pasted = start_row + len(data_batch) - 1
                            end_cell_obj = ws.Cells(last_row_pasted, len(data_batch[0]))
                            
                            st.write(f"Pasting {len(data_batch)} rows from A{start_row} to column {len(data_batch[0])}")
                            ws.Range(start_cell_obj, end_cell_obj).Formula = data_batch
                            
                            # Compute formulas and bake them immediately into permanent flat values
                            try:
                                full_range = ws.Range(start_cell_obj, end_cell_obj)
                                full_range.Copy()
                                full_range.PasteSpecial(Paste=-4163)  # xlPasteValues
                                excel.Application.CutCopyMode = False
                            except Exception as e:
                                st.warning(f"Formula baking warning: {e}")
                            
                            # Apply formatting: Copy the red background format from BT1 to all pasted rows
                            try:
                                ws.Range("BT1").Copy()
                                ws.Range(f"BT{start_row}:BT{last_row_pasted}").PasteSpecial(Paste=-4122)  # xlPasteFormats
                                ws.Range(f"CA{start_row}:CC{last_row_pasted}").PasteSpecial(Paste=-4122)
                                excel.Application.CutCopyMode = False
                            except Exception as e:
                                st.warning(f"Formatting warning: {e}")
                                
                            # =========================================================
                            # POST-PROCESSING STEPS
                            # =========================================================

                            # Grab the template formula from I1 for use on POSITIVE SKIPTRACE rows
                            try:
                                source_i1 = ws.Range("I1")
                                formula_i1 = source_i1.Formula
                                formula_i1_r1c1 = source_i1.FormulaR1C1
                            except Exception:
                                source_i1 = None
                                formula_i1 = ""
                                formula_i1_r1c1 = ""

                            # Recalculate last_row before starting post-processing steps (Fix for Step 1 error)
                            last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                            if last_row < 3: 
                                last_row = 3

                            # Step 1: Delete rows where ACTION (BO) is EXCLUDED or #N/A
                            try:
                                vals = ws.Range(f"BO3:BO{last_row}").Value
                                if vals:
                                    to_delete = []
                                    for i, row_val in enumerate(vals):
                                        val = row_val[0]
                                        val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                                        if val == -2146826246 or val_str in ("EXLUDED", "EXCLUDED", "#N/A"):
                                            to_delete.append(3 + i)
                                    if to_delete:
                                        st.write(f"Deleting {len(to_delete)} rows with excluded ACTION")
                                        to_delete.reverse()
                                        for r in to_delete:
                                            ws.Rows(r).Delete()
                            except Exception as e:
                                st.warning(f"Step 1 error: {e}")

                            # Recalculate last_row after Step 1 deletion shift
                            last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                            if last_row < 3: 
                                last_row = 3

                            # Step 2: Clear PTP DATE (BR) and PTP AMOUNT (BS) where ACTION (BO) != "PTP"
                            try:
                                vals = ws.Range(f"BO3:BO{last_row}").Value
                                if vals:
                                    cleared_count = 0
                                    for i, row_val in enumerate(vals):
                                        val = row_val[0]
                                        val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                                        if val_str != "PTP" and val_str != "":
                                            ws.Range(f"BR{3+i}:BS{3+i}").ClearContents()
                                            cleared_count += 1
                                    if cleared_count > 0:
                                        st.write(f"Cleared PTP data for {cleared_count} rows")
                            except Exception as e:
                                st.warning(f"Step 2 error: {e}")

                            # Step 3: Clear REACTION CODE (BP) where it is "0" or "#N/A"
                            try:
                                vals = ws.Range(f"BP3:BP{last_row}").Value
                                if vals:
                                    cleared_count = 0
                                    for i, row_val in enumerate(vals):
                                        val = row_val[0]
                                        val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                                        if val == -2146826246 or val_str in ("0", "0.0", "#N/A"):
                                            ws.Range(f"BP{3+i}").ClearContents()
                                            cleared_count += 1
                                    if cleared_count > 0:
                                        st.write(f"Cleared {cleared_count} reaction codes")
                            except Exception as e:
                                st.warning(f"Step 3 error: {e}")

                            # Step 4: Delete rows where PTP AMOUNT (BS) is "0"
                            try:
                                vals = ws.Range(f"BS3:BS{last_row}").Value
                                if vals:
                                    to_delete = []
                                    for i, row_val in enumerate(vals):
                                        val = row_val[0]
                                        val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                                        if val_str in ("0", "0.0"):
                                            to_delete.append(3 + i)
                                    if to_delete:
                                        st.write(f"Deleting {len(to_delete)} rows with zero PTP amount")
                                        to_delete.reverse()
                                        for r in to_delete:
                                            ws.Rows(r).Delete()
                            except Exception as e:
                                st.warning(f"Step 4 error: {e}")

                            # Recalculate last_row after Step 4 deletion shift
                            last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                            if last_row < 3: 
                                last_row = 3

                            # Step 5: Clear SKIPTRACE column (BZ) where it is "NEGATIVE" (instead of deleting rows)
                            try:
                                vals = ws.Range(f"BZ3:BZ{last_row}").Value
                                if vals:
                                    cleared_count = 0
                                    for i, row_val in enumerate(vals):
                                        val = row_val[0]
                                        val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                                        if val_str == "NEGATIVE":
                                            ws.Range(f"BZ{3+i}").ClearContents()
                                            cleared_count += 1
                                    if cleared_count > 0:
                                        st.write(f"Cleared SKIPTRACE for {cleared_count} rows with NEGATIVE value")
                            except Exception as e:
                                st.warning(f"Step 5 error: {e}")

                            # Step 6: For rows where BZ (SKIPTRACE PLATFORM) contains a platform name
                            # (VIBER, FACEBOOK, etc.), apply the following to column I in order:
                            #   1. Write the row-specific formula into each POSITIVE column I cell
                            #   2. Bake the formula result to a plain value (xlPasteValues) per cell so
                            #      the contact name is stored as readable text for later processing steps
                            #   3. Apply yellow background directly via Interior.Color per cell — this
                            #      avoids clipboard conflicts that cause PasteSpecial(formats) to silently
                            #      fail on non-contiguous union ranges.
                            # After the bulk bake earlier, BZ holds the actual platform name, not "POSITIVE",
                            # so POSITIVE rows are identified as any BZ that is non-empty and not "NEGATIVE".
                            try:
                                positive_rows = []
                                bz_vals = ws.Range(f"BZ3:BZ{last_row}").Value
                                if bz_vals:
                                    for i, row_val in enumerate(bz_vals):
                                        val = row_val[0]
                                        val_str = str(val).strip().upper() if val not in (None, "") else ""
                                        if val_str and val_str not in ("NEGATIVE", "0", "FALSE", "#N/A", "NONE"):
                                            positive_rows.append(3 + i)

                                st.write(f"Found {len(positive_rows)} POSITIVE SKIPTRACE rows (BZ contains platform name e.g. VIBER/FACEBOOK)")

                                if positive_rows:
                                    for row_num in positive_rows:
                                        # (Step 8): Copy REMARK (K=11) to BW (75)
                                        try:
                                            remark_val = ws.Cells(row_num, 11).Value
                                            ws.Cells(row_num, 75).Value = remark_val
                                        except Exception:
                                            pass

                                        target = ws.Cells(row_num, 78)  # Column BZ

                                        # 1. Write the row-specific formula
                                        if formula_i1_r1c1 and formula_i1_r1c1 not in ("", "0"):
                                            try:
                                                target.FormulaR1C1 = formula_i1_r1c1
                                            except Exception:
                                                target.Formula = f'=TRIM(RIGHT(SUBSTITUTE($J{row_num}," ",REPT(" ",100)),100))'
                                        elif formula_i1 and formula_i1 not in ("", "0"):
                                            adjusted = formula_i1.replace("$J1", f"$J{row_num}").replace("J1", f"J{row_num}")
                                            try:
                                                target.Formula = adjusted
                                            except Exception:
                                                target.Formula = f'=TRIM(RIGHT(SUBSTITUTE($J{row_num}," ",REPT(" ",100)),100))'
                                        else:
                                            target.Formula = f'=TRIM(RIGHT(SUBSTITUTE($J{row_num}," ",REPT(" ",100)),100))'

                                        # 2. Bake the formula to a plain value immediately
                                        target.Copy()
                                        target.PasteSpecial(Paste=-4163)  # xlPasteValues
                                        excel.Application.CutCopyMode = False

                                        # 3. Apply yellow background directly — reliable on individual cells
                                        target.Interior.Color = 65535  # Yellow

                                    st.write(f"Applied formula (baked to value) + yellow fill to column I for {len(positive_rows)} POSITIVE SKIPTRACE rows")
                            except Exception as e:
                                st.warning(f"Step 6 error: {e}")

                            # Step 7: Clear BW/BX values set to 0 or UNCATEGORIZED
                            try:
                                cleared_count = 0
                                for col in ("BW", "BX"):
                                    vals = ws.Range(f"{col}3:{col}{last_row}").Value
                                    if vals:
                                        for i, row_val in enumerate(vals):
                                            val = row_val[0]
                                            val_str = str(val).strip().upper() if val is not None else ""
                                            if val_str in ("0", "0.0", "UNCATEGORIZED"):
                                                ws.Range(f"{col}{3+i}").ClearContents()
                                                cleared_count += 1
                                if cleared_count > 0:
                                    st.write(f"Cleared {cleared_count} BW/BX cells with 0 or UNCATEGORIZED")
                            except Exception as e:
                                st.warning(f"Step 7 error: {e}")

                            # Step 9: Delete whole row of #N/A at the COLLECTOR at BN column (66)
                            try:
                                # Recalculate last_row after any potential deletions above
                                last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                                if last_row >= 3:
                                    vals = ws.Range(f"BN3:BN{last_row}").Value
                                    if vals:
                                        to_delete = []
                                        for i, row_val in enumerate(vals):
                                            val = row_val[0]
                                            # Handle Excel #N/A error code or string "#N/A"
                                            if val == -2146826246 or str(val).strip().upper() == "#N/A":
                                                to_delete.append(3 + i)
                                        if to_delete:
                                            st.write(f"Deleting {len(to_delete)} rows with #N/A in COLLECTOR (BN)")
                                            to_delete.reverse()
                                            for r in to_delete:
                                                ws.Rows(r).Delete()
                            except Exception as e:
                                st.warning(f"Step 9 error: {e}")

                            # Step 10: Replace "Mendoza, Joshua" in COLLECTOR column (BN=66)
                            if replacement_collector:
                                try:
                                    last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                                    if last_row >= 3:
                                        vals = ws.Range(f"BN3:BN{last_row}").Value
                                        if vals:
                                            update_count = 0
                                            for i, row_val in enumerate(vals):
                                                val = str(row_val[0]).strip()
                                                if "Mendoza, Joshua" in val:
                                                    ws.Cells(3 + i, 66).Value = replacement_collector
                                                    update_count += 1
                                            if update_count > 0:
                                                st.write(f"Replaced 'Mendoza, Joshua' with '{replacement_collector}' in {update_count} rows")
                                except Exception as e:
                                    st.warning(f"Step 10 error: {e}")

                            # Step 11: Clear cells in TYPE OF COLLECTION EFFORT (BL=64) if #N/A
                            try:
                                last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                                if last_row >= 3:
                                    vals = ws.Range(f"BL3:BL{last_row}").Value
                                    if vals:
                                        cleared_count = 0
                                        for i, row_val in enumerate(vals):
                                            val = row_val[0]
                                            # Handle Excel #N/A error code or string "#N/A"
                                            if val == -2146826246 or str(val).strip().upper() == "#N/A":
                                                ws.Cells(3 + i, 64).ClearContents()
                                                cleared_count += 1
                                        if cleared_count > 0:
                                            st.write(f"Cleared {cleared_count} #N/A cells in TYPE OF COLLECTION EFFORT (BL)")
                            except Exception as e:
                                st.warning(f"Step 11 error: {e}")

                            # Step 12: Delete rows where BUCKET (BJ=62) is #N/A
                            try:
                                # Recalculate last_row after any potential deletions above
                                last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                                if last_row >= 3:
                                    vals = ws.Range(f"BF3:BF{last_row}").Value
                                    if vals:
                                        to_delete = []
                                        for i, row_val in enumerate(vals):
                                            val = row_val[0]
                                            # Handle Excel #N/A error code or string "#N/A"
                                            if val == -2146826246 or str(val).strip().upper() == "#N/A":
                                                to_delete.append(3 + i)
                                        if to_delete:
                                            st.write(f"Deleting {len(to_delete)} rows with #N/A in BUCKET (BF)")
                                            to_delete.reverse()
                                            for r in to_delete:
                                                ws.Rows(r).Delete()
                            except Exception as e:
                                st.warning(f"Step 12 error: {e}")
                            
                            # Calculation readback (Pass 1)
                            calc_dict = {}
                            if return_buckets:
                                try:
                                    last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                                    if last_row >= start_row:
                                        # Read calculated bucket (BF=58) and original index (ZZ=702)
                                        # Use a Union or read the chunk
                                        all_vals = ws.Range(f"A{start_row}:ZZ{last_row}").Value
                                        if all_vals:
                                            for row_vals in all_vals:
                                                bucket_val = str(row_vals[57]).strip() if row_vals[57] else ""
                                                orig_idx_val = row_vals[701]
                                                if orig_idx_val is not None:
                                                    calc_dict[int(orig_idx_val)] = bucket_val
                                except Exception as e:
                                    st.warning(f"Bucket readback error: {e}")

                            # Unselect columns by selecting A1 before saving
                            try:
                                ws.Range("A1").Select()
                            except Exception:
                                pass
                        
                        # Save and close
                        wb.Save()
                        wb.Close(False)
                        
                        if return_buckets:
                            return True, calc_dict
                        
                        st.success(f"✅ Successfully processed {num_rows} rows into template")
                        # Store in session state for reuse by the splitting tool
                        st.session_state['last_master_path'] = output_path
                        return True
                        
                    except Exception as e:
                        st.error(f"Error in Excel COM processing: {e}")
                        import traceback
                        st.code(traceback.format_exc())
                        return False
                        
                    finally:
                        # Clean up Excel instance
                        if wb:
                            try:
                                wb.Close(False)
                            except:
                                pass
                        if excel:
                            try:
                                excel.Quit()
                            except:
                                pass
                        pythoncom.CoUninitialize()
                        
                except Exception as e:
                    st.error(f"Error in process_template: {e}")
                    import traceback
                    st.code(traceback.format_exc())
                    return False

        
        def process_campaign_split(master_path, tmpl_path, save_dir, date_str, bucket_list, password):
            """Reads the Master file once and splits it into Campaign files with specific formatting."""
            if not WIN32COM_AVAILABLE:
                st.error("pywin32 is required.")
                return False
                
            import win32com.client
            import pythoncom
            import time
            
            try:
                pythoncom.CoInitialize()
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.EnableEvents = False  # Suppress macros that might trigger prompts
                excel.Interactive = False   # Blocks all user-interactive dialogs
                
                # 1. Open the Master file and read data
                abs_master = os.path.abspath(master_path)
                st.write(f"Opening master: {os.path.basename(abs_master)}...")
                
                # Use more explicit parameters to suppress prompts (WriteResPassword, UpdateLinks)
                wb_master = excel.Workbooks.Open(
                    Filename=abs_master, 
                    UpdateLinks=0, 
                    ReadOnly=True,
                    Password=password, 
                    WriteResPassword=password,
                    IgnoreReadOnlyRecommended=True
                )
                
                try:
                    ws_master = wb_master.Sheets("VOLARE EXTRACTION")
                except Exception:
                    st.warning("Sheet 'VOLARE EXTRACTION' not found in Master. Using first sheet.")
                    ws_master = wb_master.Sheets(1)
                
                last_row = ws_master.Cells(ws_master.Rows.Count, "A").End(-4162).Row
                if last_row < 3:
                    st.warning("Master file is empty or has no data rows.")
                    wb_master.Close(False)
                    excel.Quit()
                    return False
                
                # Read all data from A3 to CE{last_row} (Columns 1 to 83)
                data_range = ws_master.Range(ws_master.Cells(3, 1), ws_master.Cells(last_row, 83)).Value
                wb_master.Close(False)
                
                if not data_range:
                    st.warning("Could not read data from Master file.")
                    excel.Quit()
                    return False

                # 2. Filter data into bucket groups (Column BF is index 57)
                groups = {f"{b} Days": [] for b in bucket_list}
                for row_vals in data_range:
                    if not row_vals: continue
                    bucket_val = str(row_vals[57]).strip() if row_vals[57] else ""
                    if bucket_val in groups:
                        groups[bucket_val].append(list(row_vals))
                
                for b in bucket_list:
                    bucket_name_str = f"{b} Days"
                    rows = groups[bucket_name_str]
                    if not rows:
                        st.info(f"Skipping {bucket_name_str}: No data found.")
                        continue
                    
                    campaign_tmpl = fr"C:\Users\SPM\Downloads\BPI\Template\PL_{{Year}}{{Month}}{{Day}}_{{Campaign}}_Madrid.xlsx".replace("{Year}", "2026")
                    campaign_tmpl = fr"C:\Users\SPM\Downloads\BPI\Template\PL_2026{{Month}}{{Day}}_{{Campaign}}_Madrid.xlsx"
                    
                    if not os.path.exists(campaign_tmpl):
                        st.error(f"❌ Campaign template not found: {os.path.basename(campaign_tmpl)}")
                        continue

                    bucket_filename = f"PL_{date_str}_{b}_Madrid.xlsx"
                    save_path = os.path.abspath(os.path.join(save_dir, bucket_filename))
                    local_xls_path = save_path.replace(".xlsx", ".xls")
                    
                    try:
                        st.write(f"Processing bucket {b}...")
                        shutil.copy2(campaign_tmpl, local_xls_path)
                        
                        # Re-ensure prompt suppression right before opening the bucket report
                        excel.DisplayAlerts = False
                        excel.Interactive = False
                        
                        # Open the local copy with full suppression parameters
                        wb_bucket = excel.Workbooks.Open(
                            Filename=local_xls_path, 
                            UpdateLinks=0, 
                            ReadOnly=False,
                            Password=password, 
                            WriteResPassword=password,
                            IgnoreReadOnlyRecommended=True
                        )
                        
                        # Try to get the sheet, fallback to first sheet if named differently
                        try:
                            ws_bucket = wb_bucket.Sheets("VOLARE EXTRACTION")
                        except Exception:
                            available_sheets = [s.Name for s in wb_bucket.Sheets]
                            st.warning(f"Sheet 'VOLARE EXTRACTION' not found in bucket template.")
                            st.info(f"Available sheets: {', '.join(available_sheets)}")
                            st.info(f"Falling back to first sheet: '{available_sheets[0]}'")
                            ws_bucket = wb_bucket.Sheets(1)
                        
                        # Construct Bucket Rows with correct mapping
                        # Master Row Indices: BE=56, BZ=77, CD=81, CE=82
                        # Mapping to Bucket Indices: A=0 to V=21, AA=26, AE=30
                        bucket_rows = []
                        for r_master in rows:
                            # Width 31 (Target column AE)
                            new_row = [""] * 31
                            # BE:BZ (56:index 78) -> A:V (0:index 22)
                            new_row[0:22] = r_master[56:78]
                            # CD (81) -> AA (26)
                            new_row[26] = r_master[81]
                            # CE (82) -> AE (30)
                            new_row[30] = r_master[82]
                            bucket_rows.append(tuple(new_row))
                        
                        rows_to_paste = tuple(bucket_rows)
                        
                        # Paste at the user-defined start row (Row 4)
                        start_row_bucket = 4
                        start_cell = ws_bucket.Cells(start_row_bucket, 1)
                        last_row_bucket = start_row_bucket + len(rows_to_paste) - 1
                        end_cell = ws_bucket.Cells(last_row_bucket, 31)
                        
                        st.write(f"Pasting {len(rows_to_paste)} rows into bucket sheet at Row {start_row_bucket}...")
                        ws_bucket.Range(start_cell, end_cell).Value = rows_to_paste
                        
                        # --- FORMATTING ---
                        try:
                            # 1. Apply Continuous Border to ALL pasted cells (A4:AE...)
                            pasted_range = ws_bucket.Range(start_cell, end_cell)
                            pasted_range.Borders.LineStyle = 1 # xlContinuous
                            pasted_range.Borders.Weight = 2    # xlThin

                            st.write("Applying skiptrace formatting (Yellow, Center)...")
                            for i in range(len(rows_to_paste)):
                                current_row = start_row_bucket + i
                                cell = ws_bucket.Cells(current_row, 22) # Column V
                                val = str(cell.Value).strip().upper()
                                
                                if val and val not in ("NEGATIVE", "0", "FALSE", "#N/A", "NONE", "NAN", ""):
                                    # 2. Yellow Background
                                    cell.Interior.Color = 65535 # Yellow
                                    
                                    # 3. Horizontal Center
                                    cell.HorizontalAlignment = -4108 # xlCenter
                            
                            # Auto-fit columns for professional look
                            ws_bucket.Columns("A:AE").AutoFit()
                                    
                        except Exception as fmt_err:
                            st.warning(f"Formatting warning: {fmt_err}")
                        
                        # Save specifically in modern OpenXML (.xlsx) format (FileFormat=51)
                        # Re-encrypt with the same password as requested
                        wb_bucket.SaveAs(save_path, FileFormat=51, Password=password)
                        wb_bucket.Close()
                        st.success(f"✅ Generated: {bucket_filename}")
                        
                    except Exception as bucket_err:
                        st.error(f"❌ Error processing bucket {b}: {bucket_err}")
                        import traceback
                        st.code(traceback.format_exc())
                    finally:
                        if os.path.exists(local_xls_path):
                            try:
                                os.remove(local_xls_path)
                            except:
                                pass
                
                # Restore Interactive mode before quitting
                excel.Interactive = True
                excel.Quit()
                pythoncom.CoUninitialize()
                return True
                
            except Exception as e:
                st.error(f"Error in campaign splitting: {e}")
                import traceback
                st.code(traceback.format_exc())
                return False


        st.divider()
        st.subheader("📤 Export Settings")
        output_choice = st.radio(
            "Select Output Type:",
            ["Master Report Only", "Triple Bucket Reports Only", "Both"],
            index=2, # Default to Both
            horizontal=True
        )
        
        st.divider()

        if output_choice in ["Master Report Only", "Both"]:
            st.subheader("📋 Master Report")
            col_export_1, col_export_2 = st.columns(2)
            
            with col_export_1:
                if st.button("💾 SAVE TO DRR AUTO FOLDER", width='stretch'):
                    if not os.path.exists(template_path):
                        st.error("Template file not found! Please check the path.")
                    elif not WIN32COM_AVAILABLE:
                        st.error("pywin32 is not installed. Cannot save to Excel template.")
                    else:
                        try:
                            with st.spinner(f"Writing {len(df_cleaned)} rows to template..."):
                                success = process_template(df_cleaned, template_path, full_save_path, replacement_name, password="MAD_2Q2026")
                                if success:
                                    st.success(f"Successfully saved to:\n{full_save_path}")
                        except Exception as e:
                            st.error(f"Could not save to local path: {e}")

            with col_export_2:
                if os.path.exists(template_path) and WIN32COM_AVAILABLE:
                    try:
                        # Save to a temp file first, then read back for download
                        if st.button("📥 DOWNLOAD MASTER REPORT", width='stretch'):
                            with st.spinner("Preparing download..."):
                                tmp_file = os.path.join(tempfile.gettempdir(), final_filename)
                                success = process_template(df_cleaned, template_path, tmp_file, replacement_name, password="MAD_2Q2026")
                                if success:
                                    with open(tmp_file, "rb") as f:
                                        file_bytes = f.read()
                                    # Clean up temp file
                                    try:
                                        os.remove(tmp_file)
                                    except Exception:
                                        pass
                                    
                                    mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if final_filename.endswith('.xlsm') else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    
                                    st.download_button(
                                        label="Confirm Download",
                                        data=file_bytes,
                                        file_name=final_filename,
                                        mime=mime_type,
                                        width='stretch'
                                    )
                    except Exception as e:
                        st.error(f"Error preparing download: {e}")
                        import traceback
                        st.code(traceback.format_exc())
                elif not WIN32COM_AVAILABLE:
                    st.warning("pywin32 is required for Excel export. Install with: pip install pywin32")
                else:
                    st.error("Template not found!")

        if output_choice in ["Triple Bucket Reports Only", "Both"]:
            if output_choice == "Both":
                st.divider()
            st.subheader("🚀 Triple Bucket Export")
            st.markdown("Generates 3 separate files for **120, 150, and 180 Days** buckets based on specialized templates.")
            
            if st.button("🔥 GENERATE 120/150/180 BUCKET REPORTS", width='stretch'):
                report_date_obj, date_str = get_report_date()
                buckets_to_process = [120, 150, 180]
                
                # Smart Source Selection
                master_source = None
                if 'last_master_path' in st.session_state and os.path.exists(st.session_state['last_master_path']):
                    master_source = st.session_state['last_master_path']
                    st.info(f"Reusing already processed Master file: {os.path.basename(master_source)}")
                else:
                    with st.spinner("No active Master found. Running background calculation pass... (Saved in DRR Auto folder)"):
                        tmpl_ext = os.path.splitext(template_path)[1] or ".xlsm"
                        # Generate a descriptive internal master name in the local folder
                        internal_master_name = f"Master_Calculation_{date_str}{tmpl_ext}"
                        tmp_master = os.path.join(SAVE_PATH, internal_master_name)
                        
                        success = process_template(df_cleaned, template_path, tmp_master, replacement_name, password="MAD_2Q2026")
                        if success:
                            master_source = tmp_master
                        else:
                            st.error("Background calculation failed.")
                
                if master_source:
                    with st.spinner("Splitting Master into campaign buckets..."):
                        process_campaign_split(master_source, template_path, SAVE_PATH, date_str, buckets_to_process, password="MAD_2Q2026")
            else:
                st.warning("Template file is required for downloading.")
    else:
        st.error("⚠️ Columns 'Status' and 'Remark By' not found.")
        st.write(f"Available columns: {list(df.columns)}")

else:
    st.info("Upload your Daily Prod Excel File to begin.")