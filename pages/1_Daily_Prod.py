import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
import shutil
import pythoncom
from datetime import datetime, timedelta

# Set page configuration
st.set_page_config(page_title="Daily Prod Automation", page_icon=None, layout="wide")
set_premium_style()

# Check for win32com availability
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    st.warning(" pywin32 is not installed. Excel COM automation will not work.")

# --- HELPERS ---
def get_report_date():
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        report_date = today - timedelta(days=3) 
    else:
        report_date = today - timedelta(days=1)
    return report_date, report_date.strftime("%Y%m%d")

EXCLUDED_STATUSES = [
    'CONFIRMED', 'PTP_FF', 'SMS SENT - T6 NO RESPONSE',
    'LS VIA SOCMED - T6 NO RESPONSE', 'SMS RECEIVED - UNDER NEGO',
    'SMS RECEIVED - WILL SETTLE', 'NEW', 'BP', 'ENCO', 'UNLOCKED',
    'BULK SMS', 'REACTIVE', 'LOCKED', 'PM',
    'SMS RECEIVED - NO SUCH PERSON (NSP)',
    'SERVICE REQUEST (SR) - FOR COLLECTION'
]

rfd_mapping = {
    "AWAITING_FUNDS": ["ALLOTMENT", "REMITTANCE", "ALLOWANCE", "LOANS", "PROFIT", "SALARY", "INCENTIVES", "BACKPAY", "BENEFITS", "PENSION", "COLLECTIONS", "COMMISSION"],
    "EMPLOYMENT_STATUS": ['AWAITING TO ONBOARD', 'CONTRACTUAL', 'NEWLY-HIRED', 'CONFLICT WITH EMPLOYER', 'UNEMPLOYMENT', 'FREELANCE', 'SUSPENSION', 'ON MATERNITY LEAVE'],
    "EMERGENCY": ['HOSPITALIZATION', 'VICTIM OF CALAMITY', 'ACCIDENT', 'DEATH OF RELATIVE'],
    "NO_CAPACITY_TO_PAY": ["SHORT OF FUNDS", "FINANCIAL DIFFICULTY", "BUSINESS SLOWDOWN"],
    "PRIORITIZE_OTHER_BILLS_EXPENSES": ["LOANS/OTHER DUES", "TUITION", "UTILITY BILLS", "MEDICATION"],
    "LOAN_CLARIFICATIONS": ["PAST DUE BALANCE", "MONTHLY AMORTIZATION", "MATURITY DATE", "REMAINING TERM"],
    "PAYMENT_CLARIFICATIONS": ["FUNDED ADA", "CLAIMING SETTLED ACCOUNT", "LATE CHARGES", "PAYMENT HISTORY"],
    "PERSONAL_REASON": ["SCAM VICTIM", "OUT OF COUNTRY/TOWN", "BUSY SCHEDULE", "CHILDBIRTH", "FAMILY MATTER", "ROBBERY VICTIM"],
    "OTHERS": ["DROPPED CALL", "ONLINE BANKING CONCERN", "OVERLOOKED DUEDATE", "REFUSE TO DISCLOSED REASON"]
}

def classify_remark(remark):
    if pd.isna(remark) or str(remark).strip() == "": return None, None
    remark_lower = str(remark).lower()
    if any(k in remark_lower for k in ["got scammed", "scammed", "scammer"]): return "PERSONAL_REASON", "SCAM VICTIM"
    if any(k in remark_lower for k in ['ootc', 'out of the country', 'abroad']): return 'PERSONAL_REASON', 'OUT OF COUNTRY/TOWN'
    if any(k in remark_lower for k in ['unemployed']): return 'EMPLOYMENT_STATUS', 'UNEMPLOYMENT'
    if any(k in remark_lower for k in ['dropped', 'dropped call', 'call drop']): return 'OTHERS', 'DROPPED CALL'
    if any(k in remark_lower for k in ['overlooked']): return 'OTHERS', 'OVERLOOKED DUEDATE'
    if any(k in remark_lower for k in ['financial struggle', 'financial problem']): return 'NO_CAPACITY_TO_PAY', 'FINANCIAL DIFFICULTY'
    if any(k in remark_lower for k in ['bankcruptcy']): return 'NO_CAPACITY_TO_PAY', 'BUSINESS SLOWDOWN'
    for rfd, sub_rfds in rfd_mapping.items():
        for sub_rfd in sub_rfds:
            if sub_rfd.lower() in remark_lower: return rfd, sub_rfd
    return None, None

# --- CORE AUTOMATION ---
def process_template(df_data, tmpl_path, output_path, replacement_collector=None, password=None, start_row=3):
    """Surgical cleanup matching manual entries exactly."""
    if not WIN32COM_AVAILABLE: return False
    try:
        shutil.copy2(tmpl_path, output_path)
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(os.path.abspath(output_path), Password=password)
        ws = wb.Sheets("VOLARE EXTRACTION")
        
        # --- PHASE 1: INITIAL PASTE & BAKE ---
        # Bulletproof Date Cleaning
        df_safe = df_data.copy()
        for col in df_safe.columns:
            if pd.api.types.is_datetime64_any_dtype(df_safe[col]):
                try:
                    # Strip any existing timezones
                    df_safe[col] = pd.to_datetime(df_safe[col]).dt.tz_localize(None)
                except:
                    # If already naive, just ensure it's a datetime
                    df_safe[col] = pd.to_datetime(df_safe[col], errors='coerce')
        
        # Replace NaT/NaN with None for Excel
        df_safe = df_safe.where(pd.notnull(df_safe), None)

        data_batch = []
        for r_idx, row in enumerate(df_safe.itertuples(index=False), start=start_row):
            row_vals = list(row[:53]) # S.No to SUB-RFD
            formulas = [
                f'=XLOOKUP($BJ{r_idx},BPI!D:D,BPI!BR:BR,0)', # BB (54)
                f'=$J{r_idx}', # BC (55)
                f'=VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[placement]],4,0)', # BD (56)
                f'=VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[endorsement date]],6,0)', # BE (57)
                f'=VLOOKUP($BD{r_idx},NOTES!H:J,3,0)', # BF (58)
                "", "MADRID", "", # BG, BH, BI
                f'=$E{r_idx}', # BJ (62)
                f'=IF(LEFT($BC{r_idx},12)="LS VIA EMAIL", VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[email]],24,0), VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[email]],22,0))', # BK
                f'=VLOOKUP($J{r_idx},NOTES!AA:AB,2,0)', # BL (64)
                "", # BM
                f'=IFERROR(VLOOKUP($L{r_idx},NOTES!V:W,2,0),VLOOKUP(\'VOLARE EXTRACTION\'!$BB{r_idx},NOTES!U:W,3,0))', # BN (66)
                f'=VLOOKUP($J{r_idx},NOTES!K:L,2,0)', # BO (67)
                f'=VLOOKUP($BO{r_idx},NOTES!B:E,4,0)', # BP (68)
                f'=$K{r_idx}', f'=$W{r_idx}', f'=$V{r_idx}', # BQ, BR, BS
                "", # BT (72)
                f'=VLOOKUP($BJ{r_idx},BPI!D:P,12,0)', # BU
                f'=VLOOKUP($BJ{r_idx},db_bpi[[#All],[account number]:[prin]],13,0)', # BV
                f'=VLOOKUP($BJ{r_idx},E{r_idx}:BA{r_idx},48,0)', # BW (75)
                f'=VLOOKUP($BJ{r_idx},E{r_idx}:BA{r_idx},49,0)', # BX (76)
                "", # BY
                f'=IF(COUNTIF(NOTES!Q:Q,\'VOLARE EXTRACTION\'!$J{r_idx}) > 0, "POSITIVE", "NEGATIVE")', # BZ (78)
                "", "", "", # CA, CB, CC
                f'=$B{r_idx}', f'=RIGHT($C{r_idx}, 2)' # CD, CE
            ]
            row_vals.extend(formulas)
            data_batch.append(row_vals)
        
        if data_batch:
            target_range = ws.Range(ws.Cells(start_row, 1), ws.Cells(start_row + len(data_batch) - 1, 83))
            target_range.Formula = data_batch
            target_range.Copy()
            target_range.PasteSpecial(Paste=-4163)
            
            # --- PHASE 2: SURGICAL PANDAS CLEANUP ---
            df_final = pd.DataFrame(target_range.Value)
            # Indices: K=10, BO=66, BP=67, BR=69, BS=70, BW=74, BX=75, BZ=77
            
            # 1. Action Filters
            df_final = df_final[~df_final[66].astype(str).str.upper().isin(['EXLUDED', 'EXCLUDED', '#N/A', '-2146826246'])]
            
            # 2. Clear PTP for non-PTP
            df_final.loc[df_final[66].astype(str).str.upper() != 'PTP', [69, 70]] = ""
            
            # 3. Clear Reaction "0"
            df_final.loc[df_final[67].astype(str).isin(['0', '0.0', '#N/A', '-2146826246']), 67] = ""
            
            # 4. CRUCIAL: Clear BW/BX "0" & "UNCATEGORIZED" FIRST
            for col in [74, 75]:
                df_final.loc[df_final[col].astype(str).str.upper().isin(['0', '0.0', 'UNCATEGORIZED', '#N/A']), col] = ""
            
            # 5. SkipTrace Logic
            df_final.loc[df_final[77].astype(str).str.upper() == 'NEGATIVE', 77] = ""
            mask_positive = df_final[77].astype(str).str.upper() == 'POSITIVE'
            mask_rfd_blank = df_final[74].astype(str).isin(['', 'None', 'nan'])
            
            # Remark Copy (K to BW)
            df_final.loc[mask_positive & mask_rfd_blank, 74] = df_final.loc[mask_positive & mask_rfd_blank, 10]
            
            # Last Word calculation for BZ
            def get_last_word(remark):
                words = str(remark).strip().split()
                return words[-1] if words else ""
            df_final.loc[mask_positive, 77] = df_final.loc[mask_positive, 10].apply(get_last_word)

            # 6. Delete row if PTP Amount is "0"
            df_final = df_final[~df_final[70].astype(str).isin(['0', '0.0'])]

            # 7. Collector Replacement
            if replacement_collector:
                mask_replace = df_final[65].astype(str).str.contains('|'.join(NAMES_TO_REPLACE), na=False)
                df_final.loc[mask_replace, 65] = replacement_collector

            # 8. Final Excel-safe pass (dates and nulls)
            for col in df_final.columns:
                if pd.api.types.is_datetime64_any_dtype(df_final[col]):
                    df_final[col] = pd.to_datetime(df_final[col]).dt.tz_localize(None)
            
            # Final null to None conversion
            df_final_clean = df_final.where(pd.notnull(df_final), None)

            # --- PHASE 3: FINAL FLUSH ---
            ws.Range("A3:CE20000").ClearContents()
            if not df_final_clean.empty:
                ws.Range(ws.Cells(3, 1), ws.Cells(3 + len(df_final_clean) - 1, 83)).Value = df_final_clean.values.tolist()
                for i, is_pos in enumerate(mask_positive):
                    if is_pos: ws.Cells(3 + i, 78).Interior.Color = 65535 # Yellow
            
        wb.Save()
        wb.Close()
        return True
    except Exception as e:
        st.error(f"Error: {e}")
        return False
    finally:
        try: excel.Quit()
        except: pass
        pythoncom.CoUninitialize()

def process_campaign_split(master_path, save_dir, date_str, bucket_list, password):
    if not WIN32COM_AVAILABLE: return False
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        abs_master = os.path.abspath(master_path)
        wb_master = excel.Workbooks.Open(Filename=abs_master, UpdateLinks=0, ReadOnly=True, Password=password)
        ws_master = wb_master.Sheets("VOLARE EXTRACTION")
        last_row = ws_master.Cells(ws_master.Rows.Count, "A").End(-4162).Row
        if last_row < 3:
            wb_master.Close(False)
            excel.Quit()
            return False
        
        data_range = ws_master.Range(ws_master.Cells(3, 1), ws_master.Cells(last_row, 83)).Value
        wb_master.Close(False)
        
        groups = {f"{b} Days": [] for b in bucket_list}
        for row_vals in data_range:
            if not row_vals: continue
            bucket_val = str(row_vals[57]).strip() if row_vals[57] else ""
            if bucket_val in groups: groups[bucket_val].append(list(row_vals))
        
        for b in bucket_list:
            rows = groups[f"{b} Days"]
            if not rows: continue
            campaign_tmpl = CAMPAIGN_TEMPLATE.replace("{date}", date_str).replace("{bucket}", str(b))
            bucket_filename = CAMPAIGN_FILENAME_PATTERN.format(date=date_str, bucket=b)
            save_path = os.path.abspath(os.path.join(save_dir, bucket_filename))
            
            wb_bucket = excel.Workbooks.Open(Filename=campaign_tmpl, UpdateLinks=0, ReadOnly=True)
            ws_bucket = wb_bucket.Sheets("VOLARE EXTRACTION")
            
            bucket_rows = []
            for r_master in rows:
                new_row = [""] * 31
                new_row[0:22] = r_master[56:78]
                new_row[26] = r_master[81]
                new_row[30] = r_master[82]
                bucket_rows.append(tuple(new_row))
            
            ws_bucket.Range(ws_bucket.Cells(4, 1), ws_bucket.Cells(4 + len(bucket_rows) - 1, 31)).Value = bucket_rows
            
            # Highlight POSITIVE skiptrace in buckets too
            for i, r_master in enumerate(rows):
                val = str(r_master[77]).strip().upper()
                if val and val not in ("NEGATIVE", "0", "FALSE", "#N/A", "NONE", ""):
                    ws_bucket.Cells(4 + i, 22).Interior.Color = 65535 # Yellow
            
            ws_bucket.Columns("A:AE").AutoFit()
            wb_bucket.SaveAs(save_path, FileFormat=51, Password=password)
            wb_bucket.Close()
            st.success(f" Generated: {os.path.basename(save_path)}")
        
        excel.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception: return False

def main():
    if not os.path.exists(SAVE_PATH): os.makedirs(SAVE_PATH)
    st.title("Daily Prod Automation")
    uploaded_file = st.file_uploader(" Drag and drop your file", type=["xlsx", "xls", "csv"])
    if uploaded_file:
        st.subheader(" Collector Settings")
        replacement_name = st.selectbox("Choose replacement name:", ["Select Name...", "Keep Original", "Bandola, Diana", "Reyes, Berlyn", "Galang, Jayson", "Antonio, Christine"])
        if replacement_name == "Select Name...": st.stop()
        if replacement_name == "Keep Original": replacement_name = None
        
        with st.spinner(" Processing..."):
            file_ext = uploaded_file.name.split('.')[-1]
            df = pd.read_csv(uploaded_file) if file_ext == 'csv' else pd.read_excel(uploaded_file, engine='openpyxl')
            df.columns = df.columns.str.strip()
            
            if "Status" in df.columns and "Remark By" in df.columns:
                mask_status = df['Status'].astype(str).str.startswith(tuple(EXCLUDED_STATUSES))
                mask_remark = df['Remark By'].astype(str).str.contains('SMGONZALES|CMENRIQUEZ', na=False)
                df_cleaned = df[~(mask_status | mask_remark)].copy()
                
                if "Remark" in df_cleaned.columns:
                    classifications = df_cleaned['Remark'].apply(classify_remark)
                    df_cleaned["RFD"], df_cleaned["SUB-RFD"] = zip(*classifications)
                
                # Format account number
                if "Account No." in df_cleaned.columns:
                    df_cleaned["Account No."] = "'000000" + df_cleaned["Account No."].astype(str).str.replace(r'\.0$', '', regex=True)
                
                report_date_obj, date_str = get_report_date()
                master_filename = MASTER_FILENAME_PATTERN.format(date=date_str)
                tmp_master = os.path.join(SAVE_PATH, master_filename)
                
                if process_template(df_cleaned, MASTER_TEMPLATE, tmp_master, replacement_name, password=PASSWORD):
                    process_campaign_split(tmp_master, SAVE_PATH, date_str, BUCKETS_TO_PROCESS, password=PASSWORD)
                    st.success(" Automation Complete!")
                else: st.error(" Master calculation failed.")
            else: st.error(" Required columns not found.")

# --- CONFIGURATION ---
SAVE_PATH = r"C:\Users\SPM\Downloads\BPI\PL_Daily\APRIL"
PASSWORD = "MAD_2Q2026"
NAMES_TO_REPLACE = ["Mendoza, Joshua"]
MASTER_FILENAME_PATTERN = "Master_Calculation_{date}v.xlsm"
CAMPAIGN_FILENAME_PATTERN = "PL_{date}_{bucket}_Madridv.xlsx"
MASTER_TEMPLATE = r"C:\Users\SPM\Downloads\BPI\Template\ONE PROD REPORT TEMPLATE.xlsm"
CAMPAIGN_TEMPLATE = r"C:\Users\SPM\Downloads\BPI\Template\PL_2026{date}_{bucket}_Madrid.xlsx"
BUCKETS_TO_PROCESS = [120, 150, 180]

if __name__ == "__main__":
    main()