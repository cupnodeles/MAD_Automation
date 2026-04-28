import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
import io
import os
import shutil
import tempfile
from datetime import datetime, timedelta

# Set page configuration
st.set_page_config(page_title="Daily Prod Automation", page_icon=None, layout="wide")

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

# --- DATE HELPER FOR FILENAMES ---
def get_report_date():
    """Returns (report_date_obj, filename_string) based on user's Monday/Friday logic."""
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        report_date = today - timedelta(days=3) # Use Friday last week
    else:
        report_date = today - timedelta(days=1) # Use yesterday
    return report_date, report_date.strftime("%Y%m%d")


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
            except PermissionError:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    st.error(f"Cannot copy template file after {max_retries} attempts.")
                    return False
        
        # Extract columns from S.No to SUB-RFD
        if 'S.No' in df_data.columns and 'SUB-RFD' in df_data.columns:
            start_idx = df_data.columns.get_loc('S.No')
            end_idx = df_data.columns.get_loc('SUB-RFD') + 1
            df_extracted_save = df_data.iloc[:, start_idx:end_idx]
        else:
            df_extracted_save = df_data.iloc[:, :53]
        
        num_cols = df_extracted_save.shape[1]
        num_rows = df_extracted_save.shape[0]
        
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
                except Exception:
                    if attempt < max_retries - 1:
                        time.sleep(retry_delay)
                    else:
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
                
                ws.Range(start_cell_obj, end_cell_obj).Formula = data_batch
                
                # Compute formulas and bake them immediately into permanent flat values
                try:
                    full_range = ws.Range(start_cell_obj, end_cell_obj)
                    full_range.Copy()
                    full_range.PasteSpecial(Paste=-4163)  # xlPasteValues
                    excel.Application.CutCopyMode = False
                except Exception:
                    pass
                
                # Apply formatting: Copy the red background format from BT1 to all pasted rows
                try:
                    ws.Range("BT1").Copy()
                    ws.Range(f"BT{start_row}:BT{last_row_pasted}").PasteSpecial(Paste=-4122)  # xlPasteFormats
                    ws.Range(f"CA{start_row}:CC{last_row_pasted}").PasteSpecial(Paste=-4122)
                    excel.Application.CutCopyMode = False
                except Exception:
                    pass
                
                # =========================================================
                # POST-PROCESSING STEPS
                # =========================================================

                # Grab the template formula from I1 for use on POSITIVE SKIPTRACE rows
                try:
                    source_i1 = ws.Range("I1")
                    formula_i1 = source_i1.Formula
                    formula_i1_r1c1 = source_i1.FormulaR1C1
                except Exception:
                    formula_i1 = ""
                    formula_i1_r1c1 = ""

                # Recalculate last_row before starting post-processing steps
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
                            to_delete.reverse()
                            for r in to_delete:
                                ws.Rows(r).Delete()
                except Exception:
                    pass

                # Recalculate last_row after Step 1 deletion shift
                last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                if last_row < 3: 
                    last_row = 3

                # Step 2: Clear PTP DATE (BR) and PTP AMOUNT (BS) where ACTION (BO) != "PTP"
                try:
                    vals = ws.Range(f"BO3:BO{last_row}").Value
                    if vals:
                        for i, row_val in enumerate(vals):
                            val = row_val[0]
                            val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                            if val_str != "PTP" and val_str != "":
                                ws.Range(f"BR{3+i}:BS{3+i}").ClearContents()
                except Exception:
                    pass

                # Step 3: Clear REACTION CODE (BP) where it is "0" or "#N/A"
                try:
                    vals = ws.Range(f"BP3:BP{last_row}").Value
                    if vals:
                        for i, row_val in enumerate(vals):
                            val = row_val[0]
                            val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                            if val == -2146826246 or val_str in ("0", "0.0", "#N/A"):
                                ws.Range(f"BP{3+i}").ClearContents()
                except Exception:
                    pass

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
                            to_delete.reverse()
                            for r in to_delete:
                                ws.Rows(r).Delete()
                except Exception:
                    pass

                # Recalculate last_row after Step 4 deletion shift
                last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                if last_row < 3: 
                    last_row = 3

                # Step 5: Clear SKIPTRACE column (BZ) where it is "NEGATIVE" (instead of deleting rows)
                try:
                    vals = ws.Range(f"BZ3:BZ{last_row}").Value
                    if vals:
                        for i, row_val in enumerate(vals):
                            val = row_val[0]
                            val_str = str(val).strip().upper() if not isinstance(val, int) else ""
                            if val_str == "NEGATIVE":
                                ws.Range(f"BZ{3+i}").ClearContents()
                except Exception:
                    pass

                # Step 6: For rows where BZ (SKIPTRACE PLATFORM) contains a platform name
                try:
                    positive_rows = []
                    bz_vals = ws.Range(f"BZ3:BZ{last_row}").Value
                    if bz_vals:
                        for i, row_val in enumerate(bz_vals):
                            val = row_val[0]
                            val_str = str(val).strip().upper() if val not in (None, "") else ""
                            if val_str and val_str not in ("NEGATIVE", "0", "FALSE", "#N/A", "NONE"):
                                positive_rows.append(3 + i)

                    if positive_rows:
                        for row_num in positive_rows:
                            try:
                                remark_val = ws.Cells(row_num, 11).Value
                                ws.Cells(row_num, 75).Value = remark_val
                            except Exception:
                                pass

                            target = ws.Cells(row_num, 78)  # Column BZ

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

                            target.Copy()
                            target.PasteSpecial(Paste=-4163)  # xlPasteValues
                            excel.Application.CutCopyMode = False
                            target.Interior.Color = 65535  # Yellow
                except Exception:
                    pass

                # Step 7: Clear BW/BX values set to 0 or UNCATEGORIZED
                try:
                    for col in ("BW", "BX"):
                        vals = ws.Range(f"{col}3:{col}{last_row}").Value
                        if vals:
                            for i, row_val in enumerate(vals):
                                val = row_val[0]
                                val_str = str(val).strip().upper() if val is not None else ""
                                if val_str in ("0", "0.0", "UNCATEGORIZED"):
                                    ws.Range(f"{col}{3+i}").ClearContents()
                except Exception:
                    pass

                # Step 9: Delete whole row of #N/A at the COLLECTOR at BN column (66)
                try:
                    last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                    if last_row >= 3:
                        vals = ws.Range(f"BN3:BN{last_row}").Value
                        if vals:
                            to_delete = []
                            for i, row_val in enumerate(vals):
                                val = row_val[0]
                                if val == -2146826246 or str(val).strip().upper() == "#N/A":
                                    to_delete.append(3 + i)
                            if to_delete:
                                to_delete.reverse()
                                for r in to_delete:
                                    ws.Rows(r).Delete()
                except Exception:
                    pass

                # Step 10: Replace Target Collectors (based on configuration)
                if replacement_collector:
                    try:
                        last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                        if last_row >= 3:
                            vals = ws.Range(f"BN3:BN{last_row}").Value
                            if vals:
                                for i, row_val in enumerate(vals):
                                    val = str(row_val[0]).strip()
                                    if any(name in val for name in NAMES_TO_REPLACE):
                                        ws.Cells(3 + i, 66).Value = replacement_collector
                    except Exception:
                        pass

                # Step 11: Clear cells in TYPE OF COLLECTION EFFORT (BL=64) if #N/A
                try:
                    last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                    if last_row >= 3:
                        vals = ws.Range(f"BL3:BL{last_row}").Value
                        if vals:
                            for i, row_val in enumerate(vals):
                                val = row_val[0]
                                if val == -2146826246 or str(val).strip().upper() == "#N/A":
                                    ws.Cells(3 + i, 64).ClearContents()
                except Exception:
                    pass

                # Step 12: Delete rows where BUCKET (BJ=62) is #N/A
                try:
                    last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                    if last_row >= 3:
                        vals = ws.Range(f"BF3:BF{last_row}").Value
                        if vals:
                            to_delete = []
                            for i, row_val in enumerate(vals):
                                val = row_val[0]
                                if val == -2146826246 or str(val).strip().upper() == "#N/A":
                                    to_delete.append(3 + i)
                            if to_delete:
                                to_delete.reverse()
                                for r in to_delete:
                                    ws.Rows(r).Delete()
                except Exception:
                    pass
                
                # Calculation readback
                calc_dict = {}
                if return_buckets:
                    try:
                        last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                        if last_row >= start_row:
                            all_vals = ws.Range(f"A{start_row}:ZZ{last_row}").Value
                            if all_vals:
                                for row_vals in all_vals:
                                    bucket_val = str(row_vals[57]).strip() if row_vals[57] else ""
                                    orig_idx_val = row_vals[701]
                                    if orig_idx_val is not None:
                                        calc_dict[int(orig_idx_val)] = bucket_val
                    except Exception:
                        pass

                try:
                    ws.Range("A1").Select()
                except Exception:
                    pass
            
            wb.Save()
            wb.Close(False)
            
            if return_buckets:
                return True, calc_dict
            
            st.session_state['last_master_path'] = output_path
            return True
            
        except Exception:
            return False
            
        finally:
            if wb:
                try: wb.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            pythoncom.CoUninitialize()
            
    except Exception:
        return False

def process_campaign_split(master_path, save_dir, date_str, bucket_list, password):
    """Reads the Master file once and splits it into Campaign files with specific formatting."""
    if not WIN32COM_AVAILABLE:
        st.error("pywin32 is required.")
        return False
        
    import win32com.client
    import pythoncom
    
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.Interactive = False   
        
        abs_master = os.path.abspath(master_path)
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
            ws_master = wb_master.Sheets(1)
        
        last_row = ws_master.Cells(ws_master.Rows.Count, "A").End(-4162).Row
        if last_row < 3:
            wb_master.Close(False)
            excel.Quit()
            return False
        
        data_range = ws_master.Range(ws_master.Cells(3, 1), ws_master.Cells(last_row, 83)).Value
        wb_master.Close(False)
        
        if not data_range:
            excel.Quit()
            return False

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
                continue
            
            # Use the setting from the configuration block
            campaign_tmpl = CAMPAIGN_TEMPLATE.replace("{date}", date_str).replace("{bucket}", str(b))
            
            if not os.path.exists(campaign_tmpl):
                st.error(f" Campaign template not found: {os.path.basename(campaign_tmpl)}")
                continue

            bucket_filename = f"PL_{date_str}_{b}_Madrid.xlsx"
            save_path = os.path.abspath(os.path.join(save_dir, bucket_filename))

            try:
                wb_bucket = excel.Workbooks.Open(
                    Filename=campaign_tmpl,
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True
                )

                try:
                    ws_bucket = wb_bucket.Sheets("VOLARE EXTRACTION")
                except Exception:
                    ws_bucket = wb_bucket.Sheets(1)
                
                bucket_rows = []
                for r_master in rows:
                    new_row = [""] * 31
                    new_row[0:22] = r_master[56:78]
                    new_row[26] = r_master[81]
                    new_row[30] = r_master[82]
                    bucket_rows.append(tuple(new_row))
                
                rows_to_paste = tuple(bucket_rows)
                start_row_bucket = 4
                ws_bucket.Range(ws_bucket.Cells(start_row_bucket, 1), ws_bucket.Cells(start_row_bucket + len(rows_to_paste) - 1, 31)).Value = rows_to_paste
                
                try:
                    p_range = ws_bucket.Range(ws_bucket.Cells(start_row_bucket, 1), ws_bucket.Cells(start_row_bucket + len(rows_to_paste) - 1, 31))
                    p_range.Borders.LineStyle = 1
                    p_range.Borders.Weight = 2
                    for i in range(len(rows_to_paste)):
                        cell = ws_bucket.Cells(start_row_bucket + i, 22)
                        val = str(cell.Value).strip().upper()
                        if val and val not in ("NEGATIVE", "0", "FALSE", "#N/A", "NONE", "NAN", ""):
                            cell.Interior.Color = 65535 # Yellow
                            cell.HorizontalAlignment = -4108 # xlCenter
                    ws_bucket.Columns("A:AE").AutoFit()
                except Exception: pass
                
                wb_bucket.SaveAs(save_path, FileFormat=51, Password=password)
                wb_bucket.Close()
                st.success(f" Generated: {bucket_filename}")
            except Exception: pass
        
        excel.Interactive = True
        excel.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception:
        return False

def main():
    # Ensure save directory exists
    if not os.path.exists(SAVE_PATH):
        try: os.makedirs(SAVE_PATH)
        except Exception: pass

    st.title("Daily Prod Automation")
    st.markdown("Removes unwanted statuses and remark sources for cleaner data processing")

    # FILE UPLOAD
    uploaded_file = st.file_uploader(" Drag and drop your file (Excel or CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # --- COLLECTOR SETTINGS ---
        st.subheader(" Collector Settings")
        replacement_name = st.selectbox(
            "Choose replacement name for Target Collectors (NEW):",
            ["Select Name...", "Keep Original", "Bandola, Diana", "Reyes, Berlyn", "Galang, Jayson", "Antonio, Christine"]
        )

        if replacement_name == "Select Name...":
            st.info(" Please select a replacement name (or **'Keep Original'**) to proceed with the automation.")
            st.stop()

        if replacement_name == "Keep Original":
            replacement_name = None

        st.divider()

        # --- DATA PREPARATION & CLEANING ---
        file_ext = uploaded_file.name.split('.')[-1]
        with st.spinner(" Reading data..."):
            try:
                if file_ext == 'csv': df = pd.read_csv(uploaded_file)
                else: df = pd.read_excel(uploaded_file, engine='openpyxl')
                df.columns = df.columns.str.strip()
            except Exception as e:
                st.error(f"Error loading file: {e}")
                st.stop()

        if "Status" in df.columns and "Remark By" in df.columns:
            with st.spinner(" Scrubbing data..."):
                if 'Remark' in df.columns:
                    df['Remark'] = df['Remark'].astype(str).str[:250]
                mask_status = df['Status'].astype(str).str.startswith(tuple(EXCLUDED_STATUSES))
                if 'Date' in df.columns:
                    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%m/%d/%Y')
                mask_remark = df['Remark By'].astype(str).str.contains('SMGONZALES|CMENRIQUEZ', na=False)
                df_removed = df[mask_status | mask_remark]
                df_cleaned = df[~(mask_status | mask_remark)].copy()
                
                if "Remark" in df_cleaned.columns:
                    classifications = df_cleaned['Remark'].apply(classify_remark)
                    df_cleaned["RFD"], df_cleaned["SUB-RFD"] = zip(*classifications)
                    df_cleaned["RFD"] = df_cleaned["RFD"].fillna("UNCATEGORIZED")
                    df_cleaned["SUB-RFD"] = df_cleaned["SUB-RFD"].fillna("UNCATEGORIZED")
                if "Account No." in df_cleaned.columns:
                    df_cleaned["Account No."] = df_cleaned["Account No."].astype(str).str.replace(r'\.0$', '', regex=True)
                    df_cleaned["Account No."] = df_cleaned["Account No."].replace(['nan', 'None'], '')
                    df_cleaned = df_cleaned[df_cleaned["Account No."] != ""]
                    df_cleaned["Account No."] = "'000000" + df_cleaned["Account No."]
                if "Time" in df_cleaned.columns:
                    def format_time(t):
                        try:
                            t_str = pd.to_datetime(str(t), errors='coerce').strftime('%I:%M %p')
                            if pd.isna(t_str) or t_str == 'NaT': return str(t)
                            return f"'{t_str}"
                        except Exception: return str(t)
                    df_cleaned["Time"] = df_cleaned["Time"].apply(format_time)
                
                if 'S.No' in df_cleaned.columns and 'SUB-RFD' in df_cleaned.columns:
                    start_idx = df_cleaned.columns.get_loc('S.No')
                    end_idx = df_cleaned.columns.get_loc('SUB-RFD') + 1
                    df_extracted_display = df_cleaned.iloc[:, start_idx:end_idx]
                else:
                    df_extracted_display = df_cleaned.iloc[:, :53]

            # --- EXTRACTION TRIGGER ---
            report_date_obj, date_str = get_report_date()
            tmpl_ext = os.path.splitext(MASTER_TEMPLATE)[1] or ".xlsx"
            tmp_master = os.path.join(SAVE_PATH, f"Master_Calculation_{date_str}{tmpl_ext}")

            with st.spinner("Running Master calculation and generating bucket files..."):
                with st.expander("Extraction Log", expanded=False):
                    success = process_template(df_cleaned, MASTER_TEMPLATE, tmp_master, replacement_name, password=PASSWORD)
                    if success:
                        st.success(f"Master created: {tmp_master}")
                        process_campaign_split(tmp_master, SAVE_PATH, date_str, BUCKETS_TO_PROCESS, password=PASSWORD)
                    else:
                        st.error("Master calculation failed.")

            st.divider()
            c1, c2, c3 = st.columns(3)
            c1.metric("Original Rows", f"{len(df):,}")
            c2.metric("Kept Rows", f"{len(df_cleaned):,}")
            c3.metric("Removed Rows", f"{len(df_removed):,}", delta=f"-{len(df_removed)}", delta_color="inverse")
            
            st.divider()
            st.subheader(" Removed Rows")
            if len(df_removed) > 0:
                st.dataframe(df_removed.head(100), use_container_width=True)
            
            st.divider()
            st.subheader(" Template Paste Preview")
            st.dataframe(df_extracted_display.head(50), use_container_width=True)

        else:
            st.error(" Columns 'Status' and 'Remark By' not found.")
            st.write(f"Available columns: {list(df.columns)}")
    else:
        st.info("Upload your Daily Prod Excel File to begin.")

# ==========================================
#  CONFIGURATION 
# ==========================================
SAVE_PATH = r"C:\Users\SPM\Downloads\BPI\PL_Daily\APRIL"
PASSWORD = "MAD_2Q2026"
NAMES_TO_REPLACE = ["Mendoza, Joshua"]

# Template settings
MASTER_TEMPLATE = r"C:\Users\SPM\Downloads\BPI\Template\ONE PROD REPORT TEMPLATE.xlsm"
# Use {date} for YYYYMMDD and {bucket} for the bucket number (120, 150, etc)
CAMPAIGN_TEMPLATE = r"C:\Users\SPM\Downloads\BPI\Template\PL_2026{date}_{bucket}_Madrid.xlsx"
# Which buckets to split into separate files
BUCKETS_TO_PROCESS = [120, 150, 180]

if __name__ == "__main__":
    main()