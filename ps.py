import streamlit as st
import pandas as pd
import numpy as np
import time
import re
import io

# ==========================================
# 🛠️ EXCEL MULTI-TAB & CELL HIGHLIGHTING FUNCTION
# ==========================================
def convert_to_styled_excel(df_main, df_np=None, df_tax=None, edited_cells_dict=None, is_module2=False):
    output = io.BytesIO()
    if edited_cells_dict is None:
        edited_cells_dict = {}

    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        # --- TAB 1: Report ---
        df_main_reset = df_main.reset_index()
        df_main_reset.to_excel(writer, index=False, sheet_name='Report')
        workbook = writer.book
        worksheet_main = writer.sheets['Report']
        
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#5E239D', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1})
        highlight_format = workbook.add_format({'border': 1, 'bg_color': '#FFE4E6'})

        for col_num, value in enumerate(df_main_reset.columns.values):
            worksheet_main.write(0, col_num, value, header_format)
            col_series = df_main_reset[value].fillna("").astype(str).map(str)
            max_len = max(col_series.map(len).max(), len(str(value)))
            worksheet_main.set_column(col_num, col_num, min(max_len + 2, 50))
            
            for row_num, row_idx in enumerate(df_main_reset.index):
                emp_id = str(df_main_reset.loc[row_idx, 'Employee ID']) if 'Employee ID' in df_main_reset.columns else str(df_main_reset.loc[row_idx, 'Base_ID'])
                cell_val = df_main_reset.iloc[row_num, col_num]
                
                if emp_id in edited_cells_dict and value in edited_cells_dict[emp_id]:
                    worksheet_main.write(row_num + 1, col_num, cell_val, highlight_format)
                else:
                    worksheet_main.write(row_num + 1, col_num, cell_val, cell_format)

        # --- TAB 2: NP- Absconding cases ---
        if df_np is not None and not df_np.empty:
            df_np_reset = df_np.reset_index()
            df_np_reset.to_excel(writer, index=False, sheet_name='NP- Absconding cases')
            worksheet_np = writer.sheets['NP- Absconding cases']
            
            for col_num, value in enumerate(df_np_reset.columns.values):
                worksheet_np.write(0, col_num, value, header_format)
                col_series = df_np_reset[value].fillna("").astype(str).map(str)
                max_len = max(col_series.map(len).max(), len(str(value)))
                worksheet_np.set_column(col_num, col_num, min(max_len + 2, 50), cell_format)
                for row_num, cell_val in enumerate(df_np_reset[value]):
                    worksheet_np.write(row_num + 1, col_num, cell_val, cell_format)

        # --- TAB 3: Tax exemption ---
        if df_tax is not None and not df_tax.empty:
            df_tax_reset = df_tax.reset_index()
            df_tax_reset.to_excel(writer, index=False, sheet_name='Tax exemption')
            worksheet_tax = writer.sheets['Tax exemption']
            
            for col_num, value in enumerate(df_tax_reset.columns.values):
                worksheet_tax.write(0, col_num, value, header_format)
                col_series = df_tax_reset[value].fillna("").astype(str).map(str)
                max_len = max(col_series.map(len).max(), len(str(value)))
                worksheet_tax.set_column(col_num, col_num, min(max_len + 2, 50), cell_format)
                for row_num, cell_val in enumerate(df_tax_reset[value]):
                    worksheet_tax.write(row_num + 1, col_num, cell_val, cell_format)

        # --- TAB 4: State LE ---
        if is_module2:
            state_le_data = {
                'State': [
                    'Rajasthan', 'Chhattisgarh', 'Maharashtra', 'Uttarakhand', 'Delhi', 'Karnataka', 
                    'Chandigarh', 'Gujarat', 'Haryana', 'Punjab', 'Madhya Pradesh', 'Goa', 'Assam', 
                    'Bihar', 'Odisha', 'West Bengal', 'Jharkhand', 'Kerala', 'Tamil Nadu', 
                    'Andhra Pradesh', 'Telangana', 'Uttar Pradesh'
                ],
                'Annual Leave': [26, 16, 16, 16, 15, 15, 16, 16, 16, 16, 30, 15, 16, 16, 16, 14, 16, 12, 12, 15, 15, 15],
                'Per day leave': [
                    0.071232877, 0.043835616, 0.043835616, 0.043835616, 0.04109589, 0.04109589, 
                    0.043835616, 0.043835616, 0.043835616, 0.043835616, 0.082191781, 0.04109589, 
                    0.043835616, 0.043835616, 0.043835616, 0.038356164, 0.043835616, 0.032876712, 
                    0.032876712, 0.04109589, 0.04109589, 0.04109589
                ],
                'State Limit': [45, 90, 45, 45, 45, 45, 45, 63, 45, 45, 90, 45, 45, 45, 45, 45, 45, 45, 45, 60, 60, 45]
            }
            df_state_le = pd.DataFrame(state_le_data)
            df_state_le.to_excel(writer, index=False, sheet_name='State LE')
            worksheet_state = writer.sheets['State LE']
            
            for col_num, value in enumerate(df_state_le.columns.values):
                worksheet_state.write(0, col_num, value, header_format)
                col_series = df_state_le[value].fillna("").astype(str).map(str)
                max_len = max(col_series.map(len).max(), len(str(value)))
                worksheet_state.set_column(col_num, col_num, min(max_len + 2, 50), cell_format)
                for row_num, cell_val in enumerate(df_state_le[value]):
                    worksheet_state.write(row_num + 1, col_num, cell_val, cell_format)

    return output.getvalue()

# --- PAGE CONFIGURATION & PREMIUM CSS ---
st.set_page_config(page_title="PhonePe FnF Pro", page_icon="📱", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    @keyframes gradientBG { 0% {background-position: 0% 50%;} 50% {background-position: 100% 50%;} 100% {background-position: 0% 50%;} }
    .main-header {
        background: linear-gradient(-45deg, #5E239D, #9D4EDD, #FF007F, #5E239D);
        background-size: 300% 300%; animation: gradientBG 8s ease infinite;
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        text-align: center; font-size: 3.5rem; font-weight: 900; padding-top: 10px;
    }
    .sub-header { text-align: center; color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom: 30px; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E5E7EB; box-shadow: 2px 0 15px rgba(0,0,0,0.05); }
    div.stButton > button { border-radius: 10px !important; font-weight: 700 !important; background-color: #5E239D !important; color: white !important; transition: transform 0.2s; }
    div.stButton > button:hover { transform: scale(1.02); }
    .feature-card { background: white; padding: 25px; border-radius: 12px; border-top: 4px solid #5E239D; box-shadow: 0 4px 6px rgba(0,0,0,0.05); text-align: center; height: 100%; transition: 0.3s; }
    .feature-card:hover { box-shadow: 0 8px 15px rgba(0,0,0,0.1); }
    .alert-popup { background: linear-gradient(135deg, #FFF1F2 0%, #FFE4E6 100%); border: 2px solid #F43F5E; border-radius: 12px; padding: 25px; margin: 20px 0; text-align: center; animation: pulse-red 2s infinite; }
    @keyframes pulse-red { 0% { box-shadow: 0 0 0 0 rgba(244, 63, 94, 0.4); } 70% { box-shadow: 0 0 0 15px rgba(244, 63, 94, 0); } 100% { box-shadow: 0 0 0 0 rgba(244, 63, 94, 0); } }
    </style>
""", unsafe_allow_html=True)

# --- SESSION STATE MANAGEMENT ---
if 'current_page' not in st.session_state: st.session_state.current_page = "Home"
if 'absconding_decision' not in st.session_state: st.session_state.absconding_decision = 'pending'
if 'edited_cells' not in st.session_state: st.session_state.edited_cells = {}
if 'actual_doe_inputs' not in st.session_state: st.session_state.actual_doe_inputs = {}
if 'np_absconding_df' not in st.session_state: st.session_state.np_absconding_df = None
if 'tax_exemption_df' not in st.session_state: st.session_state.tax_exemption_df = None
if 'allowed_ids' not in st.session_state: st.session_state.allowed_ids = []

def nav_home(): st.session_state.current_page = "Home"
def nav_factor(): st.session_state.current_page = "Factor"
def nav_master(): st.session_state.current_page = "Master"

# --- HELPERS ---
def load_file(uploaded_file, sheet_name=None):
    uploaded_file.seek(0)
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file, encoding='latin1', on_bad_lines='skip')
    if sheet_name:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name)
    return pd.read_excel(uploaded_file)

def standardize_id(df, possible_names, report_name):
    def scrub(val): return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower()
    clean_possible = [scrub(p) for p in possible_names]
    id_indices = [i for i, h in enumerate(df.columns) if scrub(h) in clean_possible]
    if not id_indices:
        for i, row in df.head(50).iterrows():
            found_indices = [idx for idx, val in enumerate(row.values) if scrub(val) in clean_possible]
            if found_indices:
                df.columns = [str(v).strip() for v in row.values]
                df = df.iloc[i+1:].reset_index(drop=True)
                id_indices = found_indices
                break
    if id_indices:
        df = df.copy()
        df['Base_ID'] = df.iloc[:, id_indices[0]].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        # Ensure ID text normalization (strip spaces and lowercase)
        df['Base_ID'] = df['Base_ID'].str.lower()
        
        # 🔧 SPECIFIC EDGE CASE FIX FOR HC REPORT (Adds missing 'p' prefix if it's purely digits)
        if report_name == "HC Report":
            df['Base_ID'] = df['Base_ID'].apply(lambda x: f"p{x}" if x.isdigit() and not x.startswith('p') else x)
            
        return df.dropna(subset=['Base_ID']).reset_index(drop=True)
    else:
        st.error(f"🚨 Could not locate ID in {report_name}"); st.stop()

def find_exact_or_fuzzy_column(df, target_name):
    cleaned_target = str(target_name).strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == cleaned_target:
            return col
    for col in df.columns:
        if cleaned_target in str(col).strip().lower():
            return col
    return None

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h2 style='color:#5E239D; text-align:center;'>🧭 Workspace</h2>", unsafe_allow_html=True)
    st.button("🏠 Home", on_click=nav_home, use_container_width=True)
    st.button("📊 1. Factor Calculator", on_click=nav_factor, use_container_width=True)
    st.button("📂 2. Master Builder", on_click=nav_master, use_container_width=True)
    st.markdown("---")
    st.link_button("🌐 Checkpoint portal", "https://register-validations-phonepe.streamlit.app/", use_container_width=True)
    st.link_button("🌴 Leave Encashment Calcy", "https://leave-encashment-calcy-fpt6j8bbzgdenee2jigjgl.streamlit.app/", use_container_width=True)

st.markdown("<div class='main-header'>PS Exits - FnF</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Automated Settlement Calculation & Consolidation</div>", unsafe_allow_html=True)

# ==========================================
# 🧭 PAGE ROUTING GATEWAY
# ==========================================

# --- HOME PAGE MODULE ---
if st.session_state.current_page == "Home":
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='feature-card'><h3>📊 Calculator</h3><p>Process Part A/B Factors and calculate payout data.</p></div>", unsafe_allow_html=True)
        st.button("Launch Calculator", on_click=nav_factor, use_container_width=True, key="home_calc")
    with c2:
        st.markdown("<div class='feature-card'><h3>📂 Consolidator</h3><p>Build Master FnF Reports from mapped resignation, HC and NDC data.</p></div>", unsafe_allow_html=True)
        st.button("Launch Consolidator", on_click=nav_master, use_container_width=True, key="home_master")

# --- MODULE 1: FACTOR CALCULATOR (UPDATED FORMAT MAPPINGS) ---
elif st.session_state.current_page == "Factor":
    st.markdown("### 📊 Step 1: Factor Data Calculator")
    col1, col2 = st.columns(2)
    with col1: ctc_file = st.file_uploader("1. Upload CTC Report", type=["xlsx", "csv"])
    with col2: raw_f_files = st.file_uploader("2. Raw Factor Data", type=["xlsx", "csv"], accept_multiple_files=True)

    if ctc_file and raw_f_files:
        if st.button("🔍 Extract & Preview Inputs", use_container_width=True):
            df_ctc = standardize_id(load_file(ctc_file), ['Employee Code', 'Emp ID', 'Employee ID', 'useremployeeid', 'emp code', 'id'], "CTC Report")
            
            # 🔄 INTEGRATED NEW COLUMN MAPPINGS
            ctc_rename_map = {
                '|Indicative Base Pay': 'Basic Pay',
                'O00006|House Rent Allowance': 'HRA Sales',
                '|Consistency Allowance Part B': 'Consistency Allowance',
                'Statutory Bonus - SB': 'Adv Stt Bonus SalesMaster',
                '|Sales Linked Commission - Part C': 'Sales Linked Comm. Master',
                'O00003|Mobile Allowance': 'Mobile Allow Sales Master'
            }
            for old, new in ctc_rename_map.items():
                if old in df_ctc.columns: df_ctc.rename(columns={old: new}, inplace=True)
            
            ctc_req_cols = ['Base_ID', 'Basic Pay', 'HRA Sales', 'Consistency Allowance', 'Adv Stt Bonus SalesMaster', 'Sales Linked Comm. Master', 'Mobile Allow Sales Master']
            for c in ctc_req_cols:
                if c not in df_ctc.columns: df_ctc[c] = 0
            df_ctc_clean = df_ctc[ctc_req_cols]

            def get_f_vals(df, letter):
                clean_columns = df.columns.str.strip().str.lower().str.replace('_', ' ')
                valid_patterns = [f'part {letter}', f'part{letter}']
                idx = [i for i, col in enumerate(clean_columns) if any(p in col for p in valid_patterns)]
                if idx:
                    return pd.to_numeric(df.iloc[:, idx[0]], errors='coerce').fillna(0).astype(float).values
                return np.zeros(len(df))

            all_f = []
            for f in raw_f_files:
                df_f = standardize_id(load_file(f), ['EMP Code/BT ID', 'Emp ID', 'Employee ID', 'emp id', 'employee id'], f.name)
                all_f.append(pd.DataFrame({'Base_ID': df_f['Base_ID'].values, 'Part A': get_f_vals(df_f, 'a'), 'Part B': get_f_vals(df_f, 'b'), 'Part C': get_f_vals(df_f, 'c')}))
            
            merged = pd.concat(all_f).drop_duplicates('Base_ID').merge(df_ctc_clean, on='Base_ID', how='left').fillna(0)
            
            for col in ['Part A', 'Part B', 'Part C']:
                merged[col] = merged[col].astype(float).round(3)

            st.session_state.raw_factor_merged = merged
            st.session_state.factor_preview_ready = True

    if st.session_state.get('factor_preview_ready'):
        edited = st.data_editor(
            st.session_state.raw_factor_merged, 
            use_container_width=True, 
            key="f_editor",
            column_config={
                "Part A": st.column_config.NumberColumn(format="%.3f"),
                "Part B": st.column_config.NumberColumn(format="%.3f"),
                "Part C": st.column_config.NumberColumn(format="%.3f")
            }
        )
        if st.button("✔️ Run Final Calculations", use_container_width=True):
            df = edited.copy()
            df['Monthly Basic'] = df['Basic Pay'] / 12
            df['Monthly Consistency'] = df['Consistency Allowance'] / 12
            df['Final Basic pay'] = (df['Part A'] * df['Monthly Basic']).round(0)
            df['Final Const. Bonus'] = (df['Part B'] * df['Monthly Consistency']).round(0)
            df['Final HRA'] = np.where(df['HRA Sales'] > 0, (0.05 * (df['Final Basic pay'] + df['Final Const. Bonus'])).round(0), 0)
            df['Final Sales linked'] = (df['Part C'] * (df['Sales Linked Comm. Master'] / 12)).round(0)
            df['Final Mobile allowance'] = ((df['Part A'] * df['Mobile Allow Sales Master']) / 12).round(0)
            df['Final Adv. stat bonus'] = np.where(df['Adv Stt Bonus SalesMaster'] > 0, (0.0833 * (df['Final Basic pay'] + df['Final Const. Bonus'])).round(0), 0)
            
            calculated_slice = df[[
                'Base_ID', 'Part A', 'Part B', 'Part C', 
                'Final Basic pay', 'Final HRA', 'Final Const. Bonus', 
                'Final Adv. stat bonus', 'Final Sales linked', 'Final Mobile allowance'
            ]].rename(columns={'Base_ID': 'Employee ID'})
            
            st.session_state.calculated_factor_df = calculated_slice
            st.session_state.factor_calc_done = True
            st.rerun()

    if st.session_state.get('factor_calc_done'):
        st.success("Calculations Complete!")
        st.dataframe(st.session_state.calculated_factor_df, use_container_width=True)
        st.download_button("📥 Download Factor Data", convert_to_styled_excel(st.session_state.calculated_factor_df), "Calculated_Factor_Data.xlsx", use_container_width=True)

# --- MODULE 2: MASTER BUILDER ---
elif st.session_state.current_page == "Master":
    st.markdown("### 📂 Step 2: Master Report Builder (Custom Mapping Module)")
    
    f_emp_list = st.file_uploader("📥 Step A: Upload Employee List (Base Match Roster)", type=["xlsx", "csv"])
    
    if f_emp_list is not None:
        df_target_list = standardize_id(load_file(f_emp_list), ['EMP ID', 'emp id', 'Employee ID', 'Employee Code', 'EMPLOYEECODE'], "Employee Filter List")
        st.session_state.allowed_ids = df_target_list['Base_ID'].astype(str).str.strip().tolist()
        
        st.success(f"🎯 Target Profile Roster Lock-in Successful! Identified: {len(st.session_state.allowed_ids)} unique records.")
        st.markdown("---")
        
        st.markdown("#### 📑 Step B: Upload Required Reports")
        col1, col2 = st.columns(2)
        with col1:
            f_res = st.file_uploader("1. Resignation Report", type=["xlsx", "csv"])
            f_hc = st.file_uploader("2. HC Report", type=["xlsx", "csv"])
        with col2:
            f_ndc = st.file_uploader("3. NDC Sheet", type=["xlsx", "csv"])

        if all([f_res, f_hc, f_ndc]):
            if st.button("🚀 Run Mapped Master Consolidation", use_container_width=True):
                with st.spinner("Consolidating structural mapping inputs..."):
                    
                    df_res_raw = standardize_id(load_file(f_res), ['Employee Code', 'Employee ID', 'emp id', 'id'], "Resignation Report")
                    df_hc_raw = standardize_id(load_file(f_hc), ['EMPLOYEECODE', 'Employee ID', 'user/employee id', 'id'], "HC Report")
                    df_ndc_raw = standardize_id(load_file(f_ndc), ['Employee ID', 'Employee Code', 'emp id', 'id'], "NDC Sheet")

                    master_rows = []

                    for emp_id in st.session_state.allowed_ids:
                        # Display structural ID in uppercase back on the UI
                        row_data = {'Employee ID': emp_id.upper()}
                        
                        sub_res = df_res_raw[df_res_raw['Base_ID'].astype(str).str.strip() == emp_id]
                        sub_hc = df_hc_raw[df_hc_raw['Base_ID'].astype(str).str.strip() == emp_id]
                        sub_ndc = df_ndc_raw[df_ndc_raw['Base_ID'].astype(str).str.strip() == emp_id]

                        # Designation (優先 check HC report if available, else Resignation)
                        c_desig_hc = find_exact_or_fuzzy_column(df_hc_raw, 'Designation')
                        c_desig_res = find_exact_or_fuzzy_column(df_res_raw, 'Designation')
                        if not sub_hc.empty and c_desig_hc:
                            row_data['Designation'] = str(sub_hc.iloc[0][c_desig_hc]).strip()
                        elif not sub_res.empty and c_desig_res:
                            row_data['Designation'] = str(sub_res.iloc[0][c_desig_res]).strip()
                        else:
                            row_data['Designation'] = ""

                        # Department (優先 check HC report if available, else Resignation)
                        c_dept_hc = find_exact_or_fuzzy_column(df_hc_raw, 'Department')
                        c_dept_res = find_exact_or_fuzzy_column(df_res_raw, 'Department')
                        if not sub_hc.empty and c_dept_hc:
                            row_data['Department'] = str(sub_hc.iloc[0][c_dept_hc]).strip()
                        elif not sub_res.empty and c_dept_res:
                            row_data['Department'] = str(sub_res.iloc[0][c_dept_res]).strip()
                        else:
                            row_data['Department'] = ""

                        # Date Of Joining
                        c_doj = find_exact_or_fuzzy_column(df_res_raw, 'Date Of Joining')
                        if not sub_res.empty and c_doj:
                            d_parsed = pd.to_datetime(sub_res.iloc[0][c_doj], dayfirst=True, errors='coerce')
                            row_data['Date Of Joining'] = d_parsed.strftime('%d-%m-%Y') if pd.notna(d_parsed) else str(sub_res.iloc[0][c_doj])
                        else:
                            row_data['Date Of Joining'] = ""

                        # Resignation Date
                        c_rdate = find_exact_or_fuzzy_column(df_res_raw, 'Resignation Date')
                        if not sub_res.empty and c_rdate:
                            d_parsed = pd.to_datetime(sub_res.iloc[0][c_rdate], dayfirst=True, errors='coerce')
                            row_data['Resignation Date'] = d_parsed.strftime('%d-%m-%Y') if pd.notna(d_parsed) else str(sub_res.iloc[0][c_rdate])
                        else:
                            row_data['Resignation Date'] = ""

                        # Final Approved LWD
                        c_lwd = find_exact_or_fuzzy_column(df_res_raw, 'Final Approved LWD')
                        if not sub_res.empty and c_lwd:
                            d_parsed = pd.to_datetime(sub_res.iloc[0][c_lwd], dayfirst=True, errors='coerce')
                            row_data['Final Approved LWD'] = d_parsed.strftime('%d-%m-%Y') if pd.notna(d_parsed) else str(sub_res.iloc[0][c_lwd])
                        else:
                            row_data['Final Approved LWD'] = ""

                        # Years(tenure) Formula
                        if not sub_res.empty and c_lwd and c_doj:
                            d_lwd_dt = pd.to_datetime(sub_res.iloc[0][c_lwd], dayfirst=True, errors='coerce')
                            d_doj_dt = pd.to_datetime(sub_res.iloc[0][c_doj], dayfirst=True, errors='coerce')
                            if pd.notna(d_lwd_dt) and pd.notna(d_doj_dt):
                                days_diff = (d_lwd_dt - d_doj_dt).days + 1
                                row_data['Years(tenure)'] = round(days_diff / 365.0, 2)
                            else:
                                row_data['Years(tenure)'] = 0.0
                        else:
                            row_data['Years(tenure)'] = 0.0

                        # NP recovery / NoOfWaivedOffDays
                        c_recovery_type = find_exact_or_fuzzy_column(df_res_raw, 'Recovery days Type')
                        c_waived_days = find_exact_or_fuzzy_column(df_res_raw, 'NoOfWaivedOffDays')
                        
                        if not sub_res.empty and c_recovery_type and c_waived_days:
                            rec_type_val = str(sub_res.iloc[0][c_recovery_type]).strip().lower()
                            raw_days_val = sub_res.iloc[0][c_waived_days]
                            
                            if 'waive off' in rec_type_val:
                                row_data['NP recovery'] = 0
                            elif 'recovery' in rec_type_val:
                                row_data['NP recovery'] = pd.to_numeric(raw_days_val, errors='coerce') if pd.notna(raw_days_val) else raw_days_val
                            else:
                                row_data['NP recovery'] = pd.to_numeric(raw_days_val, errors='coerce') if pd.notna(raw_days_val) else raw_days_val
                        else:
                            row_data['NP recovery'] = ""

                        row_data['NP payable'] = ""
                        row_data['Leave Encashment'] = ""
                        row_data['Onfield Allowance'] = 0
                        row_data['IT Asset Recovery'] = 0
                        row_data['Facility Recovery'] = 0

                        # Inventory Recovery
                        c_ndc_inv = find_exact_or_fuzzy_column(df_ndc_raw, 'Inventory Recovery')
                        if c_ndc_inv is None: c_ndc_inv = find_exact_or_fuzzy_column(df_ndc_raw, 'Device Recovery')
                        
                        if not sub_ndc.empty and c_ndc_inv:
                            row_data['Inventory Recovery'] = pd.to_numeric(sub_ndc.iloc[0][c_ndc_inv], errors='coerce') if pd.notna(sub_ndc.iloc[0][c_ndc_inv]) else sub_ndc.iloc[0][c_ndc_inv]
                        else:
                            row_data['Inventory Recovery'] = 0

                        row_data['Remarks'] = ""

                        c_reason = find_exact_or_fuzzy_column(df_res_raw, 'Reason')
                        if c_reason is None: c_reason = find_exact_or_fuzzy_column(df_res_raw, 'Separation Reason')
                        row_data['Reason'] = str(sub_res.iloc[0][c_reason]).strip() if (not sub_res.empty and c_reason) else ""

                        master_rows.append(row_data)

                    df_final_master = pd.DataFrame(master_rows)

                    final_req_cols = [
                        'Employee ID', 'Designation', 'Department', 'Date Of Joining', 'Resignation Date',
                        'Final Approved LWD', 'Years(tenure)', 'NP recovery', 'NP payable', 'Leave Encashment',
                        'Onfield Allowance', 'IT Asset Recovery', 'Facility Recovery', 'Inventory Recovery', 'Remarks', 'Reason'
                    ]

                    for col in final_req_cols:
                        if col not in df_final_master.columns:
                            df_final_master[col] = ""

                    st.session_state.final_master_df = df_final_master[final_req_cols].set_index('Employee ID')
                    st.session_state.absconding_decision = 'pending'
                    st.session_state.edited_cells = {}
                    st.session_state.actual_doe_inputs = {}
                    st.session_state.np_absconding_df = None
                    st.session_state.tax_exemption_df = None
                    st.rerun()

    if st.session_state.get('final_master_df') is not None:
        curr_df = st.session_state.final_master_df.reset_index()
        
        abs_mask = curr_df['Reason'].astype(str).str.contains('33|absconding|termination|resignation', case=False, na=False)
        total_absconding_cases = abs_mask.sum()
        
        edit_mask = abs_mask & (pd.to_numeric(curr_df['NP recovery'], errors='coerce').fillna(0) == 0)
        review_cases_data = curr_df[edit_mask]
        
        if total_absconding_cases > 0 and st.session_state.absconding_decision == 'pending':
            st.markdown(f"<div class='alert-popup'><h2>⚠️ Total separation/absconding cases found: {total_absconding_cases}</h2></div>", unsafe_allow_html=True)
            c1, r_btn = st.columns(2)
            with c1: 
                if st.button("✂️ Exclude all from Master", use_container_width=True):
                    st.session_state.final_master_df = curr_df[~abs_mask].set_index('Employee ID')
                    st.session_state.absconding_decision = 'tax_popup_stage'; st.rerun()
            with r_btn: 
                if st.button("📝 Review & Edit", use_container_width=True):
                    st.session_state.absconding_decision = 'review'; st.rerun()
        
        elif st.session_state.absconding_decision == 'review':
            st.markdown("### 📝 Edit Separation Cases")
            edited = st.data_editor(review_cases_data, use_container_width=True, key="abs_editor")
            if st.button("✔️ Save Edits & Proceed to Actual DOE", use_container_width=True):
                if 'Employee ID' in edited.columns:
                    edited_cols = edited.copy()
                else:
                    edited_cols = edited.reset_index()

                original_review = review_cases_data.set_index('Employee ID')
                comp_edited = edited_cols.set_index('Employee ID')
                
                cell_tracking = {}
                for emp_id in comp_edited.index:
                    cell_tracking[str(emp_id)] = []
                    for col in comp_edited.columns:
                        if col in original_review.columns:
                            if str(comp_edited.loc[emp_id, col]) != str(original_review.loc[emp_id, col]):
                                cell_tracking[str(emp_id)].append(col)
                
                st.session_state.edited_cells = cell_tracking
                curr_df.set_index('Employee ID', inplace=True)
                curr_df.update(comp_edited)
                st.session_state.final_master_df = curr_df
                st.session_state.absconding_decision = 'actual_doe_popup'
                st.rerun()

        elif st.session_state.absconding_decision == 'actual_doe_popup':
            st.markdown("### 📅 Enter Actual DOE Details")
            
            curr_df_with_idx = st.session_state.final_master_df.reset_index()
            target_mask = curr_df_with_idx['Reason'].astype(str).str.contains('33|absconding|termination|resignation', case=False, na=False) & (pd.to_numeric(curr_df_with_idx['NP recovery'], errors='coerce').fillna(0) == 0)
            target_employees = curr_df_with_idx[target_mask]
            
            form_dict = {}
            for _, row in target_employees.iterrows():
                emp_id = str(row['Employee ID'])
                form_dict[emp_id] = st.date_input(f"Employee ID: {emp_id} | Enter Actual DOE", value=None, key=f"doe_in_{emp_id}")

            if st.button("✔️ Save DOE & Proceed to Tax Exemption Configuration", use_container_width=True):
                np_rows = []
                for _, row in target_employees.iterrows():
                    emp_id = str(row['Employee ID'])
                    actual_doe_val = form_dict.get(emp_id)
                    lwd_sf_str = str(row['Final Approved LWD']).strip()
                    
                    actual_doe_dt = pd.to_datetime(actual_doe_val, errors='coerce')
                    res_doe_dt = pd.to_datetime(lwd_sf_str, dayfirst=True, errors='coerce')
                    
                    if pd.notna(actual_doe_dt) and pd.notna(res_doe_dt):
                        shortfall_days = (actual_doe_dt - res_doe_dt).days + 1
                    else:
                        shortfall_days = ""

                    np_rows.append({
                        'Employee ID': emp_id,
                        'DOE': lwd_sf_str if lwd_sf_str != "nan" else "",
                        'DOE as per resignation': actual_doe_dt.strftime('%d-%m-%Y') if pd.notna(actual_doe_dt) else "",
                        'Shortfall': shortfall_days
                    })
                
                st.session_state.np_absconding_df = pd.DataFrame(np_rows).set_index('Employee ID')
                st.session_state.absconding_decision = 'tax_popup_stage'
                st.rerun()

        elif st.session_state.absconding_decision == 'tax_popup_stage':
            st.markdown("### 📑 Step C: Tax Exemption Module Upload Configuration")
            st.info("Please supply the Gratuity & Leave Declarations form along with the updated HC Roster:")
            
            f_gratuity = st.file_uploader("1. Upload Gratuity & Leave Declaration Form", type=["xlsx", "csv"])
            f_hc_2 = st.file_uploader("2. Upload HC Report 2.0 Roster", type=["xlsx", "csv"])
            
            col_sb1, col_sb2 = st.columns(2)
            with col_sb1:
                run_tax_exemption = st.button("🚀 Process Tax Data", use_container_width=True, disabled=not (f_gratuity and f_hc_2))
            with col_sb2:
                skip_tax_exemption = st.button("⏩ Skip Step C & Generate Report", use_container_width=True)

            if run_tax_exemption and f_gratuity and f_hc_2:
                df_grat = load_file(f_gratuity)
                df_hc2_raw = load_file(f_hc_2)
                
                df_grat.columns = [str(c).strip() for c in df_grat.columns]
                df_hc2_raw.columns = [str(c).strip() for c in df_hc2_raw.columns]
                
                email_col_idx = [i for i, col in enumerate(df_grat.columns) if 'email address' in col.lower()]
                hc2_email_idx = [i for i, col in enumerate(df_hc2_raw.columns) if 'official email id' in col.lower()]
                hc2_empid_idx = [i for i, col in enumerate(df_hc2_raw.columns) if 'employee id' in col.lower() or 'user/employee id' in col.lower() or 'employeecode' in col.lower()]
                
                if email_col_idx and hc2_email_idx and hc2_empid_idx:
                    grat_email_col = df_grat.columns[email_col_idx[0]]
                    hc2_email_col = df_hc2_raw.columns[hc2_email_idx[0]]
                    hc2_empid_col = df_hc2_raw.columns[hc2_empid_idx[0]]
                    
                    hc2_mapping_slice = df_hc2_raw[[hc2_email_col, hc2_empid_col]].dropna(subset=[hc2_email_col])
                    
                    # Target and fix lookups for HC Report 2.0 as well if needed
                    hc2_mapping_slice[hc2_empid_col] = hc2_mapping_slice[hc2_empid_col].astype(str).str.strip().str.lower()
                    hc2_mapping_slice[hc2_empid_col] = hc2_mapping_slice[hc2_empid_col].apply(lambda x: f"p{x}" if x.isdigit() and not x.startswith('p') else x)
                    
                    hc2_map_dict = dict(zip(hc2_mapping_slice[hc2_email_col].astype(str).str.strip().str.lower(), hc2_mapping_slice[hc2_empid_col]))
                    
                    df_grat['lookup'] = df_grat[grat_email_col].astype(str).str.strip().str.lower().map(hc2_map_dict)
                    
                    allowed_set = set(str(i) for i in st.session_state.allowed_ids)
                    df_grat['input_sheet_lookup'] = df_grat['lookup'].apply(lambda x: x if str(x) in allowed_set else '0')
                    
                    tax_filtered_data = df_grat[df_grat['input_sheet_lookup'] != '0'].copy()
                    
                    if tax_filtered_data.empty:
                        st.warning("⚠️ No such matching employees found from the present cutoff.")
                        st.session_state.tax_exemption_df = None
                        time.sleep(2)
                    else:
                        def grab_fuzzy_tax_col(tokens):
                            for col in tax_filtered_data.columns:
                                if all(t.lower() in col.lower() for t in tokens):
                                    return col
                            return None
                        
                        le_flag = grab_fuzzy_tax_col(['availed', 'leave', 'encashment', 'tax'])
                        grat_flag = grab_fuzzy_tax_col(['availed', 'gratuity', 'tax'])
                        grat_amt = grab_fuzzy_tax_col(['sum', 'gratuity', 'received'])
                        le_amt = grab_fuzzy_tax_col(['sum', 'leave', 'encashment', 'received'])
                        
                        tax_output_rows = []
                        for _, row_tax in tax_filtered_data.iterrows():
                            tax_output_rows.append({
                                'EMP ID': str(row_tax['input_sheet_lookup']).replace('.0', '').upper(),
                                'Leave Encashment': str(row_tax[le_flag]) if le_flag else "",
                                'Gratuity': str(row_tax[grat_flag]) if grat_flag else "",
                                'Gratuity amount received from the Previous Employers': pd.to_numeric(row_tax[grat_amt], errors='coerce') if grat_amt else 0.0,
                                'Leave Encashment amount received from the Previous Employers': pd.to_numeric(row_tax[le_amt], errors='coerce') if le_amt else 0.0
                            })
                        
                        st.session_state.tax_exemption_df = pd.DataFrame(tax_output_rows).set_index('EMP ID')
                    
                    st.session_state.absconding_decision = 'done'
                    st.rerun()
                else:
                    st.error("🚨 Configuration Error: Key email parameters failed column evaluation logic.")
            
            elif skip_tax_exemption:
                st.session_state.tax_exemption_df = None
                st.session_state.absconding_decision = 'done'
                st.rerun()

        else:
            final_df_clean = st.session_state.final_master_df.copy()
            if 'Reason' in final_df_clean.columns:
                final_df_clean = final_df_clean.drop(columns=['Reason'])
                
            final = final_df_clean.reset_index()
            np_final = st.session_state.np_absconding_df
            tax_final = st.session_state.tax_exemption_df
            
            st.success("✅ Consolidation Finished Successfully! Workbook generation locks completed.")
            st.dataframe(final, use_container_width=True)
            
            if np_final is not None:
                st.markdown("#### 📑 Summary View: NP- Absconding cases Tab")
                st.dataframe(np_final.reset_index(), use_container_width=True)
                
            if tax_final is not None and not tax_final.empty:
                st.markdown("#### 📑 Summary View: Tax exemption Tab")
                st.dataframe(tax_final.reset_index(), use_container_width=True)
            elif tax_final is None:
                st.info("ℹ️ Note: Tax exemption processing omitted or returned no relevant roster records.")
            
            excel_bytes = convert_to_styled_excel(
                df_main=final_df_clean, 
                df_np=np_final, 
                df_tax=tax_final,
                edited_cells_dict=st.session_state.edited_cells,
                is_module2=True
            )
            
            st.download_button(
                label="📥 Download Final Mapped Multi-Tab Master Report", 
                data=excel_bytes, 
                file_name="Master_FnF_Report_Calculated.xlsx", 
                use_container_width=True
            )
