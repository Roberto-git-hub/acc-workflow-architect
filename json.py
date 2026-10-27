import streamlit as st
import pandas as pd
import json
import base64
import zlib
import io
import re
import requests
from requests.auth import HTTPBasicAuth
from collections import Counter
import zipfile
import os
import datetime
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as xlImage
import copy

# =====================================================================
# STREAMLIT INTERFACE - CONFIGURATION (WIDE LAYOUT)
# =====================================================================
st.set_page_config(page_title="Netbiis Automation", layout="wide", page_icon="⚙️", initial_sidebar_state="collapsed")

# --- ADVANCED CSS INJECTION (ULTRA PREMIUM UI) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"], .stMarkdown, div {
        font-family: 'Inter', sans-serif !important;
    }

    /* Hides the hamburger menu and default Streamlit footer for a cleaner look */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* More elegant titles */
    h1 {
        background: -webkit-linear-gradient(45deg, #0071e3, #40E0D0);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700 !important;
        letter-spacing: -1px;
    }
    h2, h3 {
        font-weight: 600 !important;
        color: #F5F5F7 !important;
        letter-spacing: -0.5px;
    }

    /* Main buttons styling */
    .stButton>button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
        background-color: #1E1E24 !important;
        color: #E2E2E2 !important;
    }
    .stButton>button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 8px 20px rgba(0, 113, 227, 0.2) !important;
        border-color: #0071e3 !important;
        color: #FFFFFF !important;
    }
    
    /* Primary buttons */
    .stButton>button[data-baseweb="button"].st-emotion-cache-1r6slb0 {
        background: linear-gradient(135deg, #0071e3, #005bb5) !important;
        color: white !important;
        border: none !important;
        box-shadow: 0 4px 10px rgba(0, 113, 227, 0.3) !important;
    }
    .stButton>button[data-baseweb="button"].st-emotion-cache-1r6slb0:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 8px 25px rgba(0, 113, 227, 0.5) !important;
    }

    /* Metric Cards (Dashboard style) */
    div[data-testid="metric-container"] {
        background: linear-gradient(145deg, #1A1C23, #121418);
        border: 1px solid #2A2D35;
        padding: 20px 25px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        transition: all 0.3s ease;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-5px);
        border-color: #0071e3;
        box-shadow: 0 8px 25px rgba(0, 113, 227, 0.15);
    }
    /* Metric label color */
    div[data-testid="metric-container"] > div > div > div {
        color: #8E8E93 !important;
        font-weight: 500 !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        background-color: #1A1C23;
        border-radius: 8px 8px 0px 0px;
        padding: 10px 20px;
        border: 1px solid #2A2D35;
        border-bottom: none;
        color: #8E8E93;
        font-weight: 500;
        transition: background-color 0.3s ease;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0071e3 !important;
        color: white !important;
        border-color: #0071e3 !important;
        font-weight: 600 !important;
    }
    
    /* Info/Success/Error boxes with softer corners */
    div.stAlert {
        border-radius: 10px !important;
        border: 1px solid rgba(255,255,255,0.05) !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
        backdrop-filter: blur(10px);
    }

    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #3A3D45 !important;
        border-radius: 12px !important;
        background-color: #181A20 !important;
        transition: all 0.3s ease !important;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #0071e3 !important;
        background-color: #1E222B !important;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================================
# LINK READER MODE (Operator)
# =====================================================================
if 'data' in st.query_params:
    try:
        compressed_data = base64.urlsafe_b64decode(st.query_params['data'])
        json_string = zlib.decompress(compressed_data).decode('utf-8')
        
        st.title("📄 JSON Viewer (Workflow)")
        st.success("JSON loaded successfully via secure link!")
        st.info("💡 **HOW TO COPY:** Hover over the top right corner of the black block below and click the clipboard icon (📋) to copy everything.")
        
        st.code(json_string, language='json')
        st.stop() 
    except Exception as e:
        st.error("⚠️ The link seems to be broken or corrupted. Please ask the creator to generate it again.")
        st.stop()

# =====================================================================
# TOP LOGO
# =====================================================================
if os.path.exists("logo_netbiis.png"):
    st.image("logo_netbiis.png", width=180)

# =====================================================================
# PROCESSING FUNCTIONS (Coordinator)
# =====================================================================
@st.cache_data
def extract_metadata(file_bytes, file_name, sheet_name=None):
    try:
        if file_name.endswith('.csv'):
            df_raw = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python', header=None)
        else:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    except:
        return {}

    meta = {
        "campaign_name": "", "deployment_date": "", "user": "", "cid": "",
        "notes": [], "include": [], "exclude": [], "audience": []
    }
    notes_col_idx = None
    current_mode = None 

    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(x).lower() for x in row.values if pd.notna(x)])
        if 'waterfall order' in row_str or 'dag segment name' in row_str:
            break

        col0 = str(row.values[0]).strip().lower() if pd.notna(row.values[0]) else ""
        col1 = str(row.values[1]).strip() if len(row.values) > 1 and pd.notna(row.values[1]) else ""

        if "deployment name" in col0:
            meta["campaign_name"] = re.sub(r'(?i)_v2$', '', col1).strip()
            for i, val in enumerate(row.values):
                if str(val).strip().lower() == 'notes': notes_col_idx = i
        elif "deployment date" in col0: meta["deployment_date"] = col1
        elif col0 == "user": meta["user"] = col1
        elif col0 == "cid": meta["cid"] = re.sub(r'(?i)_segment$', '', col1).strip()
        
        # 1. Identifies the Audience block exactly
        elif "audience selection" in col0: 
            current_mode = "audience"
            
        # 2. Accepts "INCLUDE" alone without breaking the Audience block
        elif col0 in ["only include", "include"]: 
            if current_mode != "audience":
                current_mode = "include"
            vals = [str(x).strip() for x in row.values[1:] if pd.notna(x) and str(x).strip().lower() != 'nan']
            if vals: meta[current_mode].append(" | ".join(vals))
            
        elif "exclude" in col0:
            current_mode = "exclude"
            vals = [str(x).strip() for x in row.values[1:] if pd.notna(x) and str(x).strip().lower() != 'nan']
            if vals: meta["exclude"].append(" | ".join(vals))
            
        elif col0 == "and" and current_mode:
            vals = [str(x).strip() for x in row.values[1:] if pd.notna(x) and str(x).strip().lower() != 'nan']
            if vals: meta[current_mode].append(" | ".join(vals))
            
        elif col0 and col0 not in ["and", "criteria:"]:
            current_mode = None

        if notes_col_idx is not None and len(row.values) > notes_col_idx:
            note_val = str(row.values[notes_col_idx]).strip()
            if pd.notna(row.values[notes_col_idx]) and note_val and note_val.lower() not in ['nan', 'notes']:
                if note_val not in meta["notes"]: meta["notes"].append(note_val)
    return meta
@st.cache_data
def load_and_clean_data(file_bytes, file_name, sheet_name=None):
    if file_name.endswith('.csv'):
        try: df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python')
        except: df = pd.read_csv(io.BytesIO(file_bytes), sep=';')
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)

    first_col = str(df.columns[0]).lower()
    if 'unnamed' in first_col or 'deployment' in first_col:
        for idx, row in df.iterrows():
            row_values = [str(val).lower().strip() for val in row.values]
            if any(k in row_values for k in ['dag segment name', 'segment name', 'cellname', 'waterfall order']):
                df.columns = df.iloc[idx]
                df = df.iloc[idx + 1:].reset_index(drop=True)
                break
    return df

def process_data(df):
    final_json_data = []
    waterfall_id = 1
    cols = {str(c).strip().lower(): c for c in df.columns}

    def get_val(row, possible_names):
        for name in possible_names:
            actual_col = cols.get(name.lower())
            if actual_col is not None:
                val = str(row.get(actual_col, '')).strip()
                if val.lower() != 'nan' and val != '':
                    return val
        return ''

    for _, row in df.iterrows():
        row_text = ' '.join([str(val).lower() for val in row.values])
        if 'reference for combination' in row_text or 'suppression segments' in row_text: break  

        waterfall_order = get_val(row, ['Waterfall Order', 'WaterfallOrder'])
        if waterfall_order.lower() == 'n/a': continue  

        segment_code = get_val(row, ['Segment Code', 'SegmentCode', 'Additional SegmentCode'])
        if not segment_code: continue

        dag_segment_name = get_val(row, ['DAG Segment Name', 'Segment Name', 'DAGSegmentName'])
        cell_name = get_val(row, ['CellName', 'Cell Name'])
        source = get_val(row, ['Source'])
        requested_volume = get_val(row, ['Requested Volume', 'Requested', 'Testing split', 'Split'])
        sline_code = get_val(row, ['SlineCode', 'Sline Code', 'SL Code'])

        if not dag_segment_name and not cell_name: continue

        check_str = (dag_segment_name or cell_name).lower()
        if check_str.startswith('where ') or check_str.startswith('pick from') or check_str.startswith('and '): continue

        if not dag_segment_name: dag_segment_name = cell_name
        if not cell_name: cell_name = dag_segment_name

        dag_segment_name = dag_segment_name.replace('&', 'AND').replace(' ', '_')
        cell_name = cell_name.replace('&', 'AND').replace(' ', '_')

        dag_count_col = None
        for name in ['dag count', 'count']:
            if name in cols:
                dag_count_col = cols[name]
                break
                
        dag_count_raw = row.get(dag_count_col, 0) if dag_count_col else 0
        try:
            if pd.isna(dag_count_raw): dag_count = 0.0
            elif isinstance(dag_count_raw, (int, float)): dag_count = float(dag_count_raw)
            else:
                cleaned = str(dag_count_raw).replace(',', '').replace('.', '').strip()
                dag_count = float(cleaned) if cleaned else 0.0
        except: dag_count = 0.0

        seg_code_lower = segment_code.lower()
        has_split_in_code = ("50%:" in seg_code_lower or "50%" in seg_code_lower or "(50%)" in seg_code_lower or "/" in segment_code or "\n" in segment_code)
        req_vol_str = str(requested_volume).lower().replace(',', '.')
        has_split_in_volume = ("50%" in req_vol_str or "50/50" in req_vol_str or req_vol_str == "0.5" or req_vol_str == "0.50" or req_vol_str == "50")

        is_split = has_split_in_code or has_split_in_volume

        if is_split:
            if "\n" in segment_code: parts = [x.strip() for x in segment_code.split("\n") if x.strip()]
            elif "/" in segment_code: parts = [x.strip() for x in segment_code.split("/", 1)]
            else: parts = [segment_code, segment_code]

            if "\n" in sline_code: sl_parts = [x.strip() for x in sline_code.split("\n") if x.strip()]
            elif "/" in sline_code: sl_parts = [x.strip() for x in sline_code.split("/", 1)]
            else: sl_parts = [sline_code, sline_code]

            if len(parts) >= 2:
                p0_lower = parts[0].lower()
                p1_lower = parts[1].lower()
                is_p1_control = 'control' in p1_lower or 'ctrl' in p1_lower
                is_p0_control = 'control' in p0_lower or 'ctrl' in p0_lower
                
                if is_p1_control and not is_p0_control:
                    parts[0], parts[1] = parts[1], parts[0]
                    if len(sl_parts) >= 2: sl_parts[0], sl_parts[1] = sl_parts[1], sl_parts[0]

            code1 = parts[0].replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()
            code2 = parts[1] if len(parts) > 1 else code1
            code2 = code2.replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()

            sl_code1 = sl_parts[0].replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()
            sl_code2 = sl_parts[1] if len(sl_parts) > 1 else sl_code1
            sl_code2 = sl_code2.replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()

            records = [
                {"WaterfallId": waterfall_id, "DAGSegmentName": dag_segment_name, "CellName": cell_name, "Split": 0.5, "SegmentCode": code1, "SlineCode": sl_code1},
                {"WaterfallId": waterfall_id + 1, "DAGSegmentName": dag_segment_name, "CellName": f"{cell_name}-test", "Split": 0.5, "SegmentCode": code2, "SlineCode": sl_code2}
            ]
            waterfall_id += 2

            for r in records:
                r["_Source"] = source  
                if source.upper() != "ACM":
                    r["DAGSCount"] = dag_count
                    r["isABTest"] = "FALSE"
                r.update({"extraColumnA": "", "extraColumnB": "", "extraColumnC": ""})
                final_json_data.append(r)
        else:
            code = segment_code.replace("(100%)", "").replace("100%:", "").replace("100%", "").strip()
            sl_code = sline_code.replace("(100%)", "").replace("100%:", "").replace("100%", "").strip()
            
            record = {"WaterfallId": waterfall_id, "DAGSegmentName": dag_segment_name, "CellName": cell_name, "Split": 1.0, "SegmentCode": code, "SlineCode": sl_code}
            waterfall_id += 1
            
            record["_Source"] = source 
            if source.upper() != "ACM":
                record["DAGSCount"] = dag_count
                record["isABTest"] = "FALSE"
            record.update({"extraColumnA": "", "extraColumnB": "", "extraColumnC": ""})
            final_json_data.append(record)

    return final_json_data

# =====================================================================
# FUNCTION: ACC SUBJECT GENERATOR
# =====================================================================
def generate_acc_subject_code(file_bytes, file_name):
    temp_subjects = []
    
    if file_name.endswith(('.xlsx', '.xls')):
        try:
            xls = pd.ExcelFile(io.BytesIO(file_bytes))
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                for idx, row in df.iterrows():
                    for i, val in enumerate(row.values):
                        if pd.notna(val):
                            val_str = str(val).strip()
                            sl_match = re.search(r'(SL[\s_-]?\d+)', val_str, re.IGNORECASE)
                            if sl_match:
                                sl_code = sl_match.group(1).strip()
                                text_after = val_str[sl_match.end():]
                                text_after = re.sub(r'^[\s:\-]+', '', text_after).strip()
                                subject_text = ""
                                
                                if text_after: subject_text = text_after
                                else:
                                    if i + 1 < len(row.values) and pd.notna(row.values[i+1]):
                                        right_val = str(row.values[i+1]).strip()
                                        if right_val: subject_text = right_val
                                    if not subject_text and idx + 1 < len(df) and pd.notna(df.iloc[idx+1, i]):
                                        bottom_val = str(df.iloc[idx+1, i]).strip()
                                        if bottom_val: subject_text = bottom_val
                                
                                if subject_text:
                                    subject_text = re.sub(r'\s+', ' ', subject_text)
                                    normalized_code = re.sub(r'^SL[\s_-]?', 'SL_', sl_code, flags=re.IGNORECASE).upper()
                                    temp_subjects.append({"originalCode": sl_code.replace(" ", ""), "normalizedCode": normalized_code, "text": subject_text})
        except Exception as e: return None, f"Error reading Excel: {e}"
    else:
        file_content = file_bytes.decode('utf-8')
        lines = file_content.splitlines()
        line_regex = re.compile(r'^\s*(SL[\s_-]?\d+)\s*;\s*(.*)', re.IGNORECASE)
        for line in lines:
            line = line.strip()
            if not line: continue
            match = line_regex.match(line)
            if match:
                original_code = match.group(1).strip()
                subject_text = match.group(2).strip()
                normalized_code = re.sub(r'^SL[\s_-]?', 'SL_', original_code, flags=re.IGNORECASE).upper()
                if subject_text:
                    temp_subjects.append({"originalCode": original_code.replace(" ", ""), "normalizedCode": normalized_code, "text": subject_text})
                    
    unique_map = {item["normalizedCode"]: item for item in temp_subjects}
    
    def sort_key(item):
        num_part = item["normalizedCode"].split('_')[1] if '_' in item["normalizedCode"] else '0'
        return int(num_part) if num_part.isdigit() else 0
        
    subject_lines = sorted(unique_map.values(), key=sort_key)
    
    if not subject_lines: return None, "No 'SL' code or Subject text found in the file."
        
    code = "<%\n"
    code += "var nameProper = recipient.firstName ? recipient.firstName.charAt(0).toUpperCase() + recipient.firstName.substring(1).toLowerCase() : \"\";\n"
    code += "var subject = \"\";\n"
    code += "var hasFirstName = recipient.firstName && recipient.firstName !== '' && recipient.firstName !== 'NONAME';\n"
    name_regex = re.compile(r'(\[Name\]\s*,?\s*)', re.IGNORECASE)
    
    for index, line in enumerate(subject_lines):
        original_text = line["text"]
        match = name_regex.search(original_text)
        has_first_name = bool(match)
        
        if has_first_name:
            prefix = original_text[:match.start()]
            text_after = original_text[match.end():]
        else:
            prefix = ""
            text_after = original_text
            
        escaped_prefix = prefix.replace('"', '\\"')
        escaped_text_after = text_after.replace('"', '\\"')
        
        if not has_first_name and escaped_text_after: escaped_text_after = escaped_text_after[0].upper() + escaped_text_after[1:]
            
        if has_first_name: subject_expr_if = f'"{escaped_prefix}" + nameProper + ", {escaped_text_after}"'
        else: subject_expr_if = f'"{escaped_prefix}{escaped_text_after}"'
            
        text_for_else = text_after.strip()
        if text_for_else: text_for_else = text_for_else[0].upper() + text_for_else[1:]
        subject_str_else = f'"{escaped_prefix}{text_for_else.replace("\"", "\\\"")}"'
        
        cond_prefix = "if" if index == 0 else "else if"
        code += f"\n  {cond_prefix} (targetData.slineCode == '{line['originalCode']}') {{\n"
        code += f"    if (hasFirstName) {{\n      subject = {subject_expr_if};\n    }} else {{\n      subject = {subject_str_else};\n    }}\n  }}"
        
    code += "\n%>\n<%= subject %>\n<%\nif( subject == '' || subject == undefined ) {\n"
    code += "  var attemptedCode = targetData.slineCode || \"UNKNOWN\";\n"
    code += "  delivery.cancelMessage( 'Empty Subject Line generated for slineCode: ' + attemptedCode );\n}\n%>"
    
    return subject_lines, code

# =====================================================================
# STREAMLIT INTERFACE - MAIN TABS
# =====================================================================
st.title("🛠️ Netbiis Automation Portal")
st.markdown("<p style='color: #8E8E93; font-size: 16px;'>Automate workflow structures, generate code, and compile package deliverables seamlessly.</p>", unsafe_allow_html=True)
st.divider()

tab1, tab2, tab3 = st.tabs(["⚙️ JSON Converter", "📦 ZIP Generator", "✉️ ACC Subject Generator"])

# ---------------------------------------------------------------------
# TAB 1: JSON GENERATOR (Waterfall)
# ---------------------------------------------------------------------
with tab1:
    
    if 'jira_dados' in st.session_state:
        jd = st.session_state['jira_dados']
        ticket_id = jd.get('ticket', 'N/A')
        lob_txt = jd.get('lob', 'N/A')
        status_txt = jd.get('status', 'N/A')
        data_deploy_texto = jd.get('data_deploy', 'N/A')
        segmentos_lista = jd.get('segmentos_ativos', [])
        segmentos_txt = ', '.join(segmentos_lista) if segmentos_lista else 'None'
        
        # Visually highlighted status banner with soft background
        st.markdown(f"""
        <div style="background-color: rgba(48, 209, 88, 0.1); border-left: 4px solid #30D158; padding: 15px; border-radius: 8px; margin-bottom: 25px;">
            <span style="color: #30D158; font-weight: 600;">🟢 Active Ticket:</span> {ticket_id} ({status_txt}) &nbsp;|&nbsp; 
            <span style="color: #E2E2E2; font-weight: 600;">LOB:</span> {lob_txt} &nbsp;|&nbsp; 
            <span style="color: #E2E2E2; font-weight: 600;">Deploy:</span> {data_deploy_texto} &nbsp;|&nbsp; 
            <span style="color: #E2E2E2; font-weight: 600;">Segments:</span> {segmentos_txt}
        </div>
        """, unsafe_allow_html=True)

    st.markdown("### Upload Matrix File")
    st.write("Upload the Excel matrix to audit, parse the waterfall, and generate the structural payload.")

    uploaded_file = st.file_uploader("Select the source file (.xlsx, .csv)", type=["xlsx", "csv"], key="json_uploader")

    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.getvalue()
            file_name = uploaded_file.name
            
            sheet_selection = None
            if file_name.endswith(('.xlsx', '.xls')):
                xls = pd.ExcelFile(io.BytesIO(file_bytes))
                sheet_names = xls.sheet_names
                if len(sheet_names) > 1:
                    sheet_selection = st.selectbox("📂 Multiple tabs detected. Select the waterfall tab:", sheet_names)
                else:
                    sheet_selection = sheet_names[0]

            meta = extract_metadata(file_bytes, file_name, sheet_selection)
            df = load_and_clean_data(file_bytes, file_name, sheet_selection)
            
            st.info("File loaded successfully. Ready for processing.")
            
            if st.button("🚀 Analyze and Generate Structure", type="primary"):
                json_objects = process_data(df)
                
                # --- METADATA IN CARDS ---
                st.markdown("<br><h3>📊 Campaign Metadata</h3>", unsafe_allow_html=True)
                colA, colB, colC = st.columns(3)
                with colA: st.metric("Campaign Name", meta.get('campaign_name', 'N/A'))
                with colB: st.metric("Expected Deployment", meta.get('deployment_date', 'N/A'))
                with colC: st.metric("Requester", meta.get('user', 'N/A'))
                    
                if meta.get("cid"): st.markdown(f"**Tracking CID:** `{meta.get('cid')}`")

                if meta.get("notes"):
                    st.markdown("#### 💡 Personalization Notes")
                    for note in meta.get("notes", []): st.info(note)

                # --- AUDIT ---
                st.markdown("<br><h3>🕵️‍♂️ Matrix Audit</h3>", unsafe_allow_html=True)
                
                errors = []
                warnings = []
                cell_names = [obj.get("CellName", "") for obj in json_objects]
                counts = Counter(cell_names)
                duplicates = [name for name, count in counts.items() if count > 1 and name != ""]
                
                if duplicates: errors.append(f"Duplicate Cell Names (Causes Adobe error): {', '.join(duplicates)}")
                    
                for obj in json_objects:
                    c_name = obj.get("CellName", "")
                    if not c_name: warnings.append(f"Empty CellName at WaterfallId {obj.get('WaterfallId')}")
                    elif len(c_name) > 64: errors.append(f"CellName exceeds 64 chars at WaterfallId {obj.get('WaterfallId')}: '{c_name}'")
                    if not obj.get("SlineCode"): warnings.append(f"Empty SlineCode at WaterfallId {obj.get('WaterfallId')}")
                        
                if errors:
                    st.error("🚨 Critical Errors Found. Please fix the spreadsheet before generating the link!")
                    for e in errors: st.write(f"- {e}")
                    st.stop() 
                    
                if warnings:
                    st.warning("⚠️ Fill Warnings (Check if intentional):")
                    for w in warnings: st.write(f"- {w}")
                        
                if not errors and not warnings:
                    st.success("✅ Flawless matrix! No structural anomalies found.")

                # --- WATERFALL OVERVIEW IN CARDS ---
                st.markdown("<br><h3>📈 Waterfall Overview</h3>", unsafe_allow_html=True)
                
                total_segments = len(json_objects)
                total_ab_tests = sum(1 for obj in json_objects if obj.get("Split") == 0.5) // 2
                total_volume = sum(float(obj.get("DAGSCount", 0)) for obj in json_objects)
                total_acm = sum(1 for obj in json_objects if str(obj.get("_Source", "")).upper() == "ACM")
                total_dag = total_segments - total_acm
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total Segments", total_segments)
                col2.metric("Native (ACM)", total_acm)
                col3.metric("External (DAG)", total_dag)
                col4.metric("Estimated Volume", f"{total_volume:,.0f}".replace(",", "."))
                
                if total_volume > 0:
                    chart_data = pd.DataFrame({"Segment": [obj.get("CellName") for obj in json_objects], "Volume": [float(obj.get("DAGSCount", 0)) for obj in json_objects]})
                    st.bar_chart(chart_data.set_index("Segment"))

                # --- TARGETING RULES ---
                st.markdown("<br><h3>🛑 Targeting Rules</h3>", unsafe_allow_html=True)
                
                # Displays the Audience in a wide highlighted box
                if meta.get("audience"):
                    st.info("**👥 AUDIENCE SELECTION:**\n" + "\n".join([f"- {item}" for item in meta["audience"]]))
                
                inc_col, exc_col = st.columns(2)
                
                with inc_col:
                    st.success("**🎯 ONLY INCLUDE (Global Filters):**")
                    if meta.get("include"):
                        for item in meta["include"]:
                            st.write(f"- {item}")
                    else:
                        st.write("None")
                        
                with exc_col:
                    st.error("**🚫 EXCLUDE (Suppressions):**")
                    if meta.get("exclude"):
                        for item in meta["exclude"]:
                            st.write(f"- {item}")
                    else:
                        st.write("None")

                for obj in json_objects: obj.pop("_Source", None)
                    
                json_string = json.dumps(json_objects, indent=2)
                compressed_data = zlib.compress(json_string.encode('utf-8'))
                encoded_data = base64.urlsafe_b64encode(compressed_data).decode('utf-8')
                shareable_link = f"https://upgraded-guide-bxpthdptstyfwaor2naznv.streamlit.app?data={encoded_data}"

                # --- EXPANDERS TO KEEP THE SCREEN CLEAN ---
                st.markdown("<br><h3>💻 Output Deliverables</h3>", unsafe_allow_html=True)
                
                with st.expander("🔗 Open Shareable Link (For Workflow Configurator)", expanded=True):
                    st.info("Copy the link below and paste it into the Jira ticket or spreadsheet. Anyone clicking it will see the JSON layout.")
                    st.code(shareable_link, language='text')
                
                with st.expander("📄 View & Download Raw JSON Output"):
                    st.code(json_string, language='json')
                    st.download_button("📥 Download .json File", data=json_string, file_name="dataSegments_output.json", mime="application/json", use_container_width=True)

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")

# ---------------------------------------------------------------------
# TAB 2: ZIP STRUCTURE GENERATOR & JIRA CONNECTION
# ---------------------------------------------------------------------
with tab2:
    st.markdown("### 📦 Generate Package & QA Folders")
    st.write("Fetch Jira metadata to automatically generate the complete folder hierarchy and populate QA spreadsheets.")
    
    with st.container(border=True):
        st.markdown("#### 🔌 1. Jira Connection")
        
        # Secure variables
        try:
            JIRA_EMAIL = st.secrets["JIRA_EMAIL"]
            JIRA_TOKEN = st.secrets["JIRA_TOKEN"]
            JIRA_URL = "https://netbiis.atlassian.net"
        except KeyError:
            JIRA_EMAIL = None
        
        col_j1, col_j2, col_j3 = st.columns([2, 1, 1])
        with col_j1:
            ticket_input_tab = st.text_input("JIRA Ticket ID (ex: SEA-7117)", key="jira_input_tab").strip().upper()
        with col_j2:
            st.write("")
            st.write("")
            fetch_btn = st.button("Fetch Data", use_container_width=True)
        with col_j3:
            st.write("")
            st.write("")
            if st.button("Clear Session", use_container_width=True):
                st.session_state.clear()
                st.rerun()
                
        if fetch_btn:
            if not JIRA_EMAIL:
                st.error("⚠️ Credentials not found. Please configure 'Secrets' on Streamlit Cloud.")
            elif ticket_input_tab:
                with st.spinner(f"Analyzing {ticket_input_tab}..."):
                    auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_TOKEN)
                    headers = {"Accept": "application/json"}
                    
                    url_main = f"{JIRA_URL}/rest/api/2/issue/{ticket_input_tab}"
                    resp_main = requests.get(url_main, headers=headers, auth=auth)
                    
                    if resp_main.status_code == 200:
                        dados_main = resp_main.json()
                        campos = dados_main.get("fields", {})
                        
                        titulo = campos.get("summary", "")
                        status_ticket = campos.get("status", {}).get("name", "Unknown")
                        lob_campo = campos.get("customfield_10476", [""])[0] if campos.get("customfield_10476") else ""
                        if lob_campo: lob = str(lob_campo)
                        else: lob = re.sub(r'[^A-Za-z0-9]', '', titulo.split(" ")[0] if titulo else "N/A")
                        
                        ano_atual = datetime.datetime.now().year
                        match_data = re.findall(r'(\d{1,2})/(\d{1,2})', titulo)
                        if match_data:
                            mes, dia = match_data[-1]
                            data_deploy = f"{dia.zfill(2)}/{mes.zfill(2)}/{ano_atual}"
                        else: data_deploy = datetime.datetime.now().strftime("%m/%d/%Y")
                        
                        ## --- Capture the main ticket owner as fallback ---
                        assignee_main = campos.get("assignee")
                        qa_reviewer_name = assignee_main.get("displayName", "N/A") if assignee_main else "N/A"
                        
                        subtasks = campos.get("subtasks", [])
                        segmentos_encontrados = set()
                        
                        for sub in subtasks:
                            sub_key = sub.get("key")
                            # Extracts the full sub-task via API
                            url_sub = f"{JIRA_URL}/rest/api/2/issue/{sub_key}"
                            resp_sub = requests.get(url_sub, headers=headers, auth=auth)
                            
                            if resp_sub.status_code == 200:
                                sub_data = resp_sub.json()
                                sub_fields = sub_data.get("fields", {})
                                
                                # 1. Segments Extraction
                                comentarios = sub_fields.get("comment", {}).get("comments", [])
                                for comment in comentarios:
                                    linhas = comment.get("body", "").split('\n')
                                    for linha in linhas:
                                        match = re.search(r'_WK\d+_([A-Za-z0-9_]+)', linha.strip())
                                        if match:
                                            seg = match.group(1).strip().upper()
                                            if seg: segmentos_encontrados.add(re.sub(r'_V\d+$', '', seg))
                                            
                                # 2. Looks for Specific Assignee ONLY for QA (Ignores HTML Proof)
                                sub_summary = sub_fields.get("summary", "").upper()
                                if "1A QA" in sub_summary or "QA REVIEW" in sub_summary:
                                    sub_assignee = sub_fields.get("assignee")
                                    if sub_assignee: # Only replaces if someone is assigned to 1A
                                        qa_reviewer_name = sub_assignee.get("displayName", qa_reviewer_name)
                        
                        st.session_state['jira_dados'] = {
                            "ticket": ticket_input_tab, "lob": str(lob), "titulo": str(titulo),
                            "status": status_ticket, "segmentos_originais": sorted(list(segmentos_encontrados)),
                            "segmentos_ativos": sorted(list(segmentos_encontrados)), "data_deploy": data_deploy,
                            "reviewer": qa_reviewer_name
                        }
                        st.toast("✅ Metadata synchronized successfully!", icon="✅")
                    else: st.error(f"Error {resp_main.status_code}: Ticket not found.")
            else:
                st.warning("Enter a valid ID.")

    if 'jira_dados' in st.session_state:
        jd = st.session_state['jira_dados']
        
        st.markdown("<br><h4>✏️ 2. Job Fine-Tuning</h4>", unsafe_allow_html=True)
        # NEW: 3 columns to accommodate the Reviewer
        colA, colB, colC = st.columns(3)
        with colA: jd['status'] = st.text_input("Ticket Status:", jd.get('status', ''))
        with colB: jd['data_deploy'] = st.text_input("Deployment Date:", jd.get('data_deploy', ''))
        with colC: jd['reviewer'] = st.text_input("Reviewer:", jd.get('reviewer', ''))
        
        selecionados = st.multiselect(
            "Segments (Remove unwanted ones):",
            options=jd.get('segmentos_originais', []),
            default=jd.get('segmentos_ativos', [])
        )
        st.session_state['jira_dados']['segmentos_ativos'] = selecionados

        st.markdown("<br><h4>📦 3. Build Architecture</h4>", unsafe_allow_html=True)
        if st.button("📥 Generate ZIP Package", type="primary"):
            job_id = jd['ticket']
            lob = jd['lob']
            segmentos = jd.get('segmentos_ativos', ["N_A"]) if jd.get('segmentos_ativos') else ["N_A"]
            data_deploy = jd.get('data_deploy', datetime.datetime.now().strftime("%m/%d/%Y"))
            # NEW: Pulls dynamic reviewer
            reviewer = jd.get('reviewer', 'N/A')
            
            zip_buffer = io.BytesIO()
            my_bar = st.progress(10, text="Initializing folder structuring...")
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                qa_geral_path = f"{job_id}/04_QA_Geral"
                my_bar.progress(30, text="Reading template inputs...")
                template_csv_path = "_Templates/Input de Subjects - Modelo.csv"
                csv_content = b""
                if os.path.exists(template_csv_path):
                    with open(template_csv_path, "rb") as f: csv_content = f.read()

                my_bar.progress(45, text="Creating hierarchical folders...")
                for seg in segmentos:
                    seg_folder = f"{job_id}/{seg}" if seg != "N_A" else job_id
                    zip_file.writestr(f"{seg_folder}/01_Inputs/", "")
                    zip_file.writestr(f"{seg_folder}/02_HTML/", "")
                    zip_file.writestr(f"{seg_folder}/03_LP/", "")
                    if csv_content:
                        seg_name = seg if seg != "N_A" else "General"
                        zip_file.writestr(f"{seg_folder}/01_Inputs/Input de Subjects - {seg_name}.csv", csv_content)
                
                template_rev_path = "_Templates/Revisao - Modelo.xlsx"
                template_lc_path = "_Templates/Last Check - Modelo.xlsx"
                
                my_bar.progress(60, text="Cloning tabs & processing Review QA...")
                if os.path.exists(template_rev_path):
                    wb_rev = load_workbook(template_rev_path)
                    ws_temp_rev = wb_rev.active
                    logos_rev = []
                    for img in ws_temp_rev._images:
                        try:
                            if hasattr(img.ref, 'getvalue'): img_bytes = img.ref.getvalue()
                            elif hasattr(img.ref, 'read'): img.ref.seek(0); img_bytes = img.ref.read()
                            else: img_bytes = img._data()
                            logos_rev.append((img_bytes, copy.copy(img.anchor)))
                        except: pass
                    
                    for sheet_name in wb_rev.sheetnames:
                        if sheet_name != ws_temp_rev.title: del wb_rev[sheet_name]
                    
                    for i, seg in enumerate(segmentos):
                        seg_name = "N/A" if seg == "N_A" else seg
                        nome_aba = re.sub(r'[/\\?*\[\]]', '', job_id if seg_name == "N/A" else f"{job_id} - {seg_name}")[:31]
                        
                        if i == 0: ws_target = ws_temp_rev
                        else:
                            ws_target = wb_rev.copy_worksheet(ws_temp_rev)
                            for img_bytes, anchor in logos_rev:
                                try:
                                    nova_img = xlImage(io.BytesIO(img_bytes)); nova_img.anchor = copy.copy(anchor)
                                    ws_target.add_image(nova_img)
                                except: pass
                            
                        ws_target.title = nome_aba
                        # NEW: Clean Excel injection
                        ws_target["F2"] = f"{job_id}"; ws_target["F3"] = f"{seg_name}"; ws_target["F4"] = reviewer; ws_target["F6"] = datetime.datetime.now().strftime("%m/%d/%Y")
                    
                    rev_stream = io.BytesIO(); wb_rev.save(rev_stream)
                    zip_file.writestr(f"{qa_geral_path}/Review Jobs Samsung - {job_id}.xlsx", rev_stream.getvalue())

                my_bar.progress(85, text="Processing Last Check QA matrices...")
                if os.path.exists(template_lc_path):
                    wb_lc = load_workbook(template_lc_path)
                    ws_temp_lc = wb_lc.active
                    logos_lc = []
                    for img in ws_temp_lc._images:
                        try:
                            if hasattr(img.ref, 'getvalue'): img_bytes = img.ref.getvalue()
                            elif hasattr(img.ref, 'read'): img.ref.seek(0); img_bytes = img.ref.read()
                            else: img_bytes = img._data()
                            logos_lc.append((img_bytes, copy.copy(img.anchor)))
                        except: pass
                    
                    for sheet_name in wb_lc.sheetnames:
                        if sheet_name != ws_temp_lc.title: del wb_lc[sheet_name]
                            
                    for i, enumerate_seg in enumerate(segmentos):
                        seg_name = "N/A" if enumerate_seg == "N_A" else enumerate_seg
                        nome_aba = re.sub(r'[/\\?*\[\]]', '', job_id if seg_name == "N/A" else f"{job_id} - {seg_name}")[:31]
                        
                        if i == 0: ws_target = ws_temp_lc
                        else:
                            ws_target = wb_lc.copy_worksheet(ws_temp_lc)
                            for img_bytes, anchor in logos_lc:
                                try:
                                    nova_img = xlImage(io.BytesIO(img_bytes)); nova_img.anchor = copy.copy(anchor)
                                    ws_target.add_image(nova_img)
                                except: pass
                            
                        ws_target.title = nome_aba
                        # NEW: Clean Excel injection
                        ws_target["D1"] = f"Reviewed by: {reviewer}"; ws_target["D2"] = f"LOB:{lob}"; ws_target["D3"] = f"JIRA Ticket: {job_id}"
                        ws_target["D4"] = f"Segment: {seg_name}"; ws_target["D5"] = data_deploy
                        
                    lc_stream = io.BytesIO(); wb_lc.save(lc_stream)
                    zip_file.writestr(f"{qa_geral_path}/Last Check - {job_id}.xlsx", lc_stream.getvalue())

            my_bar.progress(100, text="Finalizing...")
            st.success("🎉 ZIP structural package compiled successfully!")
            
            st.download_button("📦 Download Local ZIP", data=zip_buffer.getvalue(), file_name=f"{job_id}_Architecture.zip", mime="application/zip", use_container_width=True)

# ---------------------------------------------------------------------
# TAB 3: SUBJECT GENERATOR (ACC)
# ---------------------------------------------------------------------
with tab3:
    st.markdown("### ✉️ ACC Intelligent Subject Extractor")
    st.write("Upload a **URL Matrix (.xlsx)** or a CSV/TXT file. The engine will scan for SL codes and build the E4X personalization logic.")
    
    uploaded_sl_file = st.file_uploader("Select Matrix or Text File", type=["xlsx", "xls", "csv", "txt"], key="sl_uploader")
    
    if uploaded_sl_file is not None:
        file_bytes = uploaded_sl_file.getvalue()
        file_name = uploaded_sl_file.name
        
        with st.spinner("Scanning and extracting subjects..."):
            parsed_data, generated_code = generate_acc_subject_code(file_bytes, file_name)
        
        if parsed_data:
            st.success("✅ File parsed successfully!")
            
            col1, col2 = st.columns([1, 1])
            with col1:
                st.markdown("#### 👀 Discovered Subjects")
                df_preview = pd.DataFrame(parsed_data)[["originalCode", "text"]]
                df_preview.columns = ["SL Code", "Extracted Subject Text"]
                st.dataframe(df_preview, use_container_width=True, hide_index=True)
            
            with col2:
                st.markdown("#### 💻 E4X Logic Output")
                with st.container(border=True):
                    st.code(generated_code, language='javascript')
                
                st.download_button(
                    label="📥 Download .txt Script",
                    data=generated_code,
                    file_name=f"{file_name.split('.')[0]}_E4X_script.txt",
                    mime="text/plain",
                    use_container_width=True,
                    type="primary"
                )
        else:
            st.error(generated_code)
