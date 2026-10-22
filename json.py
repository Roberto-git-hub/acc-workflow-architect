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
# INTERFACE STREAMLIT - CONFIGURAÇÃO
# =====================================================================
st.set_page_config(page_title="Gerador JSON - Adobe Campaign", layout="centered", page_icon="⚙️")

# =====================================================================
# MODO LEITOR DE LINK (Operador)
# =====================================================================
if 'data' in st.query_params:
    try:
        compressed_data = base64.urlsafe_b64decode(st.query_params['data'])
        json_string = zlib.decompress(compressed_data).decode('utf-8')
        
        st.title("📄 Visualizador de JSON (Workflow)")
        st.success("JSON carregado com sucesso via link!")
        st.warning("💡 COMO COPIAR: Passe o mouse no canto superior direito do bloco preto abaixo e clique no ícone de prancheta (📋) para copiar tudo.")
        
        st.code(json_string, language='json')
        st.stop() 
    except Exception as e:
        st.error("⚠️ O link parece estar quebrado ou corrompido. Peça para gerar novamente.")
        st.stop()

# =====================================================================
# MÓDULO 1: INTEGRAÇÃO JIRA API & UI DE LIMPEZA
# =====================================================================
st.sidebar.header("🔌 Conexão Jira")

if st.sidebar.button("🔄 Limpar Sessão (Novo Job)"):
    st.session_state.clear()
    st.rerun()

ticket_input = st.sidebar.text_input("JIRA Ticket ID (ex: SEA-7117)").strip().upper()

# 🔒 CREDENCIAIS PROTEGIDAS (Via Streamlit Secrets)
JIRA_URL = "https://netbiis.atlassian.net"
try:
    JIRA_EMAIL = st.secrets["JIRA_EMAIL"]
    JIRA_TOKEN = st.secrets["JIRA_TOKEN"]
except KeyError:
    st.sidebar.error("⚠️ Credenciais não encontradas. Configure o 'Secrets' no Streamlit Cloud.")
    st.stop()

if st.sidebar.button("Puxar Dados do Ticket"):
    if ticket_input:
        with st.spinner(f"Analisando estrutura do {ticket_input}..."):
            auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_TOKEN)
            headers = {"Accept": "application/json"}
            
            url_main = f"{JIRA_URL}/rest/api/2/issue/{ticket_input}"
            resp_main = requests.get(url_main, headers=headers, auth=auth)
            
            if resp_main.status_code == 200:
                dados_main = resp_main.json()
                campos = dados_main.get("fields", {})
                
                titulo = campos.get("summary", "")
                lob_campo = campos.get("customfield_10476", [""])[0] if campos.get("customfield_10476") else ""
                
                if lob_campo:
                    lob = str(lob_campo)
                else:
                    lob_pescado = titulo.split(" ")[0] if titulo else "N/A"
                    lob = re.sub(r'[^A-Za-z0-9]', '', lob_pescado)
                
                ano_atual = datetime.datetime.now().year
                data_deploy = ""
                match_data = re.search(r'\((\d{1,2})/(\d{1,2})\)', titulo)
                if match_data:
                    mes, dia = match_data.groups()
                    data_deploy = f"{dia.zfill(2)}/{mes.zfill(2)}/{ano_atual}"
                else:
                    data_deploy = datetime.datetime.now().strftime("%d/%m/%Y")
                
                subtasks = campos.get("subtasks", [])
                segmentos_encontrados = set()
                
                for sub in subtasks:
                    sub_key = sub.get("key")
                    url_comments = f"{JIRA_URL}/rest/api/2/issue/{sub_key}/comment"
                    resp_comments = requests.get(url_comments, headers=headers, auth=auth)
                    
                    if resp_comments.status_code == 200:
                        comentarios = resp_comments.json().get("comments", [])
                        for comment in comentarios:
                            texto = comment.get("body", "")
                            linhas = texto.split('\n')
                            for linha in linhas:
                                match = re.search(r'_WK\d+_([A-Za-z0-9_]+)', linha.strip())
                                if match:
                                    seg = match.group(1).strip()
                                    if seg:
                                        seg = re.sub(r'(?i)_v\d+$', '', seg)
                                        segmentos_encontrados.add(seg)
                
                st.session_state['jira_dados'] = {
                    "ticket": ticket_input,
                    "lob": str(lob),
                    "titulo": str(titulo),
                    "segmentos": sorted(list(segmentos_encontrados)),
                    "data_deploy": data_deploy
                }
                st.sidebar.success("✅ Metadados sincronizados!")
            else:
                st.sidebar.error(f"Erro {resp_main.status_code}: Ticket não encontrado ou acesso negado.")
    else:
        st.sidebar.warning("Digite um Ticket ID válido.")

# =====================================================================
# FUNÇÕES DE PROCESSAMENTO (Coordenadora)
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
        "campaign_name": "",
        "deployment_date": "",
        "user": "",
        "cid": "",
        "notes": [],
        "include": [], 
        "exclude": []
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
                if str(val).strip().lower() == 'notes':
                    notes_col_idx = i
        elif "deployment date" in col0:
            meta["deployment_date"] = col1
        elif col0 == "user":
            meta["user"] = col1
        elif col0 == "cid":
            meta["cid"] = re.sub(r'(?i)_segment$', '', col1).strip()
        elif "only include" in col0:
            current_mode = "include"
            vals = [str(x).strip() for x in row.values[1:] if pd.notna(x) and str(x).strip().lower() != 'nan']
            if vals: meta["include"].append(" | ".join(vals))
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
                if note_val not in meta["notes"]:
                    meta["notes"].append(note_val)
    return meta

@st.cache_data
def load_and_clean_data(file_bytes, file_name, sheet_name=None):
    if file_name.endswith('.csv'):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python')
        except:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=';')
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
        if 'reference for combination' in row_text or 'suppression segments' in row_text:
            break  

        waterfall_order = get_val(row, ['Waterfall Order', 'WaterfallOrder'])
        if waterfall_order.lower() == 'n/a':
            continue  

        segment_code = get_val(row, ['Segment Code', 'SegmentCode', 'Additional SegmentCode'])
        if not segment_code:
            continue

        dag_segment_name = get_val(row, ['DAG Segment Name', 'Segment Name', 'DAGSegmentName'])
        cell_name = get_val(row, ['CellName', 'Cell Name'])
        source = get_val(row, ['Source'])
        requested_volume = get_val(row, ['Requested Volume', 'Requested', 'Testing split', 'Split'])
        sline_code = get_val(row, ['SlineCode', 'Sline Code', 'SL Code'])

        if not dag_segment_name and not cell_name:
            continue

        check_str = (dag_segment_name or cell_name).lower()
        if check_str.startswith('where ') or check_str.startswith('pick from') or check_str.startswith('and '):
            continue

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
            if pd.isna(dag_count_raw):
                dag_count = 0.0
            elif isinstance(dag_count_raw, (int, float)):
                dag_count = float(dag_count_raw)
            else:
                cleaned = str(dag_count_raw).replace(',', '').replace('.', '').strip()
                dag_count = float(cleaned) if cleaned else 0.0
        except:
            dag_count = 0.0

        seg_code_lower = segment_code.lower()
        has_split_in_code = (
            "50%:" in seg_code_lower or "50%" in seg_code_lower or "(50%)" in seg_code_lower or
            "/" in segment_code or "\n" in segment_code
        )

        req_vol_str = str(requested_volume).lower().replace(',', '.')
        has_split_in_volume = (
            "50%" in req_vol_str or "50/50" in req_vol_str or req_vol_str == "0.5" or 
            req_vol_str == "0.50" or req_vol_str == "50"
        )

        is_split = has_split_in_code or has_split_in_volume

        if is_split:
            if "\n" in segment_code:
                parts = [x.strip() for x in segment_code.split("\n") if x.strip()]
            elif "/" in segment_code:
                parts = [x.strip() for x in segment_code.split("/", 1)]
            else:
                parts = [segment_code, segment_code]

            if "\n" in sline_code:
                sl_parts = [x.strip() for x in sline_code.split("\n") if x.strip()]
            elif "/" in sline_code:
                sl_parts = [x.strip() for x in sline_code.split("/", 1)]
            else:
                sl_parts = [sline_code, sline_code]

            if len(parts) >= 2:
                p0_lower = parts[0].lower()
                p1_lower = parts[1].lower()
                is_p1_control = 'control' in p1_lower or 'ctrl' in p1_lower
                is_p0_control = 'control' in p0_lower or 'ctrl' in p0_lower
                
                if is_p1_control and not is_p0_control:
                    parts[0], parts[1] = parts[1], parts[0]
                    if len(sl_parts) >= 2:
                        sl_parts[0], sl_parts[1] = sl_parts[1], sl_parts[0]

            code1 = parts[0].replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()
            code2 = parts[1] if len(parts) > 1 else code1
            code2 = code2.replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()

            sl_code1 = sl_parts[0].replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()
            sl_code2 = sl_parts[1] if len(sl_parts) > 1 else sl_code1
            sl_code2 = sl_code2.replace("(50%)", "").replace("50%:", "").replace("50%", "").strip()

            records = [
                {
                    "WaterfallId": waterfall_id,
                    "DAGSegmentName": dag_segment_name,
                    "CellName": cell_name,
                    "Split": 0.5,
                    "SegmentCode": code1,
                    "SlineCode": sl_code1
                },
                {
                    "WaterfallId": waterfall_id + 1,
                    "DAGSegmentName": dag_segment_name,
                    "CellName": f"{cell_name}-test",
                    "Split": 0.5,
                    "SegmentCode": code2,
                    "SlineCode": sl_code2
                }
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
            
            record = {
                "WaterfallId": waterfall_id,
                "DAGSegmentName": dag_segment_name,
                "CellName": cell_name,
                "Split": 1.0,
                "SegmentCode": code,
                "SlineCode": sl_code
            }
            waterfall_id += 1
            
            record["_Source"] = source 
            if source.upper() != "ACM":
                record["DAGSCount"] = dag_count
                record["isABTest"] = "FALSE"
            record.update({"extraColumnA": "", "extraColumnB": "", "extraColumnC": ""})
            final_json_data.append(record)

    return final_json_data

# =====================================================================
# FUNÇÃO: GERADOR DE SUBJECTS ACC
# =====================================================================
def generate_acc_subject_code(file_content):
    lines = file_content.splitlines()
    line_regex = re.compile(r'^\s*(SL[_-]?\d+)\s*;\s*(.*)', re.IGNORECASE)
    
    temp_subjects = []
    for line in lines:
        line = line.strip()
        if not line: continue
        match = line_regex.match(line)
        if match:
            original_code = match.group(1).strip()
            subject_text = match.group(2).strip()
            normalized_code = re.sub(r'^SL[_-]?', 'SL_', original_code, flags=re.IGNORECASE).upper()
            if subject_text:
                temp_subjects.append({
                    "originalCode": original_code,
                    "normalizedCode": normalized_code,
                    "text": subject_text
                })
                
    unique_map = {item["normalizedCode"]: item for item in temp_subjects}
    
    def sort_key(item):
        num_part = item["normalizedCode"].split('_')[1] if '_' in item["normalizedCode"] else '0'
        return int(num_part) if num_part.isdigit() else 0
        
    subject_lines = sorted(unique_map.values(), key=sort_key)
    
    if not subject_lines:
        return None, "Nenhuma linha válida encontrada no formato 'SLxx; Texto...'."
        
    code = "<%\n"
    code += "var nameProper = recipient.firstName ? recipient.firstName.charAt(0).toUpperCase() + recipient.firstName.substring(1).toLowerCase() : \"\";\n"
    code += "var subject = \"\";\n"
    code += "var hasFirstName = recipient.firstName && recipient.firstName !== '' && recipient.firstName !== 'NONAME';\n"
    
    name_regex = re.compile(r'([Name]\s*,?\s*)', re.IGNORECASE)
    
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
        
        if not has_first_name and escaped_text_after:
            escaped_text_after = escaped_text_after[0].upper() + escaped_text_after[1:]
            
        if has_first_name:
            subject_expr_if = f'"{escaped_prefix}" + nameProper + ", {escaped_text_after}"'
        else:
            subject_expr_if = f'"{escaped_prefix}{escaped_text_after}"'
            
        text_for_else = text_after.strip()
        if text_for_else:
            text_for_else = text_for_else[0].upper() + text_for_else[1:]
        subject_str_else = f'"{escaped_prefix}{text_for_else.replace("\"", "\\\"")}"'
        
        cond_prefix = "if" if index == 0 else "else if"
        
        code += f"\n  {cond_prefix} (targetData.slineCode == '{line['originalCode']}') {{\n"
        code += f"    if (hasFirstName) {{\n"
        code += f"      subject = {subject_expr_if};\n"
        code += f"    }} else {{\n"
        code += f"      subject = {subject_str_else};\n"
        code += f"    }}\n"
        code += f"  }}"
        
    code += "\n%>\n<%= subject %>\n<%\n"
    code += "if( subject == '' || subject == undefined ) {\n"
    code += "  var attemptedCode = targetData.slineCode || \"UNKNOWN\";\n"
    code += "  delivery.cancelMessage( 'Empty Subject Line generated for slineCode: ' + attemptedCode );\n"
    code += "}\n%>"
    
    return subject_lines, code

# =====================================================================
# INTERFACE STREAMLIT - ABAS (TABS)
# =====================================================================
st.title("🛠️ Portal de Automação Netbiis")

tab1, tab2 = st.tabs(["⚙️ Waterfall & QA (JSON/ZIP)", "✉️ ACC Subject Generator"])

# ---------------------------------------------------------------------
# ABA 1: GERADOR DE JSON E ZIP
# ---------------------------------------------------------------------
with tab1:
    st.header("⚙️ Segment-to-JSON Converter")
    
    if 'jira_dados' in st.session_state:
        jd = st.session_state['jira_dados']
        data_deploy_texto = jd.get('data_deploy', 'N/A')
        st.info(f"**Ticket Ativo:** {jd['ticket']} | **LOB:** {jd['lob']} | **Deploy:** {data_deploy_texto} | **Segmentos:** {', '.join(jd['segmentos']) if jd['segmentos'] else 'Nenhum'}")

    st.write("Faça o upload da planilha para auditar, gerar a estrutura JSON e o Link Direto.")

    uploaded_file = st.file_uploader("Selecione o arquivo fonte", type=["xlsx", "csv"], key="json_uploader")

    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.getvalue()
            file_name = uploaded_file.name
            
            sheet_selection = None
            if file_name.endswith(('.xlsx', '.xls')):
                xls = pd.ExcelFile(io.BytesIO(file_bytes))
                sheet_names = xls.sheet_names
                if len(sheet_names) > 1:
                    sheet_selection = st.selectbox("📂 Múltiplas abas detectadas. Selecione a aba da cachoeira:", sheet_names)
                else:
                    sheet_selection = sheet_names[0]

            meta = extract_metadata(file_bytes, file_name, sheet_selection)
            df = load_and_clean_data(file_bytes, file_name, sheet_selection)
            
            st.success("Arquivo e aba carregados com sucesso!")
            
            if st.button("Analisar e Gerar JSON"):
                json_objects = process_data(df)
                
                st.divider()
                st.subheader("📋 1. Resumo da Campanha (Metadata)")
                colA, colB, colC = st.columns(3)
                
                with colA:
                    st.caption("Campanha")
                    st.write(f"**{meta.get('campaign_name', 'N/A')}**")
                with colB:
                    st.caption("Data Prevista (Deployment)")
                    st.write(f"**{meta.get('deployment_date', 'N/A')}**")
                with colC:
                    st.caption("Solicitante (User)")
                    st.write(f"**{meta.get('user', 'N/A')}**")
                    
                if meta.get("cid"):
                    st.markdown(f"**CID (Rastreio):** `{meta.get('cid')}`")

                if meta.get("notes"):
                    st.subheader("⚠️ 2. Alertas de Personalização (Notes)")
                    for note in meta.get("notes", []):
                        st.info(note, icon="💡")

                st.divider()
                st.subheader("🕵️‍♂️ Auditoria da Planilha")
                
                errors = []
                warnings = []
                
                cell_names = [obj.get("CellName", "") for obj in json_objects]
                counts = Counter(cell_names)
                duplicates = [name for name, count in counts.items() if count > 1 and name != ""]
                
                if duplicates:
                    errors.append(f"Nomes de Célula (CellName) Duplicados (Causa erro no Adobe): {', '.join(duplicates)}")
                    
                for obj in json_objects:
                    c_name = obj.get("CellName", "")
                    
                    if not c_name:
                        warnings.append(f"CellName vazio no WaterfallId {obj.get('WaterfallId')}")
                    elif len(c_name) > 64:
                        errors.append(f"CellName excede 64 caracteres (tem {len(c_name)}) no WaterfallId {obj.get('WaterfallId')}: '{c_name}'")
                        
                    if not obj.get("SlineCode"):
                        warnings.append(f"SlineCode vazio no WaterfallId {obj.get('WaterfallId')}")
                        
                if errors:
                    st.error("🚨 Erros Críticos Encontrados. Corrija a planilha antes de gerar o link!")
                    for e in errors:
                        st.write(f"- {e}")
                    st.stop() 
                    
                if warnings:
                    st.warning("⚠️ Avisos de Preenchimento (Verifique se é proposital):")
                    for w in warnings:
                        st.write(f"- {w}")
                        
                if not errors and not warnings:
                    st.success("✅ Planilha impecável! Nenhuma anomalia estrutural encontrada.")

                st.divider()
                st.subheader("📊 3. Visão Geral da Waterfall")
                
                total_segments = len(json_objects)
                total_ab_tests = sum(1 for obj in json_objects if obj.get("Split") == 0.5) // 2
                total_volume = sum(float(obj.get("DAGSCount", 0)) for obj in json_objects)
                
                total_acm = sum(1 for obj in json_objects if str(obj.get("_Source", "")).upper() == "ACM")
                total_dag = total_segments - total_acm
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total de Segmentos", total_segments)
                col2.metric("Nativos (ACM)", total_acm)
                col3.metric("Externos (DAG)", total_dag)
                col4.metric("Volume Estimado", f"{total_volume:,.0f}".replace(",", "."))
                
                if total_volume > 0:
                    st.write("**Distribuição Estimada por Segmento:**")
                    chart_data = pd.DataFrame({
                        "Segmento": [obj.get("CellName") for obj in json_objects],
                        "Volume": [float(obj.get("DAGSCount", 0)) for obj in json_objects]
                    })
                    st.bar_chart(chart_data.set_index("Segmento"))

                st.divider()
                st.subheader("✅ 4. Checklist de Supressão e Inclusão")
                inc_col, exc_col = st.columns(2)
                
                with inc_col:
                    st.success("**🎯 ONLY INCLUDE (Filtros Globais):**")
                    if meta.get("include"):
                        for item in meta["include"]:
                            st.write(f"- {item}")
                    else:
                        st.write("N/A")
                        
                with exc_col:
                    st.error("**🚫 EXCLUDE (Supressões):**")
                    if meta.get("exclude"):
                        for item in meta["exclude"]:
                            st.write(f"- {item}")
                    else:
                        st.write("N/A")

                for obj in json_objects:
                    obj.pop("_Source", None)
                    
                json_string = json.dumps(json_objects, indent=2)

                APP_URL = "https://upgraded-guide-bxpthdptstyfwaor2naznv.streamlit.app"
                compressed_data = zlib.compress(json_string.encode('utf-8'))
                encoded_data = base64.urlsafe_b64encode(compressed_data).decode('utf-8')
                shareable_link = f"{APP_URL}?data={encoded_data}"

                st.divider()
                st.subheader("🔗 Seu Link de Compartilhamento")
                st.info("Copie o link abaixo e cole na planilha. Quem clicar verá diretamente o JSON.")
                st.code(shareable_link, language='text')
                
                st.divider()
                st.subheader("📄 Resultado (JSON Bruto):")
                st.code(json_string, language='json')
                
                st.download_button(
                    label="📥 Baixar Arquivo .json",
                    data=json_string,
                    file_name="dataSegments_output.json",
                    mime="application/json",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo: {e}")

    st.divider()
    st.subheader("📦 5. Gerar Estrutura de Pastas e QA (ZIP)")

    if 'jira_dados' in st.session_state:
        st.info("Utilizando os dados sincronizados do Jira para criar as pastas e as planilhas de QA.")
        
        if st.button("📥 Gerar e Baixar Pacote ZIP do Job", type="primary"):
            jd = st.session_state['jira_dados']
            job_id = jd['ticket']
            lob = jd['lob']
            segmentos = jd['segmentos'] if jd['segmentos'] else ["N_A"]
            
            data_hoje = datetime.datetime.now().strftime("%d/%m/%Y")
            data_deploy = jd.get('data_deploy', data_hoje)
            reviewer = "Roberto"
            
            zip_buffer = io.BytesIO()
            progress_text = "Iniciando estruturação das pastas..."
            my_bar = st.progress(10, text=progress_text)
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                qa_geral_path = f"{job_id}/04_QA_Geral"
                my_bar.progress(30, text="Lendo arquivo CSV modelo (Inputs)...")
                template_csv_path = "_Templates/Input de Subjects - Modelo.csv"
                csv_content = b""
                if os.path.exists(template_csv_path):
                    with open(template_csv_path, "rb") as f:
                        csv_content = f.read()

                my_bar.progress(45, text="Criando as pastas de cada segmento...")
                for seg in segmentos:
                    seg_folder = f"{job_id}/{seg}" if seg != "N_A" else job_id
                    zip_file.writestr(f"{seg_folder}/01_Inputs/", "")
                    zip_file.writestr(f"{seg_folder}/02_HTML/", "")
                    zip_file.writestr(f"{seg_folder}/03_LP/", "")
                    
                    if csv_content:
                        seg_name = seg if seg != "N_A" else "Geral"
                        nome_csv = f"Input de Subjects - {seg_name}.csv"
                        zip_file.writestr(f"{seg_folder}/01_Inputs/{nome_csv}", csv_content)
                
                template_rev_path = "_Templates/Revisao - Modelo.xlsx"
                template_lc_path = "_Templates/Last Check - Modelo.xlsx"
                
                my_bar.progress(60, text="Clonando abas e injetando Logos (Revisão)...")
                if os.path.exists(template_rev_path):
                    wb_rev = load_workbook(template_rev_path)
                    ws_temp_rev = wb_rev.active
                    
                    logos_rev = []
                    for img in ws_temp_rev._images:
                        try:
                            if hasattr(img.ref, 'getvalue'):
                                img_bytes = img.ref.getvalue()
                            elif hasattr(img.ref, 'read'):
                                img.ref.seek(0)
                                img_bytes = img.ref.read()
                            else:
                                img_bytes = img._data()
                            logos_rev.append((img_bytes, copy.copy(img.anchor)))
                        except:
                            pass
                    
                    for sheet_name in wb_rev.sheetnames:
                        if sheet_name != ws_temp_rev.title:
                            del wb_rev[sheet_name]
                    
                    for i, seg in enumerate(segmentos):
                        seg_name = "N/A" if seg == "N_A" else seg
                        raw_aba_name = job_id if seg_name == "N/A" else f"{job_id} - {seg_name}"
                        nome_aba = re.sub(r'[/\\?*\[\]]', '', raw_aba_name)[:31]
                        
                        if i == 0:
                            ws_target = ws_temp_rev
                        else:
                            ws_target = wb_rev.copy_worksheet(ws_temp_rev)
                            for img_bytes, anchor in logos_rev:
                                try:
                                    nova_img = xlImage(io.BytesIO(img_bytes))
                                    nova_img.anchor = copy.copy(anchor)
                                    ws_target.add_image(nova_img)
                                except:
                                    pass
                            
                        ws_target.title = nome_aba
                        ws_target["F2"] = f"EX; {job_id}"
                        ws_target["F3"] = f"EX; {seg_name}"
                        ws_target["F4"] = reviewer
                        ws_target["F6"] = data_hoje
                    
                    rev_stream = io.BytesIO()
                    wb_rev.save(rev_stream)
                    zip_file.writestr(f"{qa_geral_path}/Revisao Jobs Samsung - {job_id}.xlsx", rev_stream.getvalue())

                my_bar.progress(85, text="Clonando abas e injetando Logos (Last Check)...")
                if os.path.exists(template_lc_path):
                    wb_lc = load_workbook(template_lc_path)
                    ws_temp_lc = wb_lc.active
                    
                    logos_lc = []
                    for img in ws_temp_lc._images:
                        try:
                            if hasattr(img.ref, 'getvalue'):
                                img_bytes = img.ref.getvalue()
                            elif hasattr(img.ref, 'read'):
                                img.ref.seek(0)
                                img_bytes = img.ref.read()
                            else:
                                img_bytes = img._data()
                            logos_lc.append((img_bytes, copy.copy(img.anchor)))
                        except:
                            pass
                    
                    for sheet_name in wb_lc.sheetnames:
                        if sheet_name != ws_temp_lc.title:
                            del wb_lc[sheet_name]
                            
                    for i, seg in enumerate(segmentos):
                        seg_name = "N/A" if seg == "N_A" else seg
                        raw_aba_name = job_id if seg_name == "N/A" else f"{job_id} - {seg_name}"
                        nome_aba = re.sub(r'[/\\?*\[\]]', '', raw_aba_name)[:31]
                        
                        if i == 0:
                            ws_target = ws_temp_lc
                        else:
                            ws_target = wb_lc.copy_worksheet(ws_temp_lc)
                            for img_bytes, anchor in logos_lc:
                                try:
                                    nova_img = xlImage(io.BytesIO(img_bytes))
                                    nova_img.anchor = copy.copy(anchor)
                                    ws_target.add_image(nova_img)
                                except:
                                    pass
                            
                        ws_target.title = nome_aba
                        ws_target["D1"] = f"Reviewed by: {reviewer} Vasconcelos"
                        ws_target["D2"] = f"LOB:{lob}"
                        ws_target["D3"] = f"JIRA Ticket: {job_id}"
                        ws_target["D4"] = f"Segment: {seg_name}"
                        ws_target["D5"] = data_deploy
                        
                    lc_stream = io.BytesIO()
                    wb_lc.save(lc_stream)
                    zip_file.writestr(f"{qa_geral_path}/Last Check - {job_id}.xlsx", lc_stream.getvalue())

            my_bar.progress(100, text="✅ Pacote ZIP finalizado com sucesso!")
            st.success("🎉 Estrutura ZIP gerada com sucesso!")
            
            st.download_button(
                label="📦 Baixar Estrutura do Job (.zip)",
                data=zip_buffer.getvalue(),
                file_name=f"Job_{job_id}_Estrutura.zip",
                mime="application/zip",
                use_container_width=True
            )
    else:
        st.warning("Puxe os dados do Jira na barra lateral primeiro para habilitar a geração das pastas automáticas.")

# ---------------------------------------------------------------------
# ABA 2: GERADOR DE SUBJECTS (ACC)
# ---------------------------------------------------------------------
with tab2:
    st.header("✉️ ACC Subject Line Code Generator")
    st.write("Faça o upload do seu ficheiro CSV ou TXT contendo as subject lines (ex: `SL1; [Name], text...`).")
    
    uploaded_sl_file = st.file_uploader("Selecione o ficheiro de Subjects (.csv, .txt)", type=["csv", "txt"], key="sl_uploader")
    
    if uploaded_sl_file is not None:
        file_content = uploaded_sl_file.getvalue().decode('utf-8')
        
        parsed_data, generated_code = generate_acc_subject_code(file_content)
        
        if parsed_data:
            st.success("Ficheiro lido com sucesso!")
            
            st.subheader("👀 Preview dos Subjects Encontrados")
            df_preview = pd.DataFrame(parsed_data)[["originalCode", "text"]]
            df_preview.columns = ["Subject Line Code", "Subject Line Text"]
            st.dataframe(df_preview, use_container_width=True)
            
            st.divider()
            st.subheader("💻 Código Gerado (E4X / JavaScript)")
            st.code(generated_code, language='javascript')
            
            st.download_button(
                label="📥 Descarregar Código (.txt)",
                data=generated_code,
                file_name=f"{uploaded_sl_file.name.split('.')[0]}_acc_code.txt",
                mime="text/plain",
                type="primary",
                use_container_width=True
            )
        else:
            st.error(generated_code)
