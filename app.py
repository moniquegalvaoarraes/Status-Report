import streamlit as st
import pandas as pd
import io
import datetime

import plotly.express as px
import plotly.graph_objects as go
import os
from pptx import Presentation
from pptx.util import Inches, Pt
import re

st.set_page_config(page_title="Status Report Assistant", layout="wide")

st.title("📊 Assistente de Status Report")

tab1, tab2, tab3 = st.tabs(["1️⃣ Gerador de Base (Excel)", "2️⃣ Dashboard & Apresentação (PPTX)", "3️⃣ Gerador de Ata (Ata de Reunião)"])

with tab1:
    st.header("Integração ServiceNow & Multidados")
    st.markdown("""
    Faça o upload dos três arquivos abaixo para gerar o seu report de status atualizado e identificar divergências.
    O sistema usará as informações do Multidados como base primária para atualizar o Template.
    """)

    col1, col2, col3 = st.columns(3)

    with col1:
        file_template = st.file_uploader("1️⃣ Template Atual (ZAMP)", type=["xlsx"])
    with col2:
        file_multidados = st.file_uploader("2️⃣ Extração Multidados", type=["xlsx", "csv"])
    with col3:
        file_sn = st.file_uploader("3️⃣ Extração ServiceNow", type=["xlsx", "csv"])

# De/Para Status
STATUS_MAPPING = {
    "Aguardando feedback do cliente": "Pendente",
    "Entendimento": "Em atendimento",
    "Entedimento": "Em atendimento", 
    "Em execução": "Em atendimento",
    "Direcionamento do chamado (Consultor)": "Em atendimento",
    "Redirecionamento do chamado (Devolução)": "Em atendimento",
    "Oportunidade de Melhoria": "Em atendimento"
}

def extract_data_from_multidados(limit_date_str=None):
    import requests
    import hashlib
    import urllib.parse
    import urllib3
    from bs4 import BeautifulSoup
    
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    # Tenta carregar as credenciais a partir do st.secrets (para nuvem)
    # com fallback para manter compatibilidade local padrão.
    USERNAME = st.secrets.get("MULTIDADOS_USERNAME", "monique.arraes")
    PASSWORD_RAW = st.secrets.get("MULTIDADOS_PASSWORD", "Abaco@2024")
    password_md5 = hashlib.md5(PASSWORD_RAW.encode('utf-8')).hexdigest()
    
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    })
    
    try:
        # Laço de repetição robusto (retry) para garantir o login e obtenção do filtro no Multidados
        zamp_filter = None
        ajax_url = "https://abaco.multidadosti.com.br/ajax_bootstrap.php"
        payload = {
            "cm": "Gerenciador_Sets_Filtros->getSetsFiltrosUsuario",
            "cfg[]": "",
            "filters[form_id]": "relatorio_servicedesk"
        }
        
        for tentativa in range(1, 4):
            try:
                # 1. GET login page
                session.get("https://abaco.multidadosti.com.br/login.php", verify=False)
                
                # 2. POST login
                login_data = {
                    "validar_login": "1",
                    "login": USERNAME,
                    "password": password_md5,
                    "password_eh_md5": "1",
                    "active_directory": "F",
                    "portal_cliente_gamaro": "F",
                    "login_autentica_mega": "F",
                    "DB": "ibabaco_novo"
                }
                session.post("https://abaco.multidadosti.com.br/login.php", data=login_data, verify=False)
                
                # 3. GET report page
                r_report = session.get("https://abaco.multidadosti.com.br/servicedesk/?m=ocorrencias&a=relatorio_servicedesk", verify=False)
                if "login.php" in r_report.url:
                    continue
                
                # 4. POST AJAX for saved filters
                r_ajax = session.post(ajax_url, data=payload, verify=False)
                if r_ajax.status_code != 200:
                    continue
                    
                res_json = r_ajax.json()
                if not res_json.get("success"):
                    continue
                    
                for item in res_json.get("data", []):
                    if isinstance(item, dict) and str(item.get("id")) == "414":
                        zamp_filter = item
                        break
                        
                if zamp_filter:
                    break
            except Exception as inner_e:
                import traceback
                try:
                    with open("sync_error.log", "a", encoding="utf-8") as f_err:
                        f_err.write(f"\n--- Erro na tentativa {tentativa} (laço interno) ---\n")
                        traceback.print_exc(file=f_err)
                except Exception:
                    pass
            
            # Pequeno delay antes do retry
            import time
            time.sleep(1)
            
        if not zamp_filter:
            st.error("Não foi possível autenticar ou obter o filtro salvo 'ZAMP - Status Report' (ID 414) no Multidados após 3 tentativas!")
            return None
            
        # 5. Parse saved filter parameters
        filtros_str = zamp_filter["filtros"]
        parsed_params = urllib.parse.parse_qs(filtros_str)
        
        form_data = {}
        for k, v in parsed_params.items():
            # Não incluir filtros de status para trazer todas as ocorrências (inclusive encerradas)
            if k.startswith('idhd_fluxos') or k.startswith('idhd_fluxos_status'):
                continue
            if len(v) == 1:
                form_data[k] = v[0]
            else:
                form_data[k] = v
                
        # Update parameters as requested
        if limit_date_str is None:
            limit_date_str = datetime.date.today().strftime('%d/%m/%Y')
        form_data['d_ini'] = '01/01/2026'
        form_data['d_fim'] = limit_date_str
        form_data['select_data_pre_definida'] = ''
        form_data['radio_data_pre_definida'] = 'data_selecionada'
        form_data['tipo_data'] = 'data_abertura'
        form_data['idconfig_rel'] = '146' # ÁBACO 360 - BASE ZAMP - MONIQUE
        
        # Forçar o filtro de fluxo como indiferente para trazer todos os chamados (ativos e encerrados)
        form_data['tipo_filtro_fluxo'] = ''
        
        # Add Excel export parameters
        form_data['exportar_excel'] = 'on'
        form_data['ok'] = 'Pesquisar'
        
        # 6. Submit report form for Excel export (which returns HTML report table GridHoras)
        report_post_url = "https://abaco.multidadosti.com.br/servicedesk/?m=ocorrencias&a=relatorio_servicedesk"
        export_resp = session.post(report_post_url, data=form_data, verify=False)
        
        if export_resp.status_code != 200:
            st.error(f"Erro na requisição do relatório: Status {export_resp.status_code}")
            return None
            
        # 7. Parse HTML table GridHoras
        soup = BeautifulSoup(export_resp.text, 'html.parser')
        table = soup.find(id='GridHoras')
        if not table:
            try:
                with open("sync_error.log", "a", encoding="utf-8") as f_err:
                    f_err.write("\n--- Tabela GridHoras nao encontrada na resposta ---\n")
                    f_err.write(f"Status da resposta: {export_resp.status_code}\n")
                    f_err.write(f"URL final da resposta: {export_resp.url}\n")
                with open("debug_response.html", "w", encoding="utf-8") as f_deb:
                    f_deb.write(export_resp.text)
            except Exception:
                pass
            st.error("Tabela de dados 'GridHoras' não encontrada no relatório retornado!")
            return None
            
        rows = table.find_all('tr')
        if len(rows) < 3:
            return pd.DataFrame()
            
        headers = [c.get_text().strip() for c in rows[1].find_all(['td', 'th'])]
        headers = [h.replace('\xa0', ' ').strip() for h in headers]
        
        data = []
        for r in rows[2:]:
            cols = [c.get_text().strip() for c in r.find_all(['td', 'th'])]
            cols = [c.replace('\xa0', ' ').strip() for c in cols]
            if len(cols) == len(headers):
                data.append(cols)
                
        df = pd.DataFrame(data, columns=headers)
        
        # Map messy headers to standard names
        cleaned_columns = []
        for col in df.columns:
            col_clean = col
            if 'n.' in col.lower():
                col_clean = 'N.º'
            elif 'ocorr' in col.lower() and 'externa' in col.lower():
                col_clean = 'Nº Ocorrência Externa'
            elif 'titu' in col.lower():
                col_clean = 'Titulo'
            elif 'proje' in col.lower():
                col_clean = 'Projeto'
            elif 'abertura' in col.lower():
                col_clean = 'Data de abertura'
            elif 'modific' in col.lower():
                col_clean = 'Data da última modificação'
            elif 'encerra' in col.lower():
                col_clean = 'Data de encerramento'
            elif 'original' in col.lower():
                col_clean = 'Solicitação Original'
            elif 'priorid' in col.lower():
                col_clean = 'Prioridade'
            elif 'solicita' in col.lower() and 'original' not in col.lower():
                col_clean = 'Solicitação'
            elif 'respons' in col.lower() and 'operador' in col.lower():
                col_clean = 'Operador responsável'
            elif 'status' in col.lower() and 'sem tempo' in col.lower():
                col_clean = 'Status (sem tempo decorrido)'
            elif 'sla de resposta' in col.lower() or 'sla' in col.lower() and 'respos' in col.lower():
                col_clean = 'SLA de resposta'
            elif 'dentro do sla' in col.lower() and 'respos' in col.lower():
                col_clean = 'Resposta dentro do SLA'
            elif 'sla de solu' in col.lower() or 'sla' in col.lower() and 'solu' in col.lower():
                col_clean = 'SLA de solução'
            elif 'dentro do sla' in col.lower() and 'solu' in col.lower():
                col_clean = 'Solução dentro do SLA'
            elif 'sistem' in col.lower():
                col_clean = 'Sistema'
                
            cleaned_columns.append(col_clean)
            
        df.columns = cleaned_columns
        return df
        
    except Exception as e:
        import traceback
        try:
            with open("sync_error.log", "a", encoding="utf-8") as f_err:
                f_err.write(f"\n--- Erro geral na funcao extract_data_from_multidados ---\n")
                traceback.print_exc(file=f_err)
        except Exception:
            pass
        st.error(f"Erro ao conectar com o Multidados: {e}")
        return None

def load_data(uploaded_file, name):
    if uploaded_file is None:
        return None
    
    filename = uploaded_file.name.lower()
    try:
        if filename.endswith(".csv"):
            return pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
        else:
            # Usando calamine para xlsx mais complexos (como o Strict Open XML do Template)
            if name == "Template":
                 # Fallback para engine openpyxl se der erro, mas tentamos calamine primeiro
                 try:
                     return pd.read_excel(uploaded_file, engine='calamine')
                 except:
                     return pd.read_excel(uploaded_file, engine='openpyxl')
            else:
                 return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo {name}: {e}")
        return None

def filter_active_tickets(df, date_cols, status_col, closed_statuses):
    if df is None or df.empty or status_col not in df.columns:
        return df
    
    hoje = datetime.date.today()
    primeiro_dia_mes = pd.Timestamp(hoje.replace(day=1))
    
    # Verifica o status
    mask_status = df[status_col].astype(str).str.strip().str.lower().isin([s.lower() for s in closed_statuses])
    
    # Tenta extrair a data de encerramento mais confiavel
    best_dates = pd.Series(pd.NaT, index=df.index)
    if isinstance(date_cols, str): date_cols = [date_cols]
    
    for col in date_cols:
        if col in df.columns:
            # errors='coerce' transforma strings vazias ou invalidas em pd.NaT
            temp_dates = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
            best_dates = best_dates.fillna(temp_dates)
            
    # É antigo se a melhor data encontrada for antes do mes atual
    mask_antigo = best_dates < primeiro_dia_mes
    
    # Desprezamos SE a regra bate: esta fechado E foi no mes anterior (ou os que estao fechados mas sem data legivel)
    mask_drop = mask_status & (mask_antigo | best_dates.isna())
        
    return df[~mask_drop].copy()

# --- Helpers para Mapeamento de Módulos e Manipulação de Slides ---
import json

MAPPING_FILE = "modulo_melhorias_map.json"

def load_melhorias_mapping():
    if os.path.exists(MAPPING_FILE):
        try:
            with open(MAPPING_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            st.error(f"Erro ao carregar o mapeamento de melhorias: {e}")
    return {}

def save_melhorias_mapping(mapping):
    try:
        with open(MAPPING_FILE, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar o mapeamento de melhorias: {e}")
        return False

def map_priority(val):
    if pd.isna(val):
        return 'Baixa'
    val_str = str(val).upper()
    if 'P1' in val_str or 'MUITO ALTA' in val_str:
        return 'Muito Alta'
    elif 'P2' in val_str or 'ALTA' in val_str:
        return 'Alta'
    elif 'P3' in val_str or 'MEDIA' in val_str or 'MÉDIA' in val_str:
        return 'Média'
    elif 'P4' in val_str or 'BAIXA' in val_str:
        return 'Baixa'
    return 'Baixa'

def delete_slide_properly(prs, index):
    slide = prs.slides[index]
    # Remove from slide ID list
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[index])
    # Drop relationship
    rId_to_remove = None
    for rId, rel in prs.part.rels.items():
        if rel.target_part == slide.part:
            rId_to_remove = rId
            break
    if rId_to_remove:
        prs.part.drop_rel(rId_to_remove)

def move_slide_to_end(prs, slide_index):
    sldIdLst = prs.slides._sldIdLst
    slide_id_el = sldIdLst[slide_index]
    sldIdLst.remove(slide_id_el)
    sldIdLst.append(slide_id_el)

def generate_pptx_presentation(df_ams, df_mel, metrics_ams, metrics_mel, fig_ams_ab, fig_ams_enc, fig_ams_ativos, fig_ams_s1, fig_ams_s2, fig_mel_ab, fig_mel_enc, fig_mel_ativos, fig_mel_s1, fig_mel_s2, slide_contents):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    import io
    
    # Colors (ZAMP / Burger King Theme)
    c_dark_brown = RGBColor(92, 53, 39)
    c_red = RGBColor(214, 35, 0)
    c_orange = RGBColor(245, 162, 0)
    c_cream = RGBColor(245, 235, 220)
    c_white = RGBColor(255, 255, 255)
    
    prs = Presentation("Status Report/Status Report - ZAMP - 260220.pptx")
    
    # 1. Update cover slide (index 0)
    if len(prs.slides) > 0:
        cover_slide = prs.slides[0]
        for shape in cover_slide.shapes:
            if shape.has_text_frame:
                text_upper = shape.text.upper()
                if "STATUS REPORT" in text_upper:
                    shape.text = f"{slide_contents['capa_titulo']}\n{slide_contents['capa_data']}"
                    p = shape.text_frame.paragraphs[0]
                    p.font.name = 'Arial'
                    p.font.size = Pt(36)
                    p.font.bold = True
                    p.font.color.rgb = c_red
                    p.alignment = 1
                    if len(shape.text_frame.paragraphs) > 1:
                        p2 = shape.text_frame.paragraphs[1]
                        p2.font.name = 'Arial'
                        p2.font.size = Pt(20)
                        p2.font.color.rgb = c_dark_brown
                        p2.alignment = 1
                    
    # 2. Delete slides 2 to 10 (indices 1 to 9)
    for _ in range(9):
        delete_slide_properly(prs, 1)
        
    # Helpers
    def add_slide_title(slide, text):
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.33), Inches(0.8))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = 'Arial'
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = c_dark_brown
        
    def add_metric_card(slide, left, top, width, height, title, value):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = c_white
        card.line.color.rgb = c_dark_brown
        card.line.width = Pt(1.5)
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = Inches(0.1)
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = 'Arial'
        p.font.size = Pt(11)
        p.font.color.rgb = c_dark_brown
        p.alignment = 1 # Center
        
        p2 = tf.add_paragraph()
        p2.text = str(value)
        p2.font.name = 'Arial'
        p2.font.size = Pt(26)
        p2.font.bold = True
        p2.font.color.rgb = c_red
        p2.alignment = 1 # Center
        
    def add_bullet_points(slide, left, top, width, height, text):
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        lines = text.split('\n')
        for idx, line in enumerate(lines):
            if idx == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = line
            p.font.name = 'Arial'
            p.font.size = Pt(15)
            p.font.color.rgb = c_dark_brown
            if line.strip().startswith('•') or line.strip().startswith('-'):
                p.level = 0
                
    def add_chart_image(slide, fig, left, top, width, height):
        if fig is None:
            return
        try:
            img_bytes = fig.to_image(format="png", width=800, height=500)
            img_io = io.BytesIO(img_bytes)
            slide.shapes.add_picture(img_io, left, top, width=width, height=height)
        except Exception as e:
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.text = f"[Erro no gráfico: {e}]"
            
    blank_layout = prs.slide_layouts[0]
    
    # Slide 2: AMS - Visão Geral
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_title(s2, slide_contents['ams_overview_title'])
    nome_mes = slide_contents.get('nome_do_mes', 'Mês Vigente')
    add_metric_card(s2, Inches(0.5), Inches(1.5), Inches(3.5), Inches(1.5), f"{nome_mes} (Abertos)", metrics_ams['mes_vigente_abertos'])
    add_metric_card(s2, Inches(0.5), Inches(3.2), Inches(3.5), Inches(1.5), "Backlog (Pendentes)", metrics_ams['backlog_pendentes'])
    add_metric_card(s2, Inches(0.5), Inches(4.9), Inches(3.5), Inches(1.5), "Encerrados no Período", metrics_ams['encerrados_periodo'])
    add_bullet_points(s2, Inches(4.5), Inches(1.5), Inches(8.33), Inches(4.9), slide_contents['ams_overview_bullets'])
    
    # Slide 3: AMS - Volumetria
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_title(s3, slide_contents['ams_vol_title'])
    add_bullet_points(s3, Inches(0.5), Inches(1.3), Inches(12.33), Inches(0.6), slide_contents['ams_vol_text'])
    add_chart_image(s3, fig_ams_ab, Inches(0.5), Inches(2.0), Inches(6.0), Inches(4.8))
    add_chart_image(s3, fig_ams_enc, Inches(6.8), Inches(2.0), Inches(6.0), Inches(4.8))
    
    # Slide 4: AMS - Chamados Ativos
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_title(s4, "ZAMP - BOLSÃO DE TICKETS - Chamados Ativos")
    add_chart_image(s4, fig_ams_ativos, Inches(1.66), Inches(1.5), Inches(10.0), Inches(5.2))
    
    # Slide 5: AMS - SLAs
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_title(s5, slide_contents['ams_sla_title'])
    add_bullet_points(s5, Inches(0.5), Inches(1.3), Inches(12.33), Inches(0.6), slide_contents['ams_sla_text'])
    add_chart_image(s5, fig_ams_s1, Inches(0.5), Inches(2.0), Inches(6.0), Inches(4.8))
    add_chart_image(s5, fig_ams_s2, Inches(6.8), Inches(2.0), Inches(6.0), Inches(4.8))
    
    # Slide 6: Melhorias - Visão Geral
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_title(s6, slide_contents['mel_overview_title'])
    add_metric_card(s6, Inches(0.5), Inches(1.5), Inches(3.5), Inches(1.5), f"{nome_mes} (Abertos)", metrics_mel['mes_vigente_abertos'])
    add_metric_card(s6, Inches(0.5), Inches(3.2), Inches(3.5), Inches(1.5), "Backlog (Pendentes)", metrics_mel['backlog_pendentes'])
    add_metric_card(s6, Inches(0.5), Inches(4.9), Inches(3.5), Inches(1.5), "Encerrados no Período", metrics_mel['encerrados_periodo'])
    add_bullet_points(s6, Inches(4.5), Inches(1.5), Inches(8.33), Inches(4.9), slide_contents['mel_overview_bullets'])
    
    # Slide 7: Melhorias - Volumetria
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_title(s7, slide_contents['mel_vol_title'])
    add_bullet_points(s7, Inches(0.5), Inches(1.3), Inches(12.33), Inches(0.6), slide_contents['mel_vol_text'])
    add_chart_image(s7, fig_mel_ab, Inches(0.5), Inches(2.0), Inches(6.0), Inches(4.8))
    add_chart_image(s7, fig_mel_enc, Inches(6.8), Inches(2.0), Inches(6.0), Inches(4.8))
    
    # Slide 8: Melhorias - Chamados Ativos
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_title(s8, "ZAMP - BOLSÃO DE MELHORIAS - Chamados Ativos")
    add_chart_image(s8, fig_mel_ativos, Inches(1.66), Inches(1.5), Inches(10.0), Inches(5.2))
    
    # Slide 9: Melhorias - SLAs
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_title(s9, slide_contents['mel_sla_title'])
    add_bullet_points(s9, Inches(0.5), Inches(1.3), Inches(12.33), Inches(0.6), slide_contents['mel_sla_text'])
    add_chart_image(s9, fig_mel_s1, Inches(0.5), Inches(2.0), Inches(6.0), Inches(4.8))
    add_chart_image(s9, fig_mel_s2, Inches(6.8), Inches(2.0), Inches(6.0), Inches(4.8))
    
    # Slide 10: Obrigado (Agradecimento)
    s10 = prs.slides.add_slide(blank_layout)
    txBox = s10.shapes.add_textbox(Inches(1.66), Inches(2.75), Inches(10.0), Inches(2.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Obrigada!"
    p.font.name = 'Arial'
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = c_red
    p.alignment = 1 # Center
    
    return prs

def generate_pdf_presentation(df_ams, df_mel, metrics_ams, metrics_mel, fig_ams_ab, fig_ams_enc, fig_ams_ativos, fig_ams_s1, fig_ams_s2, fig_mel_ab, fig_mel_enc, fig_mel_ativos, fig_mel_s1, fig_mel_s2, slide_contents):
    from fpdf import FPDF
    import io
    import os
    
    class PDFPresentation(FPDF):
        def __init__(self):
            # 13.333 x 7.5 inches = 338.67 x 190.5 mm
            super().__init__(orientation='L', unit='mm', format=(190.5, 338.67))
            self.set_margins(0, 0, 0)
            self.set_auto_page_break(False)
            
        def add_slide(self, title_text, slide_num_str):
            self.add_page()
            # Background
            self.set_fill_color(245, 235, 220)
            self.rect(0, 0, 338.67, 190.5, 'F')
            
            # Banner
            self.set_fill_color(214, 35, 0)
            self.rect(0, 0, 338.67, 4, 'F')
            
            # Title
            self.set_text_color(92, 53, 39)
            self.set_font("Arial", "B", 24)
            self.set_xy(15, 15)
            title_encoded = title_text.encode('latin1', 'ignore').decode('latin1')
            self.cell(0, 15, title_encoded, ln=True)
            
            # Footer
            self.set_font("Arial", "I", 10)
            self.set_text_color(92, 53, 39)
            self.set_xy(15, 178)
            self.cell(100, 10, "Abaco 360 - Gestao de Operacao ZAMP".encode('latin1', 'ignore').decode('latin1'))
            self.set_xy(290, 178)
            self.cell(33, 10, slide_num_str, align="R")
            
        def draw_metric_card(self, x, y, w, h, title, value):
            self.set_fill_color(255, 255, 255)
            self.set_draw_color(92, 53, 39)
            self.set_line_width(0.5)
            self.rect(x, y, w, h, 'FD')
            
            self.set_text_color(92, 53, 39)
            self.set_font("Arial", "B", 10)
            self.set_xy(x, y + 4)
            self.cell(w, 5, title.encode('latin1', 'ignore').decode('latin1'), align="C", ln=True)
            
            self.set_text_color(214, 35, 0)
            self.set_font("Arial", "B", 26)
            self.set_xy(x, y + 15)
            self.cell(w, 15, str(value), align="C")
            
        def draw_bullets(self, x, y, w, h, text):
            self.set_text_color(92, 53, 39)
            self.set_font("Arial", "", 14)
            self.set_xy(x, y)
            
            lines = text.split('\n')
            for line in lines:
                line_encoded = line.encode('latin1', 'ignore').decode('latin1')
                self.set_x(x)
                self.multi_cell(w, 7, line_encoded)
                self.ln(2)
                
        def draw_chart(self, fig, x, y, w, h):
            if fig is None:
                return
            try:
                img_bytes = fig.to_image(format="png", width=800, height=500)
                temp_file = f"temp_pdf_chart_{x}_{y}.png"
                with open(temp_file, "wb") as f:
                    f.write(img_bytes)
                self.image(temp_file, x=x, y=y, w=w, h=h)
                os.remove(temp_file)
            except Exception as e:
                self.set_xy(x, y)
                self.set_font("Arial", "I", 12)
                self.cell(w, h, f"[Erro ao renderizar grafico: {e}]".encode('latin1', 'ignore').decode('latin1'), align="C")

    pdf = PDFPresentation()
    
    # Slide 1: Capa
    pdf.add_page()
    pdf.set_fill_color(245, 235, 220)
    pdf.rect(0, 0, 338.67, 190.5, 'F')
    pdf.set_fill_color(214, 35, 0)
    pdf.rect(0, 0, 338.67, 6, 'F')
    
    pdf.set_text_color(214, 35, 0)
    pdf.set_font("Arial", "B", 34)
    pdf.set_xy(15, 55)
    title_text = slide_contents['capa_titulo'].encode('latin1', 'ignore').decode('latin1')
    pdf.multi_cell(308.67, 15, title_text, align="C")
    
    pdf.set_text_color(92, 53, 39)
    pdf.set_font("Arial", "", 20)
    pdf.set_y(pdf.get_y() + 10)
    pdf.cell(308.67, 10, "Abaco 360 - Gestao de Operacao ZAMP".encode('latin1', 'ignore').decode('latin1'), align="C", ln=True)
    pdf.ln(10)
    
    pdf.set_fill_color(92, 53, 39)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 14)
    date_str = slide_contents['capa_data']
    pdf.set_xy(139.33, pdf.get_y() + 5)
    pdf.cell(60, 10, date_str.encode('latin1', 'ignore').decode('latin1'), align="C", fill=True)
    
    nome_mes = slide_contents.get('nome_do_mes', 'Mês Vigente')
    # Slide 2: AMS Overview
    pdf.add_slide(slide_contents['ams_overview_title'], "Slide 2")
    pdf.draw_metric_card(15, 40, 90, 35, f"{nome_mes} (Abertos)", metrics_ams['mes_vigente_abertos'])
    pdf.draw_metric_card(15, 82, 90, 35, "Backlog (Pendentes)", metrics_ams['backlog_pendentes'])
    pdf.draw_metric_card(15, 124, 90, 35, "Encerrados no Periodo", metrics_ams['encerrados_periodo'])
    pdf.draw_bullets(115, 40, 208, 120, slide_contents['ams_overview_bullets'])
    
    # Slide 3: AMS Volumetria
    pdf.add_slide(slide_contents['ams_vol_title'], "Slide 3")
    pdf.set_font("Arial", "I", 12)
    pdf.set_xy(15, 30)
    pdf.cell(0, 10, slide_contents['ams_vol_text'].encode('latin1', 'ignore').decode('latin1'), ln=True)
    pdf.draw_chart(fig_ams_ab, 15, 42, 148, 125)
    pdf.draw_chart(fig_ams_enc, 175, 42, 148, 125)
    
    # Slide 4: AMS Chamados Ativos
    pdf.add_slide("ZAMP - BOLSÃO DE TICKETS - Chamados Ativos", "Slide 4")
    pdf.draw_chart(fig_ams_ativos, 44, 40, 250, 130)
    
    # Slide 5: AMS SLAs
    pdf.add_slide(slide_contents['ams_sla_title'], "Slide 5")
    pdf.set_font("Arial", "I", 12)
    pdf.set_xy(15, 30)
    pdf.cell(0, 10, slide_contents['ams_sla_text'].encode('latin1', 'ignore').decode('latin1'), ln=True)
    pdf.draw_chart(fig_ams_s1, 15, 42, 148, 125)
    pdf.draw_chart(fig_ams_s2, 175, 42, 148, 125)
    
    # Slide 6: Melhorias Overview
    pdf.add_slide(slide_contents['mel_overview_title'], "Slide 6")
    pdf.draw_metric_card(15, 40, 90, 35, f"{nome_mes} (Abertos)", metrics_mel['mes_vigente_abertos'])
    pdf.draw_metric_card(15, 82, 90, 35, "Backlog (Pendentes)", metrics_mel['backlog_pendentes'])
    pdf.draw_metric_card(15, 124, 90, 35, "Encerrados no Periodo", metrics_mel['encerrados_periodo'])
    pdf.draw_bullets(115, 40, 208, 120, slide_contents['mel_overview_bullets'])
    
    # Slide 7: Melhorias Volumetria
    pdf.add_slide(slide_contents['mel_vol_title'], "Slide 7")
    pdf.set_font("Arial", "I", 12)
    pdf.set_xy(15, 30)
    pdf.cell(0, 10, slide_contents['mel_vol_text'].encode('latin1', 'ignore').decode('latin1'), ln=True)
    pdf.draw_chart(fig_mel_ab, 15, 42, 148, 125)
    pdf.draw_chart(fig_mel_enc, 175, 42, 148, 125)
    
    # Slide 8: Melhorias Chamados Ativos
    pdf.add_slide("ZAMP - BOLSÃO DE MELHORIAS - Chamados Ativos", "Slide 8")
    pdf.draw_chart(fig_mel_ativos, 44, 40, 250, 130)
    
    # Slide 9: Melhorias SLAs
    pdf.add_slide(slide_contents['mel_sla_title'], "Slide 9")
    pdf.set_font("Arial", "I", 12)
    pdf.set_xy(15, 30)
    pdf.cell(0, 10, slide_contents['mel_sla_text'].encode('latin1', 'ignore').decode('latin1'), ln=True)
    pdf.draw_chart(fig_mel_s1, 15, 42, 148, 125)
    pdf.draw_chart(fig_mel_s2, 175, 42, 148, 125)
    
    # Slide 10: Obrigado
    pdf.add_page()
    pdf.set_fill_color(245, 235, 220)
    pdf.rect(0, 0, 338.67, 190.5, 'F')
    pdf.set_fill_color(214, 35, 0)
    pdf.rect(0, 0, 338.67, 6, 'F')
    
    pdf.set_text_color(214, 35, 0)
    pdf.set_font("Arial", "B", 48)
    pdf.set_xy(15, 80)
    pdf.cell(308.67, 20, "Obrigada!".encode('latin1', 'ignore').decode('latin1'), align="C")
    
    return pdf

with tab1:
    if file_template and file_multidados and file_sn:
        with st.spinner("Processando planilhas..."):
            df_template = load_data(file_template, "Template")
            df_md = load_data(file_multidados, "Multidados")
            df_sn = load_data(file_sn, "ServiceNow")

        if df_template is not None and df_md is not None and df_sn is not None:

            # Filtrar chamados Encerrados/Cancelados em meses anteriores (Apenas nas bases de origem por enquanto)
            closed_md = ['Encerrado', 'Encerrada', 'Cancelado', 'Cancelada']
            closed_sn = ['Encerrado', 'Cancelado']

            # Multidados
            df_md_clean = filter_active_tickets(df_md, date_cols=['Data de encerramento', 'Data da Última modificação'], status_col='Status (sem tempo decorrido)', closed_statuses=closed_md)

            # ServiceNow
            df_sn_clean = filter_active_tickets(df_sn, date_cols=['Atualizado em'], status_col='Estado', closed_statuses=closed_sn)

            st.divider()
            st.subheader("Análise e Cruzamento de Dados")

            col_md_id = 'N.º'
            col_sn_id = 'Número'
            for opt in ['Número', 'Number', 'number']:
                if opt in df_sn.columns:
                    col_sn_id = opt
                    break
            col_tpl_id = 'N.º'

            # Garante string para comparaçao
            if col_md_id in df_md.columns: df_md[col_md_id] = df_md[col_md_id].astype(str).str.strip().str.replace('.0', '', regex=False)
            if col_md_id in df_md_clean.columns: df_md_clean[col_md_id] = df_md_clean[col_md_id].astype(str).str.strip().str.replace('.0', '', regex=False)
            if col_sn_id in df_sn.columns: df_sn[col_sn_id] = df_sn[col_sn_id].astype(str).str.strip()
            if col_sn_id in df_sn_clean.columns: df_sn_clean[col_sn_id] = df_sn_clean[col_sn_id].astype(str).str.strip()
            if col_tpl_id in df_template.columns: df_template[col_tpl_id] = df_template[col_tpl_id].astype(str).str.strip().str.replace('.0', '', regex=False)

            # 1. Novos no Multidados
            new_in_md = df_md_clean[~df_md_clean[col_md_id].isin(df_template[col_tpl_id])]
            if not new_in_md.empty:
                st.info(f"✨ Foram encontrados **{len(new_in_md)} novos chamados** no Multidados. Eles foram adicionados ao Template.")
                df_template = pd.concat([df_template, new_in_md], ignore_index=True)
            else:
                st.success("✅ Nenhum chamado novo no Multidados para adicionar ao Template.")

            # 2. Novos no ServiceNow (que não possuem correspondência no Multidados)
            # O ServiceNow usa o 'Número' para cruzar com o 'Nº Ocorrência Externa' do Multidados
            if 'Nº Ocorrência Externa' in df_md.columns:
                # Lista de ocorrencias externas no MD (tirando nulos) - USAR BASE COMPLETA PARA EVITAR FALSO POSITIVO
                md_ext_ids_all = df_md['Nº Ocorrência Externa'].astype(str).str.strip().str.replace('.0', '', regex=False)
                md_ext_ids_all = md_ext_ids_all[md_ext_ids_all != 'nan']

                # SN (filtrado) que nao estao no MD Ocorrencia Externa (completo)
                if col_sn_id in df_sn_clean.columns:
                    sn_not_in_md = df_sn_clean[~df_sn_clean[col_sn_id].isin(md_ext_ids_all)]
                    if not sn_not_in_md.empty:
                        st.warning(f"⚠️ Atenção! Encontrados **{len(sn_not_in_md)} chamados** no ServiceNow que **NÃO constam no Multidados** (Nº Ocorrência Externa):")
                        cols_to_disp = [c for c in [col_sn_id, 'Estado', 'Atribuído a'] if c in sn_not_in_md.columns]
                        st.dataframe(sn_not_in_md[cols_to_disp].head(10))
                    else:
                        st.success("✅ Todos os chamados do ServiceNow constam no Multidados.")
                else:
                    st.warning(f"Não foi possível validar os chamados do ServiceNow faltantes pois a coluna '{col_sn_id}' não foi encontrada.")

                # 2.5 - Chamados no MD que não estão no SN (pelo Nº Ocorrência Externa)
                # Pegar ocorrencias no MD que não sejam nulas
                valid_md_ext = df_md_clean[df_md_clean['Nº Ocorrência Externa'].notna() & (df_md_clean['Nº Ocorrência Externa'].astype(str).str.strip() != '')]
                
                # Lista de chaves do SN completo para evitar falsos positivos de chamados ocultos pelo filtro de data
                sn_all_ids = df_sn[col_sn_id].astype(str).str.strip() if col_sn_id in df_sn.columns else pd.Series(dtype=str)
                
                md_not_in_sn = valid_md_ext[~valid_md_ext['Nº Ocorrência Externa'].astype(str).str.strip().str.replace('.0', '', regex=False).isin(sn_all_ids)]

                # Lista de Ocorrencias Externas faltando no SN para pintarmos de vermelho depois
                missing_in_sn_ids = set(md_not_in_sn['Nº Ocorrência Externa'].astype(str).str.strip().str.replace('.0', '', regex=False).tolist())

                if not md_not_in_sn.empty:
                    st.error(f"⚠️ Atenção! Encontrados **{len(md_not_in_sn)} chamados** no Multidados cujo *Nº Ocorrência Externa* **NÃO consta na extração do ServiceNow**. Eles serão destacados em vermelho na planilha final.")
                    st.dataframe(md_not_in_sn[[col_md_id, 'Nº Ocorrência Externa', 'Status (sem tempo decorrido)']].head(10))
            else:
                missing_in_sn_ids = set()
                st.error("Coluna 'Nº Ocorrência Externa' não encontrada no Multidados para fazer as validações com ServiceNow.")


            # 3. Comparar Existentes
            divergences = []

            for idx, row in df_template.iterrows():
                t_id = str(row.get(col_tpl_id, ''))

                # Localizar no MD COMPLETO para garantir atualização de status mesmo de chamados recém-encerrados
                md_match = df_md[df_md[col_md_id] == t_id]
                if md_match.empty: continue

                md_row = md_match.iloc[0]

                # Atualizar Template com dados do MD se diferente
                cols_sync = ['Status (sem tempo decorrido)', 'Data de encerramento', 'Data da Última modificação', 'Data de abertura', 'Titulo', 'Projeto', 'Prioridade']
                for c in cols_sync:
                    if c in md_row.index and c in df_template.columns:
                        val_md = md_row[c]
                        val_tpl = row[c]
                        if pd.notna(val_md) and val_md != val_tpl:
                            df_template.at[idx, c] = val_md

                # DE/PARA especificos SLA
                if 'Resposta dentro do SLA' in md_row.index and 'Status do SLA de resposta' in df_template.columns:
                    val = md_row['Resposta dentro do SLA']
                    if pd.notna(val): df_template.at[idx, 'Status do SLA de resposta'] = val

                if 'Solução dentro do SLA' in md_row.index and 'Status do SLA de Solução' in df_template.columns:
                    val = md_row['Solução dentro do SLA']
                    if pd.notna(val): df_template.at[idx, 'Status do SLA de Solução'] = val

                # Novas regras de DE/PARA Multidados -> Template
                # Nº Ocorrência Externa -> ServiceNow
                if 'Nº Ocorrência Externa' in md_row.index and 'ServiceNow' in df_template.columns:
                    val = md_row['Nº Ocorrência Externa']
                    if pd.notna(val):
                        df_template.at[idx, 'ServiceNow'] = val

                # Operador responsável -> Atribuído
                if 'Operador responsável' in md_row.index and 'Atribuído' in df_template.columns:
                    val = md_row['Operador responsável']
                    if pd.notna(val):
                        # Nao sobreescrever se a regra do ServiceNow ja tiver preenchido (ver abaixo)
                        if pd.isna(df_template.at[idx, 'Atribuído']) or str(df_template.at[idx, 'Atribuído']).strip() == '':
                             df_template.at[idx, 'Atribuído'] = val

                # Solicitação -> Módulo (com limpeza de texto)
                if 'Solicitação' in md_row.index and 'Módulo' in df_template.columns:
                    val = md_row['Solicitação']
                    if pd.notna(val):
                        val_str = str(val).strip().upper()
                        # Limpeza especial do prefixo SERVICOS
                        val_str = val_str.replace('SERVICOS ', '').replace('SERVIÇOS ', '')
                        df_template.at[idx, 'Módulo'] = val_str

                # Verificar Divergencia com ServiceNow e aplicar DE/PARA ServiceNow -> Template
                num_ocorr_ext = str(row.get('Nº Ocorrência Externa', '')).strip().replace('.0', '')
                if num_ocorr_ext and col_sn_id in df_sn.columns:
                    sn_match = df_sn[df_sn[col_sn_id] == num_ocorr_ext]
                    if not sn_match.empty:
                        sn_row = sn_match.iloc[0]

                        # Solicitante (SN) -> Solicitante (Template)
                        if 'Solicitante' in sn_row.index and 'Solicitante' in df_template.columns:
                            val_solic = sn_row['Solicitante']
                            if pd.notna(val_solic):
                                df_template.at[idx, 'Solicitante'] = val_solic

                        md_status = md_row.get('Status (sem tempo decorrido)', '')
                        sn_status = sn_row.get('Estado', '')

                        expected_sn_status = STATUS_MAPPING.get(md_status, 'Não mapeado')

                        if expected_sn_status != 'Não mapeado' and str(sn_status).strip().lower() != expected_sn_status.lower():
                            divergences.append({
                                'Nº (Multidados)': t_id,
                                'Nº (ServiceNow)': num_ocorr_ext,
                                'Operador Responsável': md_row.get('Operador responsável', ''),
                                'Status Multidados': md_status,
                                'Status ServiceNow (Atual)': sn_status,
                                'Status ServiceNow (Esperado)': expected_sn_status
                            })

                        # Regra nova: SN Encerrado/Resolvido mas MD não encerrado
                        if str(sn_status).strip().lower() in ['encerrado', 'resolvido']:
                            if str(md_status).strip().lower() not in ['encerrado', 'resolvido', 'encerrada', 'resolvida']:
                                st.error(f"🛑 Divergência Crítica: O chamado **SN {num_ocorr_ext}** (MD {t_id}) está marcado como *{sn_status}* no ServiceNow, mas no Multidados consta como *{md_status}*.")

            if divergences:
                st.warning(f"❌ Foram encontradas **{len(divergences)} divergências** de status entre Multidados e ServiceNow baseadas na tabela De/Para:")
                st.dataframe(pd.DataFrame(divergences))

            # 4. Filter final Template (drop old closed tickets) and Download
            st.divider()
            st.subheader("📥 Baixar Template Atualizado")

            # Agora que o Template foi atualizado com as datas reais do MD, podemos filtrar limpo
            df_template = filter_active_tickets(df_template, date_cols=['Data de encerramento', 'Data da Última modificação'], status_col='Status (sem tempo decorrido)', closed_statuses=closed_md)

            COLUMNS_TO_KEEP = [
                'N.º', 'ServiceNow', 'Titulo', 'Projeto', 'Data de abertura', 
                'Data da Última modificação', 'Data de encerramento', 'Solicitante', 
                'Prioridade', 'Módulo', 'Atribuído', 'Status (sem tempo decorrido)', 
                'Status do SLA de resposta', 'Status do SLA de Solução'
            ]

            # --- Construção da Base Final ---
            st.subheader("Visualização do Template Atualizado")

            # Mapeamento do sys_id do ServiceNow para construção do Link correto
            # A extração do ServiceNow geralmente tem uma coluna oculta chamda 'sys_id' ou um Link. 
            # Vamos mapear pelo Numero.
            sn_sys_id_map = {}
            # Algumas extrações tem 'sys_id' nativo ou Link nativo
            if 'sys_id' in df_sn_clean.columns:
                for _, row in df_sn_clean.iterrows():
                    num = str(row.get(col_sn_id, '')).strip()
                    if num: sn_sys_id_map[num] = row['sys_id']

            # Filtra e organiza Colunas do Template
            for col in COLUMNS_TO_KEEP:
                if col not in df_template.columns:
                    df_template[col] = None

            df_final = df_template[COLUMNS_TO_KEEP].copy()

            # Formatar campos de Data para DD/MM/YYYY
            date_cols = ['Data de abertura', 'Data da Última modificação', 'Data de encerramento']
            for dc in date_cols:
                if dc in df_final.columns:
                    df_final[dc] = pd.to_datetime(df_final[dc], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')

            output = io.BytesIO()

            import openpyxl
            from openpyxl.styles import PatternFill
            from openpyxl.worksheet.table import Table, TableStyleInfo

            red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

            # Criar um novo arquivo limpo e seguro
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Unificada'

            # 1. Escrever o Cabeçalho
            ws.append(COLUMNS_TO_KEEP)

            # Identificar as colunas chaves
            sn_col_idx = None
            for idx, col in enumerate(COLUMNS_TO_KEEP, start=1):
                if col == 'ServiceNow':
                    sn_col_idx = idx
                    break

            # 2. Escrever os Dados e Formatar
            for row_idx, row_data in enumerate(df_final[COLUMNS_TO_KEEP].itertuples(index=False), start=2):
                for col_idx, val in enumerate(row_data, start=1):
                    col_name = COLUMNS_TO_KEEP[col_idx - 1]

                    # Tratar valores vazios do pandas
                    if pd.isna(val):
                        val_str = ""
                    else:
                        val_str = str(val).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]

                    # Lógica Especial para a coluna ServiceNow (Hiperlink)
                    if col_name == 'ServiceNow' and val_str:
                        ws.cell(row=row_idx, column=col_idx, value=val_str)

                        # Tentar pegar o sys_id do mapeamento
                        sys_id = sn_sys_id_map.get(val_str)

                        if sys_id:
                            # Se for RITM usa sc_req_item
                            table = "sc_req_item" if val_str.upper().startswith("RITM") else ("change_request" if val_str.upper().startswith("CHG") else "incident")
                            link = f"https://burgerking.service-now.com/nav_to.do?uri=%2F{table}.do%3Fsys_id%3D{sys_id}"
                        else:
                            table = "sc_req_item" if val_str.upper().startswith("RITM") else ("change_request" if val_str.upper().startswith("CHG") else "incident")
                            link = f"https://burgerking.service-now.com/nav_to.do?uri=%2F{table}.do%3Fsysparm_query%3Dnumber%3D{val_str}"

                        ws.cell(row=row_idx, column=col_idx).hyperlink = link
                        ws.cell(row=row_idx, column=col_idx).style = "Hyperlink"
                    else:
                        ws.cell(row=row_idx, column=col_idx, value=val_str)

                # 3. Pintar a linha de vermelho se o ticket MD nao tiver no SN
                if sn_col_idx and missing_in_sn_ids:
                    cell_val = str(ws.cell(row=row_idx, column=sn_col_idx).value or "").strip()
                    if cell_val in missing_in_sn_ids:
                        for col_offset in range(1, len(COLUMNS_TO_KEEP) + 1):
                            ws.cell(row=row_idx, column=col_offset).fill = red_fill

            # 4. Transformar tudo numa Tabela Formatada (Azul - Estilo Médio 9)
            max_col_letter = openpyxl.utils.get_column_letter(len(COLUMNS_TO_KEEP))
            ref = f"A1:{max_col_letter}{ws.max_row}"

            tab = Table(displayName="StatusReportData", ref=ref)
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            ws.add_table(tab)

            # Salvar as dimensões de colunas pra ficar mais legível
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter # Get the column name
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                if adjusted_width > 50: adjusted_width = 50 # limite máximo
                ws.column_dimensions[column].width = adjusted_width

            wb.save(output)

            st.download_button(
                label="Baixar Template de Status Report (.xlsx)",
                data=output.getvalue(),
                file_name=f"ZAMP - Status Report Atualizado - {datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )



with tab2:
    st.header("📊 Dashboard & Geração de Apresentação (PPTX & PDF)")
    st.markdown("""
    Nesta aba, faça o upload de uma base unificada (Excel). O sistema irá estruturar os indicadores de **Status Report** de forma interativa, separar os dados pelos projetos **AMS** e **Melhorias**, permitir o mapeamento persistente de módulos SAP, exibir um **Slide Previewer** interativo para edição manual de textos e disponibilizar o download da apresentação final em **PowerPoint (PPTX)** e **PDF**.
    """)
    
    # 1. Carregar os dados extraídos pelo script local 'multidados_live_data.xlsx' no início do render
    # Usar caminho absoluto dinâmico baseado na pasta onde está o app.py para evitar erros de CWD
    base_dir = os.path.dirname(os.path.abspath(__file__))
    local_extracted_file = os.path.join(base_dir, "multidados_live_data.xlsx")
    if os.path.exists(local_extracted_file):
        try:
            try:
                df_local_live = pd.read_excel(local_extracted_file, engine='calamine')
            except Exception:
                df_local_live = pd.read_excel(local_extracted_file)
            st.session_state['df_multidados_live'] = df_local_live
        except Exception as e:
            st.session_state['df_multidados_live'] = None
    else:
        st.session_state['df_multidados_live'] = None

    # 2. Painel de Diagnóstico do Multidados (Expandível)
    with st.expander("🔍 DIAGNÓSTICO DO GRÁFICO (Dados do Multidados)"):
        st.subheader("Status de Leitura dos Dados do Multidados")
        if 'df_multidados_live' in st.session_state and st.session_state['df_multidados_live'] is not None:
            df_diag = st.session_state['df_multidados_live']
            st.success(f"Dados do Multidados carregados com sucesso a partir de 'multidados_live_data.xlsx'! {len(df_diag)} registros recuperados.")
            
            # Filtrar e contar os chamados de AMS para o diagnóstico
            df_diag_ams = df_diag[df_diag['Projeto'] == 'ZAMP - AMS - TICKETS'].copy()
            df_diag_ams['_DataAbertura'] = pd.to_datetime(df_diag_ams['Data de abertura'], errors='coerce', dayfirst=True)
            df_diag_ams['MesAnoAbertura'] = df_diag_ams['_DataAbertura'].dt.strftime('%m/%Y')
            
            status_col = 'Status (sem tempo decorrido)'
            if status_col in df_diag_ams.columns:
                status_series = df_diag_ams[status_col].astype(str).str.strip().str.lower()
                mask_exclude = status_series.isin(['cancelado', 'cancelada', 'status report']) | status_series.str.contains('status report|cancelad', case=False, na=False)
                df_diag_clean = df_diag_ams[~mask_exclude].copy()
            else:
                df_diag_clean = df_diag_ams.copy()
                
            months_diag = ['01/2026', '02/2026', '03/2026', '04/2026', '05/2026', '06/2026']
            counts_diag = df_diag_clean['MesAnoAbertura'].value_counts().reindex(months_diag, fill_value=0).reset_index()
            counts_diag.columns = ['Mês de Abertura', 'Quantidade AMS (Sem Cancelados)']
            
            st.markdown("**Dados de Abertura Coletados do Multidados para o Gráfico:**")
            st.dataframe(counts_diag, use_container_width=True)
        else:
            st.warning("O arquivo de dados extraídos 'multidados_live_data.xlsx' não foi localizado na pasta do projeto. O gráfico de abertos está operando temporariamente com os dados da planilha Excel carregada abaixo (Fallback). Por favor, execute o script 'testar_extracao_multidados.py' para gerar os dados do gráfico.")
    
    df_all = None
    
    file_consolidated = st.file_uploader("📥 Upload do Relatório de Status Consolidado (Excel)", type=["xlsx"], key="unified_status_report_uploader")
    if file_consolidated:
        with st.spinner("Processando dados do Excel..."):
            df_all = load_data(file_consolidated, "Status Report")
            
        # Tentar extrair a data limite do nome do arquivo (ex: ZAMP - Status Report - 20260624)
        limit_date_str = None
        match = re.search(r'\d{8}', file_consolidated.name)
        if match:
            date_str = match.group(0)
            try:
                dt_limit = datetime.datetime.strptime(date_str, '%Y%m%d')
                limit_date_str = dt_limit.strftime('%d/%m/%Y')
            except ValueError:
                pass
            
        # O arquivo de dados do Multidados já foi carregado dinamicamente no início do render
        pass
            
    if True:
        if df_all is not None and not df_all.empty:
            # Clean columns
            df_all.columns = [c.strip() for c in df_all.columns]
            
            # Check for Projeto column
            proj_col = 'Projeto' if 'Projeto' in df_all.columns else None
            if proj_col is None and len(df_all.columns) > 3:
                proj_col = df_all.columns[3]
                df_all.rename(columns={proj_col: 'Projeto'}, inplace=True)
                proj_col = 'Projeto'
                
            if proj_col not in df_all.columns:
                st.error("Coluna 'Projeto' não encontrada no arquivo Excel. Certifique-se de que é a coluna D.")
                st.stop()
                
            # Date conversion
            data_abertura_col = 'Data de abertura' if 'Data de abertura' in df_all.columns else None
            data_encerra_col = 'Data de encerramento' if 'Data de encerramento' in df_all.columns else None
            
            if data_abertura_col:
                df_all['_DataAbertura'] = pd.to_datetime(df_all[data_abertura_col], errors='coerce', dayfirst=True)
            else:
                df_all['_DataAbertura'] = pd.NaT
                
            if data_encerra_col:
                df_all['_DataEncerra'] = pd.to_datetime(df_all[data_encerra_col], errors='coerce', dayfirst=True)
            else:
                df_all['_DataEncerra'] = pd.NaT
                
            # Available months
            df_all['MesAnoAbertura'] = df_all['_DataAbertura'].dt.strftime('%m/%Y')
            available_months = df_all[df_all['MesAnoAbertura'].notna()]['MesAnoAbertura'].unique().tolist()
            if available_months:
                available_months.sort(key=lambda x: pd.to_datetime(x, format='%m/%Y'))
                
            if not available_months:
                st.error("Nenhuma data de abertura válida encontrada no arquivo Excel.")
                st.stop()
                
            # Period selection
            st.divider()
            st.subheader("📅 Período do Report")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                selected_month_str = st.selectbox("Selecione o Mês Vigente para o Report:", options=available_months, index=len(available_months)-1, key="selected_report_month")
                selected_vigente_dt = pd.to_datetime(selected_month_str, format='%m/%Y')
                # Traduzir o mês vigente para português
                months_pt = {
                    '01': 'Janeiro', '02': 'Fevereiro', '03': 'Março', '04': 'Abril',
                    '05': 'Maio', '06': 'Junho', '07': 'Julho', '08': 'Agosto',
                    '09': 'Setembro', '10': 'Outubro', '11': 'Novembro', '12': 'Dezembro'
                }
                vigente_month_num = selected_month_str.split('/')[0]
                nome_do_mes = months_pt.get(vigente_month_num, "Mês Vigente")
                if 'slide_contents' in st.session_state:
                    st.session_state['slide_contents']['nome_do_mes'] = nome_do_mes
            with col_p2:
                st.info(f"O relatório irá analisar as volumetrias dos últimos 12 meses e a conformidade de SLA de **{nome_do_mes}**.")
                
            # SAP Module Mapping for Melhorias
            st.divider()
            st.subheader("⚙️ Mapeamento de Módulos SAP (Melhorias)")
            st.markdown("""
            Para os chamados do projeto **Bolsão de Melhorias**, você pode atualizar a coluna de **Módulo** para mapear os dados para módulos SAP estruturados (ex: FI, MM, SD). 
            Chamados novos ainda não mapeados são destacados em **vermelho suave**. Ajuste o módulo atualizado de cada chamado e clique em **Salvar Mapeamento** para prosseguir.
            """)
            
            modulo_col = 'Módulo' if 'Módulo' in df_all.columns else None
            if modulo_col:
                is_melh_mask = df_all[modulo_col].astype(str).str.contains('melhoria', case=False, na=False)
                df_melh_tickets = df_all[is_melh_mask].copy()
            else:
                df_melh_tickets = pd.DataFrame()
                
            if not df_melh_tickets.empty:
                mapping = load_melhorias_mapping()
                df_mapping_view = []
                new_tickets_count = 0
                
                for _, row in df_melh_tickets.iterrows():
                    t_id = str(row.get('N.º', '')).strip().replace('.0', '')
                    sn_id = str(row.get('ServiceNow', '')).strip()
                    titulo = str(row.get('Titulo', '')).strip()
                    proj = str(row.get('Projeto', '')).strip()
                    mod_orig = str(row.get('Módulo', '')).strip()
                    
                    mod_mapped = mapping.get(t_id, "")
                    is_new = "Sim" if t_id not in mapping else "Não"
                    if is_new == "Sim":
                        new_tickets_count += 1
                        
                    df_mapping_view.append({
                        'N.º': t_id,
                        'ServiceNow': sn_id,
                        'Título': titulo,
                        'Projeto': proj,
                        'Módulo Original': mod_orig,
                        'Módulo Atualizado': mod_mapped,
                        'Novo?': is_new
                    })
                    
                df_mapping_view = pd.DataFrame(df_mapping_view)
                
                if new_tickets_count > 0:
                    st.warning(f"⚠️ Encontrado(s) **{new_tickets_count} novo(s) chamado(s)** de Melhorias pendente(s) de mapeamento (destacados em vermelho suave).")
                else:
                    st.success("✅ Todos os chamados de Melhorias já estão mapeados.")
                    
                def highlight_new_rows(row):
                    if row['Novo?'] == 'Sim':
                        return ['background-color: #ffe6e6; color: #cc0000; font-weight: bold'] * len(row)
                    return [''] * len(row)
                    
                styled_mapping_df = df_mapping_view.style.apply(highlight_new_rows, axis=1)
                
                sap_modules = ["FI", "MM", "ABAP", "AA", "BASIS", "SD", "PS", "ARIBA", "CO", "PP", "WM", "QM", "HR", "PM", "BI", "Outros"]
                
                edited_mapping_df = st.data_editor(
                    styled_mapping_df,
                    column_config={
                        "N.º": st.column_config.TextColumn("N.º", disabled=True),
                        "ServiceNow": st.column_config.TextColumn("ServiceNow", disabled=True),
                        "Título": st.column_config.TextColumn("Título", disabled=True),
                        "Projeto": st.column_config.TextColumn("Projeto", disabled=True),
                        "Módulo Original": st.column_config.TextColumn("Módulo Original", disabled=True),
                        "Módulo Atualizado": st.column_config.SelectboxColumn(
                            "Módulo Atualizado",
                            options=sap_modules,
                            required=True
                        ),
                        "Novo?": st.column_config.TextColumn("Novo?", disabled=True)
                    },
                    hide_index=True,
                    key="sap_module_editor",
                    use_container_width=True
                )
                
                if st.button("💾 Salvar Mapeamento", type="primary", key="save_sap_mapping_btn"):
                    updated_map = {}
                    for _, row in edited_mapping_df.iterrows():
                        t_id = str(row['N.º']).strip()
                        m_act = str(row['Módulo Atualizado']).strip()
                        if m_act:
                            updated_map[t_id] = m_act
                            
                    current_map = load_melhorias_mapping()
                    current_map.update(updated_map)
                    if save_melhorias_mapping(current_map):
                        st.success("Mapeamento salvo com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhum chamado de Melhorias encontrado na planilha.")
                
            # Apply updated modules
            mapping = load_melhorias_mapping()
            def get_module_value(row):
                t_id = str(row.get('N.º', '')).strip().replace('.0', '')
                orig = str(row.get('Módulo', '')).strip()
                if 'melhoria' in orig.lower():
                    return mapping.get(t_id, orig)
                return orig
                
            if 'Módulo' in df_all.columns:
                df_all['Módulo_Visualizacao'] = df_all.apply(get_module_value, axis=1)
            else:
                df_all['Módulo_Visualizacao'] = "N/A"
                
            # Split Data
            df_ams_raw = df_all[df_all['Projeto'].astype(str).str.strip() == 'ZAMP - AMS - TICKETS'].copy()
            df_mel_raw = df_all[df_all['Projeto'].astype(str).str.strip() == 'ZAMP - BOLSÃO MELHORIAS'].copy()
            
            # Regra de Negócio: Desconsiderar chamados com Status 'Cancelado' e/ou 'Status Report' das contagens e gráficos
            def filter_invalid_statuses(df_proj):
                if df_proj.empty:
                    return df_proj
                status_col = 'Status (sem tempo decorrido)' if 'Status (sem tempo decorrido)' in df_proj.columns else df_proj.columns[0]
                status_series = df_proj[status_col].astype(str).str.strip().str.lower()
                # Filtra fora registros que correspondam exatamente ou contenham 'cancelado', 'cancelada' ou 'status report'
                mask_exclude = status_series.isin(['cancelado', 'cancelada', 'status report']) | status_series.str.contains('status report|cancelad', case=False, na=False)
                return df_proj[~mask_exclude].copy()
                
            df_ams_raw = filter_invalid_statuses(df_ams_raw)
            df_mel_raw = filter_invalid_statuses(df_mel_raw)
            
            # Time axis calculations
            months_12 = []
            for i in range(11, -1, -1):
                m = selected_vigente_dt - pd.DateOffset(months=i)
                # O contrato começou em janeiro de 2026. Não mostrar meses antes de 01/2026.
                if m >= pd.Timestamp('2026-01-01'):
                    months_12.append(m.strftime('%m/%Y'))
                
            months_3 = []
            for i in range(2, -1, -1):
                m = selected_vigente_dt - pd.DateOffset(months=i)
                months_3.append(m.strftime('%m/%Y'))
                
            current_month_str = selected_vigente_dt.strftime('%m/%Y')
            
            def get_project_metrics(df_proj):
                if df_proj.empty:
                    return {'mes_vigente_abertos': 0, 'backlog_pendentes': 0, 'encerrados_periodo': 0}
                    
                month_start = selected_vigente_dt
                month_end = selected_vigente_dt + pd.DateOffset(months=1) - pd.Timedelta(seconds=1)
                
                status_col = 'Status (sem tempo decorrido)' if 'Status (sem tempo decorrido)' in df_proj.columns else df_proj.columns[0]
                closed_statuses_list = ['encerrado', 'encerrada', 'cancelado', 'cancelada', 'resolvido', 'resolvida']
                
                is_closed = df_proj[status_col].astype(str).str.strip().str.lower().isin(closed_statuses_list)
                
                # Encerrados no periodo
                mask_encerrados = is_closed & (df_proj['_DataEncerra'] >= month_start) & (df_proj['_DataEncerra'] <= month_end)
                encerrados_periodo = mask_encerrados.sum()
                
                # Abertos no periodo (mes vigente)
                mask_abertos = (df_proj['_DataAbertura'] >= month_start) & (df_proj['_DataAbertura'] <= month_end)
                abertos_periodo = mask_abertos.sum()
                
                # Backlog (Pendentes) - chamados abertos antes do primeiro dia do mês vigente e atualmente ativos
                mask_backlog = (df_proj['_DataAbertura'] < month_start) & (~is_closed)
                backlog_pendentes = mask_backlog.sum()
                
                return {
                    'mes_vigente_abertos': abertos_periodo,
                    'backlog_pendentes': backlog_pendentes,
                    'encerrados_periodo': encerrados_periodo
                }
                
            metrics_ams = get_project_metrics(df_ams_raw)
            metrics_mel = get_project_metrics(df_mel_raw)
            
            # Generate Plotly charts
            def generate_charts(df_proj, proj_name, df_proj_live=None):
                if df_proj.empty:
                    return None, None, None, None, None
                    
                # Estilo padrão profissional para os gráficos
                def apply_professional_layout(fig, title, show_legend=False):
                    fig.update_layout(
                        title={
                            'text': f"<b>{title}</b>",
                            'y': 0.95,
                            'x': 0.5,
                            'xanchor': 'center',
                            'yanchor': 'top',
                            'font': dict(size=16, color='#2B4C7E', family="Arial, sans-serif")
                        },
                        font_family="Arial, sans-serif",
                        font_color="#5C6B73",
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        margin=dict(l=20, r=20, t=55, b=20),
                        showlegend=show_legend
                    )
                    if show_legend:
                        fig.update_layout(
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1,
                                title_text=""
                            )
                        )
                
                # 1. Abertos por Mês
                df_abertos_source = df_proj
                if proj_name == "AMS" and df_proj_live is not None and not df_proj_live.empty:
                    df_abertos_source = df_proj_live.copy()
                    
                df_abertos_source['MesAnoAbertura'] = df_abertos_source['_DataAbertura'].dt.strftime('%m/%Y')
                abertos_df = df_abertos_source[df_abertos_source['MesAnoAbertura'].isin(months_12)]
                abertos_counts = abertos_df['MesAnoAbertura'].value_counts().reindex(months_12, fill_value=0).reset_index()
                abertos_counts.columns = ['Mês', 'Quantidade']
                
                fig_abertos = px.bar(
                    abertos_counts, x='Mês', y='Quantidade', 
                    text_auto=True,
                    color_discrete_sequence=['#2B4C7E']  # Azul marinho profissional
                )
                apply_professional_layout(fig_abertos, f"Chamados Abertos por Mês - {proj_name}", show_legend=False)
                fig_abertos.update_xaxes(showgrid=False, linecolor='#5C6B73')
                fig_abertos.update_yaxes(showgrid=True, gridcolor='#E5E7EB', linecolor='#5C6B73')
                
                # 2. Encerrados por Mês
                df_proj['MesAnoEncerra'] = df_proj['_DataEncerra'].dt.strftime('%m/%Y')
                encerra_df = df_proj[df_proj['MesAnoEncerra'].isin(months_12)]
                encerra_counts = encerra_df['MesAnoEncerra'].value_counts().reindex(months_12, fill_value=0).reset_index()
                encerra_counts.columns = ['Mês', 'Quantidade']
                
                fig_encerra = px.bar(
                    encerra_counts, x='Mês', y='Quantidade', 
                    text_auto=True,
                    color_discrete_sequence=['#4A90E2']  # Azul aço profissional
                )
                apply_professional_layout(fig_encerra, f"Chamados Encerrados por Mês - {proj_name}", show_legend=False)
                fig_encerra.update_xaxes(showgrid=False, linecolor='#5C6B73')
                fig_encerra.update_yaxes(showgrid=True, gridcolor='#E5E7EB', linecolor='#5C6B73')
                
                # 3. Chamados Ativos (substituindo Módulos SAP)
                # Todos os chamados que não estejam encerrados/cancelados/resolvidos, independente de quando foram abertos
                closed_statuses_list = ['encerrado', 'encerrada', 'cancelado', 'cancelada', 'resolvido', 'resolvida']
                df_ativos = df_proj[~df_proj['Status (sem tempo decorrido)'].astype(str).str.strip().str.lower().isin(closed_statuses_list)].copy()
                
                ativos_counts = df_ativos['Status (sem tempo decorrido)'].value_counts().reset_index()
                ativos_counts.columns = ['Status', 'Quantidade']
                
                if ativos_counts.empty:
                    ativos_counts = pd.DataFrame([{'Status': 'Sem chamados ativos', 'Quantidade': 0}])
                    
                fig_ativos = px.bar(
                    ativos_counts, x='Status', y='Quantidade',
                    text_auto=True,
                    color_discrete_sequence=['#5C6B73']  # Cinza ardósia profissional
                )
                apply_professional_layout(fig_ativos, f"Chamados Ativos - {proj_name}", show_legend=False)
                fig_ativos.update_xaxes(showgrid=False, linecolor='#5C6B73')
                fig_ativos.update_yaxes(showgrid=True, gridcolor='#E5E7EB', linecolor='#5C6B73')
                
                # SLAs
                def calculate_sla_plotly(df_sla_base, col_sla_status):
                    if col_sla_status not in df_sla_base.columns:
                        return pd.DataFrame()
                        
                    df_sla_base['Prioridade_Mapped'] = df_sla_base['Prioridade'].apply(map_priority)
                    mask_vigente = (df_sla_base['MesAnoAbertura'] == current_month_str) | (df_sla_base['_DataEncerra'].dt.strftime('%m/%Y') == current_month_str)
                    df_sla_vigente = df_sla_base[mask_vigente].copy()
                    
                    def get_sla_status(val):
                        val_str = str(val).strip().upper()
                        if val_str in ['SIM', 'DENTRO', 'NO PRAZO', 'DENTRO DO SLA', 'T']:
                            return 'Sim'
                        elif val_str in ['NÃO', 'NAO', 'FORA', 'ATRASADO', 'FORA DO SLA', 'F']:
                            return 'Não'
                        elif val_str in ['N/A', '', 'NAN', 'NONE']:
                            return 'N/A'
                        else:
                            return 'Não'
                            
                    df_sla_vigente['SLA_Status'] = df_sla_vigente[col_sla_status].apply(get_sla_status)
                    df_sla_vigente = df_sla_vigente[df_sla_vigente['SLA_Status'] != 'N/A']
                    
                    grouped = df_sla_vigente.groupby(['Prioridade_Mapped', 'SLA_Status']).size().reset_index(name='Quantidade')
                    
                    all_priorities = ['Muito Alta', 'Alta', 'Média', 'Baixa']
                    statuses = ['Sim', 'Não']
                    
                    reindexed_rows = []
                    m_label = f"{current_month_str} (Vigente)"
                    for p in all_priorities:
                        for s in statuses:
                            match = grouped[(grouped['Prioridade_Mapped'] == p) & 
                                            (grouped['SLA_Status'] == s)]
                            qty = int(match.iloc[0]['Quantidade']) if not match.empty else 0
                            reindexed_rows.append({
                                'Mês': m_label,
                                'Prioridade': p,
                                'Status': s,
                                'Quantidade': qty,
                                'Texto': str(qty) if qty > 0 else ''
                            })
                    return pd.DataFrame(reindexed_rows)
                    
                # 4. SLA Primeira Resposta
                col_sla_resp = 'Status do SLA de resposta' if 'Status do SLA de resposta' in df_proj.columns else 'Resposta dentro do SLA'
                df_sla_resp = calculate_sla_plotly(df_proj, col_sla_resp)
                if not df_sla_resp.empty:
                    fig_sla_resp = px.bar(
                        df_sla_resp, x='Prioridade', y='Quantidade', color='Status',
                        barmode='stack', text='Texto',
                        category_orders={
                            'Prioridade': ['Muito Alta', 'Alta', 'Média', 'Baixa'],
                            'Status': ['Sim', 'Não']
                        },
                        color_discrete_map={
                            'Sim': '#2E7D32',  # Verde profissional
                            'Não': '#C62828'   # Vermelho profissional
                        }
                    )
                    fig_sla_resp.update_traces(textposition='inside', texttemplate='%{text}')
                    apply_professional_layout(fig_sla_resp, f"SLA Primeira Resposta - {proj_name}", show_legend=True)
                    fig_sla_resp.update_xaxes(showgrid=False, linecolor='#5C6B73')
                    fig_sla_resp.update_yaxes(showgrid=True, gridcolor='#E5E7EB', linecolor='#5C6B73')
                else:
                    fig_sla_resp = None
                    
                # 5. SLA Solução
                col_sla_sol = 'Status do SLA de Solução' if 'Status do SLA de Solução' in df_proj.columns else ('Solução dentro do SLA' if 'Solução dentro do SLA' in df_proj.columns else 'Soluço dentro do SLA')
                df_sla_sol = calculate_sla_plotly(df_proj, col_sla_sol)
                if not df_sla_sol.empty:
                    fig_sla_sol = px.bar(
                        df_sla_sol, x='Prioridade', y='Quantidade', color='Status',
                        barmode='stack', text='Texto',
                        category_orders={
                            'Prioridade': ['Muito Alta', 'Alta', 'Média', 'Baixa'],
                            'Status': ['Sim', 'Não']
                        },
                        color_discrete_map={
                            'Sim': '#2E7D32',  # Verde profissional
                            'Não': '#C62828'   # Vermelho profissional
                        }
                    )
                    fig_sla_sol.update_traces(textposition='inside', texttemplate='%{text}')
                    apply_professional_layout(fig_sla_sol, f"SLA Solução - {proj_name}", show_legend=True)
                    fig_sla_sol.update_xaxes(showgrid=False, linecolor='#5C6B73')
                    fig_sla_sol.update_yaxes(showgrid=True, gridcolor='#E5E7EB', linecolor='#5C6B73')
                else:
                    fig_sla_sol = None
                    
                return fig_abertos, fig_encerra, fig_ativos, fig_sla_resp, fig_sla_sol
                
            # Obter dados do Multidados ao vivo para AMS
            df_ams_live = None
            if 'df_multidados_live' in st.session_state and st.session_state['df_multidados_live'] is not None:
                df_live_full = st.session_state['df_multidados_live'].copy()
                df_live_full.columns = [c.strip() for c in df_live_full.columns]
                # Conversão de datas
                df_live_full['_DataAbertura'] = pd.to_datetime(df_live_full['Data de abertura'], errors='coerce', dayfirst=True)
                df_live_full['_DataEncerra'] = pd.to_datetime(df_live_full['Data de encerramento'], errors='coerce', dayfirst=True)
                # Filtrar apenas o projeto de AMS
                df_ams_live = df_live_full[df_live_full['Projeto'] == 'ZAMP - AMS - TICKETS'].copy()
                # Aplicar filtro de status cancelado / status report
                df_ams_live = filter_invalid_statuses(df_ams_live)
                
            fig_ams_ab, fig_ams_enc, fig_ams_ativos, fig_ams_s1, fig_ams_s2 = generate_charts(df_ams_raw, "AMS", df_proj_live=df_ams_live)
            fig_mel_ab, fig_mel_enc, fig_mel_ativos, fig_mel_s1, fig_mel_s2 = generate_charts(df_mel_raw, "Melhorias")
            
            # Render interactive dashboards on screen
            st.divider()
            st.subheader("📊 Dashboards Interativos (Visualização em Tela)")
            
            tab_dash_ams, tab_dash_mel = st.tabs(["📈 ZAMP - AMS - TICKETS", "🚀 ZAMP - BOLSÃO DE MELHORIAS"])
            
            with tab_dash_ams:
                st.markdown("### Métricas do Mês Vigente - AMS")
                col_am1, col_am2, col_am3 = st.columns(3)
                col_am1.metric(f"📌 {nome_do_mes} (Abertos)", metrics_ams['mes_vigente_abertos'])
                col_am2.metric("⏳ Backlog (Pendentes)", metrics_ams['backlog_pendentes'])
                col_am3.metric("✅ Encerrados no Período", metrics_ams['encerrados_periodo'])
                
                st.markdown("### Volumetria e SLA - AMS")
                col_c1, col_c2 = st.columns(2)
                with col_c1:
                    if fig_ams_ab: st.plotly_chart(fig_ams_ab, use_container_width=True)
                with col_c2:
                    if fig_ams_enc: st.plotly_chart(fig_ams_enc, use_container_width=True)
                    
                col_c3, col_c4 = st.columns(2)
                with col_c3:
                    if fig_ams_s1: st.plotly_chart(fig_ams_s1, use_container_width=True)
                with col_c4:
                    if fig_ams_s2: st.plotly_chart(fig_ams_s2, use_container_width=True)
                    
                if fig_ams_ativos: st.plotly_chart(fig_ams_ativos, use_container_width=True)
                
            with tab_dash_mel:
                st.markdown("### Métricas do Mês Vigente - Melhorias")
                col_me1, col_me2, col_me3 = st.columns(3)
                col_me1.metric(f"📌 {nome_do_mes} (Abertos)", metrics_mel['mes_vigente_abertos'])
                col_me2.metric("⏳ Backlog (Pendentes)", metrics_mel['backlog_pendentes'])
                col_me3.metric("✅ Encerrados no Período", metrics_mel['encerrados_periodo'])
                
                st.markdown("### Volumetria e SLA - Melhorias")
                col_cx1, col_cx2 = st.columns(2)
                with col_cx1:
                    if fig_mel_ab: st.plotly_chart(fig_mel_ab, use_container_width=True)
                with col_cx2:
                    if fig_mel_enc: st.plotly_chart(fig_mel_enc, use_container_width=True)
                    
                col_cx3, col_cx4 = st.columns(2)
                with col_cx3:
                    if fig_mel_s1: st.plotly_chart(fig_mel_s1, use_container_width=True)
                with col_cx4:
                    if fig_mel_s2: st.plotly_chart(fig_mel_s2, use_container_width=True)
                    
                if fig_mel_ativos: st.plotly_chart(fig_mel_ativos, use_container_width=True)
                
            # Slide Previewer and Manual Editor
            st.divider()
            st.subheader("🖥️ Revisão e Edição dos Slides da Apresentação")
            st.markdown("""
            Revise o conteúdo que será incluído nos slides da apresentação nos campos à esquerda. Se desejar, faça qualquer **alteração manual** de textos ou tópicos. 
            A visualização à direita reflete suas atualizações em tempo real.
            """)
            
            if 'slide_contents' not in st.session_state:
                st.session_state['slide_contents'] = {
                    'capa_titulo': "STATUS REPORT",
                    'capa_data': selected_vigente_dt.strftime('%d/%m/%Y'),
                    'nome_do_mes': nome_do_mes,
                    'ams_overview_title': "ZAMP - BOLSÃO DE TICKETS",
                    'ams_overview_bullets': "- Acompanhamento mensal do bolsão de incidentes e requisições (AMS).\n- Foco na redução do backlog de meses anteriores.\n- Desempenho geral estável dentro das metas estabelecidas.",
                    'ams_vol_title': "ZAMP - BOLSÃO DE TICKETS - Volumetria",
                    'ams_vol_text': "Análise comparativa de volumetria de chamados abertos e encerrados a partir de janeiro de 2026.",
                    'ams_sla_title': "ZAMP - BOLSÃO DE TICKETS - Indicadores de SLA",
                    'ams_sla_text': "Quantidade de chamados atingidos e violados nos SLAs de Primeira Resposta e Solução por prioridade no mês vigente.",
                    'mel_overview_title': "ZAMP - BOLSÃO DE MELHORIAS",
                    'mel_overview_bullets': "- Acompanhamento de demandas evolutivas e melhorias de processo.\n- Planejamento de entregas conforme prioridades de negócio.\n- Utilização otimizada das horas do bolsão.",
                    'mel_vol_title': "ZAMP - BOLSÃO DE MELHORIAS - Volumetria",
                    'mel_vol_text': "Evolução mensal de melhorias abertas e encerradas a partir de janeiro de 2026.",
                    'mel_sla_title': "ZAMP - BOLSÃO DE MELHORIAS - Indicadores de SLA",
                    'mel_sla_text': "Quantidade de chamados atingidos e violados nos SLAs de Primeira Resposta e Solução para o bolsão de melhorias no mês vigente.",
                }
                
            col_ed, col_pr = st.columns([1, 1.2])
            
            with col_ed:
                st.markdown("#### 📝 Editor de Slides")
                
                with st.expander("Slide 1: Capa"):
                    st.session_state['slide_contents']['capa_titulo'] = st.text_input("Título da Capa:", value=st.session_state['slide_contents']['capa_titulo'], key="capa_title_input")
                    st.session_state['slide_contents']['capa_data'] = st.text_input("Data da Capa:", value=st.session_state['slide_contents']['capa_data'], key="capa_date_input")
                    
                with st.expander("Slide 2: AMS - Visão Geral"):
                    st.session_state['slide_contents']['ams_overview_title'] = st.text_input("Título do Slide 2:", value=st.session_state['slide_contents']['ams_overview_title'], key="s2_title")
                    st.session_state['slide_contents']['ams_overview_bullets'] = st.text_area("Tópicos (um por linha):", value=st.session_state['slide_contents']['ams_overview_bullets'], height=120, key="s2_bullets")
                    
                with st.expander("Slide 3: AMS - Volumetria"):
                    st.session_state['slide_contents']['ams_vol_title'] = st.text_input("Título do Slide 3:", value=st.session_state['slide_contents']['ams_vol_title'], key="s3_title")
                    st.session_state['slide_contents']['ams_vol_text'] = st.text_area("Texto Descritivo:", value=st.session_state['slide_contents']['ams_vol_text'], height=80, key="s3_text")
                    
                with st.expander("Slide 5: Melhorias - Visão Geral"):
                    st.session_state['slide_contents']['mel_overview_title'] = st.text_input("Título do Slide 5:", value=st.session_state['slide_contents']['mel_overview_title'], key="s5_title")
                    st.session_state['slide_contents']['mel_overview_bullets'] = st.text_area("Tópicos (um por linha):", value=st.session_state['slide_contents']['mel_overview_bullets'], height=120, key="s5_bullets")
                    
                with st.expander("Slide 6: Melhorias - Volumetria"):
                    st.session_state['slide_contents']['mel_vol_title'] = st.text_input("Título do Slide 6:", value=st.session_state['slide_contents']['mel_vol_title'], key="s6_title")
                    st.session_state['slide_contents']['mel_vol_text'] = st.text_area("Texto Descritivo:", value=st.session_state['slide_contents']['mel_vol_text'], height=80, key="s6_text")
                    
                with st.expander("Slide 7: Melhorias - SLAs"):
                    st.session_state['slide_contents']['mel_sla_title'] = st.text_input("Título do Slide 7:", value=st.session_state['slide_contents']['mel_sla_title'], key="s7_title")
                    st.session_state['slide_contents']['mel_sla_text'] = st.text_area("Texto Descritivo:", value=st.session_state['slide_contents']['mel_sla_text'], height=80, key="s7_text")
                    
            with col_pr:
                st.markdown("#### 📺 Visualização do Slide")
                slide_to_preview = st.selectbox("Selecione o slide para visualizar na tela:", [
                    "Slide 1: Capa",
                    "Slide 2: AMS - Visão Geral",
                    "Slide 3: AMS - Volumetria",
                    "Slide 4: AMS - Chamados Ativos",
                    "Slide 5: AMS - SLAs",
                    "Slide 6: Melhorias - Visão Geral",
                    "Slide 7: Melhorias - Volumetria",
                    "Slide 8: Melhorias - Chamados Ativos",
                    "Slide 9: Melhorias - SLAs",
                    "Slide 10: Agradecimento"
                ], key="slide_preview_selector")
                
                bg_color = "#F5EBDC"
                border_color = "#5C3527"
                text_color = "#5C3527"
                accent_color = "#D62300"
                
                slide_html = f"""
                <div style="
                    background-color: {bg_color}; 
                    border: 4px solid {border_color}; 
                    border-radius: 10px; 
                    padding: 25px; 
                    aspect-ratio: 16/9; 
                    color: {text_color}; 
                    font-family: 'Arial', sans-serif; 
                    position: relative; 
                    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
                    margin-bottom: 15px;
                ">
                    <div style="font-size: 20px; font-weight: bold; border-bottom: 2px solid {accent_color}; padding-bottom: 5px; margin-bottom: 15px; display: flex; justify-content: space-between;">
                        <span>{{SLIDE_TITLE}}</span>
                        <span style="font-size: 14px; opacity: 0.8; font-weight: normal;">ZAMP Status Report</span>
                    </div>
                    
                    <div style="font-size: 14px; line-height: 1.4; height: 75%; overflow-y: auto;">
                        {{SLIDE_CONTENT}}
                    </div>
                    
                    <div style="position: absolute; bottom: 10px; left: 25px; right: 25px; display: flex; justify-content: space-between; font-size: 10px; opacity: 0.8; border-top: 1px solid rgba(92,53,39,0.2); padding-top: 5px;">
                        <span>Ábaco 360° | Consultoria SAP</span>
                        <span>{{SLIDE_NUM}}</span>
                    </div>
                </div>
                """
                
                if slide_to_preview == "Slide 1: Capa":
                    title = st.session_state['slide_contents']['capa_titulo'].replace('\n', '<br>')
                    date = st.session_state['slide_contents']['capa_data']
                    content = f"""
                    <div style="display: flex; flex-direction: column; justify-content: center; align-items: center; height: 100%; text-align: center; margin-top: 10px;">
                        <h1 style="font-size: 32px; color: {accent_color}; margin: 0 0 10px 0;">{title}</h1>
                        <h2 style="font-size: 18px; font-weight: normal; margin: 0 0 20px 0;">Ábaco 360° - Gestão de Operação ZAMP</h2>
                        <div style="background-color: {border_color}; color: white; padding: 5px 15px; border-radius: 5px; font-size: 14px; font-weight: bold;">
                            {date}
                        </div>
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", "CAPA").replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 1"), unsafe_allow_html=True)
                    
                elif slide_to_preview == "Slide 2: AMS - Visão Geral":
                    title = st.session_state['slide_contents']['ams_overview_title']
                    bullets = st.session_state['slide_contents']['ams_overview_bullets'].replace('\n', '<br>')
                    
                    content = f"""
                    <div style="display: flex; gap: 20px; height: 100%;">
                        <div style="flex: 1; display: flex; flex-direction: column; gap: 10px; justify-content: center;">
                            <div style="background-color: white; border: 1px solid {border_color}; padding: 10px; border-radius: 5px; text-align: center;">
                                <div style="font-size: 11px; opacity: 0.8;">{nome_do_mes} (Abertos)</div>
                                <div style="font-size: 22px; font-weight: bold; color: {accent_color};">{metrics_ams['mes_vigente_abertos']}</div>
                            </div>
                            <div style="background-color: white; border: 1px solid {border_color}; padding: 10px; border-radius: 5px; text-align: center;">
                                <div style="font-size: 11px; opacity: 0.8;">Backlog (Pendentes)</div>
                                <div style="font-size: 22px; font-weight: bold; color: {accent_color};">{metrics_ams['backlog_pendentes']}</div>
                            </div>
                            <div style="background-color: white; border: 1px solid {border_color}; padding: 10px; border-radius: 5px; text-align: center;">
                                <div style="font-size: 11px; opacity: 0.8;">Encerrados no Período</div>
                                <div style="font-size: 22px; font-weight: bold; color: {accent_color};">{metrics_ams['encerrados_periodo']}</div>
                            </div>
                        </div>
                        <div style="flex: 1.5; display: flex; flex-direction: column; justify-content: center; padding-left: 10px; border-left: 1px solid rgba(92,53,39,0.2);">
                            <p style="font-size: 13px; font-weight: bold; margin-top: 0;">Destaques da Operação:</p>
                            <p style="font-size: 12px;">{bullets}</p>
                        </div>
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", title).replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 2"), unsafe_allow_html=True)
                    
                elif slide_to_preview == "Slide 3: AMS - Volumetria":
                    title = st.session_state['slide_contents']['ams_vol_title']
                    desc = st.session_state['slide_contents']['ams_vol_text']
                    content = f"""
                    <p style="font-size: 12px; margin-top: 0; margin-bottom: 10px; font-style: italic;">{desc}</p>
                    <div style="display: flex; justify-content: center; align-items: center; background-color: white; border: 1px dashed {border_color}; height: 75%; border-radius: 5px; color: {border_color}; font-weight: bold;">
                        [Gráficos de Volumetria (Abertos e Encerrados) do Projeto AMS]
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", title).replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 3"), unsafe_allow_html=True)
                    st.info("📊 Pré-visualização real dos gráficos do Slide 3:")
                    col_p1, col_p2 = st.columns(2)
                    with col_p1:
                        if fig_ams_ab: st.plotly_chart(fig_ams_ab, use_container_width=True)
                    with col_p2:
                        if fig_ams_enc: st.plotly_chart(fig_ams_enc, use_container_width=True)
                        
                elif slide_to_preview == "Slide 4: AMS - Chamados Ativos":
                    content = f"""
                    <p style="font-size: 12px; margin-top: 0; margin-bottom: 10px; font-style: italic;">Visualização de todos os chamados ativos do Projeto AMS distribuídos por status.</p>
                    <div style="display: flex; justify-content: center; align-items: center; background-color: white; border: 1px dashed {border_color}; height: 75%; border-radius: 5px; color: {border_color}; font-weight: bold;">
                        [Gráfico de Chamados Ativos - AMS]
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", "ZAMP - BOLSÃO DE TICKETS - Chamados Ativos").replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 4"), unsafe_allow_html=True)
                    st.info("📊 Pré-visualização real do gráfico de Chamados Ativos do Slide 4:")
                    if fig_ams_ativos: st.plotly_chart(fig_ams_ativos, use_container_width=True)
                    
                elif slide_to_preview == "Slide 5: AMS - SLAs":
                    title = st.session_state['slide_contents']['ams_sla_title']
                    desc = st.session_state['slide_contents']['ams_sla_text']
                    content = f"""
                    <p style="font-size: 12px; margin-top: 0; margin-bottom: 10px; font-style: italic;">{desc}</p>
                    <div style="display: flex; justify-content: center; align-items: center; background-color: white; border: 1px dashed {border_color}; height: 75%; border-radius: 5px; color: {border_color}; font-weight: bold;">
                        [Gráficos de SLAs (Primeira Resposta e Solução) do Projeto AMS]
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", title).replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 5"), unsafe_allow_html=True)
                    st.info("📊 Pré-visualização real dos gráficos de SLA do Slide 5:")
                    col_p1, col_p2 = st.columns(2)
                    with col_p1:
                        if fig_ams_s1: st.plotly_chart(fig_ams_s1, use_container_width=True)
                    with col_p2:
                        if fig_ams_s2: st.plotly_chart(fig_ams_s2, use_container_width=True)
                        
                elif slide_to_preview == "Slide 6: Melhorias - Visão Geral":
                    title = st.session_state['slide_contents']['mel_overview_title']
                    bullets = st.session_state['slide_contents']['mel_overview_bullets'].replace('\n', '<br>')
                    
                    content = f"""
                    <div style="display: flex; gap: 20px; height: 100%;">
                        <div style="flex: 1; display: flex; flex-direction: column; gap: 10px; justify-content: center;">
                            <div style="background-color: white; border: 1px solid {border_color}; padding: 10px; border-radius: 5px; text-align: center;">
                                <div style="font-size: 11px; opacity: 0.8;">{nome_do_mes} (Abertos)</div>
                                <div style="font-size: 22px; font-weight: bold; color: {accent_color};">{metrics_mel['mes_vigente_abertos']}</div>
                            </div>
                            <div style="background-color: white; border: 1px solid {border_color}; padding: 10px; border-radius: 5px; text-align: center;">
                                <div style="font-size: 11px; opacity: 0.8;">Backlog (Pendentes)</div>
                                <div style="font-size: 22px; font-weight: bold; color: {accent_color};">{metrics_mel['backlog_pendentes']}</div>
                            </div>
                            <div style="background-color: white; border: 1px solid {border_color}; padding: 10px; border-radius: 5px; text-align: center;">
                                <div style="font-size: 11px; opacity: 0.8;">Encerrados no Período</div>
                                <div style="font-size: 22px; font-weight: bold; color: {accent_color};">{metrics_mel['encerrados_periodo']}</div>
                            </div>
                        </div>
                        <div style="flex: 1.5; display: flex; flex-direction: column; justify-content: center; padding-left: 10px; border-left: 1px solid rgba(92,53,39,0.2);">
                            <p style="font-size: 13px; font-weight: bold; margin-top: 0;">Destaques da Operação (Melhorias):</p>
                            <p style="font-size: 12px;">{bullets}</p>
                        </div>
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", title).replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 6"), unsafe_allow_html=True)
                    
                elif slide_to_preview == "Slide 7: Melhorias - Volumetria":
                    title = st.session_state['slide_contents']['mel_vol_title']
                    desc = st.session_state['slide_contents']['mel_vol_text']
                    content = f"""
                    <p style="font-size: 12px; margin-top: 0; margin-bottom: 10px; font-style: italic;">{desc}</p>
                    <div style="display: flex; justify-content: center; align-items: center; background-color: white; border: 1px dashed {border_color}; height: 75%; border-radius: 5px; color: {border_color}; font-weight: bold;">
                        [Gráficos de Volumetria (Abertos e Encerrados) de Melhorias]
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", title).replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 7"), unsafe_allow_html=True)
                    st.info("📊 Pré-visualização real dos gráficos do Slide 7:")
                    col_p1, col_p2 = st.columns(2)
                    with col_p1:
                        if fig_mel_ab: st.plotly_chart(fig_mel_ab, use_container_width=True)
                    with col_p2:
                        if fig_mel_enc: st.plotly_chart(fig_mel_enc, use_container_width=True)
                        
                elif slide_to_preview == "Slide 8: Melhorias - Chamados Ativos":
                    content = f"""
                    <p style="font-size: 12px; margin-top: 0; margin-bottom: 10px; font-style: italic;">Visualização de todos os chamados ativos do Projeto Melhorias distribuídos por status.</p>
                    <div style="display: flex; justify-content: center; align-items: center; background-color: white; border: 1px dashed {border_color}; height: 75%; border-radius: 5px; color: {border_color}; font-weight: bold;">
                        [Gráfico de Chamados Ativos - Melhorias]
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", "ZAMP - BOLSÃO DE MELHORIAS - Chamados Ativos").replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 8"), unsafe_allow_html=True)
                    st.info("📊 Pré-visualização real do gráfico de Chamados Ativos do Slide 8:")
                    if fig_mel_ativos: st.plotly_chart(fig_mel_ativos, use_container_width=True)
                    
                elif slide_to_preview == "Slide 9: Melhorias - SLAs":
                    title = st.session_state['slide_contents']['mel_sla_title']
                    desc = st.session_state['slide_contents']['mel_sla_text']
                    content = f"""
                    <p style="font-size: 12px; margin-top: 0; margin-bottom: 10px; font-style: italic;">{desc}</p>
                    <div style="display: flex; justify-content: center; align-items: center; background-color: white; border: 1px dashed {border_color}; height: 75%; border-radius: 5px; color: {border_color}; font-weight: bold;">
                        [Gráficos de SLAs (Primeira Resposta e Solução) de Melhorias]
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", title).replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 9"), unsafe_allow_html=True)
                    st.info("📊 Pré-visualização real dos gráficos de SLA do Slide 9:")
                    col_p1, col_p2 = st.columns(2)
                    with col_p1:
                        if fig_mel_s1: st.plotly_chart(fig_mel_s1, use_container_width=True)
                    with col_p2:
                        if fig_mel_s2: st.plotly_chart(fig_mel_s2, use_container_width=True)
                        
                elif slide_to_preview == "Slide 10: Agradecimento":
                    content = f"""
                    <div style="display: flex; justify-content: center; align-items: center; height: 100%; text-align: center;">
                        <h1 style="font-size: 40px; color: {accent_color}; margin: 0 0 10px 0;">Obrigada!</h1>
                    </div>
                    """
                    st.markdown(slide_html.replace("{SLIDE_TITLE}", "ENCERRAMENTO").replace("{SLIDE_CONTENT}", content).replace("{SLIDE_NUM}", "Slide 10"), unsafe_allow_html=True)
                    
            # Export and Download Buttons
            st.divider()
            st.subheader("📥 Gerar e Baixar Apresentações")
            st.markdown("""
            Clique nos botões abaixo para gerar e fazer o download da sua apresentação nos formatos PowerPoint e PDF. 
            Todas as suas **edições manuais** feitas nos campos acima serão incluídas automaticamente nos arquivos.
            """)
            
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                if st.button("📥 Gerar e Baixar PowerPoint (PPTX)", type="primary", key="gen_pptx_btn"):
                    with st.spinner("Gerando arquivo PowerPoint (.pptx)..."):
                        try:
                            prs_out = generate_pptx_presentation(
                                df_ams_raw, df_mel_raw, metrics_ams, metrics_mel,
                                fig_ams_ab, fig_ams_enc, fig_ams_ativos, fig_ams_s1, fig_ams_s2,
                                fig_mel_ab, fig_mel_enc, fig_mel_ativos, fig_mel_s1, fig_mel_s2,
                                st.session_state['slide_contents']
                            )
                            pptx_buffer = io.BytesIO()
                            prs_out.save(pptx_buffer)
                            pptx_buffer.seek(0)
                            
                            st.success("PowerPoint gerado com sucesso!")
                            st.download_button(
                                label="💾 Fazer Download do PowerPoint (.pptx)",
                                data=pptx_buffer.getvalue(),
                                file_name=f"Status_Report_ZAMP_{selected_month_str.replace('/', '_')}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                key="pptx_download_action"
                            )
                        except Exception as e:
                            st.error(f"Erro ao gerar PowerPoint: {e}")
                            
            with col_d2:
                if st.button("📥 Gerar e Baixar PDF", type="primary", key="gen_pdf_btn"):
                    with st.spinner("Gerando arquivo PDF (.pdf)..."):
                        try:
                            pdf_out = generate_pdf_presentation(
                                df_ams_raw, df_mel_raw, metrics_ams, metrics_mel,
                                fig_ams_ab, fig_ams_enc, fig_ams_ativos, fig_ams_s1, fig_ams_s2,
                                fig_mel_ab, fig_mel_enc, fig_mel_ativos, fig_mel_s1, fig_mel_s2,
                                st.session_state['slide_contents']
                            )
                            
                            # Standard FPDF output in memory
                            pdf_str = pdf_out.output(dest='S')
                            # FPDF returns a string in some old versions, or bytes in others
                            # Let's ensure it is in bytes
                            if isinstance(pdf_str, str):
                                pdf_bytes = pdf_str.encode('latin1')
                            else:
                                pdf_bytes = pdf_str
                                
                            st.success("PDF gerado com sucesso!")
                            st.download_button(
                                label="💾 Fazer Download do PDF (.pdf)",
                                data=pdf_bytes,
                                file_name=f"Status_Report_ZAMP_{selected_month_str.replace('/', '_')}.pdf",
                                mime="application/pdf",
                                key="pdf_download_action"
                            )
                        except Exception as e:
                            st.error(f"Erro ao gerar PDF: {e}")
                            
    else:
        st.info("👋 Por favor, faça o upload do arquivo Excel consolidado (Status Report) no campo acima para começar.")

with tab3:
    st.header("Gerador de Ata de Reunião")
    st.markdown("Preencha as informações abaixo e faça o upload dos arquivos solicitados para gerar a Ata de Reunião.")
    
    col_ata1, col_ata2 = st.columns(2)
    with col_ata1:
        data_reuniao = st.date_input("Data da Reunião", datetime.date.today())
        participantes = st.text_input("Participantes", placeholder="Ex: Monique, Wesley e Henrique.")
        proximos_passos = st.text_area("4. Próximos Passos / Ações", placeholder="Descreva as próximas ações...")
    with col_ata2:
        temas_reuniao = st.text_area("1. Principais temas da reunião", placeholder="• Status dos chamados;\n• Alguns chamados de ARIBA...")
        
    st.markdown("### Upload de Arquivos")
    col_file1, col_file2 = st.columns(2)
    with col_file1:
        file_todos_chamados = st.file_uploader("1️⃣ Upload - Status Report - Todos os chamados", type=["xlsx", "csv"], key="file_ata_todos")
    with col_file2:
        file_txt = st.file_uploader("2️⃣ Upload - Status Report - Ata de Reunião (.txt)", type=["txt"], key="file_ata_txt")
        
    if st.button("Gerar Ata de Reunião", type="primary"):
        if not file_todos_chamados or not file_txt:
            st.error("Por favor, faça o upload dos dois arquivos.")
        else:
            with st.spinner("Processando Ata de Reunião..."):
                df_todos = load_data(file_todos_chamados, "Todos os Chamados")
                txt_content = file_txt.read().decode("utf-8")
                
                if df_todos is not None:
                    MESES_PT = {
                        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
                        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
                        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
                    }
                    mes_nome = MESES_PT[data_reuniao.month]
                    
                    def get_metrics_ata(df, project_name, base_date):
                        if 'Projeto' in df.columns:
                            df_proj = df[df['Projeto'].astype(str).str.strip().str.upper() == project_name.upper()]
                        else:
                            df_proj = df
                            
                        abertos_mes = 0
                        if 'Data de abertura' in df_proj.columns:
                            dates = pd.to_datetime(df_proj['Data de abertura'], errors='coerce', dayfirst=True)
                            mask = (dates.dt.month == base_date.month) & (dates.dt.year == base_date.year)
                            abertos_mes = mask.sum()
                            
                        encerrados = 0
                        aguardando = 0
                        em_atendimento = 0
                        
                        if 'Status (sem tempo decorrido)' in df_proj.columns:
                            status_col = df_proj['Status (sem tempo decorrido)'].astype(str).str.strip().str.lower()
                            encerrados = (status_col == 'encerrado').sum()
                            aguardando = (status_col == 'aguardando feedback do cliente').sum()
                            em_atendimento = (status_col == 'em atendimento (consultor)').sum()
                            
                        return abertos_mes, encerrados, aguardando, em_atendimento
                    
                    ams_abertos, ams_encerrados, ams_aguardando, ams_em_atendimento = get_metrics_ata(df_todos, "ZAMP - AMS - TICKETS", data_reuniao)
                    melh_abertos, melh_encerrados, melh_aguardando, melh_em_atendimento = get_metrics_ata(df_todos, "ZAMP - BOLSÃO MELHORIAS", data_reuniao)
                    
                    ams_text = []
                    melhorias_text = []
                    
                    current_block = []
                    current_project = None
                    
                    lines = txt_content.split('\n')
                    
                    for line in lines:
                        line_clean = line.rstrip('\r\n')
                        ticket_match = re.match(r'^[\s\-*•oO_+]*?(INC\d+|RITM\d+)(.*)', line_clean)
                        
                        if ticket_match:
                            if current_block:
                                block_str = "  \n".join(current_block)  # Two spaces for markdown newline
                                if current_project == "ZAMP - AMS - TICKETS":
                                    ams_text.append(block_str)
                                elif current_project == "ZAMP - BOLSÃO MELHORIAS":
                                    melhorias_text.append(block_str)
                                    
                            ticket_num = ticket_match.group(1)
                            rest_of_line = ticket_match.group(2)
                            
                            # Separar se houver datas do formato DD/MM - na mesma linha
                            parts = re.split(r'(\s*(?:o\s+)?\d{2}/\d{2}\s*-)', rest_of_line)
                            title = parts[0]
                            updates = []
                            for i in range(1, len(parts), 2):
                                d_marker = parts[i].strip()
                                u_text = parts[i+1].strip() if (i+1 < len(parts)) else ""
                                updates.append(f"{d_marker} {u_text}")
                            
                            # Apenas a linha do ticket fica em negrito
                            current_block = [f"**{ticket_num}{title}**"] + updates
                            
                            if 'ServiceNow' in df_todos.columns:
                                m_proj = df_todos[df_todos['ServiceNow'].astype(str).str.contains(ticket_num, na=False)]['Projeto']
                                current_project = m_proj.iloc[0].strip().upper() if not m_proj.empty else None
                            else:
                                current_project = None
                        else:
                            if current_block:
                                parts = re.split(r'(\s*(?:o\s+)?\d{2}/\d{2}\s*-)', line_clean)
                                header = parts[0].strip()
                                if header:
                                    current_block.append(header)
                                for i in range(1, len(parts), 2):
                                    d_marker = parts[i].strip()
                                    u_text = parts[i+1].strip() if (i+1 < len(parts)) else ""
                                    current_block.append(f"{d_marker} {u_text}")
                                
                    if current_block:
                        block_str = "  \n".join(current_block)
                        if current_project == "ZAMP - AMS - TICKETS":
                            ams_text.append(block_str)
                        elif current_project == "ZAMP - BOLSÃO MELHORIAS":
                            melhorias_text.append(block_str)

                    data_str = data_reuniao.strftime("%d/%m/%Y")
                    
                    ata_final = f"**ATA DE REUNIÃO – Status Report – {data_str}**\n\n"
                    ata_final += f"**Participantes:** {participantes if participantes else ''}\n"
                    ata_final += "________________________________________\n\n"
                    ata_final += "**1. Principais temas da reunião**\n"
                    if temas_reuniao:
                        ata_final += f"{temas_reuniao}\n\n"
                    else:
                        ata_final += "\n"
                        
                    ata_final += "________________________________________\n\n"
                    ata_final += f"**2. Status dos Chamados (Resumo) – Bolsão de Tickets - {mes_nome}**\n"
                    ata_final += f"•\tTotal de chamados abertos: {ams_abertos}\n"
                    ata_final += f"•\tEncerrados: {ams_encerrados}\n"
                    ata_final += f"•\tAguardando feedback do cliente: {ams_aguardando}\n"
                    ata_final += f"•\tEm atendimento: {ams_em_atendimento}\n\n"
                    
                    separator_line = "-" * 120
                    
                    if ams_text:
                        ata_final += f"\n{separator_line}\n\n".join(ams_text) + "\n"
                        
                    ata_final += "\n________________________________________\n\n"
                    ata_final += f"**3. Status dos Chamados (Resumo) – Bolsão de Melhorias - {mes_nome}**\n"
                    ata_final += f"•\tTotal de chamados abertos: {melh_abertos}\n"
                    ata_final += f"•\tEncerrados: {melh_encerrados}\n"
                    ata_final += f"•\tAguardando feedback do cliente: {melh_aguardando}\n"
                    ata_final += f"•\tEm atendimento: {melh_em_atendimento}\n\n"
                    
                    if melhorias_text:
                        ata_final += f"\n{separator_line}\n\n".join(melhorias_text) + "\n"

                    ata_final += "\n________________________________________\n\n"
                    ata_final += "**4. Próximos Passos / Ações**\n"
                    if proximos_passos:
                        ata_final += f"{proximos_passos}\n\n"

                    st.success("Ata de Reunião gerada com sucesso!")
                    st.markdown("### Visualização (Recomendada para copiar p/ E-mail):")
                    st.markdown(ata_final)
                    st.text_area("Texto Puro (Markdown):", value=ata_final, height=400)
                    
                    st.download_button(
                        label="📄 Baixar Ata em TXT",
                        data=ata_final.encode('utf-8'),
                        file_name=f"Ata_de_Reuniao_{data_reuniao.strftime('%Y%m%d')}.txt",
                        mime="text/plain"
                    )

