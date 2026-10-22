import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import re
import os
from io import BytesIO
from datetime import timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from google.oauth2.service_account import Credentials
import requests
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y CSS PREMIUM
# ==========================================
st.set_page_config(page_title="Cotizador SCTR Digital | YQ Corredores", layout="centered", page_icon="🛡️")

st.markdown("""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"] { display: none !important; visibility: hidden !important; }
        footer, .stApp > footer, [data-testid="stBottom"], [data-testid="stBottomBar"], [data-testid="stViewerBadge"] { display: none !important; visibility: hidden !important; }
        .block-container { padding-bottom: 0rem !important; }
        div.stButton > button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; font-size: 16px !important; font-weight: 700 !important; }
        div.stButton > button:hover { background-color: #022266 !important; border-color: #022266 !important; }
        div[data-testid="stDownloadButton"] button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; width: 100% !important; font-size: 16px !important; font-weight: 700 !important; }
    </style>
""", unsafe_allow_html=True)

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ImageRL
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
except ImportError:
    st.error("❌ Falta la librería 'reportlab'. Asegúrate de incluirla en tu requirements.txt")
    st.stop()

# --- CONSTANTES DE ACCESO ---
CODIGO_ADMIN = "ADMIN2026"
CODIGOS_ASESORES = ["ASE01", "ASE02", "ASE03", "VENTAS2026"]
RMV_VIGENTE = 1230.0

def obtener_hora_peru():
    return datetime.datetime.utcnow() - timedelta(hours=5)

# ==========================================
# 2. FUNCIONES GOOGLE SHEETS Y ZOHO CRM
# ==========================================
def get_gspread_client():
    try:
        if "gcp_service_account" not in st.secrets: return None
        creds_dict = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(credentials)
    except Exception: return None

def guardar_en_sheets(datos_fila):
    try:
        client = get_gspread_client()
        if client: client.open("historial_cotizador_sctr").sheet1.append_row(datos_fila)
    except Exception: pass

def descargar_historial_sheets():
    try:
        client = get_gspread_client()
        if client: return pd.DataFrame(client.open("historial_cotizador_sctr").sheet1.get_all_records())
    except Exception: return None

def enviar_notificacion_sctr(cliente, ruc, celular, email, rubro, actividad, trabajadores, planilla, mejor_cia, mejor_total):
    try:
        if "ZOHO_REFRESH_TOKEN" not in st.secrets: return False
        url_auth = "https://accounts.zoho.com/oauth/v2/token"
        datos_auth = {"refresh_token": st.secrets["ZOHO_REFRESH_TOKEN"], "client_id": st.secrets["ZOHO_CLIENT_ID"], "client_secret": st.secrets["ZOHO_CLIENT_SECRET"], "grant_type": "refresh_token"}
        access_token = requests.post(url_auth, data=datos_auth).json().get("access_token")
        
        if access_token:
            desc = f"Rubro: {rubro} | Actividad: {actividad} | Trabajadores: {trabajadores} | Planilla: S/ {planilla:,.2f} | Mejor Opción: {mejor_cia} (S/ {mejor_total:,.2f})"
            payload = {"data": [{"Company": cliente, "Last_Name": "Contacto Comercial", "RUC": str(ruc), "Email": email, "Mobile": str(celular), "Lead_Source": "Cotizador SCTR Web", "Seguro_de_Inter_s": "SCTR", "Description": desc}]}
            requests.post("https://www.zohoapis.com/crm/v2/Leads", headers={"Authorization": f"Zoho-oauthtoken {access_token}"}, json=payload)
    except Exception: pass

# ==========================================
# 3. MOTORES DE CÁLCULO SCTR
# ==========================================
@st.cache_data
def cargar_tarifario_sctr():
    try:
        archivos_posibles = ['Tarifario_sctr_3.xlsx', 'tarifario_sctr.xlsx', 'Tarifario_sctr.xlsx']
        ruta_activa = next((f for f in archivos_posibles if os.path.exists(f)), None)
        if not ruta_activa: return None, None
        
        xls = pd.ExcelFile(ruta_activa)
        if len(xls.sheet_names) >= 2:
            df_tarifas = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
            df_imp = pd.read_excel(xls, sheet_name=xls.sheet_names[1])
        else:
            df_tarifas = pd.read_excel(xls, sheet_name=0)
            df_imp = pd.DataFrame()
            
        df_tarifas.columns = df_tarifas.columns.str.strip()
        df_imp.columns = df_imp.columns.str.strip()
        
        renombres = {
            'Rubro Comercial (Nombre amigable para la web)': 'Rubro', 'Actividades / CIIU Principales que agrupa': 'Actividad',
            'Rímac Salud': 'Tasa_Salud_Rimac', 'Rímac Pensión': 'Tasa_Pension_Rimac',
            'Pacífico Salud': 'Tasa_Salud_Pacifico', 'Pacífico Pensión': 'Tasa_Pension_Pacifico',
            'Mapfre Salud': 'Tasa_Salud_Mapfre', 'Mapfre Pensión': 'Tasa_Pension_Mapfre',
            'Sanitas Salud': 'Tasa_Salud_Sanitas', 'Protecta Pensión': 'Tasa_Pension_Protecta'
        }
        return df_tarifas.rename(columns=renombres), df_imp
    except Exception as e:
        st.error(f"Error leyendo Excel SCTR: {e}")
        return None, None

def cotizar_sctr(df_t, df_i, actividad_elegida, planilla):
    if df_t is None or df_t.empty: return pd.DataFrame()
    fila = df_t[df_t['Actividad'] == actividad_elegida]
    if fila.empty: return pd.DataFrame()
    row = fila.iloc[0]

    def get_val(df_m, concepto, col_name, default=0.0):
        if df_m is None or df_m.empty: return default
        try:
            col_match = next((c for c in df_m.columns if str(col_name) in str(c)), None)
            if not col_match: return default
            val = df_m.loc[df_m['Concepto'].str.contains(concepto, case=False, na=False), col_match].values[0]
            if isinstance(val, str) and '%' in val: return float(val.replace('%', '').replace(',', '.')) / 100.0
            return float(str(val).replace(',', '.'))
        except: return default

    cias = [
        ("Rímac", "Tasa_Salud_Rimac", "Rímac Salud", "Tasa_Pension_Rimac", "Rímac Pens"),
        ("Pacífico", "Tasa_Salud_Pacifico", "Pacífico Salu", "Tasa_Pension_Pacifico", "Pacífico Pens"),
        ("Mapfre", "Tasa_Salud_Mapfre", "Mapfre Salu", "Tasa_Pension_Mapfre", "Mapfre Pensi"),
        ("Sanitas - Protecta", "Tasa_Salud_Sanitas", "Sanitas Salu", "Tasa_Pension_Protecta", "Protecta Per")
    ]
    
    resultados = []
    for nombre, col_ts, col_hoja2_s, col_tp, col_hoja2_p in cias:
        if col_ts in row and pd.notna(row[col_ts]) and col_tp in row and pd.notna(row[col_tp]):
            t_sal, t_pen = float(row[col_ts]), float(row[col_tp])
            factor_s, factor_p = (t_sal / 100.0) if t_sal > 0.1 else t_sal, (t_pen / 100.0) if t_pen > 0.1 else t_pen
            
            min_s, min_p = get_val(df_i, 'Prima Minima Emision', col_hoja2_s, 60.0), get_val(df_i, 'Prima Minima Emision', col_hoja2_p, 60.0)
            imp_s, imp_p = get_val(df_i, 'Impuestos', col_hoja2_s, 0.18), get_val(df_i, 'Impuestos', col_hoja2_p, 0.2154)
            if imp_s > 1: imp_s = imp_s / 100.0
            if imp_p > 1: imp_p = imp_p / 100.0
            
            neto_sal, neto_pen = max(planilla * factor_s, min_s), max(planilla * factor_p, min_p)
            bruto_sal, bruto_pen = round(neto_sal * (1 + imp_s), 2), round(neto_pen * (1 + imp_p), 2)
            
            resultados.append({
                "Aseguradora": nombre, "Tasa_Salud": t_sal, "Tasa_Pension": t_pen,
                "Bruto_Salud": bruto_sal, "Aplico_Min_Salud": (planilla * factor_s) < min_s,
                "Bruto_Pension": bruto_pen, "Aplico_Min_Pension": (planilla * factor_p) < min_p,
                "Total_Bruto": round(bruto_sal + bruto_pen, 2)
            })
            
    df_out = pd.DataFrame(resultados)
    if not df_out.empty: df_out = df_out.sort_values(by="Total_Bruto", ascending=True).reset_index(drop=True)
    return df_out

# ==========================================
# 4. MOTOR DE VENTA CRUZADA (VIDA LEY ADD-ON)
# ==========================================
@st.cache_data
def cargar_tarifario_vidaley_addon():
    try:
        ruta = next((f for f in ['tarifario_vidaley.xlsx', 'Tarifario_vidaley.xlsx'] if os.path.exists(f)), None)
        if not ruta: return None, None
        xls = pd.ExcelFile(ruta)
        df_h1, df_h2 = pd.read_excel(xls, sheet_name=xls.sheet_names[0]), pd.read_excel(xls, sheet_name=xls.sheet_names[1])
        df_h1.columns, df_h2.columns = df_h1.columns.str.strip(), df_h2.columns.str.strip()
        df_h1['ACTIVIDAD_CLEAN'] = df_h1['ACTIVIDAD'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_h1['RANGO_CLEAN'] = df_h1['RANGO DE EDAD'].astype(str).str.strip()
        df_h2['Concepto_Clean'] = df_h2['Concepto'].astype(str).str.strip()
        return df_h1, df_h2
    except Exception: return None, None

def calcular_vidaley_addon(df_vl, df_minimos, actividad_clean, categoria_sel, planilla_total):
    sub_df = df_vl[(df_vl['ACTIVIDAD_CLEAN'] == actividad_clean) & (df_vl['RANGO_CLEAN'] == '18-36')]
    if sub_df.empty: return pd.DataFrame()

    mapa_columnas = {
        "Empleado (Administrativo / Oficina)": {"Mapfre": "MAPFRE: EMPLEADO", "Pacífico": "PACIFICO: EMPLEADO", "Protecta": "PROTECTA: EMPLEADO", "Rímac": "RIMAC PLAN BASE - EMPLEADOS"},
        "Obrero (Operativo Estándar)": {"Mapfre": "MAPFRE: OBRERO", "Pacífico": "PACIFICO: OBRERO", "Protecta": "PROTECTA: OBRERO", "Rímac": "RIMAC PLAN BASE - OBREROS"},
        "Obrero Alto Riesgo (Construcción / Planta / Mina)": {"Mapfre": "MAPFRE: OBRERO ALTO RIESGO", "Pacífico": "PACIFICO: OBRERO ALTO RIESGO", "Protecta": "PROTECTA: OBRERO ALTO RIESGO", "Rímac": "RIMAC PLAN BASE - OBREROS ALTO RIESGO"}
    }
    cols_cia = mapa_columnas.get(categoria_sel, mapa_columnas["Obrero Alto Riesgo (Construcción / Planta / Mina)"])
    resultados = []

    for cia, col_tasa in cols_cia.items():
        if col_tasa not in sub_df.columns or cia not in df_minimos.columns: continue
        try:
            imp_serie = df_minimos[df_minimos['Concepto_Clean'].str.contains('Impuesto', case=False, na=False)][cia]
            imp = float(imp_serie.iloc[0]) if not imp_serie.empty and pd.notna(imp_serie.iloc[0]) else 0.18
            def get_min(palabra):
                s = df_minimos[df_minimos['Concepto_Clean'].str.contains(palabra, case=False, na=False)][cia]
                return float(s.iloc[0]) if not s.empty and pd.notna(s.iloc[0]) else 0.0
            min_m, min_t, min_s, min_a = get_min('Mensual'), get_min('Trimestral'), get_min('Semestral'), get_min('Anual')
        except Exception: continue

        tasa = float(sub_df.iloc[0][col_tasa])
        prima_neta_mensual_calc = planilla_total * tasa
        if prima_neta_mensual_calc <= 0: continue

        neto_m, neto_t, neto_s, neto_a = max(prima_neta_mensual_calc * 1.0, min_m), max(prima_neta_mensual_calc * 3.0, min_t), max(prima_neta_mensual_calc * 6.0, min_s), max(prima_neta_mensual_calc * 12.0, min_a)

        resultados.append({
            "Aseguradora VL": cia, "Bruto_Mensual": round(neto_m * (1.0 + imp), 2),
            "Bruto_Trimestral": round(neto_t * (1.0 + imp), 2), "Bruto_Semestral": round(neto_s * (1.0 + imp), 2),
            "Bruto_Anual": round(neto_a * (1.0 + imp), 2)
        })
    df_out = pd.DataFrame(resultados)
    if not df_out.empty: df_out = df_out.sort_values(by=["Bruto_Anual", "Bruto_Mensual"], ascending=[True, True]).reset_index(drop=True)
    return df_out

# ==========================================
# 5. GENERADOR DE PDF
# ==========================================
def generar_pdf_sctr(datos, df_res, folio, razon_experto=""):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15, leftMargin=15, topMargin=18, bottomMargin=18)
    estilos = getSampleStyleSheet()

    AZUL_CORP, GRIS_TEXTO, GRIS_FONDO, BORDE_SUAVE = colors.HexColor("#03318e"), colors.HexColor("#3c3c3c"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")
    st_tit = ParagraphStyle('T', parent=estilos['Heading1'], fontName='Helvetica-Bold', fontSize=13, textColor=AZUL_CORP, leading=15)
    st_sub = ParagraphStyle('S', parent=estilos['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=AZUL_CORP, spaceAfter=4)
    st_norm = ParagraphStyle('N', parent=estilos['Normal'], fontSize=8, textColor=GRIS_TEXTO, leading=11)
    st_bold = ParagraphStyle('B', parent=st_norm, fontName='Helvetica-Bold', textColor=AZUL_CORP)
    st_th = ParagraphStyle('TH', parent=estilos['Normal'], fontSize=8, fontName='Helvetica-Bold', textColor=colors.white, alignment=1)
    st_td_c = ParagraphStyle('TDC', parent=estilos['Normal'], fontSize=7.5, textColor=GRIS_TEXTO, alignment=1)

    elements = []
    img_yq = ImageRL("logo.png", width=4.0*cm, height=2.0*cm, kind='proportional') if os.path.exists("logo.png") else Paragraph("<b>YQ CORREDORES</b>", st_tit)
    p_header = Paragraph("<b>YQ CORREDORES DE SEGUROS</b><br/>Slip de Cotización SCTR (Salud y Pensión)<br/><font size='7.5' color='#666666'>Seguro Complementario de Trabajo de Riesgo - Ley N° 26790</font>", st_tit)
    p_folio = Paragraph(f"<font color='#3c3c3c'><b>Cotización:</b> {folio}<br/><b>Fecha:</b> {obtener_hora_peru().strftime('%d/%m/%Y %H:%M')}</font>", ParagraphStyle('F', parent=st_norm, alignment=2))
    t_head = Table([[img_yq, p_header, p_folio]], colWidths=[4.0*cm, 9.5*cm, 4.5*cm])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0), (-1,-1), 1, AZUL_CORP), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
    elements.extend([t_head, Spacer(1, 10)])

    elements.append(Paragraph("1. INFORMACIÓN DEL CONTRATANTE", st_sub))
    t_cli = Table([[Paragraph("<b>Contratante:</b>", st_bold), Paragraph(datos['razon_social'], st_norm), Paragraph("<b>RUC / DNI:</b>", st_bold), Paragraph(datos['ruc'], st_norm)],
                   [Paragraph("<b>Rubro:</b>", st_bold), Paragraph(datos['rubro'], st_norm), Paragraph("<b>Trabajadores:</b>", st_bold), Paragraph(f"{datos['trabajadores']} asegurados", st_norm)],
                   [Paragraph("<b>Actividad CIIU:</b>", st_bold), Paragraph(datos['actividad'], st_norm), Paragraph("<b>Planilla Mensual:</b>", st_bold), Paragraph(f"<b>S/ {datos['planilla']:,.2f}</b>", st_norm)]], colWidths=[2.8*cm, 7.7*cm, 3.0*cm, 4.5*cm])
    t_cli.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 5)]))
    elements.extend([t_cli, Spacer(1, 12)])

    elements.append(Paragraph("2. PROPUESTA ECONÓMICA (MONTOS CON IGV)", st_sub))
    data_tabla = [[Paragraph(h, st_th) for h in ['Aseguradora', 'Salud c/IGV', 'Pensión c/IGV', 'Total Mensual', 'Total Anual']]]
    for idx, row in df_res.iterrows():
        cia = str(row['Aseguradora'])
        mapa_logos = {"Rímac": "logo_rimac.png", "Pacífico": "logo_pacifico.png", "Mapfre": "logo_mapfre.png"}
        ruta_logo = mapa_logos.get(cia, "")
        if ruta_logo and os.path.exists(ruta_logo):
            celda_cia = Table([[ImageRL(ruta_logo, width=2.2*cm, height=0.7*cm, kind='proportional')]], colWidths=[3.2*cm])
            celda_cia.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
        else: celda_cia = Paragraph(f"<b>{cia}</b>", st_td_c)

        data_tabla.append([celda_cia, Paragraph(f"S/ {row['Bruto_Salud']:,.2f}", st_td_c), Paragraph(f"S/ {row['Bruto_Pension']:,.2f}", st_td_c), Paragraph(f"<font color='#089685' size='8'><b>S/ {row['Total_Bruto']:,.2f}</b></font>", st_td_c), Paragraph(f"S/ {row['Total_Bruto']*12:,.2f}", st_td_c)])
    
    t_ofertas = Table(data_tabla, colWidths=[3.8*cm, 3.4*cm, 3.4*cm, 3.4*cm, 3.4*cm], repeatRows=1)
    estilos_tabla = [('BACKGROUND', (0,0), (-1,0), AZUL_CORP), ('GRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 6)]
    if len(df_res) > 0: estilos_tabla.extend([('BACKGROUND', (0,1), (-1,1), colors.HexColor("#F0F6FF")), ('BOX', (0,1), (-1,1), 1.5, AZUL_CORP)])
    t_ofertas.setStyle(TableStyle(estilos_tabla))
    elements.extend([t_ofertas, Spacer(1, 10)])

    if razon_experto:
        elements.append(Paragraph("3. ANÁLISIS DEL EXPERTO", st_sub))
        t_box = Table([[Paragraph(f"{razon_experto}", st_norm)]], colWidths=[18.0*cm])
        t_box.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 6)]))
        elements.append(t_box)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 6. INTERFAZ DE USUARIO 
# ==========================================
df_t, df_i = cargar_tarifario_sctr()
origen_url = st.query_params.get("origen", "")

if origen_url != "web":
    if os.path.exists("logo_web.png"):
        c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
        with c_l2: st.image("logo_web.png", use_container_width=True)
    st.markdown("""
        <div style="text-align: center; padding: 10px 0px;">
            <h1 style="color: #03318e; font-size: 2.3rem; font-weight: 800; margin-bottom: 5px;">Cotiza tu SCTR al instante.</h1>
            <p style="color: #3c3c3c; font-size: 1.1rem;">Compara Pacífico, Rímac, Mapfre y Protecta.</p>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

codigo_actual = st.session_state.get('codigo_secreto', '')
es_admin = (codigo_actual == CODIGO_ADMIN or codigo_actual in CODIGOS_ASESORES)

if df_t is None:
    st.error("❌ No se encontró el archivo 'Tarifario_sctr.xlsx'.")
else:
    # Precarga cruzada de Vida Ley
    razon_url, ruc_url, cel_url, mail_url = st.query_params.get("razon", ""), st.query_params.get("ruc", ""), st.query_params.get("cel", ""), st.query_params.get("mail", "")
    try: trab_default = max(1, int(st.query_params.get("trab", "1")))
    except: trab_default = 1
    if razon_url: st.info(f"⚡ **¡Hola {razon_url}!** Datos precargados correctamente.")

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-bottom: 10px;'>1. Datos de la Empresa</h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: razon_social, ruc = st.text_input("Razón Social", value=razon_url, placeholder="Constructora SAC"), st.text_input("RUC o DNI", value=ruc_url, max_chars=11)
    with col2: celular, email = st.text_input("Celular", value=cel_url, max_chars=9), st.text_input("Correo", value=mail_url)

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-top: 15px;'>2. Actividad de Riesgo y Planilla</h3>", unsafe_allow_html=True)
    rubro_sel = st.selectbox("Rubro", df_t['Rubro'].dropna().unique().tolist())
    actividad_sel = st.selectbox("Actividad (CIIU)", df_t[df_t['Rubro'] == rubro_sel]['Actividad'].dropna().unique().tolist())
    
    col_t1, col_t2 = st.columns(2)
    with col_t1: n_trabajadores = st.number_input("N° Trabajadores", min_value=1, max_value=5000, value=trab_default)
    with col_t2: planilla_mensual = st.number_input("Planilla Total (S/)", min_value=float(n_trabajadores*RMV_VIGENTE), value=float(n_trabajadores*RMV_VIGENTE), step=100.0)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Generar Propuesta SCTR", type="primary", use_container_width=True):
        cel_cl, ruc_cl, mail_cl = "".join(filter(str.isdigit, str(celular))), ruc.strip(), email.strip()
        
        es_ruc = bool(re.match(r'^(10|20)\d{9}$', ruc_cl))
        es_dni = bool(re.match(r'^\d{8}$', ruc_cl))

        if not razon_social.strip():
            st.error("⚠️ Por favor ingresa la Razón Social o Nombre del Contratante.")
        elif not es_admin and (not ruc_cl or not (es_ruc or es_dni)):
            st.error("🏢 Ingresa un RUC válido (11 dígitos, empezando en 10 o 20) o un DNI válido (8 dígitos).")
        elif not es_admin and (len(cel_cl) != 9 or not cel_cl.startswith('9')):
            st.error("📱 Por favor ingresa un número de celular válido de 9 dígitos (Ej: 999123456).")
        elif not es_admin and (not mail_cl or not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", mail_cl)):
            st.error("📧 Por favor ingresa un correo electrónico real para enviarte tu propuesta.")
        else:
            df_res = cotizar_sctr(df_t, df_i, actividad_sel, planilla_mensual)
            st.session_state['res_sctr'], st.session_state['id_sctr'] = df_res, f"SCTR-{obtener_hora_peru().strftime('%m%d-%H%M')}"
            st.session_state['datos_sctr'] = {"razon_social": razon_social.strip(), "ruc": ruc_cl, "celular": cel_cl, "email": mail_cl, "rubro": rubro_sel, "actividad": actividad_sel, "trabajadores": n_trabajadores, "planilla": planilla_mensual}
            st.session_state['ver_vl'] = False # Reset panel cruzado
            if not df_res.empty and not es_admin:
                enviar_notificacion_sctr(razon_social.strip(), ruc_cl, cel_cl, mail_cl, rubro_sel, actividad_sel, n_trabajadores, planilla_mensual, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Total_Bruto'])
                guardar_en_sheets([obtener_hora_peru().strftime('%Y-%m-%d'), obtener_hora_peru().strftime('%H:%M:%S'), st.session_state['id_sctr'], razon_social.strip(), ruc_cl, cel_cl, mail_cl, rubro_sel, actividad_sel, n_trabajadores, planilla_mensual, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Total_Bruto']])

    if st.session_state.get('res_sctr') is not None:
        df_res, datos, folio = st.session_state['res_sctr'], st.session_state['datos_sctr'], st.session_state['id_sctr']
        if df_res.empty: st.warning("No hay tasas para esta actividad.")
        else:
            opciones_cias = df_res['Aseguradora'].tolist()
            
            if es_admin:
                st.success("🔓 MODO CORREDOR ACTIVADO - Configura la recomendación para el cliente:")
                mejor_opcion_elegida = st.selectbox("Selecciona la Aseguradora a destacar como 'Mejor Opción':", opciones_cias, index=0)
                df_res['es_top'] = df_res['Aseguradora'] == mejor_opcion_elegida
                df_res = df_res.sort_values(by=['es_top', 'Total_Bruto'], ascending=[False, True]).drop(columns=['es_top']).reset_index(drop=True)
                
                razon_default = f"Este plan de {mejor_opcion_elegida} ofrece la tarifa más competitiva del mercado para la actividad de {datos['rubro']}, optimizando los costos de planilla y cumpliendo con las exigencias de ley."
                razon_admin = st.text_area("Motivo (Análisis del Experto para el PDF):", value=razon_default)
            else:
                mejor_opcion_elegida = opciones_cias[0]
                razon_admin = (
                    "Nuestro algoritmo ha analizado el mercado SCTR filtrando las tasas de todas las aseguradoras disponibles. "
                    f"Destacamos a {mejor_opcion_elegida} como la opción de mejor relación costo-beneficio para proteger a tus trabajadores, garantizando "
                    "cobertura inmediata ante la autoridad laboral y optimizando el presupuesto de tu empresa."
                )

            st.markdown("---")
            st.success(f"¡Comparativa lista! Opción recomendada destacada: **{mejor_opcion_elegida}**")
            
            df_res['Total_Anual'] = df_res['Total_Bruto'] * 12
            df_vista = df_res[['Aseguradora', 'Bruto_Salud', 'Bruto_Pension', 'Total_Bruto', 'Total_Anual']].copy()
            df_vista.columns = ['Aseguradora', 'Salud c/IGV', 'Pensión c/IGV', 'Total Mensual', 'Total Anual']
            for c in df_vista.columns[1:]: df_vista[c] = df_vista[c].apply(lambda x: f"S/ {x:,.2f}")
            st.dataframe(df_vista, hide_index=True, use_container_width=True)

            pdf_buf = generar_pdf_sctr(datos, df_res, folio, razon_admin)
            col_b1, col_b2 = st.columns(2)
            with col_b1: st.download_button("Descargar Slip SCTR (PDF)", data=pdf_buf, file_name=f"SCTR_{folio}.pdf", mime="application/pdf")
            with col_b2:
                link_wa = f"https://api.whatsapp.com/send?phone=51906462225&text=Hola YQ, deseo emitir SCTR ({folio}) con {mejor_opcion_elegida}."
                st.markdown(f"<a href='{link_wa}' target='_blank' style='display:flex;align-items:center;justify-content:center;width:100%;height:48px;background-color:#25D366;color:white;border-radius:8px;text-decoration:none;font-weight:700;'>Emitir Póliza por WhatsApp</a>", unsafe_allow_html=True)

            # ---------------------------------------------------------
            # BLOQUE INFERIOR: ADD-ON VIDA LEY NATIVO (SIN SALIR DE LA PÁGINA)
            # ---------------------------------------------------------
            df_vl_base, df_vl_min = cargar_tarifario_vidaley_addon()
            
            if df_vl_base is not None:
                st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #F0F6FF 0%, #E6EEFF 100%); border: 2px solid #03318e; border-radius: 12px; padding: 20px; margin-top: 25px;">
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <span style="font-size: 28px;">🛡️</span>
                            <div>
                                <h3 style="color: #03318e; margin: 0; font-weight: 800; font-size: 1.15rem;">¿También requieres el Seguro Vida Ley obligatorio?</h3>
                                <p style="color: #3c3c3c; font-size: 0.92rem; margin: 3px 0 0 0;">Cotízalo <b>aquí mismo en 1 clic</b> usando tu misma planilla de <b>S/ {datos['planilla']:,.2f}</b> ({datos['trabajadores']} trabajadores).</p>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                if 'ver_vl' not in st.session_state: st.session_state['ver_vl'] = False
                ver_vl = st.checkbox("➕ Sí, deseo ver y agregar la cotización de Vida Ley para esta misma planilla", value=st.session_state['ver_vl'])
                st.session_state['ver_vl'] = ver_vl
                
                if ver_vl:
                    st.markdown("<div style='padding: 15px; border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 10px;'>", unsafe_allow_html=True)
                    c_v1, c_v2 = st.columns(2)
                    with c_v1:
                        rubros_vl = df_vl_base['ACTIVIDAD_CLEAN'].dropna().unique().tolist()
                        rubro_vl_sel = st.selectbox("Confirma tu Rubro Vida Ley:", rubros_vl, key="addon_rubro_vl")
                    with c_v2:
                        cat_vl_sel = st.selectbox("Categoría de Trabajadores:", ["Empleado (Administrativo / Oficina)", "Obrero (Operativo Estándar)", "Obrero Alto Riesgo (Construcción / Planta / Mina)"], index=2, key="addon_cat_vl")
                    
                    if st.button("Calcular Vida Ley", use_container_width=True, key="btn_calc_vl"):
                        df_res_vl = calcular_vidaley_addon(df_vl_base, df_vl_min, rubro_vl_sel, cat_vl_sel, datos['planilla'])
                        if not df_res_vl.empty:
                            st.session_state['res_vl_addon'] = df_res_vl
                            st.session_state['cia_vl_addon'] = df_res_vl.iloc[0]['Aseguradora VL']
                            st.session_state['anual_vl_addon'] = df_res_vl.iloc[0]['Bruto_Anual']
                            st.session_state['mes_vl_addon'] = df_res_vl.iloc[0]['Bruto_Mensual']
                        else: st.warning("No se encontraron tasas para esa actividad.")
                    
                    if st.session_state.get('res_vl_addon') is not None:
                        df_res_vl = st.session_state['res_vl_addon']
                        mejor_vl_cia, mejor_vl_anual, mejor_vl_mes = st.session_state['cia_vl_addon'], st.session_state['anual_vl_addon'], st.session_state['mes_vl_addon']
                        
                        df_vl_vista = df_res_vl.copy()
                        for c_m in ['Bruto_Mensual', 'Bruto_Trimestral', 'Bruto_Semestral', 'Bruto_Anual']: df_vl_vista[c_m] = df_vl_vista[c_m].apply(lambda x: f"S/ {x:,.2f}")
                        st.dataframe(df_vl_vista, hide_index=True, use_container_width=True)
                        
                        mejor_sctr_total = df_res.iloc[0]['Total_Bruto']
                        pack_mensual = mejor_vl_mes + mejor_sctr_total
                        st.info(f"💼 **PACK CORPORATIVO COMPLETO:** SCTR ({mejor_opcion_elegida}) + Vida Ley ({mejor_vl_cia}) = **S/ {pack_mensual:,.2f} / mes** (o SCTR Mensual S/ {mejor_sctr_total:,.2f} + Vida Ley Anual S/ {mejor_vl_anual:,.2f}).")
                        
                        msg_pack = f"Hola YQ, acabo de cotizar el PACK CORPORATIVO ({folio}) para {datos['razon_social']} (RUC: {datos['ruc']}, {datos['trabajadores']} trab.): 1) SCTR con {mejor_opcion_elegida} (S/ {mejor_sctr_total:,.2f}/mes) + 2) Vida Ley con {mejor_vl_cia} (Anual: S/ {mejor_vl_anual:,.2f} / Mensual: S/ {mejor_vl_mes:,.2f}). Deseo emitir ambas pólizas."
                        link_wa_pack = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_pack)}"
                        st.markdown(f"<a href='{link_wa_pack}' target='_blank' style='display:flex;align-items:center;justify-content:center;width:100%;height:48px;background-color:#089685;color:white;border-radius:8px;text-decoration:none;font-weight:700;'>📦 Solicitar Pack (SCTR + Vida Ley) por WhatsApp</a>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

with st.expander("🛡️ Acceso Interno YQ (Solo Empleados)"):
    st.text_input("Código de Autorización", type="password", key="codigo_secreto")
    if es_admin and st.button("📥 Descargar Historial CSV"):
        df_hist = descargar_historial_sheets()
        if df_hist is not None and not df_hist.empty: st.download_button("💾 Guardar CSV", df_hist.to_csv(index=False).encode('utf-8-sig'), "Historial_SCTR.csv", "text/csv")
