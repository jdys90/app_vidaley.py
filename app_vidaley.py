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
st.set_page_config(page_title="Cotizador Vida Ley Digital | YQ Corredores", layout="centered", page_icon="🛡️")

st.markdown("""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"] { display: none !important; visibility: hidden !important; }
        footer {visibility: hidden !important; display: none !important;}
        .stApp > footer {display: none !important;}
        [data-testid="stBottom"], [data-testid="stBottomBar"] {display: none !important;}
        [data-testid="stViewerBadge"] {display: none !important;}
        .block-container { padding-bottom: 0rem !important; }

        div.stButton > button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; font-size: 16px !important; font-weight: 700 !important; width: 100% !important; }
        div.stButton > button:hover { background-color: #022266 !important; border-color: #022266 !important; }
        
        div[data-testid="stDownloadButton"] button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; width: 100% !important; font-size: 16px !important; font-weight: 700 !important; }
        div[data-testid="stDownloadButton"] button:hover { background-color: #022266 !important; border-color: #022266 !important; }
    </style>
""", unsafe_allow_html=True)

CODIGO_ADMIN = "ADMIN2026"
CODIGOS_ASESORES = ["ASE01", "ASE02", "ASE03", "VENTAS2026"]
RMV_VIGENTE = 1230.0

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ImageRL
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
except ImportError:
    st.error("❌ Falta la librería 'reportlab'. Asegúrate de incluirla en tu requirements.txt")
    st.stop()

def obtener_hora_peru():
    return datetime.datetime.utcnow() - timedelta(hours=5)

# ==========================================
# 2. FUNCIONES BASE DE DATOS Y CRM
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
        if client: client.open("historial_cotizador_vidaley").sheet1.append_row(datos_fila)
    except Exception: pass

def descargar_historial_sheets():
    try:
        client = get_gspread_client()
        if client: return pd.DataFrame(client.open("historial_cotizador_vidaley").sheet1.get_all_records())
    except Exception: return None
    return None

def enviar_notificacion_vidaley(cliente, ruc, celular, email, actividad, categoria, trabajadores, planilla, mejor_cia, mejor_anual, mejor_mensual):
    try:
        if "ZOHO_REFRESH_TOKEN" not in st.secrets: return False
        url_auth = "https://accounts.zoho.com/oauth/v2/token"
        datos_auth = {"refresh_token": st.secrets["ZOHO_REFRESH_TOKEN"], "client_id": st.secrets["ZOHO_CLIENT_ID"], "client_secret": st.secrets["ZOHO_CLIENT_SECRET"], "grant_type": "refresh_token"}
        access_token = requests.post(url_auth, data=datos_auth).json().get("access_token")
        if access_token:
            desc = f"Producto: Vida Ley | Actividad: {actividad} | Categoría: {categoria} | Trabajadores: {trabajadores} | Planilla Mensual: S/ {planilla:,.2f} | Mejor Opción: {mejor_cia} (Anual: S/ {mejor_anual:,.2f} / Mensual: S/ {mejor_mensual:,.2f})"
            payload = {"data": [{"Company": cliente, "Last_Name": "Contacto Comercial", "RUC": str(ruc), "Email": email, "Mobile": str(celular), "Lead_Source": "Cotizador Vida Ley Web", "Seguro_de_Inter_s": "Vida Ley", "Description": desc}]}
            requests.post("https://www.zohoapis.com/crm/v2/Leads", headers={"Authorization": f"Zoho-oauthtoken {access_token}"}, json=payload)
    except Exception: pass

# ==========================================
# 3. MOTORES DE CÁLCULO
# ==========================================
@st.cache_data
def cargar_tarifario_vidaley():
    try:
        ruta = next((f for f in ['tarifario_vidaley.xlsx', 'Tarifario_vidaley.xlsx'] if os.path.exists(f)), None)
        if not ruta: return None, None
        xls = pd.ExcelFile(ruta)
        df_h1 = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        df_h2 = pd.read_excel(xls, sheet_name=xls.sheet_names[1])
        df_h1.columns, df_h2.columns = df_h1.columns.str.strip(), df_h2.columns.str.strip()
        df_h1['ACTIVIDAD_CLEAN'] = df_h1['ACTIVIDAD'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_h1['RANGO_CLEAN'] = df_h1['RANGO DE EDAD'].astype(str).str.strip()
        df_h2['Concepto_Clean'] = df_h2['Concepto'].astype(str).str.strip()
        return df_h1, df_h2
    except Exception: return None, None

def cotizar_vidaley(actividad_clean, categoria_sel, planilla_por_rango):
    df_h1, df_h2 = cargar_tarifario_vidaley()
    if df_h1 is None or df_h2 is None: return pd.DataFrame()
    sub_df = df_h1[df_h1['ACTIVIDAD_CLEAN'] == actividad_clean]
    if sub_df.empty: return pd.DataFrame()

    mapa_columnas = {
        "Empleado (Administrativo / Oficina)": {"Mapfre": "MAPFRE: EMPLEADO", "Pacífico": "PACIFICO: EMPLEADO", "Protecta": "PROTECTA: EMPLEADO", "Rímac": "RIMAC PLAN BASE - EMPLEADOS"},
        "Obrero (Operativo Estándar)": {"Mapfre": "MAPFRE: OBRERO", "Pacífico": "PACIFICO: OBRERO", "Protecta": "PROTECTA: OBRERO", "Rímac": "RIMAC PLAN BASE - OBREROS"},
        "Obrero Alto Riesgo (Construcción / Planta / Mina)": {"Mapfre": "MAPFRE: OBRERO ALTO RIESGO", "Pacífico": "PACIFICO: OBRERO ALTO RIESGO", "Protecta": "PROTECTA: OBRERO ALTO RIESGO", "Rímac": "RIMAC PLAN BASE - OBREROS ALTO RIESGO"}
    }
    cols_cia = mapa_columnas.get(categoria_sel, mapa_columnas["Obrero (Operativo Estándar)"])
    resultados = []

    for cia, col_tasa in cols_cia.items():
        if col_tasa not in sub_df.columns or cia not in df_h2.columns: continue
        try:
            imp = float(df_h2.loc[df_h2['Concepto_Clean'].str.contains('Impuesto', case=False, na=False), cia].values[0]) if not df_h2.loc[df_h2['Concepto_Clean'].str.contains('Impuesto', case=False, na=False), cia].empty else 0.18
            def get_min(palabra):
                s = df_h2.loc[df_h2['Concepto_Clean'].str.contains(palabra, case=False, na=False), cia]
                return float(s.values[0]) if not s.empty and pd.notna(s.values[0]) else 0.0
            min_m, min_t, min_s, min_a = get_min('Mensual'), get_min('Trimestral'), get_min('Semestral'), get_min('Anual')
        except Exception: continue

        prima_neta_mensual_calc = 0.0
        for rango, monto_planilla in planilla_por_rango.items():
            if monto_planilla > 0:
                fila_r = sub_df[sub_df['RANGO_CLEAN'] == rango]
                if not fila_r.empty and pd.notna(fila_r.iloc[0][col_tasa]):
                    prima_neta_mensual_calc += monto_planilla * float(fila_r.iloc[0][col_tasa])

        if prima_neta_mensual_calc <= 0: continue

        neto_m, neto_t, neto_s, neto_a = max(prima_neta_mensual_calc * 1.0, min_m), max(prima_neta_mensual_calc * 3.0, min_t), max(prima_neta_mensual_calc * 6.0, min_s), max(prima_neta_mensual_calc * 12.0, min_a)

        resultados.append({
            "Aseguradora": cia,
            "Bruto_Mensual": round(neto_m * (1.0 + imp), 2), "Min_M": (prima_neta_mensual_calc * 1.0) < min_m,
            "Bruto_Trimestral": round(neto_t * (1.0 + imp), 2), "Min_T": (prima_neta_mensual_calc * 3.0) < min_t,
            "Bruto_Semestral": round(neto_s * (1.0 + imp), 2), "Min_S": (prima_neta_mensual_calc * 6.0) < min_s,
            "Bruto_Anual": round(neto_a * (1.0 + imp), 2), "Min_A": (prima_neta_mensual_calc * 12.0) < min_a
        })

    df_out = pd.DataFrame(resultados)
    if not df_out.empty: df_out = df_out.sort_values(by=["Bruto_Anual", "Bruto_Mensual"], ascending=[True, True]).reset_index(drop=True)
    return df_out

@st.cache_data
def cargar_tarifario_sctr_addon():
    try:
        ruta = next((f for f in ['Tarifario_sctr_3.xlsx', 'tarifario_sctr.xlsx', 'Tarifario_sctr.xlsx'] if os.path.exists(f)), None)
        if not ruta: return None, None
        xls = pd.ExcelFile(ruta)
        df_tasas = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        df_minimos = pd.read_excel(xls, sheet_name=xls.sheet_names[1])
        df_tasas.columns, df_minimos.columns = df_tasas.columns.str.strip(), df_minimos.columns.str.strip()
        renombres = {
            'Rubro Comercial (Nombre amigable para la web)': 'Rubro',
            'Actividades / CIIU Principales que agrupa': 'Actividad',
            'Rímac Salud': 'Tasa_Salud_Rimac', 'Rímac Pensión': 'Tasa_Pension_Rimac',
            'Pacífico Salud': 'Tasa_Salud_Pacifico', 'Pacífico Pensión': 'Tasa_Pension_Pacifico',
            'Mapfre Salud': 'Tasa_Salud_Mapfre', 'Mapfre Pensión': 'Tasa_Pension_Mapfre',
            'Sanitas Salud': 'Tasa_Salud_Sanitas', 'Protecta Pensión': 'Tasa_Pension_Protecta'
        }
        df_tasas = df_tasas.rename(columns=renombres)
        # Limpiar espacios extras en Actividad para hacer el match exacto
        df_tasas['Actividad'] = df_tasas['Actividad'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        return df_tasas, df_minimos
    except Exception: return None, None

def calcular_sctr_addon(df_sctr, df_minimos, act_sel_limpia, planilla_total):
    fila = df_sctr[df_sctr['Actividad'] == act_sel_limpia]
    if fila.empty: return pd.DataFrame()
    row = fila.iloc[0]
    
    def get_val(df_m, concepto, col_name, default=0.0):
        try:
            col_match = next((c for c in df_m.columns if str(col_name) in str(c)), None)
            if not col_match: return default
            val = df_m.loc[df_m['Concepto'].str.contains(concepto, case=False, na=False), col_match].values[0]
            if isinstance(val, str):
                if '%' in val: return float(val.replace('%', '').replace(',', '.')) / 100.0
                return float(val.replace(',', '.'))
            return float(val)
        except: return default

    cias = [
        ("Rímac", "Tasa_Salud_Rimac", "Rímac Salud", "Tasa_Pension_Rimac", "Rímac Pens"),
        ("Pacífico", "Tasa_Salud_Pacifico", "Pacífico Salu", "Tasa_Pension_Pacifico", "Pacífico Pens"),
        ("Mapfre", "Tasa_Salud_Mapfre", "Mapfre Salu", "Tasa_Pension_Mapfre", "Mapfre Pensi"),
        ("Sanitas + Protecta", "Tasa_Salud_Sanitas", "Sanitas Salu", "Tasa_Pension_Protecta", "Protecta Pens")
    ]
    
    res = []
    for nombre, col_ts, col_hoja2_s, col_tp, col_hoja2_p in cias:
        if col_ts in row and pd.notna(row[col_ts]) and col_tp in row and pd.notna(row[col_tp]):
            t_sal, t_pen = float(row[col_ts]), float(row[col_tp])
            factor_s, factor_p = (t_sal / 100.0) if t_sal > 0.1 else t_sal, (t_pen / 100.0) if t_pen > 0.1 else t_pen
            
            min_s = get_val(df_minimos, 'Prima Minima Emision', col_hoja2_s, 60.0)
            min_p = get_val(df_minimos, 'Prima Minima Emision', col_hoja2_p, 60.0)
            imp_s = get_val(df_minimos, 'Impuestos', col_hoja2_s, 0.18)
            imp_p = get_val(df_minimos, 'Impuestos', col_hoja2_p, 0.2154)
            if imp_s > 1: imp_s = imp_s / 100.0
            if imp_p > 1: imp_p = imp_p / 100.0
            
            neto_sal = max(planilla_total * factor_s, min_s)
            neto_pen = max(planilla_total * factor_p, min_p)
            bruto_sal = round(neto_sal * (1 + imp_s), 2)
            bruto_pen = round(neto_pen * (1 + imp_p), 2)
            
            res.append({"Aseguradora SCTR": nombre, "Salud c/IGV": bruto_sal, "Pensión c/IGV": bruto_pen, "Total Mensual SCTR": round(bruto_sal + bruto_pen, 2), "Total Anual SCTR": round((bruto_sal + bruto_pen) * 12, 2)})
            
    df_out = pd.DataFrame(res)
    if not df_out.empty: df_out = df_out.sort_values(by="Total Mensual SCTR", ascending=True).reset_index(drop=True)
    return df_out

# ==========================================
# 4. MOTOR GENERADOR DE PDF (CON DISEÑO PREMIUM)
# ==========================================
def generar_pdf_vidaley(datos, df_res, folio, razon_experto="", df_sctr_opcional=None):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15, leftMargin=15, topMargin=16, bottomMargin=16)
    estilos = getSampleStyleSheet()

    AZUL_CORP, GRIS_TEXTO, GRIS_FONDO, BORDE_SUAVE = colors.HexColor("#03318e"), colors.HexColor("#3c3c3c"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")
    st_tit = ParagraphStyle('T', parent=estilos['Heading1'], fontName='Helvetica-Bold', fontSize=13, textColor=AZUL_CORP, leading=15)
    st_sub = ParagraphStyle('S', parent=estilos['Normal'], fontName='Helvetica-Bold', fontSize=9.5, textColor=AZUL_CORP, spaceAfter=4)
    st_norm = ParagraphStyle('N', parent=estilos['Normal'], fontSize=8, textColor=GRIS_TEXTO, leading=10.5)
    st_bold = ParagraphStyle('B', parent=st_norm, fontName='Helvetica-Bold', textColor=AZUL_CORP)
    st_th = ParagraphStyle('TH', parent=estilos['Normal'], fontSize=7.5, fontName='Helvetica-Bold', textColor=colors.white, alignment=1, leading=9)
    st_td = ParagraphStyle('TD', parent=estilos['Normal'], fontSize=7.5, textColor=GRIS_TEXTO, leading=9.5)
    st_td_c = ParagraphStyle('TDC', parent=st_td, alignment=1)
    st_rec = ParagraphStyle('REC', parent=st_td_c, fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor("#bf8d1b"))

    elements = []

    # --- 1. CABECERA ---
    img_yq = ImageRL("logo.png", width=3.8*cm, height=1.9*cm, kind='proportional') if os.path.exists("logo.png") else Paragraph("<b>YQ CORREDORES</b>", st_tit)
    txt_header = "<b>YQ CORREDORES DE SEGUROS</b><br/>Slip de Cotización Seguro Vida Ley"
    if df_sctr_opcional is not None:
        txt_header += " + SCTR Salud y Pensión"
    txt_header += "<br/><font size='7.5' color='#666666'>Decreto Legislativo N° 688 / Ley N° 29549 / D.U. N° 044-2019 (Desde el 1er día)</font>"
    
    p_header = Paragraph(txt_header, st_tit)
    fecha_str = obtener_hora_peru().strftime('%d/%m/%Y %H:%M')
    p_folio = Paragraph(f"<font color='#3c3c3c'><b>Cotización:</b> {folio}<br/><b>Fecha:</b> {fecha_str}<br/><b>Validez:</b> 15 días</font>", ParagraphStyle('F', parent=st_norm, alignment=2))

    t_head = Table([[img_yq, p_header, p_folio]], colWidths=[3.8*cm, 9.7*cm, 4.5*cm])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0), (-1,-1), 1, AZUL_CORP), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
    elements.extend([t_head, Spacer(1, 8)])

    # --- 2. DATOS DEL CONTRATANTE ---
    elements.append(Paragraph("1. INFORMACIÓN DEL CONTRATANTE Y PLANILLA ASEGURABLE", st_sub))
    
    edades_br = datos['resumen_edades'].replace(", ", "<br/>")
    
    data_cliente = [
        [Paragraph("<b>Contratante:</b>", st_bold), Paragraph(datos['razon_social'], st_norm), Paragraph("<b>RUC / DNI:</b>", st_bold), Paragraph(datos['ruc'], st_norm)],
        [Paragraph("<b>Actividad:</b>", st_bold), Paragraph(datos['actividad'], st_norm), Paragraph("<b>N° Trabajadores:</b>", st_bold), Paragraph(f"{datos['trabajadores']} asegurados ({datos['categoria']})", st_norm)],
        [Paragraph("<b>Distribución<br/>Etaria:</b>", st_bold), Paragraph(edades_br, st_norm), Paragraph("<b>Planilla Mensual:</b>", st_bold), Paragraph(f"<b>S/ {datos['planilla']:,.2f}</b>", st_norm)]
    ]
    t_cli = Table(data_cliente, colWidths=[3.0*cm, 7.5*cm, 3.0*cm, 4.5*cm])
    t_cli.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 4)]))
    elements.extend([t_cli, Spacer(1, 10)])

    # --- 3. PROPUESTA ECONÓMICA ---
    if df_sctr_opcional is not None:
        elements.append(Paragraph("2. PROPUESTA ECONÓMICA PACK CORPORATIVO (VIDA LEY + SCTR C/IGV)", st_sub))
        headers = ['Aseguradora', 'Vida Ley Mensual\n+ SCTR Mensual', 'Vida Ley Trimestral\n+ SCTR Mensual', 'Vida Ley Semestral\n+ SCTR Mensual', 'Vida Ley Anual Único\n+ SCTR Mensual', 'Acción / Emisión']
    else:
        elements.append(Paragraph("2. PROPUESTA ECONÓMICA POR FRECUENCIA DE PAGO (MONTOS FINALES CON IGV 18%)", st_sub))
        headers = ['Aseguradora', 'Pago Mensual\n(c/IGV)', 'Pago Trimestral\n(c/IGV)', 'Pago Semestral\n(c/IGV)', 'Pago Anual Único\n(c/IGV)', 'Acción / Emisión']
    
    anchos = [3.4*cm, 3.0*cm, 3.0*cm, 3.0*cm, 3.0*cm, 2.6*cm]
    data_tabla = [[Paragraph(h, st_th) for h in headers]]

    for idx, row in df_res.iterrows():
        es_mejor = (idx == 0)
        cia = str(row['Aseguradora'])
        celda_cia = []
        if es_mejor: celda_cia.append(Paragraph("<b>► MEJOR OPCIÓN</b>", st_rec))

        mapa_logos = {"Rímac": "logo_rimac.png", "Pacífico": "logo_pacifico.png", "Mapfre": "logo_mapfre.png", "Protecta": "logo_protecta.png", "La Positiva": "logo_positiva.png"}
        ruta_logo = mapa_logos.get(cia, "")
        if ruta_logo and os.path.exists(ruta_logo):
            t_lg = Table([[ImageRL(ruta_logo, width=2.3*cm, height=0.75*cm, kind='proportional')]], colWidths=[3.0*cm])
            t_lg.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            celda_cia.append(t_lg)
        else:
            celda_cia.append(Paragraph(f"<font color='#03318e'><b>{cia}</b></font>", st_td_c))

        if df_sctr_opcional is not None:
            cia_sctr_search = "Sanitas + Protecta" if cia == "Protecta" else cia
            match_sctr = df_sctr_opcional[df_sctr_opcional['Aseguradora SCTR'] == cia_sctr_search]
            sctr_mensual = match_sctr.iloc[0]['Total Mensual SCTR'] if not match_sctr.empty else 0.0
            
            txt_m = f"<b>S/ {row['Bruto_Mensual'] + sctr_mensual:,.2f}</b><br/><font size='6' color='#666666'>(VL S/{row['Bruto_Mensual']:,.0f} + SCTR S/{sctr_mensual:,.0f})</font>"
            txt_t = f"<b>S/ {row['Bruto_Trimestral'] + (sctr_mensual*3):,.2f}</b><br/><font size='6' color='#666666'>(VL S/{row['Bruto_Trimestral']:,.0f} + SCTR S/{sctr_mensual*3:,.0f})</font>"
            txt_s = f"<b>S/ {row['Bruto_Semestral'] + (sctr_mensual*6):,.2f}</b><br/><font size='6' color='#666666'>(VL S/{row['Bruto_Semestral']:,.0f} + SCTR S/{sctr_mensual*6:,.0f})</font>"
            txt_a = f"<font color='#089685' size='9'><b>S/ {row['Bruto_Anual'] + (sctr_mensual*12):,.2f}</b></font><br/><font size='6' color='#666666'>(VL S/{row['Bruto_Anual']:,.0f} + SCTR S/{sctr_mensual*12:,.0f})</font>"
        else:
            tag_m = "<br/><font size='6' color='#666666'>(Prima Mínima)</font>" if row['Min_M'] else ""
            txt_m = f"<b>S/ {row['Bruto_Mensual']:,.2f}</b>{tag_m}"
            tag_t = "<br/><font size='6' color='#666666'>(Prima Mínima)</font>" if row['Min_T'] else ""
            txt_t = f"<b>S/ {row['Bruto_Trimestral']:,.2f}</b>{tag_t}"
            tag_s = "<br/><font size='6' color='#666666'>(Prima Mínima)</font>" if row['Min_S'] else ""
            txt_s = f"<b>S/ {row['Bruto_Semestral']:,.2f}</b>{tag_s}"
            tag_a = "<br/><font size='6' color='#666666'>(Prima Mínima)</font>" if row['Min_A'] else ""
            txt_a = f"<font color='#089685' size='9'><b>S/ {row['Bruto_Anual']:,.2f}</b></font>{tag_a}"

        link_wa = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote('Hola YQ, deseo emitir póliza')}"
        txt_accion = f"<a href='{link_wa}' color='#089685'><font size='7.5'><b>► EMITIR<br/>PÓLIZA</b></font></a>"
        data_tabla.append([celda_cia, Paragraph(txt_m, st_td_c), Paragraph(txt_t, st_td_c), Paragraph(txt_s, st_td_c), Paragraph(txt_a, st_td_c), Paragraph(txt_accion, st_td_c)])
    
    t_ofertas = Table(data_tabla, colWidths=anchos, repeatRows=1)
    estilos_tabla = [('BACKGROUND', (0,0), (-1,0), AZUL_CORP), ('GRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 5)]
    if len(df_res) > 0: estilos_tabla.extend([('BACKGROUND', (0,1), (-1,1), AZUL_CLARO), ('BOX', (0,1), (-1,1), 1.5, AZUL_CORP)])
    t_ofertas.setStyle(TableStyle(estilos_tabla))
    elements.extend([t_ofertas, Spacer(1, 5)])

    # --- 4. RECOMENDACIÓN DEL EXPERTO ---
    if razon_experto:
        elements.append(Paragraph("3. ANÁLISIS Y RECOMENDACIÓN DEL EXPERTO", st_sub))
        elements.append(Spacer(1, 3))
        t_box = Table([[Paragraph(f"<b>ESTRATEGIA RECOMENDADA:</b><br/>{razon_experto}", st_norm)]], colWidths=[18.0*cm])
        t_box.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 6)]))
        elements.extend([t_box, Spacer(1, 8)])

    if df_sctr_opcional is None:
        algun_minimo = any(row['Min_M'] or row['Min_T'] or row['Min_S'] or row['Min_A'] for _, row in df_res.iterrows())
        if algun_minimo:
            nota_min = "<b>* Nota sobre Prima Mínima:</b> Los importes mostrados incluyen el 18% de IGV. Cuando un monto indica <i>(Prima Mínima)</i>, significa que la planilla declarada activa la tarifa mínima de emisión de la aseguradora para esa frecuencia. En planillas pequeñas, elegir el <b>Pago Anual Único</b> genera un ahorro de hasta el 75% frente a realizar 12 pagos mensuales."
            elements.extend([Paragraph(nota_min, ParagraphStyle('NM', parent=st_norm, fontSize=7, textColor=colors.HexColor("#555555"))), Spacer(1, 8)])

    # --- 5. COBERTURAS Y REGLAS ---
    elements.append(Paragraph("4. COBERTURAS OBLIGATORIAS DE LEY (D.L. N° 688 Y LEY N° 29549)", st_sub))
    cob_izq = "<b>INDEMNIZACIONES PRINCIPALES (DESDE EL 1ER DÍA):</b><br/>• <b>Muerte Natural:</b> 16 Remuneraciones mensuales asegurables.<br/>• <b>Muerte Accidental:</b> 32 Remuneraciones mensuales asegurables.<br/>• <b>Invalidez Total y Permanente por Accidente:</b> 32 Remuneraciones mensuales asegurables.<br/>• <b>Coberturas Adicionales Incluidas:</b> Gastos de sepelio, desgravamen, cáncer, infarto, quemaduras y accidentes graves (según plan de cada compañía hasta los 65 años)."
    cob_der = "<b>BENEFICIARIOS LEGALES Y REGLAS DE PLANILLA:</b><br/>• <b>1er Orden:</b> Cónyuge o conviviente e hijos (sin restricción de edad).<br/>• <b>2do Orden (A falta de los anteriores):</b> Ascendientes (padres) y hermanos menores de 18 años.<br/>• <b>Remuneración Asegurable:</b> Sueldo habitual en planilla hasta el tope máximo del SPP (excluye gratificaciones y utilidades).<br/>• <b>Obligación del Empleador:</b> Recabar la Declaración Jurada de Beneficiarios firmada por cada trabajador."
    t_cob = Table([[Paragraph(cob_izq, st_td), Paragraph(cob_der, st_td)]], colWidths=[9.0*cm, 9.0*cm])
    t_cob.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 6)]))
    elements.extend([t_cob, Spacer(1, 8)])

    excl_txt = "<b>CONSIDERACIONES Y EXCLUSIONES PRINCIPALES:</b><br/>• Queda excluido el suicidio o intento de suicidio, así como lesiones o fallecimiento causados intencionalmente o por acto doloso.<br/>• Seguro nominativo: requiere el envío de la trama/relación nominal de trabajadores (DNI, nombres, fecha de nacimiento y sueldo) antes del inicio de vigencia.<br/>• Las tasas para trabajadores mayores de 70 años son referenciales y están sujetas a evaluación especial de suscripción."
    t_excl = Table([[Paragraph(excl_txt, st_td)]], colWidths=[18.0*cm])
    t_excl.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F0F4F8")), ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 6)]))
    elements.extend([t_excl, Spacer(1, 8)])

    footer_txt = "<b>YQ CORREDORES DE SEGUROS (Código SBS N5410 | RUC 10421076184):</b> <i>Emisión inmediata de pólizas y constancias Vida Ley, gestión de tramas y defensa ante siniestros.</i>"
    elements.append(Paragraph(footer_txt, ParagraphStyle('FT', parent=st_norm, fontSize=7.5, textColor=AZUL_CORP, alignment=1)))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. INTERFAZ DE USUARIO (STREAMLIT)
# ==========================================
df_h1, df_h2 = cargar_tarifario_vidaley()
origen_url = st.query_params.get("origen", "")

if origen_url != "web":
    if os.path.exists("logo_web.png"):
        c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
        with c_l2: st.image("logo_web.png", use_container_width=True)
    st.markdown("""
        <div style="text-align: center; padding: 15px 0px;">
            <h1 style="color: #03318e; font-size: 2.3rem; font-weight: 800; margin-bottom: 8px;">Cotiza y emite tu Seguro Vida Ley al instante.</h1>
            <p style="color: #3c3c3c; font-size: 1.1rem;">Compara tarifas oficiales de <b>Pacífico, Rímac, Mapfre y Protecta</b> desde el primer día de labores.</p>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

codigo_actual = st.session_state.get('codigo_secreto', '')
es_admin = (codigo_actual == CODIGO_ADMIN or codigo_actual in CODIGOS_ASESORES)

if df_h1 is None:
    st.error("❌ No se encontró el archivo 'tarifario_vidaley.xlsx' en el repositorio.")
else:
    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-bottom: 10px;'>1. Datos de la Empresa / Contratante</h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        razon_social = st.text_input("Razón Social o Nombre Completo", placeholder="Ej: Constructora Peruana SAC")
        ruc = st.text_input("RUC o DNI", max_chars=11, placeholder="Ej: 20601234567 o 40123456")
    with col2:
        celular = st.text_input("Celular / WhatsApp", max_chars=9, placeholder="Ej: 999123456")
        email = st.text_input("Correo Electrónico", placeholder="rrhh@empresa.com")

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-top: 15px; margin-bottom: 10px;'>2. Actividad y Perfil de Trabajadores</h3>", unsafe_allow_html=True)
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        actividades_lista = df_h1['ACTIVIDAD_CLEAN'].dropna().unique().tolist()
        actividad_sel = st.selectbox("Rubro / Actividad Económica", actividades_lista)
    with col_a2:
        categorias_lista = ["Empleado (Administrativo / Oficina)", "Obrero (Operativo Estándar)", "Obrero Alto Riesgo (Construcción / Planta / Mina)"]
        idx_cat_def = 2 if any(k in actividad_sel for k in ["Construcción", "Seguridad", "Transporte"]) else (0 if "Oficinas" in actividad_sel else 1)
        categoria_sel = st.selectbox("Categoría de Trabajadores", categorias_lista, index=idx_cat_def)

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-top: 15px; margin-bottom: 5px;'>3. Distribución por Rango de Edad y Planilla</h3>", unsafe_allow_html=True)
    st.caption(f"*Ingresa la cantidad de trabajadores por rango de edad. El sistema sugiere automáticamente la planilla mínima legal (S/ {RMV_VIGENTE:,.2f} por trabajador), la cual puedes editar si ganan más.*")

    trabajadores_rango, planilla_rango = {}, {}
    cols_r = st.columns(3)
    for i, r_edad in enumerate(["18-36", "37-55", "56-63"]):
        with cols_r[i]:
            n_t = st.number_input(f"N° Trab. ({r_edad})", min_value=0, max_value=5000, value=1 if i==0 else 0, key=f"nt_{r_edad}")
            trabajadores_rango[r_edad] = n_t
            planilla_rango[r_edad] = st.number_input(f"Planilla {r_edad} (S/)", min_value=float(n_t*RMV_VIGENTE), value=float(n_t*RMV_VIGENTE), step=100.0) if n_t > 0 else 0.0

    with st.expander("➕ ¿Tienes trabajadores mayores de 64 años? (Clic aquí para agregarlos)"):
        cols_s = st.columns(3)
        for i, r_edad in enumerate(["64-70", "71-80", "81-110"]):
            with cols_s[i]:
                n_t = st.number_input(f"N° Trab. ({r_edad})", min_value=0, max_value=1000, value=0, key=f"nt_{r_edad}")
                trabajadores_rango[r_edad] = n_t
                planilla_rango[r_edad] = st.number_input(f"Planilla {r_edad} (S/)", min_value=float(n_t*RMV_VIGENTE), value=float(n_t*RMV_VIGENTE), step=100.0) if n_t > 0 else 0.0

    total_trab, total_planilla = sum(trabajadores_rango.values()), sum(planilla_rango.values())
    if total_trab > 0: st.info(f"📊 **Resumen a cotizar:** {total_trab} trabajador(es) | **Planilla Mensual Total:** S/ {total_planilla:,.2f}")

    msg_esp = "Hola YQ, estoy en el cotizador de Vida Ley y deseo cotizar una actividad especial o planilla corporativa."
    st.markdown(f"<div style='font-size: 0.88rem; color: #555;'>💡 ¿Tu empresa realiza actividades mineras, petróleo o no encuentras tu giro? <a href='https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_esp)}' target='_blank' style='color: #089685; font-weight: 700; text-decoration: underline;'>Cotiza directo por WhatsApp aquí</a>.</div><br>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # LOS DOS BOTONES AZULES PRINCIPALES (UX HOMOGENEIZADA)
    # ---------------------------------------------------------
    col_btn1, col_btn2 = st.columns(2)
    generar_vl = col_btn1.button("Generar Propuesta (Solo Vida Ley)", type="primary", use_container_width=True)
    generar_pack = col_btn2.button("Generar Pack (Vida Ley + SCTR)", type="primary", use_container_width=True)

    if generar_vl or generar_pack:
        celular_limpio, correo_limpio, ruc_limpio = "".join(filter(str.isdigit, str(celular))), email.strip(), ruc.strip()
        es_ruc, es_dni = bool(re.match(r'^(10|20)\d{9}$', ruc_limpio)), bool(re.match(r'^\d{8}$', ruc_limpio))

        if not razon_social.strip(): st.error("⚠️ Por favor ingresa la Razón Social o Nombre del Contratante.")
        elif total_trab <= 0: st.error("⚠️ Debes ingresar al menos 1 trabajador en algún rango de edad.")
        elif not es_admin and (not ruc_limpio or not (es_ruc or es_dni)): st.error("🏢 Ingresa un RUC válido (11 dígitos) o DNI válido (8 dígitos).")
        elif not es_admin and (len(celular_limpio) != 9 or not celular_limpio.startswith('9')): st.error("📱 Por favor ingresa un número de celular válido de 9 dígitos.")
        elif not es_admin and (not correo_limpio or not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", correo_limpio)): st.error("📧 Por favor ingresa un correo electrónico real.")
        else:
            df_res = cotizar_vidaley(actividad_sel, categoria_sel, planilla_rango)
            
            # Limpiamos el texto de SCTR si eligen solo Vida Ley, o calculamos SCTR si eligen Pack
            df_sctr_addon = None
            if generar_pack:
                df_s_b, df_s_m = cargar_tarifario_sctr_addon()
                # Como los nombres de actividad son idénticos en ambos Excels, pasamos actividad_sel directo
                if df_s_b is not None:
                    # Obrero Alto Riesgo es la asunción por defecto para construcción en SCTR, ajusta si es necesario
                    df_sctr_addon = calcular_sctr_addon(df_s_b, df_s_m, actividad_sel, total_planilla)
            
            resumen_ed = ", ".join([f"{k} años: {v} trab. (S/{planilla_rango[k]:,.0f})" for k, v in trabajadores_rango.items() if v > 0])
            st.session_state['res_vl'] = df_res
            st.session_state['res_sctr_addon'] = df_sctr_addon
            st.session_state['datos_vl'] = {
                "razon_social": razon_social.strip(), "ruc": ruc_limpio if ruc_limpio else "Por confirmar", "celular": celular_limpio, "email": correo_limpio,
                "actividad": actividad_sel, "categoria": categoria_sel.split(" (")[0], "trabajadores": total_trab, "planilla": total_planilla,
                "resumen_edades": resumen_ed, "tiene_mayores_70": (trabajadores_rango.get("71-80", 0) + trabajadores_rango.get("81-110", 0)) > 0
            }
            st.session_state['id_vl'] = f"VL-{obtener_hora_peru().strftime('%m%d-%H%M')}"
            
            if not df_res.empty and not es_admin:
                enviar_notificacion_vidaley(razon_social.strip(), ruc_limpio, celular_limpio, correo_limpio, actividad_sel, categoria_sel, total_trab, total_planilla, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Bruto_Anual'], df_res.iloc[0]['Bruto_Mensual'])
                guardar_en_sheets([obtener_hora_peru().strftime('%Y-%m-%d'), obtener_hora_peru().strftime('%H:%M:%S'), st.session_state['id_vl'], razon_social.strip(), ruc_limpio, celular_limpio, correo_limpio, actividad_sel, categoria_sel, total_trab, total_planilla, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Bruto_Anual'], df_res.iloc[0]['Bruto_Mensual']])

    if st.session_state.get('res_vl') is not None:
        df_res = st.session_state['res_vl'].copy()
        datos = st.session_state['datos_vl']
        df_sctr_opcional = st.session_state.get('res_sctr_addon')
        folio = st.session_state['id_vl']
        
        if df_res.empty: st.warning("No se encontraron tasas disponibles para esta combinación.")
        else:
            if datos.get("tiene_mayores_70"): st.warning("⚠️ **Nota Importante:** Has incluido trabajadores mayores de 70 años. Sus tasas son referenciales.")

            opciones_cias = df_res['Aseguradora'].tolist()
            top_row = df_res.iloc[0]
            ahorro_anual = max(0.0, (top_row['Bruto_Mensual'] * 12) - top_row['Bruto_Anual'])

            # --- MODO ADMIN ---
            if es_admin:
                st.success("🔓 MODO CORREDOR ACTIVADO - Configura la recomendación para el cliente:")
                mejor_opcion_elegida = st.selectbox("Selecciona la Aseguradora a destacar como 'Mejor Opción':", opciones_cias, index=0)
                df_res['es_top'] = df_res['Aseguradora'] == mejor_opcion_elegida
                df_res = df_res.sort_values(by=['es_top', 'Bruto_Anual'], ascending=[False, True]).drop(columns=['es_top']).reset_index(drop=True)
                top_row = df_res.iloc[0]
                ahorro_anual = max(0.0, (top_row['Bruto_Mensual'] * 12) - top_row['Bruto_Anual'])
                
                txt_ahorro = f" Te sugerimos elegir la modalidad de Pago Anual Único (S/ {top_row['Bruto_Anual']:,.2f}), con la cual tu empresa ahorra S/ {ahorro_anual:,.2f} al año frente a realizar 12 emisiones mensuales." if ahorro_anual > 10 else f" Si buscas la cuota mensual más baja para flujo de caja inmediato, también dispones de opciones desde S/ {df_res['Bruto_Mensual'].min():,.2f} al mes."
                razon_default = f"Nuestro algoritmo ha comparado las tasas por rango etario y primas mínimas de ley para {datos['actividad']}. Destacamos a {mejor_opcion_elegida} como la alternativa de mayor eficiencia financiera.{txt_ahorro}"
                razon_admin = st.text_area("Motivo (Análisis del Experto para el PDF):", value=razon_default)
            else:
                mejor_opcion_elegida = opciones_cias[0]
                txt_ahorro = f" Te sugerimos elegir la modalidad de Pago Anual Único (S/ {top_row['Bruto_Anual']:,.2f}), con la cual tu empresa ahorra S/ {ahorro_anual:,.2f} al año frente a realizar 12 emisiones mensuales." if ahorro_anual > 10 else f" Si buscas la cuota mensual más baja para flujo de caja inmediato, también dispones de opciones desde S/ {df_res['Bruto_Mensual'].min():,.2f} al mes."
                razon_admin = f"Nuestro algoritmo ha comparado las tasas por rango etario y primas mínimas de ley para {datos['actividad']}. Destacamos a {mejor_opcion_elegida} como la alternativa de mayor eficiencia financiera.{txt_ahorro}"

            st.markdown("---")
            if df_sctr_opcional is not None and not df_sctr_opcional.empty:
                st.success(f"¡Cotización PACK CORPORATIVO lista! Opción recomendada por eficiencia anual: **{mejor_opcion_elegida}**")
                df_vista = df_res[['Aseguradora', 'Bruto_Mensual', 'Bruto_Trimestral', 'Bruto_Semestral', 'Bruto_Anual']].copy()
                df_vista.columns = ['Aseguradora', 'VL Mensual + SCTR', 'VL Trim. + SCTR', 'VL Sem. + SCTR', 'VL Anual + SCTR']
                
                # Sumamos la cuota mensual del SCTR a cada columna de Vida Ley para la vista web
                for i, r_vl in df_vista.iterrows():
                    cia_vl = r_vl['Aseguradora']
                    cia_sctr_search = "Sanitas + Protecta" if cia_vl == "Protecta" else cia_vl
                    match_sctr = df_sctr_opcional[df_sctr_opcional['Aseguradora SCTR'] == cia_sctr_search]
                    sctr_mensual = match_sctr.iloc[0]['Total Mensual SCTR'] if not match_sctr.empty else 0.0
                    
                    df_vista.at[i, 'VL Mensual + SCTR'] = f"S/ {r_vl['VL Mensual + SCTR'] + sctr_mensual:,.2f}"
                    df_vista.at[i, 'VL Trim. + SCTR'] = f"S/ {r_vl['VL Trim. + SCTR'] + (sctr_mensual*3):,.2f}"
                    df_vista.at[i, 'VL Sem. + SCTR'] = f"S/ {r_vl['VL Sem. + SCTR'] + (sctr_mensual*6):,.2f}"
                    df_vista.at[i, 'VL Anual + SCTR'] = f"S/ {r_vl['VL Anual + SCTR'] + (sctr_mensual*12):,.2f}"
                st.dataframe(df_vista, hide_index=True, use_container_width=True)
                st.caption("*Nota: El SCTR se factura siempre de manera mensual por normativa. Los montos sumados reflejan el consolidado.*")
            else:
                st.success(f"¡Cotización Vida Ley lista! Opción recomendada por eficiencia anual: **{mejor_opcion_elegida}**")
                df_vista = df_res[['Aseguradora', 'Bruto_Mensual', 'Bruto_Trimestral', 'Bruto_Semestral', 'Bruto_Anual']].copy()
                df_vista.columns = ['Aseguradora', 'Mensual c/IGV', 'Trimestral c/IGV', 'Semestral c/IGV', 'Anual c/IGV (Recomendado)']
                for col_m in df_vista.columns[1:]: df_vista[col_m] = df_vista[col_m].apply(lambda x: f"S/ {x:,.2f}")
                st.dataframe(df_vista, hide_index=True, use_container_width=True)

            pdf_buffer = generar_pdf_vidaley(datos, df_res, folio, razon_admin, df_sctr_opcional)
            nombre_clean = re.sub(r'[^a-zA-Z0-9]', '', datos['razon_social'])[:15] or "Empresa"

            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.download_button(
                    label="Descargar Slip (PDF)" if df_sctr_opcional is None else "Descargar Slip Pack (Vida Ley + SCTR)",
                    data=pdf_buffer,
                    file_name=f"SLIP_YQ_{nombre_clean}_{obtener_hora_peru().strftime('%d%m%y')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col_b2:
                if df_sctr_opcional is not None:
                    cia_sctr_search = "Sanitas + Protecta" if mejor_opcion_elegida == "Protecta" else mejor_opcion_elegida
                    match_sctr = df_sctr_opcional[df_sctr_opcional['Aseguradora SCTR'] == cia_sctr_search]
                    sctr_mensual = match_sctr.iloc[0]['Total Mensual SCTR'] if not match_sctr.empty else 0.0
                    msg_wa = f"Hola YQ, deseo emitir el Pack Corporativo ({folio}) para {datos['razon_social']}: Vida Ley con {mejor_opcion_elegida} (Anual S/{top_row['Bruto_Anual']:,.2f}) y SCTR por S/{sctr_mensual:,.2f} al mes."
                else:
                    msg_wa = f"Hola YQ, deseo emitir mi póliza Vida Ley ({folio}) para {datos['razon_social']} con {mejor_opcion_elegida}."
                
                enlace_wa = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_wa)}"
                st.markdown(f"""
                    <a href='{enlace_wa}' target='_blank' style='display: flex; align-items: center; justify-content: center; width: 100%; height: 48px; background-color: #25D366; color: white; border-radius: 8px; text-decoration: none; font-weight: 700; font-size: 16px; box-sizing: border-box;'>
                        Emitir Póliza por WhatsApp
                    </a>
                """, unsafe_allow_html=True)

with st.expander("🛡️ Acceso Interno YQ (Solo Empleados)"):
    st.text_input("Código de Autorización", type="password", placeholder="Ingresa clave", key="codigo_secreto")
    if es_admin and st.button("📥 Descargar Historial Vida Ley (CSV)"):
        df_hist = descargar_historial_sheets()
        if df_hist is not None and not df_hist.empty:
            st.download_button("💾 Guardar Archivo CSV", df_hist.to_csv(index=False).encode('utf-8-sig'), f"Historial_VidaLey_{obtener_hora_peru().strftime('%d%m%Y')}.csv", "text/csv")
