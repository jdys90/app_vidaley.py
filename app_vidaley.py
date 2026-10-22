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
        footer, .stApp > footer, [data-testid="stBottom"], [data-testid="stBottomBar"], [data-testid="stViewerBadge"] { display: none !important; visibility: hidden !important; }
        .block-container { padding-bottom: 0rem !important; }
        div.stButton > button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; font-size: 16px !important; font-weight: 700 !important; }
        div.stButton > button:hover { background-color: #022266 !important; border-color: #022266 !important; }
        div[data-testid="stDownloadButton"] button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; width: 100% !important; font-size: 16px !important; font-weight: 700 !important; }
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
# 3. MOTORES DE CÁLCULO (VIDA LEY Y SCTR)
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
            imp_serie = df_h2[df_h2['Concepto_Clean'].str.contains('Impuesto', case=False, na=False)][cia]
            imp = float(imp_serie.iloc[0]) if not imp_serie.empty and pd.notna(imp_serie.iloc[0]) else 0.18
            def get_min(palabra):
                s = df_h2[df_h2['Concepto_Clean'].str.contains(palabra, case=False, na=False)][cia]
                return float(s.iloc[0]) if not s.empty and pd.notna(s.iloc[0]) else 0.0
            min_m, min_t, min_s, min_a = get_min('Mensual'), get_min('Trimestral'), get_min('Semestral'), get_min('Anual')
        except Exception: continue

        prima_neta_mensual_calc = sum(monto * float(sub_df[sub_df['RANGO_CLEAN'] == rango].iloc[0][col_tasa]) for rango, monto in planilla_por_rango.items() if monto > 0 and not sub_df[sub_df['RANGO_CLEAN'] == rango].empty and pd.notna(sub_df[sub_df['RANGO_CLEAN'] == rango].iloc[0][col_tasa]))
        
        if prima_neta_mensual_calc <= 0: continue

        neto_m, neto_t, neto_s, neto_a = max(prima_neta_mensual_calc * 1.0, min_m), max(prima_neta_mensual_calc * 3.0, min_t), max(prima_neta_mensual_calc * 6.0, min_s), max(prima_neta_mensual_calc * 12.0, min_a)

        resultados.append({
            "Aseguradora": cia, "Bruto_Mensual": round(neto_m * (1.0 + imp), 2), "Min_M": (prima_neta_mensual_calc * 1.0) < min_m,
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
        ruta = next((f for f in ['Tarifario_sctr.xlsx', 'tarifario_sctr.xlsx'] if os.path.exists(f)), None)
        if not ruta: return None
        df = pd.read_excel(ruta)
        df.columns = df.columns.str.strip()
        return df
    except Exception: return None

def calcular_sctr_addon(df_sctr, rubro_sel, act_sel, planilla_total):
    fila = df_sctr[(df_sctr['Rubro'] == rubro_sel) & (df_sctr['Actividad'] == act_sel)]
    if fila.empty: return pd.DataFrame()
    row = fila.iloc[0]
    
    cias = [("Rímac", "Tasa_Salud_Rimac", "Min_Salud_Rimac", "Tasa_Pension_Rimac", "Min_Pension_Rimac", 1.03),
            ("Pacífico", "Tasa_Salud_Pacifico", "Min_Salud_Pacifico", "Tasa_Pension_Pacifico", "Min_Pension_Pacifico", 1.03),
            ("Mapfre", "Tasa_Salud_Mapfre", "Min_Salud_Mapfre", "Tasa_Pension_Mapfre", "Min_Pension_Mapfre", 1.03),
            ("Sanitas + Protecta", "Tasa_Salud_Sanitas", "Min_Salud_Sanitas", "Tasa_Pension_Protecta", "Min_Pension_Protecta", 1.00)]
    res = []
    for nombre, col_ts, col_ms, col_tp, col_mp, der_emision in cias:
        if col_ts in row and pd.notna(row[col_ts]):
            t_sal, m_sal = float(row[col_ts]), float(row[col_ms]) if col_ms in row and pd.notna(row[col_ms]) else 0.0
            t_pen, m_pen = float(row[col_tp]) if col_tp in row and pd.notna(row[col_tp]) else 0.0, float(row[col_mp]) if col_mp in row and pd.notna(row[col_mp]) else 0.0
            
            factor_s, factor_p = (t_sal / 100.0) if t_sal > 0.1 else t_sal, (t_pen / 100.0) if t_pen > 0.1 else t_pen
            neto_sal, neto_pen = max(planilla_total * factor_s, m_sal), max(planilla_total * factor_p, m_pen)
            bruto_sal, bruto_pen = round(neto_sal * 1.18, 2), round(neto_pen * der_emision * 1.18, 2)
            
            res.append({"Aseguradora SCTR": nombre, "Salud c/IGV": bruto_sal, "Pensión c/IGV": bruto_pen, "Total Mensual SCTR": round(bruto_sal + bruto_pen, 2), "Total Anual SCTR": round((bruto_sal + bruto_pen) * 12, 2)})
    df_out = pd.DataFrame(res)
    if not df_out.empty: df_out = df_out.sort_values(by="Total Mensual SCTR", ascending=True).reset_index(drop=True)
    return df_out

# ==========================================
# 4. MOTOR GENERADOR DE PDF
# ==========================================
def generar_pdf_vidaley(datos, df_res, folio, razon_experto=""):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15, leftMargin=15, topMargin=16, bottomMargin=16)
    estilos = getSampleStyleSheet()

    AZUL_CORP, GRIS_TEXTO, GRIS_FONDO, BORDE_SUAVE = colors.HexColor("#03318e"), colors.HexColor("#3c3c3c"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")
    st_tit = ParagraphStyle('T', parent=estilos['Heading1'], fontName='Helvetica-Bold', fontSize=13, textColor=AZUL_CORP, leading=15)
    st_sub = ParagraphStyle('S', parent=estilos['Normal'], fontName='Helvetica-Bold', fontSize=9.5, textColor=AZUL_CORP, spaceAfter=4)
    st_norm = ParagraphStyle('N', parent=estilos['Normal'], fontSize=8, textColor=GRIS_TEXTO, leading=10.5)
    st_bold = ParagraphStyle('B', parent=st_norm, fontName='Helvetica-Bold', textColor=AZUL_CORP)
    st_th = ParagraphStyle('TH', parent=estilos['Normal'], fontSize=7.5, fontName='Helvetica-Bold', textColor=colors.white, alignment=1)
    st_td_c = ParagraphStyle('TDC', parent=estilos['Normal'], fontSize=7.5, textColor=GRIS_TEXTO, alignment=1)

    elements = []
    img_yq = ImageRL("logo.png", width=3.8*cm, height=1.9*cm, kind='proportional') if os.path.exists("logo.png") else Paragraph("<b>YQ</b>", st_tit)
    p_header = Paragraph("<b>YQ CORREDORES DE SEGUROS</b><br/>Slip de Cotización Seguro Vida Ley<br/><font size='7.5' color='#666666'>D.L. N° 688 / Ley N° 29549 (Desde el 1er día)</font>", st_tit)
    p_folio = Paragraph(f"<font color='#3c3c3c'><b>Cotización:</b> {folio}<br/><b>Fecha:</b> {obtener_hora_peru().strftime('%d/%m/%Y %H:%M')}</font>", ParagraphStyle('F', parent=st_norm, alignment=2))
    t_head = Table([[img_yq, p_header, p_folio]], colWidths=[3.8*cm, 9.7*cm, 4.5*cm])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0), (-1,-1), 1, AZUL_CORP), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
    elements.extend([t_head, Spacer(1, 8)])

    elements.append(Paragraph("1. INFORMACIÓN DEL CONTRATANTE Y PLANILLA", st_sub))
    t_cli = Table([[Paragraph("<b>Contratante:</b>", st_bold), Paragraph(datos['razon_social'], st_norm), Paragraph("<b>RUC / DNI:</b>", st_bold), Paragraph(datos['ruc'], st_norm)],
                   [Paragraph("<b>Actividad:</b>", st_bold), Paragraph(datos['actividad'], st_norm), Paragraph("<b>Trabajadores:</b>", st_bold), Paragraph(f"{datos['trabajadores']} ({datos['categoria']})", st_norm)],
                   [Paragraph("<b>Edades:</b>", st_bold), Paragraph(datos['resumen_edades'], st_norm), Paragraph("<b>Planilla Total:</b>", st_bold), Paragraph(f"<b>S/ {datos['planilla']:,.2f}</b>", st_norm)]], colWidths=[2.5*cm, 8.0*cm, 3.0*cm, 4.5*cm])
    t_cli.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 4)]))
    elements.extend([t_cli, Spacer(1, 10)])

    elements.append(Paragraph("2. PROPUESTA ECONÓMICA (MONTOS FINALES CON IGV)", st_sub))
    data_tabla = [[Paragraph(h, st_th) for h in ['Aseguradora', 'Mensual\n(c/IGV)', 'Trimestral\n(c/IGV)', 'Semestral\n(c/IGV)', 'Anual Único\n(c/IGV)']]]
    for idx, row in df_res.iterrows():
        cia = str(row['Aseguradora'])
        mapa_logos = {"Rímac": "logo_rimac.png", "Pacífico": "logo_pacifico.png", "Mapfre": "logo_mapfre.png", "Protecta": "logo_protecta.png"}
        ruta_logo = mapa_logos.get(cia, "")
        if ruta_logo and os.path.exists(ruta_logo):
            celda_cia = Table([[ImageRL(ruta_logo, width=2.3*cm, height=0.75*cm, kind='proportional')]], colWidths=[3.0*cm])
            celda_cia.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        else: celda_cia = Paragraph(f"<b>{cia}</b>", st_td_c)

        data_tabla.append([celda_cia, Paragraph(f"S/ {row['Bruto_Mensual']:,.2f}", st_td_c), Paragraph(f"S/ {row['Bruto_Trimestral']:,.2f}", st_td_c), Paragraph(f"S/ {row['Bruto_Semestral']:,.2f}", st_td_c), Paragraph(f"<font color='#089685' size='8'><b>S/ {row['Bruto_Anual']:,.2f}</b></font>", st_td_c)])
    
    t_ofertas = Table(data_tabla, colWidths=[4.0*cm, 3.5*cm, 3.5*cm, 3.5*cm, 3.5*cm], repeatRows=1)
    estilos_tabla = [('BACKGROUND', (0,0), (-1,0), AZUL_CORP), ('GRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 5)]
    if len(df_res) > 0: estilos_tabla.extend([('BACKGROUND', (0,1), (-1,1), colors.HexColor("#F0F6FF")), ('BOX', (0,1), (-1,1), 1.5, AZUL_CORP)])
    t_ofertas.setStyle(TableStyle(estilos_tabla))
    elements.extend([t_ofertas, Spacer(1, 8)])

    if razon_experto:
        elements.append(Paragraph("3. ANÁLISIS DEL EXPERTO", st_sub))
        t_box = Table([[Paragraph(f"{razon_experto}", st_norm)]], colWidths=[18.0*cm])
        t_box.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 6)]))
        elements.append(t_box)

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
        <div style="text-align: center; padding: 10px 0px;">
            <h1 style="color: #03318e; font-size: 2.3rem; font-weight: 800; margin-bottom: 5px;">Cotiza tu Seguro Vida Ley al instante.</h1>
            <p style="color: #3c3c3c; font-size: 1.1rem;">Compara Pacífico, Rímac, Mapfre y Protecta.</p>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

codigo_actual = st.session_state.get('codigo_secreto', '')
es_admin = (codigo_actual == CODIGO_ADMIN or codigo_actual in CODIGOS_ASESORES)

if df_h1 is None:
    st.error("❌ No se encontró el archivo 'tarifario_vidaley.xlsx'.")
else:
    razon_url, ruc_url, cel_url, mail_url = st.query_params.get("razon", ""), st.query_params.get("ruc", ""), st.query_params.get("cel", ""), st.query_params.get("mail", "")
    try: trab_default = max(1, int(st.query_params.get("trab", "1")))
    except: trab_default = 1
    if razon_url: st.info(f"⚡ **¡Hola {razon_url}!** Datos precargados correctamente.")

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-bottom: 10px;'>1. Datos de la Empresa</h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: razon_social, ruc = st.text_input("Razón Social", value=razon_url, placeholder="Constructora SAC"), st.text_input("RUC o DNI", value=ruc_url, max_chars=11)
    with col2: celular, email = st.text_input("Celular", value=cel_url, max_chars=9), st.text_input("Correo", value=mail_url)

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-top: 15px;'>2. Actividad y Planilla</h3>", unsafe_allow_html=True)
    col_a1, col_a2 = st.columns(2)
    with col_a1: actividad_sel = st.selectbox("Rubro", df_h1['ACTIVIDAD_CLEAN'].dropna().unique().tolist())
    with col_a2: categoria_sel = st.selectbox("Categoría", ["Empleado (Administrativo / Oficina)", "Obrero (Operativo Estándar)", "Obrero Alto Riesgo (Construcción / Planta / Mina)"], index=2 if any(k in actividad_sel for k in ["Construcción", "Seguridad", "Transporte"]) else 0)

    trabajadores_rango, planilla_rango = {}, {}
    cols_r = st.columns(3)
    for i, r_edad in enumerate(["18-36", "37-55", "56-63"]):
        with cols_r[i]:
            n_t = st.number_input(f"Trab. {r_edad}", min_value=0, max_value=5000, value=trab_default if i==0 else 0)
            trabajadores_rango[r_edad], planilla_rango[r_edad] = n_t, st.number_input(f"Planilla {r_edad}", min_value=float(n_t*RMV_VIGENTE), value=float(n_t*RMV_VIGENTE), step=100.0) if n_t > 0 else 0.0
    
    total_trab, total_planilla = sum(trabajadores_rango.values()), sum(planilla_rango.values())
    if total_trab > 0: st.info(f"📊 **Planilla Total:** S/ {total_planilla:,.2f} ({total_trab} asegurados).")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Generar Propuesta Vida Ley", type="primary", use_container_width=True):
        cel_cl, ruc_cl, mail_cl = "".join(filter(str.isdigit, str(celular))), ruc.strip(), email.strip()
        if not razon_social.strip(): st.error("⚠️ Ingresa la Razón Social.")
        elif total_trab <= 0: st.error("⚠️ Ingresa al menos 1 trabajador.")
        elif not es_admin and (not ruc_cl or len(ruc_cl) not in [8, 11]): st.error("🏢 RUC/DNI inválido.")
        else:
            df_res = cotizar_vidaley(actividad_sel, categoria_sel, planilla_rango)
            st.session_state['res_vl'], st.session_state['id_vl'] = df_res, f"VL-{obtener_hora_peru().strftime('%m%d-%H%M')}"
            st.session_state['datos_vl'] = {"razon_social": razon_social.strip(), "ruc": ruc_cl, "celular": cel_cl, "email": mail_cl, "actividad": actividad_sel, "categoria": categoria_sel.split(" (")[0], "trabajadores": total_trab, "planilla": total_planilla, "resumen_edades": ", ".join([f"{k} años: {v} trab." for k, v in trabajadores_rango.items() if v > 0])}
            st.session_state['ver_sctr'] = False # Resetear panel SCTR al cotizar
            if not df_res.empty and not es_admin:
                enviar_notificacion_vidaley(razon_social.strip(), ruc_cl, cel_cl, mail_cl, actividad_sel, categoria_sel, total_trab, total_planilla, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Bruto_Anual'], df_res.iloc[0]['Bruto_Mensual'])
                guardar_en_sheets([obtener_hora_peru().strftime('%Y-%m-%d'), obtener_hora_peru().strftime('%H:%M:%S'), st.session_state['id_vl'], razon_social.strip(), ruc_cl, cel_cl, mail_cl, actividad_sel, categoria_sel, total_trab, total_planilla, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Bruto_Anual'], df_res.iloc[0]['Bruto_Mensual']])

    if st.session_state.get('res_vl') is not None:
        df_res, datos, folio = st.session_state['res_vl'], st.session_state['datos_vl'], st.session_state['id_vl']
        if df_res.empty: st.warning("No hay tasas para esta combinación.")
        else:
            mejor_cia = df_res.iloc[0]['Aseguradora']
            ahorro = max(0, (df_res.iloc[0]['Bruto_Mensual']*12) - df_res.iloc[0]['Bruto_Anual'])
            razon = f"Recomendamos {mejor_cia} por eficiencia financiera." + (f" Al elegir Pago Anual, ahorras S/ {ahorro:,.2f} frente a pagos mensuales." if ahorro > 10 else "")

            st.markdown("---")
            st.success(f"¡Cotización lista! Opción recomendada: **{mejor_cia}**")
            
            df_vista = df_res[['Aseguradora', 'Bruto_Mensual', 'Bruto_Trimestral', 'Bruto_Semestral', 'Bruto_Anual']].copy()
            df_vista.columns = ['Aseguradora', 'Mensual c/IGV', 'Trimestral', 'Semestral', 'Anual (Recomendado)']
            for c in df_vista.columns[1:]: df_vista[c] = df_vista[c].apply(lambda x: f"S/ {x:,.2f}")
            st.dataframe(df_vista, hide_index=True, use_container_width=True)

            pdf_buf = generar_pdf_vidaley(datos, df_res, folio, razon)
            col_b1, col_b2 = st.columns(2)
            with col_b1: st.download_button("Descargar Slip Vida Ley (PDF)", data=pdf_buf, file_name=f"VIDALEY_{folio}.pdf", mime="application/pdf")
            with col_b2:
                link_wa = f"https://api.whatsapp.com/send?phone=51906462225&text=Hola YQ, deseo emitir Vida Ley ({folio}) con {mejor_cia}."
                st.markdown(f"<a href='{link_wa}' target='_blank' style='display:flex;align-items:center;justify-content:center;width:100%;height:48px;background-color:#25D366;color:white;border-radius:8px;text-decoration:none;font-weight:700;'>Emitir Póliza por WhatsApp</a>", unsafe_allow_html=True)

            # ---------------------------------------------------------
            # BLOQUE INFERIOR: ADD-ON SCTR NATIVO (SIN SALIR DE LA PÁGINA)
            # ---------------------------------------------------------
            df_sctr_base = cargar_tarifario_sctr_addon()
            
            if df_sctr_base is not None:
                st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #F0F6FF 0%, #E6EEFF 100%); border: 2px solid #03318e; border-radius: 12px; padding: 20px; margin-top: 25px;">
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <span style="font-size: 28px;">🏗️</span>
                            <div>
                                <h3 style="color: #03318e; margin: 0; font-weight: 800; font-size: 1.15rem;">¿También te exigen SCTR Salud y Pensión para ingresar a obra o planta?</h3>
                                <p style="color: #3c3c3c; font-size: 0.92rem; margin: 3px 0 0 0;">Cotízalo <b>aquí mismo en 1 clic</b> usando tu misma planilla de <b>S/ {datos['planilla']:,.2f}</b> ({datos['trabajadores']} trabajadores).</p>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                # Checkbox para desplegar
                if 'ver_sctr' not in st.session_state:
                    st.session_state['ver_sctr'] = False
                
                ver_sctr = st.checkbox("➕ Sí, deseo ver y agregar la cotización de SCTR para esta misma planilla", value=st.session_state['ver_sctr'])
                st.session_state['ver_sctr'] = ver_sctr
                
                if ver_sctr:
                    st.markdown("<div style='padding: 15px; border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 10px;'>", unsafe_allow_html=True)
                    c_s1, c_s2 = st.columns(2)
                    with c_s1:
                        rubros_sctr = df_sctr_base['Rubro'].dropna().unique().tolist()
                        rubro_sctr_sel = st.selectbox("Confirma el Giro SCTR:", rubros_sctr, key="addon_rubro_sctr")
                    with c_s2:
                        acts_sctr = df_sctr_base[df_sctr_base['Rubro'] == rubro_sctr_sel]['Actividad'].dropna().unique().tolist()
                        act_sctr_sel = st.selectbox("Actividad Específica (CIIU):", acts_sctr, key="addon_act_sctr")
                    
                    if st.button("Calcular SCTR", use_container_width=True, key="btn_calc_sctr"):
                        df_res_sctr = calcular_sctr_addon(df_sctr_base, rubro_sctr_sel, act_sctr_sel, datos['planilla'])
                        if not df_res_sctr.empty:
                            st.session_state['res_sctr_addon'] = df_res_sctr
                            st.session_state['cia_sctr_addon'] = df_res_sctr.iloc[0]['Aseguradora SCTR']
                            st.session_state['mes_sctr_addon'] = df_res_sctr.iloc[0]['Total Mensual SCTR']
                        else:
                            st.warning("No se encontraron tasas de SCTR para esa actividad.")
                    
                    if st.session_state.get('res_sctr_addon') is not None:
                        df_res_sctr = st.session_state['res_sctr_addon']
                        mejor_sctr_cia = st.session_state['cia_sctr_addon']
                        mejor_sctr_mes = st.session_state['mes_sctr_addon']
                        
                        df_sctr_vista = df_res_sctr.copy()
                        for c_m in ['Salud c/IGV', 'Pensión c/IGV', 'Total Mensual SCTR', 'Total Anual SCTR']:
                            df_sctr_vista[c_m] = df_sctr_vista[c_m].apply(lambda x: f"S/ {x:,.2f}")
                        st.dataframe(df_sctr_vista, hide_index=True, use_container_width=True)
                        
                        mejor_vl_anual = df_res.iloc[0]['Bruto_Anual']
                        mejor_vl_mes = df_res.iloc[0]['Bruto_Mensual']
                        
                        st.info(f"💼 **PACK CORPORATIVO COMPLETO:** Vida Ley ({mejor_cia}) + SCTR ({mejor_sctr_cia}) = **S/ {mejor_vl_mes + mejor_sctr_mes:,.2f} / mes** (o Vida Ley Anual S/ {mejor_vl_anual:,.2f} + SCTR Mensual S/ {mejor_sctr_mes:,.2f}).")
                        
                        msg_pack = f"Hola YQ, acabo de cotizar el PACK CORPORATIVO ({folio}) para {datos['razon_social']} (RUC: {datos['ruc']}, {datos['trabajadores']} trab.): 1) Vida Ley con {mejor_cia} (Anual: S/ {mejor_vl_anual:,.2f} / Mensual: S/ {mejor_vl_mes:,.2f}) + 2) SCTR con {mejor_sctr_cia} (S/ {mejor_sctr_mes:,.2f}/mes). Deseo emitir ambas pólizas."
                        link_wa_pack = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_pack)}"
                        st.markdown(f"<a href='{link_wa_pack}' target='_blank' style='display:flex;align-items:center;justify-content:center;width:100%;height:48px;background-color:#089685;color:white;border-radius:8px;text-decoration:none;font-weight:700;'>📦 Solicitar Pack (Vida Ley + SCTR) por WhatsApp</a>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

with st.expander("🛡️ Acceso Interno YQ (Solo Empleados)"):
    st.text_input("Código de Autorización", type="password", key="codigo_secreto")
    if es_admin and st.button("📥 Descargar Historial CSV"):
        df_hist = descargar_historial_sheets()
        if df_hist is not None and not df_hist.empty: st.download_button("💾 Guardar CSV", df_hist.to_csv(index=False).encode('utf-8-sig'), "Historial_VidaLey.csv", "text/csv")
