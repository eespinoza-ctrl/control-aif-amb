import os
import io
import base64
import requests
import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Control AIF AMB", layout="wide")

# ==========================================
# IDs DE GOOGLE SHEETS / DRIVE
# ==========================================
ID_PACO = "1nzGuTC4c00bfwiCgTL3BCFM7O2MBTBGEOhgJ7VzzuYE"                        # <--- ID PACO AQUÍ
ID_EXTINTORES = "1dvGof6Q2ftsnqvJ3X04U8D4MZTbox8j4E2T6xJcPBTw"            # <--- ID EXTINTORES AQUÍ
ID_MULTAS = "1A3osOqlI6w05vYLeos4WiEL7hJkYqJW-EjIhzVuxzZg"                    # <--- ID MULTAS AQUÍ
ID_CORRESPONDENCIA = "17ql7AXNTGxXkzyv7GaFHPNDQiJ57SMBX"  # <--- ID CORRESPONDENCIA AQUÍ
ID_RESUMEN_AMB = "10vBXajbTph0M9IDJhbd32-0hRCE12gzWH1QSjFEA26M"          # <--- ID RESUMEN AMB AQUÍ

# ==========================================
# RUTAS, AUTENTICACIÓN CLOUD/LOCAL Y LOGO
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_CREDENCIALES = os.path.join(BASE_DIR, "credenciales.json")
RUTA_LOGO_ENCABEZADO = os.path.join(BASE_DIR, "logo_encabezado.png")
RUTA_PIE_PAGINA = os.path.join(BASE_DIR, "pie_pagina.png")
FUENTE_TIPO = "Arial Narrow"

def _obtener_cliente_gspread():
    if "gcp_service_account" in st.secrets:
        return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    return gspread.service_account(filename=RUTA_CREDENCIALES)

def _obtener_credenciales_drive():
    scopes = ["https://www.googleapis.com/auth/drive.readonly"]
    if "gcp_service_account" in st.secrets:
        return Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
    return Credentials.from_service_account_file(RUTA_CREDENCIALES, scopes=scopes)

def mostrar_logo_dashboard(ruta_imagen, ancho_px=110):
    if os.path.exists(ruta_imagen):
        with open(ruta_imagen, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()
        st.markdown(
            f"""
            <style>
                .logo-fijo-amb {{
                    position: absolute;
                    top: 1.5cm;
                    right: 1.5cm;
                    width: {ancho_px}px;
                    z-index: 9999;
                }}
            </style>
            <img src="data:image/png;base64,{img_b64}" class="logo-fijo-amb">
            """,
            unsafe_allow_html=True
        )

mostrar_logo_dashboard(RUTA_LOGO_ENCABEZADO, ancho_px=110)

st.title("✈️ Control AIF AMB")

if st.button("🔄 Actualizar Datos Ahora"):
    st.cache_data.clear()

# ==========================================
# MOTOR UNIFICADO DE REPORTES WORD (FORMATO TIPO AMB)
# ==========================================
def _aplicar_fuente(run, negrita=False, tamano=None, color=(0, 0, 0)):
    run.font.name = FUENTE_TIPO
    run.bold = negrita
    if tamano:
        run.font.size = Pt(tamano)
    run.font.color.rgb = RGBColor(*color)

def _sombrear_celda(celda, hex_color="D9D9D9"):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    celda._tc.get_or_add_tcPr().append(shading)

def _crear_documento_base(titulo_texto):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
        section.header_distance = Cm(1.27)
        section.footer_distance = Cm(1.27)

        if os.path.exists(RUTA_LOGO_ENCABEZADO):
            p_header = section.header.paragraphs[0]
            p_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_header.add_run().add_picture(RUTA_LOGO_ENCABEZADO, width=Cm(2.5))

        if os.path.exists(RUTA_PIE_PAGINA):
            p_footer = section.footer.paragraphs[0]
            p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_footer.add_run().add_picture(RUTA_PIE_PAGINA, width=Cm(15.24))

    p_titulo = doc.add_paragraph(style='Title')
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_titulo = p_titulo.add_run(titulo_texto)
    _aplicar_fuente(r_titulo, negrita=True, tamano=26)
    return doc

def _agregar_titulo_seccion(doc, texto):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(texto)
    _aplicar_fuente(r, negrita=True, tamano=18)
    return p

def _agregar_tabla_antecedentes(doc, filas_datos):
    tabla = doc.add_table(rows=len(filas_datos), cols=2)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    for i, (etiqueta, valor) in enumerate(filas_datos):
        c0 = tabla.cell(i, 0)
        c1 = tabla.cell(i, 1)
        c0.width = Cm(2.59)
        c1.width = Cm(7.62)
        _sombrear_celda(c0, "D9D9D9")
        _aplicar_fuente(c0.paragraphs[0].add_run(str(etiqueta)), negrita=False)
        _aplicar_fuente(c1.paragraphs[0].add_run(str(valor)), negrita=False)
    return tabla

def _insertar_evidencia_fotografica(celda_foto, fotos_raw):
    p_foto = celda_foto.paragraphs[0]
    _aplicar_fuente(p_foto.add_run("EVIDENCIA FOTOGRÁFICA\n"), negrita=True)

    if pd.notna(fotos_raw) and "drive.google.com" in str(fotos_raw):
        lista_fotos = str(fotos_raw).split(',')
        for i, link in enumerate(lista_fotos):
            link = link.strip()
            if i > 0:
                p_foto = celda_foto.add_paragraph()
            try:
                if "id=" in link:
                    file_id = link.split("id=")[-1].split("&")[0]
                elif "/d/" in link:
                    file_id = link.split("/d/")[1].split("/")[0]
                else:
                    file_id = None

                if file_id:
                    url_directa = f"https://drive.google.com/uc?id={file_id}"
                    respuesta = requests.get(url_directa, timeout=15)
                    if respuesta.status_code == 200:
                        run_foto = p_foto.add_run()
                        run_foto.add_picture(io.BytesIO(respuesta.content), width=Inches(2.8))
                    else:
                        _aplicar_fuente(p_foto.add_run(f"\n(No se pudo descargar. Link: {link})"))
                else:
                    _aplicar_fuente(p_foto.add_run(f"\n(Link no reconocido: {link})"))
            except Exception as e:
                _aplicar_fuente(p_foto.add_run(f"\n(Error al procesar foto: {e})"))
    else:
        _aplicar_fuente(p_foto.add_run("\nSin registro fotográfico adjunto."))

def _agregar_bloque_campos(celda, bloques):
    p = celda.paragraphs[0]
    for idx_b, (titulo_b, campos) in enumerate(bloques):
        if idx_b > 0:
            p = celda.add_paragraph()
        _aplicar_fuente(p.add_run(f"{titulo_b}\n"), negrita=True)
        for etiqueta, valor in campos:
            _aplicar_fuente(p.add_run(f"{etiqueta}: "), negrita=True)
            _aplicar_fuente(p.add_run(f"{valor}\n"), negrita=False)

# ==========================================
# CARGA DE DATOS DESDE GOOGLE SHEETS / DRIVE
# ==========================================

# --- 1. FISCALIZACIÓN PACO ---
@st.cache_data(ttl=600)
def cargar_datos_amb():
    gc = _obtener_cliente_gspread()
    sh = gc.open_by_key(ID_PACO)
    worksheet = sh.worksheet("Respuestas de formulario 1")
    datos = worksheet.get_all_records()
    df = pd.DataFrame(datos)

    if df.empty:
        return df

    df.columns = df.columns.astype(str).str.strip()

    col_fecha = next((c for c in df.columns if c.lower() in ["fecha", "marca temporal"]), None)
    if col_fecha:
        fechas_limpias = df[col_fecha].astype(str).str.strip().str.split(" ").str[0]
        fechas_limpias = fechas_limpias.str.replace("/", "-")
        df["Fecha_Filtro"] = pd.to_datetime(fechas_limpias, dayfirst=True, errors="coerce")

    col_ot = next((c for c in df.columns if "numero de ot" in c.lower()), None)
    df["ID_Pantalla"] = [
        "FISC-" + str(i + 1).zfill(3) + " | OT: " + str(ot)
        for i, ot in enumerate(df[col_ot] if col_ot else [""] * len(df))
    ]

    col_cumple = next((c for c in df.columns if "cumple" in c.lower()), None)
    if col_cumple:
        df["Cumple_Limpio"] = df[col_cumple].astype(str).str.strip().str.upper()

    return df

# --- 2. EXTINTORES ---
@st.cache_data(ttl=600)
def cargar_datos_extintores():
    gc = _obtener_cliente_gspread()
    sh = gc.open_by_key(ID_EXTINTORES)
    worksheet = sh.worksheet("Respuestas de formulario 1")
    datos = worksheet.get_all_records()
    df = pd.DataFrame(datos)

    if df.empty:
        return df

    df.columns = df.columns.astype(str).str.strip()

    col_fecha = next((c for c in df.columns if c.lower() in ["fecha", "marca temporal"]), None)
    if col_fecha:
        fechas_limpias = df[col_fecha].astype(str).str.strip().str.split(" ").str[0]
        fechas_limpias = fechas_limpias.str.replace("/", "-")
        df["Fecha_Filtro"] = pd.to_datetime(fechas_limpias, dayfirst=True, errors="coerce")

    col_estado = next((c for c in df.columns if "estado del sistema" in c.lower()), None)
    if col_estado:
        df["Cumple_Limpio"] = df[col_estado].astype(str).str.strip().str.upper()

    df["ID_Pantalla"] = ["EXT-" + str(i + 1).zfill(3) for i in range(len(df))]

    return df

# --- 3. MULTAS AMB (LECTURA NUMÉRICA EXACTA) ---
@st.cache_data(ttl=600)
def cargar_datos_multas():
    gc = _obtener_cliente_gspread()
    sh = gc.open_by_key(ID_MULTAS)
    worksheet = sh.worksheet("Hoja 1")

    datos_crudos = worksheet.get_all_values()
    if not datos_crudos or len(datos_crudos) < 2:
        return pd.DataFrame()

    datos_puros = worksheet.get_all_values(value_render_option='UNFORMATTED_VALUE')

    encabezados_originales = datos_crudos[0]
    filas = datos_crudos[1:]
    filas_puras = datos_puros[1:]

    encabezados_limpios = []
    vistos = {}
    for enc in encabezados_originales:
        enc_str = str(enc).replace('\n', ' ').strip()
        enc_str = " ".join(enc_str.split())
        if enc_str in vistos:
            vistos[enc_str] += 1
            encabezados_limpios.append(f"{enc_str} {vistos[enc_str]}")
        else:
            vistos[enc_str] = 0
            encabezados_limpios.append(enc_str)

    df = pd.DataFrame(filas, columns=encabezados_limpios)

    col_fecha = next((c for c in df.columns if "FECHA" in c.upper()), None)
    if col_fecha:
        df["Fecha_Real"] = pd.to_datetime(df[col_fecha].replace('', pd.NaT), dayfirst=True, errors="coerce")
        df["Mes_Año"] = df["Fecha_Real"].dt.strftime("%Y-%m")

    col_estado = next((c for c in df.columns if c.upper().strip() == "ESTADO"), None)
    if col_estado:
        df["Estado_Limpio"] = df[col_estado].astype(str).str.strip().str.upper()
        df["Estado_Limpio"] = df["Estado_Limpio"].replace(["", "NAN", "NONE"], "SIN ESTADO")
    else:
        df["Estado_Limpio"] = "SIN ESTADO"

    col_rec_sn = next(
        (c for c in df.columns if "RECLAMACIÓN S/N" in c.upper() or "RECLAMACION S/N" in c.upper()),
        None
    )
    if col_rec_sn:
        df["Tiene_Reclamacion"] = df[col_rec_sn].astype(str).str.strip().str.upper().isin(["S", "SI", "SÍ"])
    else:
        df["Tiene_Reclamacion"] = False

    idx_monto = next((i for i, c in enumerate(encabezados_limpios) if "TOTAL UTM" in c.upper()), None)
    if idx_monto is not None:
        montos = []
        for fila_p in filas_puras:
            val = fila_p[idx_monto] if idx_monto < len(fila_p) else 0.0
            if isinstance(val, (int, float)):
                montos.append(float(val))
            else:
                v = str(val).strip().upper().replace('UTM', '').replace('$', '').replace(' ', '')
                if not v or v in ['NAN', 'NONE', '-', 'N/A']:
                    montos.append(0.0)
                else:
                    if '.' in v and ',' in v:
                        v = v.replace('.', '').replace(',', '.') if v.rfind(',') > v.rfind('.') else v.replace(',', '')
                    elif ',' in v:
                        partes = v.split(',')
                        v = v.replace(',', '') if (len(partes[-1]) == 3 and all(p.isdigit() for p in partes)) else v.replace(',', '.')
                    elif '.' in v:
                        partes = v.split('.')
                        if len(partes[-1]) == 3 and all(p.isdigit() for p in partes):
                            v = v.replace('.', '')
                    try:
                        montos.append(float(v))
                    except Exception:
                        montos.append(0.0)
        df["Monto_Numerico"] = montos
    else:
        df["Monto_Numerico"] = 0.0

    return df

# --- 4. CONTROL CORRESPONDENCIA (LECTURA DIRECTA DE .XLSX DESDE GOOGLE DRIVE) ---
@st.cache_data(ttl=600)
def cargar_datos_correspondencia():
    creds = _obtener_credenciales_drive()
    session = AuthorizedSession(creds)

    url = f"https://www.googleapis.com/drive/v3/files/{ID_CORRESPONDENCIA}?alt=media"
    resp = session.get(url, timeout=20)
    resp.raise_for_status()

    df = pd.read_excel(io.BytesIO(resp.content), sheet_name=0, dtype=str).fillna("")
    if df.empty:
        return df

    df.columns = [" ".join(str(c).replace('\n', ' ').strip().split()) for c in df.columns]

    df = df.loc[:, ~df.columns.str.upper().str.startswith("UNNAMED")].copy()
    df = df[df.apply(lambda row: any(str(v).strip() != "" for v in row), axis=1)].copy()

    if "FECHA" in df.columns:
        fechas_limpias = df["FECHA"].astype(str).str.strip().str.split(" ").str[0]
        df["Fecha_Real"] = pd.to_datetime(fechas_limpias.replace('', pd.NaT), dayfirst=True, errors="coerce")
        df["Mes_Año"] = df["Fecha_Real"].dt.strftime("%Y-%m")

        iso = df["Fecha_Real"].dt.isocalendar()
        inicio_sem = df["Fecha_Real"] - pd.to_timedelta(df["Fecha_Real"].dt.weekday, unit="D")
        fin_sem = inicio_sem + pd.Timedelta(days=6)
        df["Semana"] = (
            iso["year"].astype("Int64").astype(str) + "-S" + iso["week"].astype("Int64").astype(str).str.zfill(2)
            + " (" + inicio_sem.dt.strftime("%d/%m") + " al " + fin_sem.dt.strftime("%d/%m") + ")"
        )
        df.loc[df["Fecha_Real"].isna(), "Semana"] = pd.NA

    col_ord = next((c for c in df.columns if "ORDINARIO" in c.upper() and "FECHA" not in c.upper()), None)
    if col_ord:
        df["Estado_Corr"] = df[col_ord].apply(
            lambda x: "CON ORDINARIO" if str(x).strip() not in ["", "-", "N/A", "NAN", "NONE"] else "PENDIENTE"
        )
    else:
        df["Estado_Corr"] = "SIN ESTADO"

    return df

# --- 5. RESUMEN AMB (SOPORTA GOOGLE SHEETS NATIVO Y .XLSX EN DRIVE) ---
def _limpiar_numero_pax(val, corregir_miles_sheet=False):
    if isinstance(val, (int, float)):
        num = float(val)
        if corregir_miles_sheet and not num.is_integer():
            return float(round(num * 1000))
        return num
    v = str(val).strip().replace(' ', '')
    if not v or v.upper() in ['NAN', 'NONE', '-', 'N/A']:
        return 0.0
    if '.' in v and ',' in v:
        v = v.replace('.', '').replace(',', '.') if v.rfind(',') > v.rfind('.') else v.replace(',', '')
    elif '.' in v:
        partes = v.split('.')
        if len(partes[-1]) == 3 and all(p.lstrip('-').isdigit() for p in partes):
            v = v.replace('.', '')
    elif ',' in v:
        partes = v.split(',')
        v = v.replace(',', '') if (len(partes[-1]) == 3 and all(p.lstrip('-').isdigit() for p in partes)) else v.replace(',', '.')
    try:
        num = float(v)
        if corregir_miles_sheet and not num.is_integer():
            return float(round(num * 1000))
        return num
    except Exception:
        return 0.0

def _formatear_mes_anio_pax(val):
    meses_es = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    if isinstance(val, pd.Timestamp):
        return f"{meses_es[val.month - 1]} {val.year}"
    s = str(val).strip()
    try:
        num = float(s)
        if 30000 <= num <= 80000:
            fecha = pd.to_datetime(num, unit="D", origin="1899-12-30")
            return f"{meses_es[fecha.month - 1]} {fecha.year}"
    except Exception:
        pass
    return s

def _mes_a_numero(mes_val):
    s = str(mes_val).strip().upper()
    if s.isdigit():
        return int(s)
    mapa_meses = {
        "ENE": 1, "ENERO": 1,
        "FEB": 2, "FEBRERO": 2,
        "MAR": 3, "MARZO": 3,
        "ABR": 4, "ABRIL": 4,
        "MAY": 5, "MAYO": 5,
        "JUN": 6, "JUNIO": 6,
        "JUL": 7, "JULIO": 7,
        "AGO": 8, "AGOSTO": 8,
        "SEP": 9, "SEPT": 9, "SEPTIEMBRE": 9,
        "OCT": 10, "OCTUBRE": 10,
        "NOV": 11, "NOVIEMBRE": 11,
        "DIC": 12, "DICIEMBRE": 12,
    }
    return mapa_meses.get(s, mapa_meses.get(s[:3], 0))

@st.cache_data(ttl=600)
def _leer_hoja_resumen_amb(nombre_hoja):
    try:
        gc = _obtener_cliente_gspread()
        sh = gc.open_by_key(ID_RESUMEN_AMB)
        worksheet = sh.worksheet(nombre_hoja)
        datos_crudos = worksheet.get_all_values()
        if not datos_crudos or len(datos_crudos) < 2:
            return pd.DataFrame()
        datos_puros = worksheet.get_all_values(value_render_option='UNFORMATTED_VALUE')
        return pd.DataFrame(datos_puros[1:], columns=datos_crudos[0])
    except gspread.exceptions.APIError as e:
        if "400" in str(e):
            creds = _obtener_credenciales_drive()
            session = AuthorizedSession(creds)
            url = f"https://www.googleapis.com/drive/v3/files/{ID_RESUMEN_AMB}?alt=media"
            resp = session.get(url, timeout=20)
            resp.raise_for_status()
            return pd.read_excel(io.BytesIO(resp.content), sheet_name=nombre_hoja)
        raise e

@st.cache_data(ttl=600)
def cargar_datos_resumen_pax(nombre_hoja="PASAJEROS EMBARCADOS DGAC"):
    df = _leer_hoja_resumen_amb(nombre_hoja)
    if df.empty:
        return df

    encabezados_limpios = []
    reg_count = 0
    vistos = {}
    for enc in df.columns:
        enc_str = " ".join(str(enc).replace('\n', ' ').strip().split())
        if "REGULARIZACIONES" in enc_str.upper():
            reg_count += 1
            enc_final = "Regularizaciones Dom. (Pax)" if reg_count == 1 else "Regularizaciones Int. (Pax)"
        elif enc_str in vistos:
            vistos[enc_str] += 1
            enc_final = f"{enc_str} {vistos[enc_str]}"
        else:
            vistos[enc_str] = 0
            enc_final = enc_str
        encabezados_limpios.append(enc_final)

    df.columns = encabezados_limpios
    df = df.loc[:, ~df.columns.str.upper().str.startswith("UNNAMED")].copy()
    df = df[df.apply(lambda row: any(str(v).strip() not in ["", "nan", "None"] for v in row), axis=1)].reset_index(drop=True)

    col_mes = next((c for c in df.columns if "MES" in c.upper()), None)
    if col_mes:
        df[col_mes] = df[col_mes].apply(_formatear_mes_anio_pax)

    cols_texto = [
        c for c in df.columns
        if ("MES" in c.upper() or "CARTA" in c.upper() or "OFICIO" in c.upper())
        and "TOTAL" not in c.upper()
    ]
    for col in df.columns:
        if col in cols_texto:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            es_regularizacion = "REGULARIZACION" in col.upper()
            df[col] = df[col].apply(lambda x: _limpiar_numero_pax(x, corregir_miles_sheet=not es_regularizacion))

    if col_mes:
        df = df[~df[col_mes].str.upper().str.contains("TOTAL|ACUMULADO", na=False)].reset_index(drop=True)
        df["Año_Filtro"] = df[col_mes].astype(str).str.extract(r'(20\d{2})')[0].fillna("Sin Año")
    else:
        df["Año_Filtro"] = "Sin Año"

    return df

@st.cache_data(ttl=600)
def cargar_datos_mano_obra(nombre_hoja="MANO DE OBRA"):
    df = _leer_hoja_resumen_amb(nombre_hoja)
    if df.empty:
        return df

    df.columns = [" ".join(str(c).replace('\n', ' ').strip().split()) for c in df.columns]
    df = df.loc[:, ~df.columns.str.upper().str.startswith("UNNAMED")].copy()
    df = df[df.apply(lambda row: any(str(v).strip() not in ["", "nan", "None"] for v in row), axis=1)].reset_index(drop=True)

    cols_mo = ["M.O CALIFICADA", "M.O SEMICALIFICADA", "M.O NO CALIFICADA"]
    for col_obj in cols_mo:
        col_real = next((c for c in df.columns if c.upper() == col_obj), None)
        df[col_obj] = df[col_real].apply(_limpiar_numero_pax) if col_real else 0.0

    df["TOTAL M.O"] = df["M.O CALIFICADA"] + df["M.O SEMICALIFICADA"] + df["M.O NO CALIFICADA"]

    col_ent = next((c for c in df.columns if c.upper() == "ENTIDAD"), None)
    df["ENTIDAD"] = df[col_ent].fillna("").astype(str).str.strip().replace("", "SIN ENTIDAD") if col_ent else "SIN ENTIDAD"

    col_anio = next((c for c in df.columns if c.upper() in ["AÑO", "ANO"]), None)
    df["AÑO"] = df[col_anio].fillna("").astype(str).str.replace(".0", "", regex=False).str.strip() if col_anio else "Sin Año"

    col_mes = next((c for c in df.columns if c.upper() == "MES"), None)
    df["MES"] = df[col_mes].fillna("").astype(str).str.strip() if col_mes else ""

    df["Mes_Num"] = df["MES"].apply(_mes_a_numero)
    df["Anio_Num"] = pd.to_numeric(df["AÑO"], errors="coerce").fillna(0).astype(int)
    df["Orden_Periodo"] = df["Anio_Num"] * 100 + df["Mes_Num"]
    df["Periodo"] = df["AÑO"] + "-" + df["Mes_Num"].astype(str).str.zfill(2) + " (" + df["MES"] + ")"

    return df.sort_values("Orden_Periodo").reset_index(drop=True)

# ==========================================
# FUNCIONES GENERADORAS DE INFORMES WORD
# ==========================================

def generar_informe_word(fila):
    doc = _crear_documento_base('Informe de Fiscalización en Terreno - AIF Aeropuerto AMB.')

    _agregar_titulo_seccion(doc, '1. Antecedentes Generales')
    fecha_txt = str(fila.get('Fecha', 'N/A')).replace('0:00:00', '').strip()
    hora_txt = str(fila.get('Hora', '')).replace('0:00:00', '').strip()
    _agregar_tabla_antecedentes(doc, [
        ("ID de Registro", fila.get('ID_Pantalla', 'N/A')),
        ("Fecha y Hora", f"{fecha_txt} {hora_txt}".strip()),
        ("Inspector AIF", fila.get('Fiscalizador', 'N/A')),
    ])

    _agregar_titulo_seccion(doc, '2. Detalle del Hallazgo y Evidencia')
    tabla2 = doc.add_table(rows=1, cols=2)
    tabla2.style = 'Table Grid'
    tabla2.autofit = False
    tabla2.cell(0, 0).width = Cm(7.62)
    tabla2.cell(0, 1).width = Cm(7.62)

    _agregar_bloque_campos(tabla2.cell(0, 0), [
        ("UBICACIÓN Y ACTIVO", [
            ("Sector / Espigón", f"{fila.get('Sector fiscalizado', 'N/A')} / {fila.get('Espigón', '')}"),
            ("OT", fila.get('Numero de OT', 'N/A')),
            ("Activo", fila.get('Código del activo', '')),
            ("Grupo", fila.get('Grupo de Trabajadores', '')),
        ]),
        ("HALLAZGO CONTRACTUAL", [
            ("Estado (SLA)", fila.get('Cumple_Limpio', 'N/A')),
            ("Detalle", fila.get('Detalle Observación', 'Sin observaciones')),
        ]),
    ])

    fotos_raw = fila.get('Registro Fotográfico (hasta 10 fotografias)', '')
    _insertar_evidencia_fotografica(tabla2.cell(0, 1), fotos_raw)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generar_informe_word_extintor(fila, col_foto):
    doc = _crear_documento_base('Informe de Fiscalización de Extintores - AIF Aeropuerto AMB.')

    _agregar_titulo_seccion(doc, '1. Antecedentes Generales')
    _agregar_tabla_antecedentes(doc, [
        ("ID de Registro", fila.get('ID_Pantalla', 'N/A')),
        ("Fecha y Hora", fila.get('Marca temporal', 'N/A')),
        ("Inspector AIF", fila.get('Profesional que realiza la fiscalización', 'N/A')),
    ])

    _agregar_titulo_seccion(doc, '2. Detalle del Hallazgo y Evidencia')
    tabla2 = doc.add_table(rows=1, cols=2)
    tabla2.style = 'Table Grid'
    tabla2.autofit = False
    tabla2.cell(0, 0).width = Cm(7.62)
    tabla2.cell(0, 1).width = Cm(7.62)

    _agregar_bloque_campos(tabla2.cell(0, 0), [
        ("UBICACIÓN", [
            ("Zona", fila.get('Ingresa zona de fiscalización', 'N/A')),
            ("Espigón", fila.get('Ingresa espigón', 'N/A')),
        ]),
        ("HALLAZGO EN EXTINTOR", [
            ("Estado (SLA)", fila.get('Estado del sistema fiscalizado', 'N/A')),
            ("Observación", fila.get('Observación', 'Sin observaciones')),
        ]),
    ])

    fotos_raw = fila.get(col_foto, '') if col_foto else ''
    _insertar_evidencia_fotografica(tabla2.cell(0, 1), fotos_raw)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generar_informe_word_multas(mes_sel, estado_sel, df_resumen, total_pagado, total_pagado_reclamo, cant_pagado_reclamo, total_en_proceso):
    hoy_str = pd.Timestamp.now(tz='America/Santiago').strftime('%d/%m/%Y %H:%M')
    doc = _crear_documento_base('Informe de Control de Multas - AIF Aeropuerto AMB.')

    _agregar_titulo_seccion(doc, '1. Antecedentes Generales')
    _agregar_tabla_antecedentes(doc, [
        ("Período Consultado", mes_sel),
        ("Filtro de Estado", estado_sel),
        ("Fecha de Emisión", hoy_str),
        ("Total Registros", str(len(df_resumen))),
    ])

    _agregar_titulo_seccion(doc, '2. Resumen Financiero (UTM)')
    _agregar_tabla_antecedentes(doc, [
        ("Total UTM Pagadas", f"{total_pagado:,.2f} UTM"),
        ("Pagadas con Reclamación", f"{total_pagado_reclamo:,.2f} UTM ({cant_pagado_reclamo} multas)"),
        ("UTM en Trámite", f"{total_en_proceso:,.2f} UTM"),
    ])

    _agregar_titulo_seccion(doc, '3. Distribución por Estado')
    resumen_estados = (
        df_resumen.groupby("Estado_Limpio", as_index=False)
        .agg(Cantidad=("Monto_Numerico", "count"), Total_UTM=("Monto_Numerico", "sum"))
        .sort_values("Cantidad", ascending=False)
    )
    tabla_est = doc.add_table(rows=len(resumen_estados) + 1, cols=3)
    tabla_est.style = 'Table Grid'
    tabla_est.autofit = False
    encabezados = [("Estado", 6.24), ("N° Multas", 3.5), ("Total UTM", 5.5)]
    for j, (titulo_col, ancho) in enumerate(encabezados):
        celda = tabla_est.cell(0, j)
        celda.width = Cm(ancho)
        _sombrear_celda(celda, "D9D9D9")
        _aplicar_fuente(celda.paragraphs[0].add_run(titulo_col), negrita=True)

    for i, row in enumerate(resumen_estados.itertuples(), start=1):
        valores = [str(row.Estado_Limpio), str(row.Cantidad), f"{row.Total_UTM:,.2f} UTM"]
        for j, (_, ancho) in enumerate(encabezados):
            celda = tabla_est.cell(i, j)
            celda.width = Cm(ancho)
            _aplicar_fuente(celda.paragraphs[0].add_run(valores[j]), negrita=False)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generar_informe_word_correspondencia(periodo_sel, estado_sel, area_sel, df_resumen):
    hoy_str = pd.Timestamp.now(tz='America/Santiago').strftime('%d/%m/%Y %H:%M')
    doc = _crear_documento_base('Informe de Control de Correspondencia - AIF Aeropuerto AMB.')

    con_ord = len(df_resumen[df_resumen["Estado_Corr"] == "CON ORDINARIO"])
    pendientes = len(df_resumen[df_resumen["Estado_Corr"] == "PENDIENTE"])

    _agregar_titulo_seccion(doc, '1. Antecedentes Generales')
    _agregar_tabla_antecedentes(doc, [
        ("Período Consultado", periodo_sel),
        ("Área Filtrada", area_sel),
        ("Estado Filtrado", estado_sel),
        ("Fecha de Emisión", hoy_str),
        ("Total Registros", f"{len(df_resumen)} ({con_ord} con Ordinario / {pendientes} pendientes)"),
    ])

    _agregar_titulo_seccion(doc, '2. Distribución por Área')
    if "AREA" in df_resumen.columns and len(df_resumen) > 0:
        resumen_area = df_resumen["AREA"].replace("", "SIN ÁREA").value_counts().reset_index()
        resumen_area.columns = ["Área", "Cantidad"]

        tabla_area = doc.add_table(rows=len(resumen_area) + 1, cols=2)
        tabla_area.style = 'Table Grid'
        tabla_area.autofit = False
        for j, (titulo_col, ancho) in enumerate([("Área", 10.24), ("N° Documentos", 5.0)]):
            celda = tabla_area.cell(0, j)
            celda.width = Cm(ancho)
            _sombrear_celda(celda, "D9D9D9")
            _aplicar_fuente(celda.paragraphs[0].add_run(titulo_col), negrita=True)

        for i, row in enumerate(resumen_area.itertuples(), start=1):
            for j, (val, ancho) in enumerate([(str(row.Área), 10.24), (str(row.Cantidad), 5.0)]):
                celda = tabla_area.cell(i, j)
                celda.width = Cm(ancho)
                _aplicar_fuente(celda.paragraphs[0].add_run(val), negrita=False)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generar_informe_word_resumen_pax(periodo_sel, df_resumen, tot_dom, tot_int, tot_emb, tot_ajus):
    hoy_str = pd.Timestamp.now(tz='America/Santiago').strftime('%d/%m/%Y %H:%M')
    doc = _crear_documento_base('Informe Resumen AMB - Pasajeros Embarcados (DGAC).')

    _agregar_titulo_seccion(doc, '1. Antecedentes Generales')
    _agregar_tabla_antecedentes(doc, [
        ("Período Consultado", periodo_sel),
        ("Meses Registrados", str(len(df_resumen))),
        ("Fecha de Emisión", hoy_str),
    ])

    _agregar_titulo_seccion(doc, '2. Totales Consolidados de Pasajeros')
    _agregar_tabla_antecedentes(doc, [
        ("Total Doméstico", f"{tot_dom:,.0f} pax".replace(",", ".")),
        ("Total Internacional", f"{tot_int:,.0f} pax".replace(",", ".")),
        ("Total Embarcados Mes", f"{tot_emb:,.0f} pax".replace(",", ".")),
        ("Total Pasajeros Ajustados", f"{tot_ajus:,.0f} pax".replace(",", ".")),
    ])

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# CREACIÓN DE MÓDULOS PRINCIPALES (RESUMEN AMB PRIMERO)
# ==========================================
tab_resumen_amb, tab_paco, tab_extintores, tab_multas, tab_correspondencia = st.tabs([
    "📊 Resumen AMB",
    "✈️ Fiscalización PACO",
    "🧯 Fiscalización de Extintores",
    "⚖️ Control Multas AMB",
    "✉️ Control Correspondencia"
])

# ==========================================
# MÓDULO 1: RESUMEN AMB
# ==========================================
with tab_resumen_amb:
    st.header("📊 Resumen AMB")

    subtab_pax, subtab_mo = st.tabs([
        "🛫 Pasajeros Embarcados (DGAC)",
        "👷 Evolutivo Mano de Obra"
    ])

    # --- SUB-PESTAÑA 1: PASAJEROS EMBARCADOS DGAC ---
    with subtab_pax:
        try:
            df_pax = cargar_datos_resumen_pax("PASAJEROS EMBARCADOS DGAC")

            if df_pax.empty:
                st.warning("La hoja 'PASAJEROS EMBARCADOS DGAC' se conectó correctamente, pero actualmente no tiene registros.")
            else:
                col_mes = next((c for c in df_pax.columns if "MES" in c.upper()), df_pax.columns[0])
                col_oficio = next((c for c in df_pax.columns if "CARTA" in c.upper() or "OFICIO" in c.upper()), None)
                col_tot_dom = next((c for c in df_pax.columns if "TOTAL DOM" in c.upper()), None)
                col_tot_int = next((c for c in df_pax.columns if "TOTAL INT" in c.upper()), None)
                col_tot_emb = next((c for c in df_pax.columns if "EMBARCADOS" in c.upper()), None)
                col_tot_aju = next((c for c in df_pax.columns if "AJUSTADOS" in c.upper()), None)

                # --- DOBLE FILTRO: AÑO Y MES ---
                f_pax1, f_pax2 = st.columns(2)
                with f_pax1:
                    anios_disp = sorted(df_pax["Año_Filtro"].dropna().unique().tolist(), reverse=True)
                    anio_sel = st.selectbox(
                        "📅 Filtrar por Año:",
                        ["Todos los Años"] + anios_disp,
                        key="anio_resumen_pax"
                    )

                if anio_sel != "Todos los Años":
                    df_base_meses = df_pax[df_pax["Año_Filtro"] == anio_sel].copy()
                else:
                    df_base_meses = df_pax.copy()

                with f_pax2:
                    meses_disp = df_base_meses[col_mes].dropna().unique().tolist()[::-1]
                    mes_pax_sel = st.selectbox(
                        "📆 Filtrar por Mes / Período:",
                        ["Todos los Meses (Acumulado)"] + meses_disp,
                        key="mes_resumen_pax"
                    )

                if mes_pax_sel != "Todos los Meses (Acumulado)":
                    df_pax_filt = df_base_meses[df_base_meses[col_mes] == mes_pax_sel].copy()
                    etiqueta_periodo_pax = mes_pax_sel
                else:
                    df_pax_filt = df_base_meses.copy()
                    etiqueta_periodo_pax = "Acumulado Histórico" if anio_sel == "Todos los Años" else f"Año {anio_sel}"

                # --- BLOQUE KPI 1: TOTAL HISTÓRICO / PERÍODO FILTRADO ---
                tot_dom = pd.to_numeric(df_pax_filt[col_tot_dom], errors="coerce").fillna(0).sum() if col_tot_dom else 0.0
                tot_int = pd.to_numeric(df_pax_filt[col_tot_int], errors="coerce").fillna(0).sum() if col_tot_int else 0.0
                tot_emb = pd.to_numeric(df_pax_filt[col_tot_emb], errors="coerce").fillna(0).sum() if col_tot_emb else 0.0
                tot_aju = pd.to_numeric(df_pax_filt[col_tot_aju], errors="coerce").fillna(0).sum() if col_tot_aju else 0.0

                st.subheader(f"📊 Total Acumulado ({etiqueta_periodo_pax})")
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("🇨🇱 Total Doméstico", f"{tot_dom:,.0f}".replace(",", "."))
                k2.metric("🌎 Total Internacional", f"{tot_int:,.0f}".replace(",", "."))
                k3.metric("🛫 Total Embarcados", f"{tot_emb:,.0f}".replace(",", "."))
                k4.metric("✅ Pasajeros Ajustados", f"{tot_aju:,.0f}".replace(",", "."))

                # --- BLOQUE KPI 2: ÚLTIMO MES CARGADO EN LA BASE ---
                ultima_fila = df_pax.iloc[-1]
                ult_mes_nombre = str(ultima_fila[col_mes])
                ult_oficio_txt = f" | {ultima_fila[col_oficio]}" if col_oficio and str(ultima_fila[col_oficio]).strip() else ""
                ult_dom = float(ultima_fila[col_tot_dom]) if col_tot_dom else 0.0
                ult_int = float(ultima_fila[col_tot_int]) if col_tot_int else 0.0
                ult_emb = float(ultima_fila[col_tot_emb]) if col_tot_emb else 0.0
                ult_aju = float(ultima_fila[col_tot_aju]) if col_tot_aju else 0.0

                st.subheader(f"🗓️ Último Mes Cargado: {ult_mes_nombre}{ult_oficio_txt}")
                u1, u2, u3, u4 = st.columns(4)
                u1.metric("🇨🇱 Doméstico (Último Mes)", f"{ult_dom:,.0f}".replace(",", "."))
                u2.metric("🌎 Internacional (Último Mes)", f"{ult_int:,.0f}".replace(",", "."))
                u3.metric("🛫 Embarcados (Último Mes)", f"{ult_emb:,.0f}".replace(",", "."))
                u4.metric("✅ Ajustados (Último Mes)", f"{ult_aju:,.0f}".replace(",", "."))

                st.divider()

                g1, g2 = st.columns([2, 1])

                with g1:
                    st.subheader("📈 Evolución Mensual: Doméstico vs. Internacional")
                    if col_tot_dom and col_tot_int:
                        df_graf = df_pax_filt[[col_mes, col_tot_dom, col_tot_int]].melt(
                            id_vars=col_mes,
                            var_name="Tipo Tráfico",
                            value_name="Pasajeros"
                        )
                        fig_pax = px.bar(
                            df_graf,
                            x=col_mes,
                            y="Pasajeros",
                            color="Tipo Tráfico",
                            barmode="group",
                            text_auto=".3s",
                            color_discrete_map={
                                col_tot_dom: "#3B82F6",
                                col_tot_int: "#10B981"
                            }
                        )
                        fig_pax.update_xaxes(type="category")
                        fig_pax.update_layout(xaxis_title="", yaxis_title="N° Pasajeros")
                        st.plotly_chart(fig_pax, use_container_width=True)

                with g2:
                    st.subheader("🥧 Participación de Tráfico")
                    if (tot_dom + tot_int) > 0:
                        df_pie = pd.DataFrame({
                            "Segmento": ["Doméstico", "Internacional"],
                            "Pasajeros": [tot_dom, tot_int]
                        })
                        fig_pie = px.pie(
                            df_pie,
                            values="Pasajeros",
                            names="Segmento",
                            hole=0.45,
                            color="Segmento",
                            color_discrete_map={"Doméstico": "#3B82F6", "Internacional": "#10B981"}
                        )
                        fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                        st.plotly_chart(fig_pie, use_container_width=True)

                st.divider()

                st.subheader("📋 Detalle Mensual Comunicado por DGAC")
                df_pax_mostrar = df_pax_filt.drop(columns=["Año_Filtro"], errors="ignore")
                df_pax_visual = df_pax_mostrar.copy()
                for c in df_pax_visual.select_dtypes(include="number").columns:
                    df_pax_visual[c] = df_pax_visual[c].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                st.dataframe(df_pax_visual, use_container_width=True, hide_index=True)

                c_dl1, c_dl2 = st.columns(2)
                etiqueta_pax = etiqueta_periodo_pax.replace(" ", "_")

                with c_dl1:
                    buffer_excel_pax = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel_pax, engine="openpyxl") as writer:
                        df_pax_mostrar.to_excel(writer, index=False, sheet_name="Pasajeros_DGAC")
                    buffer_excel_pax.seek(0)

                    st.download_button(
                        label=f"📥 Descargar Excel ({len(df_pax_mostrar)} meses)",
                        data=buffer_excel_pax,
                        file_name=f"Resumen_AMB_Pasajeros_{etiqueta_pax}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_excel_pax"
                    )

                with c_dl2:
                    buffer_word_pax = generar_informe_word_resumen_pax(
                        etiqueta_periodo_pax, df_pax_mostrar, tot_dom, tot_int, tot_emb, tot_aju
                    )
                    st.download_button(
                        label="📄 Descargar Informe Ejecutivo (Word .docx)",
                        data=buffer_word_pax,
                        file_name=f"Informe_Resumen_AMB_Pasajeros_{etiqueta_pax}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="btn_word_pax"
                    )

        except Exception as e:
            st.error(f"Error en el módulo Resumen AMB (Pasajeros): {e}")

    # --- SUB-PESTAÑA 2: EVOLUTIVO MANO DE OBRA ---
    with subtab_mo:
        try:
            df_mo = cargar_datos_mano_obra("MANO DE OBRA")

            if df_mo.empty:
                st.warning("La hoja 'MANO DE OBRA' se conectó correctamente, pero actualmente no tiene registros.")
            else:
                f_mo1, f_mo2, f_mo3 = st.columns(3)
                with f_mo1:
                    entidades_mo = sorted(df_mo["ENTIDAD"].dropna().unique().tolist())
                    entidad_sel = st.selectbox(
                        "🏢 Filtrar por Entidad:",
                        ["TODAS (Consolidado)"] + entidades_mo,
                        key="entidad_resumen_mo"
                    )
                with f_mo2:
                    anios_mo = sorted(df_mo["AÑO"].dropna().unique().tolist(), reverse=True)
                    anio_mo_sel = st.selectbox(
                        "📅 Filtrar por Año:",
                        ["Todos los Años"] + anios_mo,
                        key="anio_resumen_mo"
                    )

                df_mo_base = df_mo.copy()
                if entidad_sel != "TODAS (Consolidado)":
                    df_mo_base = df_mo_base[df_mo_base["ENTIDAD"] == entidad_sel]
                if anio_mo_sel != "Todos los Años":
                    df_mo_base = df_mo_base[df_mo_base["AÑO"] == anio_mo_sel]

                with f_mo3:
                    meses_mo_df = (
                        df_mo_base[["MES", "Mes_Num"]]
                        .drop_duplicates()
                        .query("MES != ''")
                        .sort_values("Mes_Num")
                    )
                    meses_mo = meses_mo_df["MES"].tolist()
                    mes_mo_sel = st.selectbox(
                        "📆 Filtrar por Mes:",
                        ["Todos los Meses"] + meses_mo,
                        key="mes_resumen_mo"
                    )

                df_mo_filt = df_mo_base.copy()
                if mes_mo_sel != "Todos los Meses":
                    df_mo_filt = df_mo_filt[df_mo_filt["MES"] == mes_mo_sel]

                if df_mo_filt.empty:
                    st.info("No hay registros de mano de obra para los filtros seleccionados.")
                else:
                    df_mo_agg = (
                        df_mo_filt.groupby(["Orden_Periodo", "Periodo"], as_index=False)[
                            ["M.O CALIFICADA", "M.O SEMICALIFICADA", "M.O NO CALIFICADA", "TOTAL M.O"]
                        ]
                        .sum()
                        .sort_values("Orden_Periodo")
                    )

                    if mes_mo_sel != "Todos los Meses":
                        tot_mo_gen = df_mo_filt["TOTAL M.O"].sum()
                        tot_mo_cal = df_mo_filt["M.O CALIFICADA"].sum()
                        tot_mo_sem = df_mo_filt["M.O SEMICALIFICADA"].sum()
                        tot_mo_noc = df_mo_filt["M.O NO CALIFICADA"].sum()
                        etiqueta_kpi_mo = f"Total Mensual Consultado ({mes_mo_sel} - {anio_mo_sel} | {entidad_sel})"
                    else:
                        ultimo = df_mo_agg.iloc[-1]
                        tot_mo_gen = ultimo["TOTAL M.O"]
                        tot_mo_cal = ultimo["M.O CALIFICADA"]
                        tot_mo_sem = ultimo["M.O SEMICALIFICADA"]
                        tot_mo_noc = ultimo["M.O NO CALIFICADA"]
                        etiqueta_kpi_mo = f"Total Último Período Registrado ({ultimo['Periodo']} | {entidad_sel})"

                    st.subheader(f"📊 {etiqueta_kpi_mo}")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("👷 Total Mensual M.O", f"{tot_mo_gen:,.0f}".replace(",", "."))
                    m2.metric("🎓 M.O Calificada", f"{tot_mo_cal:,.0f}".replace(",", "."))
                    m3.metric("🛠️ M.O Semicalificada", f"{tot_mo_sem:,.0f}".replace(",", "."))
                    m4.metric("🧰 M.O No Calificada", f"{tot_mo_noc:,.0f}".replace(",", "."))
                    st.divider()

                    st.subheader("📈 Evolutivo de Mano de Obra")
                    cols_graf_mo = ["M.O CALIFICADA", "M.O SEMICALIFICADA", "M.O NO CALIFICADA", "TOTAL M.O"]
                    df_mo_long = df_mo_agg.melt(
                        id_vars=["Orden_Periodo", "Periodo"],
                        value_vars=cols_graf_mo,
                        var_name="Categoría",
                        value_name="Dotación"
                    )
                    fig_mo = px.line(
                        df_mo_long,
                        x="Periodo",
                        y="Dotación",
                        color="Categoría",
                        markers=True,
                        color_discrete_map={
                            "TOTAL M.O": "#F8FAFC",
                            "M.O CALIFICADA": "#3B82F6",
                            "M.O SEMICALIFICADA": "#10B981",
                            "M.O NO CALIFICADA": "#F59E0B"
                        }
                    )
                    fig_mo.update_xaxes(type="category")
                    fig_mo.update_layout(xaxis_title="", yaxis_title="N° Trabajadores")
                    st.plotly_chart(fig_mo, use_container_width=True)

                    st.divider()

                    st.subheader("📋 Detalle de Mano de Obra")
                    cols_mostrar_mo = [
                        "ENTIDAD", "AÑO", "MES",
                        "M.O CALIFICADA", "M.O SEMICALIFICADA", "M.O NO CALIFICADA", "TOTAL M.O"
                    ]
                    df_mo_mostrar = df_mo_filt[cols_mostrar_mo].copy()
                    df_mo_visual = df_mo_mostrar.copy()
                    for c in ["M.O CALIFICADA", "M.O SEMICALIFICADA", "M.O NO CALIFICADA", "TOTAL M.O"]:
                        df_mo_visual[c] = df_mo_visual[c].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                    st.dataframe(df_mo_visual, use_container_width=True, hide_index=True)

                    buffer_excel_mo = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel_mo, engine="openpyxl") as writer:
                        df_mo_mostrar.to_excel(writer, index=False, sheet_name="Mano_de_Obra")
                    buffer_excel_mo.seek(0)

                    etiqueta_mo = f"{entidad_sel}_{anio_mo_sel}_{mes_mo_sel}".replace(" ", "_").replace("(", "").replace(")", "")
                    st.download_button(
                        label=f"📥 Descargar Excel Mano de Obra ({len(df_mo_mostrar)} registros)",
                        data=buffer_excel_mo,
                        file_name=f"Resumen_AMB_Mano_de_Obra_{etiqueta_mo}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_excel_mo"
                    )

        except Exception as e:
            st.error(f"Error en el módulo Resumen AMB (Mano de Obra): {e}")

# ==========================================
# MÓDULO 2: FISCALIZACIÓN PACO
# ==========================================
with tab_paco:
    try:
        df = cargar_datos_amb()

        if df.empty:
            st.warning("La base de datos de PACO se conectó correctamente, pero actualmente no tiene registros.")
        else:
            hoy_chile = pd.Timestamp.now(tz='America/Santiago')
            hoy_str = hoy_chile.strftime('%Y-%m-%d')

            if 'Fecha_Filtro' in df.columns:
                df['Fecha_Str'] = df['Fecha_Filtro'].dt.strftime('%Y-%m-%d')
                fiscalizaciones_hoy = len(df[df['Fecha_Str'] == hoy_str])
            else:
                fiscalizaciones_hoy = 0

            total_fiscalizaciones = len(df)
            incumplimientos_total = (
                len(df[~df["Cumple_Limpio"].isin(["SÍ", "SI", "CUMPLE", "OPERATIVO", "BUENO", ""])])
                if "Cumple_Limpio" in df.columns else 0
            )

            col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
            col_kpi1.metric("📋 Total Histórico", total_fiscalizaciones)
            col_kpi2.metric("📅 Fiscalizaciones Hoy", fiscalizaciones_hoy)
            col_kpi3.metric("🚨 Total Incumplimientos", incumplimientos_total)
            st.divider()

            col_graficos, col_gestion = st.columns([2, 1.5])

            with col_graficos:
                st.subheader("🏆 Ranking de Fiscalizadores (PACO)")
                if 'Fiscalizador' in df.columns:
                    ranking = df['Fiscalizador'].value_counts().reset_index()
                    ranking.columns = ['Fiscalizador', 'Cantidad de Registros']
                    fig = px.bar(
                        ranking,
                        x='Fiscalizador',
                        y='Cantidad de Registros',
                        text_auto=True,
                        color='Cantidad de Registros',
                        color_continuous_scale="Blues"
                    )
                    fig.update_layout(xaxis_title="", yaxis_title="N° Registros")
                    st.plotly_chart(fig, use_container_width=True)

                col_top_fisc, col_top_incumple = st.columns(2)
                with col_top_fisc:
                    st.subheader("📍 Top 3: Mayor Fiscalización")
                    if 'Sector fiscalizado' in df.columns:
                        top3_fisc = df['Sector fiscalizado'].value_counts().head(3).reset_index()
                        top3_fisc.columns = ['Sector', 'N° Registros']
                        st.dataframe(top3_fisc, use_container_width=True, hide_index=True)

                with col_top_incumple:
                    st.subheader("⚠️ Top 3: Más Incumplimientos")
                    if 'Sector fiscalizado' in df.columns and 'Cumple_Limpio' in df.columns:
                        df_no_cumple_graf = df[~df["Cumple_Limpio"].isin(["SÍ", "SI", "CUMPLE", "OPERATIVO", "BUENO", ""])]
                        if len(df_no_cumple_graf) > 0:
                            top3_inc = df_no_cumple_graf['Sector fiscalizado'].value_counts().head(3).reset_index()
                            top3_inc.columns = ['Sector', 'N° Incumplimientos']
                            st.dataframe(top3_inc, use_container_width=True, hide_index=True)
                        else:
                            st.success("✅ Sin incumplimientos.")

                st.divider()
                st.subheader("📥 Exportación Rápida de Base de Datos PACO")
                df_export = df.drop(columns=['Fecha_Filtro', 'Fecha_Str', 'Cumple_Limpio'], errors='ignore')
                cols_paco = ['ID_Pantalla'] + [c for c in df_export.columns if c != 'ID_Pantalla']
                df_export = df_export[cols_paco]
                csv_completo = df_export.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="Descargar Registro Histórico Completo (CSV)",
                    data=csv_completo,
                    file_name=f"Base_Datos_PACO_{hoy_str}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

            with col_gestion:
                tab_ficha, tab_incumple = st.tabs(["🔍 Ficha de Fiscalización", "🚨 Solo Incumplimientos"])
                with tab_ficha:
                    st.write("Auditoría de registro individual:")
                    registro_sel = st.selectbox("Seleccione el registro:", df['ID_Pantalla'].tolist())
                    if registro_sel:
                        fila = df[df['ID_Pantalla'] == registro_sel].iloc[0]
                        st.markdown(f"**Fecha:** {fila.get('Fecha', 'N/A')} {fila.get('Hora', '')}")
                        st.markdown(f"**Fiscalizador:** {fila.get('Fiscalizador', 'N/A')}")
                        st.markdown(f"**Sector:** {fila.get('Sector fiscalizado', 'N/A')} / {fila.get('Espigón', 'N/A')}")
                        st.markdown(f"**OT / Activo:** {fila.get('Numero de OT', 'N/A')} | {fila.get('Código del activo', 'N/A')}")
                        st.markdown(f"**Grupo de Trabajo:** {fila.get('Grupo de Trabajadores', 'N/A')}")
                        cumple_txt = fila.get('Cumple_Limpio', 'N/A')
                        color_cumple = "green" if cumple_txt in ["SÍ", "SI", "CUMPLE", "OPERATIVO", "BUENO"] else "red"
                        st.markdown(f"**¿Cumple?:** :{color_cumple}[**{cumple_txt}**]")
                        st.markdown(f"**Detalle Observación:** {fila.get('Detalle Observación', 'Sin observaciones')}")

                        st.divider()
                        st.markdown("**📄 Generar Anexo para Oficio (Formato Tipo AMB):**")
                        buffer_word = generar_informe_word(fila)
                        st.download_button(
                            label="Descargar Informe de Fiscalización (Word .docx)",
                            data=buffer_word,
                            file_name=f"Informe_FISC_{fila.get('Numero de OT', 'N_A')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

                with tab_incumple:
                    st.write("Vista rápida de hallazgos que no cumplen el SLA:")
                    if "Cumple_Limpio" in df.columns:
                        df_no_cumple = df[~df["Cumple_Limpio"].isin(["SÍ", "SI", "CUMPLE", "OPERATIVO", "BUENO", ""])].copy()
                        if len(df_no_cumple) > 0:
                            if 'Sector fiscalizado' in df_no_cumple.columns and 'Espigón' in df_no_cumple.columns:
                                df_no_cumple['Localización'] = (
                                    df_no_cumple['Sector fiscalizado'].astype(str) + " - " + df_no_cumple['Espigón'].astype(str)
                                )
                            else:
                                df_no_cumple['Localización'] = df_no_cumple.get('Sector fiscalizado', 'N/A')

                            def extraer_enlace(fotos_str):
                                return str(fotos_str).split(',')[0].strip() if pd.notna(fotos_str) and str(fotos_str).strip() else None

                            df_no_cumple['Evidencia'] = df_no_cumple.get('Registro Fotográfico (hasta 10 fotografias)', '').apply(extraer_enlace)
                            cols_validas = [c for c in ['Fecha', 'Localización', 'Detalle Observación', 'Evidencia'] if c in df_no_cumple.columns]
                            st.dataframe(
                                df_no_cumple[cols_validas],
                                use_container_width=True,
                                hide_index=True,
                                column_config={"Evidencia": st.column_config.LinkColumn("Evidencia Fotográfica", display_text="Ver Fotografía 📷")}
                            )
                        else:
                            st.success("✅ No hay incumplimientos registrados.")

            st.divider()

            st.subheader("📋 Registro Histórico PACO por Estado")

            if "Cumple_Limpio" in df.columns:
                df_hist_paco = df.copy()
                df_hist_paco["Estado_PACO"] = df_hist_paco["Cumple_Limpio"].replace(["", "NAN", "NONE"], "SIN ESTADO")
                estados_paco = sorted(df_hist_paco["Estado_PACO"].dropna().unique().tolist())

                estado_paco_sel = st.radio(
                    "Filtrar registro histórico y descarga por Estado:",
                    options=["TODOS"] + estados_paco,
                    horizontal=True,
                    key="filtro_estado_paco"
                )

                if estado_paco_sel != "TODOS":
                    df_paco_tabla = df_hist_paco[df_hist_paco["Estado_PACO"] == estado_paco_sel].copy()
                    nombre_excel_paco = f"Fiscalizacion_PACO_{estado_paco_sel}_{hoy_str}.xlsx".replace(" ", "_")
                    texto_btn_paco = f"📥 Descargar Excel: Estado {estado_paco_sel} ({len(df_paco_tabla)} registros)"
                else:
                    df_paco_tabla = df_hist_paco.copy()
                    nombre_excel_paco = f"Fiscalizacion_PACO_TODOS_{hoy_str}.xlsx"
                    texto_btn_paco = f"📥 Descargar Excel: Registro Histórico Completo ({len(df_paco_tabla)} registros)"
            else:
                df_paco_tabla = df.copy()
                nombre_excel_paco = f"Fiscalizacion_PACO_TODOS_{hoy_str}.xlsx"
                texto_btn_paco = f"📥 Descargar Excel: Registro Histórico Completo ({len(df_paco_tabla)} registros)"

            cols_ocultas_paco = ["Fecha_Filtro", "Fecha_Str", "Cumple_Limpio", "Estado_PACO"]
            df_paco_mostrar = df_paco_tabla.drop(
                columns=[c for c in cols_ocultas_paco if c in df_paco_tabla.columns],
                errors="ignore"
            )
            if "ID_Pantalla" in df_paco_mostrar.columns:
                cols_orden = ["ID_Pantalla"] + [c for c in df_paco_mostrar.columns if c != "ID_Pantalla"]
                df_paco_mostrar = df_paco_mostrar[cols_orden]

            if len(df_paco_mostrar) > 0:
                st.dataframe(df_paco_mostrar, use_container_width=True, hide_index=True)

                buffer_excel_paco = io.BytesIO()
                with pd.ExcelWriter(buffer_excel_paco, engine="openpyxl") as writer:
                    df_paco_mostrar.to_excel(writer, index=False, sheet_name="Historico_PACO")
                buffer_excel_paco.seek(0)

                st.download_button(
                    label=texto_btn_paco,
                    data=buffer_excel_paco,
                    file_name=nombre_excel_paco,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_excel_paco"
                )
            else:
                st.info("No hay registros de fiscalización para el estado seleccionado.")

    except Exception as e:
        st.error(f"Error al procesar datos PACO: {e}")

# ==========================================
# MÓDULO 3: FISCALIZACIÓN DE EXTINTORES
# ==========================================
with tab_extintores:
    st.header("🧯 Módulo de Fiscalización de Extintores")
    try:
        df_ext = cargar_datos_extintores()

        if df_ext.empty:
            st.warning("La base de datos de Extintores se conectó correctamente, pero actualmente no tiene registros.")
        else:
            hoy_chile = pd.Timestamp.now(tz='America/Santiago')
            hoy_str = hoy_chile.strftime('%Y-%m-%d')

            if 'Fecha_Filtro' in df_ext.columns:
                df_ext['Fecha_Str'] = df_ext['Fecha_Filtro'].dt.strftime('%Y-%m-%d')
                fiscalizaciones_hoy_ext = len(df_ext[df_ext['Fecha_Str'] == hoy_str])
            else:
                fiscalizaciones_hoy_ext = 0

            total_ext = len(df_ext)
            if "Cumple_Limpio" in df_ext.columns:
                df_no_cumple_ext = df_ext[~df_ext["Cumple_Limpio"].isin(["SÍ", "SI", "CUMPLE", "OPERATIVO", "BUENO", ""])].copy()
                inc_total_ext = len(df_no_cumple_ext)
            else:
                df_no_cumple_ext = pd.DataFrame()
                inc_total_ext = 0

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📋 Total Registros (Extintores)", total_ext)
            c2.metric("📅 Fiscalizaciones Hoy", fiscalizaciones_hoy_ext)
            c3.metric("🚨 Total Incumplimientos", inc_total_ext)
            st.divider()

            col_top_sec, col_top_inc = st.columns(2)
            with col_top_sec:
                st.subheader("📍 Top 5: Sectores Fiscalizados")
                if 'Ingresa zona de fiscalización' in df_ext.columns:
                    top5_sec = df_ext['Ingresa zona de fiscalización'].value_counts().head(5).reset_index()
                    top5_sec.columns = ['Sector', 'N° Registros']
                    st.dataframe(top5_sec, use_container_width=True, hide_index=True)

            with col_top_inc:
                st.subheader("⚠️ Top 5: Sectores con Incumplimiento")
                if 'Ingresa zona de fiscalización' in df_ext.columns and not df_no_cumple_ext.empty:
                    top5_inc = df_no_cumple_ext['Ingresa zona de fiscalización'].value_counts().head(5).reset_index()
                    top5_inc.columns = ['Sector', 'N° Incumplimientos']
                    st.dataframe(top5_inc, use_container_width=True, hide_index=True)
                elif inc_total_ext == 0:
                    st.success("✅ Sin incumplimientos en extintores.")

            st.divider()
            st.subheader("🚨 Tabla de Gestión: Incumplimientos de Extintores")

            if not df_no_cumple_ext.empty:
                if 'Ingresa zona de fiscalización' in df_no_cumple_ext.columns and 'Ingresa espigón' in df_no_cumple_ext.columns:
                    df_no_cumple_ext['Localización'] = (
                        df_no_cumple_ext['Ingresa zona de fiscalización'].astype(str) + " - " + df_no_cumple_ext['Ingresa espigón'].astype(str)
                    )
                elif 'Ingresa zona de fiscalización' in df_no_cumple_ext.columns:
                    df_no_cumple_ext['Localización'] = df_no_cumple_ext['Ingresa zona de fiscalización']
                else:
                    df_no_cumple_ext['Localización'] = "N/A"

                col_foto_ext = next((c for c in df_no_cumple_ext.columns if "Respaldo fotográfico" in c), None)

                def extraer_enlace_ext(fotos_str):
                    return str(fotos_str).split(',')[0].strip() if pd.notna(fotos_str) and str(fotos_str).strip() else None

                if col_foto_ext:
                    df_no_cumple_ext['Evidencia'] = df_no_cumple_ext[col_foto_ext].apply(extraer_enlace_ext)
                else:
                    df_no_cumple_ext['Evidencia'] = None

                df_no_cumple_ext['Observación'] = df_no_cumple_ext.get('Observación', 'Sin observaciones')
                df_no_cumple_ext['Fecha'] = df_no_cumple_ext.get('Marca temporal', 'N/A')

                cols_validas_ext = [c for c in ['Fecha', 'Localización', 'Observación', 'Evidencia'] if c in df_no_cumple_ext.columns]
                st.dataframe(
                    df_no_cumple_ext[cols_validas_ext],
                    use_container_width=True,
                    hide_index=True,
                    column_config={"Evidencia": st.column_config.LinkColumn("Evidencia Fotográfica", display_text="Ver Fotografía 📷")}
                )

                st.divider()
                st.subheader("📄 Generar Anexo para Oficio (Formato Tipo AMB)")
                st.write("Seleccione el ID de un registro con incumplimiento para descargar su informe:")

                opciones_ext = df_no_cumple_ext['ID_Pantalla'].tolist()
                registro_ext_sel = st.selectbox("Seleccione el ID del Extintor:", opciones_ext)

                if registro_ext_sel:
                    fila_ext = df_no_cumple_ext[df_no_cumple_ext['ID_Pantalla'] == registro_ext_sel].iloc[0]
                    st.markdown(f"**Ubicación:** {fila_ext.get('Ingresa zona de fiscalización', 'N/A')} - {fila_ext.get('Ingresa espigón', 'N/A')}")
                    st.markdown(f"**Falla:** {fila_ext.get('Estado del sistema fiscalizado', 'N/A')}")

                    buffer_word_ext = generar_informe_word_extintor(fila_ext, col_foto_ext)
                    st.download_button(
                        label="Descargar Informe de Extintor (Word .docx)",
                        data=buffer_word_ext,
                        file_name=f"Informe_EXTINTOR_{registro_ext_sel}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            else:
                st.info("No hay hallazgos reprobados que gestionar.")

            st.write("")
            df_ext_export = df_ext.drop(columns=['Fecha_Filtro', 'Fecha_Str', 'Cumple_Limpio'], errors='ignore')
            cols_ext = ['ID_Pantalla'] + [c for c in df_ext_export.columns if c != 'ID_Pantalla']
            df_ext_export = df_ext_export[cols_ext]
            csv_ext = df_ext_export.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Descargar Base de Extintores Completa (CSV)",
                data=csv_ext,
                file_name=f"Base_Extintores_AMB_{hoy_str}.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Error en módulo Extintores: {e}")

# ==========================================
# MÓDULO 4: CONTROL MULTAS AMB
# ==========================================
with tab_multas:
    st.header("⚖️ Panel de Control de Multas (AMB)")
    try:
        df_multas = cargar_datos_multas()

        if df_multas.empty:
            st.warning("No hay datos en la base de multas o hubo un error al cargar.")
        else:
            if "Mes_Año" in df_multas.columns:
                meses_disponibles = sorted(df_multas["Mes_Año"].dropna().unique().tolist(), reverse=True)
                mes_sel = st.selectbox(
                    "📅 Consultar por Período Mensual (Año-Mes):",
                    ["Acumulado Histórico (Todos)"] + meses_disponibles
                )
                if mes_sel != "Acumulado Histórico (Todos)":
                    df_filtro_multas = df_multas[df_multas["Mes_Año"] == mes_sel].copy()
                else:
                    df_filtro_multas = df_multas.copy()
            else:
                st.warning("No se encontró una columna de Fecha para crear el filtro mensual.")
                mes_sel = "Histórico"
                df_filtro_multas = df_multas.copy()

            st.subheader(f"Resumen del Estado Contractual: {mes_sel}")

            estados_serie = df_filtro_multas["Estado_Limpio"]
            propuestas = len(df_filtro_multas[estados_serie.str.contains("PROPUEST", na=False)])
            cursadas = len(df_filtro_multas[estados_serie.str.contains("CURSAD", na=False)])
            pagadas = len(df_filtro_multas[estados_serie.str.contains("PAGAD", na=False)])
            suspendidas = len(df_filtro_multas[estados_serie.str.contains("SUSPEND|RECLAM|DESIST", na=False)])

            es_pagada = estados_serie.str.contains("PAGAD", na=False)
            es_reclamada = df_filtro_multas["Tiene_Reclamacion"] | estados_serie.str.contains("RECLAM", na=False)
            es_en_proceso = estados_serie.str.contains("PROPUEST|CURSAD", na=False)

            total_pagado = df_filtro_multas[es_pagada]["Monto_Numerico"].sum()
            total_pagado_reclamo = df_filtro_multas[es_pagada & es_reclamada]["Monto_Numerico"].sum()
            cant_pagado_reclamo = len(df_filtro_multas[es_pagada & es_reclamada])
            total_en_proceso = df_filtro_multas[es_en_proceso]["Monto_Numerico"].sum()

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("📝 Propuestas", propuestas)
            col2.metric("📩 Cursadas", cursadas)
            col3.metric("⛔ Suspendidas / Otros", suspendidas)
            col4.metric("✅ Pagadas", pagadas)

            st.divider()

            col_graf, col_finanzas = st.columns([2, 1.5])

            with col_graf:
                st.write("**Distribución de Multas por Estado (%)**")
                if len(df_filtro_multas) > 0:
                    estado_counts = df_filtro_multas["Estado_Limpio"].value_counts().reset_index()
                    estado_counts.columns = ["Estado", "Cantidad"]
                    color_map = {
                        "PAGADA": "#00CC96",
                        "PROPUESTA": "#636EFA",
                        "CURSADA": "#FFA15A",
                        "SUSPENDIDA": "#EF553B",
                        "DESISTIDA": "#AB63FA"
                    }

                    fig_multas = px.pie(
                        estado_counts,
                        values="Cantidad",
                        names="Estado",
                        hole=0.4,
                        color="Estado",
                        color_discrete_map=color_map
                    )
                    fig_multas.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(fig_multas, use_container_width=True)
                else:
                    st.info("No hay datos suficientes para graficar.")

            with col_finanzas:
                st.write("**Resumen Financiero (En UTM)**")
                st.markdown(f"""
                <div style="padding: 15px; border-radius: 10px; background-color: #1E293B; margin-bottom: 12px; text-align: center; border-left: 5px solid #00CC96;">
                    <h4 style="margin: 0; color: #00CC96;">Total UTM Pagadas</h4>
                    <h2 style="margin: 5px 0 0 0; color: white;">{total_pagado:,.2f} UTM</h2>
                </div>

                <div style="padding: 15px; border-radius: 10px; background-color: #2D1E2F; margin-bottom: 12px; text-align: center; border-left: 5px solid #AB63FA;">
                    <h4 style="margin: 0; color: #D6A2E8;">Pagadas con Reclamación ({cant_pagado_reclamo})</h4>
                    <h2 style="margin: 5px 0 0 0; color: white;">{total_pagado_reclamo:,.2f} UTM</h2>
                </div>

                <div style="padding: 15px; border-radius: 10px; background-color: #334155; text-align: center; border-left: 5px solid #FFA15A;">
                    <h4 style="margin: 0; color: #FFA15A;">UTM en Trámite (Propuestas + Cursadas)</h4>
                    <h2 style="margin: 5px 0 0 0; color: white;">{total_en_proceso:,.2f} UTM</h2>
                </div>
                """, unsafe_allow_html=True)

            with st.expander("🔍 Auditar suma de UTM por Estado"):
                resumen_auditoria = (
                    df_filtro_multas.groupby("Estado_Limpio", as_index=False)
                    .agg(Cantidad=("Monto_Numerico", "count"), Total_UTM=("Monto_Numerico", "sum"))
                    .sort_values("Total_UTM", ascending=False)
                )
                st.dataframe(resumen_auditoria, use_container_width=True, hide_index=True)

            st.divider()

            st.subheader("📋 Detalle de Multas por Estado")

            estados_unicos = sorted(df_filtro_multas["Estado_Limpio"].dropna().unique().tolist())
            estado_seleccionado = st.radio(
                "Filtrar tabla y descargas por Estado:",
                options=["TODOS"] + estados_unicos,
                horizontal=True,
                key="filtro_estado_multas"
            )

            if estado_seleccionado != "TODOS":
                df_tabla = df_filtro_multas[df_filtro_multas["Estado_Limpio"] == estado_seleccionado].copy()
                nombre_archivo_excel = f"Multas_AMB_{estado_seleccionado}_{mes_sel}.xlsx".replace(" ", "_")
                nombre_archivo_word = f"Informe_Multas_AMB_{estado_seleccionado}_{mes_sel}.docx".replace(" ", "_")
                texto_boton_excel = f"📥 Descargar Excel: Estado {estado_seleccionado} ({len(df_tabla)} registros)"
            else:
                df_tabla = df_filtro_multas.copy()
                nombre_archivo_excel = f"Multas_AMB_TODOS_{mes_sel}.xlsx".replace(" ", "_")
                nombre_archivo_word = f"Informe_Multas_AMB_TODOS_{mes_sel}.docx".replace(" ", "_")
                texto_boton_excel = f"📥 Descargar Excel: Listado Completo ({len(df_tabla)} registros)"

            columnas_ocultas = ["Fecha_Real", "Mes_Año", "Estado_Limpio", "Monto_Numerico", "Tiene_Reclamacion"]
            df_mostrar = df_tabla.drop(columns=[col for col in columnas_ocultas if col in df_tabla.columns], errors='ignore')

            if len(df_mostrar) > 0:
                st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    buffer_excel = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                        df_mostrar.to_excel(writer, index=False, sheet_name='Detalle_Multas')
                    buffer_excel.seek(0)

                    st.download_button(
                        label=texto_boton_excel,
                        data=buffer_excel,
                        file_name=nombre_archivo_excel,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_excel_multas"
                    )

                with col_dl2:
                    buffer_word_multas = generar_informe_word_multas(
                        mes_sel, estado_seleccionado, df_tabla,
                        total_pagado, total_pagado_reclamo, cant_pagado_reclamo, total_en_proceso
                    )
                    st.download_button(
                        label="📄 Descargar Informe Ejecutivo de Multas (Word .docx)",
                        data=buffer_word_multas,
                        file_name=nombre_archivo_word,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="btn_word_multas"
                    )
            else:
                st.info("No hay multas registradas para el filtro consultado.")

    except Exception as e:
        st.error(f"Error en el módulo de Multas: {e}")

# ==========================================
# MÓDULO 5: CONTROL CORRESPONDENCIA
# ==========================================
with tab_correspondencia:
    st.header("✉️ Panel de Control de Correspondencia")
    try:
        df_corr = cargar_datos_correspondencia()

        if df_corr.empty:
            st.warning("La base de Correspondencia se conectó correctamente, pero actualmente no tiene registros.")
        else:
            col_f_mes, col_f_sem = st.columns(2)
            with col_f_mes:
                if "Mes_Año" in df_corr.columns:
                    meses_corr = sorted(df_corr["Mes_Año"].dropna().unique().tolist(), reverse=True)
                    mes_corr_sel = st.selectbox(
                        "📅 Consultar por Período Mensual (Año-Mes):",
                        ["Acumulado Histórico (Todos)"] + meses_corr,
                        key="mes_correspondencia"
                    )
                    if mes_corr_sel != "Acumulado Histórico (Todos)":
                        df_corr_filt = df_corr[df_corr["Mes_Año"] == mes_corr_sel].copy()
                    else:
                        df_corr_filt = df_corr.copy()
                else:
                    mes_corr_sel = "Histórico"
                    df_corr_filt = df_corr.copy()

            with col_f_sem:
                if "Semana" in df_corr_filt.columns:
                    semanas_corr = sorted(df_corr_filt["Semana"].dropna().unique().tolist(), reverse=True)
                    sem_corr_sel = st.selectbox(
                        "📆 Filtrar por Semana:",
                        ["Todas las semanas"] + semanas_corr,
                        key="sem_correspondencia"
                    )
                    if sem_corr_sel != "Todas las semanas":
                        df_corr_filt = df_corr_filt[df_corr_filt["Semana"] == sem_corr_sel].copy()
                else:
                    sem_corr_sel = "Todas las semanas"

            periodo_corr_txt = f"{mes_corr_sel} | {sem_corr_sel}" if sem_corr_sel != "Todas las semanas" else mes_corr_sel
            etiqueta_archivo = (
                sem_corr_sel.split(" ")[0] if sem_corr_sel != "Todas las semanas" else mes_corr_sel
            ).replace(" ", "_")

            total_corr = len(df_corr_filt)
            con_ordinario = len(df_corr_filt[df_corr_filt["Estado_Corr"] == "CON ORDINARIO"])
            pendientes_ord = len(df_corr_filt[df_corr_filt["Estado_Corr"] == "PENDIENTE"])
            total_areas = df_corr_filt["AREA"].replace("", pd.NA).dropna().nunique() if "AREA" in df_corr_filt.columns else 0

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("📨 Total Correspondencia", total_corr)
            k2.metric("✅ Con N° Ordinario", con_ordinario)
            k3.metric("⏳ Pendientes de Ordinario", pendientes_ord)
            k4.metric("🏢 Áreas Involucradas", total_areas)
            st.divider()

            g1, g2 = st.columns(2)
            with g1:
                st.subheader("🏢 Documentos por Área")
                if "AREA" in df_corr_filt.columns and len(df_corr_filt) > 0:
                    conteo_area = df_corr_filt["AREA"].replace("", "SIN ÁREA").value_counts().reset_index()
                    conteo_area.columns = ["Área", "Cantidad"]
                    fig_area = px.bar(
                        conteo_area, x="Área", y="Cantidad", text_auto=True,
                        color="Cantidad", color_continuous_scale="Blues"
                    )
                    fig_area.update_layout(xaxis_title="", yaxis_title="N° Documentos")
                    st.plotly_chart(fig_area, use_container_width=True)

            with g2:
                st.subheader("👤 Carga por Profesional")
                if "PROFESIONAL" in df_corr_filt.columns and len(df_corr_filt) > 0:
                    conteo_prof = df_corr_filt["PROFESIONAL"].replace("", "SIN ASIGNAR").value_counts().reset_index()
                    conteo_prof.columns = ["Profesional", "Cantidad"]
                    fig_prof = px.bar(
                        conteo_prof, x="Cantidad", y="Profesional", orientation="h", text_auto=True,
                        color="Cantidad", color_continuous_scale="Teal"
                    )
                    fig_prof.update_layout(xaxis_title="N° Documentos", yaxis_title="")
                    st.plotly_chart(fig_prof, use_container_width=True)

            st.divider()

            st.subheader("📋 Registro Detallado de Correspondencia")

            f1, f2, f3 = st.columns(3)
            with f1:
                areas_disp = sorted(df_corr_filt["AREA"].replace("", "SIN ÁREA").unique().tolist()) if "AREA" in df_corr_filt.columns else []
                area_sel = st.selectbox("Filtrar por Área:", ["TODAS"] + areas_disp, key="filtro_area_corr")
            with f2:
                profs_disp = sorted(df_corr_filt["PROFESIONAL"].replace("", "SIN ASIGNAR").unique().tolist()) if "PROFESIONAL" in df_corr_filt.columns else []
                prof_sel = st.selectbox("Filtrar por Profesional:", ["TODOS"] + profs_disp, key="filtro_prof_corr")
            with f3:
                busqueda = st.text_input("Buscar por BO, Materia o N° Ordinario:", "", key="busqueda_corr")

            estado_corr_sel = st.radio(
                "Filtrar por Estado de Tramitación:",
                options=["TODOS", "CON ORDINARIO", "PENDIENTE"],
                horizontal=True,
                key="filtro_estado_corr"
            )

            df_corr_tabla = df_corr_filt.copy()
            if area_sel != "TODAS" and "AREA" in df_corr_tabla.columns:
                df_corr_tabla = df_corr_tabla[df_corr_tabla["AREA"].replace("", "SIN ÁREA") == area_sel]
            if prof_sel != "TODOS" and "PROFESIONAL" in df_corr_tabla.columns:
                df_corr_tabla = df_corr_tabla[df_corr_tabla["PROFESIONAL"].replace("", "SIN ASIGNAR") == prof_sel]
            if estado_corr_sel != "TODOS":
                df_corr_tabla = df_corr_tabla[df_corr_tabla["Estado_Corr"] == estado_corr_sel]
            if busqueda.strip():
                texto = busqueda.strip().lower()
                mask = df_corr_tabla.astype(str).apply(lambda col: col.str.lower().str.contains(texto, na=False)).any(axis=1)
                df_corr_tabla = df_corr_tabla[mask]

            cols_ocultas_corr = ["Fecha_Real", "Mes_Año", "Semana", "Estado_Corr"]
            df_corr_mostrar = df_corr_tabla.drop(columns=[c for c in cols_ocultas_corr if c in df_corr_tabla.columns], errors="ignore")

            if len(df_corr_mostrar) > 0:
                st.dataframe(df_corr_mostrar, use_container_width=True, hide_index=True)

                c_dl1, c_dl2 = st.columns(2)
                with c_dl1:
                    buffer_excel_corr = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel_corr, engine="openpyxl") as writer:
                        df_corr_mostrar.to_excel(writer, index=False, sheet_name="Correspondencia")
                    buffer_excel_corr.seek(0)

                    st.download_button(
                        label=f"📥 Descargar Excel ({len(df_corr_mostrar)} registros)",
                        data=buffer_excel_corr,
                        file_name=f"Control_Correspondencia_{estado_corr_sel}_{etiqueta_archivo}.xlsx".replace(" ", "_"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_excel_corr"
                    )

                with c_dl2:
                    buffer_word_corr = generar_informe_word_correspondencia(
                        periodo_corr_txt, estado_corr_sel, area_sel, df_corr_tabla
                    )
                    st.download_button(
                        label="📄 Descargar Informe Ejecutivo (Word .docx)",
                        data=buffer_word_corr,
                        file_name=f"Informe_Correspondencia_{etiqueta_archivo}.docx".replace(" ", "_"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="btn_word_corr"
                    )
            else:
                st.info("No hay registros de correspondencia para los filtros seleccionados.")

    except Exception as e:
        st.error(f"Error en el módulo de Correspondencia: {e}")
