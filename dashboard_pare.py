# ============================================================
# DASHBOARD PARE Y SIGA - SST
# Versión: Google Sheets + IA Groq + SCP + Predicción + Semanal
# ============================================================

# ============ IMPORTS ============
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import numpy as np
import sys
import re
from google.oauth2.service_account import Credentials
import gspread

try:
    from utils import mostrar_ip_sidebar
    UTILS_DISPONIBLE = True
except ImportError:
    UTILS_DISPONIBLE = False

print(f"🐍 PYTHON: {sys.executable}")
print(f"📁 Working dir: {os.getcwd()}")

# ============ CONFIGURACIÓN IA (Groq) ============
# ============ CONFIGURACIÓN IA (Groq) ============
import os

try:
    from groq import Groq
    GROQ_DISPONIBLE = True
except ImportError as e:
    GROQ_DISPONIBLE = False
    print(f"ERROR IMPORTANDO GROQ: {e}")

MODELO_GROQ = "openai/gpt-oss-120b"

# --- LECTURA A PRUEBA DE BALAS ---
GROQ_API_KEY = ""

# Intento 1: Leer archivo directamente con Python puro
try:
    if os.path.exists(".streamlit/secrets.toml"):
        with open(".streamlit/secrets.toml", "r", encoding="utf-8") as f:
            for linea in f:
                if linea.strip().startswith("GROQ_API_KEY"):
                    # Extrae lo que está entre comillas
                    GROQ_API_KEY = linea.split("=")[1].strip().strip('"').strip("'")
                    break
    else:
        st.error("❌ ERROR CRÍTICO: No se encuentra el archivo .streamlit/secrets.toml")
except Exception as e:
    st.error(f"❌ Error leyendo archivo: {e}")

# Intento 2: Si el archivo falló, usar st.secrets
if not GROQ_API_KEY:
    try:
        GROQ_API_KEY = str(st.secrets["GROQ_API_KEY"]).strip()
    except Exception:
        pass

# Intento 3: Variable de entorno (por si acaso)
if not GROQ_API_KEY:
    GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")

print("=" * 60)
print(f"🔑 LONGITUD DE LA CLAVE DETECTADA: {len(GROQ_API_KEY)}")
print("=" * 60)

def ia_disponible():
    return bool(GROQ_DISPONIBLE and len(GROQ_API_KEY) > 10)

cliente_ia = None
if ia_disponible():
    try:
        cliente_ia = Groq(api_key=GROQ_API_KEY)
        print("✅ Cliente Groq creado correctamente.")
    except Exception as e:
        print(f"❌ Error creando cliente Groq: {e}")


def generar_analisis_gemini(resumen_dict):
    """Genera análisis SST usando Groq."""
    if cliente_ia is None:
        return "⚠️ Groq no está configurado."

    prompt = f"""
Eres un experto en Seguridad y Salud en el Trabajo (SST) con especialización 
en programas de observación conductual (BBS - Behavior Based Safety).

Analiza los siguientes datos del programa PARE y SIGA:

{resumen_dict}

Genera un análisis ejecutivo en español con:
1. **Diagnóstico general** (3-4 líneas)
2. **Hallazgos críticos** (máximo 5 bullets)
3. **Análisis de causas raíz** (NO QUIERE = motivacional, NO SABE = entrenamiento, 
   NO PUEDE = recursos)
4. **Recomendaciones accionables** (5 bullets con horizonte temporal)
5. **Meta SMART** para los próximos 3 meses
6. **Riesgos de no actuar**

Sé específico con los números.
"""
    try:
        response = cliente_ia.chat.completions.create(
            model=MODELO_GROQ,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=2000
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error al consultar Groq: {e}"


# ============ CONFIG STREAMLIT ============
st.set_page_config(
    page_title="Dashboard PARE y SIGA - SST",
    page_icon="🦺",
    layout="wide"
)
# Forzar la carga de la clave cuando la app ya está corriendo
#cargar_clave_groq()
print(f"🔍 DEBUG: Longitud de la clave cargada: {len(GROQ_API_KEY)}")  

# --- Rutas de datos locales (Fallback) ---
RUTA_EXCEL   = r"C:\Users\SEGURIDAD\Pare y Siga\datos_pare_unificados.xlsx"
RUTA_CSV     = r"C:\Users\SEGURIDAD\Pare y Siga\datos_pare_unificados.csv"
RUTA_PARQUET = r"C:\Users\SEGURIDAD\Pare y Siga\datos_pare_unificados.parquet"

# ============ CARGA DESDE GOOGLE SHEETS ============ 
def cargar_datos_google_sheets():
    try:
        import os
        from google.oauth2.service_account import Credentials
        import gspread
        
        # 1. Intentar leer el archivo local (para cuando pruebas en Codespaces)
        if os.path.exists("google_creds.json"):
            creds = Credentials.from_service_account_file(
                "google_creds.json",
                scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            )
        else:
            # 2. Si no hay archivo, leer desde los Secrets de Streamlit (para la nube)
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(
                creds_dict,
                scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            )
            
        client = gspread.authorize(creds)
        
        # Usa la URL y el nombre de hoja desde secrets
        url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        worksheet_name = st.secrets["connections"]["gsheets"]["worksheet"]
        
        sheet = client.open_by_url(url)
        worksheet = sheet.worksheet(worksheet_name)
        
        data = worksheet.get_all_values()
        df = pd.DataFrame(data[1:], columns=data[0])
        df = df.dropna(how="all")
        print(f"✅ Datos cargados desde Google Sheets: {df.shape}")
        return df
    except Exception as e:
        print(f"⚠️ Error con Google Sheets: {e}")
        raise e 

# ============ CARGA DE DATOS ============
@st.cache_data(ttl=300, show_spinner=False)
def cargar_datos():
    """Carga desde Google Sheets (preferido) o Excel local (fallback)."""

    try:
        df = cargar_datos_google_sheets()
        print("📡 Fuente: Google Sheets (en vivo)")
    except Exception as e:
        print(f"⚠️ Error con Google Sheets: {e}")
        print("📁 Usando Excel local como fallback...")
        if os.path.exists(RUTA_EXCEL):
            df = pd.read_excel(RUTA_EXCEL, engine='openpyxl')
        elif os.path.exists(RUTA_CSV):
            df = pd.read_csv(RUTA_CSV, sep=';', encoding='utf-8-sig')
        elif os.path.exists(RUTA_PARQUET):
            df = pd.read_parquet(RUTA_PARQUET)
        else:
            raise FileNotFoundError("No se pudo cargar desde Google Sheets ni archivos locales")

    # ============================================================
    # NORMALIZAR NOMBRES DE COLUMNAS
    # ============================================================
    df.columns = [
        str(c).replace('\ufeff', '').replace('"', '').replace('\n', ' ').strip()
        for c in df.columns
    ]
    df.columns = [' '.join(c.split()) for c in df.columns]

    # ============================================================
    # UNIFICAR COLUMNAS DE REPORTANTE
    # ============================================================
    columnas_reportante = [
        'PERSONA QUE REPORTA HUARAZ',
        'PERSONA QUE REPORTA CHIMBOTE',
        'PERSONA QUIEN REPORTA ALJER',
        'PERSONA QUIEN REPORTA CHOCOPE',
        'PERSONA QUIEN REPORTA GUADALUPE',
        'PERSONA QUIEN REPORTA SOL NACIENTE',
        'PERSONA QUIEN REPORTA PMA TRUJILLO',
        'PERSONA QUIEN REPORTA HUAMACHUCO ',
        'PERSONA QUIEN REPORTA HUAMACHUCO',
    ]
    cols_existentes = [c for c in columnas_reportante if c in df.columns]

    if cols_existentes:
        df['reportante'] = (
            df[cols_existentes]
            .apply(
                lambda row: next(
                    (str(v).strip() for v in row
                     if pd.notna(v) and str(v).strip() != ''),
                    'SIN DATO'
                ),
                axis=1
            )
        )
    else:
        df['reportante'] = 'SIN DATO'

    # ============================================================
    # DERIVAR CDA
    # ============================================================
    mapa_sede = {
        'PERSONA QUE REPORTA HUARAZ': 'HUARAZ',
        'PERSONA QUE REPORTA CHIMBOTE': 'CHIMBOTE',
        'PERSONA QUIEN REPORTA ALJER': 'ALJER',
        'PERSONA QUIEN REPORTA CHOCOPE': 'CHOCOPE',
        'PERSONA QUIEN REPORTA GUADALUPE': 'GUADALUPE',
        'PERSONA QUIEN REPORTA SOL NACIENTE': 'SOL NACIENTE',
        'PERSONA QUIEN REPORTA PMA TRUJILLO': 'PMA TRUJILLO',
        'PERSONA QUIEN REPORTA HUAMACHUCO ': 'HUAMACHUCO',
    }

    def obtener_sede(row):
        for col, sede in mapa_sede.items():
            if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != '':
                return sede
        return None

    col_cda = None
    for c in df.columns:
        c_up = c.upper()
        if 'ELIJE EL CDA' in c_up or 'ELIGE EL CDA' in c_up:
            col_cda = c
            break

    if col_cda is not None:
        print(f"✅ Columna CDA directa: {repr(col_cda)}")
        df['cda'] = df[col_cda].fillna('SIN DATO').astype(str).str.strip().str.upper()
    else:
        df['cda'] = df.apply(obtener_sede, axis=1)
        df['cda'] = df['cda'].fillna('SIN DATO').astype(str).str.upper()

    # ============================================================
    # FECHA Y LIMPIEZA
    # ============================================================
    col_fecha = None
    for c in df.columns:
        if 'marca' in c.lower() and 'temporal' in c.lower():
            col_fecha = c
            break
    if col_fecha is None:
        col_fecha = df.columns[0]

    print(f"✅ Columna de fecha: {repr(col_fecha)}")
    df['fecha'] = pd.to_datetime(df[col_fecha], errors='coerce', dayfirst=True)

    filas_antes = len(df)
    df = df[df['fecha'].notna()].copy()
    if filas_antes - len(df) > 0:
        print(f"⚠️ Eliminadas {filas_antes - len(df)} filas sin fecha válida")

    df['anio'] = df['fecha'].dt.year.astype('Int64')

    # ============================================================
    # UNIFICAR COLUMNAS ESPECÍFICAS
    # ============================================================
    def coalesce(df, cols):
        cols = [c for c in cols if c in df.columns]
        if not cols:
            return pd.Series(['SIN DATO'] * len(df), index=df.index)
        s = df[cols].bfill(axis=1).iloc[:, 0]
        return s.fillna('SIN DATO').astype(str).str.strip().str.upper()

    # Colaborador observado
    cols_obs = [c for c in df.columns if 'NOMBRES Y APELLIDOS DEL COLABORADOR' in c.upper()]
    df['colaborador_observado'] = coalesce(df, cols_obs)

    # Área
    cols_area = [
        c for c in df.columns
        if 'AREA A LA QUE PERTENECE' in c.upper()
        or 'AREA DEL COLABORADOR' in c.upper()
    ]
    print(f"🔍 Columnas de área: {cols_area}")
    df['area'] = coalesce(df, cols_area)

    # Cargo (búsqueda ampliada)
    cols_cargo = [
        c for c in df.columns
        if 'CARGO DEL COLABORADOR' in c.upper()
        or 'MARCA EL CARGO' in c.upper()
        or 'SELECCIONE EL CARGO' in c.upper()
    ]
    print(f"🔍 Columnas de cargo: {cols_cargo}")
    df['cargo_observado'] = coalesce(df, cols_cargo)

    # Acto inseguro
    cols_acto = [
        c for c in df.columns
        if 'ACTOS INSEGUROS' in c.upper() or 'OTROS' in c.upper()
    ]
    print(f"🔍 Columnas de acto inseguro: {cols_acto}")
    df['acto_inseguro'] = coalesce(df, cols_acto)

    # Acto seguro
    col_seguro = next((c for c in df.columns if c.upper().startswith('ACTO SEGURO')), None)
    if col_seguro:
        df['acto_seguro'] = df[col_seguro].fillna('SIN DATO').astype(str).str.strip().str.upper()
    else:
        df['acto_seguro'] = 'SIN DATO'

    # Causa
    col_causa = next((c for c in df.columns if 'SELECCIONE LA CAUSA' in c.upper()), None)
    if col_causa:
        df['causa'] = df[col_causa].fillna('SIN DATO').astype(str).str.strip().str.upper()
    else:
        df['causa'] = 'SIN DATO'

    # Riesgo
    col_riesgo = next((c for c in df.columns if 'RIESGO POTENCIAL' in c.upper()), None)
    if col_riesgo:
        df['riesgo'] = df[col_riesgo].fillna('SIN DATO').astype(str).str.strip().str.upper()
    else:
        df['riesgo'] = 'SIN DATO'

    # Tipo de tarjeta
    col_tipo = next((c for c in df.columns if 'TIPO DE TARJETA' in c.upper()), None)
    if col_tipo:
        df['tipo_tarjeta'] = df[col_tipo].fillna('SIN DATO').astype(str).str.strip().str.upper()
    else:
        df['tipo_tarjeta'] = 'SIN DATO'

    # Acción correctiva
    col_accion = next((c for c in df.columns if 'ACCION CORRECTIVA' in c.upper()), None)
    if col_accion:
        df['accion_correctiva'] = df[col_accion].fillna('SIN DATO').astype(str).str.strip()
    else:
        df['accion_correctiva'] = 'SIN DATO'

    # ============================================================
    # ✅ FIX DEFINITIVO: Limpiar vacíos en TODAS las columnas clave
    # ============================================================
    columnas_a_limpiar = ['area', 'cargo_observado', 'causa', 'acto_inseguro', 'reportante', 'cda']
    
    for col in columnas_a_limpiar:
        if col in df.columns:
            df[col] = df[col].replace(['', ' ', 'nan', 'NaN', 'NAN', 'None', 'NONE', 'null', 'NULL'], 'SIN DATO')
            df[col] = df[col].fillna('SIN DATO')
            df[col] = df[col].astype(str).str.strip().str.upper()
            
    # ============================================================
    # CAMPOS DE FECHA AVANZADOS
    # ============================================================
    meses_orden = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
                   'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']

    df['mes_orden'] = df['fecha'].dt.month
    df['mes_etiqueta'] = df['mes_orden'].apply(
        lambda m: meses_orden[int(m)-1] if pd.notna(m) else None
    )

    df['dia_semana'] = df['fecha'].dt.day_name()
    df['dia_semana_num'] = df['fecha'].dt.dayofweek + 1
    df['anio_mes'] = df['fecha'].dt.strftime('%Y-%m')

    df['semana_iso'] = df['fecha'].dt.isocalendar().week.astype('Int64')
    df['anio_semana'] = df['fecha'].dt.isocalendar().year.astype('Int64')
    df['anio_semana_label'] = (df['anio_semana'].astype(str) + '-S' +
                                df['semana_iso'].astype(str).str.zfill(2))
    df['inicio_semana'] = df['fecha'].dt.to_period('W').dt.start_time
    df['fin_semana'] = df['fecha'].dt.to_period('W').dt.end_time

    print(f"✅ Dataset procesado: {df.shape}")
    print(f"   Rango: {df['fecha'].min().date()} → {df['fecha'].max().date()}")

    return df


# ⬇️ LLAMADA A LA FUNCIÓN
df = cargar_datos()


# ============ SIDEBAR: FILTROS ============
st.sidebar.header("🔎 Filtros globales")

anios = sorted([int(a) for a in df['anio'].dropna().unique()]) if 'anio' in df.columns else []
anio_sel = st.sidebar.multiselect("Año", anios, default=anios)

cdas = sorted(df['cda'].dropna().unique().tolist())
cda_sel = st.sidebar.multiselect("CDA / Sede", cdas, default=cdas)

areas = sorted(df['area'].dropna().unique().tolist())
area_sel = st.sidebar.multiselect("Área", areas, default=areas)

cargos = sorted(df['cargo_observado'].dropna().unique().tolist())
cargo_sel = st.sidebar.multiselect("Cargo observado", cargos, default=cargos)

tipos = sorted(df['tipo_tarjeta'].dropna().unique().tolist())
tipo_sel = st.sidebar.multiselect("Tipo de tarjeta", tipos, default=tipos)

riesgos = sorted(df['riesgo'].dropna().unique().tolist())
riesgo_sel = st.sidebar.multiselect("Nivel de riesgo", riesgos, default=riesgos)

# --- Rango de fechas ---
df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
fechas_validas = df['fecha'].dropna()

if len(fechas_validas) == 0:
    st.error("⚠️ No hay fechas válidas en los datos.")
    st.stop()

fmin = fechas_validas.min().date()
fmax = fechas_validas.max().date()

if fmin > fmax:
    fmin, fmax = fmax, fmin

rango = st.sidebar.date_input(
    "Rango de fechas",
    value=(fmin, fmax),
    min_value=fmin,
    max_value=fmax
)

# --- Filtro de MES ---
st.sidebar.divider()
st.sidebar.subheader("📅 Filtro temporal")
meses_orden = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
               'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
meses_disponibles = sorted(
    [m for m in df['mes_etiqueta'].dropna().unique().tolist() if m in meses_orden],
    key=lambda m: meses_orden.index(m)
)
mes_sel = st.sidebar.multiselect("Mes", meses_disponibles, default=meses_disponibles)

# --- Filtro de SEMANA ---
st.sidebar.markdown("**📆 Filtro por semana**")
filtro_semana = st.sidebar.radio(
    "Modo de filtrado:",
    ["Todas las semanas", "Últimas 4 semanas", "Últimas 8 semanas",
     "Últimas 12 semanas", "Semana específica"],
    index=0,
    horizontal=False,
    key="filtro_semana_radio"
)

semana_especifica = None
semanas_recientes = None

if filtro_semana == "Últimas 4 semanas":
    semanas_recientes = 4
elif filtro_semana == "Últimas 8 semanas":
    semanas_recientes = 8
elif filtro_semana == "Últimas 12 semanas":
    semanas_recientes = 12
elif filtro_semana == "Semana específica":
    semanas_disponibles = sorted(
        df['anio_semana_label'].dropna().unique().tolist(), reverse=True
    )
    semana_especifica = st.sidebar.selectbox(
        "Selecciona la semana:",
        semanas_disponibles[:52]
    )

# --- Análisis detallado por cargo ---
st.sidebar.divider()
st.sidebar.header("🎯 Análisis detallado")
cargos_unicos = sorted(df['cargo_observado'].dropna().astype(str).unique().tolist())
cargo_detalle = st.sidebar.selectbox(
    "Selecciona un cargo:", ["— Ninguno —"] + cargos_unicos
)

# --- IP en sidebar ---
if UTILS_DISPONIBLE:
    st.sidebar.divider()
    mostrar_ip_sidebar(st, puerto=8501, estilo="caja")


# ============ APLICAR FILTROS ============
dff = df[
    df['anio'].isin(anio_sel) &
    df['mes_etiqueta'].isin(mes_sel) &
    df['cda'].isin(cda_sel) &
    df['area'].isin(area_sel) &
    df['cargo_observado'].isin(cargo_sel) &
    df['tipo_tarjeta'].isin(tipo_sel) &
    df['riesgo'].isin(riesgo_sel) &
    (df['fecha'].dt.date >= rango[0]) &
    (df['fecha'].dt.date <= rango[1])
].copy()

if semanas_recientes is not None:
    semanas_unicas = sorted(dff['anio_semana_label'].dropna().unique(), reverse=True)
    semanas_sel = semanas_unicas[:semanas_recientes]
    dff = dff[dff['anio_semana_label'].isin(semanas_sel)].copy()
elif semana_especifica is not None:
    dff = dff[dff['anio_semana_label'] == semana_especifica].copy()


# ============ HEADER ============
st.title("🦺 Dashboard de Observaciones de Comportamiento — PARE y SIGA")
st.caption(
    f"📍 {', '.join(cda_sel[:3])}{'...' if len(cda_sel) > 3 else ''} · "
    f"Registros filtrados: **{len(dff):,}** de {len(df):,}"
)


# ============ KPIs ============
c1, c2, c3, c4, c5, c6 = st.columns(6)
total            = len(dff)
reportantes_uniq = dff['reportante'].nunique()
areas_n          = dff['area'].nunique()
causas_n         = dff[dff['causa'] != 'SIN DATO']['causa'].nunique()
pct_pare         = (dff['tipo_tarjeta'].str.contains('PARE', na=False).mean() * 100) if total else 0
pct_alto         = (dff['riesgo'].eq('ALTO').mean() * 100) if total else 0

c1.metric("📋 Total observaciones", f"{total:,}")
c2.metric("👥 Reportantes únicos",   reportantes_uniq)
c3.metric("🏭 Áreas cubiertas",      areas_n)
c4.metric("⚠️ Causas distintas",     causas_n)
c5.metric("✅ % Tarjeta PARE",       f"{pct_pare:.1f}%")
c6.metric("🚦 % Riesgo ALTO",        f"{pct_alto:.1f}%")

st.divider()


# ============ FILA 1: Tendencia + Tipo de tarjeta ============
col1, col2 = st.columns([2, 1])

with col1:
    trend = (dff.groupby([dff['fecha'].dt.to_period('M'), 'anio'])
                .size().reset_index(name='observaciones'))
    trend['fecha'] = trend['fecha'].dt.to_timestamp()
    trend['anio'] = trend['anio'].astype(str)
    fig = px.line(trend, x='fecha', y='observaciones', color='anio',
                  markers=True, title="📈 Tendencia mensual por año",
                  color_discrete_sequence=px.colors.qualitative.Set1)
    fig.update_traces(line_width=3)
    st.plotly_chart(fig, use_container_width=True)

with col2:
    tp = dff['tipo_tarjeta'].value_counts().reset_index()
    tp.columns = ['tipo', 'n']
    fig = px.pie(tp, names='tipo', values='n', hole=0.5,
                 title="🎫 Tipo de tarjeta",
                 color_discrete_sequence=px.colors.qualitative.Set2)
    st.plotly_chart(fig, use_container_width=True)


# ============ TENDENCIA SEMANAL ============
st.subheader("📆 Tendencia semanal")

trend_semanal = (dff.groupby('anio_semana_label')
                    .agg(observaciones=('fecha', 'size'),
                         inicio=('inicio_semana', 'first'))
                    .reset_index()
                    .sort_values('anio_semana_label'))

if len(trend_semanal) > 0:
    fig = px.line(
        trend_semanal, x='anio_semana_label', y='observaciones',
        markers=True, title="Observaciones por semana",
        labels={'anio_semana_label': 'Semana', 'observaciones': 'Nº observaciones'},
        text='observaciones'
    )
    fig.update_traces(
        line_color='#1f77b4', line_width=3,
        marker=dict(size=12), textposition='top center',
        textfont=dict(size=11, color='black')
    )

    promedio = trend_semanal['observaciones'].mean()
    fig.add_hline(y=promedio, line_dash='dash', line_color='orange',
                  annotation_text=f"Promedio: {promedio:.0f}",
                  annotation_position="right")

    fig.update_layout(height=400, hovermode='x unified',
                      xaxis=dict(type='category', tickangle=-45))
    st.plotly_chart(fig, use_container_width=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📊 Semanas analizadas", len(trend_semanal))
    k2.metric("📈 Promedio semanal", f"{promedio:.1f}")
    k3.metric("⬆️ Semana más alta", f"{trend_semanal['observaciones'].max()}")
    k4.metric("⬇️ Semana más baja", f"{trend_semanal['observaciones'].min()}")

    with st.expander("📄 Ver Top 5 semanas con más observaciones"):
        top_sem = (trend_semanal
                   .nlargest(5, 'observaciones')[['anio_semana_label', 'inicio', 'observaciones']]
                   .copy())
        top_sem.columns = ['Semana', 'Inicio', 'Observaciones']
        top_sem['Inicio'] = pd.to_datetime(top_sem['Inicio']).dt.strftime('%d/%m/%Y')
        st.dataframe(top_sem, use_container_width=True, hide_index=True)
else:
    st.info("No hay datos semanales con los filtros actuales.")


# ============ DETALLE DE REPORTANTES POR SEMANA ============
st.subheader("👥 Reportantes de la semana — Desglose PARE/SIGA")

semanas_disp_detalle = sorted(dff['anio_semana_label'].dropna().unique().tolist(), reverse=True)

if len(semanas_disp_detalle) == 0:
    st.info("No hay datos de reportantes en el filtro actual.")
else:
    col_sel1, col_sel2 = st.columns([2, 2])
    with col_sel1:
        semana_detalle = st.selectbox(
            "📅 Selecciona la semana a detallar:",
            semanas_disp_detalle[:20],
            index=0,
            key="semana_detalle_reportantes"
        )

    df_sem = dff[dff['anio_semana_label'] == semana_detalle].copy()

    if len(df_sem) == 0:
        st.warning("No hay datos para esta semana.")
    else:
        fecha_ini = df_sem['inicio_semana'].min()
        fecha_fin = df_sem['fin_semana'].max()
        st.caption(f"📅 Período: **{fecha_ini.strftime('%d/%m/%Y')}** al "
                   f"**{fecha_fin.strftime('%d/%m/%Y')}**  |  "
                   f"Total: **{len(df_sem)} observaciones**")

        total_pare_sem = (df_sem['tipo_tarjeta'].str.contains('PARE', na=False)).sum()
        total_siga_sem = (df_sem['tipo_tarjeta'].str.contains('SIGA', na=False)).sum()
        total_rep_sem  = df_sem['reportante'].nunique()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📋 Total tarjetas", f"{len(df_sem)}")
        k2.metric("✅ PARE", f"{total_pare_sem}")
        k3.metric("🎫 SIGA", f"{total_siga_sem}")
        k4.metric("👥 Reportantes activos", f"{total_rep_sem}")

        st.markdown("#### 📊 Ranking de reportantes de la semana")

        ranking_rep = (df_sem.groupby('reportante')
                            .agg(total=('tipo_tarjeta', 'size'),
                                 pare=('tipo_tarjeta', lambda x: x.str.contains('PARE', na=False).sum()),
                                 siga=('tipo_tarjeta', lambda x: x.str.contains('SIGA', na=False).sum()),
                                 areas=('area', lambda x: ', '.join(sorted(x.unique())[:3])))
                            .reset_index()
                            .sort_values('total', ascending=False))

        ranking_rep['pct_pare'] = ((ranking_rep['pare'] / ranking_rep['total']) * 100).round(1)
        ranking_rep.columns = ['Reportante', 'Total', 'PARE', 'SIGA', 'Áreas', '% PARE']

        st.dataframe(
            ranking_rep, use_container_width=True, hide_index=True,
            column_config={
                'Total': st.column_config.NumberColumn('Total', format="%d"),
                'PARE': st.column_config.NumberColumn('PARE ✅', format="%d"),
                'SIGA': st.column_config.NumberColumn('SIGA 🎫', format="%d"),
                '% PARE': st.column_config.ProgressColumn(
                    '% PARE', min_value=0, max_value=100, format="%.1f%%"),
            }
        )

        st.markdown("#### 📈 Distribución visual por reportante")

        graf_data = ranking_rep[['Reportante', 'PARE', 'SIGA']].copy()
        graf_data = graf_data.melt(id_vars='Reportante',
                                    value_vars=['PARE', 'SIGA'],
                                    var_name='Tipo', value_name='Cantidad')

        fig = px.bar(graf_data, x='Reportante', y='Cantidad',
                     color='Tipo', barmode='stack',
                     title=f"Tarjetas por reportante — {semana_detalle}",
                     color_discrete_map={'PARE': '#EF553B', 'SIGA': '#00CC96'},
                     text='Cantidad')
        fig.update_traces(textposition='inside')
        fig.update_layout(xaxis_tickangle=-45, height=450,
                          legend_title="Tipo", yaxis_title="Nº tarjetas")
        st.plotly_chart(fig, use_container_width=True)


# ============ RIESGO POTENCIAL ============
st.subheader("🚦 Distribución por nivel de riesgo potencial")

col_r1, col_r2 = st.columns([1, 2])
color_map = {'ALTO': '#EF553B', 'MEDIO': '#FFA15A',
             'BAJO': '#00CC96', 'SIN DATO': '#999999'}

with col_r1:
    riesgo_cnt = dff['riesgo'].value_counts().reset_index()
    riesgo_cnt.columns = ['riesgo', 'n']
    fig = px.pie(riesgo_cnt, names='riesgo', values='n', hole=0.5,
                 title="% por nivel de riesgo",
                 color='riesgo', color_discrete_map=color_map)
    st.plotly_chart(fig, use_container_width=True)

with col_r2:
    riesgo_area = pd.crosstab(dff['area'], dff['riesgo'])
    orden = [c for c in ['ALTO', 'MEDIO', 'BAJO', 'SIN DATO']
             if c in riesgo_area.columns]
    riesgo_area = riesgo_area[orden]
    riesgo_area['total'] = riesgo_area.sum(axis=1)
    riesgo_area = (riesgo_area.sort_values('total', ascending=False)
                              .head(10).drop(columns='total'))
    fig = px.bar(riesgo_area, barmode='stack',
                 title="Riesgo por área (top 10)",
                 color_discrete_map=color_map)
    st.plotly_chart(fig, use_container_width=True)


# ============ TOP ÁREAS Y CARGOS ============
col1, col2 = st.columns(2)

with col1:
    top_area = dff['area'].value_counts().head(10).reset_index()
    top_area.columns = ['area', 'n']
    fig = px.bar(top_area, x='n', y='area', orientation='h',
                 title="🏭 Top 10 áreas observadas",
                 color='n', color_continuous_scale='Blues')
    fig.update_layout(yaxis={'categoryorder': 'total ascending'})
    st.plotly_chart(fig, use_container_width=True)

with col2:
    top_cargo = dff['cargo_observado'].value_counts().head(10).reset_index()
    top_cargo.columns = ['cargo', 'n']
    fig = px.bar(top_cargo, x='n', y='cargo', orientation='h',
                 title="👷 Top 10 cargos observados",
                 color='n', color_continuous_scale='Oranges')
    fig.update_layout(yaxis={'categoryorder': 'total ascending'})
    st.plotly_chart(fig, use_container_width=True)


# ============ CAUSAS ============
st.subheader("⚠️ Causas de actos inseguros (NO QUIERE / NO SABE / NO PUEDE)")

causas_validas = dff[
    (dff['causa'] != 'SIN DATO') &
    (dff['causa'] != '') &
    (dff['causa'].notna())
]

pareto = causas_validas['causa'].value_counts().reset_index()
pareto.columns = ['causa', 'n']

if len(pareto) > 0:
    pareto['pct'] = (pareto['n'] / pareto['n'].sum() * 100).round(1)
    color_map_causa = {'NO QUIERE': '#EF553B', 'NO SABE': '#FFA15A', 'NO PUEDE': '#636EFA'}
    col_p1, col_p2 = st.columns(2)

    with col_p1:
        fig = px.bar(pareto, x='n', y='causa', orientation='h',
                     color='causa', color_discrete_map=color_map_causa,
                     title="Frecuencia por causa",
                     text=pareto.apply(lambda r: f"{r['n']} ({r['pct']}%)", axis=1))
        fig.update_traces(textposition='outside')
        fig.update_layout(showlegend=False,
                          yaxis={'categoryorder': 'total ascending'},
                          height=350)
        st.plotly_chart(fig, use_container_width=True)

    with col_p2:
        fig = px.pie(pareto, names='causa', values='n', hole=0.55,
                     color='causa', color_discrete_map=color_map_causa,
                     title="Distribución % por causa")
        fig.update_traces(textinfo='percent+label')
        fig.update_layout(height=350)
        st.plotly_chart(fig, use_container_width=True)

    causa_top = pareto.iloc[0]['causa']
    pct_top   = pareto.iloc[0]['pct']
    interpretacion = {
        'NO QUIERE': "Problema **motivacional / cultural** → refuerzo positivo, liderazgo visible, reconocimiento.",
        'NO SABE':   "Problema de **entrenamiento** → reentrenamiento específico y verificación de competencia.",
        'NO PUEDE':  "Problema de **recursos / condiciones** → revisar EPP, herramientas, ergonomía y diseño del puesto."
    }
    st.info(f"**Causa dominante:** `{causa_top}` ({pct_top}% de los actos). "
            f"{interpretacion.get(causa_top, '')}")


# ============ PARETO 80/20 ============
st.subheader("📊 Pareto 80/20 de actos inseguros específicos")

actos_validos = dff[
    (dff['acto_inseguro'] != 'SIN DATO') &
    (dff['acto_inseguro'] != '') &
    (dff['acto_inseguro'].notna())
]

actos_p = actos_validos['acto_inseguro'].value_counts().reset_index()
actos_p.columns = ['acto', 'n']

if len(actos_p) >= 5:
    actos_p['pct']      = actos_p['n'] / actos_p['n'].sum() * 100
    actos_p['pct_acum'] = actos_p['pct'].cumsum()

    actos_p['acto_corto'] = actos_p['acto'].apply(
        lambda x: x[:35] + '...' if len(x) > 35 else x
    )

    fig = go.Figure()
    fig.add_bar(x=actos_p['acto_corto'], y=actos_p['n'],
                name='Frecuencia', marker_color='#EF553B',
                text=actos_p['n'], textposition='outside',
                hovertemplate='<b>%{x}</b><br>Frecuencia: %{y}<extra></extra>')
    fig.add_trace(go.Scatter(x=actos_p['acto_corto'], y=actos_p['pct_acum'],
                             name='% Acumulado', yaxis='y2',
                             line=dict(color='#636EFA', width=3),
                             mode='lines+markers'))
    fig.add_hline(y=80, line_dash='dash', line_color='green',
                  annotation_text="80%", yref='y2')
    fig.update_layout(
        title="Pareto de actos inseguros específicos (regla 80/20)",
        xaxis=dict(
            title='Actos inseguros específicos',
            tickangle=-45,
            tickfont=dict(size=10)
        ),
        yaxis=dict(title='Frecuencia'),
        yaxis2=dict(title='% Acumulado', overlaying='y', side='right', range=[0, 105]),
        height=600,
        margin=dict(b=200),
        legend=dict(x=0.75, y=1.15, orientation='h'))
    st.plotly_chart(fig, use_container_width=True)

    n_80 = max(len(actos_p[actos_p['pct_acum'] <= 80]), 1)
    top_80 = actos_p.head(n_80).copy()

    st.caption(f"💡 Los **{n_80}** actos más frecuentes explican el **80%** de los eventos.")

    st.markdown("### 🎯 Plan de acción sobre los desvíos críticos (Pareto 80%)")

    def clasificar_accion(acto):
        a = acto.upper()
        if any(k in a for k in ['LENTES','GUANTES','CASCO','ZAPATOS','RESPIRADOR','ARNES','TAPONES','EPP']):
            return ("🦺 EPP", "Refuerzo diario del uso de EPP + inspección aleatoria.")
        if any(k in a for k in ['SUBIR','BAJAR','TRANSPORTAR','CARGA','POSTURA','ERGONOM','LEVANTAR']):
            return ("💪 Ergonómico", "Capacitación en levantamiento seguro.")
        if any(k in a for k in ['VELOCIDAD','CONDUCIR','CINTURON','RETROCESO','ESTACIONAR','PASOS','REPARTO','MONTACARGA','CAMION']):
            return ("🚛 Conducción segura", "Reentrenamiento en manejo defensivo.")
        if any(k in a for k in ['HERRAMIENTA','EQUIPO','MANTENIMIENTO','INSPECCION','DANAD','DAÑAD']):
            return ("🔧 Herramientas y equipos", "Inspección y mantenimiento preventivo.")
        if any(k in a for k in ['ORDEN','LIMPIEZA','ALMACENAR','APILAR']):
            return ("🧹 Orden y limpieza (5S)", "Implementar 5S + auditoría semanal.")
        return ("⚠️ Comportamiento general", "Observación conductual focalizada.")

    recomendaciones = []
    for _, row in top_80.iterrows():
        categoria, accion = clasificar_accion(row['acto'])
        recomendaciones.append({
            'Desvío crítico': row['acto'],
            'Frecuencia': row['n'],
            '% Acumulado': f"{row['pct_acum']:.1f}%",
            'Categoría SST': categoria,
            'Acción recomendada': accion
        })

    df_rec = pd.DataFrame(recomendaciones)
    st.dataframe(df_rec, use_container_width=True, hide_index=True)

elif len(actos_p) > 0:
    fig = px.bar(actos_p, x='n', y='acto', orientation='h',
                 color='n', color_continuous_scale='Reds',
                 title="Frecuencia de actos inseguros específicos", text='n')
    fig.update_traces(textposition='outside')
    fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=500)
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay actos inseguros específicos en el filtro actual.")


# ============ TOP 15 ACTOS ============
st.subheader("🚨 Top 15 actos inseguros específicos observados")

if len(actos_validos) > 0:
    top_actos = (actos_validos['acto_inseguro'].value_counts().head(15).reset_index())
    top_actos.columns = ['acto', 'n']
    fig = px.bar(top_actos, x='n', y='acto', orientation='h',
                 color='n', color_continuous_scale='Reds',
                 title="Comportamientos inseguros más frecuentes")
    fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=550)
    st.plotly_chart(fig, use_container_width=True)


# ============ HEATMAP ÁREA × CAUSA ============
st.subheader("🔥 Mapa de calor: Área × Causa")

dff_area_valida = dff[
    (dff['area'] != 'SIN DATO') & (dff['area'] != '') & (dff['area'].notna()) &
    (dff['causa'] != 'SIN DATO') & (dff['causa'] != '') & (dff['causa'].notna())
]

pivot = pd.crosstab(dff_area_valida['area'], dff_area_valida['causa'])
top_areas  = dff_area_valida['area'].value_counts().head(10).index
top_causas = dff_area_valida['causa'].value_counts().head(10).index
pivot = pivot.loc[pivot.index.intersection(top_areas),
                  pivot.columns.intersection(top_causas)]

if pivot.size > 0:
    fig = px.imshow(pivot, aspect='auto', color_continuous_scale='YlOrRd',
                    text_auto=True, title="Concentración de causas por área")
    fig.update_layout(height=550)
    st.plotly_chart(fig, use_container_width=True)


# ============ HEATMAP CARGO × CAUSA ============
st.subheader("🚨 Comportamiento crítico por cargo (Cargo × Causa)")

dff_cargo_valido = dff[
    (dff['cargo_observado'] != 'SIN DATO') & (dff['cargo_observado'] != '') & (dff['cargo_observado'].notna()) &
    (dff['causa'] != 'SIN DATO') & (dff['causa'] != '') & (dff['causa'].notna())
]

pivot_cargo = pd.crosstab(dff_cargo_valido['cargo_observado'], dff_cargo_valido['causa'])
top_cargos_h = dff_cargo_valido['cargo_observado'].value_counts().head(10).index
top_causas_h = dff_cargo_valido['causa'].value_counts().head(10).index
pivot_cargo = pivot_cargo.loc[
    pivot_cargo.index.intersection(top_cargos_h),
    pivot_cargo.columns.intersection(top_causas_h)
]

if pivot_cargo.size > 0:
    fig = px.imshow(pivot_cargo, aspect='auto', color_continuous_scale='Reds',
                    text_auto=True,
                    title="Concentración de causas por cargo")
    fig.update_layout(height=550)
    st.plotly_chart(fig, use_container_width=True)


# ============ TOP 3 CAUSAS POR CARGO ============
st.subheader("📊 Top 3 causas por cargo")

df_causas = dff[
    (dff['causa'] != 'SIN DATO') & (dff['causa'] != '') & (dff['causa'].notna()) &
    (dff['cargo_observado'] != 'SIN DATO') & (dff['cargo_observado'] != '') & (dff['cargo_observado'].notna())
]

if len(df_causas) > 0:
    top_cargos = df_causas['cargo_observado'].value_counts().head(8).index
    sub = df_causas[df_causas['cargo_observado'].isin(top_cargos)]
    top3 = (sub.groupby(['cargo_observado', 'causa']).size().reset_index(name='n')
               .sort_values(['cargo_observado','n'], ascending=[True, False])
               .groupby('cargo_observado').head(3))
    fig = px.bar(top3, x='n', y='cargo_observado', color='causa',
                 orientation='h', barmode='stack',
                 title="Top 3 causas de actos inseguros por cargo",
                 color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=550)
    st.plotly_chart(fig, use_container_width=True)


# ============ RANKING CRITICIDAD ============
st.subheader("🎯 Ranking de criticidad — Score = Frecuencia × % Riesgo ALTO")

df_rank = dff[(dff['causa'] != 'SIN DATO') &
              (dff['riesgo'].isin(['ALTO', 'MEDIO', 'BAJO']))].copy()

if len(df_rank) > 0:
    ranking = (df_rank.groupby(['cargo_observado','causa'])
                      .agg(n=('causa','size'),
                           pct_alto=('riesgo', lambda x: (x=='ALTO').mean()*100))
                      .reset_index())
    idx_max = ranking.groupby('cargo_observado')['n'].idxmax()
    criticos = ranking.loc[idx_max].copy()
    totales = df_rank['cargo_observado'].value_counts().rename('total')
    criticos = criticos.merge(totales, left_on='cargo_observado',
                              right_index=True, how='left')
    criticos['pct']      = (criticos['n'] / criticos['total'] * 100).round(1)
    criticos['pct_alto'] = criticos['pct_alto'].round(1)
    criticos['score']    = (criticos['n'] * criticos['pct_alto'] / 100).round(1)
    criticos = criticos.sort_values('score', ascending=False).reset_index(drop=True)
    criticos = criticos[['cargo_observado','causa','n','pct','pct_alto','score']]
    criticos.columns = ['Cargo','Causa más crítica','Frecuencia','% del cargo',
                        '% Riesgo ALTO','Score']
    st.dataframe(criticos, use_container_width=True, hide_index=True)

    top_crit = criticos.head(10).copy()
    top_crit['label'] = top_crit['Cargo'] + " — " + top_crit['Causa más crítica']
    fig = px.bar(top_crit, x='Score', y='label', orientation='h',
                 color='Score', color_continuous_scale='Reds',
                 title="Top 10 comportamientos más críticos", text='% Riesgo ALTO')
    fig.update_traces(texttemplate='%{text}%', textposition='outside')
    fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=550)
    st.plotly_chart(fig, use_container_width=True)


# ============ SCORE DE CRITICIDAD PREDICTIVO (SCP) ============
st.divider()
st.subheader("🎯 Score de Criticidad Predictivo (SCP)")
st.caption("Score compuesto: Frecuencia (35%) + Severidad (35%) + Tendencia (20%) + Concentración (10%)")


@st.cache_data(show_spinner=False)
def calcular_scp_dashboard(df, min_obs=5):
    """Calcula el SCP por zona (sede + área + cargo)."""
    from sklearn.preprocessing import MinMaxScaler

    pare = df[df['tipo_tarjeta'].str.contains('PARE', na=False)].copy()
    pesos = {'ALTO': 3, 'MEDIO': 2, 'BAJO': 1}
    pare['peso_riesgo'] = pare['riesgo'].map(pesos).fillna(0)

    grupos = ['cda', 'area', 'cargo_observado']

    base = (pare.groupby(grupos)
                .agg(
                    frecuencia=('peso_riesgo', 'size'),
                    peso_total=('peso_riesgo', 'sum'),
                    n_causas=('causa', 'nunique'),
                    n_reportantes=('reportante', 'nunique'),
                ).reset_index())

    base['indice_severidad'] = (base['peso_total'] / (3 * base['frecuencia'])).round(3)
    base['indice_concentracion'] = (1 / (1 + base['n_causas'] + base['n_reportantes'] * 0.5)).round(3)

    pare['anio_mes'] = pare['fecha'].dt.to_period('M')
    ult_mes = pare['anio_mes'].max()
    meses_analisis = [ult_mes - i for i in range(5, -1, -1)]

    tendencias = []
    for grupo, sub in pare.groupby(grupos):
        serie = sub.groupby('anio_mes').size().reindex(meses_analisis, fill_value=0)
        if len(serie) >= 3 and serie.sum() > 0:
            x = np.arange(len(serie))
            y = serie.values
            try:
                pendiente = np.polyfit(x, y, 1)[0]
            except Exception:
                pendiente = 0
            media = serie.mean()
            pend_norm = pendiente / (media + 1) if media > 0 else 0
            pend_norm = float(np.clip(pend_norm, -1, 1))
        else:
            pend_norm = 0.0
        if not isinstance(grupo, tuple):
            grupo = (grupo,)
        fila = dict(zip(grupos, grupo))
        fila['pendiente'] = pend_norm
        tendencias.append(fila)

    tend_df = pd.DataFrame(tendencias)
    score_df = base.merge(tend_df, on=grupos, how='left')
    score_df['pendiente'] = score_df['pendiente'].fillna(0)

    # Normalización con MinMaxScaler
    score_df['log_frecuencia'] = np.log1p(score_df['frecuencia'])
    scaler = MinMaxScaler(feature_range=(0, 100))
    score_df['freq_norm'] = scaler.fit_transform(score_df[['log_frecuencia']])
    score_df['sev_norm']  = score_df['indice_severidad'] * 100
    score_df['tend_norm'] = (score_df['pendiente'] + 1) * 50
    score_df['conc_norm'] = score_df['indice_concentracion'] * 100

    score_df['SCP'] = (
        score_df['freq_norm'] * 0.35 +
        score_df['sev_norm']  * 0.35 +
        score_df['tend_norm'] * 0.20 +
        score_df['conc_norm'] * 0.10
    ).round(1)

    def clasificar(row):
        if row['frecuencia'] < min_obs:
            return '⚪ DATOS INSUFICIENTES'
        scp = row['SCP']
        if scp >= 70:   return '🔴 CRÍTICO'
        elif scp >= 50: return '🟠 ALTO'
        elif scp >= 30: return '🟡 MEDIO'
        else:           return '🟢 BAJO'

    score_df['nivel'] = score_df.apply(clasificar, axis=1)
    score_df = score_df.sort_values('SCP', ascending=False).reset_index(drop=True)
    return score_df


try:
    scp_df = calcular_scp_dashboard(dff, min_obs=5)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("🔴 CRÍTICAS",   (scp_df['nivel'] == '🔴 CRÍTICO').sum())
    k2.metric("🟠 ALTAS",      (scp_df['nivel'] == '🟠 ALTO').sum())
    k3.metric("🟡 MEDIAS",     (scp_df['nivel'] == '🟡 MEDIO').sum())
    k4.metric("🟢 BAJAS",      (scp_df['nivel'] == '🟢 BAJO').sum())
    k5.metric("⚪ SIN DATOS",  (scp_df['nivel'] == '⚪ DATOS INSUFICIENTES').sum())

    st.markdown("#### 🚨 Top 15 zonas críticas (con ≥5 observaciones)")
    top15 = scp_df[scp_df['nivel'] != '⚪ DATOS INSUFICIENTES'].head(15).copy()

    if len(top15) > 0:
        tabla_scp = top15[['cda', 'area', 'cargo_observado', 'frecuencia',
                            'indice_severidad', 'pendiente', 'SCP', 'nivel']].copy()
        tabla_scp.columns = ['Sede', 'Área', 'Cargo', 'Frecuencia',
                              'Índice Severidad', 'Tendencia', 'SCP', 'Nivel']
        st.dataframe(tabla_scp, use_container_width=True, hide_index=True)

        top10 = top15.head(10).copy()
        top10['label'] = (top10['cda'] + " — " +
                           top10['cargo_observado'] + "\n(" +
                           top10['area'] + ")")

        color_map_scp = {
            '🔴 CRÍTICO': '#EF553B',
            '🟠 ALTO':    '#FFA15A',
            '🟡 MEDIO':   '#FFD700',
            '🟢 BAJO':    '#00CC96',
        }

        fig = px.bar(
            top10, x='SCP', y='label', orientation='h',
            color='nivel', color_discrete_map=color_map_scp,
            title="Top 10 zonas por Score de Criticidad Predictivo",
            text='SCP'
        )
        fig.update_traces(texttemplate='%{text}', textposition='outside')
        fig.update_layout(
            yaxis={'categoryorder': 'total ascending'},
            height=500, showlegend=True, legend_title="Nivel"
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### 🏭 Distribución de zonas críticas por sede")
        sede_nivel = pd.crosstab(scp_df['cda'], scp_df['nivel'])
        orden_niveles = ['🔴 CRÍTICO', '🟠 ALTO', '🟡 MEDIO', '🟢 BAJO', '⚪ DATOS INSUFICIENTES']
        orden_niveles = [n for n in orden_niveles if n in sede_nivel.columns]
        sede_nivel = sede_nivel[orden_niveles]

        fig = px.bar(sede_nivel, barmode='stack',
                     title="Zonas por nivel de criticidad y sede",
                     color_discrete_map=color_map_scp)
        fig.update_layout(xaxis_title="Sede", yaxis_title="N° de zonas",
                          height=400, legend_title="Nivel")
        st.plotly_chart(fig, use_container_width=True)

        criticos = scp_df[scp_df['nivel'] == '🔴 CRÍTICO']
        if len(criticos) > 0:
            top1 = criticos.iloc[0]
            st.error(f"""
            **🚨 INTERVENCIÓN INMEDIATA REQUERIDA**

            - **Zona:** {top1['cda']} · {top1['area']} · {top1['cargo_observado']}
            - **Frecuencia:** {top1['frecuencia']} observaciones
            - **Índice de severidad:** {top1['indice_severidad']:.2f} (1.0 = todos ALTO)
            - **SCP:** {top1['SCP']}
            - **Acción:** Intervención inmediata esta semana
            """)

        st.download_button(
            "⬇️ Descargar Score completo (CSV)",
            scp_df.to_csv(index=False, sep=';').encode('utf-8-sig'),
            "score_criticidad_predictivo.csv", "text/csv", key="dl_scp"
        )
    else:
        st.info("No hay zonas con suficientes observaciones para calcular el SCP.")

except ImportError:
    st.warning("⚠️ Falta instalar scikit-learn: `pip install scikit-learn`")
except Exception as e:
    st.error(f"Error al calcular SCP: {e}")


# ============ PREDICCIÓN TEMPORAL DEL SCP ============
st.divider()
st.subheader("🔮 Predicción Temporal del SCP")
st.caption("Proyección a 3 meses alineada con el ciclo PHVA (ISO 45001). "
           "Indicador direccional, no predicción exacta.")


@st.cache_data(show_spinner=False)
def predecir_scp_dashboard(dff, min_obs=5, meses_horizonte=3):
    """Proyección del SCP usando la MISMA lógica que calcular_scp_dashboard."""

    # ✅ PASO 1: SCP base
    scp_base = calcular_scp_dashboard(dff, min_obs=min_obs)

    # ✅ PASO 2: Proyecciones de frecuencia
    pare = dff[dff['tipo_tarjeta'].str.contains('PARE', na=False)].copy()
    pesos = {'ALTO': 3, 'MEDIO': 2, 'BAJO': 1}
    pare['peso_riesgo'] = pare['riesgo'].map(pesos).fillna(0)
    pare['anio_mes'] = pare['fecha'].dt.to_period('M')

    grupos = ['cda', 'area', 'cargo_observado']

    tendencias = []
    for grupo, sub in pare.groupby(grupos):
        if not isinstance(grupo, tuple):
            grupo = (grupo,)
        serie = sub.groupby('anio_mes').size().sort_index()
        fila = dict(zip(grupos, grupo))

        if len(serie) < 3:
            fila.update({
                'proyeccion_m1': serie.mean() if len(serie) > 0 else 0,
                'proyeccion_m2': serie.mean() if len(serie) > 0 else 0,
                'proyeccion_m3': serie.mean() if len(serie) > 0 else 0,
                'media_6m': serie.mean() if len(serie) > 0 else 0,
            })
            tendencias.append(fila)
            continue

        serie.index = serie.index.to_timestamp()
        serie_6m = serie.tail(6)
        x = np.arange(len(serie_6m))
        y = serie_6m.values
        try:
            pendiente, intercepto = np.polyfit(x, y, 1)
        except Exception:
            pendiente, intercepto = 0, serie_6m.mean()

        proyecciones = []
        for i in range(1, meses_horizonte + 1):
            y_pred = pendiente * (len(serie_6m) + i - 1) + intercepto
            proyecciones.append(max(0, round(y_pred, 1)))

        fila.update({
            'proyeccion_m1': proyecciones[0] if len(proyecciones) > 0 else 0,
            'proyeccion_m2': proyecciones[1] if len(proyecciones) > 1 else 0,
            'proyeccion_m3': proyecciones[2] if len(proyecciones) > 2 else 0,
            'media_6m': round(serie_6m.mean(), 2),
        })
        tendencias.append(fila)

    tend_df = pd.DataFrame(tendencias)

    # ✅ PASO 3: Combinar SCP base con proyecciones
    score_df = scp_base.merge(tend_df, on=grupos, how='left')
    for col in ['proyeccion_m1', 'proyeccion_m2', 'proyeccion_m3', 'media_6m']:
        score_df[col] = score_df[col].fillna(0)

    # ✅ PASO 4: Proyectar SCP usando el SCP real
    def calcular_delta(row, proyeccion):
        if row['media_6m'] == 0 or row['frecuencia'] < min_obs:
            return 0
        delta_pct = (proyeccion - row['media_6m']) / row['media_6m']
        delta_pct = np.clip(delta_pct, -0.3, 0.3)
        factor = 0.6 if row['indice_severidad'] >= 0.7 else (0.45 if row['indice_severidad'] >= 0.5 else 0.3)
        return delta_pct * 100 * factor

    score_df['delta_m1'] = score_df.apply(lambda r: calcular_delta(r, r['proyeccion_m1']), axis=1)
    score_df['delta_m2'] = score_df.apply(lambda r: calcular_delta(r, r['proyeccion_m2']), axis=1)
    score_df['delta_m3'] = score_df.apply(lambda r: calcular_delta(r, r['proyeccion_m3']), axis=1)

    score_df['SCP_m1'] = (score_df['SCP'] + score_df['delta_m1']).clip(30, 95).round(1)
    score_df['SCP_m2'] = (score_df['SCP'] + score_df['delta_m2']).clip(30, 95).round(1)
    score_df['SCP_m3'] = (score_df['SCP'] + score_df['delta_m3']).clip(30, 95).round(1)

    # ✅ PASO 5: Clasificación
    def meses_hasta(row, umbral=70):
        if row['frecuencia'] < min_obs: return '⚪ SIN DATOS'
        if row['SCP'] >= umbral: return '🚨 YA CRÍTICO'
        velocidad = (row['SCP_m3'] - row['SCP']) / 3 if row['SCP_m3'] > row['SCP'] + 0.5 else 0
        if velocidad <= 0.5: return '✅ ESTABLE'
        meses = (umbral - row['SCP']) / velocidad
        if meses > 12: return '🟢 >12 meses'
        elif meses > 6: return '🟡 6-12 meses'
        elif meses > 3: return '🟠 3-6 meses'
        else: return f'🔴 {int(np.ceil(meses))} meses'

    score_df['meses_critico'] = score_df.apply(meses_hasta, axis=1)

    def direccion(row):
        if row['frecuencia'] < min_obs: return '⚪ SIN DATOS'
        d = row['SCP_m3'] - row['SCP']
        if d > 3: return '📈 EMPEORA RÁPIDO'
        elif d > 1: return '📈 EMPEORA'
        elif d < -3: return '📉 MEJORA RÁPIDO'
        elif d < -1: return '📉 MEJORA'
        else: return '➡️ ESTABLE'

    def alerta(row):
        if row['frecuencia'] < min_obs: return '⚪ SIN DATOS'
        if row['SCP'] >= 70: return '🚨 CRÍTICO ACTUAL'
        if row['SCP'] >= 60: return '🟠 ATENCIÓN PRIORITARIA'
        if row['SCP_m3'] >= 70: return '🔴 SERÁ CRÍTICO'
        if row['SCP_m3'] >= 60: return '🟠 ATENCIÓN PRIORITARIA'
        if row['SCP_m3'] > row['SCP'] + 3: return '🟡 VIGILAR'
        return '✅ ESTABLE'

    score_df['direccion'] = score_df.apply(direccion, axis=1)
    score_df['alerta'] = score_df.apply(alerta, axis=1)

    return score_df.sort_values('SCP', ascending=False).reset_index(drop=True)


try:
    pred_df = predecir_scp_dashboard(dff, min_obs=5)
    pred_valid = pred_df[pred_df['nivel'] != '⚪ DATOS INSUFICIENTES'].copy()

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("🚨 Críticos actuales", (pred_df['alerta'] == '🚨 CRÍTICO ACTUAL').sum())
    a2.metric("🔴 Serán críticos (3m)", (pred_df['alerta'] == '🔴 SERÁ CRÍTICO').sum())
    a3.metric("🟠 Atención prioritaria", (pred_df['alerta'] == '🟠 ATENCIÓN PRIORITARIA').sum())
    a4.metric("📈 Empeoran", pred_valid['direccion'].str.contains('EMPEORA', na=False).sum())

    st.markdown("#### 🔮 Top 15 zonas — Proyección a 3 meses")

    top_pred = pred_valid.head(15).copy()
    tabla_pred = top_pred[[
        'cda', 'area', 'cargo_observado', 'frecuencia',
        'SCP', 'SCP_m1', 'SCP_m2', 'SCP_m3',
        'meses_critico', 'alerta'
    ]].copy()
    tabla_pred.columns = [
        'Sede', 'Área', 'Cargo', 'Frec.',
        'SCP actual', 'SCP +1m', 'SCP +2m', 'SCP +3m',
        'Meses hasta 🔴', 'Alerta'
    ]

    st.dataframe(tabla_pred, use_container_width=True, hide_index=True)

    st.markdown("#### 📈 Curva de evolución — Top 5 zonas")

    top5 = pred_valid.head(5).copy()
    top5['label'] = top5['cda'].str[:20] + " / " + top5['cargo_observado'].str[:15]

    fig = go.Figure()
    for _, row in top5.iterrows():
        fig.add_trace(go.Scatter(
            x=['Actual', '+1m', '+2m', '+3m'],
            y=[row['SCP'], row['SCP_m1'], row['SCP_m2'], row['SCP_m3']],
            mode='lines+markers', name=row['label'],
            line=dict(width=3), marker=dict(size=10)
        ))

    fig.add_hline(y=70, line_dash='dash', line_color='red',
                  annotation_text="CRÍTICO (70)", annotation_position="top right")
    fig.add_hline(y=50, line_dash='dot', line_color='orange',
                  annotation_text="ALTO (50)", annotation_position="bottom right")

    fig.update_layout(
        xaxis_title="Período", yaxis_title="SCP",
        height=500, hovermode='x unified', yaxis=dict(range=[0, 100]),
        legend=dict(orientation='h', y=-0.2)
    )
    st.plotly_chart(fig, use_container_width=True)

    seran_criticos = pred_df[pred_df['alerta'] == '🔴 SERÁ CRÍTICO']
    if len(seran_criticos) > 0:
        st.error(f"""
        ### 🚨 ALERTA TEMPRANA — {len(seran_criticos)} zona(s) alcanzarán nivel CRÍTICO en ≤3 meses

        {chr(10).join([f"- **{r['cda']} / {r['cargo_observado']}**: SCP {r['SCP']} → {r['SCP_m3']} ({r['meses_critico']})" for _, r in seran_criticos.iterrows()])}

        **Acción:** Iniciar plan de intervención **antes** de que crucen el umbral.
        """)

    st.download_button(
        "⬇️ Descargar predicción completa (CSV)",
        pred_df.to_csv(index=False, sep=';').encode('utf-8-sig'),
        "prediccion_scp.csv", "text/csv", key="dl_pred_scp"
    )

except Exception as e:
    st.error(f"Error en predicción: {e}")


# ============ SIGA ============
st.divider()
st.subheader("✅ Reforzamiento positivo — Tarjeta SIGA")
df_siga = dff[dff['tipo_tarjeta'].str.contains('SIGA', na=False)].copy()

if len(df_siga) == 0:
    st.info("No hay registros SIGA en el filtro actual.")
else:
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("✅ Total SIGA", f"{len(df_siga):,}")
    k2.metric("👥 Reportantes SIGA", df_siga['reportante'].nunique())
    k3.metric("🏭 Áreas reforzadas", df_siga['area'].nunique())
    k4.metric("👷 Cargos reforzados", df_siga['cargo_observado'].nunique())

    col_s1, col_s2 = st.columns(2)

    with col_s1:
        top_a_siga = df_siga['area'].value_counts().head(10).reset_index()
        top_a_siga.columns = ['area', 'n']
        fig = px.bar(top_a_siga, x='n', y='area', orientation='h',
                     color='n', color_continuous_scale='Greens',
                     title="Áreas con más refuerzos positivos")
        fig.update_layout(yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

    with col_s2:
        top_c_siga = df_siga['cargo_observado'].value_counts().head(10).reset_index()
        top_c_siga.columns = ['cargo', 'n']
        fig = px.bar(top_c_siga, x='n', y='cargo', orientation='h',
                     color='n', color_continuous_scale='Greens',
                     title="Cargos con más refuerzos positivos")
        fig.update_layout(yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

    actos_seguros = df_siga[df_siga['acto_seguro'].str.upper() != 'SIN DATO']['acto_seguro']
    if len(actos_seguros) > 0:
        st.markdown("#### 📝 Comportamientos seguros reforzados")
        top_seguros = actos_seguros.value_counts().head(10).reset_index()
        top_seguros.columns = ['acto_seguro', 'n']
        fig = px.bar(top_seguros, x='n', y='acto_seguro', orientation='h',
                     color='n', color_continuous_scale='Greens')
        fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=450)
        st.plotly_chart(fig, use_container_width=True)


# ============ ACADEMIA DE SEGURIDAD ============
st.divider()
st.subheader("🎓 Personal derivado a Academia de Seguridad")
st.caption("Colaboradores con tarjetas **PARE** que requieren reentrenamiento focalizado.")

df_pare = dff[
    dff['tipo_tarjeta'].str.contains('PARE', na=False) &
    (dff['colaborador_observado'] != 'SIN DATO')
].copy()

if len(df_pare) == 0:
    st.info("No hay tarjetas PARE con colaborador identificado en el filtro actual.")
else:
    academia = (df_pare.groupby(['colaborador_observado','cargo_observado','area'])
                       .agg(total_actos=('acto_inseguro','size'),
                            riesgos_altos=('riesgo', lambda x: (x=='ALTO').sum()),
                            causas_distintas=('causa','nunique'),
                            ultima_observacion=('fecha','max'))
                       .reset_index())
    academia['pct_alto'] = (academia['riesgos_altos'] / academia['total_actos'] * 100).round(1)
    academia['score_academia'] = (academia['total_actos']*0.4 +
                                  academia['pct_alto']*0.5 +
                                  academia['causas_distintas']*0.1).round(1)

    def clasificar_prioridad(row):
        if row['riesgos_altos'] >= 3 or row['pct_alto'] >= 70:
            return '🔴 ALTA — Derivar urgente'
        elif row['riesgos_altos'] >= 2 or row['pct_alto'] >= 40:
            return '🟡 MEDIA — Derivar programado'
        return '🟢 BAJA — Refuerzo en sitio'

    def tipo_academia(cola):
        causas = df_pare[df_pare['colaborador_observado'] == cola]['causa']
        causas = causas[causas != 'SIN DATO']
        cd = causas.mode()[0] if len(causas) else 'N/A'
        if 'SABE' in cd:
            return ('📚 Reentrenamiento técnico', 'Capacitación + verificación de competencia')
        elif 'QUIERE' in cd:
            return ('🧠 Coaching conductual', 'Sesión de liderazgo + refuerzo positivo')
        elif 'PUEDE' in cd:
            return ('🔧 Revisión de condiciones', 'Evaluación del puesto + ajuste EPP')
        return ('⚠️ Evaluación general', 'Entrevista + observación en campo')

    academia['Prioridad'] = academia.apply(clasificar_prioridad, axis=1)
    academia[['Tipo academia','Detalle acción']] = academia[
        'colaborador_observado'].apply(lambda c: pd.Series(tipo_academia(c)))

    causas_dom = (df_pare[df_pare['causa'] != 'SIN DATO']
                  .groupby('colaborador_observado')['causa']
                  .agg(lambda x: x.mode()[0] if len(x) else 'N/A')
                  .reset_index().rename(columns={'causa':'Causa dominante'}))
    academia = academia.merge(causas_dom, on='colaborador_observado', how='left')
    academia = academia.sort_values('score_academia', ascending=False).reset_index(drop=True)

    ka1, ka2, ka3, ka4 = st.columns(4)
    ka1.metric("🎓 Total a derivar", len(academia))
    ka2.metric("🔴 ALTA", len(academia[academia['Prioridad'].str.contains('ALTA')]))
    ka3.metric("🟡 MEDIA", len(academia[academia['Prioridad'].str.contains('MEDIA')]))
    ka4.metric("🟢 BAJA", len(academia[academia['Prioridad'].str.contains('BAJA')]))

    st.markdown("#### 📋 Lista de derivación")
    tabla_academia = academia[[
        'colaborador_observado','cargo_observado','area',
        'total_actos','riesgos_altos','pct_alto',
        'Causa dominante','Tipo academia','Prioridad','score_academia',
        'ultima_observacion']].copy()
    tabla_academia.columns = ['Colaborador','Cargo','Área',
                              'Nº Actos PARE','Riesgos ALTO','% Riesgo ALTO',
                              'Causa dominante','Tipo academia','Prioridad','Score',
                              'Última observación']
    tabla_academia['Última observación'] = pd.to_datetime(
        tabla_academia['Última observación']).dt.strftime('%d/%m/%Y')
    st.dataframe(tabla_academia, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Descargar lista de Academia (CSV)",
        tabla_academia.to_csv(index=False, sep=';').encode('utf-8-sig'),
        "academia_seguridad.csv", "text/csv", key="dl_academia"
    )


# ============ ANÁLISIS DETALLADO POR CARGO ============
if cargo_detalle != "— Ninguno —":
    st.divider()
    st.header(f"🔍 Análisis detallado: `{cargo_detalle}`")
    df_cargo = dff[dff['cargo_observado'] == cargo_detalle].copy()

    if len(df_cargo) == 0:
        st.warning("No hay observaciones para este cargo con los filtros actuales.")
    else:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📋 Observaciones", f"{len(df_cargo):,}")
        causas_cargo = df_cargo[df_cargo['causa'] != 'SIN DATO']['causa']
        k2.metric("⚠️ Causa dominante", causas_cargo.mode()[0] if len(causas_cargo) else 'N/A')
        pct_pare_cargo = df_cargo['tipo_tarjeta'].str.contains('PARE', na=False).mean()*100
        k3.metric("✅ % PARE", f"{pct_pare_cargo:.1f}%")
        k4.metric("🚦 % Riesgo ALTO", f"{df_cargo['riesgo'].eq('ALTO').mean()*100:.1f}%")

        col_a, col_b = st.columns(2)
        with col_a:
            causas_c = causas_cargo.value_counts().reset_index()
            causas_c.columns = ['causa','n']
            if len(causas_c) > 0:
                fig = px.bar(causas_c, x='n', y='causa', orientation='h',
                             color='n', color_continuous_scale='Reds',
                             title=f"Causas — {cargo_detalle}")
                fig.update_layout(yaxis={'categoryorder': 'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
        with col_b:
            tipos_c = df_cargo['tipo_tarjeta'].value_counts().reset_index()
            tipos_c.columns = ['tipo','n']
            fig = px.pie(tipos_c, names='tipo', values='n', hole=0.5,
                         title=f"Tipo tarjeta — {cargo_detalle}",
                         color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig, use_container_width=True)


# ============ TOP REPORTANTES ============
st.divider()
st.subheader("👥 Top reportantes (cultura de reporte)")

top_rep = (dff[dff['reportante'] != 'SIN DATO']['reportante']
           .value_counts().head(15).reset_index())

if len(top_rep) > 0:
    top_rep.columns = ['reportante', 'n']
    fig = px.bar(top_rep, x='reportante', y='n',
                 title="Colaboradores que más registran",
                 color='n', color_continuous_scale='Greens')
    fig.update_xaxes(tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("⚠️ No hay datos de reportantes válidos en el filtro actual.")


# ============ DÍA DE SEMANA ============
orden_dias = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
etiquetas  = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom']
dia = dff['dia_semana'].value_counts().reindex(orden_dias).fillna(0).reset_index()
dia.columns = ['dia','n']
dia['etiqueta'] = etiquetas
fig = px.bar(dia, x='etiqueta', y='n',
             title="📅 Observaciones por día de la semana",
             color='n', color_continuous_scale='Purples')
st.plotly_chart(fig, use_container_width=True)


# ============ DESCARGA ============
st.divider()
st.subheader("⬇️ Descarga de datos filtrados")
st.download_button(
    "Descargar CSV filtrado",
    dff.to_csv(index=False, sep=';').encode('utf-8-sig'),
    "observaciones_filtradas.csv", "text/csv", key="dl_filtrado"
)


# ============ INSIGHTS SST ============
st.divider()
st.subheader("🧠 Lectura SST automática")

if total > 0:
    area_top  = dff['area'].value_counts().idxmax()
    cargo_top = dff['cargo_observado'].value_counts().idxmax()
    causas_valid = dff[dff['causa'] != 'SIN DATO']['causa']
    causa_top = causas_valid.mode()[0] if len(causas_valid) else 'N/A'

    st.markdown(f"""
    ### 📌 Hallazgos clave

    - **Área crítica:** `{area_top}`.
    - **Cargo más observado:** `{cargo_top}`.
    - **Causa predominante:** `{causa_top}`.
    - **Riesgo ALTO:** {pct_alto:.1f}% → {"crítico" if pct_alto > 50 else "aceptable"}.
    - **Cultura de reporte:** {reportantes_uniq} reportantes distintos.
    - **Balance PARE/SIGA:** {pct_pare:.0f}% refuerzos positivos.
    """)


# ============ ANÁLISIS IA ============
st.divider()
st.subheader("🤖 Análisis inteligente con IA (Groq)")

if ia_disponible():
    st.success("✅ Groq configurado correctamente.")
    
    if st.button("🧠 Generar análisis IA de estos datos", type="primary"):
        with st.spinner("La IA está analizando tus datos..."):
            resumen = {
                "total_observaciones": int(len(dff)),
                "periodo": f"{dff['fecha'].min().date()} a {dff['fecha'].max().date()}",
                "pct_pare": round(pct_pare, 1),
                "pct_riesgo_alto": round(pct_alto, 1),
                "top_areas": dff['area'].value_counts().head(5).to_dict(),
                "top_cargos": dff['cargo_observado'].value_counts().head(5).to_dict(),
                "causas": dff[dff['causa'] != 'SIN DATO']['causa'].value_counts().to_dict(),
                "riesgos": dff['riesgo'].value_counts().to_dict(),
                "top_actos_inseguros": (
                    dff[dff['acto_inseguro'] != 'SIN DATO']['acto_inseguro']
                    .value_counts().head(10).to_dict()
                ),
            }

            analisis = generar_analisis_gemini(resumen)

            if analisis.startswith("Error") or analisis.startswith("⚠️"):
                st.error(analisis)
            else:
                st.success("✅ Análisis generado")
                st.markdown(analisis)
                st.download_button(
                    "⬇️ Descargar análisis IA (.md)",
                    analisis.encode('utf-8'),
                    "analisis_ia_sst.md",
                    "text/markdown",
                    key="dl_ia"
                )
else:
    st.warning(
        f"⚠️ **Falta la API Key de Groq**\n\n"
        f"Longitud detectada: `{len(GROQ_API_KEY)}` caracteres.\n"
        f"Si dice 0, el archivo no se está leyendo. Revisa la terminal."
    ) 

st.caption("Dashboard SST · PARE y SIGA · Elaborado por Seguridad y Salud en el Trabajo JLRL")  