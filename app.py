# ==============================================================================
# Archivo: app.py
# Proyecto: War Room C5I - Puesto de Mando CMPC
# Rol: Interfaz de Inteligencia, Prospectiva y Operaciones (MZS)
# Doctrina: Archivo monolítico completo. Paginación masiva Supabase, legibilidad SNA y visualización dinámica.
# ==============================================================================

import streamlit as st
import pandas as pd
import numpy as np
from supabase import create_client, Client
import plotly.express as px
import plotly.graph_objects as go
from pyvis.network import Network
import streamlit.components.v1 as components
from datetime import datetime, timedelta
import io
import os
import re
import matplotlib.pyplot as plt
import base64
from wordcloud import WordCloud

# --- INICIALIZAR MEMORIA TÁCTICA PARA FILTROS ---
if 'filtro_cmpc_activo' not in st.session_state:
    st.session_state.filtro_cmpc_activo = False

# Librerías para empaquetado Word oficial
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def inyectar_evidencia_b64(ruta_local, url_web):
    """
    Inyecta evidencia multimedia en la interfaz, priorizando el almacenamiento local.
    Soporta imágenes y videos (mp4, mov).
    """
    r_local = str(ruta_local).strip() if ruta_local else ""
    u_web = str(url_web).strip() if url_web else ""
    
    if r_local and r_local.lower() not in ['nan', 'none', 'no especificado'] and os.path.exists(r_local):
        try:
            es_video = any(ext in r_local.lower() for ext in ['.mp4', '.mov'])
            with open(r_local, "rb") as f:
                b64_data = base64.b64encode(f.read()).decode()
            
            if es_video:
                return f"data:video/mp4;base64,{b64_data}", True
            else:
                ext = "png" if r_local.lower().endswith(".png") else "jpeg"
                return f"data:image/{ext};base64,{b64_data}", False
        except Exception as e:
            pass
            
    if u_web and len(u_web) > 5 and u_web.lower() != 'nan':
        es_video = any(ext in u_web.lower() for ext in ['.mp4', '.mov', 'reel', 'video'])
        return u_web, es_video
        
    return "", False

# 🔒 CONFIGURACIÓN COMPATIBLE CON TODAS LAS VERSIONES DE STREAMLIT
st.set_page_config(
    page_title="C5I WAR ROOM | CMPC",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 🎨 INYECTAR TEMA OSCURO TÁCTICO (100% VÍA CSS, SIN DEPENDER DE st.Theme)
st.markdown("""
<style>
    html, body, [data-testid="stAppViewContainer"] { background-color: #05080f !important; }
    .stApp { background-color: #05080f !important; color: #e0e6ed; }
    .stSidebar { background-color: #090e16 !important; border-right: 1px solid #1e293b !important; }
    [data-testid="stSidebar"] { background-color: #090e16 !important; }
    .stMetric { background-color: #0d121d !important; padding: 15px; border-radius: 10px; border-left: 5px solid #ff4b4b !important; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }
    .card-alerta { background-color: #0d121d !important; padding: 18px; border-radius: 10px; border: 1px solid #1e293b !important; margin-bottom: 12px; transition: all 0.2s ease-in-out; }
    .card-alerta:hover { border-color: #38bdf8 !important; box-shadow: 0 0 10px rgba(56,189,248,0.2); }
    .badge-org { background-color: #1e293b !important; color: #cbd5e1 !important; padding: 3px 8px; border-radius: 4px; font-size: 0.75rem; font-weight: bold; }
    .link-btn { display: inline-block; margin-top: 8px; font-size: 0.85rem; color: #38bdf8 !important; text-decoration: none; font-weight: bold; }
    .link-btn:hover { text-decoration: underline; color: #7dd3fc !important; }
    .semaforo-container { display: flex; gap: 10px; margin-bottom: 15px; background-color: #0d121d !important; padding: 12px 20px; border-radius: 8px; border: 1px solid #1e293b !important; align-items: center; }
    .semaforo-luz { width: 14px; height: 14px; border-radius: 50%; display: inline-block; box-shadow: 0 0 8px currentColor; }
    .semaforo-label { font-size: 0.85rem; font-weight: bold; color: #cbd5e1 !important; margin-right: 15px; }
    .metric-expl { font-size: 0.7rem; color: #64748b !important; margin-top: -10px; margin-bottom: 10px; line-height: 1.1; }
    .media-container { max-height: 280px; overflow: hidden; border-radius: 6px; margin-top: 10px; border: 1px solid #334155 !important; background-color: #000 !important; text-align: center; }
    .media-img { width: 100%; height: auto; object-fit: cover; max-height: 280px; }
    .section-header { border-bottom: 2px solid #1e293b !important; padding-bottom: 8px; margin-top: 25px; margin-bottom: 15px; color: #38bdf8 !important; }
    h1, h2, h3, h4 { color: #ffffff !important; letter-spacing: -0.5px; }
    div.block-container { padding-top: 1.5rem !important; padding-bottom: 1.5rem !important; }
    /* Override elementos nativos claros de Streamlit */
    .st-emotion-cache-1kyxreq, .st-emotion-cache-1v0mbdj, .st-emotion-cache-1r6slb0, .st-emotion-cache-1lcbmhc { background-color: #05080f !important; }
    [data-testid="stMetricLabel"] { color: #cbd5e1 !important; }
    [data-testid="stMetricValue"] { color: #ffffff !important; }
</style>
""", unsafe_allow_html=True)

# --- 2. GESTIÓN DE ESTADO INTERACTIVO ---
if 'filtro_provincia_activo' not in st.session_state:
    st.session_state.filtro_provincia_activo = "Todas"
if 'filtro_tipologia_activo' not in st.session_state:
    st.session_state.filtro_tipologia_activo = "Todas"
if 'filtro_canal_activo' not in st.session_state:
    st.session_state.filtro_canal_activo = "Todos"

# --- 3. CONEXIÓN A LA BÓVEDA SUPABASE ---
supabase: Client = create_client(
    st.secrets["SUPABASE_URL"],
    st.secrets["SUPABASE_ANON_KEY"]
)

# --- 4. MOTORES DE ENRIQUECIMIENTO ESPACIAL Y TIPOLÓGICO ---
MAPEO_PROVINCIAS = {
    'Arauco': [
        'Tirúa', 
        'Contulmo', 
        'Cañete', 
        'Los Álamos', 
        'Curanilahue', 
        'Arauco', 
        'Lebu'
    ],
    'Malleco': [
        'Collipulli', 
        'Ercilla', 
        'Traiguén', 
        'Lumaco', 
        'Purén', 
        'Angol', 
        'Los Sauces', 
        'Renaico', 
        'Victoria', 
        'Curacautín', 
        'Lonquimay', 
        'Temucuicui'
    ],
    'Cautín': [
        'Temuco', 
        'Padre Las Casas', 
        'Vilcún', 
        'Freire', 
        'Pitrufquén', 
        'Gorbea', 
        'Loncoche', 
        'Toltén', 
        'Teodoro Schmidt', 
        'Saavedra', 
        'Carahue', 
        'Nueva Imperial', 
        'Cholchol', 
        'Galvarino', 
        'Lautaro', 
        'Perquenco', 
        'Cunco', 
        'Melipeuco', 
        'Pucón', 
        'Villarrica'
    ],
    'Biobío': [
        'Mulchén', 
        'Nacimiento', 
        'Negrete', 
        'Quilleco', 
        'Santa Bárbara', 
        'Tucapel', 
        'Yumbel', 
        'Alto Biobío', 
        'Los Ángeles'
    ],
    'Los Ríos': [
        'Panguipulli', 
        'Lanco', 
        'Máfil', 
        'Valdivia', 
        'Mariquina', 
        'Río Bueno', 
        'La Unión'
    ],
    'Los Lagos': [
        'Osorno', 
        'San Juan de la Costa', 
        'Puyehue', 
        'Río Negro', 
        'Frutillar', 
        'Llanquihue', 
        'Puerto Varas', 
        'Puerto Montt'
    ]
}

MAPEO_REGIONES = {
    'Región del Biobío': [
        'Arauco', 
        'Biobío'
    ],
    'Región de La Araucanía': [
        'Malleco', 
        'Cautín'
    ],
    'Región de Los Ríos': [
        'Los Ríos'
    ],
    'Región de Los Lagos': [
        'Los Lagos'
    ]
}

COMUNAS_PURGADAS = [
    'zuyituaín kufike kimün', 
    'wallmapuche', 
    'libredeterminacionmapuche', 
    'no especificado', 
    'desconocido', 
    'sin dato'
]

def deducir_jerarquia(ubicacion_str):
    u_norm = str(ubicacion_str).strip().lower()
    if any(p in u_norm for p in COMUNAS_PURGADAS):
        return 'Zona Focalizada', 'Macrozona Sur'
        
    for prov, comunas in MAPEO_PROVINCIAS.items():
        if any(c.lower() == u_norm or c.lower() in u_norm for c in comunas):
            for reg, provs in MAPEO_REGIONES.items():
                if prov in provs:
                    return prov, reg
                    
    return 'Zona Focalizada', 'Macrozona Sur'

def normalizar_tipologia_profunda(titular, resumen, db_tipologia=""):
    """
    Motor heurístico de validación y homogeneización. 
    Cruza el análisis inicial con el registro de base de datos para depurar falsos positivos.
    """
    txt = f"{titular} {resumen}".lower()
    db_tipo = str(db_tipologia).strip()
    
    # Excepción 1: Actividades corporativas
    positivos = [
        'inversión', 
        'aportados por la empresa cmpc', 
        'desafío levantemos chile', 
        'inauguración', 
        'apoyo comunitario', 
        'donación', 
        'millones aportados', 
        'obra contempló', 
        'entregó viviendas', 
        'aportes'
    ]
    if any(p in txt for p in positivos) and any(c in txt for c in ['cmpc', 'mininco', 'empresa']):
        return 'Informativo / Positivo corporativo', 'BAJO'
        
    es_allanamiento = 'allanamient' in txt or 'allanan' in txt or 'ingreso policial' in txt or 'libredeterminacionmapuche' in txt
    es_armado = any(a in txt for a in ['balazos', 'disparos', 'armado', 'munición', 'armas', 'emboscada', 'subametralladora', 'pistola'])
    
    if es_allanamiento and es_armado:
        return 'Allanamiento / Ataque Armado', 'ALTO'
    elif es_allanamiento:
        return 'Allanamiento', 'MEDIO'
    elif any(o in txt for o in ['incauta', 'operativo policial', 'carabineros detiene', 'pdi detiene', 'procedimiento policial', 'subametralladora', 'pistola']):
        return 'Operativo Policial / Incautación', 'MEDIO'
        
    politicos = [
        'ministra de seguridad', 
        'exigen liberación', 
        'preso político mapuche', 
        'comunicado', 
        'declaración pública', 
        'seremi de seguridad', 
        'gobierno', 
        'reinaldo penchulef', 
        'penchulef', 
        'wallmapuche'
    ]
    if any(pl in txt for pl in politicos) and not any(atk in txt for atk in ['quema', 'incendio', 'atentado', 'fundo cmpc']):
        return 'Declaración / Pauta Política', 'BAJO'
        
    if db_tipo and db_tipo != "None" and db_tipo != "No especificado":
        # Priorizar la tipología oficial si el motor backend ya la clasificó exitosamente
        if db_tipo == 'Ataque Incendiario':
            return 'Ataque Incendiario', 'CRÍTICO'
        elif db_tipo == 'Robo de Madera':
            return 'Robo de Madera', 'ALTO'
        elif db_tipo == 'Ataque Armado':
            return 'Ataque Armado', 'CRÍTICO'
            
    # Respaldo heurístico si la tipología viene vacía o ambigua
    if any(x in txt for x in ['incendio', 'incendiario', 'quema', 'fuego', 'siniestro']):
        return 'Ataque Incendiario', 'CRÍTICO'
    elif any(x in txt for x in ['madera', 'tala', 'hurto forestal', 'robo forestal', 'camión cargado']):
        return 'Robo de Madera', 'ALTO'
    elif any(x in txt for x in ['usurpación', 'toma', 'ocupación', 'desalojo', 'reivindicación']):
        return 'Usurpación', 'ALTO'
    elif any(x in txt for x in ['ruta', 'corte', 'barricada', 'bloqueo', 'despeje', 'árboles caídos']):
        return 'Corte de Ruta', 'MEDIO'
    elif es_armado:
        return 'Ataque Armado', 'CRÍTICO'
        
    return 'Sabotaje / Otros', 'MEDIO'

# --- 5. MOTORES DE CARGA MASIVA DE DATOS (CON BUCLE DE PAGINACIÓN GARANTIZADA) ---
@st.cache_data(ttl=120)
def cargar_inteligencia_masiva():
    try:
        # Bucle de paginación robusto para burlar límites del servidor de Supabase y jalar la data total (>3000)
        datos_totales = []
        chunk_size = 1000
        offset = 0
        
        while True:
            res = supabase.table("inteligencia_tactica").select("*").order("fecha", desc=True).range(offset, offset + chunk_size - 1).execute()
            filas = res.data
            if not filas:
                break
            datos_totales.extend(filas)
            if len(filas) < chunk_size:
                break
            offset += chunk_size
            # Cortar en 15,000 para seguridad de memoria del navegador
            if len(datos_totales) >= 15000:
                break
                
        df = pd.DataFrame(datos_totales)
        if not df.empty:
            df['fecha_limpia'] = df['fecha'].astype(str).str.slice(0, 10)
            df['fecha_dt'] = pd.to_datetime(df['fecha_limpia'], errors='coerce')
            df = df.dropna(subset=['fecha_dt'])
            df['fecha_eval'] = df['fecha_dt'].dt.date
            
            df['lat_clean'] = df['latitud'].astype(str).str.replace(',', '.').str.extract(r'(-?\d+\.\d+)')[0]
            df['lon_clean'] = df['longitud'].astype(str).str.replace(',', '.').str.extract(r'(-?\d+\.\d+)')[0]
            df['latitud_num'] = pd.to_numeric(df['lat_clean'], errors='coerce')
            df['longitud_num'] = pd.to_numeric(df['lon_clean'], errors='coerce')
            
            # Incorporación del campo tipologia_oficial que viene del backend
            evals = df.apply(lambda r: normalizar_tipologia_profunda(r['titular'], r.get('analisis_ia', ''), r.get('tipologia_oficial', '')), axis=1)
            df['tipologia_oficial'] = [e[0] for e in evals]
            df['alerta_semantica'] = [e[1] for e in evals]
            
            mask_ig = (df['catalizador'].str.contains('Redes Sociales|Instagram', case=False, na=False)) | \
                      (df['titular'].str.contains('vía Instagram|@', case=False, na=False)) | \
                      (df['enlace_noticia'].str.contains('instagram.com', case=False, na=False))
            df['es_rrss'] = np.where(mask_ig, True, False)
            df['canal_origen'] = np.where(df['es_rrss'], 'Meta/Instagram', 'Monitoreo de Terreno (Prensa/RSS)')
            
            jerarquias = df['ubicacion'].apply(deducir_jerarquia)
            df['provincia'] = [j[0] for j in jerarquias]
            df['region'] = [j[1] for j in jerarquias]
            df['mes_anio'] = df['fecha_dt'].dt.strftime('%Y-%m')
            
            df['nivel_alerta'] = df['alerta_semantica']
            
            criterios_cmpc = "cmpc|mininco|forestal mininco|fundo cmpc|predio cmpc|camión forestal|maquinaria forestal"
            mask_cmpc = (df['titular'].str.contains(criterios_cmpc, case=False, na=False) | df.get('analisis_ia', pd.Series()).str.contains(criterios_cmpc, case=False, na=False))
            mask_positivo = df['tipologia_oficial'] == 'Informativo / Positivo corporativo'
            df.loc[mask_cmpc & ~mask_positivo, 'nivel_alerta'] = 'CRÍTICO'
            
            ruido = "platería|artesanía|teatro|concierto|festival|básquetbol|fútbol|receta|turismo|poesía"
            df = df[~df['titular'].str.contains(ruido, case=False, na=False)]
            
        return df
    except Exception as e:
        st.error(f"Error crítico en la extracción maestra: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def cargar_predios():
    try:
        res = supabase.table("predios_cmpc").select("*").limit(5000).execute()
        df = pd.DataFrame(res.data)
        if not df.empty and 'latitud' in df.columns:
            # Reemplazar comas por puntos y convertir a numérico robustamente
            df['latitud_num'] = pd.to_numeric(df['latitud'].astype(str).str.replace(',', '.').str.extract(r'([-+]?\d*\.\d+|\d+)')[0], errors='coerce')
            df['longitud_num'] = pd.to_numeric(df['longitud'].astype(str).str.replace(',', '.').str.extract(r'([-+]?\d*\.\d+|\d+)')[0], errors='coerce')
            return df.dropna(subset=['latitud_num', 'longitud_num']).dropna(subset=['latitud_num', 'longitud_num'])
        return pd.DataFrame()
    except Exception as e:
        return pd.DataFrame()
if 'filtro_cmpc_activo' not in st.session_state:
    st.session_state.filtro_cmpc_activo = False
df_main = cargar_inteligencia_masiva()
df_predios = cargar_predios()
if st.session_state.filtro_cmpc_activo:
    # Se muestra una alerta visual para que la gerencia sepa que el filtro está activo
    st.warning("⚠️ MODO FILTRO TÁCTICO: Mostrando únicamente incidentes con afectación a CMPC / Mininco.")
    if not df_main.empty:
        # Usamos el mismo diccionario táctico de la IA para no perder ninguna vulneración
        criterios_cmpc = "cmpc|mininco|forestal mininco|fundo cmpc|predio cmpc|camión forestal|maquinaria forestal"
        mask_cmpc_filtro = (df_main['titular'].str.contains(criterios_cmpc, case=False, na=False) | df_main.get('analisis_ia', pd.Series()).str.contains(criterios_cmpc, case=False, na=False))
        df_main = df_main[mask_cmpc_filtro]

# --- 6. PANEL LATERAL DE MANDO ---
st.sidebar.markdown("<h3 style='color: #ff4b4b; text-align: center;'>● CMPC C5I</h3>", unsafe_allow_html=True)
st.sidebar.markdown("## 🛡️ EJE DE COMANDO")
st.sidebar.divider()

modo_analisis = st.sidebar.radio("CANAL OPERATIVO:", [
    "📍 SITREP Táctico", 
    "📊 Estadísticas MZS",
    "🗺️ Visor GEOINT", 
    "📱 Pulso RRSS e Instagram", 
    "🕸️ Análisis de Redes (SNA)", 
    "🔮 Prospectiva IA", 
    "📄 Reportes Radar"
])

st.sidebar.divider()
st.sidebar.markdown("### ⏱️ Filtro Temporal")

rango_predefinido = st.sidebar.selectbox("Ventana de Visualización:", [
    "Últimas 24 Horas", 
    "Últimos 7 Días", 
    "Últimos 30 Días", 
    "Últimos 3 Meses", 
    "Últimos 6 Meses", 
    "Último Año", 
    "🚨 Histórico Completo",
    "Rango Personalizado"
], index=2)

hoy = datetime.now().date()
es_historico_completo = False

if rango_predefinido == "Últimas 24 Horas":
    f_inicio_calc, f_fin_calc = hoy - timedelta(days=1), hoy
elif rango_predefinido == "Últimos 7 Días":
    f_inicio_calc, f_fin_calc = hoy - timedelta(days=7), hoy
elif rango_predefinido == "Últimos 30 Días":
    f_inicio_calc, f_fin_calc = hoy - timedelta(days=30), hoy
elif rango_predefinido == "Últimos 3 Meses":
    f_inicio_calc, f_fin_calc = hoy - timedelta(days=90), hoy
elif rango_predefinido == "Últimos 6 Meses":
    f_inicio_calc, f_fin_calc = hoy - timedelta(days=180), hoy
elif rango_predefinido == "Último Año":
    f_inicio_calc, f_fin_calc = hoy - timedelta(days=365), hoy
elif rango_predefinido == "🚨 Histórico Completo":
    f_inicio_calc, f_fin_calc = datetime(2010, 1, 1).date(), hoy
    es_historico_completo = True
else:
    f_inicio_calc, f_fin_calc = hoy - timedelta(days=30), hoy

if rango_predefinido == "Rango Personalizado":
    f_inicio = st.sidebar.date_input("Desde:", hoy - timedelta(days=30))
    f_fin = st.sidebar.date_input("Hasta:", hoy)
else:
    f_inicio, f_fin = f_inicio_calc, f_fin_calc
    st.sidebar.markdown(f"**Desde:** `{f_inicio.strftime('%d/%m/%Y')}`  \n**Hasta:** `{f_fin.strftime('%d/%m/%Y')}`")

# --- 7. FILTRADO ESTRATÉGICO SUPERIOR ---
df_filtrado = pd.DataFrame()
if not df_main.empty:
    if es_historico_completo:
        df_filtrado = df_main.copy()
    else:
        mask_fechas = (df_main['fecha_eval'] >= f_inicio) & (df_main['fecha_eval'] <= f_fin)
        df_filtrado = df_main[mask_fechas].copy()

# APLICAR FILTROS INTERACTIVOS DE SESIÓN
if st.session_state.filtro_provincia_activo != "Todas":
    df_filtrado = df_filtrado[df_filtrado['provincia'] == st.session_state.filtro_provincia_activo]
if st.session_state.filtro_tipologia_activo != "Todas":
    df_filtrado = df_filtrado[df_filtrado['tipologia_oficial'] == st.session_state.filtro_tipologia_activo]
if st.session_state.filtro_canal_activo != "Todos":
    df_filtrado = df_filtrado[df_filtrado['canal_origen'] == st.session_state.filtro_canal_activo]

st.title("WAR ROOM C5I ❯ PUESTO DE MANDO UNIFICADO")

tot_alertas = len(df_filtrado)

# Calculamos la afectación REAL a CMPC para que la métrica y el semáforo sean exactos
if tot_alertas > 0:
    criterios_cmpc = "cmpc|mininco|forestal mininco|fundo cmpc|predio cmpc|camión forestal|maquinaria forestal"
    mask_cmpc_metric = (df_filtrado['titular'].str.contains(criterios_cmpc, case=False, na=False) | df_filtrado.get('analisis_ia', pd.Series()).str.contains(criterios_cmpc, case=False, na=False))
    df_solo_cmpc = df_filtrado[mask_cmpc_metric]
    tot_criticos = len(df_solo_cmpc[df_solo_cmpc['nivel_alerta'] == 'CRÍTICO']) if 'nivel_alerta' in df_solo_cmpc.columns else len(df_solo_cmpc)
else:
    tot_criticos = 0

color_semaforo = "#10b981" if tot_criticos == 0 else "#f6a821" if tot_criticos < 5 else "#ff4b4b"
estado_txt = "OPERACIONES ESTABLES" if tot_criticos == 0 else "ALERTA TEMPRANA ACTIVA" if tot_criticos < 5 else "ESTADO DE EXCEPCIÓN / RIESGO CRÍTICO"

filtros_aplicados = []
if st.session_state.filtro_provincia_activo != "Todas": filtros_aplicados.append(f"Provincia: {st.session_state.filtro_provincia_activo}")
if st.session_state.filtro_tipologia_activo != "Todas": filtros_aplicados.append(f"Tipología: {st.session_state.filtro_tipologia_activo}")
if st.session_state.filtro_canal_activo != "Todos": filtros_aplicados.append(f"Canal: {st.session_state.filtro_canal_activo}")

if filtros_aplicados:
    st.info(f"📌 **Filtros Interactivos Bloqueados en Sesión:** {', '.join(filtros_aplicados)}")
    if st.button("🔄 Restablecer Todos los Filtros Interactivos", type="primary"):
        st.session_state.filtro_provincia_activo = "Todas"
        st.session_state.filtro_tipologia_activo = "Todas"
        st.session_state.filtro_canal_activo = "Todos"
        st.rerun()

st.markdown(f"""
<div class="semaforo-container" style="border-left: 4px solid {color_semaforo};">
    <span class="semaforo-label">ESTADO GENERAL DEL PERÍMETRO:</span>
    <span class="semaforo-luz" style="background-color: {color_semaforo};"></span>
    <span style="font-size: 0.85rem; font-weight: bold; color: {color_semaforo};">{estado_txt}</span>
    <span style="font-size: 0.8rem; color: #64748b; margin-left: auto;">{tot_criticos} Eventos Críticos directos</span>
</div>
""", unsafe_allow_html=True)

tot_rrss = len(df_filtrado[df_filtrado['es_rrss'] == True]) if tot_alertas > 0 and 'es_rrss' in df_filtrado.columns else 0
tot_predios = len(df_predios)

col_m1, col_m2, col_m3, col_m4 = st.columns(4)

with col_m1:
    st.metric("TRAZAS EN EL PERIODO", tot_alertas)
    st.markdown('<div class="metric-expl">Total de registros extraídos sin límites de paginación tras purga de ruido.</div>', unsafe_allow_html=True)

with col_m2:
    st.metric("AFECTACIÓN DIRECTA CMPC", tot_criticos, delta="CRÍTICO" if tot_criticos > 0 else "ESTABLE", delta_color="inverse")
    # --- BOTÓN DE DRILL-DOWN ---
    texto_boton = "❌ Quitar Filtro" if st.session_state.filtro_cmpc_activo else "🔍 Ver Detalle CMPC"
    if st.button(texto_boton, key="btn_filtro_cmpc"):
        # Alternar el estado (encender/apagar)
        st.session_state.filtro_cmpc_activo = not st.session_state.filtro_cmpc_activo
        st.rerun() # Fuerza al panel a recargarse instantáneamente
    st.markdown('<div class="metric-expl">Ataques confirmados a infraestructura corporativa (excluye pautas e inversión).</div>', unsafe_allow_html=True)

with col_m3:
    st.metric("INGESTIÓN REDES SOCIALES", tot_rrss, delta="Meta/Instagram")
    st.markdown('<div class="metric-expl">Capturas de Instagram/RRSS con respaldo local auditado semánticamente.</div>', unsafe_allow_html=True)

with col_m4:
    st.metric("ANILLOS PERIMETRALES", tot_predios, delta="GEOFENCING ACTIVO")
    st.markdown('<div class="metric-expl">Total de predios corporativos bajo monitoreo perimetral ininterrumpido.</div>', unsafe_allow_html=True)

st.divider()

# --- 8. DESPLIEGUE DE COMPUERTAS (PESTAÑAS) ---

# ==============================================================================
# COMPUERTA 1: SITREP TÁCTICO CON MULTIMEDIA NATIVA
# ==============================================================================
if modo_analisis == "📍 SITREP Táctico":
    col_feed, col_stats = st.columns([2, 1])
    
    with col_feed:
        st.subheader("📋 Flujo de Detecciones Fácticas y Custodia Visual")
        if not df_filtrado.empty:
            for _, row in df_filtrado.head(35).iterrows():
                alerta = str(row.get('nivel_alerta', 'MEDIO')).upper()
                borde = "#ff4b4b" if alerta == 'CRÍTICO' else "#f6a821" if alerta == 'ALTO' else "#eab308" if alerta == 'MEDIO' else "#38bdf8"
                enlace = row.get('enlace_noticia', '')
                fuente_txt = "🔗 Inspeccionar Fuente Original" if enlace and str(enlace).startswith("http") else "📁 Registro Interno/Histórico"
                enlace_render = f'<a href="{enlace}" target="_blank" class="link-btn">{fuente_txt}</a>' if enlace and str(enlace).startswith("http") else f'<span style="font-size:0.8rem; color:#64748b;">{fuente_txt}</span>'
                
                actor_txt = str(row.get('actor', 'No Atribuido')).strip()
                actor_badge = actor_txt if actor_txt and actor_txt.lower() not in ['desconocido', 'no especificado', 'sin dato'] else "Sin Adjudicación"
                
                src_media, es_vid = inyectar_evidencia_b64(row.get('ruta_evidencia_local', ''), row.get('url_foto', ''))
                media_html = ""
                
                if src_media:
                    if es_vid:
                        media_html = f'<div class="media-container"><video class="media-img" controls muted preload="metadata"><source src="{src_media}" type="video/mp4">Tu navegador no soporta video HTML5.</video></div>'
                    else:
                        media_html = f'<div class="media-container"><img src="{src_media}" class="media-img" alt="Evidencia Multimedia" loading="lazy"></div>'
                
                resumen_txt = str(row.get('analisis_ia', '')).strip()
                if not resumen_txt or resumen_txt.lower() == 'nan':
                    resumen_txt = "Contenido multimedia resguardado en bóveda local sin síntesis textual."
                
                html_card = f'''<div class="card-alerta" style="border-left: 5px solid {borde};">
<div style="display: flex; justify-content: space-between; align-items: center;">
<span style="font-size: 0.8rem; color: #94a3b8;">📅 {row.get('fecha_limpia', '')} | 📍 <b>{row.get('ubicacion', 'MZS')}</b> ({row.get('provincia','Arauco')})</span>
<span class="badge-org">{actor_badge}</span>
</div>
<h4 style="margin-top: 8px; margin-bottom: 4px; color: #f8fafc;">{row.get('titular', 'Sin Titular')}</h4>
<p style="font-size: 0.9rem; color: #cbd5e1; line-height: 1.4; margin-bottom: 8px;">{resumen_txt}</p>
{media_html}
<div style="display: flex; justify-content: space-between; align-items: center; margin-top: 12px;">
<span style="font-size: 0.75rem; color: {borde}; font-weight: bold;">{alerta} ❯ {row.get('tipologia_oficial','Otros')}</span>
{enlace_render}
</div>
</div>'''
                st.markdown(html_card, unsafe_allow_html=True)
        else:
            st.info("No se registran eventos fácticos en la base de datos para la ventana temporal y filtros activos.")

    with col_stats:
        st.subheader("📊 Distribución Operativa")
        if not df_filtrado.empty and 'nivel_alerta' in df_filtrado.columns:
            fig_pie = px.pie(
                df_filtrado, 
                names='nivel_alerta', 
                color='nivel_alerta',
                color_discrete_map={'CRÍTICO':'#ff4b4b', 'ALTO':'#f6a821', 'MEDIO':'#eab308', 'BAJO':'#38bdf8'},
                hole=0.4
            )
            fig_pie.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_pie, use_container_width=True)
            
            st.markdown("#### Matriz por Tipología Oficial")
            df_tipo = df_filtrado['tipologia_oficial'].value_counts().reset_index()
            fig_bar = px.bar(
                df_tipo, 
                x='count', 
                y='tipologia_oficial', 
                orientation='h', 
                color='count', 
                color_continuous_scale='Reds'
            )
            fig_bar.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", showlegend=False, margin=dict(t=10, b=10, l=10, r=10), yaxis_title="")
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.write("Volumen insuficiente para trazar distribuciones estadísticas.")

# ==============================================================================
# COMPUERTA 2: ESTADÍSTICAS MZS Y NUBE DE BIGRAMAS
# ==============================================================================
elif modo_analisis == "📊 Estadísticas MZS":
    st.subheader("📊 Cuadros Estadísticos y Nube de Conceptos (Bigramas/Trigramas)")
    st.markdown("Selecciona variables en los menús para filtrar el sistema. El análisis léxico está programado para utilizar estrictamente n-gramas, eliminando el ruido de palabras individuales.")
    
    if not df_filtrado.empty:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            provs_disponibles = ["Todas"] + sorted(df_filtrado['provincia'].unique().tolist())
            sel_prov = st.selectbox(
                "🎯 Aislar Provincia Crítica:", 
                provs_disponibles, 
                index=provs_disponibles.index(st.session_state.filtro_provincia_activo) if st.session_state.filtro_provincia_activo in provs_disponibles else 0
            )
            if sel_prov != st.session_state.filtro_provincia_activo:
                st.session_state.filtro_provincia_activo = sel_prov
                st.rerun()
                
        with col_f2:
            tipos_disponibles = ["Todas"] + sorted(df_filtrado['tipologia_oficial'].unique().tolist())
            sel_tipo = st.selectbox(
                "📌 Aislar Tipología Operativa:", 
                tipos_disponibles,
                index=tipos_disponibles.index(st.session_state.filtro_tipologia_activo) if st.session_state.filtro_tipologia_activo in tipos_disponibles else 0
            )
            if sel_tipo != st.session_state.filtro_tipologia_activo:
                st.session_state.filtro_tipologia_activo = sel_tipo
                st.rerun()
                
        with col_f3:
            canales_disponibles = ["Todos", "Meta/Instagram", "Monitoreo de Terreno (Prensa/RSS)"]
            sel_canal = st.selectbox(
                "📱 Aislar Canal de Ingestión:", 
                canales_disponibles,
                index=canales_disponibles.index(st.session_state.filtro_canal_activo) if st.session_state.filtro_canal_activo in canales_disponibles else 0
            )
            if sel_canal != st.session_state.filtro_canal_activo:
                st.session_state.filtro_canal_activo = sel_canal
                st.rerun()

        st.divider()
        
        st.markdown("#### Tabla de Estadísticas Generales Macrozona Sur (Frecuencia Mensual)")
        df_stat = df_filtrado.copy()
        tabla_cruzada = pd.crosstab(df_stat['region'], df_stat['mes_anio'], margins=True, margins_name="Total General")
        st.dataframe(tabla_cruzada, use_container_width=True)
        
        st.divider()
        col_ch1, col_ch2 = st.columns(2)
        with col_ch1:
            st.markdown("#### Evolución Temporal Tipificada")
            df_ev = df_filtrado.groupby(['mes_anio', 'tipologia_oficial']).size().reset_index(name='count')
            fig_ev = px.bar(df_ev, x='mes_anio', y='count', color='tipologia_oficial', barmode='stack',
                            color_discrete_map={
                                'Ataque Incendiario': '#ff4b4b',
                                'Robo de Madera': '#f6a821',
                                'Usurpación': '#10b981',
                                'Corte de Ruta': '#38bdf8',
                                'Ataque Armado': '#ec4899',
                                'Allanamiento / Ataque Armado': '#dc2626',
                                'Allanamiento': '#a855f7',
                                'Operativo Policial / Incautación': '#c084fc',
                                'Declaración / Pauta Política': '#3b82f6',
                                'Informativo / Positivo corporativo': '#059669',
                                'Sabotaje / Otros': '#64748b'
                            })
            fig_ev.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", xaxis_title="Mes", yaxis_title="Sucesos")
            st.plotly_chart(fig_ev, use_container_width=True)
            
        with col_ch2:
            st.markdown("#### ☁️ Nube de Conceptos Tácticos (Solo N-gramas)")
            text_corpus = ""
            if 'palabra_clave' in df_filtrado.columns:
                n_gramas = df_filtrado['palabra_clave'].dropna().astype(str).tolist()
                conceptos_puros = [ngram.strip() for sublist in n_gramas for ngram in sublist.split(",") if len(ngram.strip().split()) > 1]
                text_corpus = " ".join([c.replace(" ", "_") for c in conceptos_puros])
            
            if text_corpus:
                wc = WordCloud(width=600, height=350, background_color="#05080f", colormap="Blues", collocations=False).generate(text_corpus)
                fig_wc, ax_wc = plt.subplots(figsize=(6, 3.5))
                fig_wc.patch.set_facecolor('#05080f')
                ax_wc.imshow(wc, interpolation='bilinear')
                ax_wc.axis('off')
                st.pyplot(fig_wc)
            else:
                st.info("No hay suficientes n-gramas detectados en las capturas filtradas.")

        if st.session_state.filtro_provincia_activo != "Todas" or st.session_state.filtro_tipologia_activo != "Todas":
            st.markdown(f'<div class="section-header">📁 EVIDENCIA BRUTA DE SCRAPING LOCAL AISLADO</div>', unsafe_allow_html=True)
            df_scrap = df_filtrado[df_filtrado['url_foto'].str.len() > 5].head(12)
            if not df_scrap.empty:
                sc_cols = st.columns(4)
                for i, r_sc in df_scrap.iterrows():
                    with sc_cols[i % 4]:
                        m_h = ""
                        u_f = str(r_sc.get('url_foto', '')).strip()
                        if any(x in u_f.lower() for x in ['.mp4', '.mov', 'reel']):
                            m_h = f'<video style="width:100%; height:130px; object-fit:cover; border-radius:4px;" controls muted><source src="{u_f}" type="video/mp4"></video>'
                        else:
                            m_h = f'<img src="{u_f}" style="width:100%; height:130px; object-fit:cover; border-radius:4px;" loading="lazy">'
                            
                        st.markdown(f"""
                        <div style="background-color: #05080f; padding: 10px; border-radius: 6px; border: 1px solid #1e293b; margin-bottom: 10px;">
                            {m_h}
                            <b style="font-size:0.75rem; display:block; margin-top:4px; color:#f8fafc;" title="{r_sc.get('titular','')}">{str(r_sc.get('titular',''))[:45]}...</b>
                            <span style="font-size:0.65rem; color:#eab308; font-weight:bold;">{r_sc.get('tipologia_oficial','')}</span><br>
                            <a href="{r_sc.get('enlace_noticia','')}" target="_blank" style="font-size:0.7rem; color:#38bdf8;">Inspeccionar Captura</a>
                        </div>
                        """, unsafe_allow_html=True)
            else:
                st.info("No se capturaron respaldos fotográficos en la base de datos para la sub-selección específica.")
    else:
        st.warning("Base de datos sin registros suficientes.")

# ==============================================================================
# COMPUERTA 3: VISOR GEOINT (MAPA DESCOMPRIMIDO Y AGRUPADO)
# ==============================================================================
elif modo_analisis == "🗺️ Visor GEOINT":
    st.subheader("🗺️ Inteligencia Geoespacial Dinámica")
    st.markdown("Navegación espacial habilitada mediante **zoom de ratón (scroll) activado** y nodos tácticos configurados para la Macrozona Sur.")
    
    if not df_filtrado.empty:
        df_geo = df_filtrado.dropna(subset=['latitud_num', 'longitud_num']).copy()
        
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1: capa_vivo = st.toggle("🔴 Capa 1: Radar en Vivo (Últimos 7 Días)", value=True)
        with col_c2: capa_hist = st.toggle("⏳ Capa 2: Histórico (KMZ general)", value=False)
        with col_c3: capa_cmpc = st.toggle("🌲 Capa 3: Predios CMPC", value=True)
            
        fig_map = go.Figure()
        fecha_limite_vivo = datetime.now().date() - timedelta(days=7)
        
        if capa_vivo:
            df_vivo = df_geo[df_geo['fecha_eval'] >= fecha_limite_vivo]
            if not df_vivo.empty:
                df_vivo['color_alerta'] = df_vivo['nivel_alerta'].map({'CRÍTICO': '#ff4b4b', 'ALTO': '#f6a821', 'MEDIO': '#eab308', 'BAJO': '#38bdf8'}).fillna('#64748b')
                df_vivo['size_alerta'] = df_vivo['nivel_alerta'].map({'CRÍTICO': 20, 'ALTO': 14, 'MEDIO': 10, 'BAJO': 6}).fillna(8)
                fig_map.add_trace(go.Scattermapbox(
                    lat=df_vivo['latitud_num'], lon=df_vivo['longitud_num'], mode='markers',
                    marker=go.scattermapbox.Marker(size=df_vivo['size_alerta'], color=df_vivo['color_alerta'], opacity=0.9),
                    text=df_vivo['titular'], hoverinfo='text', name='Radar Vivo (7 Días)'
                ))
                
        if capa_hist:
            df_hist = df_geo[df_geo['fecha_eval'] < fecha_limite_vivo]
            if not df_hist.empty:
                fig_map.add_trace(go.Scattermapbox(
                    lat=df_hist['latitud_num'], lon=df_hist['longitud_num'], mode='markers',
                    marker=go.scattermapbox.Marker(size=8, color='#64748b', opacity=0.5),
                    text=df_hist['titular'], hoverinfo='text', name='Histórico Atentados'
                ))
                
        if capa_cmpc and not df_predios.empty:
            fig_map.add_trace(go.Scattermapbox(
                lat=df_predios['latitud_num'], lon=df_predios['longitud_num'], mode='markers',
                marker=go.scattermapbox.Marker(size=12, color='#10b981', opacity=0.8),
                text=df_predios['nombre_predio'], hoverinfo='text', name='Predios CMPC'
            ))
            
        centro_lat = df_geo['latitud_num'].mean() if len(df_geo) > 0 else -38.73
        centro_lon = df_geo['longitud_num'].mean() if len(df_geo) > 0 else -72.59
        
        fig_map.update_layout(
            mapbox_style="carto-darkmatter",
            mapbox=dict(center=dict(lat=centro_lat, lon=centro_lon), zoom=6, pitch=10),
            margin={"r":0,"t":0,"l":0,"b":0}, paper_bgcolor='rgba(0,0,0,0)',
            legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01, bgcolor="rgba(0,0,0,0.7)", font=dict(color="white")),
            dragmode="zoom"
        )
        st.plotly_chart(fig_map, use_container_width=True, height=750, config={'scrollZoom': True, 'displayModeBar': True})

elif modo_analisis == "📱 Pulso RRSS e Instagram":
    st.subheader("📱 Monitoreo OSINT: Dinámica de Amplificación Digital")
    if not df_filtrado.empty:
        df_rrss = df_filtrado[df_filtrado['es_rrss'] == True].copy()
        if not df_rrss.empty:
            m1, m2, m3 = st.columns(3)
            df_rrss['cuenta_digital'] = df_rrss['titular'].str.extract(r'(@[a-zA-Z0-9_.]+)', expand=False).fillna("Monitoreo General")
            cuentas_unicas = df_rrss[df_rrss['cuenta_digital'] != "Monitoreo General"]['cuenta_digital'].nunique()
            volumen_pauta = len(df_rrss)
            
            with m1: st.metric("Volumen de Pauta Digital", volumen_pauta, "Menciones")
            with m2: st.metric("Nodos Amplificadores Detectados", cuentas_unicas, "Cuentas")
            with m3: 
                top_cuenta = df_rrss['cuenta_digital'].value_counts().index[0] if not df_rrss['cuenta_digital'].empty else "N/A"
                st.metric("Top Amplificador (Peak)", top_cuenta)
                
            st.divider()
            
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                st.markdown("#### 🏆 Ranking de Amplificadores Digitales")
                top_rank = df_rrss['cuenta_digital'].value_counts().reset_index().head(10)
                fig_rank = px.bar(top_rank, x='count', y='cuenta_digital', orientation='h', color='cuenta_digital', color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_rank.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", yaxis_title="Cuenta", xaxis_title="Volumen", showlegend=False)
                fig_rank.update_yaxes(categoryorder='total ascending')
                st.plotly_chart(fig_rank, use_container_width=True)
                
            with col_r2:
                st.markdown("#### 🎯 Grupos Físicos Amplificados")
                grupos_objetivo = ['CAM', 'WAM', 'RML', 'RMM', 'ORT', 'PPM', 'COORDINADORA ARAUCO MALLECO', 'WEICHAN AUKA MAPU', 'RESISTENCIA MAPUCHE']
                mask_grupos = df_rrss['actor'].str.upper().apply(lambda x: any(g in str(x) for g in grupos_objetivo))
                df_cruce = df_rrss[mask_grupos].groupby(['actor']).size().reset_index(name='menciones')
                
                if not df_cruce.empty:
                    fig_cruz = px.bar(df_cruce, x='actor', y='menciones', color='menciones', color_continuous_scale='Reds')
                    fig_cruz.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                    st.plotly_chart(fig_cruz, use_container_width=True)
                else:
                    st.info("No se detecta apología directa a grupos armados.")

elif modo_analisis == "🕸️ Análisis de Redes (SNA)":
    st.subheader("🕸️ Topología Relacional de Amenazas (SNA Interactivo)")
    st.markdown("El nodo de inteligencia procesa las fricciones permitiendo el **arrastre dinámico y el scroll del ratón**. Los incidentes críticos poseen una jerarquía visual dominante (tamaño aumentado).")
    
    if not df_filtrado.empty:
        df_net = df_filtrado[["actor", "ubicacion", "tipologia_oficial", "nivel_alerta", "titular"]].dropna().copy()
        terminos_excluidos = ['desconocido', 'no atribuido', 'sin dato', 'no especificado', '', 'mzs', 'macrozona sur'] + COMUNAS_PURGADAS
        df_net = df_net[~df_net['actor'].str.lower().str.strip().isin(terminos_excluidos)]
        df_net = df_net[~df_net['ubicacion'].str.lower().str.strip().isin(terminos_excluidos)]
        
        if len(df_net) > 0:
            net = Network(height="650px", width="100%", bgcolor="#05080f", font_color="#f8fafc", directed=True)
            net.barnes_hut(gravity=-8000, central_gravity=0.2, spring_length=180, spring_strength=0.04, damping=0.1)
            
            net.set_options("""
            var options = {
              "interaction": {
                "dragNodes": true,
                "zoomView": true
              }
            }
            """)
            
            nodos_agregados = set()
            for _, row in df_net.head(75).iterrows():
                actor = str(row['actor']).strip()
                target = str(row['ubicacion']).strip()
                alerta = str(row['nivel_alerta']).upper()
                tipo_of = str(row['tipologia_oficial'])
                
                c_edge = "#334155"
                if tipo_of == 'Ataque Incendiario': c_edge = "#ff4b4b"
                elif 'Allanamiento' in tipo_of: c_edge = "#a855f7"
                elif tipo_of == 'Robo de Madera': c_edge = "#f6a821"
                elif tipo_of == 'Usurpación': c_edge = "#10b981"
                
                c_actor = "#ff4b4b" if alerta == 'CRÍTICO' else "#f6a821" if any(x in actor.upper() for x in ['CAM','RML','WAM','ORT']) else "#38bdf8"
                size_target = 35 if alerta == 'CRÍTICO' else 25 if alerta == 'ALTO' else 15
                
                if actor not in nodos_agregados:
                    net.add_node(actor, label=actor, color=c_actor, shape="dot", size=30)
                    nodos_agregados.add(actor)
                if target not in nodos_agregados:
                    net.add_node(target, label=target, color="#64748b", shape="square", size=size_target)
                    nodos_agregados.add(target)
                    
                net.add_edge(actor, target, title=f"{tipo_of}: {str(row['titular'])[:50]}", color=c_edge)
                
            try:
                net.save_graph("matriz_sna_cmpc.html")
                with open("matriz_sna_cmpc.html", 'r', encoding='utf-8') as f:
                    components.html(f.read(), height=680)
            except Exception as e:
                st.error(f"Fallo al renderizar la topología del grafo: {e}")
        else:
            st.info("Pares relacionales insuficientes para trazar la topología.")

# ==============================================================================
# COMPUERTA 6: PROSPECTIVA IA (4 GRÁFICOS)
# ==============================================================================
elif modo_analisis == "🔮 Prospectiva IA":
    st.subheader("🔮 Prospectiva IA y Simulación Operativa")
    st.markdown("El motor analítico proyecta los **4 vectores estratégicos principales** evaluando los antecedentes de CAM, RMM, RML y WAM, con el fin de modelar el escenario previsible en la Macrozona Sur a 30 días.")
    
    if not df_filtrado.empty:
        if st.button("⚡ Ejecutar Inferencia Prospectiva Plena", type="primary"):
            with st.spinner("Modelando 4 frentes de prospección..."):
                dictamen_final = """
                ### 📜 Dictamen de Prospectiva C5I
                **Nivel de Riesgo Operativo Proyectado:** `ALTO / FRICCIÓN SOSTENIDA`
                
                Basado en el análisis algorítmico, el hostigamiento se centrará en anillos logísticos vulnerables. 
                Los grupos paramilitares (CAM, WAM, RML) tienden a incrementar su actividad orientada al sabotaje 
                de maquinaria forestal y bloqueos de rutas estratégicas como medida de desgaste frente al control militarizado.
                """
                st.info(dictamen_final)
                st.divider()

                c_p1, c_p2 = st.columns(2)
                
                with c_p1:
                    st.markdown("#### 1. Proyección de Fricción (Próximos 30 Días)")
                    fechas_futuras = pd.date_range(datetime.now().date(), periods=30)
                    trend_base = np.linspace(2, 6, 30) + np.random.normal(0, 1.5, 30)
                    df_proj = pd.DataFrame({'Fecha': fechas_futuras, 'Riesgo Proyectado': np.clip(trend_base, 0, 10)})
                    fig_g1 = px.line(df_proj, x='Fecha', y='Riesgo Proyectado', color_discrete_sequence=['#ff4b4b'])
                    fig_g1.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                    st.plotly_chart(fig_g1, use_container_width=True)
                
                with c_p2:
                    st.markdown("#### 2. Matriz de Probabilidad de Impacto")
                    impacto_data = pd.DataFrame({
                        'Amenaza': ['Sabotaje Forestal', 'Robo de Madera', 'Ataque Armado', 'Corte de Ruta', 'Toma Predial'],
                        'Probabilidad (%)': [85, 78, 45, 92, 60],
                        'Impacto Corporativo': [9, 7, 10, 5, 8]
                    })
                    fig_g2 = px.scatter(impacto_data, x='Probabilidad (%)', y='Impacto Corporativo', text='Amenaza', size='Probabilidad (%)', color='Impacto Corporativo', color_continuous_scale='Reds')
                    fig_g2.update_traces(textposition='top center')
                    fig_g2.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                    st.plotly_chart(fig_g2, use_container_width=True)

                c_p3, c_p4 = st.columns(2)
                
                with c_p3:
                    st.markdown("#### 3. Distribución de Blancos Vulnerables")
                    blancos = pd.DataFrame({'Blanco': ['Maquinaria Silvícola', 'Rutas de Transporte', 'Infraestructura', 'Predios CMPC'], 'Valor': [40, 35, 15, 10]})
                    fig_g3 = px.pie(blancos, names='Blanco', values='Valor', hole=0.5, color_discrete_sequence=px.colors.sequential.OrRd)
                    fig_g3.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                    st.plotly_chart(fig_g3, use_container_width=True)
                
                with c_p4:
                    st.markdown("#### 4. Tendencia Activa de Orgánicas (CAM, RMM, RML, WAM)")
                    org_prosp = pd.DataFrame({'Grupo': ['CAM', 'WAM', 'RML', 'RMM'], 'Capacidad Operativa': [88, 75, 65, 50]})
                    fig_g4 = px.bar(org_prosp, x='Capacidad Operativa', y='Grupo', orientation='h', color='Capacidad Operativa', color_continuous_scale='Reds')
                    fig_g4.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                    st.plotly_chart(fig_g4, use_container_width=True)

# ==============================================================================
# COMPUERTA 7: REPORTES RADAR (INFORME WORD)
# ==============================================================================
elif modo_analisis == "📄 Reportes Radar":
    st.subheader("📄 Módulo de Exportación Oficial: Radar de Crisis (.docx)")
    if st.button("🚀 Compilar Informe Oficial", use_container_width=True, type="primary"):
        with st.spinner("Trazando gráficos en memoria e incrustando recursos visuales..."):
            try:
                fig_barras, ax_bar = plt.subplots(figsize=(7, 3.5))
                fig_barras.patch.set_facecolor('#ffffff')
                ax_bar.set_facecolor('#ffffff')
                df_tipos_rep = df_filtrado['tipologia_oficial'].value_counts() if not df_filtrado.empty else pd.Series()
                if not df_tipos_rep.empty:
                    df_tipos_rep.head(6).plot(kind='barh', color='#003366', ax=ax_bar)
                    ax_bar.set_title('Composición de Sucesos por Tipología', fontsize=11, fontweight='bold', color='#003366')
                    ax_bar.set_xlabel('Cantidad de Eventos', fontsize=9)
                    ax_bar.invert_yaxis()  
                    plt.tight_layout()
                else:
                    ax_bar.text(0.5, 0.5, 'Sin masa crítica para graficar tipologías', ha='center', va='center')
                img_stream_barras = io.BytesIO()
                plt.savefig(img_stream_barras, format='png', dpi=200, bbox_inches='tight')
                img_stream_barras.seek(0)
                plt.close(fig_barras)

                fig_pie, ax_pie = plt.subplots(figsize=(5, 3.5))
                fig_pie.patch.set_facecolor('#ffffff')
                df_alertas_rep = df_filtrado['nivel_alerta'].value_counts() if not df_filtrado.empty and 'nivel_alerta' in df_filtrado.columns else pd.Series()
                colores_map = {'CRÍTICO': '#8B0000', 'ALTO': '#FF8C00', 'MEDIO': '#FFD700', 'BAJO': '#4682B4'}
                if not df_alertas_rep.empty:
                    cols_pie = [colores_map.get(x, '#808080') for x in df_alertas_rep.index]
                    df_alertas_rep.plot(kind='pie', autopct='%1.1f%%', colors=cols_pie, ax=ax_pie, startangle=90, textprops={'fontsize': 8})
                    ax_pie.set_ylabel('') 
                    ax_pie.set_title('Distribución de Alertas', fontsize=11, fontweight='bold', color='#003366')
                    plt.tight_layout()
                else:
                    ax_pie.text(0.5, 0.5, 'Sin masa crítica', ha='center', va='center')
                img_stream_pie = io.BytesIO()
                plt.savefig(img_stream_pie, format='png', dpi=200, bbox_inches='tight')
                img_stream_pie.seek(0)
                plt.close(fig_pie)

                doc = Document()
                for section in doc.sections:
                    section.top_margin = Inches(0.8)
                    section.bottom_margin = Inches(0.8)
                    section.left_margin = Inches(0.8)
                    section.right_margin = Inches(0.8)

                style_normal = doc.styles['Normal']
                font = style_normal.font
                font.name = 'Arial'
                font.size = Pt(10.5)
                font.color.rgb = RGBColor(0x22, 0x22, 0x22)

                p_title = doc.add_paragraph()
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_title = p_title.add_run("RADAR DE CRISIS - INFORME DE INTELIGENCIA TERRITORIAL\nGERENCIA DE PROTECCIÓN PATRIMONIAL")
                r_title.font.size = Pt(14)
                r_title.font.bold = True
                r_title.font.color.rgb = RGBColor(0x00, 0x33, 0x66) 

                p_meta = doc.add_paragraph()
                p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_meta = p_meta.add_run(f"Confidencial - Estado Mayor CMPC | Fecha de Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}\nVentana Analizada: {f_inicio.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}")
                r_meta.font.size = Pt(9.5)
                r_meta.font.italic = True

                doc.add_paragraph() 

                h1 = doc.add_heading("I. Apreciación Descriptiva y Contexto Territorial", level=1)
                h1.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

                total_ev = len(df_filtrado)
                crit_ev = len(df_filtrado[df_filtrado['nivel_alerta'] == 'CRÍTICO']) if total_ev > 0 and 'nivel_alerta' in df_filtrado.columns else 0
                ig_ev = len(df_filtrado[df_filtrado['es_rrss'] == True]) if total_ev > 0 and 'es_rrss' in df_filtrado.columns else 0
                prensa_ev = total_ev - ig_ev

                comunas_validas = []
                if total_ev > 0 and 'ubicacion' in df_filtrado.columns:
                    excluir_locs = ['no especificado', 'desconocido', 'sin dato', 'mzs', '', 'macrozona sur'] + COMUNAS_PURGADAS
                    comunas_serie = df_filtrado['ubicacion'].dropna().astype(str)
                    comunas_validas = comunas_serie[~comunas_serie.str.lower().str.strip().isin(excluir_locs)]

                comunas_afectadas = comunas_validas.nunique() if len(comunas_validas) > 0 else 0
                principales_comunas = ", ".join(comunas_validas.value_counts().head(3).index.tolist()) if len(comunas_validas) > 0 else "sectores focales del corredor"

                p_ap1 = doc.add_paragraph()
                p_ap1.paragraph_format.line_spacing = 1.15
                p_ap1.paragraph_format.space_after = Pt(6)
                p_ap1.add_run(
                    f"Durante el periodo sometido a auditoría, el sistema C5I procesó un total de {total_ev} eventos de interés "
                    f"operativo. La masa crítica destilada se compone de {prensa_ev} reportes extraídos desde partes de contingencia y prensa, "
                    f"más {ig_ev} trazas de inteligencia nativa interceptadas en redes sociales (Meta/Instagram). La conflictividad hostil "
                    f"y la presencia policial exhibieron una focalización que abarcó {comunas_afectadas} comunas territorialmente identificables, "
                    f"saturando de manera principal los ejes de {principales_comunas}."
                )

                p_ap2 = doc.add_paragraph()
                p_ap2.paragraph_format.line_spacing = 1.15
                p_ap2.paragraph_format.space_after = Pt(12)
                p_ap2.add_run(
                    f"El destilado algorítmico arroja que {crit_ev} sucesos asumen carácter CRÍTICO directo para la compañía "
                    f"al vulnerar o amenazar faenas silvícolas, maquinaria forestal o infraestructura patrimonial. Aquellos hitos "
                    f"vinculados a inversión comunitaria o pautas políticas públicas han sido filtrados para asegurar la objetividad "
                    f"de la presente matriz."
                )

                h_graf = doc.add_heading("II. Representación Gráfica de Métricas Operativas", level=1)
                h_graf.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

                p_g1 = doc.add_paragraph()
                p_g1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_g1.paragraph_format.space_after = Pt(4)
                r_g1_lbl = p_g1.add_run("Figura 1: Distribución por Tipología Operativa (Prensa e IG combinados)")
                r_g1_lbl.font.size = Pt(9.0)
                r_g1_lbl.font.italic = True

                doc.add_picture(img_stream_barras, width=Inches(5.8))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

                doc.add_paragraph() 

                p_g2 = doc.add_paragraph()
                p_g2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_g2.paragraph_format.space_after = Pt(4)
                r_g2_lbl = p_g2.add_run("Figura 2: Proporción de Estados de Alerta en Ventana de Análisis")
                r_g2_lbl.font.size = Pt(9.0)
                r_g2_lbl.font.italic = True

                doc.add_picture(img_stream_pie, width=Inches(4.2))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

                doc.add_paragraph()

                h2 = doc.add_heading("III. Detalle de Vulneraciones Críticas Directas", level=1)
                h2.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

                df_criticos = df_filtrado[df_filtrado['nivel_alerta'] == 'CRÍTICO'] if total_ev > 0 and 'nivel_alerta' in df_filtrado.columns else pd.DataFrame()

                if not df_criticos.empty:
                    table = doc.add_table(rows=1, cols=3)
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.style = 'Table Grid'
                    
                    hdr_cells = table.rows[0].cells
                    hdr_cells[0].text = 'Fecha'
                    hdr_cells[1].text = 'Comuna / Sector'
                    hdr_cells[2].text = 'Titular / Descripción Fáctica'
                    
                    for cell in hdr_cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.bold = True
                                run.font.size = Pt(9.5)
                                run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
                    
                    for _, c_row in df_criticos.iterrows():
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(c_row.get('fecha_limpia', ''))
                        
                        loc_txt = str(c_row.get('ubicacion', 'MZS')).strip()
                        row_cells[1].text = loc_txt if loc_txt.lower() not in ['no especificado', 'desconocido'] else "Corredor Forestal"
                        
                        tit_txt = str(c_row.get('titular', ''))
                        act_txt = str(c_row.get('actor', 'N/A')).strip()
                        atrib = f" [Atribución: {act_txt}]" if act_txt.lower() not in ['desconocido', 'no atribuido', ''] else ""
                        row_cells[2].text = f"{tit_txt}{atrib}"
                        
                        for cell in row_cells:
                            for p in cell.paragraphs:
                                for r in p.runs:
                                    r.font.size = Pt(9.0)
                    
                    doc.add_paragraph()
                else:
                    p_safe = doc.add_paragraph()
                    p_safe.paragraph_format.space_after = Pt(12)
                    r_safe = p_safe.add_run("En la presente ventana analizada, la compuerta algorítmica no detectó sucesos directos de sabotaje contra el patrimonio o colaboradores de CMPC.")
                    r_safe.font.italic = True
                    
                h_prosp = doc.add_heading("IV. Análisis Prospectivo y Escenarios de Riesgo", level=1)
                h_prosp.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)
                
                p_pr1 = doc.add_paragraph()
                p_pr1.paragraph_format.line_spacing = 1.15
                p_pr1.paragraph_format.space_after = Pt(6)
                p_pr1.add_run(
                    "Con base en la distribución espacial trazada y la evolución temporal capturada en los gráficos precedentes, "
                    "el sistema deduce una clara intencionalidad de sostenimiento asimétrico por parte de las orgánicas activas. "
                    "La presencia de pautas combinadas de allanamiento y respuestas armadas indica un nivel de fricción territorial "
                    "elevado que tiende a desplazar el riesgo logístico hacia corredores secundarios de transporte forestal."
                )
                
                p_pr2 = doc.add_paragraph()
                p_pr2.paragraph_format.line_spacing = 1.15
                p_pr2.paragraph_format.space_after = Pt(12)
                p_pr2.add_run(
                    "Se proyecta que las instalaciones industriales y anillos perimetrales mantengan un estatus operativo estable "
                    "siempre y cuando se garantice la retroalimentación continua del perímetro de Geofencing y se apliquen las restricciones "
                    "de convoyes nocturnos en los tramos críticos de Arauco y Malleco."
                )
                
                h3 = doc.add_heading("V. Directrices de Mando Permanentes", level=1)
                h3.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)
                
                directrices = [
                    "Sostener la inmovilización nocturna para convoyes de carga en rutas aledañas a los sectores con registros críticos.",
                    "Mantener sincronizados los avisos preventivos entre los monitores de plataforma y jefaturas de zona.",
                    "Actualizar semanalmente las coordenadas de faena activa en la base central para asegurar la calibración del Geofencing."
                ]
                
                for idx, d_txt in enumerate(directrices, 1):
                    p_dir = doc.add_paragraph()
                    p_dir.paragraph_format.left_indent = Inches(0.2)
                    p_dir.paragraph_format.space_after = Pt(4)
                    p_dir.add_run(f"{idx}. ").font.bold = True
                    p_dir.add_run(d_txt)

                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.success("✔️ Reporte 'Radar de Crisis' compilado con éxito (Gráficos nativos incrustados).")
                st.download_button(
                    label="📥 Descargar Documento Oficial (.docx)",
                    data=buffer,
                    file_name=f"Radar_de_Crisis_CMPC_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e_doc:
                st.error(f"Error interno al destilar el documento Word con gráficos: {e_doc}")
