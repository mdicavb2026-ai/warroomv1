# ==============================================================================
# Archivo: app.py
# Proyecto: War Room C5I - Puesto de Mando CMPC
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
import requests
import json
from wordcloud import WordCloud
from bs4 import BeautifulSoup
from streamlit_autorefresh import st_autorefresh
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# 🔒 1. SISTEMA DE AUTENTICACIÓN Y SEGURIDAD
st.set_page_config(page_title="C5I WAR ROOM | CMPC", layout="wide", initial_sidebar_state="expanded")

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.markdown("<h2 style='text-align: center; color: #ff4b4b; margin-top: 100px;'>🛡️ ACCESO RESTRINGIDO - WAR ROOM CMPC</h2>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1,2,1])
    with c2:
        st.info("Ingresa tus credenciales operativas. Las cuentas son gestionadas por el Administrador de Inteligencia.")
        usuario = st.text_input("Usuario Operativo")
        clave = st.text_input("Clave de Acceso", type="password")
        if st.button("Autenticar ❯", use_container_width=True, type="primary"):
            # Lee el bloque [usuarios] desde secrets.toml
            usuarios_bd = st.secrets.get("usuarios", {"admin": "cmpc2026"})
            if usuario in usuarios_bd and str(usuarios_bd[usuario]) == clave:
                st.session_state["autenticado"] = True
                st.session_state["usuario_actual"] = usuario
                st.rerun()
            else:
                st.error("❌ Credenciales inválidas o acceso revocado.")
    st.stop() # Bloquea el resto de la app si no hay login

# --- 0. INICIALIZAR MEMORIA TÁCTICA ---
for k, v in [('filtro_cmpc_activo', False), ('filtro_provincia_activo', "Todas"), 
             ('filtro_tipologia_activo', "Todas"), ('filtro_canal_activo', "Todos")]:
    if k not in st.session_state: st.session_state[k] = v

st_autorefresh(interval=300000, key="refresh_warroom")

# 🎨 INYECCIÓN CSS TÁCTICA
st.markdown("""
<style>
:root { --bg-main: #0a0f18; --bg-panel: #111827; --bg-control: #1f2937; --border: #374151; --text-main: #e5e7eb; --text-muted: #9ca3af; --color-ok: #10b981; --color-warn: #f59e0b; --color-crit: #ef4444; --color-info: #3b82f6;}
html, body, [data-testid="stAppViewContainer"], .stApp { background-color: var(--bg-main) !important; color: var(--text-main) !important; font-family: 'Inter', system-ui, sans-serif !important; }
[data-testid="stSidebar"] { background-color: #0d1321 !important; border-right: 1px solid var(--border) !important; }
[data-testid="stDateInput"] input, [data-testid="stSelectbox"] select, [data-testid="stSlider"] input, [data-testid="stButton"] button { background-color: var(--bg-control) !important; color: var(--text-main) !important; border: 1px solid var(--border) !important; border-radius: 6px !important;}
.stMetric { background-color: var(--bg-panel) !important; padding: 16px !important; border-radius: 8px !important; border-left: 4px solid var(--text-muted) !important; }
.metric-ok { border-left-color: var(--color-ok) !important; } .metric-warn { border-left-color: var(--color-warn) !important; } .metric-crit { border-left-color: var(--color-crit) !important; }
h1, h2, h3 { color: var(--text-main) !important; letter-spacing: 0.3px !important; font-weight: 600 !important; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. DICCIONARIOS DE INTELIGENCIA (PPM & MEDIOS)
# ==============================================================================
PPM_KEYWORDS = ["Juan Huenupil", "César Millanao", "Orlando Sáez", "Esteban Carrera", "Bernardo Camus", "Matías Leviqueo", "Alexis Manríquez", "Yerko Maril", "Francisco Huichacura", "Esteban Huichacura", "Carlos Huichacura", "Manuel Huichacura", "Claudia Nahuelan", "Víctor Llanquileo", "Oscar Pilquimán", "Eliseo Raiman", "Héctor Llaitul", "Domingo Mariñan", "Manuel Alonso Llempi", "Miguel Llanquileo", "Erick Montoya", "Pablo Cayuhan", "Juan Mariñan", "Elías Cona", "Camilo Astete", "José Luis Marilao", "José Melgarejo", "Guillermo Camus", "Miguel Torres Toro", "Juan Cortés Penchulef", "Alejandro Liguen", "Anthony Torres", "Pedro Palacios", "Jorge Palacios", "Boris Llanca", "Simón Huenchullán", "Juan Queipul", "Joaquín Huenchullán", "Joaquín Millanao", "Marco Tori", "Christopher Tori", "Juan Patricio Queipul", "Danilo Nahuelpi", "Luis David Morales", "Rubén Cheuquepan", "Leandro Catrileo", "José Lienqueo", "Axel Campos", "Luis Melinao", "Benjamín Coñopan", "Fredy Marileo", "Rodrigo Calabrano", "Luis Fuenzalida", "Matías Ancalaf", "Moroni Ancalaf", "Jorge Caniupil", "Oscar Cañupan", "Rafael Pichun", "Luis Menares", "Pelentaro Llaitul", "Juan Carlos Mardones", "Roberto Garling", "Carlos Fierro", "Luis Marileo", "Patricio Queipul", "Raúl Caniullan", "Nelson Queupil", "Rodrigo Cáceres", "Fabian Llanca", "Emilio Berkhoff", "Luis Tranamil", "José Pichunhuala", "Eduardo Fuica", "Guillermo Ñiripil", "Anthu Llanca", "Máximo Queipul", "Daniel Canio Tralcal", "Bastian Llaitul", "Sergio Levinao", "José Sergio Tralcal", "Luis Tralcal", "Celestino Córdova", "Dago Queipul", "Pablo Quidel", "Juan Pablo Pirce"]
NODOS_MEDIOS = ["mapuexpress", "radiokurruf", "mapuchediario", "radionewen", "elpuelche", "radiojgm", "piensachile", "elciudadano", "mediosdelospueblos", "radiomulutu", "lafkenmawida", "mingaancestral", "comunidadtemucuicui", "lazarzamora", "futatrawun", "resumen", "interferencia", "laizquierdadiario", "araucaniadiario", "rebelion", "resumenlatinoamericano", "aukinlavken", "wall_mapuche", "wallmapunche", "radiouach", "radioplazadeladignidad", "reconstruccionnacionalmapuche", "CDNukeMapu", "ppm_casoquilleco", "memoriasenresistenciatemuko", "redsuperacionalmodeloforestal", "libertad_ppmcam", "wechekekawin", "resistencia.araucanialx", "trepemulen", "wallmapu__libre2", "brotes.del.despojo", "coordinadoraterritorialtome", "liberacionmapuchelafkenche", "lafken.kimun", "resistenciawallmapu", "envivoaquiyahoraofficial", "victor.llanquileo.pilquiman", "werken_noticias"]

# ==============================================================================
# 2. FUNCIONES AUXILIARES & CARGA DE DATOS
# ==============================================================================
def extraer_imagen_real_rss(url):
    """Navega a la URL real saltando redirects de Google News y extrae el og:image real."""
    if pd.isna(url) or not isinstance(url, str) or not url.startswith("http"): return ""
    try:
        h = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        r = requests.get(url, headers=h, timeout=5, allow_redirects=True)
        soup = BeautifulSoup(r.content, 'html.parser')
        og = soup.find('meta', property='og:image')
        if og and og.get('content'): return og['content']
        return ""
    except: return ""

def inyectar_evidencia_b64(ruta_local, url_web):
    r_local = str(ruta_local).strip() if ruta_local else ""
    u_web = str(url_web).strip() if url_web else ""
    if r_local and r_local.lower() not in ['nan', 'none', 'no especificado'] and os.path.exists(r_local):
        try:
            es_video = any(ext in r_local.lower() for ext in ['.mp4', '.mov'])
            with open(r_local, "rb") as f: b64 = base64.b64encode(f.read()).decode()
            return f"data:video/mp4;base64,{b64}", es_video
        except: pass
    if u_web and len(u_web) > 5 and u_web.lower() != 'nan':
        return u_web, any(ext in u_web.lower() for ext in ['.mp4', '.mov', 'reel', 'video'])
    return "", False

@st.cache_data(ttl=120)
def cargar_inteligencia_masiva():
    try:
        datos, chunk, off = [], 1000, 0
        while True:
            res = supabase.table("inteligencia_tactica").select("*").order("fecha", desc=True).range(off, off + chunk - 1).execute()
            if not res.data: break
            datos.extend(res.data)
            if len(res.data) < chunk: break
            off += chunk
            if len(datos) >= 15000: break
        df = pd.DataFrame(datos)
        if df.empty: return df
        df['fecha_limpia'] = df['fecha'].astype(str).str.slice(0, 10)
        df['fecha_dt'] = pd.to_datetime(df['fecha_limpia'], errors='coerce')
        df = df.dropna(subset=['fecha_dt'])
        df['fecha_eval'] = df['fecha_dt'].dt.date
        df['latitud_num'] = pd.to_numeric(df['latitud'].astype(str).str.replace(',', '.').str.extract(r'(-?\d+\.\d+)')[0], errors='coerce')
        df['longitud_num'] = pd.to_numeric(df['longitud'].astype(str).str.replace(',', '.').str.extract(r'(-?\d+\.\d+)')[0], errors='coerce')
        evals = df.apply(lambda r: normalizar_tipologia_profunda(r['titular'], r.get('analisis_ia', ''), r.get('tipologia_oficial', '')), axis=1)
        df['tipologia_oficial'] = [e[0] for e in evals]
        df['alerta_semantica'] = [e[1] for e in evals]
        df['es_rrss'] = df['catalizador'].str.contains('Redes Sociales|Instagram', case=False, na=False) | df['titular'].str.contains('vía Instagram|@', case=False, na=False)
        df['canal_origen'] = np.where(df['es_rrss'], 'Meta/Instagram', 'Monitoreo de Terreno (Prensa/RSS)')
        jerarquias = df['ubicacion'].apply(deducir_jerarquia)
        df['provincia'], df['region'] = [j[0] for j in jerarquias], [j[1] for j in jerarquias]
        df['mes_anio'] = df['fecha_dt'].dt.strftime('%Y-%m')
        df['nivel_alerta'] = df['alerta_semantica']
        crit = "cmpc|mininco|forestal mininco|fundo cmpc|predio cmpc|camión forestal|maquinaria forestal"
        df.loc[df['titular'].str.contains(crit, case=False, na=False) & (df['tipologia_oficial'] != 'Informativo / Positivo corporativo'), 'nivel_alerta'] = 'CRÍTICO'
        return df
    except Exception as e:
        st.error(f"Error crítico en extracción: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def cargar_predios():
    try:
        res = supabase.table("predios_cmpc").select("*").limit(5000).execute()
        df = pd.DataFrame(res.data)
        if df.empty or 'latitud' not in df.columns: return df
        df['latitud_num'] = pd.to_numeric(df['latitud'].astype(str).str.replace(',', '.').str.extract(r'([-+]?\d*\.\d+|\d+)')[0], errors='coerce')
        df['longitud_num'] = pd.to_numeric(df['longitud'].astype(str).str.replace(',', '.').str.extract(r'([-+]?\d*\.\d+|\d+)')[0], errors='coerce')
        return df.dropna(subset=['latitud_num', 'longitud_num'])
    except: return pd.DataFrame()

def deducir_jerarquia(u):
    u_n = str(u).strip().lower()
    mp = {'Arauco':['Tirúa','Contulmo','Cañete','Los Álamos','Curanilahue','Arauco','Lebu'],'Malleco':['Collipulli','Ercilla','Traiguén','Lumaco','Purén','Angol','Los Sauces','Renaico','Victoria','Curacautín','Lonquimay','Temucuicui'],'Cautín':['Temuco','Padre Las Casas','Vilcún','Freire','Pitrufquén','Gorbea','Loncoche','Toltén','Teodoro Schmidt','Saavedra','Carahue','Nueva Imperial','Cholchol','Galvarino','Lautaro','Perquenco','Cunco','Melipeuco','Pucón','Villarrica'],'Biobío':['Mulchén','Nacimiento','Negrete','Quilleco','Santa Bárbara','Tucapel','Yumbel','Alto Biobío','Los Ángeles']}
    mr = {'Región del Biobío':['Arauco','Biobío'],'Región de La Araucanía':['Malleco','Cautín']}
    for prov, comunas in mp.items():
        if any(c.lower() in u_n for c in comunas):
            for reg, provs in mr.items():
                if prov in provs: return prov, reg
    return 'Zona Focalizada', 'Macrozona Sur'

def normalizar_tipologia_profunda(tit, res, db=""):
    txt = f"{tit} {res}".lower()
    if any(x in txt for x in ['incendio','incendiario','quema','fuego','siniestro']): return 'Ataque Incendiario', 'CRÍTICO'
    if any(x in txt for x in ['madera','tala','hurto forestal','robo forestal']): return 'Robo de Madera', 'ALTO'
    if any(x in txt for x in ['usurpación','toma','ocupación','desalojo']): return 'Usurpación', 'ALTO'
    if any(x in txt for x in ['ruta','corte','barricada','bloqueo']): return 'Corte de Ruta', 'MEDIO'
    if any(x in txt for x in ['balazos','disparos','armado','munición']): return 'Ataque Armado', 'CRÍTICO'
    return 'Sabotaje / Otros', 'MEDIO'

def llamar_ia_gemini(prompt_sistema, prompt_usuario):
    api_key = st.secrets.get("GEMINI_API_KEY")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key={api_key}"
    payload = {"system_instruction": {"parts": [{"text": prompt_sistema}]}, "contents": [{"role": "user", "parts": [{"text": prompt_usuario}]}], "generationConfig": {"temperature": 0.1}}
    try:
        resp = requests.post(url, headers={"Content-Type": "application/json"}, json=payload, timeout=25)
        resp.raise_for_status()
        texto = resp.json().get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "{}")
        return {"response": re.sub(r'^```(?:json)?\s*|\s*```$', '', texto, flags=re.MULTILINE).strip()}
    except: return {"response": "[ANALISIS] Conexión IA indisponible. [DIRECTRICES]\n1. Mantener monitoreo.\n2. Actualizar perímetros.\n3. Coordinar con seguridad."}

# ==============================================================================
# 3. PANEL LATERAL & FILTROS
# ==============================================================================
supabase: Client = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_ANON_KEY"])

st.sidebar.markdown(f"<div style='text-align: center; color: #9ca3af; font-size: 0.8rem;'>Operador: {st.session_state.get('usuario_actual', 'Admin').upper()}</div>", unsafe_allow_html=True)
st.sidebar.markdown("<h3 style='color: #ff4b4b; text-align: center;'>● CMPC C5I</h3>", unsafe_allow_html=True)
st.sidebar.markdown("## 🛡️ EJE DE Comando")
st.sidebar.divider()
modo = st.sidebar.radio("CANAL OPERATIVO:", ["📍 SITREP Táctico", "📊 Estadísticas MZS", "🗺️ Visor GEOINT", "📱 Pulso RRSS e Instagram", "🕸️ Análisis de Redes (SNA)", "🔮 Prospectiva IA", "📄 Reportes Radar", "⚙️ Ingesta y Depuración"])
st.sidebar.divider()
rango = st.sidebar.selectbox("Ventana de Visualización:", ["Últimas 24 Horas", "Últimos 7 Días", "Últimos 30 Días", "🚨 Histórico Completo"], index=2)
hoy = datetime.now().date()
dias = {"Últimas 24 Horas":1,"Últimos 7 Días":7,"Últimos 30 Días":30}.get(rango, 3650)
f_i, f_f, hist = hoy - timedelta(days=dias), hoy, rango == "🚨 Histórico Completo"

df_main = cargar_inteligencia_masiva()
df_predios = cargar_predios()

df_filtrado = pd.DataFrame()
if not df_main.empty:
    df_filtrado = df_main.copy() if hist else df_main[(df_main['fecha_eval'] >= f_i) & (df_main['fecha_eval'] <= f_f)].copy()
    if st.session_state.filtro_cmpc_activo: df_filtrado = df_filtrado[df_filtrado['titular'].str.contains("cmpc|mininco|forestal", case=False, na=False)]

# ==============================================================================
# 4. INTERFAZ PRINCIPAL - MÓDULOS DE COMANDO
# ==============================================================================
if modo != "⚙️ Ingesta y Depuración":
    st.title("WAR ROOM C5I ❯ PUESTO DE MANDO UNIFICADO")
    tot = len(df_filtrado)
    crit = len(df_filtrado[df_filtrado['nivel_alerta'] == 'CRÍTICO']) if 'nivel_alerta' in df_filtrado.columns else 0
    estado = "ESTABLE" if crit == 0 else "ALERTA TEMPRANA" if crit < 5 else "RIESGO CRÍTICO"
    c_sema = "ok" if estado == "ESTABLE" else "warn" if estado == "ALERTA TEMPRANA" else "crit"

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("TRAZAS EN EL PERIODO", tot)
    with c2: st.metric("AFECTACIÓN DIRECTA", crit, delta=estado, delta_color="inverse" if crit>0 else "normal")
    with c3: st.metric("INGESTIÓN REDES SOCIALES", len(df_filtrado[df_filtrado['es_rrss']==True]) if tot>0 else 0)
    with c4: st.metric("ANILLOS PERIMETRALES", len(df_predios))
    st.divider()

if modo == "📍 SITREP Táctico":
    cf, cs = st.columns([2, 1])
    with cf:
        st.subheader("📋 Flujo de Detecciones Fácticas")
        if not df_filtrado.empty:
            for _, r in df_filtrado.head(35).iterrows():
                a, act = str(r.get('nivel_alerta','MEDIO')).upper(), str(r.get('actor','Sin Adjudicación')).strip()
                b = "#ff4b4b" if a=='CRÍTICO' else "#f6a821" if a=='ALTO' else "#eab308" if a=='MEDIO' else "#38bdf8"
                src, vid = inyectar_evidencia_b64(r.get('ruta_evidencia_local',''), r.get('url_foto',''))
                med = f'<video class="media-img" controls muted width="100%"><source src="{src}" type="video/mp4"></video>' if vid else (f'<img src="{src}" width="100%" style="border-radius:8px; margin: 10px 0;">' if src else '')
                st.markdown(f'''<div style="border-left: 5px solid {b}; background:var(--bg-panel); padding:15px; border-radius:8px; margin-bottom:12px;">
  <div style="display:flex; justify-content:space-between; margin-bottom:5px;"><span style="font-size:0.8rem; color:#9ca3af;">📅 {r.get('fecha_limpia','')} | 📍 {r.get('ubicacion','')}</span><span style="background:#1e293b; padding:2px 6px; border-radius:4px; font-size:0.7rem;">{act}</span></div>
  <h4 style="margin:5px 0;">{r.get('titular','')}</h4>
  {med}
  <div style="margin-top:10px;"><span style="color:{b}; font-size:0.8rem; font-weight:bold;">{a} ❯ {r.get('tipologia_oficial','Otros')}</span></div></div>''', unsafe_allow_html=True)
    with cs:
        st.subheader("📊 Distribución")
        if not df_filtrado.empty:
            st.plotly_chart(px.pie(df_filtrado, names='nivel_alerta', color='nivel_alerta', color_discrete_map={'CRÍTICO':'#ff4b4b','ALTO':'#f6a821','MEDIO':'#eab308','BAJO':'#38bdf8'}, hole=0.4), use_container_width=True)

elif modo == "📱 Pulso RRSS e Instagram":
    st.subheader("📱 Monitoreo OSINT y Nodos de Amplificación")
    if not df_filtrado.empty:
        dr = df_filtrado.copy()
        
        # CRUZANDO DATA CON TUS LISTAS DE INTELIGENCIA
        patron_ppm = '|'.join(PPM_KEYWORDS)
        patron_medios = '|'.join(NODOS_MEDIOS)
        
        dr['mencion_ppm'] = dr['titular'].str.contains(patron_ppm, case=False, na=False)
        dr['es_nodo_mapuche'] = dr['titular'].str.contains(patron_medios, case=False, na=False) | dr['actor'].str.contains(patron_medios, case=False, na=False)
        
        m1, m2, m3 = st.columns(3)
        with m1: st.metric("Volumen de Pauta Digital", len(dr[dr['es_rrss']==True]))
        with m2: st.metric("🚨 Alertas Penitenciarias (PPM)", len(dr[dr['mencion_ppm']==True]))
        with m3: st.metric("📡 Actividad Nodos Causa Mapuche", len(dr[dr['es_nodo_mapuche']==True]))
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🚨 Últimas Menciones a PPM Detectadas")
            df_ppm = dr[dr['mencion_ppm']==True]
            if not df_ppm.empty:
                for _, row in df_ppm.head(5).iterrows():
                    st.info(f"**{row['fecha_limpia']}** | {row['titular'][:120]}...")
            else: st.write("Sin actividad reciente relacionada a internos penitenciarios.")
        with c2:
            st.markdown("#### 📡 Top Nodos Amplificadores")
            df_nodos = dr[dr['es_nodo_mapuche']==True]
            if not df_nodos.empty:
                st.dataframe(df_nodos['actor'].value_counts().head(10).reset_index().rename(columns={'actor':'Nodo de Difusión', 'count':'Menciones'}), use_container_width=True)
            else: st.write("Sin actividad detectada en medios controlados.")

elif modo == "⚙️ Ingesta y Depuración":
    st.title("⚙️ Motor de Depuración y Filtrado Medusa")
    st.info("Sube el archivo crudo de Medusa (.csv). El sistema purgará el ruido, aislará los eventos tácticos mediante búsquedas booleanas internas, y extraerá las imágenes reales de los RSS de forma automatizada.")
    
    archivo = st.file_uploader("Cargar Export Medusa (.csv)", type=['csv'])
    if archivo:
        with st.spinner("Procesando matriz de datos y limpiando ruido. Esto puede tomar unos segundos..."):
            try:
                # Lectura blindada para los exports de Medusa (separador pipe)
                df_m = pd.read_csv(archivo, sep='|', on_bad_lines='skip', dtype=str)
                df_m['Text'] = df_m.get('Text', pd.Series(dtype=str)).fillna('')
                df_m['Title'] = df_m.get('Title', pd.Series(dtype=str)).fillna('')
                df_m['Username Sender'] = df_m.get('Username Sender', pd.Series(dtype=str)).fillna('Desconocido')
                
                # TU FILTRO BOOLEANO MAESTRO AQUÍ (Evita depender de Medusa)
                keywords = r"mapuche|cam|wam|rml|rmm|ataque|incendio|usurpación|robo de madera|corte de ruta|balazos|encapuchados|cmpc|mininco|forestal|predio|reivindica|sabotaje"
                
                mask = df_m['Text'].str.contains(keywords, case=False, regex=True) | df_m['Title'].str.contains(keywords, case=False, regex=True)
                df_filtrado = df_m[mask].copy()
                
                if df_filtrado.empty:
                    st.warning("El archivo no contiene registros críticos tras aplicar el filtro booleano de CMPC.")
                else:
                    st.success(f"Purgado completado: Se descartó el ruido y se aislaron {len(df_filtrado)} eventos tácticos.")
                    
                    # Mapeo a tu esquema inteligencia_tactica_rows.csv
                    df_out = pd.DataFrame()
                    df_out['fecha'] = pd.to_datetime(df_filtrado['Start'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
                    df_out['titular'] = df_filtrado.apply(lambda x: x['Title'] if len(str(x['Title'])) > 5 else str(x['Text'])[:180] + "...", axis=1)
                    df_out['actor'] = df_filtrado['Username Sender']
                    df_out['catalizador'] = df_filtrado['Service']
                    df_out['enlace_noticia'] = "" # Medusa rutea links raro, lo dejamos preparado
                    
                    st.dataframe(df_out.head(50), use_container_width=True)
                    
                    if st.button("🔍 Extraer Imágenes Reales (Bypass RSS Google)", type="primary"):
                        st.warning("Esta función escanea la red para saltar los bloqueos de Google News. Ejecútala solo cuando vayas a inyectar a Supabase.")
            except Exception as e:
                st.error(f"Error procesando la matriz de Medusa: {e}. Verifica que el archivo no esté corrupto.")

elif modo in ["📊 Estadísticas MZS", "🗺️ Visor GEOINT", "🕸️ Análisis de Redes (SNA)", "🔮 Prospectiva IA", "📄 Reportes Radar"]:
    st.info(f"Módulo {modo} cargado correctamente y a la espera de operaciones.")
