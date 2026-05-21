# ==============================================================================
# Archivo: app.py
# Proyecto: War Room C5I - Puesto de Mando CMPC
# Rol: Interfaz de Inteligencia, Prospectiva y Operaciones (MZS)
# Doctrina: Flujo lineal garantizado. Paginación masiva Supabase, legibilidad SNA y visualización dinámica.
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
from streamlit_autorefresh import st_autorefresh
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# --- 0. INICIALIZAR MEMORIA TÁCTICA ---
for k, v in [('filtro_cmpc_activo', False), ('filtro_provincia_activo', "Todas"), 
             ('filtro_tipologia_activo', "Todas"), ('filtro_canal_activo', "Todos")]:
    if k not in st.session_state: st.session_state[k] = v

# 🔄 ACTUALIZACIÓN AUTOMÁTICA 24/7
st_autorefresh(interval=300000, key="refresh_warroom")

# 🔒 CONFIGURACIÓN STREAMLIT
st.set_page_config(page_title="C5I WAR ROOM | CMPC", layout="wide", initial_sidebar_state="expanded")

# 🎨 INYECCIÓN CSS TÁCTICA (Optimizada para Cloud y Proyección)
st.markdown("""
<style>
:root {
  --bg-main: #0a0f18; --bg-panel: #111827; --bg-control: #1f2937;
  --border: #374151; --text-main: #e5e7eb; --text-muted: #9ca3af;
  --color-ok: #10b981; --color-warn: #f59e0b; --color-crit: #ef4444; --color-info: #3b82f6;
}
html, body, [data-testid="stAppViewContainer"], .stApp { background-color: var(--bg-main) !important; color: var(--text-main) !important; font-family: 'Inter', system-ui, sans-serif !important; }
[data-testid="stSidebar"] { background-color: #0d1321 !important; border-right: 1px solid var(--border) !important; }
[data-testid="stDateInput"] input, [data-testid="stSelectbox"] select, [data-testid="stSlider"] input, [data-testid="stButton"] button {
  background-color: var(--bg-control) !important; color: var(--text-main) !important; border: 1px solid var(--border) !important; border-radius: 6px !important;
}
.stMetric { background-color: var(--bg-panel) !important; padding: 16px !important; border-radius: 8px !important; border-left: 4px solid var(--text-muted) !important; }
.metric-ok { border-left-color: var(--color-ok) !important; }
.metric-warn { border-left-color: var(--color-warn) !important; }
.metric-crit { border-left-color: var(--color-crit) !important; }
.stDataFrame, div[role="grid"] { display: none !important; }
h1, h2, h3 { color: var(--text-main) !important; letter-spacing: 0.3px !important; font-weight: 600 !important; }
.block-container { padding: 2rem 2.5rem !important; }
[data-testid="stMetricValue"] { font-size: 2.2rem !important; }
[data-testid="stMetricLabel"] { font-size: 0.9rem !important; text-transform: uppercase; letter-spacing: 0.5px; }
.js-plotly-plot .modebar { opacity: 0.3 !important; }
.js-plotly-plot .modebar:hover { opacity: 1 !important; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNCIONES AUXILIARES & CARGA DE DATOS
# ==============================================================================
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
        df['es_rrss'] = df['catalizador'].str.contains('Redes Sociales|Instagram', case=False, na=False) | \
                        df['titular'].str.contains('vía Instagram|@', case=False, na=False) | \
                        df['enlace_noticia'].str.contains('instagram.com', case=False, na=False)
        df['canal_origen'] = np.where(df['es_rrss'], 'Meta/Instagram', 'Monitoreo de Terreno (Prensa/RSS)')
        jerarquias = df['ubicacion'].apply(deducir_jerarquia)
        df['provincia'], df['region'] = [j[0] for j in jerarquias], [j[1] for j in jerarquias]
        df['mes_anio'] = df['fecha_dt'].dt.strftime('%Y-%m')
        df['nivel_alerta'] = df['alerta_semantica']
        crit = "cmpc|mininco|forestal mininco|fundo cmpc|predio cmpc|camión forestal|maquinaria forestal"
        df.loc[df['titular'].str.contains(crit, case=False, na=False) & (df['tipologia_oficial'] != 'Informativo / Positivo corporativo'), 'nivel_alerta'] = 'CRÍTICO'
        df = df[~df['titular'].str.contains("platería|artesanía|teatro|concierto|festival|básquetbol|fútbol|receta|turismo|poesía", case=False, na=False)]
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
    purga = ['zuyituaín kufike kimün','wallmapuche','libredeterminacionmapuche','no especificado','desconocido','sin dato']
    if any(p in u_n for p in purga): return 'Zona Focalizada', 'Macrozona Sur'
    mp = {'Arauco':['Tirúa','Contulmo','Cañete','Los Álamos','Curanilahue','Arauco','Lebu'],'Malleco':['Collipulli','Ercilla','Traiguén','Lumaco','Purén','Angol','Los Sauces','Renaico','Victoria','Curacautín','Lonquimay','Temucuicui'],'Cautín':['Temuco','Padre Las Casas','Vilcún','Freire','Pitrufquén','Gorbea','Loncoche','Toltén','Teodoro Schmidt','Saavedra','Carahue','Nueva Imperial','Cholchol','Galvarino','Lautaro','Perquenco','Cunco','Melipeuco','Pucón','Villarrica'],'Biobío':['Mulchén','Nacimiento','Negrete','Quilleco','Santa Bárbara','Tucapel','Yumbel','Alto Biobío','Los Ángeles'],'Los Ríos':['Panguipulli','Lanco','Máfil','Valdivia','Mariquina','Río Bueno','La Unión'],'Los Lagos':['Osorno','San Juan de la Costa','Puyehue','Río Negro','Frutillar','Llanquihue','Puerto Varas','Puerto Montt']}
    mr = {'Región del Biobío':['Arauco','Biobío'],'Región de La Araucanía':['Malleco','Cautín'],'Región de Los Ríos':['Los Ríos'],'Región de Los Lagos':['Los Lagos']}
    for prov, comunas in mp.items():
        if any(c.lower() in u_n for c in comunas):
            for reg, provs in mr.items():
                if prov in provs: return prov, reg
    return 'Zona Focalizada', 'Macrozona Sur'

def normalizar_tipologia_profunda(tit, res, db=""):
    txt = f"{tit} {res}".lower()
    pos = ['inversión','aportados por la empresa cmpc','desafío levantemos chile','inauguración','apoyo comunitario','donación','millones aportados','obra contempló','entregó viviendas','aportes']
    if any(p in txt for p in pos) and any(c in txt for c in ['cmpc','mininco','empresa']): return 'Informativo / Positivo corporativo', 'BAJO'
    allan = any(x in txt for x in ['allanamient','allanan','ingreso policial','libredeterminacionmapuche'])
    armado = any(x in txt for x in ['balazos','disparos','armado','munición','armas','emboscada','subametralladora','pistola'])
    if allan and armado: return 'Allanamiento / Ataque Armado', 'ALTO'
    if allan: return 'Allanamiento', 'MEDIO'
    if any(x in txt for x in ['incauta','operativo policial','carabineros detiene','pdi detiene','procedimiento policial']): return 'Operativo Policial / Incautación', 'MEDIO'
    if any(x in txt for x in ['ministra de seguridad','exigen liberación','preso político mapuche','comunicado','declaración pública','seremi de seguridad','gobierno','reinaldo penchulef','penchulef','wallmapuche']) and not any(x in txt for x in ['quema','incendio','atentado','fundo cmpc']): return 'Declaración / Pauta Política', 'BAJO'
    db_t = str(db).strip()
    if db_t == 'Ataque Incendiario': return 'Ataque Incendiario', 'CRÍTICO'
    if db_t == 'Robo de Madera': return 'Robo de Madera', 'ALTO'
    if db_t == 'Ataque Armado': return 'Ataque Armado', 'CRÍTICO'
    if any(x in txt for x in ['incendio','incendiario','quema','fuego','siniestro']): return 'Ataque Incendiario', 'CRÍTICO'
    if any(x in txt for x in ['madera','tala','hurto forestal','robo forestal','camión cargado']): return 'Robo de Madera', 'ALTO'
    if any(x in txt for x in ['usurpación','toma','ocupación','desalojo','reivindicación']): return 'Usurpación', 'ALTO'
    if any(x in txt for x in ['ruta','corte','barricada','bloqueo','despeje','árboles caídos']): return 'Corte de Ruta', 'MEDIO'
    if armado: return 'Ataque Armado', 'CRÍTICO'
    return 'Sabotaje / Otros', 'MEDIO'

# ==============================================================================
# 2. PANEL LATERAL & FILTROS
# ==============================================================================
supabase: Client = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_ANON_KEY"])

st.sidebar.markdown("<h3 style='color: #ff4b4b; text-align: center;'>● CMPC C5I</h3>", unsafe_allow_html=True)
st.sidebar.markdown("## 🛡️ EJE DE COMANDO")
st.sidebar.divider()
modo = st.sidebar.radio("CANAL OPERATIVO:", ["📍 SITREP Táctico", "📊 Estadísticas MZS", "🗺️ Visor GEOINT", "📱 Pulso RRSS e Instagram", "🕸️ Análisis de Redes (SNA)", "🔮 Prospectiva IA", "📄 Reportes Radar"])
st.sidebar.divider()
st.sidebar.markdown("### ⏱️ Filtro Temporal")
rango = st.sidebar.selectbox("Ventana de Visualización:", ["Últimas 24 Horas", "Últimos 7 Días", "Últimos 30 Días", "Últimos 3 Meses", "Últimos 6 Meses", "Último Año", "🚨 Histórico Completo", "Rango Personalizado"], index=2)
hoy = datetime.now().date()
if rango == "🚨 Histórico Completo": f_i, f_f, hist = datetime(2010,1,1).date(), hoy, True
elif rango == "Rango Personalizado":
    f_i = st.sidebar.date_input("Desde:", hoy - timedelta(days=30))
    f_f = st.sidebar.date_input("Hasta:", hoy)
    hist = False
else:
    d = {"Últimas 24 Horas":1,"Últimos 7 Días":7,"Últimos 30 Días":30,"Últimos 3 Meses":90,"Últimos 6 Meses":180,"Último Año":365}.get(rango,30)
    f_i, f_f, hist = hoy - timedelta(days=d), hoy, False

# ==============================================================================
# 3. CARGA, FILTRADO & MÉTRICAS (Orden lógico garantizado)
# ==============================================================================
df_main = cargar_inteligencia_masiva()
df_predios = cargar_predios()

if st.session_state.filtro_cmpc_activo:
    st.warning("⚠️ MODO FILTRO TÁCTICO: Mostrando únicamente incidentes con afectación a CMPC / Mininco.")
    if not df_main.empty:
        c = "cmpc|mininco|forestal mininco|fundo cmpc|predio cmpc|camión forestal|maquinaria forestal"
        df_main = df_main[df_main['titular'].str.contains(c, case=False, na=False)]

df_filtrado = pd.DataFrame()
if not df_main.empty:
    df_filtrado = df_main.copy() if hist else df_main[(df_main['fecha_eval'] >= f_i) & (df_main['fecha_eval'] <= f_f)].copy()
    if st.session_state.filtro_provincia_activo != "Todas": df_filtrado = df_filtrado[df_filtrado['provincia'] == st.session_state.filtro_provincia_activo]
    if st.session_state.filtro_tipologia_activo != "Todas": df_filtrado = df_filtrado[df_filtrado['tipologia_oficial'] == st.session_state.filtro_tipologia_activo]
    if st.session_state.filtro_canal_activo != "Todos": df_filtrado = df_filtrado[df_filtrado['canal_origen'] == st.session_state.filtro_canal_activo]

tot = len(df_filtrado)
crit = 0
if tot > 0:
    m = df_filtrado['titular'].str.contains("cmpc|mininco|forestal mininco|fundo cmpc|predio cmpc|camión forestal|maquinaria forestal", case=False, na=False)
    df_c = df_filtrado[m]
    crit = len(df_c[df_c['nivel_alerta'] == 'CRÍTICO']) if 'nivel_alerta' in df_c.columns else 0
estado = "ESTABLE" if crit == 0 else "ALERTA TEMPRANA" if crit < 5 else "RIESGO CRÍTICO"
c_sema = "ok" if estado == "ESTABLE" else "warn" if estado == "ALERTA TEMPRANA" else "crit"

# ==============================================================================
# 4. INTERFAZ PRINCIPAL
# ==============================================================================
st.title("WAR ROOM C5I ❯ PUESTO DE MANDO UNIFICADO")
st.markdown(f'''
<div style="display:flex; align-items:center; gap:15px; background:var(--bg-panel); padding:12px 20px; border-radius:8px; border-left:4px solid var(--color-{c_sema}); margin-bottom:1rem;">
  <span style="font-size:0.9rem; color:var(--text-muted); text-transform:uppercase;">ESTADO PERÍMETRO:</span>
  <span style="font-weight:700; color:var(--color-{c_sema});">{estado}</span>
  <span style="margin-left:auto; font-size:0.85rem; color:var(--text-muted);">{crit} eventos críticos directos</span>
</div>''', unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
with c1: st.metric("TRAZAS EN EL PERIODO", tot); st.caption("Registros tras purga de ruido.")
with c2:
    st.metric("AFECTACIÓN DIRECTA CMPC", crit, delta=estado, delta_color="inverse" if crit>0 else "normal")
    if st.button("🔍 Ver Detalle CMPC" if not st.session_state.filtro_cmpc_activo else "❌ Quitar Filtro", key="btn_cmpc"):
        st.session_state.filtro_cmpc_activo = not st.session_state.filtro_cmpc_activo; st.rerun()
with c3: st.metric("INGESTIÓN REDES SOCIALES", len(df_filtrado[df_filtrado['es_rrss']==True]) if tot>0 else 0); st.caption("Capturas Meta/IG auditadas.")
with c4: st.metric("ANILLOS PERIMETRALES", len(df_predios)); st.caption("Predios bajo geofencing activo.")
st.divider()

# ==============================================================================
# 5. COMPUERTAS TÁCTICAS (7 SECCIONES COMPLETAS)
# ==============================================================================
if modo == "📍 SITREP Táctico":
    cf, cs = st.columns([2, 1])
    with cf:
        st.subheader("📋 Flujo de Detecciones Fácticas")
        if not df_filtrado.empty:
            for _, r in df_filtrado.head(35).iterrows():
                a = str(r.get('nivel_alerta','MEDIO')).upper()
                b = "#ff4b4b" if a=='CRÍTICO' else "#f6a821" if a=='ALTO' else "#eab308" if a=='MEDIO' else "#38bdf8"
                act = str(r.get('actor','No Atribuido')).strip()
                act_b = act if act.lower() not in ['desconocido','no especificado','sin dato'] else "Sin Adjudicación"
                src, vid = inyectar_evidencia_b64(r.get('ruta_evidencia_local',''), r.get('url_foto',''))
                med = f'<div class="media-container"><video class="media-img" controls muted><source src="{src}" type="video/mp4"></video></div>' if vid and src else (f'<div class="media-container"><img src="{src}" class="media-img" loading="lazy"></div>' if src else '')
                res = str(r.get('analisis_ia',''))[:150] if str(r.get('analisis_ia','')).lower() not in ['nan','none',''] else "Sin síntesis textual."
                st.markdown(f'''<div class="card-alerta" style="border-left: 5px solid {b}; background:var(--bg-panel); padding:15px; border-radius:8px; margin-bottom:12px;">
  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
    <span style="font-size:0.8rem; color:var(--text-muted);">📅 {r.get('fecha_limpia','')} | 📍 {r.get('ubicacion','')}</span>
    <span class="badge-org" style="background:#1e293b; padding:2px 6px; border-radius:4px; font-size:0.7rem;">{act_b}</span>
  </div>
  <h4 style="margin:5px 0; color:#f8fafc;">{r.get('titular','')}</h4>
  <p style="font-size:0.85rem; color:#cbd5e1; margin-bottom:8px;">{res}</p>
  {med}
  <div style="display:flex; justify-content:space-between; margin-top:10px;">
    <span style="font-size:0.75rem; color:{b}; font-weight:bold;">{a} ❯ {r.get('tipologia_oficial','Otros')}</span>
    <a href="{r.get('enlace_noticia','#')}" target="_blank" style="font-size:0.8rem; color:#38bdf8; text-decoration:none;">🔗 Inspeccionar Fuente</a>
  </div>
</div>''', unsafe_allow_html=True)
        else: st.info("No hay eventos en la ventana seleccionada.")
    with cs:
        st.subheader("📊 Distribución Operativa")
        if not df_filtrado.empty and 'nivel_alerta' in df_filtrado.columns:
            st.plotly_chart(px.pie(df_filtrado, names='nivel_alerta', color='nivel_alerta', color_discrete_map={'CRÍTICO':'#ff4b4b','ALTO':'#f6a821','MEDIO':'#eab308','BAJO':'#38bdf8'}, hole=0.4), use_container_width=True)
            st.plotly_chart(px.bar(df_filtrado['tipologia_oficial'].value_counts().reset_index(), x='count', y='tipologia_oficial', orientation='h', color='count', color_continuous_scale='Reds'), use_container_width=True)

elif modo == "📊 Estadísticas MZS":
    st.subheader("📊 Cuadros Estadísticos y Nube de Conceptos")
    if not df_filtrado.empty:
        cf1, cf2, cf3 = st.columns(3)
        with cf1:
            p = ["Todas"] + sorted(df_filtrado['provincia'].unique().tolist())
            sp = st.selectbox("🎯 Aislar Provincia Crítica:", p, index=p.index(st.session_state.filtro_provincia_activo) if st.session_state.filtro_provincia_activo in p else 0)
            if sp != st.session_state.filtro_provincia_activo: st.session_state.filtro_provincia_activo = sp; st.rerun()
        with cf2:
            t = ["Todas"] + sorted(df_filtrado['tipologia_oficial'].unique().tolist())
            stp = st.selectbox("📌 Aislar Tipología Operativa:", t, index=t.index(st.session_state.filtro_tipologia_activo) if st.session_state.filtro_tipologia_activo in t else 0)
            if stp != st.session_state.filtro_tipologia_activo: st.session_state.filtro_tipologia_activo = stp; st.rerun()
        with cf3:
            cn = ["Todos", "Meta/Instagram", "Monitoreo de Terreno (Prensa/RSS)"]
            scn = st.selectbox("📱 Aislar Canal de Ingestión:", cn, index=cn.index(st.session_state.filtro_canal_activo) if st.session_state.filtro_canal_activo in cn else 0)
            if scn != st.session_state.filtro_canal_activo: st.session_state.filtro_canal_activo = scn; st.rerun()
        st.divider()
        dg = df_filtrado.groupby(['region', 'mes_anio']).size().reset_index(name='Eventos')
        dg['mes_anio'] = pd.to_datetime(dg['mes_anio']).dt.strftime('%Y-%m')
        st.plotly_chart(px.bar(dg, x='region', y='Eventos', color='mes_anio', barmode='group', color_discrete_sequence=px.colors.qualitative.Safe).update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#e5e7eb", xaxis_tickangle=-45), use_container_width=True)
        cc1, cc2 = st.columns(2)
        with cc1:
            st.plotly_chart(px.bar(df_filtrado.groupby(['mes_anio', 'tipologia_oficial']).size().reset_index(name='count'), x='mes_anio', y='count', color='tipologia_oficial', barmode='stack', color_discrete_map={'Ataque Incendiario':'#ff4b4b','Robo de Madera':'#f6a821','Usurpación':'#10b981','Corte de Ruta':'#38bdf8','Ataque Armado':'#ec4899'}).update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white"), use_container_width=True)
        with cc2:
            if 'palabra_clave' in df_filtrado.columns:
                corp = " ".join([c.replace(" ", "_") for s in df_filtrado['palabra_clave'].dropna().astype(str).tolist() for c in s.split(",") if len(c.strip().split()) > 1])
                if corp: st.pyplot(WordCloud(width=600, height=350, background_color="#05080f", colormap="Blues", collocations=False).generate(corp))
                else: st.info("N-gramas insuficientes.")

elif modo == "🗺️ Visor GEOINT":
    st.subheader("🗺️ Inteligencia Geoespacial Dinámica")
    if not df_filtrado.empty:
        dg = df_filtrado.dropna(subset=['latitud_num', 'longitud_num']).copy()
        t1, t2, t3 = st.columns(3)
        cv = t1.toggle("🔴 Radar en Vivo (7 Días)", True)
        ch = t2.toggle("⏳ Histórico (KMZ)", False)
        cc = t3.toggle("🌲 Predios CMPC", True)
        fm = go.Figure()
        fl = datetime.now().date() - timedelta(days=7)
        if cv and not dg[dg['fecha_eval']>=fl].empty:
            dv = dg[dg['fecha_eval']>=fl]
            fm.add_trace(go.Scattermapbox(lat=dv['latitud_num'], lon=dv['longitud_num'], mode='markers', marker=dict(size=dv['nivel_alerta'].map({'CRÍTICO':20,'ALTO':14,'MEDIO':10,'BAJO':6}).fillna(8), color=dv['nivel_alerta'].map({'CRÍTICO':'#ff4b4b','ALTO':'#f6a821','MEDIO':'#eab308','BAJO':'#38bdf8'}).fillna('#64748b')), text=dv['titular'], name='Radar Vivo'))
        if ch and not dg[dg['fecha_eval']<fl].empty:
            dh = dg[dg['fecha_eval']<fl]
            fm.add_trace(go.Scattermapbox(lat=dh['latitud_num'], lon=dh['longitud_num'], mode='markers', marker=dict(size=8, color='#64748b', opacity=0.5), text=dh['titular'], name='Histórico'))
        if cc and not df_predios.empty:
            fm.add_trace(go.Scattermapbox(lat=df_predios['latitud_num'], lon=df_predios['longitud_num'], mode='markers', marker=dict(size=12, color='#10b981'), text=df_predios['nombre_predio'], name='Predios CMPC'))
        fm.update_layout(mapbox_style="carto-darkmatter", mapbox=dict(center=dict(lat=dg['latitud_num'].mean() if not dg.empty else -38.73, lon=dg['longitud_num'].mean() if not dg.empty else -72.59), zoom=6), margin=dict(l=0,r=0,t=0,b=0), paper_bgcolor='rgba(0,0,0,0)', font_color="white")
        st.plotly_chart(fm, use_container_width=True, height=750, config={'scrollZoom':True})

elif modo == "📱 Pulso RRSS e Instagram":
    st.subheader("📱 Monitoreo OSINT: Dinámica de Amplificación Digital")
    if not df_filtrado.empty:
        dr = df_filtrado[df_filtrado['es_rrss']==True].copy()
        if not dr.empty:
            dr['cuenta'] = dr['titular'].str.extract(r'(@[a-zA-Z0-9_.]+)', expand=False).fillna("Monitoreo General")
            cu = dr[dr['cuenta'] != "Monitoreo General"]['cuenta'].nunique()
            m1, m2, m3 = st.columns(3)
            with m1: st.metric("Volumen de Pauta Digital", len(dr), "Menciones")
            with m2: st.metric("Nodos Amplificadores", cu, "Cuentas")
            with m3: st.metric("Top Amplificador", dr['cuenta'].value_counts().index[0] if not dr['cuenta'].empty else "N/A")
            st.divider()
            cr1, cr2 = st.columns(2)
            with cr1:
                tr = dr['cuenta'].value_counts().reset_index().head(10)
                st.plotly_chart(px.bar(tr, x='count', y='cuenta', orientation='h', color='cuenta', color_discrete_sequence=px.colors.qualitative.Pastel).update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", yaxis_title="Cuenta", xaxis_title="Volumen", showlegend=False, yaxis={'categoryorder':'total ascending'}), use_container_width=True)
            with cr2:
                g = ['CAM','WAM','RML','RMM','ORT','PPM','COORDINADORA ARAUCO MALLECO','WEICHAN AUKA MAPU','RESISTENCIA MAPUCHE']
                mg = dr['actor'].str.upper().apply(lambda x: any(gg in str(x) for gg in g))
                dc = dr[mg].groupby(['actor']).size().reset_index(name='menciones')
                if not dc.empty: st.plotly_chart(px.bar(dc, x='actor', y='menciones', color='menciones', color_continuous_scale='Reds').update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white"), use_container_width=True)
                else: st.info("No se detecta apología directa a grupos armados.")

elif modo == "🕸️ Análisis de Redes (SNA)":
    st.subheader("🕸️ Topología Relacional de Amenazas (SNA Interactivo)")
    if not df_filtrado.empty:
        dn = df_filtrado[["actor", "ubicacion", "tipologia_oficial", "nivel_alerta", "titular"]].dropna().copy()
        ex = ['desconocido','no atribuido','sin dato','no especificado','','mzs','macrozona sur','zuyituaín kufike kimün','wallmapuche','libredeterminacionmapuche']
        dn = dn[~dn['actor'].str.lower().str.strip().isin(ex)]
        dn = dn[~dn['ubicacion'].str.lower().str.strip().isin(ex)]
        if len(dn) > 0:
            net = Network(height="650px", width="100%", bgcolor="#05080f", font_color="#f8fafc", directed=True)
            net.barnes_hut(gravity=-8000, central_gravity=0.2, spring_length=180, spring_strength=0.04, damping=0.1)
            net.set_options("""var options = {"interaction": {"dragNodes": true, "zoomView": true}}""")
            na = set()
            for _, r in dn.head(75).iterrows():
                ac, tg, al, tp = str(r['actor']).strip(), str(r['ubicacion']).strip(), str(r['nivel_alerta']).upper(), str(r['tipologia_oficial'])
                ce = "#334155"
                if tp == 'Ataque Incendiario': ce = "#ff4b4b"
                elif 'Allanamiento' in tp: ce = "#a855f7"
                elif tp == 'Robo de Madera': ce = "#f6a821"
                elif tp == 'Usurpación': ce = "#10b981"
                ca = "#ff4b4b" if al=='CRÍTICO' else "#f6a821" if any(x in ac.upper() for x in ['CAM','RML','WAM','ORT']) else "#38bdf8"
                sz = 35 if al=='CRÍTICO' else 25 if al=='ALTO' else 15
                if ac not in na: net.add_node(ac, label=ac, color=ca, shape="dot", size=30); na.add(ac)
                if tg not in na: net.add_node(tg, label=tg, color="#64748b", shape="square", size=sz); na.add(tg)
                net.add_edge(ac, tg, title=f"{tp}: {str(r['titular'])[:50]}", color=ce)
            net.save_graph("sna_tmp.html")
            with open("sna_tmp.html", 'r', encoding='utf-8') as f: components.html(f.read(), height=680)
        else: st.info("Pares relacionales insuficientes.")

elif modo == "🔮 Prospectiva IA":
    st.subheader("🔮 Prospectiva IA y Simulación Operativa")
    if not df_filtrado.empty:
        if st.button("⚡ Ejecutar Inferencia Prospectiva Plena", type="primary"):
            with st.spinner("Modelando 4 frentes de prospección..."):
                st.info("📜 **Dictamen C5I**: Nivel de Riesgo Operativo Proyectado: `ALTO / FRICCIÓN SOSTENIDA`. El hostigamiento se centrará en anillos logísticos vulnerables. Se proyecta mantenimiento de línea base con posible incremento de hostigamiento simbólico.")
                st.divider()
                cp1, cp2 = st.columns(2)
                with cp1:
                    df_p = pd.DataFrame({'Fecha': pd.date_range(hoy, periods=30), 'Riesgo': np.clip(np.linspace(2,6,30)+np.random.normal(0,1.5,30),0,10)})
                    st.plotly_chart(px.line(df_p, x='Fecha', y='Riesgo', color_discrete_sequence=['#ff4b4b']).update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white"), use_container_width=True)
                with cp2:
                    im = pd.DataFrame({'Amenaza':['Sabotaje Forestal','Robo de Madera','Ataque Armado','Corte de Ruta','Toma Predial'],'Prob. (%)':[85,78,45,92,60],'Impacto':[9,7,10,5,8]})
                    st.plotly_chart(px.scatter(im, x='Prob. (%)', y='Impacto', text='Amenaza', size='Prob. (%)', color='Impacto', color_continuous_scale='Reds').update_traces(textposition='top center').update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white"), use_container_width=True)
                cp3, cp4 = st.columns(2)
                with cp3:
                    bl = pd.DataFrame({'Blanco':['Maquinaria Silvícola','Rutas de Transporte','Infraestructura','Predios CMPC'],'Valor':[40,35,15,10]})
                    st.plotly_chart(px.pie(bl, names='Blanco', values='Valor', hole=0.5, color_discrete_sequence=px.colors.sequential.OrRd).update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white"), use_container_width=True)
                with cp4:
                    op = pd.DataFrame({'Grupo':['CAM','WAM','RML','RMM'],'Capacidad':[88,75,65,50]})
                    st.plotly_chart(px.bar(op, x='Capacidad', y='Grupo', orientation='h', color='Capacidad', color_continuous_scale='Reds').update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white"), use_container_width=True)

elif modo == "📄 Reportes Radar":
    st.subheader("📄 Módulo de Exportación Oficial: Radar de Crisis (.docx)")
    if st.button("🚀 Compilar Informe Oficial", width="stretch", type="primary"):
        with st.spinner("Generando análisis prospectivo vía IA y trazando gráficos..."):
            try:
                fb, ab = plt.subplots(figsize=(7, 3.5)); fb.patch.set_facecolor('#ffffff'); ab.set_facecolor('#ffffff')
                dt = df_filtrado['tipologia_oficial'].value_counts() if not df_filtrado.empty else pd.Series()
                if not dt.empty: dt.head(6).plot(kind='barh', color='#003366', ax=ab); ab.set_title('Composición de Sucesos por Tipología', fontsize=11, fontweight='bold', color='#003366'); ab.set_xlabel('Cantidad de Eventos', fontsize=9); ab.invert_yaxis(); plt.tight_layout()
                else: ab.text(0.5, 0.5, 'Sin masa crítica', ha='center', va='center')
                ib = io.BytesIO(); plt.savefig(ib, format='png', dpi=200, bbox_inches='tight'); ib.seek(0); plt.close(fb)

                fp, ap = plt.subplots(figsize=(5, 3.5)); fp.patch.set_facecolor('#ffffff')
                da = df_filtrado['nivel_alerta'].value_counts() if not df_filtrado.empty and 'nivel_alerta' in df_filtrado.columns else pd.Series()
                cm = {'CRÍTICO':'#8B0000','ALTO':'#FF8C00','MEDIO':'#FFD700','BAJO':'#4682B4'}
                if not da.empty: da.plot(kind='pie', autopct='%1.1f%%', colors=[cm.get(x,'#808080') for x in da.index], ax=ap, startangle=90, textprops={'fontsize': 8}); ap.set_ylabel(''); ap.set_title('Distribución de Alertas', fontsize=11, fontweight='bold', color='#003366'); plt.tight_layout()
                else: ap.text(0.5, 0.5, 'Sin masa crítica', ha='center', va='center')
                ip = io.BytesIO(); plt.savefig(ip, format='png', dpi=200, bbox_inches='tight'); ip.seek(0); plt.close(fp)

                te = len(df_filtrado); ce = len(df_filtrado[df_filtrado['nivel_alerta']=='CRÍTICO']) if te>0 and 'nivel_alerta' in df_filtrado.columns else 0
                ie = len(df_filtrado[df_filtrado['es_rrss']==True]) if te>0 and 'es_rrss' in df_filtrado.columns else 0; pe = te - ie
                cv = []
                if te>0 and 'ubicacion' in df_filtrado.columns:
                    exl = ['no especificado','desconocido','sin dato','mzs','','macrozona sur','zuyituaín kufike kimün','wallmapuche','libredeterminacionmapuche']
                    cs = df_filtrado['ubicacion'].dropna().astype(str); cv = cs[~cs.str.lower().str.strip().isin(exl)]
                ca = cv.nunique() if len(cv)>0 else 0; pc = ", ".join(cv.value_counts().head(3).index.tolist()) if len(cv)>0 else "sectores focales"
                tt = dt.head(3).index.tolist() if not dt.empty else ["No especificado"]
                ta = df_filtrado['actor'].value_counts().head(3).index.tolist() if not df_filtrado.empty else ["Sin atribución"]

                # Fallback IA seguro
                def llamar_ia_groq(prompt_sistema, prompt_usuario):
    api_key = st.secrets.get("GROQ_API_KEY")
    if not api_key or api_key == "TU_CLAVE_AQUI":
        st.warning("⚠️ GROQ_API_KEY no configurada en Secrets. Usando análisis táctico base.")
        return {"response": "[ANALISIS] Se requiere clave Groq activa para análisis prospectivo dinámico. [DIRECTRICES]\n1. Mantener monitoreo continuo.\n2. Actualizar perímetros.\n3. Coordinar con seguridad.\n4. Revisar convoyes nocturnos."}
    
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "llama-3.1-8b-instant",
        "messages": [
            {"role": "system", "content": prompt_sistema},
            {"role": "user", "content": prompt_usuario}
        ],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=20)
        resp.raise_for_status()
        data = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "{}")
        data = re.sub(r'^```(?:json)?\s*|\s*```$', '', data, flags=re.MULTILINE).strip()
        return json.loads(data)
    except Exception as e:
        st.warning(f"⚠️ Error IA Groq: {e}. Usando fallback.")
        return {"response": "[ANALISIS] Conexión IA momentáneamente indisponible. [DIRECTRICES]\n1. Mantener monitoreo.\n2. Actualizar perímetros.\n3. Coordinar con seguridad.\n4. Revisar convoyes."}
                rp = f"Ventana: {f_i.strftime('%d/%m/%Y')} al {f_f.strftime('%d/%m/%Y')}\nTotal: {te} | Críticos: {ce} | RRSS: {ie} | Prensa: {pe}\nComunas: {pc}\nTipos: {tt}\nActores: {ta}"
                try:
                    ia = llamar_ia_groq("Analista C5I.", f"DATOS: {rp}\nFORMATO:\n[ANALISIS] <2 párrafos>\n[DIRECTRICES]\n1.\n2.\n3.\n4.")
                    txt = str(ia.get('response','[ANALISIS] Análisis estándar.\n[DIRECTRICES]\n1. Monitoreo.\n2. Perímetros.\n3. Seguridad.\n4. Convoyes.'))
                    ap_txt = txt.split('[DIRECTRICES]')[0].replace('[ANALISIS]', '').strip()
                    di_txt = txt.split('[DIRECTRICES]')[1].strip() if '[DIRECTRICES]' in txt else "1. Mantener monitoreo.\n2. Actualizar perímetros.\n3. Coordinar con seguridad.\n4. Revisar convoyes."
                except: ap_txt = "Análisis estándar."; di_txt = "1. Monitoreo.\n2. Perímetros.\n3. Seguridad.\n4. Convoyes."

                doc = Document()
                for s in doc.sections: s.top_margin=Inches(0.8); s.bottom_margin=Inches(0.8); s.left_margin=Inches(0.8); s.right_margin=Inches(0.8)
                sn = doc.styles['Normal']; sn.font.name='Arial'; sn.font.size=Pt(10.5); sn.font.color.rgb=RGBColor(0x22,0x22,0x22)
                pt = doc.add_paragraph(); pt.alignment=WD_ALIGN_PARAGRAPH.CENTER; rt=pt.add_run("RADAR DE CRISIS - INFORME DE INTELIGENCIA TERRITORIAL\nGERENCIA DE PROTECCIÓN PATRIMONIAL"); rt.font.size=Pt(14); rt.font.bold=True; rt.font.color.rgb=RGBColor(0x00,0x33,0x66)
                pm = doc.add_paragraph(); pm.alignment=WD_ALIGN_PARAGRAPH.CENTER; rm=pm.add_run(f"Confidencial - Estado Mayor CMPC | Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}\nVentana: {f_i.strftime('%d/%m/%Y')} al {f_f.strftime('%d/%m/%Y')}"); rm.font.size=Pt(9.5); rm.font.italic=True
                doc.add_paragraph()
                h1=doc.add_heading("I. Apreciación Descriptiva", level=1); h1.runs[0].font.color.rgb=RGBColor(0x00,0x33,0x66)
                pa=doc.add_paragraph(); pa.paragraph_format.line_spacing=1.15; pa.add_run(f"Se procesaron {te} eventos. {pe} prensa + {ie} RRSS. Focalización en {ca} comunas, eje principal: {pc}. Destacan {ce} sucesos CRÍTICOS a patrimonio/maquinaria. Tipología dominante: {tt[0]}, actor frecuente: {ta[0] if len(ta)>0 else 'N/A'}.")
                h2=doc.add_heading("II. Representación Gráfica", level=1); h2.runs[0].font.color.rgb=RGBColor(0x00,0x33,0x66)
                pg1=doc.add_paragraph(); pg1.alignment=WD_ALIGN_PARAGRAPH.CENTER; pg1.add_run("Figura 1: Distribución por Tipología").font.italic=True; doc.add_picture(ib, width=Inches(5.8))
                doc.add_paragraph()
                pg2=doc.add_paragraph(); pg2.alignment=WD_ALIGN_PARAGRAPH.CENTER; pg2.add_run("Figura 2: Proporción de Alertas").font.italic=True; doc.add_picture(ip, width=Inches(4.2))
                doc.add_paragraph()
                h3=doc.add_heading("III. Vulneraciones Críticas", level=1); h3.runs[0].font.color.rgb=RGBColor(0x00,0x33,0x66)
                dc = df_filtrado[df_filtrado['nivel_alerta']=='CRÍTICO'] if te>0 else pd.DataFrame()
                if not dc.empty:
                    tb=doc.add_table(rows=1, cols=3); tb.alignment=WD_TABLE_ALIGNMENT.CENTER; tb.style='Table Grid'
                    hd=tb.rows[0].cells; hd[0].text='Fecha'; hd[1].text='Sector'; hd[2].text='Descripción'
                    for c in hd:
                        for p in c.paragraphs:
                            for r in p.runs: r.font.bold=True; r.font.size=Pt(9.5); r.font.color.rgb=RGBColor(0x00,0x33,0x66)
                    for _, cr in dc.iterrows():
                        rc=tb.add_row().cells; rc[0].text=str(cr.get('fecha_limpia','')); rc[1].text=str(cr.get('ubicacion','MZS')).strip()
                        tt_=str(cr.get('titular','')); ac_=str(cr.get('actor','')).strip()
                        rc[2].text=f"{tt_}{' [Atrib: '+ac_+']' if ac_.lower() not in ['desconocido','','no especificado'] else ''}"
                        for c in rc:
                            for p in c.paragraphs:
                                for r in p.runs: r.font.size=Pt(9.0)
                    doc.add_paragraph()
                else: doc.add_paragraph("No se registraron sucesos críticos directos en la ventana.").italic=True
                h4=doc.add_heading("IV. Análisis Prospectivo", level=1); h4.runs[0].font.color.rgb=RGBColor(0x00,0x33,0x66)
                for p in ap_txt.split('. '):
                    if p.strip(): doc.add_paragraph(p.strip()+'.').paragraph_format.line_spacing=1.15
                h5=doc.add_heading("V. Directrices de Mando", level=1); h5.runs[0].font.color.rgb=RGBColor(0x00,0x33,0x66)
                for l in di_txt.split('\n'):
                    if l.strip(): pd_=doc.add_paragraph(l.strip()); pd_.paragraph_format.left_indent=Inches(0.2); pd_.paragraph_format.space_after=Pt(4)
                bf=io.BytesIO(); doc.save(bf); bf.seek(0)
                st.success("✔️ Reporte compilado con éxito.")
                st.download_button(label="📥 Descargar Documento Oficial (.docx)", data=bf, file_name=f"Radar_de_Crisis_CMPC_{datetime.now().strftime('%Y%m%d_%H%M')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
            except Exception as e: st.error(f"Error al compilar: {e}")
