import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, timedelta
import math
import re
import sqlite3
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Generador PDF oficial para Auditoría ENRE
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Cliente Supabase
try:
    from supabase import create_client
    SUPABASE_LIB_AVAILABLE = True
except ImportError:
    SUPABASE_LIB_AVAILABLE = False

# =============================================================================
# 1. CONFIGURACIÓN GENERAL Y ESTILOS
# =============================================================================
st.set_page_config(
    page_title="Gestión y Auditoría ENRE | Transnea Región Centro",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .banner-enre {
        background: linear-gradient(135deg, #062E1B 0%, #0E5E35 60%, #10B981 100%);
        padding: 18px 26px;
        border-radius: 14px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(7, 56, 32, 0.25);
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 15px;
    }
    .banner-title {
        font-size: 22px;
        font-weight: 800;
        margin: 0;
    }
    .banner-sub {
        font-size: 13.5px;
        color: #D1FAE5;
        margin-top: 4px;
    }
    .banner-badge {
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(52, 211, 153, 0.45);
        padding: 8px 16px;
        border-radius: 10px;
        font-size: 12px;
        color: #ECFDF5;
        text-align: right;
    }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# 2. CATÁLOGOS MAESTROS REGIÓN CENTRO Y PARÁMETROS AEA 95301
# =============================================================================
ANCHO_FRANJA_M = 22.0  # 11m izquierda + 11m derecha desde hilo de guardia (AEA 95301)
FACTOR_HA_POR_KM = ANCHO_FRANJA_M * 1000.0 / 10000.0  # 2.2 ha por cada km lineal

CATALOGO_LAT = {
    "ECO-SMA": {"tramo": "El Colorado (Formosa) / San Martín (Chaco)", "km": 30.47, "piquetes": 126, "color": "#E11D48"},
    "LLE-MAN": {"tramo": "Mansilla (Formosa) / La Leonesa (Chaco)", "km": 46.60, "piquetes": 162, "color": "#F97316"},
    "RES-LLE": {"tramo": "Resistencia (Chaco) / La Leonesa (Chaco)", "km": 62.80, "piquetes": 223, "color": "#EAB308"},
    "RES-BAR1": {"tramo": "Resistencia (Chaco) / Barranqueras 1 (Chaco)", "km": 30.20, "piquetes": 149, "color": "#10B981"},
    "RES-BAR2": {"tramo": "Resistencia (Chaco) / Barranqueras 2 (Chaco)", "km": 31.40, "piquetes": 154, "color": "#06B6D4"},
    "RES-CCE": {"tramo": "Resistencia (Chaco) / Corrientes Centro (Corrientes)", "km": 34.20, "piquetes": 100, "color": "#3B82F6"},
    "RES-SCA": {"tramo": "Resistencia (Chaco) / Santa Catalina (Corrientes)", "km": 42.00, "piquetes": 150, "color": "#8B5CF6"},
    "SCA-BVI": {"tramo": "Santa Catalina (Corrientes) / Bella Vista (Corrientes)", "km": 121.00, "piquetes": 463, "color": "#D946EF"},
    "BVI-GOY": {"tramo": "Bella Vista (Corrientes) / Goya (Corrientes)", "km": 77.70, "piquetes": 303, "color": "#EC4899"},
}

LINEAS_SIGLAS = list(CATALOGO_LAT.keys())
TOTAL_KM_RED = round(sum(d["km"] for d in CATALOGO_LAT.values()), 2)
TOTAL_PIQ_RED = sum(d["piquetes"] for d in CATALOGO_LAT.values())
TOTAL_HA_RED = round(TOTAL_KM_RED * FACTOR_HA_POR_KM, 2)

CUADRILLAS_CONFIG = {
    "Cuadrilla 1 - Desmonte y Limpieza": {
        "tareas": ["Limpieza de Electroducto / Desmonte (Franja 22m)"],
        "unidad_principal": "km"
    },
    "Cuadrilla 2 - Alturas y TCT": {
        "tareas": ["Perfilado de Aisladores", "Cambio de Aisladores (Correctivo)"],
        "unidad_principal": "piquetes"
    },
    "Cuadrilla 3 - Puesta a Tierra (PAT)": {
        "tareas": ["Medición de PAT", "Reparación de PAT"],
        "unidad_principal": "piquetes"
    },
    "Cuadrilla 4 - Recorrido e Inspección": {
        "tareas": ["Recorrido Visual Lento", "Recorrido Visual Rápido", "Recorrido con Dron"],
        "unidad_principal": "km"
    }
}

ESTADOS_TAREA = ["Programado", "En Ejecución", "Ejecutado", "Reprogramado", "Suspendido"]
COLORES_ESTADO = {
    "Programado": "#3B82F6",
    "En Ejecución": "#F59E0B",
    "Ejecutado": "#10B981",
    "Reprogramado": "#8B5CF6",
    "Suspendido": "#EF4444"
}

MESES_ES = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"
}

def natural_piq_key(piq_str):
    """Ordena piquetes numéricamente manteniendo intercalados en su lugar (ej: 111 -> 111b -> 112)."""
    s = str(piq_str).strip()
    m = re.match(r"^(\d+)(.*)$", s)
    if m:
        return (int(m.group(1)), m.group(2).strip().lower())
    return (999999, s.lower())

def normalizar_sigla_lat(nombre):
    if not nombre:
        return ""
    s = str(nombre).strip().upper().replace("#", "").replace("_", "-")
    if "ECO" in s and "SMA" in s:
        return "ECO-SMA"
    equivalencias = {
        "SMA-ECO": "ECO-SMA",
        "RES-BAR 1": "RES-BAR1",
        "RES-BAR 2": "RES-BAR2",
        "RES BAR 1": "RES-BAR1",
        "RES BAR 2": "RES-BAR2",
    }
    return equivalencias.get(s, s)

def distancia_geodesica_m(lat1, lon1, lat2, lon2):
    """Fórmula geodésica esférica idéntica a la hoja Auxiliar Distancia (en metros)."""
    try:
        r1 = math.radians(90.0 - float(lat1))
        r2 = math.radians(90.0 - float(lat2))
        dlon = math.radians(float(lon1) - float(lon2))
        cos_val = math.cos(r1) * math.cos(r2) + math.sin(r1) * math.sin(r2) * math.cos(dlon)
        cos_val = max(-1.0, min(1.0, cos_val))
        return 1000.0 * 6371.0 * math.acos(cos_val)
    except Exception:
        return 0.0

def corregir_par_lat_lon(lat_val, lon_val):
    """En NEA Argentina: Latitud está entre -22 y -32; Longitud está entre -54 y -63."""
    lat = float(lat_val)
    lon = float(lon_val)
    if lat < -45.0 and lon > -45.0:
        return lon, lat
    return lat, lon

def parsear_fecha_semana(f_val):
    try:
        s = str(f_val).strip()[:10]
        if len(s) == 10 and s[4] == "-":
            dt = date(int(s[:4]), int(s[5:7]), int(s[8:10]))
            iso = dt.isocalendar()
            return iso[0], iso[1]
        elif len(s) == 10 and s[2] == "/":
            dt = date(int(s[6:10]), int(s[3:5]), int(s[:2]))
            iso = dt.isocalendar()
            return iso[0], iso[1]
    except Exception:
        pass
    return None, None

def generar_mapa_semanas(anio):
    primer_dia = date(anio, 1, 1)
    inicio_s1 = primer_dia - timedelta(days=primer_dia.weekday())
    semanas = []
    for s in range(1, 53):
        f_ini = inicio_s1 + timedelta(weeks=s - 1)
        f_fin = f_ini + timedelta(days=6)
        f_jueves = f_ini + timedelta(days=3)
        mes_num = f_jueves.month if f_jueves.year == anio else (1 if s == 1 else 12)
        semanas.append({
            "semana": s,
            "etiqueta": f"S{s:02d}",
            "mes_num": mes_num,
            "mes_nombre": MESES_ES[mes_num],
            "rango_str": f"{f_ini.strftime('%d/%m')} - {f_fin.strftime('%d/%m')}"
        })
    return pd.DataFrame(semanas)

def crear_grafico_torta(titulo, ejecutado, total, unidad, color_ejec):
    pendiente = max(0.0, round(float(total) - float(ejecutado), 2))
    ejec_val = round(float(ejecutado), 2)
    pct = (ejec_val / total * 100) if total > 0 else 0.0

    fig = go.Figure(data=[go.Pie(
        labels=["Ejecutado", "Pendiente"],
        values=[ejec_val, pendiente],
        hole=0.58,
        marker=dict(colors=[color_ejec, "#1E293B"], line=dict(color="#0F172A", width=2)),
        textinfo="percent",
        textfont=dict(size=12, color="#FFFFFF"),
        hovertemplate=f"<b>%{{label}}</b><br>%{{value:.1f}} {unidad} (%{{percent}})<extra></extra>"
    )])
    fig.update_layout(
        template="plotly_dark",
        title=dict(text=f"<b>{titulo}</b>", font=dict(size=13, color="#E2E8F0"), x=0.5, xanchor="center"),
        annotations=[dict(
            text=f"<b>{pct:.1f}%</b><br><span style='font-size:10px;color:#94A3B8'>{ejec_val:g}/{total:g} {unidad}</span>",
            x=0.5, y=0.5, font_size=15, showarrow=False, font=dict(color="#FFFFFF")
        )],
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5, font=dict(size=11)),
        height=275,
        margin=dict(l=10, r=10, t=42, b=30),
        plot_bgcolor="#0F172A",
        paper_bgcolor="#0F172A"
    )
    return fig

# =============================================================================
# 3. PERSISTENCIA DE DATOS (PLAN ENRE + GPS PIQUETES + CAMPO REAL)
# =============================================================================
@st.cache_resource
def init_supabase():
    if SUPABASE_LIB_AVAILABLE and "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
        try:
            url = str(st.secrets["SUPABASE_URL"]).strip()
            key = str(st.secrets["SUPABASE_KEY"]).strip()
            return create_client(url, key)
        except Exception as e:
            st.error(f"Error conectando a Supabase: {e}")
            return None
    return None

sb_client = init_supabase()
MODO_NUBE = sb_client is not None

def get_sqlite_conn():
    conn = sqlite3.connect("mantenimiento_enre_transnea.db")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS plan_mantenimiento_enre (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            anio INTEGER,
            linea TEXT,
            cuadrilla TEXT,
            tarea TEXT,
            tramo_desc TEXT,
            semana_inicio INTEGER,
            semana_fin INTEGER,
            prog_km REAL DEFAULT 0.0,
            ejec_km REAL DEFAULT 0.0,
            prog_piq INTEGER DEFAULT 0,
            ejec_piq INTEGER DEFAULT 0,
            ot TEXT,
            estado TEXT,
            observaciones TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS piquetes_gps (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            linea TEXT NOT NULL,
            piquete TEXT NOT NULL,
            orden INTEGER NOT NULL,
            latitud REAL NOT NULL,
            longitud REAL NOT NULL,
            dist_km REAL DEFAULT 0.0,
            tipo TEXT DEFAULT '',
            estructura TEXT DEFAULT '',
            pct_desmonte REAL DEFAULT 0.0,
            UNIQUE(linea, piquete)
        )
    """)
    conn.commit()
    return conn

def db_fetch_plan():
    cols = [
        "id", "anio", "linea", "cuadrilla", "tarea", "tramo_desc",
        "semana_inicio", "semana_fin", "prog_km", "ejec_km",
        "prog_piq", "ejec_piq", "ot", "estado", "observaciones"
    ]
    if MODO_NUBE:
        try:
            res = sb_client.table("plan_mantenimiento_enre").select("*").order("semana_inicio", desc=False).execute()
            if not res.data:
                return pd.DataFrame(columns=cols)
            df = pd.DataFrame(res.data)
            for c in cols:
                if c not in df.columns:
                    df[c] = 0 if "km" in c or "piq" in c else None
            return df[cols]
        except Exception as e:
            st.warning(f"Aviso consultando plan_mantenimiento_enre: {e}")
            return pd.DataFrame(columns=cols)

    conn = get_sqlite_conn()
    df = pd.read_sql_query("SELECT * FROM plan_mantenimiento_enre ORDER BY semana_inicio ASC", conn)
    conn.close()
    return df

def db_fetch_gps():
    cols = ["id", "linea", "piquete", "orden", "latitud", "longitud", "dist_km", "tipo", "estructura", "pct_desmonte"]
    df = pd.DataFrame(columns=cols)
    if MODO_NUBE:
        try:
            all_rows = []
            offset = 0
            while True:
                r = sb_client.table("piquetes_gps").select("*").order("id", desc=False).range(offset, offset + 999).execute()
                if not r.data:
                    break
                all_rows.extend(r.data)
                if len(r.data) < 1000:
                    break
                offset += 1000
            if all_rows:
                df = pd.DataFrame(all_rows)
                for c in cols:
                    if c not in df.columns:
                        df[c] = 0.0 if c in ["dist_km", "pct_desmonte"] else ""
                df = df[cols]
        except Exception:
            return pd.DataFrame(columns=cols)
    else:
        conn = get_sqlite_conn()
        df = pd.read_sql_query("SELECT * FROM piquetes_gps", conn)
        conn.close()

    if not df.empty:
        # Aplicar orden natural por número de piquete para ubicar 3A, 44A, 111b en su posición física real
        df["_nat_key"] = df["piquete"].apply(natural_piq_key)
        df = df.sort_values(["linea", "_nat_key"]).drop(columns=["_nat_key"]).reset_index(drop=True)
        df["orden"] = df.groupby("linea").cumcount() + 1
    return df

def db_upsert_gps_batch(records):
    """Guarda o actualiza piquetes GPS masivamente (Supabase o SQLite)."""
    if not records:
        return
    if MODO_NUBE:
        try:
            for i in range(0, len(records), 500):
                sb_client.table("piquetes_gps").upsert(records[i:i+500], on_conflict="linea,piquete").execute()
            return
        except Exception as e:
            st.error(f"Error guardando coordenadas GPS en Supabase (verificá que creaste la tabla piquetes_gps): {e}")
            return

    conn = get_sqlite_conn()
    for r in records:
        conn.execute("""
            INSERT INTO piquetes_gps (linea, piquete, orden, latitud, longitud, dist_km, tipo, estructura, pct_desmonte)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(linea, piquete) DO UPDATE SET
                orden=excluded.orden,
                latitud=excluded.latitud,
                longitud=excluded.longitud,
                dist_km=excluded.dist_km,
                tipo=excluded.tipo,
                estructura=excluded.estructura,
                pct_desmonte=excluded.pct_desmonte
        """, (
            r["linea"], r["piquete"], int(r["orden"]), float(r["latitud"]), float(r["longitud"]),
            float(r.get("dist_km", 0.0)), str(r.get("tipo", "")), str(r.get("estructura", "")), float(r.get("pct_desmonte", 0.0))
        ))
    conn.commit()
    conn.close()

def db_update_pct_desmonte_linea(linea_sigla, mapa_pct_piq):
    """Actualiza el % de desmonte por piquete para una línea."""
    df_curr = db_fetch_gps()
    sub = df_curr[df_curr["linea"] == linea_sigla].copy()
    if sub.empty:
        return 0
    recs = []
    actualizados = 0
    for row in sub.itertuples():
        piq_s = str(row.piquete).strip()
        pct_val = mapa_pct_piq.get(piq_s, float(row.pct_desmonte or 0.0))
        if piq_s in mapa_pct_piq:
            actualizados += 1
        recs.append({
            "linea": row.linea,
            "piquete": piq_s,
            "orden": int(row.orden),
            "latitud": float(row.latitud),
            "longitud": float(row.longitud),
            "dist_km": float(row.dist_km or 0.0),
            "tipo": str(row.tipo or ""),
            "estructura": str(row.estructura or ""),
            "pct_desmonte": float(pct_val)
        })
    db_upsert_gps_batch(recs)
    return actualizados

@st.cache_data(ttl=60)
def fetch_datos_reales_supabase():
    df_perf = pd.DataFrame()
    df_pat = pd.DataFrame()
    if not MODO_NUBE:
        return df_perf, df_pat

    try:
        rows_perf = []
        offset = 0
        while True:
            r = sb_client.table("relevamientos").select("linea, piquete, fecha").range(offset, offset + 999).execute()
            if not r.data:
                break
            rows_perf.extend(r.data)
            if len(r.data) < 1000:
                break
            offset += 1000
        if rows_perf:
            df_perf = pd.DataFrame(rows_perf).drop_duplicates(subset=["linea", "piquete", "fecha"])
            df_perf["sigla"] = df_perf["linea"].apply(normalizar_sigla_lat)
            df_perf["piq_id"] = df_perf["piquete"].astype(str).str.strip()
            parsed = df_perf["fecha"].apply(parsear_fecha_semana)
            df_perf["anio"] = [p[0] for p in parsed]
            df_perf["semana"] = [p[1] for p in parsed]
    except Exception:
        pass

    try:
        rows_pat = []
        offset = 0
        while True:
            r = sb_client.table("mediciones_pat").select("*").range(offset, offset + 999).execute()
            if not r.data:
                break
            rows_pat.extend(r.data)
            if len(r.data) < 1000:
                break
            offset += 1000
        if rows_pat:
            df_pat = pd.DataFrame(rows_pat)
            col_lin = next((c for c in ["linea", "lat", "linea_lat"] if c in df_pat.columns), None)
            col_piq = next((c for c in ["piquete", "estructura", "nro_piquete"] if c in df_pat.columns), None)
            col_fec = next((c for c in ["fecha", "created_at", "fecha_medicion"] if c in df_pat.columns), None)

            if col_lin and col_piq:
                df_pat["sigla"] = df_pat[col_lin].apply(normalizar_sigla_lat)
                df_pat["piq_id"] = df_pat[col_piq].astype(str).str.strip()
                if col_fec:
                    parsed_p = df_pat[col_fec].apply(parsear_fecha_semana)
                    df_pat["anio"] = [p[0] for p in parsed_p]
                    df_pat["semana"] = [p[1] for p in parsed_p]
                else:
                    df_pat["anio"] = date.today().year
                    df_pat["semana"] = 1
                df_pat = df_pat.drop_duplicates(subset=["sigla", "piq_id", "anio"])
    except Exception:
        pass

    return df_perf, df_pat

def db_insert_registro(reg):
    if MODO_NUBE:
        try:
            sb_client.table("plan_mantenimiento_enre").insert(reg).execute()
            return
        except Exception as e:
            st.error(f"Error guardando en Supabase: {e}")
            return

    conn = get_sqlite_conn()
    conn.execute("""
        INSERT INTO plan_mantenimiento_enre
        (anio, linea, cuadrilla, tarea, tramo_desc, semana_inicio, semana_fin, prog_km, ejec_km, prog_piq, ejec_piq, ot, estado, observaciones)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        reg["anio"], reg["linea"], reg["cuadrilla"], reg["tarea"], reg["tramo_desc"],
        reg["semana_inicio"], reg["semana_fin"], reg["prog_km"], reg["ejec_km"],
        reg["prog_piq"], reg["ejec_piq"], reg["ot"], reg["estado"], reg["observaciones"]
    ))
    conn.commit()
    conn.close()

def db_update_avance(reg_id, estado, ejec_km, ejec_piq, s_ini, s_fin, obs):
    if MODO_NUBE:
        try:
            sb_client.table("plan_mantenimiento_enre").update({
                "estado": estado, "ejec_km": float(ejec_km), "ejec_piq": int(ejec_piq),
                "semana_inicio": int(s_ini), "semana_fin": int(s_fin), "observaciones": obs
            }).eq("id", int(reg_id)).execute()
            return
        except Exception as e:
            st.error(f"Error actualizando en Supabase: {e}")
            return

    conn = get_sqlite_conn()
    conn.execute("""
        UPDATE plan_mantenimiento_enre
        SET estado = ?, ejec_km = ?, ejec_piq = ?, semana_inicio = ?, semana_fin = ?, observaciones = ?
        WHERE id = ?
    """, (estado, float(ejec_km), int(ejec_piq), int(s_ini), int(s_fin), obs, int(reg_id)))
    conn.commit()
    conn.close()

def db_delete_registro(reg_id):
    if MODO_NUBE:
        try:
            sb_client.table("plan_mantenimiento_enre").delete().eq("id", int(reg_id)).execute()
            return
        except Exception as e:
            st.error(f"Error eliminando en Supabase: {e}")
            return

    conn = get_sqlite_conn()
    conn.execute("DELETE FROM plan_mantenimiento_enre WHERE id = ?", (int(reg_id),))
    conn.commit()
    conn.close()

def db_vaciar_plan_completo():
    if MODO_NUBE:
        try:
            sb_client.table("plan_mantenimiento_enre").delete().gte("id", 0).execute()
            return
        except Exception as e:
            st.error(f"Error vaciando tabla en Supabase: {e}")
            return

    conn = get_sqlite_conn()
    conn.execute("DELETE FROM plan_mantenimiento_enre")
    conn.commit()
    conn.close()

# =============================================================================
# 4. GENERADORES DE INFORMES DE AUDITORÍA ENRE (PDF Y EXCEL)
# =============================================================================
def generar_pdf_auditoria_enre(df_resumen_lineas, df_tareas, anio, horizonte_txt):
    if not REPORTLAB_AVAILABLE:
        return b""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle(
        "EnreTitle", parent=styles["Heading1"], fontName="Helvetica-Bold",
        fontSize=14, textColor=colors.HexColor("#0E5E35"), spaceAfter=4
    )
    sub_style = ParagraphStyle(
        "EnreSub", parent=styles["Normal"], fontName="Helvetica",
        fontSize=9, textColor=colors.HexColor("#334155"), spaceAfter=10
    )

    story.append(Paragraph("TRANSNEA S.A. — INFORME TÉCNICO DE MANTENIMIENTO LAT (AUDITORÍA ENRE)", title_style))
    story.append(Paragraph(
        f"<b>Región Centro</b> | Año Plan: {anio} | Horizonte: {horizonte_txt} | "
        f"Marco Normativo: <b>Reglamentación AEA 95301</b> (Franja de Servidumbre: 22 m totales — 11 m a cada lado del eje/hilo de guardia). "
        f"Fecha de emisión: {date.today().strftime('%d/%m/%Y')}",
        sub_style
    ))

    story.append(Paragraph("<b>1. INDICADORES DE LIMPIEZA DE ELECTRODUCTO (CUADRILLA 1) Y AVANCE REAL DE CAMPO POR LÍNEA</b>", styles["Normal"]))
    story.append(Spacer(1, 5))

    t1_data = [[
        "Línea", "Tramo Geográfico", "Long.\n(km)", "Franja\n(ha 22m)",
        "Desm. Prog\n(km)", "Desm. Ejec\n(km)", "Limpio\n(ha)",
        "% Desm.\nLínea", "Perf. Real\n(piq / %)", "PAT Real\n(piq / %)"
    ]]

    for r in df_resumen_lineas.itertuples():
        t1_data.append([
            r.Sigla,
            r.Tramo[:32],
            f"{r.Km_Total:.2f}",
            f"{r.Ha_Total:.1f}",
            f"{r.Desm_Prog_Km:.2f}",
            f"{r.Desm_Ejec_Km:.2f}",
            f"{r.Desm_Ejec_Ha:.1f}",
            f"{r.Pct_Linea_Limpia:.1f}%",
            f"{r.Perf_Real_Piq}/{r.Piquetes} ({r.Pct_Perf_Real:.0f}%)",
            f"{r.PAT_Real_Piq}/{r.Piquetes} ({r.Pct_PAT_Real:.0f}%)"
        ])

    t1 = Table(t1_data, colWidths=[52, 165, 46, 50, 54, 54, 48, 52, 82, 82])
    t1.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0E5E35")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.2),
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAF9")]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t1)
    story.append(Spacer(1, 12))

    if not df_tareas.empty:
        story.append(Paragraph("<b>2. CRONOGRAMA Y EJECUCIÓN PROGRAMADA POR CUADRILLA</b>", styles["Normal"]))
        story.append(Spacer(1, 5))

        t2_data = [["Línea", "Cuadrilla", "Actividad", "Semanas", "Prog.\n(km / piq)", "Ejec.\n(km / piq)", "OT", "Estado"]]
        for t in df_tareas.itertuples():
            prog_str = f"{t.prog_km:.1f} km" if t.prog_km > 0 else f"{t.prog_piq} piq"
            ejec_str = f"{t.ejec_km:.1f} km" if t.prog_km > 0 else f"{t.ejec_piq} piq"
            t2_data.append([
                t.linea,
                t.cuadrilla.split("-")[0].strip(),
                t.tarea[:34],
                f"S{t.semana_inicio:02d}-S{t.semana_fin:02d}",
                prog_str,
                ejec_str,
                t.ot if t.ot else "-",
                t.estado
            ])

        t2 = Table(t2_data, colWidths=[58, 72, 185, 58, 65, 65, 55, 75])
        t2.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("ALIGN", (3, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ]))
        story.append(t2)

    doc.build(story)
    return buf.getvalue()

def generar_excel_auditoria_enre(df_resumen_lineas, df_tareas, df_semanas_vista, anio, horizonte_txt):
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Indicadores_ENRE_Lineas"

    fill_green = PatternFill(start_color="0E5E35", end_color="0E5E35", fill_type="solid")
    fill_dark = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    font_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_head = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
    )

    ws1.merge_cells("A1:N1")
    ws1["A1"] = f"TRANSNEA S.A. - INDICADORES DE MANTENIMIENTO Y DESMONTE (AEA 95301 - FRANJA 22M) - AÑO {anio}"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_green
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    headers_h1 = [
        "Sigla LAT", "Tramo", "Longitud (km)", "Piquetes Totales", "Franja Total (ha)",
        "Desmonte Prog (km)", "Desmonte Ejec (km)", "Superficie Limpia (ha)",
        "% Cumpl. Desm s/ Prog", "% Desmonte s/ Línea",
        "Perfilado Real (piq)", "% Perfilado Línea",
        "PAT Real (piq)", "% PAT Línea"
    ]
    for c_i, h_txt in enumerate(headers_h1, start=1):
        cell = ws1.cell(row=3, column=c_i, value=h_txt)
        cell.font = font_head
        cell.fill = fill_dark
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for r_i, r in enumerate(df_resumen_lineas.itertuples(), start=4):
        vals = [
            r.Sigla, r.Tramo, r.Km_Total, r.Piquetes, r.Ha_Total,
            r.Desm_Prog_Km, r.Desm_Ejec_Km, r.Desm_Ejec_Ha,
            round(r.Pct_Cumpl_Prog, 2), round(r.Pct_Linea_Limpia, 2),
            r.Perf_Real_Piq, round(r.Pct_Perf_Real, 2),
            r.PAT_Real_Piq, round(r.Pct_PAT_Real, 2)
        ]
        for c_i, v in enumerate(vals, start=1):
            c = ws1.cell(row=r_i, column=c_i, value=v)
            c.border = thin_border
            if c_i >= 3:
                c.alignment = Alignment(horizontal="center")

    for col in ws1.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws1.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_len + 3, 14)

    ws2 = wb.create_sheet(title="Cronograma_Semanal")
    sems = df_semanas_vista["semana"].tolist()
    tot_cols = 6 + len(sems)
    last_col = openpyxl.utils.get_column_letter(tot_cols)

    ws2.merge_cells(f"A1:{last_col}1")
    ws2["A1"] = f"PROGRAMACIÓN SEMANAL DE CUADRILLAS - {horizonte_txt.upper()} ({anio})"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_green
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    fijos = ["Línea", "Cuadrilla", "Actividad", "Tramo", "Meta (km/piq)", "Estado"]
    for idx, f_h in enumerate(fijos, start=1):
        c = ws2.cell(row=3, column=idx, value=f_h)
        c.font = font_head
        c.fill = fill_dark
        c.alignment = Alignment(horizontal="center")

    for idx_s, s_row in enumerate(df_semanas_vista.itertuples(), start=7):
        c = ws2.cell(row=3, column=idx_s, value=f"{s_row.etiqueta}\n{s_row.mes_nombre}")
        c.font = Font(name="Calibri", size=8.5, bold=True, color="FFFFFF")
        c.fill = fill_green
        c.alignment = Alignment(horizontal="center", vertical="center")

    colores_xl = {"Programado": "3B82F6", "En Ejecución": "F59E0B", "Ejecutado": "10B981", "Reprogramado": "8B5CF6", "Suspendido": "EF4444"}
    for r_i, t in enumerate(df_tareas.itertuples(), start=4):
        meta_s = f"{t.prog_km} km" if t.prog_km > 0 else f"{t.prog_piq} piq"
        row_v = [t.linea, t.cuadrilla.split("-")[0].strip(), t.tarea, t.tramo_desc, meta_s, t.estado]
        for c_i, val in enumerate(row_v, start=1):
            cell = ws2.cell(row=r_i, column=c_i, value=val)
            cell.border = thin_border

        fill_c = PatternFill(start_color=colores_xl.get(t.estado, "3B82F6"), end_color=colores_xl.get(t.estado, "3B82F6"), fill_type="solid")
        for idx_s, s_num in enumerate(sems, start=7):
            cs = ws2.cell(row=r_i, column=idx_s)
            cs.border = thin_border
            if t.semana_inicio <= s_num <= t.semana_fin:
                cs.fill = fill_c
                cs.value = "■"
                cs.font = Font(color="FFFFFF", size=8)
                cs.alignment = Alignment(horizontal="center")

    anchos_f = [14, 16, 32, 24, 14, 14]
    for i, w in enumerate(anchos_f, start=1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for i in range(7, tot_cols + 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 5.0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# =============================================================================
# 5. ENCABEZADO Y SELECTOR DE HORIZONTE TEMPORAL
# =============================================================================
estado_cloud_txt = "☁️ Conectado a Supabase (Perfilado + PAT + GPS en vivo)" if MODO_NUBE else "⚠️ Modo Local (Configurar Secrets)"

st.markdown(f"""
<div class="banner-enre">
    <div>
        <h1 class="banner-title">⚡ TRANSNEA S.A. | Programación LAT, Mapa de Red e Indicadores ENRE</h1>
        <div class="banner-sub">Región Centro ({TOTAL_KM_RED} km • {TOTAL_PIQ_RED:,} Piquetes) | Reglamentación AEA 95301 (Franja 22m: 11m + 11m)</div>
    </div>
    <div class="banner-badge">
        <strong>Coordinación de Mantenimiento</strong><br>
        {estado_cloud_txt}
    </div>
</div>
""", unsafe_allow_html=True)

with st.container(border=True):
    ch1, ch2, ch3, ch4 = st.columns([1.1, 2.2, 2.3, 1.9])
    with ch1:
        anio_sel = st.selectbox("📅 Año Plan", [2026, 2027, 2025], index=0, key="top_anio_sel")
    with ch2:
        horizonte_sel = st.selectbox(
            "🔭 Horizonte de Planificación",
            [
                "12 Meses (Anual - 52 Semanas)",
                "6 Meses (Semestral - 26 Semanas)",
                "4 Meses (Cuatrimestral - 17 Semanas)",
                "3 Meses (Trimestral - 13 Semanas)",
                "1 Mes (Mensual - 4/5 Semanas)"
            ],
            index=0,
            key="top_horizonte_sel"
        )

    df_semanas_anio = generar_mapa_semanas(anio_sel)

    with ch3:
        if horizonte_sel.startswith("12 Meses"):
            s_desde, s_hasta = 1, 52
            st.info("Horizonte completo: **S01 a S52** (Ene – Dic)")
        elif horizonte_sel.startswith("6 Meses"):
            sem = st.radio("Semestre:", ["1° Semestre (Ene-Jun)", "2° Semestre (Jul-Dic)"], horizontal=True, key="top_semestre_sel")
            s_desde, s_hasta = (1, 26) if sem.startswith("1°") else (27, 52)
        elif horizonte_sel.startswith("4 Meses"):
            cua = st.selectbox("Cuatrimestre:", ["1° Cuatrimestre (Ene-Abr)", "2° Cuatrimestre (May-Ago)", "3° Cuatrimestre (Sep-Dic)"], key="top_cuatri_sel")
            s_desde, s_hasta = (1, 17) if cua.startswith("1°") else ((18, 35) if cua.startswith("2°") else (36, 52))
        elif horizonte_sel.startswith("3 Meses"):
            tri = st.selectbox("Trimestre:", ["Q1 (Ene-Mar)", "Q2 (Abr-Jun)", "Q3 (Jul-Sep)", "Q4 (Oct-Dic)"], key="top_trim_sel")
            s_desde, s_hasta = (1, 13) if tri.startswith("Q1") else ((14, 26) if tri.startswith("Q2") else ((27, 39) if tri.startswith("Q3") else (40, 52)))
        else:
            mes_sel = st.selectbox("Mes:", list(MESES_ES.keys()), format_func=lambda m: f"{MESES_ES[m]} {anio_sel}", index=datetime.now().month - 1, key="top_mes_sel")
            sems_m = df_semanas_anio[df_semanas_anio["mes_num"] == mes_sel]["semana"].tolist()
            s_desde, s_hasta = min(sems_m), max(sems_m)

    with ch4:
        filtro_cuadrilla = st.selectbox("👷 Filtrar Cuadrilla:", ["TODAS LAS CUADRILLAS"] + list(CUADRILLAS_CONFIG.keys()), key="top_cuad_sel")

df_semanas_vista = df_semanas_anio[(df_semanas_anio["semana"] >= s_desde) & (df_semanas_anio["semana"] <= s_hasta)].copy()
df_raw = db_fetch_plan()
df_gps_all = db_fetch_gps()
df_perf_real, df_pat_real = fetch_datos_reales_supabase()

tab_indicadores, tab_mapa, tab_cronograma, tab_cargar, tab_gestion = st.tabs([
    "📊 Indicadores y Auditoría ENRE",
    "🗺️ Mapa de Red",
    "🗓️ Cronograma Semanal",
    "➕ Programar Actividad",
    "✏️ Actualizar Avance / Limpiar Datos"
])

# =============================================================================
# PESTAÑA 1: INDICADORES, DIAGRAMAS DE TORTA Y BARRAS INTEGRANDO DATOS REALES
# =============================================================================
with tab_indicadores:
    df_anio = df_raw[df_raw["anio"] == anio_sel].copy() if not df_raw.empty else pd.DataFrame(columns=df_raw.columns)
    df_horizonte = df_anio[(df_anio["semana_fin"] >= s_desde) & (df_anio["semana_inicio"] <= s_hasta)].copy() if not df_anio.empty else df_anio

    perf_ventana = pd.DataFrame()
    if not df_perf_real.empty:
        perf_ventana = df_perf_real[
            (df_perf_real["anio"] == anio_sel) &
            (df_perf_real["semana"] >= s_desde) &
            (df_perf_real["semana"] <= s_hasta)
        ].drop_duplicates(subset=["sigla", "piq_id"])

    pat_ventana = pd.DataFrame()
    if not df_pat_real.empty:
        pat_ventana = df_pat_real[
            (df_pat_real["anio"] == anio_sel) &
            (df_pat_real["semana"] >= s_desde) &
            (df_pat_real["semana"] <= s_hasta)
        ].drop_duplicates(subset=["sigla", "piq_id"])

    df_desm = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 1 - Desmonte y Limpieza"] if not df_horizonte.empty else pd.DataFrame()
    km_desm_prog = df_desm["prog_km"].sum() if not df_desm.empty else 0.0
    km_desm_ejec = df_desm["ejec_km"].sum() if not df_desm.empty else 0.0
    ha_desm_ejec = round(km_desm_ejec * FACTOR_HA_POR_KM, 2)
    pct_cumpl_desm = (km_desm_ejec / km_desm_prog * 100) if km_desm_prog > 0 else 0.0
    pct_red_limpia = (km_desm_ejec / TOTAL_KM_RED * 100) if TOTAL_KM_RED > 0 else 0.0

    piq_perf_real_total = len(perf_ventana)
    df_c2 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 2 - Alturas y TCT"] if not df_horizonte.empty else pd.DataFrame()
    piq_aisl_camb = df_c2[df_c2["tarea"] == "Cambio de Aisladores (Correctivo)"]["ejec_piq"].sum() if not df_c2.empty else 0

    piq_pat_real_total = len(pat_ventana)
    df_c3 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 3 - Puesta a Tierra (PAT)"] if not df_horizonte.empty else pd.DataFrame()
    piq_pat_rep = df_c3[df_c3["tarea"] == "Reparación de PAT"]["ejec_piq"].sum() if not df_c3.empty else 0

    df_c4 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 4 - Recorrido e Inspección"] if not df_horizonte.empty else pd.DataFrame()
    km_rec_dron = df_c4[df_c4["tarea"] == "Recorrido con Dron"]["ejec_km"].sum() if not df_c4.empty else 0.0
    km_rec_vis = df_c4[df_c4["tarea"].str.contains("Visual")]["ejec_km"].sum() if not df_c4.empty else 0.0

    st.markdown("#### 🌲 Cuadrilla 1: Limpieza de Electroducto (Franja Servidumbre 22 m — AEA 95301)")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Desmonte Programado", f"{km_desm_prog:.2f} km", f"{km_desm_prog * FACTOR_HA_POR_KM:.1f} ha")
    k2.metric("Desmonte Ejecutado", f"{km_desm_ejec:.2f} km", f"{ha_desm_ejec:.1f} ha limpias")
    k3.metric("% Cumpl. s/ Programado", f"{pct_cumpl_desm:.1f}%", delta=f"{km_desm_ejec - km_desm_prog:.2f} km")
    k4.metric("% Avance s/ Red Total", f"{pct_red_limpia:.2f}%", f"de {TOTAL_KM_RED} km totales")
    k5.metric("Ancho Normativo AEA", "22 metros", "11m + 11m eje línea")

    st.markdown("#### 👷 Resumen Operativo Cuadrillas 2, 3 y 4 (Sincronizado con Campo)")
    c_kpi1, c_kpi2, c_kpi3, c_kpi4 = st.columns(4)
    c_kpi1.metric("Cuadrilla 2 (Perfilado Real)", f"{int(piq_perf_real_total)} piq.", f"{int(piq_aisl_camb)} cambios aislador")
    c_kpi2.metric("Cuadrilla 3 (PAT Real)", f"{int(piq_pat_real_total)} piq.", f"{int(piq_pat_rep)} PAT reparadas")
    c_kpi3.metric("Cuadrilla 4 (Dron)", f"{km_rec_dron:.2f} km", "Inspección aérea HD/IR")
    c_kpi4.metric("Cuadrilla 4 (Visual)", f"{km_rec_vis:.2f} km", "Recorrido lento/rápido")

    st.divider()

    st.markdown("#### 🍩 Distribución Porcentual Global — Región Centro")
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        st.plotly_chart(
            crear_grafico_torta("🌲 Desmonte Electroducto (22m)", km_desm_ejec, TOTAL_KM_RED, "km", "#10B981"),
            use_container_width=True, key="pie_desmonte"
        )
    with col_t2:
        st.plotly_chart(
            crear_grafico_torta("⚡ Perfilado Aisladores (Campo)", piq_perf_real_total, TOTAL_PIQ_RED, "piq", "#3B82F6"),
            use_container_width=True, key="pie_perfilado"
        )
    with col_t3:
        st.plotly_chart(
            crear_grafico_torta("🔌 Medición PAT (Campo)", piq_pat_real_total, TOTAL_PIQ_RED, "piq", "#F59E0B"),
            use_container_width=True, key="pie_pat"
        )

    st.divider()

    filas_lineas = []
    for sigla, info in CATALOGO_LAT.items():
        km_tot = info["km"]
        piq_tot = info["piquetes"]
        ha_tot = round(km_tot * FACTOR_HA_POR_KM, 2)

        sub_l_desm = df_desm[df_desm["linea"] == sigla] if not df_desm.empty else pd.DataFrame()
        d_prog = sub_l_desm["prog_km"].sum() if not sub_l_desm.empty else 0.0
        d_ejec = sub_l_desm["ejec_km"].sum() if not sub_l_desm.empty else 0.0
        d_ha = round(d_ejec * FACTOR_HA_POR_KM, 2)

        pct_prog = (d_ejec / d_prog * 100) if d_prog > 0 else 0.0
        pct_linea = (d_ejec / km_tot * 100) if km_tot > 0 else 0.0

        perf_l_real = int(perf_ventana[perf_ventana["sigla"] == sigla]["piq_id"].nunique()) if not perf_ventana.empty else 0
        pat_l_real = int(pat_ventana[pat_ventana["sigla"] == sigla]["piq_id"].nunique()) if not pat_ventana.empty else 0

        pct_perf_l = (perf_l_real / piq_tot * 100) if piq_tot > 0 else 0.0
        pct_pat_l = (pat_l_real / piq_tot * 100) if piq_tot > 0 else 0.0

        filas_lineas.append({
            "Sigla": sigla,
            "Tramo": info["tramo"],
            "Km_Total": km_tot,
            "Piquetes": piq_tot,
            "Ha_Total": ha_tot,
            "Desm_Prog_Km": d_prog,
            "Desm_Ejec_Km": d_ejec,
            "Desm_Ejec_Ha": d_ha,
            "Pct_Cumpl_Prog": pct_prog,
            "Pct_Linea_Limpia": pct_linea,
            "Perf_Real_Piq": perf_l_real,
            "Pct_Perf_Real": pct_perf_l,
            "PAT_Real_Piq": pat_l_real,
            "Pct_PAT_Real": pct_pat_l
        })

    df_resumen_lineas = pd.DataFrame(filas_lineas)

    col_g1, col_g2 = st.columns([1.15, 0.85])
    with col_g1:
        st.markdown("##### 📊 Avance Comparativo por Línea (% Desmonte vs. Perfilado Real vs. PAT Real)")
        fig_bar = go.Figure()

        fig_bar.add_trace(go.Bar(
            y=df_resumen_lineas["Sigla"],
            x=df_resumen_lineas["Pct_Linea_Limpia"],
            name="🌲 Desmonte (22m)",
            orientation="h",
            text=[f"{v:.1f}%" if v > 0 else "" for v in df_resumen_lineas["Pct_Linea_Limpia"]],
            textposition="outside",
            marker=dict(color="#10B981")
        ))

        fig_bar.add_trace(go.Bar(
            y=df_resumen_lineas["Sigla"],
            x=df_resumen_lineas["Pct_Perf_Real"],
            name="⚡ Perfilado Real",
            orientation="h",
            text=[f"{v:.1f}%" if v > 0 else "" for v in df_resumen_lineas["Pct_Perf_Real"]],
            textposition="outside",
            marker=dict(color="#3B82F6")
        ))

        fig_bar.add_trace(go.Bar(
            y=df_resumen_lineas["Sigla"],
            x=df_resumen_lineas["Pct_PAT_Real"],
            name="🔌 PAT Real",
            orientation="h",
            text=[f"{v:.1f}%" if v > 0 else "" for v in df_resumen_lineas["Pct_PAT_Real"]],
            textposition="outside",
            marker=dict(color="#F59E0B")
        ))

        fig_bar.update_layout(
            template="plotly_dark",
            barmode="group",
            plot_bgcolor="#0F172A", paper_bgcolor="#0F172A",
            xaxis=dict(range=[0, 118], title="% de Avance sobre Total de Línea", gridcolor="#334155"),
            yaxis=dict(autorange="reversed", gridcolor="#334155"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=450, margin=dict(l=20, r=35, t=35, b=20)
        )
        st.plotly_chart(fig_bar, use_container_width=True, key="bar_comparativo_lineas")

    with col_g2:
        st.markdown("##### 📋 Tabla Consolidada por Línea (Plan + Campo Real)")
        df_show_l = df_resumen_lineas[[
            "Sigla", "Km_Total", "Desm_Ejec_Km", "Pct_Linea_Limpia", "Perf_Real_Piq", "Pct_Perf_Real", "PAT_Real_Piq", "Pct_PAT_Real"
        ]].rename(columns={
            "Sigla": "Línea", "Km_Total": "km Tot", "Desm_Ejec_Km": "Desm km",
            "Pct_Linea_Limpia": "% Desm", "Perf_Real_Piq": "Perf Piq",
            "Pct_Perf_Real": "% Perf", "PAT_Real_Piq": "PAT Piq", "Pct_PAT_Real": "% PAT"
        })
        st.dataframe(
            df_show_l.style.format({
                "km Tot": "{:.1f}", "Desm km": "{:.1f}", "% Desm": "{:.1f}%",
                "Perf Piq": "{:,}", "% Perf": "{:.1f}%", "PAT Piq": "{:,}", "% PAT": "{:.1f}%"
            }),
            use_container_width=True, hide_index=True, height=450
        )

    st.divider()
    st.markdown("#### 🏛️ Exportación de Informes Oficiales para Auditoría ENRE (AEA 95301)")
    col_ex1, col_ex2 = st.columns(2)
    horizonte_corto = horizonte_sel.split("(")[0].strip()

    with col_ex1:
        xl_bytes = generar_excel_auditoria_enre(df_resumen_lineas, df_horizonte, df_semanas_vista, anio_sel, horizonte_corto)
        st.download_button(
            label="📊 Descargar Matriz e Indicadores ENRE en Excel (.xlsx)",
            data=xl_bytes,
            file_name=f"Auditoria_ENRE_Transnea_RegionCentro_{anio_sel}_S{s_desde:02d}_S{s_hasta:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_excel_enre"
        )

    with col_ex2:
        if REPORTLAB_AVAILABLE:
            pdf_bytes = generar_pdf_auditoria_enre(df_resumen_lineas, df_horizonte, anio_sel, f"{horizonte_corto} (S{s_desde:02d} a S{s_hasta:02d})")
            st.download_button(
                label="📄 Descargar Informe Técnico Oficial ENRE en PDF (.pdf)",
                data=pdf_bytes,
                file_name=f"Informe_ENRE_AEA95301_Transnea_{anio_sel}_S{s_desde:02d}_S{s_hasta:02d}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
                key="dl_pdf_enre"
            )
        else:
            st.info("Instalá 'reportlab' en requirements.txt para habilitar la exportación PDF.")

# =============================================================================
# PESTAÑA 2: MAPA DE RED (REGIÓN OPERATIVA CENTRO - COLOR POR LÍNEA + GPS)
# =============================================================================
with tab_mapa:
    st.markdown("### 🗺️ Mapa de Red Georreferenciado — Región Operativa Centro")
    st.caption("Visualización diferenciada por color de las 9 líneas LAT de Región Centro (Chaco, Corrientes y Formosa), con ordenamiento natural de piquetes intercalados y cálculo de superficie.")

    cm1, cm2, cm3, cm4 = st.columns([2, 1.5, 1.5, 2])
    with cm1:
        mapa_linea_sel = st.selectbox(
            "⚡ Seleccionar Línea LAT:",
            options=["TODAS (REGIÓN CENTRO MULTICOLOR)"] + LINEAS_SIGLAS,
            key="map_linea_sel"
        )
    with cm2:
        ancho_serv_sel = st.selectbox(
            "📏 Ancho Franja Servidumbre:",
            options=[22.0, 25.0],
            format_func=lambda w: f"{w:.0f} m (AEA 95301: 11+11m)" if w == 22.0 else f"{w:.0f} m (Salvoconducto: 12.5+12.5m)",
            key="map_ancho_serv"
        )
    with cm3:
        capa_operativa = st.selectbox(
            "🎨 Capa Operativa sobre Piquetes:",
            options=["Color Distintivo por Línea", "Avance Desmonte (%)", "Perfilado Real (Campo)", "PAT Medida Real (Campo)"],
            key="map_capa_op"
        )
    with cm4:
        estilo_mapa = st.selectbox(
            "🌍 Fondo Cartográfico:",
            options=["open-street-map", "carto-darkmatter", "carto-positron"],
            format_func=lambda s: {"open-street-map": "Rutas y Localidades (OSM)", "carto-darkmatter": "Oscuro Técnico", "carto-positron": "Claro Catastral"}[s],
            key="map_estilo_fondo"
        )

    if df_gps_all.empty:
        st.warning("⚠️ Aún no hay coordenadas GPS cargadas en la base de datos. Subí tu archivo **`Base de datos piquetes.xlsm`** (o `piquetes_region_centro_gps.csv`) en el importador de abajo para cargar los 1.830 piquetes de la Región Centro.")
    else:
        df_map = df_gps_all.copy() if mapa_linea_sel.startswith("TODAS") else df_gps_all[df_gps_all["linea"] == mapa_linea_sel].copy()
        df_map["_nat"] = df_map["piquete"].apply(natural_piq_key)
        df_map = df_map.sort_values(["linea", "_nat"]).drop(columns=["_nat"]).reset_index(drop=True)

        dist_m_list = []
        for sig, grp in df_map.groupby("linea", sort=False):
            lats = grp["latitud"].tolist()
            lons = grp["longitud"].tolist()
            for i in range(len(grp)):
                if i < len(grp) - 1:
                    d_m = distancia_geodesica_m(lats[i], lons[i], lats[i+1], lons[i+1])
                else:
                    d_m = float(grp.iloc[i]["dist_km"] or 0.0) * 1000.0
                dist_m_list.append(d_m)

        df_map["dist_m"] = dist_m_list
        df_map["area_m2"] = df_map["dist_m"] * ancho_serv_sel
        df_map["hectareas"] = df_map["area_m2"] / 10000.0
        df_map["ha_limpias"] = df_map["hectareas"] * (df_map["pct_desmonte"].fillna(0.0) / 100.0)

        set_perf = set(zip(perf_ventana["sigla"], perf_ventana["piq_id"])) if not perf_ventana.empty else set()
        set_pat = set(zip(pat_ventana["sigla"], pat_ventana["piq_id"])) if not pat_ventana.empty else set()

        df_map["tiene_perf"] = [((r.linea, str(r.piquete).strip()) in set_perf) for r in df_map.itertuples()]
        df_map["tiene_pat"] = [((r.linea, str(r.piquete).strip()) in set_pat) for r in df_map.itertuples()]

        km_vista = df_map["dist_m"].sum() / 1000.0
        m2_vista = df_map["area_m2"].sum()
        ha_vista = df_map["hectareas"].sum()
        ha_limpias_vista = df_map["ha_limpias"].sum()
        pct_limpio_vista = (ha_limpias_vista / ha_vista * 100) if ha_vista > 0 else 0.0

        mk1, mk2, mk3, mk4, mk5 = st.columns(5)
        mk1.metric("Piquetes Georreferenciados", f"{len(df_map):,}")
        mk2.metric("Longitud GPS Calculada", f"{km_vista:.2f} km")
        mk3.metric(f"Superficie Electroducto ({ancho_serv_sel:.0f}m)", f"{m2_vista:,.0f} m²", f"{ha_vista:.2f} ha totales")
        mk4.metric("Superficie Limpia (Desmonte)", f"{ha_limpias_vista:.2f} ha", f"{pct_limpio_vista:.1f}% limpio")
        mk5.metric("Piquetes Campo (Perf / PAT)", f"{int(df_map['tiene_perf'].sum())} / {int(df_map['tiene_pat'].sum())}")

        # Construcción del Mapa Interactivo con colores únicos por línea
        fig_map = go.Figure()

        for sig, grp in df_map.groupby("linea", sort=False):
            color_linea = CATALOGO_LAT.get(sig, {}).get("color", "#D946EF")

            hover_txts = [
                f"<b>LAT {r.linea} — Piquete {r.piquete}</b><br>"
                f"Tipo: {r.tipo} | Estructura: {r.estructura}<br>"
                f"GPS: {r.latitud:.6f}, {r.longitud:.6f}<br>"
                f"Vano siguiente: {r.dist_m:.1f} m ({r.hectareas:.3f} ha)<br>"
                f"Desmonte vano: {r.pct_desmonte:.0f}% ({r.ha_limpias:.3f} ha limpias)<br>"
                f"Perfilado: {'✅ Sí' if r.tiene_perf else '⏳ Pendiente'} | PAT: {'✅ Sí' if r.tiene_pat else '⏳ Pendiente'}"
                for r in grp.itertuples()
            ]

            fig_map.add_trace(go.Scattermap(
                lat=grp["latitud"],
                lon=grp["longitud"],
                mode="lines",
                line=dict(width=4, color=color_linea),
                name=f"LAT {sig}",
                hoverinfo="skip"
            ))

            if capa_operativa == "Avance Desmonte (%)":
                colores_pts = ["#10B981" if p >= 80 else ("#F59E0B" if p > 0 else color_linea) for p in grp["pct_desmonte"]]
            elif capa_operativa == "Perfilado Real (Campo)":
                colores_pts = ["#3B82F6" if ok else color_linea for ok in grp["tiene_perf"]]
            elif capa_operativa == "PAT Medida Real (Campo)":
                colores_pts = ["#F59E0B" if ok else color_linea for ok in grp["tiene_pat"]]
            else:
                colores_pts = color_linea

            fig_map.add_trace(go.Scattermap(
                lat=grp["latitud"],
                lon=grp["longitud"],
                mode="markers",
                marker=dict(size=7.5 if not mapa_linea_sel.startswith("TODAS") else 6, color=colores_pts),
                name=f"Piquetes {sig}",
                text=hover_txts,
                hovertemplate="%{text}<extra></extra>",
                showlegend=False
            ))

        centro_lat = float(df_map["latitud"].mean())
        centro_lon = float(df_map["longitud"].mean())
        zoom_val = 7.2 if mapa_linea_sel.startswith("TODAS") else 9.5

        fig_map.update_layout(
            map=dict(
                style=estilo_mapa,
                center=dict(lat=centro_lat, lon=centro_lon),
                zoom=zoom_val
            ),
            height=540,
            margin=dict(l=0, r=0, t=10, b=0),
            legend=dict(
                orientation="h", yanchor="top", y=0.99, xanchor="left", x=0.01,
                bgcolor="rgba(15, 23, 42, 0.88)", font=dict(color="#FFFFFF", size=11)
            )
        )
        st.plotly_chart(fig_map, use_container_width=True, key="mapa_red_region_centro")

    st.divider()

    # =========================================================================
    # SECCIÓN DE GESTIÓN GPS: 1) DECLARAR PIQUETE | 2) IMPORTAR EXCEL | 3) CALCULAR TRAMO
    # =========================================================================
    subtab_decl, subtab_imp, subtab_calc = st.tabs([
        "📍 Declarar / Editar Coordenadas GPS",
        "📥 Importar Base Piquetes / Planilla Desmonte (.xlsm / .xlsx)",
        "📐 Calculadora de Desmonte por Tramo (m² y Hectáreas)"
    ])

    with subtab_decl:
        st.markdown("##### 📍 Declarar Coordenadas GPS de Piquete (Latitud y Longitud)")
        with st.container(border=True):
            dc1, dc2, dc3, dc4 = st.columns([1.8, 1.2, 1.2, 1.5])
            with dc1:
                d_linea = st.selectbox("Línea LAT (Región Centro)", LINEAS_SIGLAS, key="decl_linea")
            with dc2:
                d_piq = st.text_input("Piquete N°", placeholder="Ej: 1 o 111b", key="decl_piq")
            with dc3:
                d_orden = st.number_input("Orden Secuencial", min_value=1, max_value=1000, value=1, step=1, key="decl_orden")
            with dc4:
                d_tipo = st.selectbox("Tipo Estructura", ["SS", "RT", "SD"], key="decl_tipo")

            dc5, dc6, dc7, dc8 = st.columns([1.5, 1.5, 1.5, 1.5])
            with dc5:
                d_lat = st.number_input("Latitud GPS (Sur)", min_value=-35.0, max_value=-20.0, value=-27.450000, format="%.6f", step=0.0001, key="decl_lat")
            with dc6:
                d_lon = st.number_input("Longitud GPS (Oeste)", min_value=-65.0, max_value=-54.0, value=-58.980000, format="%.6f", step=0.0001, key="decl_lon")
            with dc7:
                d_est = st.selectbox("Material Poste", ["Hormigón", "Metálica", "Madera"], key="decl_est")
            with dc8:
                d_pct = st.number_input("% Desmonte Vano", min_value=0.0, max_value=100.0, value=0.0, step=5.0, key="decl_pct")

            dist_calc_km = 0.0
            if not df_gps_all.empty:
                sub_prev = df_gps_all[(df_gps_all["linea"] == d_linea) & (df_gps_all["orden"] == int(d_orden) - 1)]
                if not sub_prev.empty:
                    lat_p = float(sub_prev.iloc[0]["latitud"])
                    lon_p = float(sub_prev.iloc[0]["longitud"])
                    d_m_prev = distancia_geodesica_m(lat_p, lon_p, d_lat, d_lon)
                    dist_calc_km = d_m_prev / 1000.0
                    area_p = d_m_prev * ancho_serv_sel
                    st.info(f"📏 Distancia geodésica desde el piquete anterior (Orden {int(d_orden)-1}): **{d_m_prev:.2f} m** | Superficie: **{area_p:,.1f} m²** (**{area_p/10000.0:.4f} ha**).")

            if st.button("💾 Guardar Coordenada GPS del Piquete", type="primary", use_container_width=True, key="btn_save_gps_single"):
                if not d_piq.strip():
                    st.error("Ingresá el número o identificador del piquete.")
                else:
                    lat_ok, lon_ok = corregir_par_lat_lon(d_lat, d_lon)
                    db_upsert_gps_batch([{
                        "linea": d_linea,
                        "piquete": d_piq.strip(),
                        "orden": int(d_orden),
                        "latitud": round(lat_ok, 6),
                        "longitud": round(lon_ok, 6),
                        "dist_km": round(dist_calc_km, 5),
                        "tipo": d_tipo,
                        "estructura": d_est,
                        "pct_desmonte": float(d_pct)
                    }])
                    st.success(f"✅ Coordenada del Piquete {d_piq} ({d_linea}) guardada correctamente.")
                    st.rerun()

    with subtab_imp:
        col_i1, col_i2 = st.columns(2)

        with col_i1:
            st.markdown("##### 📂 1. Importar Coordenadas (`Base de datos piquetes.xlsm` o `.csv`)")
            st.caption("Importa automáticamente los 1.830 piquetes de las 9 líneas de la Región Centro con sus coordenadas GPS y ordenamiento natural.")
            file_bd_gps = st.file_uploader("Subir Base de datos piquetes (.xlsm / .xlsx / .csv)", type=["xlsm", "xlsx", "csv"], key="upl_bd_piquetes")

            if file_bd_gps is not None:
                try:
                    if file_bd_gps.name.lower().endswith(".csv"):
                        df_in = pd.read_csv(file_bd_gps)
                    else:
                        df_raw_xl = pd.read_excel(file_bd_gps, sheet_name="Base de datos", header=1)
                        df_raw_xl["linea_norm"] = df_raw_xl["LAT"].apply(normalizar_sigla_lat)
                        df_sub_rc = df_raw_xl[df_raw_xl["linea_norm"].isin(LINEAS_SIGLAS)].copy()
                        df_in = pd.DataFrame({
                            "linea": df_sub_rc["linea_norm"],
                            "piquete": df_sub_rc["PIQ N°"].astype(str).str.strip(),
                            "latitud": pd.to_numeric(df_sub_rc["LAT.1"], errors="coerce"),
                            "longitud": pd.to_numeric(df_sub_rc["LONG"], errors="coerce"),
                            "dist_km": pd.to_numeric(df_sub_rc["Entre piq."], errors="coerce").fillna(0.0),
                            "tipo": df_sub_rc["TIPO"].fillna("").astype(str).str.strip(),
                            "estructura": df_sub_rc["ESTRUCTURA"].fillna("").astype(str).str.strip(),
                        }).dropna(subset=["latitud", "longitud"])

                    df_in["_nat"] = df_in["piquete"].apply(natural_piq_key)
                    df_in = df_in.sort_values(["linea", "_nat"]).drop(columns=["_nat"]).reset_index(drop=True)
                    df_in["orden"] = df_in.groupby("linea").cumcount() + 1

                    st.success(f"✅ Detectados **{len(df_in):,} piquetes** válidos de la Región Centro.")
                    if st.button("🚀 Cargar Coordenadas al Mapa de Red", type="primary", use_container_width=True, key="btn_imp_gps_bd"):
                        recs_imp = []
                        for r in df_in.itertuples():
                            lat_c, lon_c = corregir_par_lat_lon(r.latitud, r.longitud)
                            recs_imp.append({
                                "linea": r.linea,
                                "piquete": str(r.piquete).strip(),
                                "orden": int(r.orden),
                                "latitud": round(lat_c, 6),
                                "longitud": round(lon_c, 6),
                                "dist_km": round(float(r.dist_km), 5),
                                "tipo": str(r.tipo),
                                "estructura": str(r.estructura),
                                "pct_desmonte": 0.0
                            })
                        with st.spinner("Guardando 1.830 coordenadas GPS..."):
                            db_upsert_gps_batch(recs_imp)
                        st.toast("¡Base GPS de Región Centro cargada con éxito!", icon="🗺️")
                        st.rerun()
                except Exception as e:
                    st.error(f"Error leyendo archivo de coordenadas: {e}")

        with col_i2:
            st.markdown("##### 🌲 2. Importar Planilla de Desmonte (`DESMONTE LAT *.xlsx`)")
            st.caption("Lee los porcentajes por piquete de 'Porcentaje Piquetes', calcula distancias desde GPS, superficie en m² y hectáreas limpias.")
            linea_desm_xl = st.selectbox("Línea LAT correspondiente a la planilla:", LINEAS_SIGLAS, key="imp_desm_linea")
            file_desm_xl = st.file_uploader("Subir Excel de Desmonte (.xlsx)", type=["xlsx", "xlsm"], key="upl_desmonte_xl")

            if file_desm_xl is not None:
                try:
                    df_pct_xl = pd.read_excel(file_desm_xl, sheet_name="Porcentaje Piquetes", header=1)
                    col_p = df_pct_xl.columns[0]
                    df_pct_xl = df_pct_xl.dropna(subset=[col_p]).copy()

                    mapa_pct = {}
                    for r in df_pct_xl.itertuples():
                        try:
                            p_str = str(int(float(r[1])))
                        except Exception:
                            p_str = str(r[1]).strip()
                        val_raw = float(r[2]) if pd.notnull(r[2]) else 0.0
                        pct_100 = val_raw * 100.0 if val_raw <= 1.0 else val_raw
                        mapa_pct[p_str] = round(pct_100, 2)

                    sub_gps_l = df_gps_all[df_gps_all["linea"] == linea_desm_xl].copy() if not df_gps_all.empty else pd.DataFrame()
                    if sub_gps_l.empty:
                        st.warning(f"Primero cargá las coordenadas GPS de **{linea_desm_xl}** en el panel izquierdo.")
                    else:
                        lats = sub_gps_l["latitud"].tolist()
                        lons = sub_gps_l["longitud"].tolist()
                        piqs = sub_gps_l["piquete"].tolist()
                        tot_m = 0.0
                        limpio_m = 0.0
                        for i in range(len(sub_gps_l) - 1):
                            dm = distancia_geodesica_m(lats[i], lons[i], lats[i+1], lons[i+1])
                            pct_i = mapa_pct.get(str(piqs[i]).strip(), 0.0) / 100.0
                            tot_m += dm
                            limpio_m += dm * pct_i

                        area_limpia_m2 = limpio_m * ancho_serv_sel
                        ha_limpia = area_limpia_m2 / 10000.0
                        pct_avance_tot = (limpio_m / tot_m * 100) if tot_m > 0 else 0.0

                        st.info(
                            f"**Resultado Cálculo GPS ({linea_desm_xl} — Franja {ancho_serv_sel:.0f}m):**\n"
                            f"- Distancia limpia: **{limpio_m/1000.0:.2f} km** de {tot_m/1000.0:.2f} km (**{pct_avance_tot:.1f}%**)\n"
                            f"- Superficie limpia: **{area_limpia_m2:,.0f} m²** (**{ha_limpia:.2f} ha**)"
                        )

                        if st.button("💾 Aplicar Avance al Mapa y Actualizar Plan ENRE", type="primary", use_container_width=True, key="btn_apply_desm_xl"):
                            db_update_pct_desmonte_linea(linea_desm_xl, mapa_pct)
                            db_insert_registro({
                                "anio": int(anio_sel),
                                "linea": linea_desm_xl,
                                "cuadrilla": "Cuadrilla 1 - Desmonte y Limpieza",
                                "tarea": "Limpieza de Electroducto / Desmonte (Franja 22m)",
                                "tramo_desc": f"Relevamiento GPS ({pct_avance_tot:.1f}% traza)",
                                "semana_inicio": int(s_desde),
                                "semana_fin": int(s_hasta),
                                "prog_km": round(tot_m / 1000.0, 2),
                                "ejec_km": round(limpio_m / 1000.0, 2),
                                "prog_piq": len(sub_gps_l),
                                "ejec_piq": sum(1 for v in mapa_pct.values() if v > 0),
                                "ot": "GPS-DESM",
                                "estado": "En Ejecución" if pct_avance_tot < 99.9 else "Ejecutado",
                                "observaciones": f"Calculado por GPS ({ancho_serv_sel:.0f}m): {ha_limpia:.2f} ha limpias."
                            })
                            st.toast(f"¡Desmonte de {linea_desm_xl} actualizado en Mapa e Indicadores!", icon="🌲")
                            st.rerun()
                except Exception as e:
                    st.error(f"Error procesando planilla de desmonte: {e}")

    with subtab_calc:
        st.markdown("##### 📐 Calculadora de Distancia, Superficie (m²) y Hectáreas entre Piquetes")
        if df_gps_all.empty:
            st.info("Cargá primero las coordenadas GPS en la pestaña anterior para usar la calculadora.")
        else:
            cc1, cc2, cc3, cc4 = st.columns([2, 1.3, 1.3, 1.4])
            with cc1:
                calc_linea = st.selectbox("Línea LAT:", LINEAS_SIGLAS, key="calc_tramo_linea")
            sub_c = df_gps_all[df_gps_all["linea"] == calc_linea].reset_index(drop=True)
            lista_piqs_c = sub_c["piquete"].tolist()

            with cc2:
                piq_ini_c = st.selectbox("Piquete Inicio:", lista_piqs_c, index=0, key="calc_piq_ini")
            with cc3:
                piq_fin_c = st.selectbox("Piquete Fin:", lista_piqs_c, index=len(lista_piqs_c) - 1, key="calc_piq_fin")
            with cc4:
                pct_tramo_c = st.number_input("% Limpieza Tramo:", min_value=0.0, max_value=100.0, value=100.0, step=5.0, key="calc_pct_tramo")

            idx1 = lista_piqs_c.index(piq_ini_c)
            idx2 = lista_piqs_c.index(piq_fin_c)
            if idx1 > idx2:
                idx1, idx2 = idx2, idx1

            tramo_df = sub_c.iloc[idx1:idx2+1].reset_index(drop=True)
            dist_tramo_m = 0.0
            for i in range(len(tramo_df) - 1):
                dist_tramo_m += distancia_geodesica_m(
                    tramo_df.loc[i, "latitud"], tramo_df.loc[i, "longitud"],
                    tramo_df.loc[i+1, "latitud"], tramo_df.loc[i+1, "longitud"]
                )

            dist_ejec_m = dist_tramo_m * (pct_tramo_c / 100.0)
            sup_total_m2 = dist_tramo_m * ancho_serv_sel
            sup_ejec_m2 = dist_ejec_m * ancho_serv_sel
            ha_total_tramo = sup_total_m2 / 10000.0
            ha_ejec_tramo = sup_ejec_m2 / 10000.0
            km_linea_cat = CATALOGO_LAT[calc_linea]["km"]
            pct_sobre_linea = ((dist_ejec_m / 1000.0) / km_linea_cat * 100) if km_linea_cat > 0 else 0.0

            rk1, rk2, rk3, rk4 = st.columns(4)
            rk1.metric("Distancia Tramo", f"{dist_tramo_m:,.1f} m", f"{dist_tramo_m/1000.0:.2f} km")
            rk2.metric(f"Superficie ({ancho_serv_sel:.0f}m)", f"{sup_ejec_m2:,.0f} m²", f"de {sup_total_m2:,.0f} m²")
            rk3.metric("Hectáreas Limpias", f"{ha_ejec_tramo:.2f} ha", f"de {ha_total_tramo:.2f} ha")
            rk4.metric("% Avance s/ Línea Total", f"{pct_sobre_linea:.2f}%", f"sobre {km_linea_cat} km")

# =============================================================================
# PESTAÑA 3: CRONOGRAMA SEMANAL (PLANIFICADO + EJECUCIÓN REAL DE CAMPO)
# =============================================================================
with tab_cronograma:
    incluir_reales_gantt = st.checkbox("🔗 Mostrar también ejecución real registrada en Perfilado y PAT en el cronograma", value=True, key="chk_reales_gantt")

    filas_gantt = []

    if not df_raw.empty:
        df_crono = df_raw[(df_raw["anio"] == anio_sel) & (df_raw["semana_fin"] >= s_desde) & (df_raw["semana_inicio"] <= s_hasta)].copy()
        if filtro_cuadrilla != "TODAS LAS CUADRILLAS":
            df_crono = df_crono[df_crono["cuadrilla"] == filtro_cuadrilla]

        for r in df_crono.itertuples():
            c_short = r.cuadrilla.split("-")[0].strip()
            meta_txt = f"{r.ejec_km}/{r.prog_km} km" if r.prog_km > 0 else f"{r.ejec_piq}/{r.prog_piq} piq"
            filas_gantt.append({
                "linea": r.linea,
                "label": f"<b>{r.linea}</b> | {c_short}: {r.tarea[:22]}",
                "s_ini": int(r.semana_inicio),
                "s_fin": int(r.semana_fin),
                "estado": r.estado,
                "texto_barra": f"S{r.semana_inicio:02d}-S{r.semana_fin:02d} ({meta_txt})"
            })

    if incluir_reales_gantt:
        if (filtro_cuadrilla in ["TODAS LAS CUADRILLAS", "Cuadrilla 2 - Alturas y TCT"]) and not perf_ventana.empty:
            for sig, grp in perf_ventana.groupby("sigla"):
                if sig in CATALOGO_LAT:
                    s_min = int(grp["semana"].min())
                    s_max = int(grp["semana"].max())
                    cant_p = int(grp["piq_id"].nunique())
                    filas_gantt.append({
                        "linea": sig,
                        "label": f"<b>{sig}</b> | [CAMPO] Perfilado Real",
                        "s_ini": s_min,
                        "s_fin": s_max,
                        "estado": "Ejecutado",
                        "texto_barra": f"S{s_min:02d}-S{s_max:02d} ({cant_p} piq. reales)"
                    })

        if (filtro_cuadrilla in ["TODAS LAS CUADRILLAS", "Cuadrilla 3 - Puesta a Tierra (PAT)"]) and not pat_ventana.empty:
            for sig, grp in pat_ventana.groupby("sigla"):
                if sig in CATALOGO_LAT:
                    s_min = int(grp["semana"].min())
                    s_max = int(grp["semana"].max())
                    cant_p = int(grp["piq_id"].nunique())
                    filas_gantt.append({
                        "linea": sig,
                        "label": f"<b>{sig}</b> | [CAMPO] Medición PAT Real",
                        "s_ini": s_min,
                        "s_fin": s_max,
                        "estado": "Ejecutado",
                        "texto_barra": f"S{s_min:02d}-S{s_max:02d} ({cant_p} piq. reales)"
                    })

    if not filas_gantt:
        st.info(f"No hay actividades programadas ni ejecuciones reales entre S{s_desde:02d} y S{s_hasta:02d} para este filtro.")
    else:
        df_gantt_all = pd.DataFrame(filas_gantt).sort_values(by=["s_ini", "linea"], ascending=[False, True])
        st.markdown(f"#### 🗓️ Diagrama Gantt Semanal ({horizonte_sel.split('(')[0].strip()} • S{s_desde:02d} a S{s_hasta:02d})")
        fig_gantt = go.Figure()

        meses_u = df_semanas_vista["mes_num"].unique()
        for idx_m, m_num in enumerate(meses_u):
            sems_m = df_semanas_vista[df_semanas_vista["mes_num"] == m_num]["semana"]
            x0_m, x1_m = sems_m.min() - 0.5, sems_m.max() + 0.5
            if idx_m % 2 == 0:
                fig_gantt.add_vrect(x0=x0_m, x1=x1_m, fillcolor="rgba(255, 255, 255, 0.04)", line_width=0)
            fig_gantt.add_annotation(
                x=(x0_m + x1_m) / 2, y=1.06, yref="paper",
                text=f"<b>{MESES_ES[m_num].upper()}</b>", showarrow=False, font=dict(size=11, color="#34D399")
            )

        for r in df_gantt_all.itertuples():
            ini_c = max(r.s_ini, s_desde)
            fin_c = min(r.s_fin, s_hasta)
            dur = (fin_c - ini_c) + 1

            fig_gantt.add_trace(go.Bar(
                y=[r.label], x=[dur], base=[ini_c - 0.45], orientation="h",
                marker=dict(color=COLORES_ESTADO.get(r.estado, "#3B82F6"), line=dict(color="#FFFFFF", width=1)),
                text=r.texto_barra,
                textposition="inside", insidetextanchor="middle",
                textfont=dict(color="#FFFFFF", size=10),
                showlegend=False
            ))

        fig_gantt.update_layout(
            template="plotly_dark", barmode="overlay",
            plot_bgcolor="#0F172A", paper_bgcolor="#0F172A",
            height=max(340, len(df_gantt_all) * 42 + 100),
            margin=dict(l=20, r=20, t=45, b=30),
            xaxis=dict(
                range=[s_desde - 0.5, s_hasta + 0.5],
                tickvals=df_semanas_vista["semana"].tolist(),
                ticktext=df_semanas_vista["etiqueta"].tolist(),
                title="Semanas del Año", gridcolor="#334155"
            ),
            yaxis=dict(gridcolor="#334155")
        )
        st.plotly_chart(fig_gantt, use_container_width=True, key="gantt_semanal_main")

# =============================================================================
# PESTAÑA 4: PROGRAMAR NUEVA ACTIVIDAD POR CUADRILLA
# =============================================================================
with tab_cargar:
    st.markdown("### ➕ Cargar Programación Real de Mantenimiento")
    with st.container(border=True):
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            sel_lat = st.selectbox(
                "⚡ Línea de Alta Tensión (Región Centro)",
                LINEAS_SIGLAS,
                format_func=lambda s: f"{s} — {CATALOGO_LAT[s]['tramo']} ({CATALOGO_LAT[s]['km']} km)",
                key="add_sel_lat"
            )
        with c2:
            sel_cuad = st.selectbox("👷 Cuadrilla Asignada", list(CUADRILLAS_CONFIG.keys()), key="add_sel_cuad")
        with c3:
            sel_tarea = st.selectbox("🛠️ Actividad / Tarea", CUADRILLAS_CONFIG[sel_cuad]["tareas"], key="add_sel_tarea")

        info_lat = CATALOGO_LAT[sel_lat]
        st.caption(f"📏 Datos físicos de **{sel_lat}**: Longitud total **{info_lat['km']} km** | **{info_lat['piquetes']} piquetes** | Franja total 22m: **{round(info_lat['km']*2.2, 1)} ha**.")

        c4, c5, c6 = st.columns([2.5, 1.2, 1.2])
        with c4:
            tramo_desc = st.text_input("📍 Tramo / Piquetes a Intervenir", value=f"Piq. 1 a {info_lat['piquetes']}", key="add_tramo_desc")
        with c5:
            s_ini = st.number_input("Semana Inicio", min_value=1, max_value=52, value=s_desde, key="add_s_ini")
        with c6:
            s_fin = st.number_input("Semana Fin", min_value=int(s_ini), max_value=52, value=min(52, int(s_ini) + 3), key="add_s_fin")

        c7, c8, c9, c10 = st.columns(4)
        with c7:
            prog_km = st.number_input("Meta Programada (km)", min_value=0.0, max_value=float(info_lat["km"]), value=float(info_lat["km"]) if CUADRILLAS_CONFIG[sel_cuad]["unidad_principal"] == "km" else 0.0, step=0.5, key="add_prog_km")
        with c8:
            ejec_km = st.number_input("Avance Ejecutado (km)", min_value=0.0, max_value=float(info_lat["km"]), value=0.0, step=0.5, key="add_ejec_km")
        with c9:
            prog_piq = st.number_input("Meta Programada (Piquetes)", min_value=0, max_value=int(info_lat["piquetes"]), value=int(info_lat["piquetes"]) if CUADRILLAS_CONFIG[sel_cuad]["unidad_principal"] == "piquetes" else 0, step=1, key="add_prog_piq")
        with c10:
            ejec_piq = st.number_input("Avance Ejecutado (Piquetes)", min_value=0, max_value=int(info_lat["piquetes"]), value=0, step=1, key="add_ejec_piq")

        if sel_cuad == "Cuadrilla 1 - Desmonte y Limpieza" and prog_km > 0:
            st.info(f"🌲 Superficie de desmonte programada (ancho 22m): **{prog_km * FACTOR_HA_POR_KM:.2f} hectáreas** (Ejecutado: **{ejec_km * FACTOR_HA_POR_KM:.2f} ha**).")

        c11, c12, c13 = st.columns([1.5, 1.5, 3])
        with c11:
            ot_val = st.text_input("📋 N° OT", placeholder="Ej: OT-260", key="add_ot_val")
        with c12:
            est_val = st.selectbox("🚦 Estado", ESTADOS_TAREA, index=0, key="add_est_val")
        with c13:
            obs_val = st.text_input("📝 Observaciones", placeholder="Ej: Cumplimiento plan anual ENRE.", key="add_obs_val")

        if st.button("💾 Guardar Programación", type="primary", use_container_width=True, key="btn_save_prog"):
            db_insert_registro({
                "anio": int(anio_sel), "linea": sel_lat, "cuadrilla": sel_cuad,
                "tarea": sel_tarea, "tramo_desc": tramo_desc.strip(),
                "semana_inicio": int(s_ini), "semana_fin": int(s_fin),
                "prog_km": float(prog_km), "ejec_km": float(ejec_km),
                "prog_piq": int(prog_piq), "ejec_piq": int(ejec_piq),
                "ot": ot_val.strip(), "estado": est_val, "observaciones": obs_val.strip()
            })
            st.success("✅ Actividad registrada correctamente.")
            st.rerun()

# =============================================================================
# PESTAÑA 5: ACTUALIZAR AVANCE Y LIMPIAR DATOS
# =============================================================================
with tab_gestion:
    st.markdown("### ✏️ Actualizar Avances Reales y Reprogramaciones")
    df_g = df_raw[df_raw["anio"] == anio_sel].copy() if not df_raw.empty else pd.DataFrame()
    if df_g.empty:
        st.info("No hay registros cargados en el plan.")
    else:
        opciones_ed = df_g.apply(
            lambda r: f"ID {r['id']} | {r['linea']} | {r['cuadrilla'].split('-')[0].strip()} — {r['tarea']} (S{r['semana_inicio']:02d}-S{r['semana_fin']:02d})",
            axis=1
        ).tolist()
        sel_ed = st.selectbox("Seleccionar registro a actualizar:", opciones_ed, key="ed_sel_registro")
        id_ed = int(sel_ed.split(" ")[1])
        fila = df_g[df_g["id"] == id_ed].iloc[0]

        with st.container(border=True):
            e1, e2, e3, e4 = st.columns(4)
            with e1:
                ed_est = st.selectbox("Estado", ESTADOS_TAREA, index=ESTADOS_TAREA.index(fila["estado"]) if fila["estado"] in ESTADOS_TAREA else 0, key=f"ed_est_{id_ed}")
            with e2:
                ed_km = st.number_input("Km Ejecutados", min_value=0.0, value=float(fila["ejec_km"]), step=0.5, key=f"ed_km_{id_ed}")
            with e3:
                ed_piq = st.number_input("Piquetes Ejecutados", min_value=0, value=int(fila["ejec_piq"]), step=1, key=f"ed_piq_{id_ed}")
            with e4:
                ed_si = st.number_input("Semana Inicio", min_value=1, max_value=52, value=int(fila["semana_inicio"]), key=f"ed_si_{id_ed}")

            e5, e6 = st.columns([1, 3])
            with e5:
                ed_sf = st.number_input("Semana Fin", min_value=int(ed_si), max_value=52, value=max(int(ed_si), int(fila["semana_fin"])), key=f"ed_sf_{id_ed}")
            with e6:
                ed_obs = st.text_input("Observaciones", value=str(fila["observaciones"] or ""), key=f"ed_obs_{id_ed}")

            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                if st.button("💾 Guardar Avance", type="primary", use_container_width=True, key=f"btn_upd_{id_ed}"):
                    db_update_avance(id_ed, ed_est, ed_km, ed_piq, ed_si, ed_sf, ed_obs)
                    st.toast("Avance actualizado.", icon="✅")
                    st.rerun()
            with col_b2:
                if st.button("🗑️ Eliminar Registro", use_container_width=True, key=f"btn_del_{id_ed}"):
                    db_delete_registro(id_ed)
                    st.toast("Registro eliminado.", icon="🗑️")
                    st.rerun()

        st.divider()
        with st.expander("🧹 Limpieza de Datos de Prueba / Ficticios", expanded=True):
            st.warning("Esta acción eliminará todos los registros cargados en `plan_mantenimiento_enre` (los datos reales de Perfilado y PAT en sus respectivas tablas **no** se tocan).")
            if st.button("🗑️ Borrar Todos los Datos Ficticios / Vaciar Plan", type="primary", key="btn_vaciar_todo"):
                db_vaciar_plan_completo()
                st.toast("Datos ficticios eliminados correctamente.", icon="🧹")
                st.rerun()
