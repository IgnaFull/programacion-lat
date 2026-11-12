import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, timedelta
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
    "ECO-SMA": {"tramo": "El Colorado (Formosa) / San Martín (Chaco)", "km": 30.47, "piquetes": 126},
    "LLE-MAN": {"tramo": "Mansilla (Formosa) / La Leonesa (Chaco)", "km": 46.60, "piquetes": 162},
    "RES-LLE": {"tramo": "Resistencia (Chaco) / La Leonesa (Chaco)", "km": 62.80, "piquetes": 223},
    "RES-BAR1": {"tramo": "Resistencia (Chaco) / Barranqueras 1 (Chaco)", "km": 30.20, "piquetes": 149},
    "RES-BAR2": {"tramo": "Resistencia (Chaco) / Barranqueras 2 (Chaco)", "km": 31.40, "piquetes": 154},
    "RES-CCE": {"tramo": "Resistencia (Chaco) / Corrientes Centro (Corrientes)", "km": 34.20, "piquetes": 100},
    "RES-SCA": {"tramo": "Resistencia (Chaco) / Santa Catalina (Corrientes)", "km": 42.00, "piquetes": 150},
    "SCA-BVI": {"tramo": "Santa Catalina (Corrientes) / Bella Vista (Corrientes)", "km": 121.00, "piquetes": 463},
    "BVI-GOY": {"tramo": "Bella Vista (Corrientes) / Goya (Corrientes)", "km": 77.70, "piquetes": 303},
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

def normalizar_sigla_lat(nombre):
    if not nombre:
        return ""
    s = str(nombre).strip().upper()
    equivalencias = {
        "SMA-ECO": "ECO-SMA",
        "RES-BAR 1": "RES-BAR1",
        "RES-BAR 2": "RES-BAR2",
        "RES BAR 1": "RES-BAR1",
        "RES BAR 2": "RES-BAR2",
    }
    return equivalencias.get(s, s)

def parsear_fecha_semana(f_val):
    """Convierte fecha string a (anio, semana_iso)."""
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
# 3. PERSISTENCIA DE DATOS E INTEGRACIÓN CON TABLAS REALES DE SUPABASE
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

@st.cache_data(ttl=60)
def fetch_datos_reales_supabase():
    """Lee los avances reales registrados en campo desde 'relevamientos' (Perfilado) y 'mediciones_pat' (PAT)."""
    df_perf = pd.DataFrame()
    df_pat = pd.DataFrame()
    if not MODO_NUBE:
        return df_perf, df_pat

    # 1. Leer Perfilado de Aisladores (relevamientos)
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

    # 2. Leer Mediciones de Puesta a Tierra (mediciones_pat)
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
estado_cloud_txt = "☁️ Conectado a Supabase (Perfilado + PAT en vivo)" if MODO_NUBE else "⚠️ Modo Local (Configurar Secrets)"

st.markdown(f"""
<div class="banner-enre">
    <div>
        <h1 class="banner-title">⚡ TRANSNEA S.A. | Programación LAT e Indicadores ENRE</h1>
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
df_perf_real, df_pat_real = fetch_datos_reales_supabase()

tab_indicadores, tab_cronograma, tab_cargar, tab_gestion = st.tabs([
    "📊 Indicadores y Auditoría ENRE",
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

    # Cálculos Cuadrilla 1 (Desmonte)
    df_desm = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 1 - Desmonte y Limpieza"] if not df_horizonte.empty else pd.DataFrame()
    km_desm_prog = df_desm["prog_km"].sum() if not df_desm.empty else 0.0
    km_desm_ejec = df_desm["ejec_km"].sum() if not df_desm.empty else 0.0
    ha_desm_ejec = round(km_desm_ejec * FACTOR_HA_POR_KM, 2)
    pct_cumpl_desm = (km_desm_ejec / km_desm_prog * 100) if km_desm_prog > 0 else 0.0
    pct_red_limpia = (km_desm_ejec / TOTAL_KM_RED * 100) if TOTAL_KM_RED > 0 else 0.0

    # Cálculos Cuadrilla 2 (Datos reales de Supabase 'relevamientos' + cambios aisladores del plan)
    piq_perf_real_total = len(perf_ventana)
    df_c2 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 2 - Alturas y TCT"] if not df_horizonte.empty else pd.DataFrame()
    piq_aisl_camb = df_c2[df_c2["tarea"] == "Cambio de Aisladores (Correctivo)"]["ejec_piq"].sum() if not df_c2.empty else 0

    # Cálculos Cuadrilla 3 (Datos reales de Supabase 'mediciones_pat' + reparaciones del plan)
    piq_pat_real_total = len(pat_ventana)
    df_c3 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 3 - Puesta a Tierra (PAT)"] if not df_horizonte.empty else pd.DataFrame()
    piq_pat_rep = df_c3[df_c3["tarea"] == "Reparación de PAT"]["ejec_piq"].sum() if not df_c3.empty else 0

    # Cálculos Cuadrilla 4 (Recorridos)
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
# PESTAÑA 2: CRONOGRAMA SEMANAL (PLANIFICADO + EJECUCIÓN REAL DE CAMPO)
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
# PESTAÑA 3: PROGRAMAR NUEVA ACTIVIDAD POR CUADRILLA
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
# PESTAÑA 4: ACTUALIZAR AVANCE Y LIMPIAR DATOS FICTICIOS
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
