import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
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

# Cliente Supabase opcional
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
# 2. CATÁLOGO MAESTRO REGIÓN CENTRO Y PARÁMETROS AEA 95301
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

# =============================================================================
# 3. PERSISTENCIA DE DATOS (SUPABASE + SQLITE LOCAL)
# =============================================================================
@st.cache_resource
def init_supabase():
    if SUPABASE_LIB_AVAILABLE and "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
        try:
            return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
        except Exception:
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

def db_fetch_all():
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
        except Exception:
            pass

    conn = get_sqlite_conn()
    df = pd.read_sql_query("SELECT * FROM plan_mantenimiento_enre ORDER BY semana_inicio ASC", conn)
    conn.close()
    return df

def db_insert_registro(reg):
    if MODO_NUBE:
        try:
            sb_client.table("plan_mantenimiento_enre").insert(reg).execute()
            return
        except Exception:
            pass

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
        except Exception:
            pass

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
        except Exception:
            pass

    conn = get_sqlite_conn()
    conn.execute("DELETE FROM plan_mantenimiento_enre WHERE id = ?", (int(reg_id),))
    conn.commit()
    conn.close()

def cargar_datos_demo(anio_target):
    demos = [
        {"anio": anio_target, "linea": "ECO-SMA", "cuadrilla": "Cuadrilla 1 - Desmonte y Limpieza", "tarea": "Limpieza de Electroducto / Desmonte (Franja 22m)", "tramo_desc": "Piq. 1 a 126 (Traza Completa)", "semana_inicio": 2, "semana_fin": 7, "prog_km": 30.47, "ejec_km": 30.47, "prog_piq": 126, "ejec_piq": 126, "ot": "OT-201", "estado": "Ejecutado", "observaciones": "Franja de 22m (11m+11m) despejada s/ AEA 95301."},
        {"anio": anio_target, "linea": "SCA-BVI", "cuadrilla": "Cuadrilla 1 - Desmonte y Limpieza", "tarea": "Limpieza de Electroducto / Desmonte (Franja 22m)", "tramo_desc": "Piq. 1 a 250", "semana_inicio": 8, "semana_fin": 18, "prog_km": 65.00, "ejec_km": 48.50, "prog_piq": 250, "ejec_piq": 185, "ot": "OT-215", "estado": "En Ejecución", "observaciones": "Monte nativo denso en zona bañados."},
        {"anio": anio_target, "linea": "BVI-GOY", "cuadrilla": "Cuadrilla 1 - Desmonte y Limpieza", "tarea": "Limpieza de Electroducto / Desmonte (Franja 22m)", "tramo_desc": "Piq. 1 a 303", "semana_inicio": 20, "semana_fin": 30, "prog_km": 77.70, "ejec_km": 0.0, "prog_piq": 303, "ejec_piq": 0, "ot": "OT-240", "estado": "Programado", "observaciones": "Programado 2do trimestre."},
        {"anio": anio_target, "linea": "RES-LLE", "cuadrilla": "Cuadrilla 2 - Alturas y TCT", "tarea": "Perfilado de Aisladores", "tramo_desc": "Piq. 1 a 223", "semana_inicio": 4, "semana_fin": 10, "prog_km": 62.80, "ejec_km": 62.80, "prog_piq": 223, "ejec_piq": 223, "ot": "OT-208", "estado": "Ejecutado", "observaciones": "Detectadas 4 cadenas críticas en retenciones."},
        {"anio": anio_target, "linea": "RES-LLE", "cuadrilla": "Cuadrilla 2 - Alturas y TCT", "tarea": "Cambio de Aisladores (Correctivo)", "tramo_desc": "Retenciones Piq. 42, 88, 140, 195", "semana_inicio": 12, "semana_fin": 14, "prog_km": 0.0, "ejec_km": 0.0, "prog_piq": 4, "ejec_piq": 4, "ot": "OT-222", "estado": "Ejecutado", "observaciones": "Reemplazo con TCT sin corte de servicio."},
        {"anio": anio_target, "linea": "LLE-MAN", "cuadrilla": "Cuadrilla 3 - Puesta a Tierra (PAT)", "tarea": "Medición de PAT", "tramo_desc": "Piq. 1 a 162", "semana_inicio": 6, "semana_fin": 12, "prog_km": 46.60, "ejec_km": 46.60, "prog_piq": 162, "ejec_piq": 162, "ot": "OT-211", "estado": "Ejecutado", "observaciones": "14 estructuras superaron 40 Ohm."},
        {"anio": anio_target, "linea": "LLE-MAN", "cuadrilla": "Cuadrilla 3 - Puesta a Tierra (PAT)", "tarea": "Reparación de PAT", "tramo_desc": "14 Estructuras Observadas", "semana_inicio": 13, "semana_fin": 16, "prog_km": 0.0, "ejec_km": 0.0, "prog_piq": 14, "ejec_piq": 10, "ot": "OT-229", "estado": "En Ejecución", "observaciones": "Reconexión de bajadas y jabalinas."},
        {"anio": anio_target, "linea": "RES-BAR1", "cuadrilla": "Cuadrilla 4 - Recorrido e Inspección", "tarea": "Recorrido con Dron", "tramo_desc": "Traza Completa", "semana_inicio": 5, "semana_fin": 6, "prog_km": 30.20, "ejec_km": 30.20, "prog_piq": 149, "ejec_piq": 149, "ot": "OT-205", "estado": "Ejecutado", "observaciones": "Inspección termográfica y visual HD."},
        {"anio": anio_target, "linea": "RES-BAR2", "cuadrilla": "Cuadrilla 4 - Recorrido e Inspección", "tarea": "Recorrido Visual Rápido", "tramo_desc": "Traza Completa", "semana_inicio": 9, "semana_fin": 10, "prog_km": 31.40, "ejec_km": 31.40, "prog_piq": 154, "ejec_piq": 154, "ot": "OT-218", "estado": "Ejecutado", "observaciones": "Sin novedades críticas en franja."},
        {"anio": anio_target, "linea": "RES-CCE", "cuadrilla": "Cuadrilla 4 - Recorrido e Inspección", "tarea": "Recorrido Visual Lento", "tramo_desc": "Cruce Río Paraná y accesos", "semana_inicio": 15, "semana_fin": 17, "prog_km": 34.20, "ejec_km": 18.00, "prog_piq": 100, "ejec_piq": 52, "ot": "OT-233", "estado": "En Ejecución", "observaciones": "Verificación detallada de herrajes."},
        {"anio": anio_target, "linea": "RES-SCA", "cuadrilla": "Cuadrilla 2 - Alturas y TCT", "tarea": "Perfilado de Aisladores", "tramo_desc": "Piq. 1 a 150", "semana_inicio": 26, "semana_fin": 33, "prog_km": 42.00, "ejec_km": 0.0, "prog_piq": 150, "ejec_piq": 0, "ot": "OT-255", "estado": "Programado", "observaciones": "Planificado 2do semestre."},
    ]
    for d in demos:
        db_insert_registro(d)

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

    story.append(Paragraph(f"TRANSNEA S.A. — INFORME TÉCNICO DE MANTENIMIENTO LAT (AUDITORÍA ENRE)", title_style))
    story.append(Paragraph(
        f"<b>Región Centro</b> | Año Plan: {anio} | Horizonte: {horizonte_txt} | "
        f"Marco Normativo: <b>Reglamentación AEA 95301</b> (Franja de Servidumbre: 22 m totales — 11 m a cada lado del eje/hilo de guardia). "
        f"Fecha de emisión: {date.today().strftime('%d/%m/%Y')}",
        sub_style
    ))

    # Sección 1: Estado de Desmonte y Cobertura por Línea (Cuadrilla 1 y Red)
    story.append(Paragraph("<b>1. INDICADORES DE LIMPIEZA DE ELECTRODUCTO (CUADRILLA 1) Y COBERTURA POR LÍNEA</b>", styles["Normal"]))
    story.append(Spacer(1, 5))

    t1_data = [[
        "Línea", "Tramo Geográfico", "Long. Total\n(km)", "Franja Total\n(ha - 22m)",
        "Desmonte Prog.\n(km)", "Desmonte Ejec.\n(km)", "Sup. Limpia\n(ha)",
        "% Cumpl.\ns/ Prog.", "% Avance\ns/ Línea"
    ]]

    for r in df_resumen_lineas.itertuples():
        t1_data.append([
            r.Sigla,
            r.Tramo[:38],
            f"{r.Km_Total:.2f}",
            f"{r.Ha_Total:.1f}",
            f"{r.Desm_Prog_Km:.2f}",
            f"{r.Desm_Ejec_Km:.2f}",
            f"{r.Desm_Ejec_Ha:.1f}",
            f"{r.Pct_Cumpl_Prog:.1f}%",
            f"{r.Pct_Linea_Limpia:.1f}%"
        ])

    t1 = Table(t1_data, colWidths=[58, 195, 58, 62, 68, 68, 60, 62, 62])
    t1.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0E5E35")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAF9")]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t1)
    story.append(Spacer(1, 12))

    # Sección 2: Detalle de Actividades Programadas y Ejecutadas (4 Cuadrillas)
    story.append(Paragraph("<b>2. CRONOGRAMA Y EJECUCIÓN POR CUADRILLA (ALTAS TENSIONES REGIÓN CENTRO)</b>", styles["Normal"]))
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

    # HOJA 1: INDICADORES DE DESMONTE Y CUMPLIMIENTO POR LÍNEA
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

    ws1.merge_cells("A1:K1")
    ws1["A1"] = f"TRANSNEA S.A. - INDICADORES DE MANTENIMIENTO Y DESMONTE (AEA 95301 - FRANJA 22M) - AÑO {anio}"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_green
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    headers_h1 = [
        "Sigla LAT", "Tramo", "Longitud (km)", "Piquetes", "Franja Total (ha)",
        "Desmonte Prog (km)", "Desmonte Ejec (km)", "Superficie Limpia (ha)",
        "% Cumpl. s/ Programado", "% Limpio s/ Total Línea", "Piquetes Intervenidos (TCT/PAT)"
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
            round(r.Pct_Cumpl_Prog, 2), round(r.Pct_Linea_Limpia, 2), r.Piq_Intervenidos
        ]
        for c_i, v in enumerate(vals, start=1):
            c = ws1.cell(row=r_i, column=c_i, value=v)
            c.border = thin_border
            if c_i >= 3:
                c.alignment = Alignment(horizontal="center")

    for col in ws1.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws1.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_len + 3, 14)

    # HOJA 2: CRONOGRAMA SEMANAL (MATRIZ GANTT)
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
st.markdown(f"""
<div class="banner-enre">
    <div>
        <h1 class="banner-title">⚡ TRANSNEA S.A. | Programación LAT e Indicadores ENRE</h1>
        <div class="banner-sub">Región Centro ({TOTAL_KM_RED} km • {TOTAL_PIQ_RED:,} Piquetes) | Reglamentación AEA 95301 (Franja 22m: 11m + 11m)</div>
    </div>
    <div class="banner-badge">
        <strong>Coordinación de Mantenimiento</strong><br>
        4 Cuadrillas Operativas • Control ENRE
    </div>
</div>
""", unsafe_allow_html=True)

with st.container(border=True):
    ch1, ch2, ch3, ch4 = st.columns([1.1, 2.2, 2.3, 1.9])
    with ch1:
        anio_sel = st.selectbox("📅 Año Plan", [2026, 2027, 2025], index=0)
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
            index=0
        )

    df_semanas_anio = generar_mapa_semanas(anio_sel)

    with ch3:
        if horizonte_sel.startswith("12 Meses"):
            s_desde, s_hasta = 1, 52
            st.info("Horizonte completo: **S01 a S52** (Ene – Dic)")
        elif horizonte_sel.startswith("6 Meses"):
            sem = st.radio("Semestre:", ["1° Semestre (Ene-Jun)", "2° Semestre (Jul-Dic)"], horizontal=True)
            s_desde, s_hasta = (1, 26) if sem.startswith("1°") else (27, 52)
        elif horizonte_sel.startswith("4 Meses"):
            cua = st.selectbox("Cuatrimestre:", ["1° Cuatrimestre (Ene-Abr)", "2° Cuatrimestre (May-Ago)", "3° Cuatrimestre (Sep-Dic)"])
            s_desde, s_hasta = (1, 17) if cua.startswith("1°") else ((18, 35) if cua.startswith("2°") else (36, 52))
        elif horizonte_sel.startswith("3 Meses"):
            tri = st.selectbox("Trimestre:", ["Q1 (Ene-Mar)", "Q2 (Abr-Jun)", "Q3 (Jul-Sep)", "Q4 (Oct-Dic)"])
            s_desde, s_hasta = (1, 13) if tri.startswith("Q1") else ((14, 26) if tri.startswith("Q2") else ((27, 39) if tri.startswith("Q3") else (40, 52)))
        else:
            mes_sel = st.selectbox("Mes:", list(MESES_ES.keys()), format_func=lambda m: f"{MESES_ES[m]} {anio_sel}", index=datetime.now().month - 1)
            sems_m = df_semanas_anio[df_semanas_anio["mes_num"] == mes_sel]["semana"].tolist()
            s_desde, s_hasta = min(sems_m), max(sems_m)

    with ch4:
        filtro_cuadrilla = st.selectbox("👷 Filtrar Cuadrilla:", ["TODAS LAS CUADRILLAS"] + list(CUADRILLAS_CONFIG.keys()))

df_semanas_vista = df_semanas_anio[(df_semanas_anio["semana"] >= s_desde) & (df_semanas_anio["semana"] <= s_hasta)].copy()
df_raw = db_fetch_all()

tab_indicadores, tab_cronograma, tab_cargar, tab_gestion = st.tabs([
    "📊 Indicadores y Auditoría ENRE",
    "🗓️ Cronograma Semanal",
    "➕ Programar Actividad",
    "✏️ Actualizar Avance / Reprogramar"
])

# =============================================================================
# PESTAÑA 1: INDICADORES DE DESMONTE (AEA 95301) Y 4 CUADRILLAS + INFORMES ENRE
# =============================================================================
with tab_indicadores:
    if df_raw.empty:
        st.warning("No hay registros cargados para el año seleccionado.")
        if st.button("🚀 Cargar Plan de Ejemplo Región Centro (4 Cuadrillas) para Presentación", type="primary"):
            cargar_datos_demo(anio_sel)
            st.rerun()
    else:
        df_anio = df_raw[df_raw["anio"] == anio_sel].copy()
        df_horizonte = df_anio[(df_anio["semana_fin"] >= s_desde) & (df_anio["semana_inicio"] <= s_hasta)].copy()

        # Cálculo específico Cuadrilla 1 (Desmonte y Limpieza de Electroducto - Franja 22m)
        df_desm = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 1 - Desmonte y Limpieza"]
        km_desm_prog = df_desm["prog_km"].sum()
        km_desm_ejec = df_desm["ejec_km"].sum()
        ha_desm_ejec = round(km_desm_ejec * FACTOR_HA_POR_KM, 2)
        pct_cumpl_desm = (km_desm_ejec / km_desm_prog * 100) if km_desm_prog > 0 else 0.0
        pct_red_limpia = (km_desm_ejec / TOTAL_KM_RED * 100) if TOTAL_KM_RED > 0 else 0.0

        # Cálculo Cuadrilla 2 (Alturas y TCT)
        df_c2 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 2 - Alturas y TCT"]
        piq_perf_ejec = df_c2[df_c2["tarea"] == "Perfilado de Aisladores"]["ejec_piq"].sum()
        piq_aisl_camb = df_c2[df_c2["tarea"] == "Cambio de Aisladores (Correctivo)"]["ejec_piq"].sum()

        # Cálculo Cuadrilla 3 (PAT)
        df_c3 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 3 - Puesta a Tierra (PAT)"]
        piq_pat_med = df_c3[df_c3["tarea"] == "Medición de PAT"]["ejec_piq"].sum()
        piq_pat_rep = df_c3[df_c3["tarea"] == "Reparación de PAT"]["ejec_piq"].sum()

        # Cálculo Cuadrilla 4 (Recorridos)
        df_c4 = df_horizonte[df_horizonte["cuadrilla"] == "Cuadrilla 4 - Recorrido e Inspección"]
        km_rec_dron = df_c4[df_c4["tarea"] == "Recorrido con Dron"]["ejec_km"].sum()
        km_rec_vis = df_c4[df_c4["tarea"].str.contains("Visual")]["ejec_km"].sum()

        st.markdown("#### 🌲 Cuadrilla 1: Limpieza de Electroducto (Franja Servidumbre 22 m — AEA 95301)")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Desmonte Programado", f"{km_desm_prog:.2f} km", f"{km_desm_prog * FACTOR_HA_POR_KM:.1f} ha")
        k2.metric("Desmonte Ejecutado", f"{km_desm_ejec:.2f} km", f"{ha_desm_ejec:.1f} ha limpias")
        k3.metric("% Cumpl. s/ Programado", f"{pct_cumpl_desm:.1f}%", delta=f"{km_desm_ejec - km_desm_prog:.2f} km")
        k4.metric("% Avance s/ Red Total", f"{pct_red_limpia:.2f}%", f"de {TOTAL_KM_RED} km totales")
        k5.metric("Ancho Normativo AEA", "22 metros", "11m + 11m eje línea")

        st.markdown("#### 👷 Resumen Operativo Cuadrillas 2, 3 y 4")
        c_kpi1, c_kpi2, c_kpi3, c_kpi4 = st.columns(4)
        c_kpi1.metric("Cuadrilla 2 (TCT)", f"{int(piq_perf_ejec)} piq. perfilados", f"{int(piq_aisl_camb)} cambios aislador")
        c_kpi2.metric("Cuadrilla 3 (PAT)", f"{int(piq_pat_med)} PAT medidas", f"{int(piq_pat_rep)} PAT reparadas")
        c_kpi3.metric("Cuadrilla 4 (Dron)", f"{km_rec_dron:.2f} km", "Inspección aérea HD/IR")
        c_kpi4.metric("Cuadrilla 4 (Visual)", f"{km_rec_vis:.2f} km", "Recorrido lento/rápido")

        st.divider()

        # Tabla consolidada por Línea LAT
        filas_lineas = []
        for sigla, info in CATALOGO_LAT.items():
            km_tot = info["km"]
            piq_tot = info["piquetes"]
            ha_tot = round(km_tot * FACTOR_HA_POR_KM, 2)

            sub_l_desm = df_desm[df_desm["linea"] == sigla]
            d_prog = sub_l_desm["prog_km"].sum()
            d_ejec = sub_l_desm["ejec_km"].sum()
            d_ha = round(d_ejec * FACTOR_HA_POR_KM, 2)

            pct_prog = (d_ejec / d_prog * 100) if d_prog > 0 else 0.0
            pct_linea = (d_ejec / km_tot * 100) if km_tot > 0 else 0.0

            sub_l_all = df_horizonte[df_horizonte["linea"] == sigla]
            piq_interv = int(sub_l_all[sub_l_all["cuadrilla"] != "Cuadrilla 1 - Desmonte y Limpieza"]["ejec_piq"].sum())

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
                "Piq_Intervenidos": piq_interv
            })

        df_resumen_lineas = pd.DataFrame(filas_lineas)

        col_g1, col_g2 = st.columns([1.15, 0.85])
        with col_g1:
            st.markdown("##### 📊 Avance de Limpieza de Electroducto por Línea (% sobre Traza Total)")
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                y=df_resumen_lineas["Sigla"],
                x=df_resumen_lineas["Pct_Linea_Limpia"],
                orientation="h",
                text=[f" <b>{r.Pct_Linea_Limpia:.1f}%</b> ({r.Desm_Ejec_Km:.1f} / {r.Km_Total:.1f} km)" for r in df_resumen_lineas.itertuples()],
                textposition="outside",
                cliponaxis=False,
                marker=dict(color="#10B981", line=dict(color="#FFFFFF", width=1)),
            ))
            fig_bar.update_layout(
                template="plotly_dark",
                plot_bgcolor="#0F172A", paper_bgcolor="#0F172A",
                xaxis=dict(range=[0, 125], title="% de Línea Limpia (Ancho 22m)", gridcolor="#334155"),
                yaxis=dict(autorange="reversed", gridcolor="#334155"),
                height=380, margin=dict(l=20, r=40, t=20, b=20)
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_g2:
            st.markdown("##### 📋 Tabla de Control por Línea LAT")
            df_show_l = df_resumen_lineas[[
                "Sigla", "Km_Total", "Desm_Prog_Km", "Desm_Ejec_Km", "Desm_Ejec_Ha", "Pct_Cumpl_Prog", "Pct_Linea_Limpia"
            ]].rename(columns={
                "Sigla": "Línea", "Km_Total": "Total km", "Desm_Prog_Km": "Prog. km",
                "Desm_Ejec_Km": "Ejec. km", "Desm_Ejec_Ha": "Limpio (ha)",
                "Pct_Cumpl_Prog": "% s/ Prog.", "Pct_Linea_Limpia": "% s/ Línea"
            })
            st.dataframe(
                df_show_l.style.format({
                    "Total km": "{:.2f}", "Prog. km": "{:.2f}", "Ejec. km": "{:.2f}",
                    "Limpio (ha)": "{:.1f}", "% s/ Prog.": "{:.1f}%", "% s/ Línea": "{:.1f}%"
                }),
                use_container_width=True, hide_index=True, height=380
            )

        # Botones de Exportación Oficial para Auditoría ENRE
        st.divider()
        st.markdown("#### 🏛️ Exportación de Informes Oficiales para Auditoría ENRE (AEA 95301)")
        col_ex1, col_ex2 = st.columns(2)
        horizonte_corto = horizonte_sel.split("(")[0].strip()

        with col_ex1:
            xl_bytes = generar_excel_auditoria_enre(df_resumen_lineas, df_horizonte, df_semanas_vista, anio_sel, horizonte_corto)
            st.download_button(
                label=f"📊 Descargar Matriz e Indicadores ENRE en Excel (.xlsx)",
                data=xl_bytes,
                file_name=f"Auditoria_ENRE_Transnea_RegionCentro_{anio_sel}_S{s_desde:02d}_S{s_hasta:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_ex2:
            if REPORTLAB_AVAILABLE:
                pdf_bytes = generar_pdf_auditoria_enre(df_resumen_lineas, df_horizonte, anio_sel, f"{horizonte_corto} (S{s_desde:02d} a S{s_hasta:02d})")
                st.download_button(
                    label=f"📄 Descargar Informe Técnico Oficial ENRE en PDF (.pdf)",
                    data=pdf_bytes,
                    file_name=f"Informe_ENRE_AEA95301_Transnea_{anio_sel}_S{s_desde:02d}_S{s_hasta:02d}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
            else:
                st.info("Instalá 'reportlab' en requirements.txt para habilitar la exportación PDF.")

# =============================================================================
# PESTAÑA 2: CRONOGRAMA SEMANAL (12, 6, 4, 3 Y 1 MES)
# =============================================================================
with tab_cronograma:
    if df_raw.empty:
        st.info("Cargá actividades o iniciá los datos de ejemplo en la pestaña de Indicadores.")
    else:
        df_crono = df_raw[(df_raw["anio"] == anio_sel) & (df_raw["semana_fin"] >= s_desde) & (df_raw["semana_inicio"] <= s_hasta)].copy()
        if filtro_cuadrilla != "TODAS LAS CUADRILLAS":
            df_crono = df_crono[df_crono["cuadrilla"] == filtro_cuadrilla]

        if df_crono.empty:
            st.info(f"No hay actividades programadas entre S{s_desde:02d} y S{s_hasta:02d} para este filtro.")
        else:
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

            for r in df_crono.sort_values(by=["semana_inicio", "linea"], ascending=[False, True]).itertuples():
                ini_c = max(r.semana_inicio, s_desde)
                fin_c = min(r.semana_fin, s_hasta)
                dur = (fin_c - ini_c) + 1
                c_short = r.cuadrilla.split("-")[0].strip()
                y_lbl = f"<b>{r.linea}</b> | {c_short}: {r.tarea[:22]}"
                meta_txt = f"{r.ejec_km}/{r.prog_km} km" if r.prog_km > 0 else f"{r.ejec_piq}/{r.prog_piq} piq"

                fig_gantt.add_trace(go.Bar(
                    y=[y_lbl], x=[dur], base=[ini_c - 0.45], orientation="h",
                    marker=dict(color=COLORES_ESTADO.get(r.estado, "#3B82F6"), line=dict(color="#FFFFFF", width=1)),
                    text=f"S{r.semana_inicio:02d}-S{r.semana_fin:02d} ({meta_txt})",
                    textposition="inside", insidetextanchor="middle",
                    textfont=dict(color="#FFFFFF", size=10),
                    showlegend=False
                ))

            fig_gantt.update_layout(
                template="plotly_dark", barmode="overlay",
                plot_bgcolor="#0F172A", paper_bgcolor="#0F172A",
                height=max(340, len(df_crono) * 42 + 100),
                margin=dict(l=20, r=20, t=45, b=30),
                xaxis=dict(
                    range=[s_desde - 0.5, s_hasta + 0.5],
                    tickvals=df_semanas_vista["semana"].tolist(),
                    ticktext=df_semanas_vista["etiqueta"].tolist(),
                    title="Semanas del Año", gridcolor="#334155"
                ),
                yaxis=dict(gridcolor="#334155")
            )
            st.plotly_chart(fig_gantt, use_container_width=True)

# =============================================================================
# PESTAÑA 3: PROGRAMAR NUEVA ACTIVIDAD POR CUADRILLA
# =============================================================================
with tab_cargar:
    st.markdown("### ➕ Cargar Programación de Mantenimiento")
    with st.container(border=True):
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            sel_lat = st.selectbox(
                "⚡ Línea de Alta Tensión (Región Centro)",
                LINEAS_SIGLAS,
                format_func=lambda s: f"{s} — {CATALOGO_LAT[s]['tramo']} ({CATALOGO_LAT[s]['km']} km)"
            )
        with c2:
            sel_cuad = st.selectbox("👷 Cuadrilla Asignada", list(CUADRILLAS_CONFIG.keys()))
        with c3:
            sel_tarea = st.selectbox("🛠️ Actividad / Tarea", CUADRILLAS_CONFIG[sel_cuad]["tareas"])

        info_lat = CATALOGO_LAT[sel_lat]
        st.caption(f"📏 Datos físicos de **{sel_lat}**: Longitud total **{info_lat['km']} km** | **{info_lat['piquetes']} piquetes** | Franja total 22m: **{round(info_lat['km']*2.2, 1)} ha**.")

        c4, c5, c6 = st.columns([2.5, 1.2, 1.2])
        with c4:
            tramo_desc = st.text_input("📍 Tramo / Piquetes a Intervenir", value=f"Piq. 1 a {info_lat['piquetes']}")
        with c5:
            s_ini = st.number_input("Semana Inicio", min_value=1, max_value=52, value=s_desde)
        with c6:
            s_fin = st.number_input("Semana Fin", min_value=int(s_ini), max_value=52, value=min(52, int(s_ini) + 3))

        c7, c8, c9, c10 = st.columns(4)
        with c7:
            prog_km = st.number_input("Meta Programada (km)", min_value=0.0, max_value=float(info_lat["km"]), value=float(info_lat["km"]) if CUADRILLAS_CONFIG[sel_cuad]["unidad_principal"] == "km" else 0.0, step=0.5)
        with c8:
            ejec_km = st.number_input("Avance Ejecutado (km)", min_value=0.0, max_value=float(info_lat["km"]), value=0.0, step=0.5)
        with c9:
            prog_piq = st.number_input("Meta Programada (Piquetes)", min_value=0, max_value=int(info_lat["piquetes"]), value=int(info_lat["piquetes"]) if CUADRILLAS_CONFIG[sel_cuad]["unidad_principal"] == "piquetes" else 0, step=1)
        with c10:
            ejec_piq = st.number_input("Avance Ejecutado (Piquetes)", min_value=0, max_value=int(info_lat["piquetes"]), value=0, step=1)

        if sel_cuad == "Cuadrilla 1 - Desmonte y Limpieza" and prog_km > 0:
            st.info(f"🌲 Superficie de desmonte programada (ancho 22m): **{prog_km * FACTOR_HA_POR_KM:.2f} hectáreas** (Ejecutado: **{ejec_km * FACTOR_HA_POR_KM:.2f} ha**).")

        c11, c12, c13 = st.columns([1.5, 1.5, 3])
        with c11:
            ot_val = st.text_input("📋 N° OT", placeholder="Ej: OT-260")
        with c12:
            est_val = st.selectbox("🚦 Estado", ESTADOS_TAREA, index=0)
        with c13:
            obs_val = st.text_input("📝 Observaciones", placeholder="Ej: Cumplimiento plan anual ENRE.")

        if st.button("💾 Guardar Programación", type="primary", use_container_width=True):
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
# PESTAÑA 4: ACTUALIZAR AVANCE Y REPROGRAMAR
# =============================================================================
with tab_gestion:
    st.markdown("### ✏️ Actualizar Avances Reales y Reprogramaciones")
    df_g = df_raw[df_raw["anio"] == anio_sel].copy()
    if df_g.empty:
        st.info("No hay registros cargados para editar.")
    else:
        opciones_ed = df_g.apply(
            lambda r: f"ID {r['id']} | {r['linea']} | {r['cuadrilla'].split('-')[0].strip()} — {r['tarea']} (S{r['semana_inicio']:02d}-S{r['semana_fin']:02d})",
            axis=1
        ).tolist()
        sel_ed = st.selectbox("Seleccionar registro a actualizar:", opciones_ed)
        id_ed = int(sel_ed.split(" ")[1])
        fila = df_g[df_g["id"] == id_ed].iloc[0]

        with st.container(border=True):
            e1, e2, e3, e4 = st.columns(4)
            with e1:
                ed_est = st.selectbox("Estado", ESTADOS_TAREA, index=ESTADOS_TAREA.index(fila["estado"]) if fila["estado"] in ESTADOS_TAREA else 0)
            with e2:
                ed_km = st.number_input("Km Ejecutados", min_value=0.0, value=float(fila["ejec_km"]), step=0.5)
            with e3:
                ed_piq = st.number_input("Piquetes Ejecutados", min_value=0, value=int(fila["ejec_piq"]), step=1)
            with e4:
                ed_si = st.number_input("Semana Inicio", min_value=1, max_value=52, value=int(fila["semana_inicio"]))

            e5, e6 = st.columns([1, 3])
            with e5:
                ed_sf = st.number_input("Semana Fin", min_value=int(ed_si), max_value=52, value=max(int(ed_si), int(fila["semana_fin"])))
            with e6:
                ed_obs = st.text_input("Observaciones", value=str(fila["observaciones"] or ""))

            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                if st.button("💾 Guardar Avance", type="primary", use_container_width=True):
                    db_update_avance(id_ed, ed_est, ed_km, ed_piq, ed_si, ed_sf, ed_obs)
                    st.toast("Avance actualizado.", icon="✅")
                    st.rerun()
            with col_b2:
                if st.button("🗑️ Eliminar", use_container_width=True):
                    db_delete_registro(id_ed)
                    st.toast("Registro eliminado.", icon="🗑️")
                    st.rerun()
