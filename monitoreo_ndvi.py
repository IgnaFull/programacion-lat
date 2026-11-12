from datetime import date, timedelta
import os
import sys
import numpy as np
import pandas as pd
from pyproj import Transformer
from pystac_client import Client
import rasterio
from rasterio.mask import mask
from shapely.geometry import LineString
from shapely.ops import transform
from supabase import create_client

# =============================================================================
# 1. CONEXIÓN A SUPABASE MEDIANTE VARIABLES DE ENTORNO
# =============================================================================
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip()
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "").strip()

if not SUPABASE_URL or not SUPABASE_KEY:
    print("❌ Error: Faltan las variables de entorno SUPABASE_URL o SUPABASE_KEY.")
    sys.exit(1)

sb = create_client(SUPABASE_URL, SUPABASE_KEY)

# =============================================================================
# 2. CONFIGURACIÓN GEOGRÁFICA NEA (UTM ZONA 21S - EPSG:32721)
# =============================================================================
LINEAS_REGION_CENTRO = [
    "ECO-SMA",
    "LLE-MAN",
    "RES-LLE",
    "RES-BAR1",
    "RES-BAR2",
    "RES-CCE",
    "RES-SCA",
    "SCA-BVI",
    "BVI-GOY",
]

to_utm = Transformer.from_crs("EPSG:4326", "EPSG:32721", always_xy=True).transform


def ndvi_a_pct_limpio(ndvi_val):
    """Convierte el NDVI medio de la franja de 22m a porcentaje de limpieza."""
    if ndvi_val <= 0.25:
        return 100.0
    elif ndvi_val >= 0.65:
        return 0.0
    return round((1.0 - (ndvi_val - 0.25) / (0.65 - 0.25)) * 100.0, 1)


def analizar_linea_sentinel2(sigla_linea, dias_atras=45, max_nubes=20):
    res = (
        sb.table("piquetes_gps")
        .select("*")
        .eq("linea", sigla_linea)
        .order("orden", desc=False)
        .execute()
    )
    df = pd.DataFrame(res.data)
    if df.empty or len(df) < 2:
        print(f"⚠️ [{sigla_linea}] Sin coordenadas GPS suficientes en Supabase.")
        return

    min_lon, min_lat = float(df["longitud"].min()) - 0.02, float(df["latitud"].min()) - 0.02
    max_lon, max_lat = float(df["longitud"].max()) + 0.02, float(df["latitud"].max()) + 0.02
    bbox = [min_lon, min_lat, max_lon, max_lat]

    f_fin = date.today()
    f_ini = f_fin - timedelta(days=dias_atras)

    catalog = Client.open("https://earth-search.aws.element84.com/v1")
    search = catalog.search(
        collections=["sentinel-2-l2a"],
        bbox=bbox,
        datetime=f"{f_ini.isoformat()}/{f_fin.isoformat()}",
        query={"eo:cloud_cover": {"lt": max_nubes}},
        sortby=[{"field": "properties.datetime", "direction": "desc"}],
    )
    items = list(search.items())
    if not items:
        print(f"⚠️ [{sigla_linea}] No se encontraron imágenes despejadas (<{max_nubes}% nubes) en los últimos {dias_atras} días.")
        return

    item = items[0]
    fecha_imagen = item.properties["datetime"][:10]
    url_red = item.assets["red"].href
    url_nir = item.assets["nir"].href
    nubes = item.properties.get("eo:cloud_cover", 0.0)
    print(f"🛰️ [{sigla_linea}] Escena {item.id} | Fecha: {fecha_imagen} | Nubes: {nubes:.1f}%")

    actualizaciones = []
    with rasterio.open(url_red) as src_red, rasterio.open(url_nir) as src_nir:
        to_raster_crs = Transformer.from_crs("EPSG:32721", src_red.crs, always_xy=True).transform

        for i in range(len(df) - 1):
            p1 = df.iloc[i]
            p2 = df.iloc[i + 1]

            linea_wgs = LineString([(float(p1["longitud"]), float(p1["latitud"])), (float(p2["longitud"]), float(p2["latitud"]))])
            linea_utm = transform(to_utm, linea_wgs)
            franja_22m_utm = linea_utm.buffer(11.0, cap_style=2)  # 11m izquierda + 11m derecha (AEA 95301)
            franja_raster = transform(to_raster_crs, franja_22m_utm)

            try:
                red_crop, _ = mask(src_red, [franja_raster], crop=True, filled=False)
                nir_crop, _ = mask(src_nir, [franja_raster], crop=True, filled=False)

                red_arr = red_crop[0].compressed().astype(float)
                nir_arr = nir_crop[0].compressed().astype(float)

                if len(red_arr) > 0 and len(nir_arr) > 0:
                    ndvi_arr = (nir_arr - red_arr) / (nir_arr + red_arr + 1e-6)
                    ndvi_medio = float(np.nanmean(ndvi_arr))
                    pct_limpio = ndvi_a_pct_limpio(ndvi_medio)
                else:
                    pct_limpio = float(p1.get("pct_desmonte") or 0.0)
            except Exception:
                pct_limpio = float(p1.get("pct_desmonte") or 0.0)

            actualizaciones.append({
                "linea": str(p1["linea"]),
                "piquete": str(p1["piquete"]),
                "orden": int(p1["orden"]),
                "latitud": float(p1["latitud"]),
                "longitud": float(p1["longitud"]),
                "dist_km": float(p1.get("dist_km") or 0.0),
                "tipo": str(p1.get("tipo") or ""),
                "estructura": str(p1.get("estructura") or ""),
                "pct_desmonte": pct_limpio,
                "fecha_limpieza": fecha_imagen,
                "tipo_vegetacion": str(p1.get("tipo_vegetacion") or "Monte Nativo / Chilcal"),
            })

    if actualizaciones:
        for idx in range(0, len(actualizaciones), 500):
            sb.table("piquetes_gps").upsert(actualizaciones[idx:idx + 500], on_conflict="linea,piquete").execute()
        print(f"✅ [{sigla_linea}] {len(actualizaciones)} vanos actualizados en Supabase.")


if __name__ == "__main__":
    print("🚀 Iniciando análisis satelital Sentinel-2 para Región Centro...")
    for sigla in LINEAS_REGION_CENTRO:
        try:
            analizar_linea_sentinel2(sigla)
        except Exception as err:
            print(f"❌ Error procesando {sigla}: {err}")
    print("🏁 Proceso finalizado.")
