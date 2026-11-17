from datetime import date, timedelta
import os
import sys
import numpy as np
import pandas as pd
from pyproj import Transformer
from pystac_client import Client
import rasterio
from rasterio.mask import mask
from rasterio.warp import reproject, Resampling
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

# Clases SCL Sentinel-2 L2A:
# 4 = Vegetación, 5 = Suelo desnudo / despejado, 6 = Agua, 2/3/8/9/10 = Sombras y Nubes
SCL_VALIDOS_TIERRA = [4, 5]
SCL_NUBES_SOMBRAS = [2, 3, 8, 9, 10]

to_utm = Transformer.from_crs("EPSG:4326", "EPSG:32721", always_xy=True).transform


def leer_banda_alineada(src, ref_transform, ref_shape, ref_crs):
    """Reproyecta una banda de 20m (SCL o SWIR1) a la grilla exacta de 10m de Red/NIR."""
    dst_arr = np.zeros(ref_shape, dtype=np.float32)
    try:
        reproject(
            source=rasterio.band(src, 1),
            destination=dst_arr,
            src_transform=src.transform,
            src_crs=src.crs,
            dst_transform=ref_transform,
            dst_crs=ref_crs,
            resampling=Resampling.nearest,
        )
    except Exception:
        return None
    return dst_arr


def evaluar_corredor_espectral(red_arr, nir_arr, swir_arr, scl_arr, veg_actual):
    """
    Filtra nubes/agua con SCL y separa pastizales bajos de monte alto usando NDVI + NDMI.
    Retorna: (pct_limpio, tipo_vegetacion_detectado, fraccion_nubes)
    """
    total_px = len(scl_arr)
    if total_px == 0:
        return None, veg_actual, 1.0

    nubes_px = int(np.sum(np.isin(scl_arr, SCL_NUBES_SOMBRAS)))
    agua_px = int(np.sum(scl_arr == 6))
    frac_nubes = nubes_px / float(total_px)

    # Si más del 50% del vano está cubierto por nubes o sombras, conservar dato histórico
    if frac_nubes > 0.50:
        return None, veg_actual, frac_nubes

    mask_valid = np.isin(scl_arr, SCL_VALIDOS_TIERRA) & (red_arr > 0) & (nir_arr > 0)
    if int(np.sum(mask_valid)) == 0:
        if agua_px > (total_px * 0.5):
            return 100.0, "Bañado / Estero / Juncal", frac_nubes
        return None, veg_actual, frac_nubes

    r = red_arr[mask_valid]
    n = nir_arr[mask_valid]
    s = swir_arr[mask_valid] if swir_arr is not None else np.zeros_like(r)
    scl_v = scl_arr[mask_valid]

    ndvi = (n - r) / (n + r + 1e-6)
    ndmi = np.where(s > 0, (n - s) / (n + s + 1e-6), 0.0)

    # 1. Despejado / Suelo desnudo (SCL=5 o NDVI <= 0.28) -> 100% limpio
    es_limpio = (scl_v == 5) | (ndvi <= 0.28)

    # 2. Monte alto / Leñoso denso (NDVI > 0.52 y NDMI > 0.18, o NDVI > 0.62) -> Crítico (0% a 35% limpio)
    es_monte_alto = (~es_limpio) & (((ndvi > 0.52) & (ndmi > 0.18)) | (ndvi > 0.62))

    # 3. Pastizal / Herbáceo / Rebrote bajo -> 75% a 90% limpio
    es_pastizal = (~es_limpio) & (~es_monte_alto)

    scores_px = np.zeros_like(ndvi, dtype=np.float32)
    scores_px[es_limpio] = 100.0

    if np.any(es_pastizal):
        ndvi_p = np.clip(ndvi[es_pastizal], 0.28, 0.62)
        scores_px[es_pastizal] = 90.0 - ((ndvi_p - 0.28) / 0.34) * 15.0

    if np.any(es_monte_alto):
        ndvi_m = np.clip(ndvi[es_monte_alto], 0.52, 0.75)
        scores_px[es_monte_alto] = np.maximum(0.0, 35.0 * (1.0 - (ndvi_m - 0.52) / 0.23))

    pct_limpio_medio = round(float(np.mean(scores_px)), 1)

    prop_monte = float(np.mean(es_monte_alto))
    prop_pastizal = float(np.mean(es_pastizal))

    if agua_px > 0 and (prop_monte > 0.15 or prop_pastizal > 0.25):
        tipo_veg_out = "Bañado / Estero / Juncal"
    elif prop_monte >= 0.25:
        tipo_veg_out = "Monte Nativo / Chilcal"
    elif prop_pastizal >= 0.30:
        tipo_veg_out = "Pastizal / Campo Abierto"
    else:
        tipo_veg_out = veg_actual or "Monte Nativo / Chilcal"

    return pct_limpio_medio, tipo_veg_out, frac_nubes


def analizar_linea_sentinel2(sigla_linea, dias_atras=45, max_nubes=25):
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
        print(f"⚠️ [{sigla_linea}] Sin escenas despejadas (<{max_nubes}% nubes) en los últimos {dias_atras} días.")
        return

    item = items[0]
    fecha_imagen = item.properties["datetime"][:10]
    url_red = item.assets["red"].href
    url_nir = item.assets["nir"].href
    url_swir = item.assets["swir16"].href if "swir16" in item.assets else None
    url_scl = item.assets["scl"].href if "scl" in item.assets else None
    nubes_escena = item.properties.get("eo:cloud_cover", 0.0)

    print(f"🛰️ [{sigla_linea}] Escena {item.id} | Fecha: {fecha_imagen} | Nubes escena: {nubes_escena:.1f}%")

    actualizaciones = []
    vanos_omitidos_nubes = 0

    with rasterio.open(url_red) as src_red, rasterio.open(url_nir) as src_nir:
        src_swir = rasterio.open(url_swir) if url_swir else None
        src_scl = rasterio.open(url_scl) if url_scl else None

        to_raster_crs = Transformer.from_crs("EPSG:32721", src_red.crs, always_xy=True).transform

        for i in range(len(df) - 1):
            p1 = df.iloc[i]
            p2 = df.iloc[i + 1]

            p1_id = str(p1["piquete"]).strip()
            p2_id = str(p2["piquete"]).strip()

            # Omitir tramos sin servidumbre terrestre (Cruce Río Paraná y Cable Subterráneo Corrientes)
            es_subt_o_rio = (
                (sigla_linea in ["RES-CCE", "RES-SCA"] and p1_id in ["94", "95"] and p2_id in ["95", "96"])
                or (sigla_linea == "RES-SCA" and p1_id == "96" and p2_id == "101")
            )
            if es_subt_o_rio:
                actualizaciones.append({
                    "linea": str(p1["linea"]),
                    "piquete": p1_id,
                    "orden": int(p1["orden"]),
                    "latitud": float(p1["latitud"]),
                    "longitud": float(p1["longitud"]),
                    "dist_km": float(p1.get("dist_km") or 0.0),
                    "tipo": str(p1.get("tipo") or ""),
                    "estructura": str(p1.get("estructura") or ""),
                    "pct_desmonte": 100.0,
                    "fecha_limpieza": fecha_imagen,
                    "tipo_vegetacion": "Pastizal / Campo Abierto",
                })
                continue

            pct_previo = float(p1.get("pct_desmonte") or 0.0)
            fecha_previa = str(p1.get("fecha_limpieza") or fecha_imagen)
            veg_previa = str(p1.get("tipo_vegetacion") or "Monte Nativo / Chilcal")

            linea_wgs = LineString([
                (float(p1["longitud"]), float(p1["latitud"])),
                (float(p2["longitud"]), float(p2["latitud"])),
            ])
            linea_utm = transform(to_utm, linea_wgs)

            # Doble corredor: 10m interno (5m+5m conductores) y 22m total (11m+11m AEA 95301)
            franja_22m_raster = transform(to_raster_crs, linea_utm.buffer(11.0, cap_style=2))
            franja_10m_raster = transform(to_raster_crs, linea_utm.buffer(5.0, cap_style=2))

            try:
                red_22, out_trans = mask(src_red, [franja_22m_raster], crop=True, filled=True, nodata=0)
                nir_22, _ = mask(src_nir, [franja_22m_raster], crop=True, filled=True, nodata=0)

                ref_shape = red_22[0].shape
                valid_geom_mask = red_22[0] > 0

                if src_scl is not None:
                    scl_grid = leer_banda_alineada(src_scl, out_trans, ref_shape, src_red.crs)
                    scl_22 = scl_grid[valid_geom_mask].astype(int) if scl_grid is not None else np.full(np.sum(valid_geom_mask), 4)
                else:
                    scl_22 = np.full(np.sum(valid_geom_mask), 4)

                if src_swir is not None:
                    swir_grid = leer_banda_alineada(src_swir, out_trans, ref_shape, src_red.crs)
                    swir_22 = swir_grid[valid_geom_mask] if swir_grid is not None else None
                else:
                    swir_22 = None

                r_22_vec = red_22[0][valid_geom_mask].astype(float)
                n_22_vec = nir_22[0][valid_geom_mask].astype(float)

                pct_22, veg_22, _ = evaluar_corredor_espectral(r_22_vec, n_22_vec, swir_22, scl_22, veg_previa)

                # Evaluar también corredor central de 10m (eje conductores)
                red_10, out_trans_10 = mask(src_red, [franja_10m_raster], crop=True, filled=True, nodata=0)
                nir_10, _ = mask(src_nir, [franja_10m_raster], crop=True, filled=True, nodata=0)
                valid_10 = red_10[0] > 0

                if src_scl is not None and np.any(valid_10):
                    scl_g10 = leer_banda_alineada(src_scl, out_trans_10, red_10[0].shape, src_red.crs)
                    scl_10 = scl_g10[valid_10].astype(int) if scl_g10 is not None else np.full(np.sum(valid_10), 4)
                else:
                    scl_10 = np.full(np.sum(valid_10), 4)

                if src_swir is not None and np.any(valid_10):
                    swir_g10 = leer_banda_alineada(src_swir, out_trans_10, red_10[0].shape, src_red.crs)
                    swir_10 = swir_g10[valid_10] if swir_g10 is not None else None
                else:
                    swir_10 = None

                pct_10, veg_10, _ = evaluar_corredor_espectral(
                    red_10[0][valid_10].astype(float),
                    nir_10[0][valid_10].astype(float),
                    swir_10,
                    scl_10,
                    veg_previa,
                )

                if pct_22 is None:
                    vanos_omitidos_nubes += 1
                    pct_final = pct_previo
                    fecha_final = fecha_previa
                    veg_final = veg_previa
                else:
                    if pct_10 is not None:
                        pct_final = round(0.70 * pct_10 + 0.30 * pct_22, 1)
                        veg_final = veg_10
                    else:
                        pct_final = pct_22
                        veg_final = veg_22
                    fecha_final = fecha_imagen

            except Exception:
                pct_final = pct_previo
                fecha_final = fecha_previa
                veg_final = veg_previa

            actualizaciones.append({
                "linea": str(p1["linea"]),
                "piquete": p1_id,
                "orden": int(p1["orden"]),
                "latitud": float(p1["latitud"]),
                "longitud": float(p1["longitud"]),
                "dist_km": float(p1.get("dist_km") or 0.0),
                "tipo": str(p1.get("tipo") or ""),
                "estructura": str(p1.get("estructura") or ""),
                "pct_desmonte": pct_final,
                "fecha_limpieza": fecha_final,
                "tipo_vegetacion": veg_final,
            })

        if src_swir is not None:
            src_swir.close()
        if src_scl is not None:
            src_scl.close()

    if actualizaciones:
        for idx in range(0, len(actualizaciones), 500):
            sb.table("piquetes_gps").upsert(actualizaciones[idx : idx + 500], on_conflict="linea,piquete").execute()
        print(
            f"✅ [{sigla_linea}] {len(actualizaciones)} vanos procesados | "
            f"Omitidos por nube/sombra local: {vanos_omitidos_nubes}."
        )


if __name__ == "__main__":
    print("🚀 Iniciando análisis satelital Sentinel-2 (SCL + NDVI + NDMI) para Región Centro...")
    for sigla in LINEAS_REGION_CENTRO:
        try:
            analizar_linea_sentinel2(sigla)
        except Exception as err:
            print(f"❌ Error procesando {sigla}: {err}")
    print("🏁 Proceso finalizado.")
