# -*- coding: utf-8 -*-
"""
Preprocesado geográfico del dashboard
=====================================

Se ejecuta UNA VEZ (o cuando cambien los datos de origen) y deja en
`data_cache/` los ficheros ligeros que la app lee al instante:

  municipios.parquet          límites municipales de los 8.132 municipios,
                              a resolución completa, en WGS84
  grids_es_1km.parquet        rejillas de 1 km habitadas de España, ya
                              asignadas a su municipio, en WGS84
  municipios_peninbal.geojson geometría SIMPLIFICADA para el mapa nacional
  municipios_canarias.geojson ídem, para el recuadro de Canarias

Por qué hace falta: la rejilla europea original ocupa 20 GB -su tabla de
atributos sola son 18 GB, con 7.055.226 celdas y ocho campos de texto de 254
bytes cada uno- y no se puede abrir en una app interactiva. Aquí se lee una
sola vez leyendo únicamente las cuatro columnas necesarias, recortando por el
rectángulo de España y filtrando por país y población, y se deja en unos pocos
MB.

Ejecutar con:
    py -3.10 preprocessGrids.py
"""

from __future__ import annotations

import json
import struct
import sys
import time
from pathlib import Path

import geopandas as gpd
import numpy as np
import pandas as pd
import pyogrio
from pyproj import CRS, Transformer

# ==========================================================================
# CONFIGURACIÓN
# ==========================================================================

BASE_DIR = Path(__file__).resolve().parent
CACHE_DIR = BASE_DIR / "data_cache"
LIMITES_DIR = BASE_DIR / "lineas_limite"

# La rejilla vive fuera del proyecto y NO se toca: solo se lee.
GRID_SHP = Path(r"C:\Users\Víctor Díaz Barba\Documents\QGIS\GridEuropa1km.shp")

# El IGN separa Canarias del resto en dos carpetas con proyecciones distintas.
CAPAS_MUNICIPALES = [
    # carpeta, subcarpeta/fichero
    ("SHP_ETRS89", "recintos_municipales_inspire_peninbal_etrs89"),   # EPSG:4258
    ("SHP_WGS84",  "recintos_municipales_inspire_canarias_wgs84"),    # EPSG:4326
]

COL_POBLACION = "TOT_P_2021"     # población de la celda en 2021
COL_PAIS = "NUTS2021_0"          # país NUTS0 de la celda
COL_PAIS_ALT = "CNTR_ID"         # código de país, con más detalle en fronteras

# Las celdas fronterizas llevan un código compuesto, y NO siempre con España
# delante: en la frontera con Portugal el valor real es 'PT-ES', no 'ES-PT'.
# Enumerar combinaciones se presta a olvidos -así se perdían 230 celdas con
# 33.537 habitantes, entre ellas las dos pobladas de La Alamedilla-, así que se
# filtra por subcadena. Ningún código NUTS0 de dos letras contiene "ES" salvo
# el propio ES, así que no hay falsos positivos. Se mira también CNTR_ID porque
# recoge fronteras que NUTS0 no distingue (AD-ES con Andorra, GI-ES con
# Gibraltar). Lo que sobre lo descarta después el cruce con los municipios.
PATRON_PAIS = "%ES%"

# Rectángulo que envuelve España incluidas Canarias, en lon/lat.
BBOX_ESPANA_LONLAT = (-18.4, 27.4, 4.7, 44.1)

CRS_SALIDA = "EPSG:4326"         # lo que necesitan los mapas web

# --- Geometría simplificada para el mapa nacional -------------------------
# Los 8.132 municipios a resolución completa son 4,4 millones de vértices y
# 71 MB: imposible mandarlos al navegador. A escala nacional España mide unos
# 15,5 grados de longitud, así que en un lienzo de 1.400 px un pixel son
# 0,011 grados. Con 0,004 grados de tolerancia (444 m, 0,36 px) la
# simplificación es sub-pixel y el GeoJSON baja a 3,9 MB.
TOLERANCIA_SIMPLIFICADO = 0.004
# Coordenadas a cuatro decimales (~11 m): recorta el texto del GeoJSON a la
# mitad sin efecto visible.
PRECISION_COORDENADAS = 1e-4
# Provincias canarias, que van en el recuadro aparte. Se separan en dos
# ficheros porque Plotly incrusta el GeoJSON en cada traza: con uno solo, el
# mapa principal y el recuadro duplicarían los 3,9 MB.
PROVINCIAS_CANARIAS = ("35", "38")
# Área mínima de un trozo de provincia para dibujar su contorno, en grados
# cuadrados. Un pixel a escala nacional son 0,011 grados, así que por debajo de
# 0,011² el trozo no llega a verse.
AREA_MINIMA_PARTE = 0.011 ** 2

# --- Geometría ligera para el mapa municipal de la Ficha ------------------
# municipios.parquet a resolución completa pesa 68 MB, más de lo que conviene
# subir a un repositorio. La Ficha solo lee de él un municipio a la vez, para
# dibujar su contorno sobre el mapa de calles, y a esa escala el detalle
# submétrico del IGN no se ve. Aquí se escribe una copia ligera para subir; el
# original se queda intacto en local como fuente de la que regenerar todo.
#
# La clave es que la tolerancia sea proporcional al tamaño de cada municipio,
# no un valor único: en la Ficha cada uno ocupa los mismos ~700 px de ancho,
# pero miden desde 0,0014 grados (Emperador) hasta 1,06 (Cáceres). Una
# tolerancia fija sub-pixel para Cáceres deformaría Emperador en 98 px.
DIVISOR_SIMPLIFICADO_MUNICIPAL = 5600
# Coordenadas a seis decimales (~11 cm). No se nota ni en el municipio más
# pequeño, y como todas comparten la misma rejilla decimal zstd comprime el
# fichero a poco más de la tercera parte.
PRECISION_MUNICIPAL = 1e-6


# ==========================================================================
# MUNICIPIOS
# ==========================================================================


def codigos_del_censo() -> set[str]:
    """Los 8.132 códigos INE del censo, para validar la geometría.

    Hace falta porque el fichero peninbal del IGN trae 88 polígonos que NO son
    municipios: comunidades de villa y tierra y mancomunidades ("Comunero de
    Ansó y Hecho", "Cuarto del Madroño"...) a las que asigna códigos ficticios
    53xxx y 54xxx. Cruzar contra el censo los descarta sin ambigüedad.
    """
    parquet = CACHE_DIR / "datos_largo.parquet"
    if parquet.exists():
        df = pd.read_parquet(parquet, columns=["Territorio"])
    else:
        df = pd.read_excel(BASE_DIR / "DatosCensoINE.xlsx",
                           sheet_name="Datos_largo", usecols=["Territorio"],
                           engine="openpyxl")
    cod = df["Territorio"].astype(str).str.extract(r"^(\d{5})\s")[0]
    return set(cod.dropna())


def construir_municipios(censo: set[str]) -> gpd.GeoDataFrame:
    """Une las dos capas del IGN y deja un municipio por fila, en WGS84."""
    trozos = []
    for carpeta, capa in CAPAS_MUNICIPALES:
        ruta = LIMITES_DIR / carpeta / capa / f"{capa}.shp"
        if not ruta.exists():
            raise FileNotFoundError(ruta)
        gdf = pyogrio.read_dataframe(ruta, columns=["NATCODE", "NAMEUNIT"])
        # NATCODE = 34 + CCAA(2) + provincia(2) + código INE(5).
        gdf["cod"] = gdf["NATCODE"].astype(str).str[-5:]
        print(f"  {capa}: {len(gdf)} polígonos, CRS {gdf.crs.to_string()}")
        trozos.append(gdf.to_crs(CRS_SALIDA))

    muni = pd.concat(trozos, ignore_index=True)
    muni = gpd.GeoDataFrame(muni, geometry="geometry", crs=CRS_SALIDA)

    antes = len(muni)
    muni = muni[muni["cod"].isin(censo)].copy()
    print(f"  descartados {antes - len(muni)} polígonos que no son municipios")

    # Un municipio puede venir en varias piezas (exclaves); se unen en una.
    if muni["cod"].duplicated().any():
        n = int(muni["cod"].duplicated().sum())
        muni = muni.dissolve(by="cod", aggfunc="first").reset_index()
        print(f"  unidas {n} piezas de municipios con exclaves")

    return muni[["cod", "NAMEUNIT", "geometry"]]


# ==========================================================================
# REJILLA DE 1 KM
# ==========================================================================


def leer_rejilla_espana() -> gpd.GeoDataFrame:
    """Lee de la rejilla europea solo las celdas habitadas de España."""
    info = pyogrio.read_info(GRID_SHP)
    crs_grid = CRS.from_user_input(info["crs"])
    print(f"  origen: {info['features']:,} celdas, CRS "
          f"{crs_grid.name} (EPSG:{crs_grid.to_epsg()})")

    tr = Transformer.from_crs("EPSG:4326", crs_grid, always_xy=True)
    bbox = tr.transform_bounds(*BBOX_ESPANA_LONLAT, densify_pts=51)
    print(f"  recorte: {tuple(round(v) for v in bbox)}")

    t0 = time.time()
    gdf = pyogrio.read_dataframe(
        GRID_SHP,
        # Solo cuatro columnas: leer las 20 supondría arrastrar los ocho campos
        # de texto de 254 bytes que hacen que el .dbf pese 18 GB.
        columns=["GRD_ID", COL_POBLACION, COL_PAIS, COL_PAIS_ALT],
        where=(f"{COL_POBLACION} > 0 AND ("
               f"{COL_PAIS} LIKE '{PATRON_PAIS}' OR "
               f"{COL_PAIS_ALT} LIKE '{PATRON_PAIS}')"),
        bbox=bbox,
    )
    print(f"  leídas {len(gdf):,} celdas habitadas en {time.time()-t0:,.0f} s")
    print(f"  {COL_PAIS}: "
          + ", ".join(f"{k}={v:,}"
                      for k, v in gdf[COL_PAIS].value_counts().items()))
    return gdf


def asignar_municipio(grid: gpd.GeoDataFrame,
                      muni: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Asigna cada celda a todos los municipios con los que SOLAPA.

    Basta que la celda entre un poco en el término para que le corresponda: no
    hace falta que quepa entera. Consecuencias:

      · Una celda a caballo entre varios municipios aparece en todos ellos, así
        que sale una fila por cada pareja celda-municipio y el total de filas
        supera al número de celdas.
      · Por lo mismo, sumar la población de todas las filas cuenta de más a la
        gente de las celdas compartidas. Dentro de un municipio la suma sí es
        coherente; lo que no se puede es sumar entre municipios.
      · A cambio, ningún municipio se queda sin mapa. Con el criterio del
        centro de celda, 16 municipios más pequeños que 1 km² se quedaban
        vacíos, entre ellos Barañáin, con 19.588 habitantes en 1,39 km².
    """
    muni_proj = muni.to_crs(grid.crs)

    t0 = time.time()
    unido = gpd.sjoin(grid, muni_proj[["cod", "geometry"]],
                      how="inner", predicate="intersects")
    unido = unido.drop(columns=["index_right"])
    print(f"  unión espacial por solape en {time.time()-t0:,.0f} s")
    print(f"  {len(unido):,} parejas celda-municipio de {len(grid):,} celdas "
          f"({len(unido)/max(len(grid),1):.2f} municipios por celda)")
    return unido


# ==========================================================================
# GEOMETRÍA SIMPLIFICADA PARA EL MAPA NACIONAL
# ==========================================================================


def escribir_geojson_simplificado(muni: gpd.GeoDataFrame) -> None:
    """Escribe la geometría simplificada, separada en península y Canarias."""
    from shapely import set_precision

    g = muni[["cod", "geometry"]].copy()
    vertices_antes = int(g.geometry.count_coordinates().sum())

    g["geometry"] = g.geometry.simplify(TOLERANCIA_SIMPLIFICADO,
                                        preserve_topology=True)
    g["geometry"] = set_precision(g.geometry.values, PRECISION_COORDENADAS)
    vertices_despues = int(g.geometry.count_coordinates().sum())
    print(f"  vértices: {vertices_antes:,} -> {vertices_despues:,} "
          f"({100*vertices_despues/vertices_antes:.1f} %)")

    invalidas = int((~g.geometry.is_valid).sum())
    vacias = int(g.geometry.is_empty.sum())
    if invalidas or vacias:
        print(f"  !! AVISO: {invalidas} inválidas, {vacias} vacías")

    es_canaria = g["cod"].str[:2].isin(PROVINCIAS_CANARIAS)
    for nombre, trozo in [("municipios_peninbal", g[~es_canaria]),
                          ("municipios_canarias", g[es_canaria])]:
        # El índice es el código INE, así que en el GeoJSON queda como "id" de
        # cada feature: es la clave con la que Plotly une datos y geometría.
        texto = trozo.set_index("cod").to_json()
        salida = CACHE_DIR / f"{nombre}.geojson"
        salida.write_text(texto, encoding="utf-8")
        print(f"  escrito {salida.name}: {len(trozo):,} municipios, "
              f"{len(texto)/1e6:.1f} MB")


def escribir_municipios_ligero(muni: gpd.GeoDataFrame) -> None:
    """Copia ligera de municipios.parquet, la que se sube al servidor.

    No sobrescribe el original: escribe municipios_ligero.parquet aparte. El
    dashboard prefiere este si existe y si no cae al completo, así que en
    local funcionan los dos y en el servidor basta con subir el ligero.
    """
    from shapely import set_precision

    # Solo las dos columnas que lee el dashboard: NAMEUNIT no se usa, el
    # nombre del municipio sale del censo.
    g = muni[["cod", "geometry"]].copy()
    vertices_antes = int(g.geometry.count_coordinates().sum())

    limites = g.geometry.bounds
    tamano = np.maximum(limites["maxx"] - limites["minx"],
                        limites["maxy"] - limites["miny"])
    tolerancias = tamano / DIVISOR_SIMPLIFICADO_MUNICIPAL
    g["geometry"] = [geometria.simplify(tol, preserve_topology=True)
                     for geometria, tol in zip(g.geometry, tolerancias)]
    g["geometry"] = set_precision(g.geometry.values, PRECISION_MUNICIPAL)

    vertices_despues = int(g.geometry.count_coordinates().sum())
    print(f"  vértices: {vertices_antes:,} -> {vertices_despues:,} "
          f"({100*vertices_despues/vertices_antes:.1f} %)")
    invalidas = int((~g.geometry.is_valid).sum())
    vacias = int(g.geometry.is_empty.sum())
    if invalidas or vacias:
        print(f"  !! AVISO: {invalidas} inválidas, {vacias} vacías")

    # Ordenado por código: así el filtro al leer descarta grupos de filas
    # enteros en vez de recorrer el fichero completo.
    g = g.sort_values("cod")
    salida = CACHE_DIR / "municipios_ligero.parquet"
    g.to_parquet(salida, index=False, compression="zstd", compression_level=19)
    print(f"  escrito {salida.name}: {len(g):,} municipios, "
          f"{salida.stat().st_size/1e6:,.1f} MB")


def escribir_lineas_provinciales(muni: gpd.GeoDataFrame) -> None:
    """Límites provinciales como líneas, para superponerlos al mapa nacional.

    OJO AL ORDEN: se disuelven los municipios a RESOLUCIÓN COMPLETA y solo
    después se simplifica el resultado. Al revés no funciona. Simplificar antes
    deja las aristas compartidas de municipios vecinos sin encajar -la
    simplificación recorre cada anillo por separado y no las recorta igual- y
    la unión abre 16.983 rendijas en vez de los 71 huecos reales.

    Se guardan como listas de coordenadas con None de separador, que es lo que
    consume una traza de líneas, en vez de como GeoJSON: pesa mucho menos.
    """
    from shapely import set_precision

    g = muni[["cod", "geometry"]].copy()
    g["prov"] = g["cod"].str[:2]

    t0 = time.time()
    prov = g.dissolve(by="prov")
    print(f"  {len(prov)} provincias disueltas en {time.time()-t0:,.0f} s")

    prov["geometry"] = prov.geometry.simplify(TOLERANCIA_SIMPLIFICADO,
                                              preserve_topology=True)
    prov["geometry"] = set_precision(prov.geometry.values,
                                     PRECISION_COORDENADAS)

    es_canaria = prov.index.isin(PROVINCIAS_CANARIAS)
    salida: dict[str, dict[str, list]] = {}
    descartadas = 0

    for nombre, trozo in [("peninbal", prov[~es_canaria]),
                          ("canarias", prov[es_canaria])]:
        lon: list = []
        lat: list = []
        for geom in trozo.geometry:
            partes = (list(geom.geoms) if geom.geom_type == "MultiPolygon"
                      else [geom])
            for parte in partes:
                # Los islotes de la costa son cientos de polígonos diminutos.
                # Por debajo de un pixel a escala nacional no se ven, así que
                # solo añadirían peso.
                if parte.area < AREA_MINIMA_PARTE:
                    descartadas += 1
                    continue
                for anillo in [parte.exterior, *parte.interiors]:
                    xs, ys = anillo.coords.xy
                    lon.extend(xs)
                    lon.append(None)
                    lat.extend(ys)
                    lat.append(None)
        salida[nombre] = {"lon": lon, "lat": lat}
        print(f"  {nombre}: {len(lon):,} puntos de línea")

    print(f"  partes descartadas por diminutas: {descartadas:,}")
    ruta = CACHE_DIR / "provincias_lineas.json"
    ruta.write_text(json.dumps(salida), encoding="utf-8")
    print(f"  escrito {ruta.name} ({ruta.stat().st_size/1e6:.2f} MB)")


# ==========================================================================
# MAIN
# ==========================================================================


def main() -> None:
    CACHE_DIR.mkdir(exist_ok=True)

    print("\n[1/5] Municipios")
    censo = codigos_del_censo()
    print(f"  códigos en el censo: {len(censo)}")
    muni = construir_municipios(censo)
    print(f"  -> {len(muni)} municipios")
    faltan = censo - set(muni["cod"])
    if faltan:
        print(f"  !! AVISO: {len(faltan)} municipios del censo sin geometría: "
              f"{sorted(faltan)[:8]}")

    salida_muni = CACHE_DIR / "municipios.parquet"
    muni.to_parquet(salida_muni, index=False)
    print(f"  escrito {salida_muni.name} "
          f"({salida_muni.stat().st_size/1e6:,.1f} MB)")

    print("\n[2/5] Copia ligera de los municipios, para subir")
    escribir_municipios_ligero(muni)

    print("\n[3/5] Geometría simplificada para el mapa nacional")
    escribir_geojson_simplificado(muni)
    print("  --- límites provinciales ---")
    escribir_lineas_provinciales(muni)

    print("\n[4/5] Rejilla de 1 km")
    grid = leer_rejilla_espana()

    print("\n[5/5] Asignación de cada celda a los municipios que solapa")
    grid = asignar_municipio(grid, muni)
    grid = grid.to_crs(CRS_SALIDA)
    grid = grid.rename(columns={COL_POBLACION: "poblacion"})
    # Ordenado por municipio: así el filtro por código al leer el parquet puede
    # descartar grupos de filas enteros en vez de recorrerlo todo.
    grid = grid.sort_values("cod")[["cod", "GRD_ID", "poblacion", "geometry"]]

    salida_grid = CACHE_DIR / "grids_es_1km.parquet"
    grid.to_parquet(salida_grid, index=False)
    print(f"  -> {len(grid):,} parejas celda-municipio")
    print(f"  celdas distintas: {grid['GRD_ID'].nunique():,}")
    print(f"  municipios con al menos una celda: {grid['cod'].nunique():,} "
          f"de {len(muni):,}")
    sin_celda = sorted(set(muni["cod"]) - set(grid["cod"]))
    if sin_celda:
        print(f"  !! {len(sin_celda)} municipios sin celda habitada: "
              f"{sin_celda[:8]}")
    print(f"  escrito {salida_grid.name} "
          f"({salida_grid.stat().st_size/1e6:,.1f} MB)")

    print("\nListo.")


if __name__ == "__main__":
    main()
