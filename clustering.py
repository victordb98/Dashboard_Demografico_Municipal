# -*- coding: utf-8 -*-
"""
Clustering de municipios — Paso 1: construcción del dataframe
=============================================================

Deja en `data_cache/` una tabla con un municipio por fila y una variable
demográfica por columna, lista para agrupar.

Las definiciones NO se reimplementan: el script importa `app.py` y usa sus
propias funciones, así que los indicadores del clustering y los que muestra el
dashboard no pueden divergir. Cuesta unos 60 segundos recorrer los 8.132
municipios llamando a `calcular_kpis`, y es tiempo bien gastado.

Tres decisiones tomadas al construirla:

  1. TASA DE ACTIVIDAD sobre la población de 16 y más años, que es la
     definición del INE. Con el denominador de 15 a 64 -el que usa la tasa de
     empleo- 77 municipios pasarían del 100 % y al estandarizar actuarían como
     atípicos tirando de los grupos.

  2. "NO CONSTA" FUERA DEL DENOMINADOR en estado civil y ocupación. Si se deja
     dentro, la tasa de no respuesta del municipio contamina todas sus
     variables a la vez: Torre del Burgo tiene el 62 % del estado civil sin
     constar, y su porcentaje de solteros sale 15,4 % con "No consta" en la
     base frente a 40,6 % sin ella. La primera cifra no habla del municipio,
     habla del registro.

  3. SOLO MUNICIPIOS COMPLETOS. Se conservan los 7.502 que tienen las
     diecinueve variables y se descartan 630 (7,7 %), casi todos los más
     pequeños, a los que el INE no publica los datos laborales.

Ejecutar con:
    py -3.10 clustering.py
"""

from __future__ import annotations

import logging
import sys
import time
from pathlib import Path

import pandas as pd

# Streamlit configura su logging al importarse, así que hay que bajarle el
# nivel DESPUÉS de importarlo: si no, los decoradores de caché de app.py
# llenan la salida de avisos de "no runtime found".
import streamlit  # noqa: F401
logging.getLogger("streamlit").setLevel(logging.ERROR)

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import app  # noqa: E402


# ==========================================================================
# §1  CONFIGURACIÓN
# ==========================================================================

CACHE_DIR = BASE_DIR / "data_cache"
SALIDA_PARQUET = CACHE_DIR / "clustering_municipios.parquet"
SALIDA_CSV = CACHE_DIR / "clustering_municipios.csv"

# Columnas que identifican la fila, no son variables de agrupación.
IDENTIFICADORAS = ["cod", "nombre"]

# Las diecinueve variables, con su descripción. El orden es el de la tabla.
VARIABLES = [
    ("pct_menores_18",        "% población menor de 18 años"),
    ("pct_mayores_67",        "% población de 67 y más años"),
    ("pct_migrantes",         "% nacidos en el extranjero"),
    ("indice_fecundidad",     "Hijos por mujer"),
    ("renta_bruta_persona",   "Renta bruta media por persona (€)"),
    ("pct_solteros",          "% solteros"),
    ("pct_casados",           "% casados"),
    ("pct_divorciados",       "% divorciados o separados"),
    ("pct_cuenta_propia",     "% trabajadores por cuenta propia"),
    ("pct_agricultura",       "% ocupados en agricultura, ganadería y pesca"),
    ("pct_secundario",        "% ocupados en industria y construcción"),
    ("pct_directores",        "% directores, gerentes y técnicos"),
    ("pct_cualificados",      "% trabajadores cualificados y operarios"),
    ("pct_primaria_inferior", "% con educación primaria o inferior"),
    ("pct_secundaria_2a",     "% con segunda etapa de secundaria"),
    ("pct_superior",          "% con educación superior"),
    ("tasa_paro",             "Tasa de paro"),
    ("tasa_empleo",           "Tasa de empleo"),
    ("tasa_actividad",        "Tasa de actividad"),
]

# Columnas del Excel agrupadas por hoja. La BASE de cada bloque de porcentajes
# es la suma de estas columnas, que excluye "No consta" a propósito.
COLS_CIVIL = {
    "pct_solteros": "Soltero/a",
    "pct_casados": "Casado/a",
    "pct_divorciados": "Divorciado/a o separado/a",
}
BASE_CIVIL = ["Soltero/a", "Casado/a", "Viudo/a", "Divorciado/a o separado/a"]

COLS_OCUPACION = {
    "pct_directores":
        "Directores/gerentes y profesionales/técnicos de nivel medio o alto",
    "pct_cualificados":
        "Trabajadores cualificados y oficiales/operarios de nivel bajo",
}
BASE_OCUPACION = [
    "Directores/gerentes y profesionales/técnicos de nivel medio o alto",
    "Trabajadores cualificados y oficiales/operarios de nivel bajo",
    "Ocupaciones elementales",
]

COLS_ACTIVIDAD = {
    "pct_agricultura": ["Agricultura, ganadería y pesca"],
    "pct_secundario": ["Industria", "Construcción"],
}
BASE_ACTIVIDAD = ["Agricultura, ganadería y pesca", "Industria",
                  "Construcción", "Servicios"]

COLS_ESTUDIOS = {
    "pct_primaria_inferior": "Educación primaria e inferior",
    "pct_secundaria_2a": ("Segunda etapa de Educación Secundaria y Educación"
                          " Postsecundaria no Superior"),
    "pct_superior": "Educación superior",
}
BASE_ESTUDIOS = [
    "Educación primaria e inferior",
    "Primera etapa de Educación Secundaria y similar",
    ("Segunda etapa de Educación Secundaria y Educación"
     " Postsecundaria no Superior"),
    "Educación superior",
]

COL_CUENTA_PROPIA = "Trabajador por cuenta propia"
BASE_SITUACION_PROF = [
    "Trabajador por cuenta propia",
    "Trabajador por cuenta ajena y otra situación",
]


# ==========================================================================
# §2  BLOQUES DE VARIABLES
# ==========================================================================


def _porcentajes(tabla: pd.DataFrame, columnas: dict,
                 base_cols: list[str]) -> pd.DataFrame:
    """Porcentaje de cada columna sobre la suma de `base_cols`.

    La base excluye "No consta", así que los porcentajes de cada bloque suman
    100 entre las categorías conocidas.
    """
    datos = tabla.copy()
    for c in set(base_cols) | {c for v in columnas.values()
                               for c in ([v] if isinstance(v, str) else v)}:
        datos[c] = pd.to_numeric(datos[c], errors="coerce").fillna(0.0)

    base = datos[base_cols].sum(axis=1)
    base = base.where(base > 0)

    salida = pd.DataFrame(index=datos.index)
    for nombre, cols in columnas.items():
        cols = [cols] if isinstance(cols, str) else cols
        salida[nombre] = 100 * datos[cols].sum(axis=1) / base
    return salida


def bloque_demografico(muni: pd.DataFrame) -> pd.DataFrame:
    """Edades, migrantes y fecundidad, con la función del dashboard.

    Se recorre municipio a municipio llamando a `app.calcular_kpis` en vez de
    vectorizar: es más lento, pero así el reparto 3/5 de los umbrales de 18 y
    67 años y el índice de fecundidad son literalmente los mismos que muestra
    la ficha municipal, sin código duplicado que pueda divergir.
    """
    t0 = time.time()
    por_codigo = {cod: grupo for cod, grupo
                  in muni.groupby("cod", observed=True)}

    filas = {}
    for i, (cod, grupo) in enumerate(por_codigo.items(), 1):
        k = app.calcular_kpis(grupo)
        filas[cod] = {
            "pct_menores_18": k["pct_menores18"],
            "pct_mayores_67": k["pct_mayores67"],
            "pct_migrantes": k["pct_extranjero"],
            "indice_fecundidad": k["fecundidad"],
        }
        if i % 2000 == 0:
            print(f"    {i:,} de {len(por_codigo):,}...")

    salida = pd.DataFrame.from_dict(filas, orient="index")
    salida.index.name = "cod"
    print(f"    {len(salida):,} municipios en {time.time()-t0:,.0f} s")
    return salida


def bloque_laboral(mtime: float, act: pd.DataFrame) -> pd.DataFrame:
    """Tasas de paro, empleo y actividad.

    Paro y empleo salen de `app.tasas_laborales_municipales`, que es la misma
    función que alimenta los mapas y está contrastada contra la ficha.

    La tasa de actividad no existe en el dashboard y se define aquí: activos
    entre la población de 16 y más años, que es el Total de la hoja de
    actividad y la definición del INE.
    """
    tasas = app.tasas_laborales_municipales(mtime)
    base_16 = pd.to_numeric(act["Total"], errors="coerce")
    activos = tasas["ocupados"] + tasas["parados"]

    return pd.DataFrame({
        "tasa_paro": tasas["paro"],
        "tasa_empleo": tasas["empleo"],
        "tasa_actividad": 100 * activos / base_16.where(base_16 > 0),
    })


# ==========================================================================
# §3  ENSAMBLADO
# ==========================================================================


def construir_dataframe() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (completos, todos). `todos` conserva los huecos, para el informe."""
    mtime = app.EXCEL_PATH.stat().st_mtime

    print("[1/4] Cargando los datos del censo")
    muni, _ref = app.cargar_poblacion(mtime)
    catalogo = (muni.drop_duplicates("cod")[["cod", "nombre"]]
                .set_index("cod").sort_index())
    renta, _ctx = app.cargar_renta(mtime)
    act = app.cargar_tabla_municipal("relacion_actividad", mtime)
    civil = app.cargar_tabla_municipal("estado_civil", mtime)
    sitprof = app.cargar_tabla_municipal("situacion_prof", mtime)
    actividad = app.cargar_tabla_municipal("actividad", mtime)
    ocupacion = app.cargar_tabla_municipal("ocupacion", mtime)
    estudios = app.cargar_tabla_municipal("nivel_estudios", mtime)
    print(f"    {len(catalogo):,} municipios en el censo")

    print("[2/4] Edades, migrantes y fecundidad")
    demo = bloque_demografico(muni)

    print("[3/4] Renta, estado civil, ocupación, actividad y estudios")
    trozos = [
        demo,
        pd.DataFrame({"renta_bruta_persona": renta}),
        _porcentajes(civil, COLS_CIVIL, BASE_CIVIL),
        _porcentajes(sitprof, {"pct_cuenta_propia": COL_CUENTA_PROPIA},
                     BASE_SITUACION_PROF),
        _porcentajes(actividad, COLS_ACTIVIDAD, BASE_ACTIVIDAD),
        _porcentajes(ocupacion, COLS_OCUPACION, BASE_OCUPACION),
        _porcentajes(estudios, COLS_ESTUDIOS, BASE_ESTUDIOS),
        bloque_laboral(mtime, act),
    ]

    print("[4/4] Ensamblando")
    todos = catalogo.join(trozos, how="left")
    columnas = [v for v, _d in VARIABLES]
    todos = todos[["nombre"] + columnas]

    completos = todos.dropna(subset=columnas).copy()
    return completos, todos


def informar(completos: pd.DataFrame, todos: pd.DataFrame) -> None:
    """Resumen de cobertura y de las variables, para poder confiar en la tabla."""
    columnas = [v for v, _d in VARIABLES]

    print("\n" + "=" * 74)
    print("COBERTURA")
    print("=" * 74)
    print(f"  municipios en el censo : {len(todos):,}")
    print(f"  con las 19 variables   : {len(completos):,}")
    print(f"  descartados            : {len(todos)-len(completos):,} "
          f"({100*(len(todos)-len(completos))/len(todos):.1f} %)")

    print("\n  huecos por variable, antes de descartar:")
    faltan = todos[columnas].isna().sum()
    for var, n in faltan[faltan > 0].sort_values(ascending=False).items():
        print(f"    {var:24s} {n:>5,} sin dato")
    if not (faltan > 0).any():
        print("    ninguno")

    # ¿Quiénes se pierden? Importa saberlo: si son todos diminutos, el
    # clustering se queda sin ese perfil.
    perdidos = todos.index.difference(completos.index)
    if len(perdidos):
        amb = todos.loc[perdidos]
        print(f"\n  los descartados son, sobre todo, municipios pequeños:")
        print(f"    ejemplos: "
              + ", ".join(amb['nombre'].head(4).tolist()))

    print("\n" + "=" * 74)
    print("VARIABLES")
    print("=" * 74)
    desc = completos[columnas].describe().T
    print(f"  {'variable':24s} {'media':>9s} {'desv':>8s} "
          f"{'min':>8s} {'mediana':>9s} {'max':>9s}")
    for var, etiqueta in VARIABLES:
        f = desc.loc[var]
        print(f"  {var:24s} {f['mean']:9.2f} {f['std']:8.2f} "
              f"{f['min']:8.2f} {f['50%']:9.2f} {f['max']:9.2f}")

    print("\n  descripción de cada variable:")
    for var, etiqueta in VARIABLES:
        print(f"    {var:24s} {etiqueta}")

    # Comprobación de coherencia: los bloques de porcentajes que deben sumar
    # 100 entre las categorías conocidas.
    print("\n" + "=" * 74)
    print("COMPROBACIONES")
    print("=" * 74)
    est = completos[["pct_primaria_inferior", "pct_secundaria_2a",
                     "pct_superior"]].sum(axis=1)
    print(f"  estudios: los tres niveles pedidos suman entre "
          f"{est.min():.1f} y {est.max():.1f} % (falta la 1ª etapa de "
          f"secundaria, que no se pidió)")
    civ = completos[["pct_solteros", "pct_casados",
                     "pct_divorciados"]].sum(axis=1)
    print(f"  estado civil: los tres pedidos suman entre {civ.min():.1f} y "
          f"{civ.max():.1f} % (falta viudos, que no se pidió)")
    fuera = int((completos["tasa_empleo"] > 100).sum())
    print(f"  tasa de empleo por encima del 100 %: {fuera} municipios "
          f"(numerador de 16 y más, denominador de 15 a 64)")
    print(f"  tasa de actividad por encima del 100 %: "
          f"{int((completos['tasa_actividad'] > 100).sum())} municipios")


# ==========================================================================
# §4  MAIN
# ==========================================================================


def main() -> None:
    CACHE_DIR.mkdir(exist_ok=True)
    completos, todos = construir_dataframe()
    informar(completos, todos)

    completos.to_parquet(SALIDA_PARQUET)
    # CSV con separadores españoles, para poder abrirlo en Excel sin tocar nada.
    completos.to_csv(SALIDA_CSV, sep=";", decimal=",", encoding="utf-8-sig")

    print("\n" + "=" * 74)
    print(f"  escrito {SALIDA_PARQUET.name} "
          f"({SALIDA_PARQUET.stat().st_size/1e6:.2f} MB)")
    print(f"  escrito {SALIDA_CSV.name} "
          f"({SALIDA_CSV.stat().st_size/1e6:.2f} MB)")
    print(f"  {len(completos):,} municipios x {len(VARIABLES)} variables")
    print("=" * 74)
    print(completos.shape)
    print(completos.columns.tolist())


if __name__ == "__main__":
    main()
