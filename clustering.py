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

import json
import logging
import sys
import time
from pathlib import Path

import numpy as np
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

# Etiquetas del corte en GRUPOS grupos, y la ficha del analisis. Son lo unico
# que necesita el dashboard: app.py no puede recalcular el arbol -scipy no
# esta en requirements.txt, es dependencia solo de este script- asi que lee
# estos dos ficheros y no reimplementa nada.
SALIDA_GRUPOS = CACHE_DIR / "clustering_grupos.parquet"
SALIDA_META = CACHE_DIR / "clustering_meta.json"

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
# §4  ESTANDARIZACIÓN
# ==========================================================================


def estandarizar(completos: pd.DataFrame) -> pd.DataFrame:
    """Las diecinueve variables en puntuaciones z, en una tabla nueva.

    z = (valor - media) / desviación típica, columna a columna. Se usa la
    desviación muestral (ddof=1, la que da pandas por defecto); con 7.502
    filas la diferencia con la poblacional es del 0,007 %.

    Hace falta porque las variables están en unidades incomparables: la renta
    va en decenas de miles de euros, los porcentajes en decenas y el índice de
    fecundidad en unidades. Cualquier agrupación basada en distancias la
    decidiría la renta ella sola, por ser la de números más grandes.

    `completos` no se modifica: lo que se devuelve es una tabla aparte.
    """
    columnas = [v for v, _d in VARIABLES]
    medias = completos[columnas].mean()
    desviaciones = completos[columnas].std()

    tipificadas = completos[columnas].sub(medias).div(desviaciones)
    # El nombre se conserva para poder identificar cada fila, pero no es una
    # variable y por tanto no se estandariza.
    tipificadas.insert(0, "nombre", completos["nombre"])
    return tipificadas


def informar_estandarizadas(tipificadas: pd.DataFrame,
                            completos: pd.DataFrame) -> None:
    """Comprueba la estandarización y señala los valores extremos."""
    columnas = [v for v, _d in VARIABLES]

    print("\n" + "=" * 74)
    print("ESTANDARIZACIÓN")
    print("=" * 74)
    print(f"  {'variable':24s} {'media':>9s} {'desv':>8s} "
          f"{'z min':>7s} {'z max':>7s}")
    print(f"  {'':24s} {'(usada)':>9s} {'(usada)':>8s}")
    for var in columnas:
        print(f"  {var:24s} {completos[var].mean():9.2f} "
              f"{completos[var].std():8.2f} "
              f"{tipificadas[var].min():7.2f} {tipificadas[var].max():7.2f}")

    # Comprobación de que ha salido bien: toda columna tipificada tiene que
    # quedar con media 0 y desviación 1.
    medias = tipificadas[columnas].mean().abs().max()
    desv = (tipificadas[columnas].std() - 1).abs().max()
    print(f"\n  comprobación: media máxima {medias:.2e}, "
          f"desviación más alejada de 1 en {desv:.2e}")

    # Los valores muy extremos importan: en una agrupación por distancias, un
    # municipio a 20 desviaciones tira de su grupo él solo.
    print("\n  colas largas (algún municipio por encima de 5 desviaciones):")
    extremos = tipificadas[columnas].abs().max()
    hay = False
    for var, z in extremos[extremos > 5].sort_values(ascending=False).items():
        cual = tipificadas[var].abs().idxmax()
        print(f"    {var:24s} z = {z:6.2f}  "
              f"({tipificadas.loc[cual, 'nombre']})")
        hay = True
    if not hay:
        print("    ninguna")


# ==========================================================================
# §5  MATRIZ DE DISIMILITUD
# ==========================================================================

# La tasa de actividad se queda fuera del cálculo de distancias. Es, casi por
# definición, una combinación de las otras dos tasas (activos = ocupados +
# parados), así que incluirla haría que el mercado laboral pesara tres veces
# en la distancia mientras el estado civil pesa una. Sigue en la tabla de
# variables: solo no se usa para medir.
VARIABLE_EXCLUIDA = "tasa_actividad"


def matriz_disimilitud(tipificadas: pd.DataFrame) -> tuple[np.ndarray,
                                                           list[str]]:
    """Distancias euclídeas entre cada par de municipios.

    Devuelve la matriz en forma condensada: en vez de las 56,3 millones de
    celdas de la matriz cuadrada, solo las 28,1 millones del triángulo
    superior, que es donde está toda la información (la matriz es simétrica y
    su diagonal es cero). Ocupa 215 MB en lugar de 429 MB, y es el formato que
    esperan directamente las funciones de agrupación jerárquica.

    Para leer una distancia concreta hay que convertirla con
    `scipy.spatial.distance.squareform`, o usar `distancia_entre()`.
    """
    from scipy.spatial.distance import pdist

    columnas = [v for v, _d in VARIABLES if v != VARIABLE_EXCLUIDA]
    datos = tipificadas[columnas].to_numpy()
    return pdist(datos, metric="euclidean"), columnas


def distancia_entre(condensada: np.ndarray, tipificadas: pd.DataFrame,
                    cod_a: str, cod_b: str) -> float:
    """Distancia entre dos municipios, buscándola en la matriz condensada.

    Se calcula la posición sin desplegar la matriz, que ocuparía el doble.
    """
    n = len(tipificadas)
    i, j = sorted((tipificadas.index.get_loc(cod_a),
                   tipificadas.index.get_loc(cod_b)))
    return float(condensada[n * i - i * (i + 1) // 2 + j - i - 1])


def informar_disimilitud(condensada: np.ndarray, columnas: list[str],
                         tipificadas: pd.DataFrame) -> None:
    """Resumen de la matriz, para comprobar que mide lo que debe."""
    n = len(tipificadas)

    print("\n" + "=" * 74)
    print("MATRIZ DE DISIMILITUD")
    print("=" * 74)
    print(f"  método      : distancia euclídea sobre variables estandarizadas")
    print(f"  variables   : {len(columnas)} de {len(VARIABLES)} "
          f"(fuera: {VARIABLE_EXCLUIDA})")
    print(f"  municipios  : {n:,}")
    print(f"  parejas     : {len(condensada):,}")
    print(f"  memoria     : {condensada.nbytes/2**20:.0f} MB "
          f"(condensada; la cuadrada serían {n*n*8/2**20:.0f})")

    print(f"\n  distancias: mínima {condensada.min():.3f}   "
          f"mediana {np.median(condensada):.3f}   "
          f"media {condensada.mean():.3f}   "
          f"máxima {condensada.max():.3f}")

    # La pareja más parecida y la más distinta. Sirve de comprobación: si la
    # más parecida no son dos municipios plausiblemente gemelos, algo falla.
    for etiqueta, pos in (("más parecidos", int(condensada.argmin())),
                          ("más distintos", int(condensada.argmax()))):
        # Inversa de la fórmula de `distancia_entre`: de la posición en el
        # triángulo a la pareja (i, j) que le corresponde.
        i = int(n - 2 - np.floor(np.sqrt(-8 * pos + 4 * n * (n - 1) - 7) / 2
                                 - 0.5))
        j = int(pos - n * i + i * (i + 1) // 2 + i + 1)
        print(f"\n  {etiqueta} ({condensada[pos]:.3f}): "
              f"{tipificadas['nombre'].iloc[i]} "
              f"({tipificadas.index[i]}) y "
              f"{tipificadas['nombre'].iloc[j]} "
              f"({tipificadas.index[j]})")

    # A quién le queda lejos todo el mundo: son los que, al agrupar, tienden a
    # quedarse solos en su propio grupo. Se calcula por bloques de filas para
    # no desplegar la matriz cuadrada, que costaría 429 MB de golpe.
    print("\n  los más aislados (mayor distancia media al resto):")
    from scipy.spatial.distance import cdist
    datos = tipificadas[columnas].to_numpy()
    sumas = np.empty(n)
    for ini in range(0, n, 512):
        bloque = cdist(datos[ini:ini + 512], datos)
        sumas[ini:ini + 512] = bloque.sum(axis=1)
    medias = sumas / (n - 1)
    for pos in np.argsort(-medias)[:5]:
        print(f"    {tipificadas['nombre'].iloc[pos]:<28} "
              f"{medias[pos]:6.2f}  (la media general es "
              f"{condensada.mean():.2f})")


# ==========================================================================
# §6  DENDROGRAMA
# ==========================================================================

FIGURAS_DIR = BASE_DIR / "figuras"
# Ramas que se muestran en el árbol recortado. Con 7.502 hojas el árbol
# completo es una mancha negra: no se distingue ninguna hoja porque cada una
# mide una centésima de pixel. El recortado agrupa las hojas y sí se lee.
RAMAS_RECORTADAS = 30


def enlazar(condensada: np.ndarray) -> np.ndarray:
    """Agrupación jerárquica por el método de Ward.

    Ward une en cada paso los dos grupos que menos aumentan la varianza
    interna, así que tiende a formar grupos de tamaño parecido y compactos.
    Es el que mejor funciona con variables estandarizadas como las nuestras.

    Se le pasa la matriz condensada de §5, ya calculada, en vez de la tabla de
    datos: el resultado es idéntico y no hay que recalcular las distancias.
    """
    import scipy.cluster.hierarchy as sch

    return sch.linkage(condensada, method="ward")


def dendrograma(enlace: np.ndarray, tipificadas: pd.DataFrame) -> None:
    """Dibuja el árbol, completo y recortado, y lo guarda en figuras/."""
    import matplotlib.pyplot as plt
    import scipy.cluster.hierarchy as sch

    FIGURAS_DIR.mkdir(exist_ok=True)
    fig, (izq, der) = plt.subplots(1, 2, figsize=(16, 6))

    # Izquierda: el árbol completo, como referencia de la forma general.
    sch.dendrogram(enlace, ax=izq, no_labels=True)
    izq.set_title(f"Dendrograma completo ({len(tipificadas):,} municipios)")
    izq.set_xlabel("Municipios")
    # Ojo: en Ward la altura no es la distancia entre dos municipios, sino el
    # aumento de varianza que causa la unión. Por eso llega a 400 y pico
    # cuando la mayor distancia entre dos municipios era 29,6.
    izq.set_ylabel("Distancia de enlace (Ward)")

    # Derecha: recortado a las últimas uniones, que es donde se decide en
    # cuántos grupos conviene cortar.
    sch.dendrogram(enlace, ax=der, truncate_mode="lastp",
                   p=RAMAS_RECORTADAS, show_leaf_counts=True,
                   show_contracted=True)
    der.set_title(f"Últimas {RAMAS_RECORTADAS} uniones "
                  f"(entre paréntesis, municipios de cada rama)")
    der.set_xlabel("Ramas")
    der.set_ylabel("Distancia de enlace (Ward)")
    der.tick_params(axis="x", labelsize=7)

    fig.tight_layout()
    salida = FIGURAS_DIR / "dendrograma.png"
    fig.savefig(salida, dpi=140)
    print(f"\n  escrito {salida.relative_to(BASE_DIR)}")
    plt.show()


# Hasta cuántos grupos se dibuja la curva del codo. Más allá de 20 las
# alturas son ya casi planas y no aportan nada a la decisión.
GRUPOS_MAXIMO = 20


def grafico_codo(enlace: np.ndarray) -> None:
    """Altura de unión frente a número de grupos: el gráfico del codo.

    Dos paneles porque son dos magnitudes distintas y meterlas en un mismo eje
    con dos escalas engañaría sobre su tamaño relativo:

    - izquierda, la altura a la que se deshace cada número de grupos. Donde la
      curva se dobla (el codo) es donde seguir uniendo empieza a costar caro.
    - derecha, el salto de cada unión respecto a la anterior. Es la misma
      información derivada, pero el máximo salta a la vista en vez de haber
      que estimarlo mirando la curvatura.
    """
    import matplotlib.pyplot as plt

    alturas = enlace[:, 2]
    ks = np.arange(2, GRUPOS_MAXIMO + 1)
    valores = np.array([alturas[-(k - 1)] for k in ks])
    saltos = np.array([alturas[-(k - 1)] - alturas[-k] for k in ks])
    corte = int(ks[saltos.argmax()])

    FIGURAS_DIR.mkdir(exist_ok=True)
    fig, (izq, der) = plt.subplots(1, 2, figsize=(14, 5.5))

    izq.plot(ks, valores, marker="o", color="#3987e5", linewidth=2,
             markersize=6)
    izq.set_title("Altura de unión según el número de grupos")
    izq.set_xlabel("Número de grupos")
    izq.set_ylabel("Distancia de enlace (Ward)")

    # Barras para el salto: el máximo en color distinto y etiquetado, que es
    # la única lectura que se le pide a este panel.
    colores = ["#c9302c" if k == corte else "#8fa8c8" for k in ks]
    der.bar(ks, saltos, color=colores)
    der.set_title("Cuánto sube cada unión respecto a la anterior")
    der.set_xlabel("Número de grupos")
    der.set_ylabel("Salto en la distancia de enlace")
    der.annotate(f"máximo: cortar en {corte} grupos",
                 xy=(corte, saltos.max()),
                 xytext=(corte + 1.5, saltos.max() * 0.92),
                 color="#c9302c", fontsize=10,
                 arrowprops=dict(arrowstyle="->", color="#c9302c"))

    for eje in (izq, der):
        eje.set_xticks(ks)
        eje.tick_params(axis="x", labelsize=8)
        eje.grid(axis="y", alpha=0.25, linewidth=0.6)
        eje.set_axisbelow(True)
        for lado in ("top", "right"):
            eje.spines[lado].set_visible(False)

    fig.tight_layout()
    salida = FIGURAS_DIR / "codo.png"
    fig.savefig(salida, dpi=140)
    print(f"  escrito {salida.relative_to(BASE_DIR)}")
    plt.show()


def informar_dendrograma(enlace: np.ndarray,
                         tipificadas: pd.DataFrame) -> None:
    """Las últimas uniones del árbol: es lo que dice en cuántos grupos cortar."""
    alturas = enlace[:, 2]

    print("\n" + "=" * 74)
    print("DENDROGRAMA")
    print("=" * 74)
    print(f"  método: Ward sobre {len(tipificadas):,} municipios")
    print(f"  altura de la última unión: {alturas[-1]:.1f}")

    # El salto entre uniones consecutivas es la pista habitual: donde el árbol
    # da un estirón grande, unir más grupos ya cuesta mucho, y ahí es donde
    # tiene sentido cortar.
    print(f"\n  {'grupos':>7s} {'altura de la union':>19s} "
          f"{'salto respecto a la anterior':>29s}")
    for k in range(2, 13):
        altura = alturas[-(k - 1)]
        salto = alturas[-(k - 1)] - alturas[-k]
        print(f"  {k:>7d} {altura:>19.1f} {salto:>29.1f}")

    mejor = int(np.argmax([alturas[-(k - 1)] - alturas[-k]
                           for k in range(2, 13)])) + 2
    print(f"\n  el salto mayor es el que deshace {mejor} grupos, así que "
          f"cortar en {mejor} es el primer candidato")

    # Con los tamaños delante se ve si el corte reparte o si solo está
    # apartando rarezas. Un grupo de uno o dos municipios significa que las
    # colas largas mandan y que hay que tratarlas antes de agrupar.
    from scipy.cluster.hierarchy import fcluster
    print(f"\n  cómo reparte cada corte:")
    for k in range(2, 11):
        tam = np.bincount(fcluster(enlace, k, criterion="maxclust"))[1:]
        print(f"    k={k:>2}: " + "  ".join(f"{x:,}" for x in
                                            sorted(tam, reverse=True)))
    print(f"\n  no hay grupos residuales: incluso con k=10 el menor tiene "
          f"{np.bincount(fcluster(enlace, 10, criterion='maxclust'))[1:].min():,} "
          f"municipios")


# ==========================================================================
# §7  CARACTERIZACIÓN DE LOS GRUPOS
# ==========================================================================

GRUPOS = 7
# Provincias que se listan por grupo. Con más de seis la lista deja de leerse
# y no añade nada: la cola son provincias con dos o tres municipios.
PROVINCIAS_LISTADAS = 6
# A partir de esta puntuación z se considera que una variable caracteriza al
# grupo. 0,5 desviaciones es suficiente para que la diferencia se note en los
# datos originales sin llenar el informe de ruido.
Z_CARACTERISTICA = 0.5

# Nombres propuestos para k=7, leídos del perfil de cada grupo.
#
# ATENCIÓN: la numeración la asigna fcluster y depende del árbol. Si cambian
# los datos, las variables o el k, los números se reordenan y estos nombres
# dejarían de corresponder. Hay que volver a mirar el informe antes de fiarse.
# El 1 y el 7 se nombraron primero por sus provincias, y era engañoso: el 1
# lleva las capitales del sur (Sevilla, Badajoz, Málaga), así que no es "norte
# próspero" sino el perfil urbano; y el 7 lleva Sabadell, Terrassa y 53
# municipios de Madrid, así que no es solo litoral turístico.
NOMBRES_GRUPOS = {
    1: "Urbano o cualificados de renta alta",
    2: "Rural en despoblación extrema",
    3: "Agrario próspero",
    4: "Rural agrario envejecido",
    5: "Sur agrario de rentas bajas",
    6: "Perfil medio con industria",
    7: "Urbano de rentas bajas y migración",
}

# El otro corte del árbol. La curva del codo señala 14 con un salto de 6,65,
# el mayor después del de 7; pasar de 14 a 15 cuesta solo 0,41.
GRUPOS_FINO = 14
NOMBRES_GRUPOS_FINO = {
    1: "Industrial cualificado del norte",
    2: "Metrópolis y residencial acomodado",
    3: "Aldeas agrarias en extinción",
    4: "Rural despoblado sin relevo",
    5: "Agrario próspero del Ebro",
    6: "Aldeas con fecundidad aparente alta",
    7: "Agrario envejecido de pleno empleo",
    8: "Rural envejecido e inactivo",
    9: "Agrario intensivo del sur",
    10: "Interior meridional con paro alto",
    11: "Comarcas industriales en declive",
    12: "Industrial joven y migratorio",
    13: "Litoral turístico de residentes extranjeros",
    14: "Periferias urbanas jóvenes",
}


def poblacion_municipal() -> pd.Series:
    """Población total por municipio, para poder identificar los grandes.

    No es una de las variables del agrupamiento: sirve solo para nombrar los
    municipios reconocibles de cada grupo y para ordenar los grupos por
    tamaño típico de municipio.
    """
    muni, _ref = app.cargar_poblacion(app.EXCEL_PATH.stat().st_mtime)
    total = muni[(muni["Sexo"] == "Ambos sexos")
                 & (muni["grupo_edad"] == "Todas las edades")]
    return pd.to_numeric(total.set_index("cod")["Total"], errors="coerce")


def municipios_del_grupo(etiquetas: pd.Series, tipificadas: pd.DataFrame,
                         completos: pd.DataFrame, poblacion: pd.Series,
                         grupo: int, cuantos: int = 6) -> dict:
    """Los municipios que mejor identifican un grupo.

    Dos listas, porque responden a preguntas distintas: los más típicos son
    los más cercanos al centro del grupo, y describen el perfil; los más
    poblados son los reconocibles, y sitúan el grupo en el mapa mental.
    """
    columnas = [v for v, _d in VARIABLES if v != VARIABLE_EXCLUIDA]
    dentro = etiquetas.values == grupo
    indices = completos.index[dentro]
    datos = tipificadas.loc[indices, columnas].to_numpy()
    distancias = np.linalg.norm(datos - datos.mean(axis=0), axis=1)

    pobl = poblacion.reindex(indices).fillna(0)
    return {
        "n": int(dentro.sum()),
        "poblacion": float(pobl.sum()),
        "mediana_habitantes": float(pobl.median()),
        "tipicos": [completos.loc[i, "nombre"]
                    for i in indices[np.argsort(distancias)[:cuantos]]],
        "poblados": [(completos.loc[i, "nombre"], int(v))
                     for i, v in pobl.sort_values(
                         ascending=False).head(cuantos).items()],
    }


def agrupar(enlace: np.ndarray, k: int,
            tipificadas: pd.DataFrame) -> pd.Series:
    """Corta el árbol en k grupos y devuelve la etiqueta de cada municipio."""
    from scipy.cluster.hierarchy import fcluster

    etiquetas = fcluster(enlace, k, criterion="maxclust")
    return pd.Series(etiquetas, index=tipificadas.index, name="grupo")


def guardar_etiquetas(etiquetas: pd.Series, completos: pd.DataFrame,
                      enlace: np.ndarray, descartados: int) -> None:
    """Escribe las etiquetas de grupo y la ficha del analisis para app.py.

    El dashboard no puede rehacer el agrupamiento: cortar el arbol exige la
    matriz de disimilitud y scipy, que no esta instalado en el servidor. Y
    tampoco debe reimplementar los nombres de las variables ni los de los
    grupos, porque entonces podrian divergir de los de aqui. Asi que este
    script deja las dos cosas escritas y app.py se limita a leerlas.

    Se guardan solo las etiquetas, no los perfiles: las medias por grupo son
    un groupby sobre 7.502 filas, instantaneo, y calcularlas en el dashboard
    garantiza que cuadran con la tabla de variables que se este leyendo.
    """
    pd.DataFrame({"grupo": etiquetas.astype("int8")},
                 index=etiquetas.index).to_parquet(SALIDA_GRUPOS)

    ficha = {
        "grupos": GRUPOS,
        "nombres": {str(g): n for g, n in NOMBRES_GRUPOS.items()},
        "variables": [{"clave": v, "descripcion": d} for v, d in VARIABLES],
        "variable_excluida": VARIABLE_EXCLUIDA,
        "z_caracteristica": Z_CARACTERISTICA,
        "provincias_listadas": PROVINCIAS_LISTADAS,
        "municipios": int(len(completos)),
        "descartados": int(descartados),
        # La altura de la ultima union mide lo lejos que estaban los dos
        # bloques que el arbol junta al final: es la cifra que cierra el
        # informe y da la escala del dendrograma.
        "altura_ultima_union": float(enlace[-1, 2]),
        "metodo": "Ward",
        "distancia": "euclidea",
    }
    SALIDA_META.write_text(
        json.dumps(ficha, ensure_ascii=False, indent=2), encoding="utf-8")


def nombres_provincias() -> pd.Series:
    """Código de provincia de dos dígitos -> nombre, sacado del censo.

    Están en la misma columna `Territorio` que los municipios, distinguidas
    porque su código tiene dos dígitos en vez de cinco.
    """
    largo = pd.read_parquet(CACHE_DIR / "datos_largo.parquet")
    territorios = largo["Territorio"].astype(str).drop_duplicates()
    provincias = territorios[territorios.str.match(r"^\d{2}\s")]
    return pd.Series(provincias.str[3:].values,
                     index=provincias.str[:2].values).sort_index()


def perfil_de_grupos(etiquetas: pd.Series, completos: pd.DataFrame,
                     tipificadas: pd.DataFrame) -> tuple[pd.DataFrame,
                                                         pd.DataFrame]:
    """Media de cada variable por grupo, en unidades originales y en z.

    Las dos hacen falta: las unidades originales dicen cómo es el grupo, y las
    z dicen en qué se distingue del resto de España, que no es lo mismo. Un
    grupo puede tener un 20 % de paro (alto en absoluto) y estar en la media
    si todos lo tienen.
    """
    columnas = [v for v, _d in VARIABLES]
    medias = completos[columnas].groupby(etiquetas).mean().T
    zetas = tipificadas[columnas].groupby(etiquetas).mean().T
    return medias, zetas


def informar_grupos(etiquetas: pd.Series, completos: pd.DataFrame,
                    tipificadas: pd.DataFrame,
                    etiquetas_texto: dict[int, str] | None = None) -> None:
    """Informe por grupo: tamaño, rasgos que lo distinguen y provincias."""
    medias, zetas = perfil_de_grupos(etiquetas, completos, tipificadas)
    descripciones = dict(VARIABLES)
    provincias = nombres_provincias()
    tamanos = etiquetas.value_counts().sort_index()
    # Provincia de cada municipio: los dos primeros dígitos del código INE.
    prov_de_muni = completos.index.str[:2]

    print("\n" + "=" * 74)
    print(f"LOS {len(tamanos)} GRUPOS, VARIABLE A VARIABLE")
    print("=" * 74)
    # La referencia es la media de los municipios, sin ponderar por
    # población: es la correcta para comparar grupos de municipios, pero no
    # coincide con la cifra nacional, porque aquí Villarejo pesa lo mismo que
    # Madrid.
    print("  medias en unidades originales; la última columna es la media de")
    print(f"  los {len(completos):,} municipios, sin ponderar por población")
    print()
    cab = "  ".join(f"{'g' + str(g):>8s}" for g in medias.columns)
    print(f"  {'variable':24s} {cab}  {'MEDIA':>9s}")
    for var, _d in VARIABLES:
        fila = "  ".join(f"{medias.loc[var, g]:8.1f}" for g in medias.columns)
        print(f"  {var:24s} {fila}  {completos[var].mean():9.1f}")
    print()
    print(f"  {'municipios':24s} "
          + "  ".join(f"{tamanos[g]:8,}" for g in medias.columns)
          + f"  {len(completos):9,}")

    for g in medias.columns:
        titulo = f"GRUPO {g}"
        if etiquetas_texto and g in etiquetas_texto:
            titulo += f" — {etiquetas_texto[g]}"
        print("\n" + "-" * 74)
        print(f"{titulo}   ({tamanos[g]:,} municipios, "
              f"{100*tamanos[g]/len(completos):.1f} %)")
        print("-" * 74)

        # Lo que distingue al grupo: las variables más alejadas de la media
        # nacional, en desviaciones típicas.
        z = zetas[g].sort_values()
        print("  en qué se sale de la media nacional:")
        for var in list(z.index[::-1][:4]) + list(z.index[:4]):
            if abs(z[var]) < Z_CARACTERISTICA:
                continue
            signo = "+" if z[var] > 0 else "-"
            print(f"    {signo} {descripciones[var]:<48} "
                  f"{medias.loc[var, g]:8.1f}  (media municipal "
                  f"{completos[var].mean():.1f}, {z[var]:+.2f} desv)")

        cuenta = pd.Series(prov_de_muni[etiquetas.values == g]).value_counts()
        print(f"\n  provincias ({len(cuenta)} de 52, "
              f"las {min(PROVINCIAS_LISTADAS, len(cuenta))} con más "
              f"municipios):")
        for cod, n in cuenta.head(PROVINCIAS_LISTADAS).items():
            # Cuántos de los municipios de esa provincia caen en este grupo:
            # distingue "provincia grande" de "provincia característica".
            total_prov = int((prov_de_muni == cod).sum())
            print(f"    {provincias.get(cod, cod):<24} {n:>4} municipios "
                  f"({100*n/total_prov:.0f} % de la provincia)")

        # Provincias donde el grupo es dominante: es lo que de verdad lo
        # localiza en el mapa, más que el recuento en bruto. Se exigen 20
        # municipios para no premiar a una provincia con tres.
        total_por_prov = pd.Series(prov_de_muni).value_counts()
        suficientes = cuenta[cuenta >= 20]
        cuota = (suficientes / total_por_prov.reindex(suficientes.index))
        dominante = cuota.sort_values(ascending=False).head(3)
        if len(dominante):
            print("  donde más pesa (sobre el total de cada provincia):")
            for cod, frac in dominante.items():
                print(f"    {provincias.get(cod, cod):<24} {100*frac:.0f} %")


# ==========================================================================
# §8  MAIN
# ==========================================================================


def main() -> None:
    CACHE_DIR.mkdir(exist_ok=True)
    completos, todos = construir_dataframe()
    informar(completos, todos)

    # El parquet guarda la precisión completa: es el que se usará para agrupar.
    completos.to_parquet(SALIDA_PARQUET)
    # El CSV es para leerlo, así que va redondeado. Con 16 decimales los
    # porcentajes son ilegibles, y la diferencia es ruido de coma flotante: dos
    # decimales en un porcentaje ya es una centésima de punto. Separadores
    # españoles, para abrirlo en Excel sin tocar nada.
    # Si el CSV está abierto en Excel, Windows no deja sobrescribirlo. Eso no
    # es motivo para tumbar el script y perder el clustering que viene
    # después: se avisa y se sigue con el fichero anterior.
    try:
        completos.round(2).to_csv(SALIDA_CSV, sep=";", decimal=",",
                                  encoding="utf-8-sig")
        csv_escrito = True
    except PermissionError:
        csv_escrito = False

    print("\n" + "=" * 74)
    print(f"  escrito {SALIDA_PARQUET.name} "
          f"({SALIDA_PARQUET.stat().st_size/1e6:.2f} MB)")
    if csv_escrito:
        print(f"  escrito {SALIDA_CSV.name} "
              f"({SALIDA_CSV.stat().st_size/1e6:.2f} MB)")
    else:
        print(f"  !! {SALIDA_CSV.name} NO actualizado: está abierto en otro "
              f"programa. Ciérralo y vuelve a ejecutar si lo necesitas.")
    print(f"  {len(completos):,} municipios x {len(VARIABLES)} variables")
    print("=" * 74)

    # Tabla estandarizada, la que se usará para agrupar. Se queda en memoria a
    # propósito: no se escribe ningún fichero, se recalcula en cada ejecución
    # a partir de `completos`.
    tipificadas = estandarizar(completos)
    informar_estandarizadas(tipificadas, completos)

    print(f"\n  tabla estandarizada: {tipificadas.shape[0]:,} municipios x "
          f"{len(VARIABLES)} variables (mas la columna nombre)")
    print(tipificadas.head(3).to_string(max_cols=6))

    # Matriz de disimilitud, también en memoria: 215 MB no son para un fichero.
    t0 = time.time()
    condensada, usadas = matriz_disimilitud(tipificadas)
    informar_disimilitud(condensada, usadas, tipificadas)
    print(f"\n  calculada en {time.time()-t0:.1f} s")

    t0 = time.time()
    enlace = enlazar(condensada)
    informar_dendrograma(enlace, tipificadas)
    print(f"\n  enlazado en {time.time()-t0:.1f} s")
    dendrograma(enlace, tipificadas)
    grafico_codo(enlace)

    etiquetas = agrupar(enlace, GRUPOS, tipificadas)
    informar_grupos(etiquetas, completos, tipificadas, NOMBRES_GRUPOS)

    guardar_etiquetas(etiquetas, completos, enlace, len(todos) - len(completos))
    print()
    print("=" * 74)
    print(f"  escrito {SALIDA_GRUPOS.name} y {SALIDA_META.name}: "
          f"las etiquetas que lee la pantalla Clustering del dashboard")
    print("=" * 74)


if __name__ == "__main__":
    main()
