# -*- coding: utf-8 -*-
"""
Dashboard Demográfico — Municipios y Comarcas de España
=======================================================

Pantallas previstas:
  1. Ficha Municipal          <-- IMPLEMENTADA
  2. Mercado laboral
  3. Mapas (municipios / comarcas / rejilla 1 km)
  4. Clusters

Ejecutar con:
    streamlit run app.py

Estructura del fichero (mantener este orden al ir añadiendo cosas):
    §1  Configuración y constantes
    §2  Paleta y plantilla de Plotly
    §3  Carga de datos (con caché en parquet)
    §4  Utilidades de transformación
    §5  Componentes visuales (KPIs, gráficos)
    §6  Pantalla: Ficha Municipal
    §7  Pantalla: Mapas
    §8  Main
"""

from __future__ import annotations

import math
import re
from pathlib import Path

import json
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

# El mapa necesita la pila geoespacial. Si falta, el resto del dashboard sigue
# funcionando y solo ese bloque avisa de que hay que instalarla.
try:
    import geopandas as gpd
    GEO_DISPONIBLE = True
except ImportError:                                   # pragma: no cover
    gpd = None
    GEO_DISPONIBLE = False

# ==========================================================================
# §1  CONFIGURACIÓN Y CONSTANTES
# ==========================================================================

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "DatosCensoINE.xlsx"
CACHE_DIR = BASE_DIR / "data_cache"

# Hojas del Excel. Las no usadas todavía quedan aquí documentadas para
# las pantallas siguientes.
HOJAS = {
    "poblacion": "Datos_largo",       # Territorio | Sexo | Grupo de edad | Total | España | Extranjero
    "relacion_actividad": "Relacion_actividad",
    "nivel_estudios": "Nivel_estudios",
    "ocupacion": "Ocupacion",
    "actividad": "Actividad",
    "situacion_prof": "Situación_prof",
    "estado_civil": "Estado_civil",
    "renta": "Renta",
}

# Orden canónico de los grupos de edad (el Excel no viene ordenado alfabéticamente
# de forma útil: "De 5 a 9 años" iría después de "De 45 a 49 años").
GRUPOS_EDAD = [
    "De 0 a 4 años", "De 5 a 9 años", "De 10 a 14 años", "De 15 a 19 años",
    "De 20 a 24 años", "De 25 a 29 años", "De 30 a 34 años", "De 35 a 39 años",
    "De 40 a 44 años", "De 45 a 49 años", "De 50 a 54 años", "De 55 a 59 años",
    "De 60 a 64 años", "De 65 a 69 años", "De 70 a 74 años", "De 75 a 79 años",
    "De 80 a 84 años", "De 85 a 89 años", "De 90 a 94 años", "De 95 a 99 años",
    "100 y más años",
]
TOTAL_EDADES = "Todas las edades"

# Punto medio de cada intervalo, para estimar la edad media a partir de datos
# agrupados (el censo no publica la edad exacta por municipio).
PUNTO_MEDIO_EDAD = {g: 2.5 + 5 * i for i, g in enumerate(GRUPOS_EDAD[:-1])}
PUNTO_MEDIO_EDAD["100 y más años"] = 101.0

# Los umbrales de 18 y 67 años caen dentro de un intervalo quinquenal, así que
# hay que repartirlo. Cada grupo abarca 5 edades simples y en los dos casos nos
# quedamos con 3 de ellas:
#   menores de 18  -> de "De 15 a 19 años" cuentan 15, 16 y 17
#   67 y más       -> de "De 65 a 69 años" cuentan 67, 68 y 69
# Se asume reparto uniforme dentro del intervalo.
FRACCION_INTERVALO = 3 / 5
GRUPO_CORTE_MENORES = "De 15 a 19 años"   # umbral de 18 años
GRUPO_CORTE_MAYORES = "De 65 a 69 años"   # umbral de 67 años

# Columnas de lugar de nacimiento en Datos_largo. Se cumple, en las 540.210
# filas y sin una sola desviación, que ESPAÑA + EXTRANJERO == TOTAL.
COL_TOTAL, COL_ESP, COL_EXT = "Total", "España", "Extranjero"

# --- Mercado laboral (hoja Relacion_actividad) ----------------------------
# La hoja da la partición exacta OCUPADO + PARADO + INACTIVO == TOTAL, sobre la
# población de 16 y más años. NO trae columna "Activo": los activos son la suma
# de ocupados y parados, así que pintarlos junto a sus dos componentes
# duplicaría población. Por eso el gráfico usa las tres categorías reales.
CAT_ACTIVIDAD = [
    # columna del Excel, etiqueta, color de relleno
    ("Ocupado/a",  "Ocupados",  "s3"),
    ("Parado/a",   "Parados",   "s4"),
    ("Inactivo/a", "Inactivos", "neutro"),
]

# --- Situación profesional (hoja Situación_prof) --------------------------
# También partición exacta: propia + ajena == Total, y exactamente los mismos
# 7.563 municipios que la hoja de actividad. Su Total ronda el de ocupados
# (ratio mediana 0,987): describe el reparto de la población ocupada.
# La etiqueta corta se queda en "Cuenta ajena", pero el nombre largo del INE
# viaja al tooltip, porque incluye "y otra situación".
CAT_SITUACION_PROF = [
    # columna del Excel, etiqueta corta, color de relleno
    ("Trabajador por cuenta propia", "Cuenta propia", "s7_claro"),
    ("Trabajador por cuenta ajena y otra situación", "Cuenta ajena", "s7"),
]

# --- Desgloses seleccionables del bloque inferior -------------------------
# Tres hojas con la misma forma, intercambiables desde un desplegable. Las
# etiquetas del INE son larguísimas ("Segunda etapa de Educación Secundaria y
# Educación Postsecundaria no Superior"), así que en el eje va una versión
# corta partida en dos líneas y el nombre completo viaja al tooltip.
#
# COLORES. Como los tres desgloses nunca se ven a la vez -los cambia el
# desplegable-, cada uno lleva su propia paleta sin chocar con los otros. Todas
# son de croma bajo (<= 6 sobre 100 en OKLab) para que no resulten estridentes,
# y todas superan 3:1 de contraste sobre el fondo oscuro.
#   · Actividad: color por sector. Verde salvia para el campo, gris acero para
#     la industria, marrón tierra para la construcción y azul apagado para los
#     servicios. Separación mínima ΔE 13,4, buscada por optimización.
#   · Nivel de estudios y Ocupación: son variables ORDENADAS, así que llevan una
#     rampa de un solo tono con pasos de luminosidad regulares (~0,11 en OKLab).
#     Más claro = nivel más alto, que sobre fondo oscuro es lo que destaca. Los
#     pasos contiguos quedan cerca (ΔE ~10-12) a propósito: eso es lo que hace
#     que se lea como una progresión y no como categorías sueltas.
# "No consta" va siempre al gris neutro, porque es un residuo y no una categoría.
#
# El ΔE de estas paletas queda por debajo del suelo de 15 que se exige cuando el
# color IDENTIFICA una serie. Aquí no la identifica: cada barra lleva su nombre
# rotulado en el eje, y el color solo refuerza. Subirlo más obligaría a saturar,
# que es justo lo que no se quiere.
#
# Los porcentajes se calculan sobre la SUMA DE LAS CATEGORÍAS MOSTRADAS y no
# sobre la columna Total de la hoja. En Nivel_estudios y Ocupacion da igual
# porque cuadran exacto, pero en Actividad la suma se queda por debajo del
# "Total CNAE" (hasta un 3,56 % en el peor municipio, exacta en 5.834 de
# 7.563): normalizando sobre lo mostrado, las barras suman siempre 100 % y no
# queda un resto invisible.
DESGLOSES = {
    "Nivel de estudios": {
        "hoja": "nivel_estudios",
        # El año va solo en el título de la tarjeta, nunca en la opción del
        # desplegable: ahí sería ruido al elegir.
        "titulo": "Nivel de estudios (2024)",
        "base": "Población de 16 y más años",
        "categorias": [
            # rampa azul: cuanto más alto el nivel, más claro
            ("Educación primaria e inferior",
             "Primaria<br>o inferior", "#647d97"),
            ("Primera etapa de Educación Secundaria y similar",
             "Secundaria<br>1ª etapa", "#849eba"),
            ("Segunda etapa de Educación Secundaria y Educación Postsecundaria"
             " no Superior",
             "Secundaria<br>2ª etapa", "#a7c1dc"),
            ("Educación superior",
             "Educación<br>superior", "#cee0f4"),
        ],
    },
    "Ocupación": {
        "hoja": "ocupacion",
        "titulo": "Ocupación (2023)",
        "base": "Población ocupada",
        "categorias": [
            # rampa ocre: cuanto más cualificada la ocupación, más clara
            ("Directores/gerentes y profesionales/técnicos de nivel medio o alto",
             "Directivos<br>y técnicos", "#e9cbab"),
            ("Trabajadores cualificados y oficiales/operarios de nivel bajo",
             "Cualificados<br>y operarios", "#c4a582"),
            ("Ocupaciones elementales",
             "Ocupaciones<br>elementales", "#9e805e"),
            ("No consta", "No consta", "#666666"),
        ],
    },
    "Actividad": {
        "hoja": "actividad",
        "titulo": "Actividad (2024)",
        "base": "Población ocupada, por rama de actividad (CNAE)",
        "categorias": [
            ("Agricultura, ganadería y pesca",
             "Agricultura<br>y pesca", "#72a473"),      # verde salvia
            ("Industria", "Industria", "#b6bbc1"),       # gris acero
            ("Construcción", "Construcción", "#ad744d"),  # marrón tierra
            ("Servicios", "Servicios", "#6f8eb6"),       # azul apagado
            # Columna casi vacía: 238 personas en toda España, nula en 4.144
            # municipios. Se mantiene por coherencia con la hoja, pero su barra
            # sale a cero.
            ("No consta", "No consta", "#666666"),
        ],
    },
}

# --- Renta (hoja Renta) ---------------------------------------------------
# Única hoja que cubre los 8.132 municipios. Sus columnas vienen como TEXTO,
# no como número: 73 municipios traen un '.' -el marcador de dato suprimido
# del INE- y eso convierte toda la columna en object. Hay que forzar a numérico.
COL_RENTA = "Renta bruta media por persona"

# --- Estado civil (hoja Estado_civil) -------------------------------------
# Partición exacta sobre la población de 16 y más años (su total coincide con
# el de Relacion_actividad). Cinco porciones, dentro del límite de un donut.
# Aquí el color SÍ identifica cada porción, así que los tonos se eligieron por
# optimización: todos los pares separan ΔE >= 18, por encima del suelo de 15,
# manteniendo el croma en 11 sobre 100 para que no resulten estridentes.
CAT_ESTADO_CIVIL = [
    # columna del Excel, etiqueta corta, color
    ("Soltero/a",                 "Soltero/a",   "#5e9ad6"),   # azul
    ("Casado/a",                  "Casado/a",    "#6dc799"),   # verde azulado
    ("Viudo/a",                   "Viudo/a",     "#d7c4ee"),   # violeta claro
    ("Divorciado/a o separado/a", "Div. o sep.", "#c08843"),   # ocre
    ("No consta",                 "No consta",   "#666666"),   # gris
]

# --- Índice de fecundidad -------------------------------------------------
#   numerador   = población de 0 a 4 años (ambos sexos) / 5   -> nacimientos al
#                 año, aproximados por el tamaño de la cohorte
#   denominador = mujeres de 15 a 49 años / 35                -> mujeres por
#                 cada año de edad fértil
#   índice      = numerador / denominador                     -> hijos por mujer
#
# "De 15 a 49" son exactamente siete grupos quinquenales, así que el 35 del
# denominador cuadra sin repartir ningún intervalo. Contrastado: para España da
# 1,123 frente al ~1,19 que publica el INE, así que la aproximación se sostiene.
GRUPO_NACIMIENTOS = "De 0 a 4 años"
FERTIL_DESDE = "De 15 a 19 años"
FERTIL_HASTA = "De 45 a 49 años"
GRUPOS_EDAD_FERTIL = GRUPOS_EDAD[
    GRUPOS_EDAD.index(FERTIL_DESDE):GRUPOS_EDAD.index(FERTIL_HASTA) + 1
]

# Población en edad de trabajar: diez grupos completos, de 15 a 64 años. Es la
# definición estándar (INE, Eurostat) y es el denominador de la tasa de empleo.
EDAD_TRABAJAR_DESDE = "De 15 a 19 años"
EDAD_TRABAJAR_HASTA = "De 60 a 64 años"
GRUPOS_EDAD_TRABAJAR = GRUPOS_EDAD[
    GRUPOS_EDAD.index(EDAD_TRABAJAR_DESDE):GRUPOS_EDAD.index(EDAD_TRABAJAR_HASTA) + 1
]

# --- Mapa de rejillas de 1 km ---------------------------------------------
# Los dos ficheros los genera preprocessGrids.py, que hace el trabajo pesado
# una sola vez: la rejilla europea de origen ocupa 20 GB y aquí queda en unos
# pocos MB con solo las celdas habitadas de España ya asignadas a su municipio.
GEO_GRIDS = CACHE_DIR / "grids_es_1km.parquet"
# De los municipios hay dos versiones: la completa del IGN (68 MB, demasiado
# para subirla a un repositorio) y una copia aligerada (14 MB) con el error de
# contorno por debajo de la décima de pixel a la escala a la que se dibuja.
# Se prefiere la ligera y se cae a la completa si no está, así que en local
# funcionan las dos y al servidor solo hace falta subir la ligera.
GEO_MUNICIPIOS = next(
    (ruta for ruta in (CACHE_DIR / "municipios_ligero.parquet",
                       CACHE_DIR / "municipios.parquet") if ruta.exists()),
    CACHE_DIR / "municipios_ligero.parquet",
)

# Mapa base de OpenStreetMap: colorido, con las calles, los parques en verde y
# las masas de agua en azul bien diferenciados. No necesita clave. Una
# alternativa más suave con los mismos códigos de color es "carto-voyager".
ESTILO_MAPA = "open-street-map"

# Rampa secuencial de un solo tono, de rojo claro a granate. Sobre un mapa base
# claro manda al revés que sobre uno oscuro: más población = más oscuro. La
# luminosidad baja de forma monótona (OKLab 0,91 -> 0,35), que es lo que hace
# que se lea como una progresión, y los pasos contiguos separan ΔE 13-17.
ESCALA_POBLACION = [
    [0.00, "#f8c9c1"],   # rojo muy claro
    [0.25, "#f09a8e"],
    [0.50, "#dd5347"],   # rojo
    [0.75, "#ac2620"],
    [1.00, "#6b1210"],   # granate
]

# --- Mapa nacional por municipios -----------------------------------------
# Geometría simplificada que genera preprocessGrids.py. Dos ficheros porque
# Plotly incrusta el GeoJSON en cada traza, y el recuadro de Canarias con el
# fichero completo duplicaría los 3,8 MB.
GEOJSON_PENINBAL = CACHE_DIR / "municipios_peninbal.geojson"
GEOJSON_CANARIAS = CACHE_DIR / "municipios_canarias.geojson"
# Límites provinciales, para superponerlos con trazo más grueso sobre la
# retícula municipal. Van como listas de coordenadas y no como GeoJSON porque
# es lo que consume una traza de líneas y pesa una fracción.
LINEAS_PROVINCIALES = CACHE_DIR / "provincias_lineas.json"
# Trazo del límite provincial. Oscuro y no blanco: la mayoría de los colores de
# las capas son de tono medio o claro, así que una línea oscura se lee sobre
# ellos, y en la costa se funde con el fondo, que es justo lo deseable porque
# ahí el contorno ya lo marca el borde del relleno.
COLOR_LINEA_PROVINCIA = "rgba(18,18,17,0.78)"
GROSOR_LINEA_PROVINCIA = 1.1

# --- Mapas de categoría predominante --------------------------------------
# Cada uno pinta, para cada municipio, cuál de sus categorías manda. Hay dos
# formas de decidirlo y cambian mucho el mapa:
#
#   "absoluto" : gana la categoría con mayor porcentaje. Directo y literal.
#   "relativo" : gana la que más se desvía de lo normal, dividiendo el
#                porcentaje del municipio por la media de los porcentajes
#                municipales de España. Es un índice de especialización, y
#                reparte el mapa mucho más porque descuenta lo que es común a
#                todos los municipios.
#
# Los grupos van en el ORDEN DE LA LEYENDA, de arriba abajo, con su etiqueta,
# las columnas del Excel que suma y su color.
MAPAS_CATEGORIA = {
    "actividad": {
        "hoja": "actividad",
        "metodo": "relativo",
        "titulo_leyenda": "Sector predominante",
        # Verde para el campo, gris para la industria y ámbar para los
        # servicios. Los dos primeros son los mismos que usa el bloque de
        # Actividad en la ficha municipal, así que el código de color se
        # mantiene entre pantallas.
        "grupos": [
            ("Agrícola",   ["Agricultura, ganadería y pesca"], "#72a473"),
            ("Industrial", ["Industria", "Construcción"],      "#b6bbc1"),
            ("Servicios",  ["Servicios"],                      "#e0a83c"),
        ],
    },
    "nivel_estudios": {
        "hoja": "nivel_estudios",
        "metodo": "absoluto",
        "titulo_leyenda": "Nivel educativo más común",
        # Cuatro tonos claramente distintos, no una rampa: el nivel ganador es
        # una categoría, no una magnitud. Son los mismos cuatro del donut de
        # estado civil, que es el juego mejor validado del dashboard: todos los
        # pares separan ΔE >= 18, y >= 12 simulando deuteranopia, frente a
        # ΔE 1,6 a 7,2 de las alternativas que probé.
        # El violeta claro va a Superior porque es el tono más luminoso y esos
        # municipios son la minoría interesante; el ocre, al nivel que domina
        # dos tercios del mapa.
        "grupos": [
            ("Superior",
             ["Educación superior"], "#d7c4ee"),            # violeta claro
            ("Secundaria 2ª etapa",
             ["Segunda etapa de Educación Secundaria y Educación"
              " Postsecundaria no Superior"], "#6dc799"),   # verde azulado
            ("Secundaria 1ª etapa",
             ["Primera etapa de Educación Secundaria y similar"],
             "#c08843"),                                    # ocre
            ("Primaria o inferior",
             ["Educación primaria e inferior"], "#5e9ad6"),  # azul
        ],
    },
}

# Quintiles de renta, de Q1 a Q5: rojo, naranja oscuro, amarillo, verde
# amarillento y verde fuerte. Los cinco pasos superan 3:1 de contraste sobre la
# superficie oscura.
#
# Aviso conocido: una escala rojo-verde es la más difícil para la deficiencia de
# visión del color más común. Los tonos están elegidos para separar al máximo
# los dos extremos -Q1 y Q5 quedan a ΔE 12,9 simulando deuteranopia, frente a
# 6,1 con una versión más literal-, pero con esa deficiencia siguen pareciéndose
# más de lo deseable. Lo que evita el equívoco es que la leyenda rotula Q1..Q5 y
# el tooltip da la renta exacta, así que el color nunca es el único dato.
COLORES_QUINTIL = ["#c9302c", "#e5842c", "#f0d84e", "#9dbb45", "#4fb168"]

# Municipios sin dato publicado. Gris muy apagado: se distinguen del mapa pero
# no compiten con los que sí tienen dato.
COLOR_SIN_DATO = "#3a3a36"

# Reparto del lienzo. La península ocupa TODO el lienzo, no una parte: el área
# por la que se puede arrastrar el mapa es su dominio, así que recortarlo hacía
# que al desplazarlo se perdiera de vista aunque siguiera habiendo fondo gris.
# La leyenda se superpone encima, con su propio panel.
DOMINIO_PENINSULA = dict(x=[0.0, 1.0], y=[0.0, 1.0])
# Recuadro de Canarias, abajo a la derecha, sobre el Mediterráneo, que es donde
# queda hueco al encajar la península. Apaisado a propósito, para que el
# archipiélago -4,9 grados de longitud por 2,1 de latitud- lo llene.
#
# El borde izquierdo, 0,615, no es arbitrario: es el mayor recuadro que cabe
# sin tapar ningún municipio peninsular, comprobado contra los 151.187 vértices
# de la geometría. Con 0,560 se comía 62 municipios de Alicante y Murcia.
# Su proporción, 2,08, es la de ENCUADRE_CANARIAS: si no coinciden, las islas
# dejan bandas dentro del marco en vez de llenarlo.
DOMINIO_CANARIAS = dict(x=[0.615, 0.995], y=[0.012, 0.367])
# Encuadre del encarte, fijado a mano. Con fitbounds Plotly añade tanto margen
# alrededor de las islas que el mapa dibujado se queda en menos de la mitad del
# recuadro; con el rango explícito lo llena. El aspecto de esta ventana (unas
# 2,0 veces más ancha que alta, en Mercator) es el que marca la proporción del
# dominio de arriba.
ENCUADRE_CANARIAS = dict(lon=[-18.35, -13.25], lat=[27.45, 29.60])

# Encuadre de la península, también fijado a mano y por el mismo motivo:
# `fitbounds` no solo encuadra, además RECORTA. Define una ventana geográfica
# ceñida a España y todo lo que sale de ella desaparece, aunque el dominio del
# subgráfico llegue al borde del lienzo. Por eso el mapa se cortaba antes de
# llegar al recuadro al arrastrarlo.
#
# Encuadre de la península. La clave está en que el ZOOM y el ÁREA DIBUJABLE
# son cosas distintas, y aquí se usan por separado:
#
#   · Los rangos de longitud y latitud NO encuadran el mapa: definen el área
#     por la que se puede arrastrar. Lo que sale de ella se recorta.
#   · `scale` hace el zoom DENTRO de esa área, y es lo que decide el tamaño al
#     que se ve España.
#
# Con los rangos ceñidos a España -que es lo natural- las dos cosas quedaban
# atadas: o el mapa llenaba el lienzo y no se podía arrastrar en vertical, o
# había recorrido pero se veía un hueco vacío enorme. Separándolas se tienen
# las dos: la ventana es 3 veces España, así que hay un lienzo entero de
# recorrido en cada dirección, y el zoom la compensa para que España siga
# llenando el alto en la vista inicial.
#
# El aspecto de la ventana, 1,944, es el del lienzo de 1.400 x 720: cuando
# coinciden, el área dibujable es el lienzo completo y no queda ninguna banda
# recortada. Si se cambia el alto del mapa hay que recalcular la longitud.
ENCUADRE_PENINSULA = dict(lon=[-34.80, 29.82], lat=[25.74, 51.26])
CENTRO_PENINSULA = dict(lon=-2.487, lat=39.660)
# Calibrado midiendo el render: con este zoom España ocupa 886 x 720 px, es
# decir toca el borde de arriba y el de abajo sin dejar hueco.
ZOOM_PENINSULA = 3.038

# Pantallas del dashboard.
PANTALLAS = ["Ficha municipal", "Mapas"]

# Municipio preseleccionado al arrancar la app.
MUNICIPIO_INICIAL = "Madrid"


# ==========================================================================
# §2  PALETA Y PLANTILLA DE PLOTLY
#
# Paleta validada para superficie oscura. Los slots categóricos se asignan
# SIEMPRE en este orden, nunca cíclicamente, para garantizar separación
# suficiente también con daltonismo.
# ==========================================================================

C = {
    # Superficies
    "page": "#0d0d0d",
    "surface": "#1a1a19",
    # Tinta
    "ink": "#ffffff",
    "ink_2": "#c3c2b7",
    "ink_muted": "#898781",
    # Cuadrícula y ejes (líneas de un solo pelo, siempre sólidas)
    "grid": "#2c2c2a",
    "axis": "#383835",
    "border": "rgba(255,255,255,0.10)",
    # Slots categóricos (orden fijo)
    "s1": "#3987e5",   # azul     -> Hombres
    "s2": "#d95926",   # naranja  -> Mujeres
    "s3": "#199e70",   # aqua
    "s4": "#c98500",   # amarillo
    "s5": "#d55181",   # magenta
    "s6": "#008300",   # verde
    "s7": "#9085e9",   # violeta
    "s8": "#e66767",   # rojo
    # Variantes claras para el desglose por lugar de nacimiento. Codificación
    # compuesta: el TONO indica el sexo (azul/naranja) y la CLARIDAD el lugar de
    # nacimiento. Ambos ejes son robustos al daltonismo -azul frente a naranja
    # no se confunde en las deficiencias rojo-verde, y la claridad se percibe
    # igual en todas ellas-, y además el sexo va reforzado por el lado del eje
    # y el origen por el orden de apilado, así que ningún dato depende del color.
    # Los dos tonos claros tienen la misma luminosidad en OKLab (L=0,765 y
    # L=0,764) para que ninguno pese más que el otro.
    "s1_claro": "#86b6ef",   # azul claro    -> Hombres nacidos en el extranjero
    "s2_claro": "#e8a06f",   # naranja claro -> Mujeres nacidas en el extranjero
    # Gris recesivo para la categoría residual "Inactivos": 3,21:1 sobre la
    # superficie, suficiente para verse pero sin competir con ocupados y
    # parados, que son lo que se quiere leer.
    "neutro": "#6b6a65",
    # Carril de fondo de la barra de renta. Sobre fondo oscuro, un "gris muy
    # claro" sería lo más llamativo de la tarjeta y taparía el dato; el
    # equivalente recesivo aquí es un gris apenas por encima de la superficie.
    "carril": "#3a3a36",
    # Color del dato de renta del municipio.
    "renta": "#3f9e8c",
    # Violeta claro para la segunda categoría de Situación profesional. Se usa
    # la familia violeta -mismo tono, dos claridades- porque es la que más se
    # separa del bloque de actividad, que es su vecino directo (ΔE 24,6), y
    # porque una diferencia de claridad se percibe igual con cualquier tipo de
    # daltonismo. Entre sí separan ΔE 17,0, por encima del suelo de 15.
    "s7_claro": "#c4bef3",
}

FONT_STACK = 'system-ui, -apple-system, "Segoe UI", sans-serif'


def registrar_plantilla() -> None:
    """Plantilla Plotly común a todos los gráficos del dashboard.

    Se registra una sola vez; las pantallas nuevas la heredan sin
    reconfigurar nada.
    """
    pio.templates["demografia"] = go.layout.Template(
        layout=dict(
            font=dict(family=FONT_STACK, size=13, color=C["ink_2"]),
            # El gráfico pinta su propia superficie: así no depende de cómo
            # Streamlit rellene el contenedor en cada versión.
            paper_bgcolor=C["surface"],
            plot_bgcolor=C["surface"],
            colorway=[C[f"s{i}"] for i in range(1, 9)],
            title=dict(
                font=dict(size=16, color=C["ink"]),
                x=0, xanchor="left", y=0.97, pad=dict(b=12),
            ),
            # Márgenes con sitio para las etiquetas de los ejes: si se aprietan,
            # los rótulos del eje Y y del eje X se cortan.
            margin=dict(l=72, r=24, t=64, b=56),
            hoverlabel=dict(
                bgcolor=C["surface"],
                bordercolor=C["border"],
                font=dict(family=FONT_STACK, size=13, color=C["ink"]),
            ),
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02,
                xanchor="right", x=1,
                bgcolor="rgba(0,0,0,0)", borderwidth=0,
                font=dict(size=12, color=C["ink_2"]),
            ),
            xaxis=dict(
                gridcolor=C["grid"], gridwidth=1, griddash="solid",
                zeroline=False, showline=False,
                tickfont=dict(size=11, color=C["ink_muted"]),
                title=dict(font=dict(size=12, color=C["ink_muted"])),
            ),
            yaxis=dict(
                gridcolor=C["grid"], gridwidth=1, griddash="solid",
                zeroline=False, showline=False,
                tickfont=dict(size=11, color=C["ink_muted"]),
                title=dict(font=dict(size=12, color=C["ink_muted"])),
            ),
        )
    )
    pio.templates.default = "demografia"


CSS = f"""
<style>
  /* Plano de página y ancho de contenido */
  .stApp {{ background: {C["page"]}; }}
  .block-container {{ padding-top: 2.2rem; max-width: 1500px; }}

  /* Cabecera de pantalla */
  .dash-eyebrow {{
      font: 600 11px/1 {FONT_STACK};
      letter-spacing: .14em; text-transform: uppercase;
      color: {C["ink_muted"]}; margin-bottom: .5rem;
  }}
  .dash-title {{
      font: 700 30px/1.15 {FONT_STACK};
      color: {C["ink"]}; margin: 0 0 .35rem 0;
  }}
  .dash-sub {{ font: 400 14px/1.5 {FONT_STACK}; color: {C["ink_muted"]}; margin: 0; }}
  .dash-rule {{
      height: 1px; background: {C["border"]};
      border: 0; margin: 1.2rem 0 1.4rem 0;
  }}

  /* Cabecera de una tarjeta de gráfico. Los títulos van aquí y no dentro de
     la figura: Plotly ancla el título a una fracción del lienzo, así que en
     gráficos bajos un título de dos líneas se sale por arriba. */
  .card-title {{
      font: 600 16px/1.25 {FONT_STACK};
      color: {C["ink"]}; margin: 2px 0 2px 0;
  }}
  .card-sub {{
      font: 400 12px/1.35 {FONT_STACK};
      color: {C["ink_muted"]}; margin: 0 0 2px 0;
  }}

  /* Tarjetas de KPI. El recuadro lo pone st.metric(border=True); aquí solo
     se ajusta la tipografía, que es lo que Streamlit no deja configurar. */
  div[data-testid="stMetricLabel"] p {{
      font-size: 12px !important; color: {C["ink_muted"]} !important;
      letter-spacing: .02em;
  }}
  div[data-testid="stMetricValue"] {{
      font-size: 26px !important; color: {C["ink"]} !important;
      font-variant-numeric: proportional-nums;  /* nada de tabular en cifras grandes */
  }}
</style>
"""


# ==========================================================================
# §3  CARGA DE DATOS
#
# Leer 540.000 filas de un .xlsx con openpyxl tarda ~1 minuto. Se convierte
# a parquet la primera vez y a partir de ahí la carga es instantánea. La
# caché se invalida sola si el Excel cambia (se compara la fecha de
# modificación).
# ==========================================================================


@st.cache_data(show_spinner=False)
def cargar_hoja(nombre_hoja: str, mtime_excel: float) -> pd.DataFrame:
    """Lee una hoja del Excel usando parquet como caché en disco.

    `mtime_excel` no se usa dentro: está en la firma para que Streamlit
    invalide la caché en memoria cuando el Excel se modifique.
    """
    CACHE_DIR.mkdir(exist_ok=True)
    slug = re.sub(r"[^0-9A-Za-z]+", "_", nombre_hoja).strip("_").lower()
    parquet = CACHE_DIR / f"{slug}.parquet"

    if parquet.exists() and parquet.stat().st_mtime >= mtime_excel:
        return pd.read_parquet(parquet)

    with st.spinner(f"Leyendo «{nombre_hoja}» del Excel (solo la primera vez)…"):
        df = pd.read_excel(EXCEL_PATH, sheet_name=nombre_hoja, engine="openpyxl")
    try:
        df.to_parquet(parquet, index=False)
    except Exception:
        pass  # sin pyarrow seguimos funcionando, solo perdemos la caché en disco
    return df


@st.cache_data(show_spinner=False)
def cargar_poblacion(mtime_excel: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (municipios, referencias).

    - municipios:  solo filas con código de 5 dígitos, con el código separado
                   del nombre y la etiqueta lista para el selector.
    - referencias: provincias (2 dígitos) y Total Nacional, para comparar.
    """
    df = cargar_hoja(HOJAS["poblacion"], mtime_excel)
    df = df.rename(columns={"Grupo de edad": "grupo_edad"})
    df["Territorio"] = df["Territorio"].astype(str).str.strip()

    # El campo Territorio mezcla tres niveles:
    #   "Total Nacional"          -> nacional
    #   "01 Araba/Álava"          -> provincia (2 dígitos)
    #   "01001 Alegría-Dulantzi"  -> municipio (5 dígitos)
    codigo = df["Territorio"].str.extract(r"^(\d{2,5})\s+(.*)$")
    df["cod"] = codigo[0]
    df["nombre"] = codigo[1].str.strip()

    es_muni = df["cod"].str.len() == 5
    muni = df[es_muni].copy()
    ref = df[~es_muni].copy()
    ref.loc[ref["cod"].isna(), "nombre"] = "España"

    # cod_prov permitirá unir con los shapefiles del IGN y comparar cada
    # municipio con su provincia.
    muni["cod_prov"] = muni["cod"].str[:2]

    # 17 nombres de municipio están duplicados a nivel nacional (Cieza, Mieres,
    # Sada, Torrent...). Al quitar el código quedarían ambiguos, así que a esos
    # -y solo a esos- se les añade la provincia.
    prov_por_cod = (
        ref[ref["cod"].notna()].drop_duplicates("cod").set_index("cod")["nombre"].to_dict()
    )
    nombres_unicos = muni.drop_duplicates("cod")["nombre"]
    duplicados = set(nombres_unicos[nombres_unicos.duplicated(keep=False)])

    muni["etiqueta"] = muni["nombre"].where(
        ~muni["nombre"].isin(duplicados),
        muni["nombre"] + " (" + muni["cod_prov"].map(prov_por_cod).fillna("?") + ")",
    )

    # Categórica ordenada: así cualquier gráfico sale ya en orden de edad.
    orden = GRUPOS_EDAD + [TOTAL_EDADES]
    for d in (muni, ref):
        d["grupo_edad"] = pd.Categorical(d["grupo_edad"], categories=orden, ordered=True)

    cols = ["cod", "cod_prov", "nombre", "etiqueta", "Sexo", "grupo_edad",
            "Total", "España", "Extranjero"]
    return muni[cols].reset_index(drop=True), ref.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def cargar_tabla_municipal(clave_hoja: str, mtime_excel: float) -> pd.DataFrame:
    """Carga una hoja «municipio x categorías», indexada por código INE.

    Vale para Relacion_actividad, Situación_prof, Nivel_estudios, Ocupacion,
    Actividad y Estado_civil: todas comparten la misma forma -primera columna
    sin nombre con "CCCCC Nombre", luego Total y las categorías- y todas suman
    exactamente Total.

    OJO: estas hojas solo cubren 7.563 de los 8.132 municipios. El INE omite
    los 569 más pequeños (mediana de 32 habitantes) por confidencialidad, así
    que siempre hay que contemplar que un municipio no esté.

    Si más adelante hacen falta agregados por provincia o nacionales, salen de
    aquí con un groupby sobre los dos primeros dígitos del código; ojo entonces
    con que a esas sumas les faltan los municipios omitidos.
    """
    df = cargar_hoja(HOJAS[clave_hoja], mtime_excel)

    # La primera columna de estas hojas viene sin nombre (es un espacio).
    df = df.rename(columns={df.columns[0]: "Territorio"})
    df["Territorio"] = df["Territorio"].astype(str).str.strip()
    df["cod"] = df["Territorio"].str.extract(r"^(\d{5})")[0]
    df = df[df["cod"].notna()].copy()

    return df.set_index("cod").drop(columns=["Territorio"])


@st.cache_data(show_spinner=False)
def cargar_geometria_municipio(cod: str):
    """Límite del municipio en WGS84, o None si no está.

    Se lee con filtro por código para no traer a memoria los 8.132 polígonos.
    """
    if not (GEO_DISPONIBLE and GEO_MUNICIPIOS.exists()):
        return None
    gdf = gpd.read_parquet(GEO_MUNICIPIOS, filters=[("cod", "==", cod)])
    return gdf if len(gdf) else None


@st.cache_data(show_spinner=False)
def cargar_celdas_municipio(cod: str):
    """Celdas de 1 km habitadas que solapan el municipio, o None.

    Basta que la celda entre un poco en el término. Las que quedan a caballo
    entre varios municipios aparecen en todos ellos, así que la suma de sus
    poblaciones no debe compararse entre municipios.
    """
    if not (GEO_DISPONIBLE and GEO_GRIDS.exists()):
        return None
    gdf = gpd.read_parquet(GEO_GRIDS, filters=[("cod", "==", cod)])
    return gdf if len(gdf) else None


@st.cache_resource(show_spinner=False)
def cargar_geojson_municipal(ambito: str) -> dict | None:
    """GeoJSON simplificado, cacheado como recurso.

    Va en cache_resource y no en cache_data porque es un diccionario grande
    que se reutiliza tal cual: no hace falta copiarlo en cada ejecución.
    """
    ruta = GEOJSON_PENINBAL if ambito == "peninbal" else GEOJSON_CANARIAS
    if not ruta.exists():
        return None
    return json.loads(ruta.read_text(encoding="utf-8"))


@st.cache_resource(show_spinner=False)
def cargar_lineas_provinciales() -> dict | None:
    """Coordenadas de los límites provinciales, por ámbito."""
    if not LINEAS_PROVINCIALES.exists():
        return None
    return json.loads(LINEAS_PROVINCIALES.read_text(encoding="utf-8"))


@st.cache_data(show_spinner=False)
def cargar_renta(mtime_excel: float) -> tuple[pd.Series, dict]:
    """Renta bruta media por persona y su contexto nacional.

    Devuelve (serie por código INE, contexto). El contexto trae el máximo
    nacional -que es el carril de fondo de la barra-, la mediana y los cuatro
    cortes de quintil, calculados una sola vez sobre los 8.059 municipios con
    dato publicado.
    """
    df = cargar_tabla_municipal("renta", mtime_excel)
    # La columna llega como texto porque 73 municipios traen un '.'.
    serie = pd.to_numeric(df[COL_RENTA], errors="coerce")
    validos = serie.dropna()

    contexto = {
        "maximo": float(validos.max()),
        "minimo": float(validos.min()),
        "mediana": float(validos.median()),
        # Cortes del 20/40/60/80: el quintil 1 es el de renta más baja.
        "cortes": [float(validos.quantile(q)) for q in (0.2, 0.4, 0.6, 0.8)],
        "n_con_dato": int(len(validos)),
        "n_sin_dato": int(serie.isna().sum()),
    }
    return serie, contexto


# ==========================================================================
# §4  UTILIDADES DE TRANSFORMACIÓN
# ==========================================================================


def fmt_int(x: float) -> str:
    """12345 -> '12.345' (formato español)."""
    return f"{x:,.0f}".replace(",", ".")


def fmt_dec(x: float, dec: int = 1) -> str:
    """12.34 -> '12,3' (formato español)."""
    return f"{x:,.{dec}f}".replace(",", "@").replace(".", ",").replace("@", ".")


def perfil_edad_sexo(df: pd.DataFrame, col_origen: str = COL_TOTAL) -> pd.DataFrame:
    """Tabla ancha grupo_edad x {Hombres, Mujeres} para un territorio.

    Se excluye la fila agregada 'Todas las edades'.
    """
    d = df[
        df["Sexo"].isin(["Hombres", "Mujeres"])
        & (df["grupo_edad"] != TOTAL_EDADES)
    ]
    ancho = (
        d.pivot_table(index="grupo_edad", columns="Sexo",
                      values=col_origen, aggfunc="sum", observed=True)
        .reindex(GRUPOS_EDAD)
        .fillna(0)
    )
    for s in ("Hombres", "Mujeres"):
        if s not in ancho.columns:
            ancho[s] = 0
    return ancho[["Hombres", "Mujeres"]]


# Columnas del perfil desglosado y cómo se pintan. El orden importa: dentro de
# cada sexo, el segmento de nacidos en España va primero y queda pegado al eje
# central; el de nacidos en el extranjero se apila hacia fuera.
SEGMENTOS = [
    # clave     sexo        lado  color            etiqueta de leyenda
    ("H_esp", "Hombres", -1, "s1",       "Hombres · España"),
    ("H_ext", "Hombres", -1, "s1_claro", "Hombres · extranjero"),
    ("M_esp", "Mujeres", +1, "s2",       "Mujeres · España"),
    ("M_ext", "Mujeres", +1, "s2_claro", "Mujeres · extranjero"),
]


def perfil_por_origen(df: pd.DataFrame) -> pd.DataFrame:
    """Tabla grupo_edad x [H_esp, H_ext, M_esp, M_ext] para un territorio.

    Las cuatro columnas suman la población total: España y Extranjero son
    una partición exacta del Total en los datos del censo.
    """
    d = df[
        df["Sexo"].isin(["Hombres", "Mujeres"])
        & (df["grupo_edad"] != TOTAL_EDADES)
    ]
    salida = pd.DataFrame(index=pd.Index(GRUPOS_EDAD, name="grupo_edad"))
    for clave, sexo, _lado, _color, _etq in SEGMENTOS:
        col = COL_ESP if clave.endswith("_esp") else COL_EXT
        serie = (
            d[d["Sexo"] == sexo]
            .groupby("grupo_edad", observed=True)[col].sum()
            .reindex(GRUPOS_EDAD)
            .fillna(0)
        )
        salida[clave] = serie.to_numpy()
    return salida


def _reparte_umbral(por_edad: pd.Series, grupo_corte: str,
                    hacia_arriba: bool) -> float:
    """Suma los efectivos a un lado de un umbral que cae dentro de un intervalo.

    `hacia_arriba=True` -> parte del intervalo de corte + todos los grupos de
    más edad (caso "67 y más"). `False` -> todos los grupos más jóvenes + parte
    del intervalo de corte (caso "menores de 18").
    """
    i = GRUPOS_EDAD.index(grupo_corte)
    parcial = FRACCION_INTERVALO * por_edad.loc[grupo_corte]
    if hacia_arriba:
        return float(parcial + por_edad.loc[GRUPOS_EDAD[i + 1:]].sum())
    return float(por_edad.loc[GRUPOS_EDAD[:i]].sum() + parcial)


def calcular_kpis(df_muni: pd.DataFrame) -> dict:
    """Indicadores de cabecera para la ficha del municipio."""
    fila_total = df_muni[
        (df_muni["Sexo"] == "Ambos sexos") & (df_muni["grupo_edad"] == TOTAL_EDADES)
    ]
    poblacion = float(fila_total[COL_TOTAL].sum())
    extranjero = float(fila_total[COL_EXT].sum())

    perfil = perfil_edad_sexo(df_muni)
    por_edad = perfil.sum(axis=1)

    # Índice de fecundidad. El denominador se va a cero en 58 municipios que no
    # tienen ninguna mujer de 15 a 49 años; ahí sale None y la interfaz pinta
    # un guion.
    nacimientos_ano = float(por_edad.get(GRUPO_NACIMIENTOS, 0.0)) / 5
    mujeres_fertiles = float(perfil.loc[GRUPOS_EDAD_FERTIL, "Mujeres"].sum())
    fecundidad = (nacimientos_ano / (mujeres_fertiles / 35)
                  if mujeres_fertiles else None)

    # Edad media estimada a partir de los puntos medios de cada intervalo.
    pesos = por_edad.sum()
    edad_media = (
        sum(por_edad[g] * PUNTO_MEDIO_EDAD[g] for g in GRUPOS_EDAD) / pesos
        if pesos else 0.0
    )

    menores18 = _reparte_umbral(por_edad, GRUPO_CORTE_MENORES, hacia_arriba=False)
    mayores67 = _reparte_umbral(por_edad, GRUPO_CORTE_MAYORES, hacia_arriba=True)

    def pct(x: float) -> float:
        return 100 * x / poblacion if poblacion else 0.0

    return {
        "poblacion": poblacion,
        "edad_media": edad_media,
        "n_mayores67": mayores67, "pct_mayores67": pct(mayores67),
        "n_menores18": menores18, "pct_menores18": pct(menores18),
        "n_extranjero": extranjero, "pct_extranjero": pct(extranjero),
        "fecundidad": fecundidad,
        "nacimientos_ano": nacimientos_ano,
        "mujeres_fertiles": mujeres_fertiles,
    }


def poblacion_en_edad_trabajar(df_muni: pd.DataFrame) -> float:
    """Población de 15 a 64 años: denominador de la tasa de empleo."""
    por_edad = perfil_edad_sexo(df_muni).sum(axis=1)
    return float(por_edad.loc[GRUPOS_EDAD_TRABAJAR].sum())


def calcular_kpis_laborales(fila_act: pd.Series | None,
                            pob_15_64: float) -> dict:
    """Tasa de paro y tasa de empleo.

        tasa de paro   = parados / (parados + ocupados)
        tasa de empleo = ocupados / población de 15 a 64 años

    Con `fila_act` a None (municipio sin datos en la hoja) todo sale a None,
    que es lo que la interfaz pinta como «—».
    """
    if fila_act is None:
        return {k: None for k in
                ("ocupados", "parados", "inactivos", "activos",
                 "tasa_paro", "tasa_empleo", "pob_15_64")}

    ocupados = float(fila_act["Ocupado/a"])
    parados = float(fila_act["Parado/a"])
    inactivos = float(fila_act["Inactivo/a"])
    activos = ocupados + parados

    return {
        "ocupados": ocupados,
        "parados": parados,
        "inactivos": inactivos,
        "activos": activos,
        "tasa_paro": 100 * parados / activos if activos else None,
        "tasa_empleo": 100 * ocupados / pob_15_64 if pob_15_64 else None,
        "pob_15_64": pob_15_64,
    }


@st.cache_data(show_spinner=False)
def categoria_predominante(clave_mapa: str,
                           mtime_excel: float) -> tuple[pd.DataFrame, pd.Series]:
    """Categoría que manda en cada municipio, por el método que toque.

    Devuelve (tabla, medias). La tabla lleva, por municipio, el porcentaje de
    cada categoría, su ratio frente a la media y la categoría ganadora.

    Con método "absoluto" gana el mayor porcentaje. Como todas las categorías
    comparten denominador, es lo mismo que quedarse con el mayor recuento.

    Con método "relativo" gana la que más se desvía de lo normal, y la
    referencia es la MEDIA DE LOS PORCENTAJES MUNICIPALES -el municipio
    típico-, no la cuota agregada de España. En la actividad económica la
    diferencia es enorme: la media municipal es 14,6 / 22,9 / 62,5 % y la cuota
    agregada 4,2 / 18,2 / 77,6 %, de forma que con la primera el mapa reparte
    36/31/33 % y con la segunda el 69 % de los municipios saldría agrícola.
    """
    cfg = MAPAS_CATEGORIA[clave_mapa]
    tabla_hoja = cargar_tabla_municipal(cfg["hoja"], mtime_excel)

    columnas = {etq: cols for etq, cols, _color in cfg["grupos"]}
    todas = [c for cols in columnas.values() for c in cols]
    datos = tabla_hoja[todas].astype(float).fillna(0)
    base = datos.sum(axis=1)

    pct = pd.DataFrame({
        etq: 100 * datos[cols].sum(axis=1) / base
        for etq, cols in columnas.items()
    })
    pct = pct[base > 0]

    medias = pct.mean()
    ratios = pct.div(medias, axis=1)

    tabla = pct.add_prefix("pct_").join(ratios.add_prefix("ratio_"))
    # En los empates idxmax se queda con la primera columna, que es la de más
    # arriba en la leyenda. Son pocos casos y cualquier regla sería arbitraria.
    referencia = ratios if cfg["metodo"] == "relativo" else pct
    tabla["ganador"] = referencia.idxmax(axis=1)
    return tabla, medias


@st.cache_data(show_spinner=False)
def tasas_laborales_municipales(mtime_excel: float) -> pd.DataFrame:
    """Tasa de paro y tasa de empleo de todos los municipios.

    Con las mismas definiciones que la ficha municipal:
        paro   = parados / (parados + ocupados)
        empleo = ocupados / población de 15 a 64 años
    """
    act = cargar_tabla_municipal("relacion_actividad", mtime_excel)
    poblacion, _ref = cargar_poblacion(mtime_excel)

    ocupados = act["Ocupado/a"].astype(float)
    parados = act["Parado/a"].astype(float)
    activos = ocupados + parados

    en_edad = poblacion[
        (poblacion["Sexo"] == "Ambos sexos")
        & (poblacion["grupo_edad"].isin(GRUPOS_EDAD_TRABAJAR))
    ]
    pob_15_64 = (en_edad.groupby("cod", observed=True)["Total"].sum()
                 .reindex(act.index))

    return pd.DataFrame({
        "paro": 100 * parados / activos.where(activos > 0),
        "empleo": 100 * ocupados / pob_15_64.where(pob_15_64 > 0),
        "ocupados": ocupados,
        "parados": parados,
        "pob_15_64": pob_15_64,
    })


def etiquetas_de_rangos(bordes, decimales: int = 1) -> list[str]:
    """Etiqueta cada clase con su intervalo: "< 6,0 %", "6,0 – 8,5 %"...

    Los extremos se dejan abiertos porque el primero y el último tramo llegan
    hasta el mínimo y el máximo observados, que no aportan nada como cifra.
    """
    n = len(bordes) - 1
    etiquetas = []
    for i in range(n):
        if i == 0:
            etiquetas.append(f"< {fmt_dec(bordes[1], decimales)} %")
        elif i == n - 1:
            etiquetas.append(f"≥ {fmt_dec(bordes[n - 1], decimales)} %")
        else:
            etiquetas.append(f"{fmt_dec(bordes[i], decimales)} – "
                             f"{fmt_dec(bordes[i + 1], decimales)} %")
    return etiquetas


def quintil_renta(valor: float, cortes: list[float]) -> int:
    """Quintil en el que cae `valor`. 1 = renta más baja, 5 = más alta.

    Un valor que caiga justo en un corte se asigna al quintil inferior.
    """
    return 1 + sum(valor > c for c in cortes)


def _ticks_simetricos(vmax: float, en_pct: bool,
                      limite: float) -> tuple[list[float], list[str]]:
    """Ticks del eje X espejados en torno a cero, con etiquetas en valor absoluto.

    En una pirámide el eje X va de -max a +max, pero el lector debe ver
    cantidades positivas a los dos lados.

    `limite` es el extremo real del eje: no se generan ticks más allá, para
    no dejar etiquetas colgando fuera del área de dibujo. Con efectivos
    (en_pct=False) el paso se fuerza a entero: "2,5 personas" no significa nada.
    """
    if vmax <= 0:
        return [0], ["0"]

    objetivo = vmax / 3.2                       # apuntamos a ~4 ticks por lado
    exp = math.floor(math.log10(objetivo)) if objetivo > 0 else 0
    magnitud = 10.0 ** exp
    paso = next(
        (magnitud * m for m in (1, 2, 2.5, 5, 10) if magnitud * m >= objetivo),
        magnitud * 10,
    )
    if not en_pct:
        paso = max(1.0, round(paso))

    n = int(limite / paso)
    vals = [i * paso for i in range(-n, n + 1)]
    dec = 1 if en_pct else 0
    txt = [
        (fmt_dec(abs(v), dec) + " %") if en_pct else fmt_int(abs(v))
        for v in vals
    ]
    return vals, txt


# ==========================================================================
# §5  COMPONENTES VISUALES
# ==========================================================================


def fila_kpis(k: dict) -> None:
    """Los indicadores de cabecera, apilados en la mitad izquierda.

    Uno por fila, salvo los dos umbrales de edad -67 y más, menores de 18-,
    que comparten fila porque se leen en pareja. El conjunto ocupa
    aproximadamente el alto del mapa que va al lado.
    Los porcentajes llevan la cifra absoluta en el tooltip, para no perder el
    dato de partida.
    """
    reparto = (f"Reparto uniforme del grupo de corte: se toman "
               f"{FRACCION_INTERVALO:.0%} de sus efectivos.")

    fec = k["fecundidad"]
    ayuda_fec = (
        "Hijos por mujer. Se estima como la población de 0 a 4 años entre 5 "
        "-los nacimientos de un año- dividida entre las mujeres de 15 a 49 "
        f"años entre 35: {fmt_int(k['nacimientos_ano'])} / "
        f"{fmt_int(k['mujeres_fertiles'] / 35)}. "
        "En municipios muy pequeños el resultado es inestable, porque el "
        "denominador es de unas pocas mujeres."
    )

    filas = [
        [("Población (2025)", fmt_int(k["poblacion"]), None)],
        [("Edad media (est.)", fmt_dec(k["edad_media"], 1) + " años",
          "Estimada con el punto medio de cada intervalo de edad: el censo no "
          "publica la edad exacta por municipio.")],
        [("67 y más años", fmt_dec(k["pct_mayores67"], 1) + " %",
          f"{fmt_int(k['n_mayores67'])} personas. {reparto}"),
         ("Menores de 18 años", fmt_dec(k["pct_menores18"], 1) + " %",
          f"{fmt_int(k['n_menores18'])} personas. {reparto}")],
        [("Nacidos fuera de España", fmt_dec(k["pct_extranjero"], 1) + " %",
          f"{fmt_int(k['n_extranjero'])} personas.")],
        [("Índice de fecundidad",
          "—" if fec is None else fmt_dec(fec, 2),
          None if fec is None else ayuda_fec)],
    ]
    for tramo in filas:
        columnas = st.columns(len(tramo), gap="small")
        for col, (etiqueta, valor, ayuda) in zip(columnas, tramo):
            col.metric(etiqueta, valor, help=ayuda, border=True)


def bloque_mapa(cod: str, nombre: str) -> None:
    """Mapa del municipio con su rejilla, o el aviso que corresponda."""
    if not GEO_DISPONIBLE:
        st.info(
            "Falta la pila geoespacial. Instálala con "
            "`py -3.10 -m pip install geopandas pyogrio`.",
            icon="🧩",
        )
        return
    if not (GEO_MUNICIPIOS.exists() and GEO_GRIDS.exists()):
        st.info(
            "Faltan los ficheros de geometría. Genéralos una vez con "
            "`py -3.10 preprocessGrids.py`.",
            icon="🧩",
        )
        return

    geo_muni = cargar_geometria_municipio(cod)
    if geo_muni is None:
        st.info(f"No hay geometría para **{nombre}** (código {cod}).", icon="🗺️")
        return

    celdas = cargar_celdas_municipio(cod)
    st.plotly_chart(
        mapa_municipio(geo_muni, celdas),
        width="stretch",
        config={"displayModeBar": False, "scrollZoom": True},
    )

    if celdas is None:
        st.caption(
            "Ninguna celda de 1 km habitada solapa este término municipal."
        )
    else:
        n = len(celdas)
        st.caption(
            f"{fmt_int(n)} celda{'s' if n != 1 else ''} habitada"
            f"{'s' if n != 1 else ''} · "
            f"la más poblada, {fmt_int(celdas['poblacion'].max())} habitantes. "
            "Basta que la celda entre un poco en el término para que aparezca, "
            "así que las del borde se comparten con los municipios vecinos."
        )


def resolver_color(clave: str) -> str:
    """Admite una clave de la paleta C ("s1") o un hexadecimal literal.

    Los desgloses del bloque inferior llevan colores semánticos propios, que
    viven en su configuración y no en la paleta general.
    """
    return clave if clave.startswith("#") else C[clave]


def cabecera_grafico(titulo: str, subtitulo: str = "") -> None:
    """Título y subtítulo de una tarjeta de gráfico, fuera de la figura."""
    st.markdown(f'<p class="card-title">{titulo}</p>', unsafe_allow_html=True)
    if subtitulo:
        st.markdown(f'<p class="card-sub">{subtitulo}</p>',
                    unsafe_allow_html=True)


def fila_kpis_laborales(k: dict) -> None:
    """Tasa de paro y tasa de empleo, una al lado de la otra."""
    sin_datos = k["tasa_paro"] is None and k["tasa_empleo"] is None
    datos = [
        ("Tasa de paro (2023)",
         "—" if k["tasa_paro"] is None else fmt_dec(k["tasa_paro"], 1) + " %",
         None if sin_datos else
         f"Parados / (parados + ocupados) = {fmt_int(k['parados'])} / "
         f"{fmt_int(k['activos'])}"),
        ("Tasa de empleo (2023)",
         "—" if k["tasa_empleo"] is None else fmt_dec(k["tasa_empleo"], 1) + " %",
         None if sin_datos else
         f"Ocupados / población de 15 a 64 años = {fmt_int(k['ocupados'])} / "
         f"{fmt_int(k['pob_15_64'])}"
         + ("  ⚠ Supera el 100 % porque el numerador incluye a los ocupados "
            "de 65 y más años, que el denominador no cuenta. La hoja de "
            "actividad no viene desglosada por edad."
            if (k["tasa_empleo"] or 0) > 100 else "")),
    ]
    for col, (etiqueta, valor, ayuda) in zip(st.columns(2, gap="small"), datos):
        col.metric(etiqueta, valor, help=ayuda, border=True)


# Alto reservado a cada barra horizontal, en píxeles. La figura crece con el
# número de categorías en vez de apretarlas en un alto fijo.
ALTO_POR_BARRA = 74
ALTO_BASE_BARRAS = 32


def barras_categorias(fila: pd.Series, categorias: list[tuple[str, str, str]],
                      col_total: str = "Total") -> go.Figure:
    """Barras horizontales con el reparto de un municipio entre categorías.

    fila       : Series de una hoja municipal (Total + las categorías).
    categorias : lista de (columna del Excel, etiqueta, clave de color), en el
                 orden en que se quieren leer de arriba abajo.

    Al ser una sola serie no lleva leyenda: cada barra va rotulada en el eje.
    El valor absoluto se escribe al final de la barra y el porcentaje justo
    debajo, así que ninguna cifra depende del tooltip.
    """
    total = float(fila[col_total])

    columnas, etiquetas, valores, pcts, colores = [], [], [], [], []
    for col, etq, color in categorias:
        v = float(fila[col])
        columnas.append(col)
        etiquetas.append(etq)
        valores.append(v)
        pcts.append(100 * v / total if total else 0.0)
        colores.append(resolver_color(color))

    # Plotly ordena el eje categórico de abajo hacia arriba: se invierte para
    # que la primera categoría quede arriba.
    columnas, etiquetas, valores, pcts, colores = (
        list(reversed(x))
        for x in (columnas, etiquetas, valores, pcts, colores)
    )

    fig = go.Figure(go.Bar(
        y=etiquetas, x=pcts, orientation="h",
        marker=dict(color=colores),
        width=0.46,   # marcas finas, con aire entre ellas
        # El nombre largo del INE va al tooltip: las etiquetas del eje se
        # acortan para que quepan, pero la definición no se pierde.
        customdata=[[fmt_int(v), fmt_dec(p, 1), col]
                    for v, p, col in zip(valores, pcts, columnas)],
        hovertemplate="<b>%{y}</b><br>%{customdata[0]} personas"
                      " · %{customdata[1]} %"
                      "<br><i>%{customdata[2]}</i><extra></extra>",
        showlegend=False,
    ))

    fig.update_layout(
        height=ALTO_BASE_BARRAS + ALTO_POR_BARRA * len(etiquetas),
        margin=dict(l=92, r=32, t=16, b=16),
        # Espacio a la derecha para que quepa el rótulo de dos líneas.
        xaxis=dict(range=[0, max(pcts) * 1.34], showgrid=False,
                   showticklabels=False, visible=False),
        yaxis=dict(showgrid=False, ticksuffix="   "),
    )

    # Rótulos como anotaciones y no como `text` de la barra: la cifra absoluta
    # arriba y el porcentaje debajo, alineados a la izquierda. El `text` de
    # Plotly centra el bloque de dos líneas y el porcentaje queda indentado.
    # Van con tinta de texto, no con el color de la serie, así que tampoco hay
    # problema de contraste sobre el relleno.
    for etq, v, p in zip(etiquetas, valores, pcts):
        fig.add_annotation(
            x=p, y=etq, xref="x", yref="y",
            text=f"{fmt_int(v)}<br>{fmt_dec(p, 1)} %",
            showarrow=False, align="left",
            xanchor="left", yanchor="middle", xshift=10,
            font=dict(color=C["ink_2"], size=13, family=FONT_STACK),
        )
    try:
        fig.update_layout(barcornerradius=4)
    except Exception:
        pass
    return fig


def _vista_del_mapa(bounds, ancho_px: int, alto_px: int) -> tuple[dict, float]:
    """Centro y zoom para que el municipio quepa entero en el lienzo.

    `bounds` es (oeste, sur, este, norte) en grados. El zoom sale de la
    aritmética de Web Mercator: a zoom z el mundo mide TESELA·2^z píxeles.
    Ojo con la tesela: MapLibre -el motor que usa Plotly en las trazas `map`-
    trabaja con teselas de 512 px, no de 256. Usar 256 da un zoom un nivel más
    alto y el término se sale del lienzo por la mitad.
    """
    oeste, sur, este, norte = bounds
    centro = {"lon": (oeste + este) / 2, "lat": (sur + norte) / 2}
    TESELA = 512

    span_lon = max(este - oeste, 1e-6)
    zoom_lon = math.log2(ancho_px * 360 / (TESELA * span_lon))

    def mercator_y(lat: float) -> float:
        lat = max(min(lat, 89.9), -89.9)
        return math.log(math.tan(math.pi / 4 + math.radians(lat) / 2))

    span_y = max(mercator_y(norte) - mercator_y(sur), 1e-9)
    zoom_lat = math.log2(alto_px * 2 * math.pi / (TESELA * span_y))

    # Se queda con la dimensión más restrictiva y deja un margen para que el
    # término no toque los bordes.
    zoom = min(zoom_lon, zoom_lat) - 0.35
    return centro, max(min(zoom, 16.0), 3.0)


def mapa_municipio(geo_muni, celdas, alto: int = 430) -> go.Figure:
    """Mapa de calles con el término municipal y su rejilla de 1 km.

    Tres capas, de abajo arriba: el mapa de calles, las celdas de 1 km
    coloreadas por población y el límite del municipio por encima.

    La escala de color es logarítmica: la población por celda va de 1 a más de
    50.000 y en lineal todo se vería del mismo tono menos la celda más densa.
    La barra de color rotula los valores reales, no sus logaritmos.
    """
    fig = go.Figure()
    geo_limite = json.loads(geo_muni.to_json())

    if celdas is not None and len(celdas):
        celdas = celdas.copy()
        # Una celda puede repetirse si aparece por varios municipios; aquí ya
        # viene filtrada a uno, pero por si acaso se deduplica por su id.
        celdas = celdas.drop_duplicates(subset="GRD_ID").set_index("GRD_ID")
        geo_celdas = json.loads(celdas.to_json())

        pobl = celdas["poblacion"].astype(float)
        z = pobl.map(math.log10)
        # La rampa se estira hasta el máximo del municipio, para aprovechar
        # todo el contraste dentro de este mapa. Los colores no son
        # comparables entre municipios, pero las marcas de la barra sí: van
        # rotuladas con cifras absolutas, así que un tono siempre se puede
        # traducir a habitantes leyendo la escala.
        # El margen de 0,08 décadas evita que la última marca caiga justo en
        # el extremo de la barra, donde Plotly la esconde.
        zmax = max(float(z.max()) + 0.08, 1.0)

        # Marcas de la barra de color en potencias de diez dentro del rango.
        tickvals, ticktext = [], []
        for e in range(0, 8):
            if e <= zmax:
                tickvals.append(e)
                ticktext.append(fmt_int(10 ** e))

        fig.add_trace(go.Choroplethmap(
            geojson=geo_celdas,
            locations=list(celdas.index),
            z=z.tolist(),
            zmin=0, zmax=zmax,
            colorscale=ESCALA_POBLACION,
            # Opacidad moderada: la idea de poner un mapa de calles debajo es
            # que se sigan viendo.
            marker=dict(opacity=0.66, line=dict(width=0.5,
                                                color="rgba(255,255,255,0.45)")),
            # Barra horizontal dentro del mapa, abajo a la izquierda: puesta
            # fuera le quitaría anchura al mapa, y a la derecha choca con la
            # atribución de Carto.
            # Panel claro con tinta oscura: sobre un mapa base claro, un panel
            # oscuro se vería como una caja negra pegada encima.
            colorbar=dict(
                orientation="h",
                title=dict(text="habitantes por km²", side="top",
                           font=dict(size=11, color="#33322f")),
                tickvals=tickvals, ticktext=ticktext,
                tickfont=dict(size=11, color="#33322f"),
                # Sin esto Plotly esconde la última marca, que cae justo en el
                # extremo de la barra y es la más importante.
                ticklabeloverflow="allow",
                thickness=10, len=0.44,
                x=0.02, xanchor="left", y=0.04, yanchor="bottom",
                xpad=10, ypad=8,
                bgcolor="rgba(252,252,251,0.86)", outlinewidth=0,
            ),
            customdata=[[fmt_int(v)] for v in pobl],
            hovertemplate="%{customdata[0]} habitantes<extra></extra>",
        ))
    else:
        # Sin celdas habitadas no hay traza: el mapa se queda con el término.
        fig.add_trace(go.Scattermap(lat=[], lon=[], mode="markers",
                                    showlegend=False, hoverinfo="skip"))

    centro, zoom = _vista_del_mapa(tuple(geo_muni.total_bounds), 700, alto)

    fig.update_layout(
        map=dict(
            style=ESTILO_MAPA,
            center=centro,
            zoom=zoom,
            # Las capas de layout se dibujan por encima de las trazas, así que
            # el límite queda visible sobre las celdas. Van dos: una blanca
            # ancha debajo y otra oscura fina encima. Ese doble trazo es lo que
            # hace que la línea se lea igual sobre un parque verde, sobre el
            # agua azul o sobre una celda violeta oscura.
            layers=[
                dict(sourcetype="geojson", source=geo_limite, type="line",
                     color="rgba(255,255,255,0.9)", line=dict(width=5.0)),
                dict(sourcetype="geojson", source=geo_limite, type="line",
                     color="#1b1b1a", line=dict(width=1.8)),
            ],
        ),
        height=alto,
        margin=dict(l=0, r=0, t=0, b=0),
        showlegend=False,
    )
    return fig


def _escala_discreta(colores: list[str]) -> list[list]:
    """Convierte una lista de colores en una escala Plotly de tramos planos.

    Con n colores, un valor z = i + 0,5 cae de lleno en el color i, así que
    una variable categórica se pinta sin degradados intermedios.
    """
    n = len(colores)
    escala = []
    for i, color in enumerate(colores):
        escala.append([i / n, color])
        escala.append([(i + 1) / n, color])
    return escala


def mapa_nacional(categoria: pd.Series, etiquetas: list[str],
                  colores: list[str], texto: pd.Series,
                  titulo_leyenda: str, alto: int = 720) -> go.Figure:
    """Coropleta municipal de España con Canarias en un recuadro aparte.

    categoria : Serie indexada por código INE con el índice de categoría de
                cada municipio (0..n-1). Los que falten se pintan como
                "Sin datos", que es la última categoría.
    etiquetas : nombre de cada categoría, sin incluir "Sin datos".
    colores   : color de cada categoría, en el mismo orden.
    texto     : Serie indexada por código con el texto del tooltip.
    """
    etiquetas = list(etiquetas) + ["Sin datos"]
    colores = list(colores) + [COLOR_SIN_DATO]
    n = len(colores)
    sin_dato = n - 1

    # La barra de color se dibuja de abajo arriba, así que se invierte todo
    # para que la leyenda se lea de arriba abajo en el orden dado y "Sin datos"
    # quede al final. El índice interno es n-1-categoría.
    escala = _escala_discreta(list(reversed(colores)))
    etiquetas_barra = list(reversed(etiquetas))

    fig = go.Figure()
    ambitos = [("peninbal", "geo", DOMINIO_PENINSULA, True),
               ("canarias", "geo2", DOMINIO_CANARIAS, False)]

    for ambito, subplot, dominio, con_leyenda in ambitos:
        geo = cargar_geojson_municipal(ambito)
        if geo is None:
            continue

        codigos = [f["id"] for f in geo["features"]]
        # Los municipios sin dato van a la última categoría en vez de dejar un
        # hueco: un agujero en el mapa se lee como error, no como ausencia.
        z = [n - 1 - float(categoria.get(c, sin_dato)) + 0.5 for c in codigos]
        htxt = [texto.get(c, "Sin datos publicados") for c in codigos]

        fig.add_trace(go.Choropleth(
            geojson=geo, featureidkey="id",
            locations=codigos, z=z,
            zmin=0, zmax=n,
            colorscale=escala,
            marker=dict(line=dict(width=0.15, color="rgba(0,0,0,0.45)")),
            showscale=con_leyenda,
            colorbar=dict(
                title=dict(text=titulo_leyenda, side="top",
                           font=dict(size=12, color=C["ink_2"])),
                # Una marca en el centro de cada tramo: así la barra de color
                # funciona como leyenda de una variable categórica.
                tickmode="array",
                tickvals=[i + 0.5 for i in range(n)],
                ticktext=etiquetas_barra,
                tickfont=dict(size=12, color=C["ink_2"]),
                ticks="", ticklen=0,
                thickness=14, len=0.44,
                x=0.985, xanchor="right", y=0.98, yanchor="top",
                # Panel más opaco que antes: ahora el mapa llega por debajo de
                # la leyenda, así que tiene que taparlo para leerse.
                outlinewidth=0, bgcolor="rgba(13,13,13,0.88)",
                ypad=12, xpad=12,
            ),
            geo=subplot,
            text=htxt,
            hovertemplate="%{text}<extra></extra>",
        ))

        # Límites provinciales por encima de la retícula municipal, con trazo
        # más grueso. Se dibujan después para que queden sobre los rellenos.
        lineas = cargar_lineas_provinciales()
        if lineas and ambito in lineas:
            fig.add_trace(go.Scattergeo(
                lon=lineas[ambito]["lon"], lat=lineas[ambito]["lat"],
                mode="lines",
                line=dict(width=GROSOR_LINEA_PROVINCIA,
                          color=COLOR_LINEA_PROVINCIA),
                hoverinfo="skip", showlegend=False, geo=subplot,
            ))

        # El recuadro de Canarias lleva marco para que se lea como lo que es:
        # un encarte, no parte del encuadre peninsular. Y fondo OPACO, porque
        # el mapa principal ocupa ahora todo el lienzo y pasa por debajo: sin
        # fondo se vería el Mediterráneo asomando entre las islas.
        es_encarte = subplot != "geo"
        # Los dos llevan encuadre fijo en vez de fitbounds, porque fitbounds
        # recorta la ventana y el mapa se pierde al arrastrarlo.
        encuadre = ENCUADRE_CANARIAS if es_encarte else ENCUADRE_PENINSULA
        ajuste = dict(lonaxis=dict(range=encuadre["lon"]),
                      lataxis=dict(range=encuadre["lat"]),
                      projection=dict(type="mercator"))
        if not es_encarte:
            # La península añade zoom y centro: los rangos de arriba solo
            # marcan hasta dónde se puede arrastrar, y es el zoom el que fija
            # el tamaño al que se ve el mapa. El encarte no los necesita: ahí
            # los rangos sí encuadran, porque no hace falta arrastrarlo.
            ajuste["center"] = CENTRO_PENINSULA
            ajuste["projection"]["scale"] = ZOOM_PENINSULA
        fig.update_layout({subplot: dict(
            visible=False,        # sin costas ni relleno: los municipios ya
                                  # dibujan la tierra
            showframe=es_encarte,
            framecolor="rgba(255,255,255,0.22)", framewidth=1,
            bgcolor=C["surface"] if es_encarte else "rgba(0,0,0,0)",
            domain=dominio,
            **ajuste,
        )})

    fig.update_layout(
        height=alto,
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor=C["surface"],
        plot_bgcolor=C["surface"],
        annotations=[dict(
            text="Canarias",
            # Centrado en el dominio del encarte y no pegado a su esquina:
            # fitbounds deja margen al encajar, así que el marco visible es
            # más estrecho que el dominio y una etiqueta en la esquina se queda
            # fuera. Centrada siempre cae sobre el encarte.
            x=sum(DOMINIO_CANARIAS["x"]) / 2,
            y=DOMINIO_CANARIAS["y"][1] - 0.012,
            xref="paper", yref="paper",
            xanchor="center", yanchor="top", showarrow=False,
            font=dict(size=11, color=C["ink_muted"], family=FONT_STACK),
        )],
    )
    return fig


def barra_renta(valor: float, contexto: dict) -> go.Figure:
    """Barra de renta del municipio sobre un carril que llega al máximo nacional.

    El carril es contexto visual, no un dato que haya que leer: va sin rótulo y
    sin tooltip, así que al pasar el ratón por encima no aparece nada.
    """
    maximo = contexto["maximo"]

    fig = go.Figure()
    # Carril de fondo: llega hasta la renta del municipio más alto de España.
    fig.add_trace(go.Bar(
        y=[""], x=[maximo], orientation="h", width=0.52,
        marker=dict(color=C["carril"]),
        hoverinfo="skip", hovertemplate=None,
        showlegend=False,
    ))
    # Dato del municipio, encima del carril.
    fig.add_trace(go.Bar(
        y=[""], x=[valor], orientation="h", width=0.52,
        marker=dict(color=C["renta"]),
        hovertemplate=f"{fmt_int(valor)} € por persona<extra></extra>",
        showlegend=False,
    ))

    # Rótulo del valor al final de la barra de color. Si la barra casi llena el
    # carril, el rótulo se mete dentro para no salirse del lienzo.
    casi_lleno = bool(maximo) and valor > 0.78 * maximo
    fig.add_annotation(
        x=valor, xref="x",
        # El eje Y es categórico y con una sola categoría no centra el rótulo,
        # así que la vertical se fija en coordenadas de lienzo.
        y=0.5, yref="paper", yanchor="middle",
        text=f"<b>{fmt_int(valor)} €</b>",
        showarrow=False,
        xanchor="right" if casi_lleno else "left",
        xshift=-12 if casi_lleno else 12,
        font=dict(color=C["ink"] if casi_lleno else C["ink_2"],
                  size=15, family=FONT_STACK),
    )

    fig.update_layout(
        barmode="overlay",
        height=104,
        margin=dict(l=8, r=8, t=8, b=8),
        xaxis=dict(range=[0, maximo], visible=False),
        yaxis=dict(visible=False),
    )
    try:
        fig.update_layout(barcornerradius=4)
    except Exception:
        pass
    return fig


def donut_estado_civil(fila: pd.Series) -> go.Figure:
    """Donut con el reparto por estado civil del municipio.

    Sin leyenda: cada porción va rotulada por fuera con su nombre y su peso,
    así que la identidad no depende del color. En el centro, el total.
    """
    etiquetas, valores, colores, columnas = [], [], [], []
    for col, etq, color in CAT_ESTADO_CIVIL:
        v = fila[col]
        valores.append(0.0 if pd.isna(v) else float(v))
        etiquetas.append(etq)
        colores.append(resolver_color(color))
        columnas.append(col)

    total = sum(valores)
    pcts = [100 * v / total if total else 0.0 for v in valores]

    fig = go.Figure(go.Pie(
        labels=etiquetas, values=valores,
        hole=0.58,
        sort=False, direction="clockwise",
        marker=dict(colors=colores,
                    # Filete del color de la superficie: separa las porciones
                    # sin dibujarles un borde.
                    line=dict(color=C["surface"], width=2)),
        text=[f"{e}<br>{fmt_dec(p, 1)} %" for e, p in zip(etiquetas, pcts)],
        textinfo="text", textposition="outside",
        textfont=dict(color=C["ink_2"], size=12, family=FONT_STACK),
        customdata=[[fmt_int(v), col] for v, col in zip(valores, columnas)],
        hovertemplate="<b>%{label}</b><br>%{customdata[0]} personas"
                      " · %{percent}<br><i>%{customdata[1]}</i><extra></extra>",
        showlegend=False,
    ))

    fig.update_layout(
        height=330,
        margin=dict(l=70, r=70, t=24, b=24),
        annotations=[dict(
            text=f"<span style='font-size:20px;color:{C['ink']}'>"
                 f"{fmt_int(total)}</span><br>"
                 f"<span style='font-size:11px;color:{C['ink_muted']}'>"
                 "personas de 16 y más años</span>",
            x=0.5, y=0.5, showarrow=False, align="center",
            font=dict(family=FONT_STACK),
        )],
    )
    return fig


def barras_verticales_categorias(
    fila: pd.Series, categorias: list[tuple[str, str, str]]
) -> go.Figure:
    """Barras verticales con el reparto de un municipio entre categorías.

    fila       : Series de una hoja municipal.
    categorias : lista de (columna del Excel, etiqueta corta para el eje,
                 clave de color), en el orden en que se quieren leer de
                 izquierda a derecha.

    Los porcentajes van sobre la suma de las categorías mostradas, así que
    siempre totalizan 100 % (ver DESGLOSES). Cada barra lleva encima la cifra
    absoluta y, debajo, su peso; el eje vertical se oculta porque no aporta
    nada que no digan ya los rótulos.
    """
    valores, etiquetas, colores, columnas = [], [], [], []
    for col, etq, color in categorias:
        v = fila[col]
        # La columna "No consta" de la hoja Actividad viene vacía en 4.144
        # municipios; un hueco cuenta como cero, no como dato ausente.
        valores.append(0.0 if pd.isna(v) else float(v))
        etiquetas.append(etq)
        colores.append(resolver_color(color))
        columnas.append(col)

    base = sum(valores)
    pcts = [100 * v / base if base else 0.0 for v in valores]

    fig = go.Figure(go.Bar(
        x=etiquetas, y=valores,
        marker=dict(color=colores),
        # A todo el ancho y con solo 4-6 categorías hay que contenerlas: sin
        # esto salen bloques de 150 px.
        width=0.36,
        # Encima de cada barra: la cifra absoluta y debajo su porcentaje. En
        # vertical el centrado que aplica Plotly es el correcto.
        text=[f"{fmt_int(v)}<br>{fmt_dec(p, 1)} %"
              for v, p in zip(valores, pcts)],
        textposition="outside",
        textfont=dict(color=C["ink_2"], size=13, family=FONT_STACK),
        cliponaxis=False,
        # El nombre completo del INE, que en el eje no cabría, va al tooltip.
        customdata=[[fmt_int(v), fmt_dec(p, 1), col]
                    for v, p, col in zip(valores, pcts, columnas)],
        hovertemplate="%{customdata[0]} personas · %{customdata[1]} %"
                      "<br><i>%{customdata[2]}</i><extra></extra>",
        showlegend=False,
    ))

    fig.update_layout(
        height=340,
        margin=dict(l=24, r=24, t=24, b=72),
        # Aire arriba para el rótulo de dos líneas.
        yaxis=dict(range=[0, max(valores) * 1.20 if max(valores) else 1],
                   showgrid=False, showticklabels=False, visible=False),
        xaxis=dict(showgrid=False,
                   tickfont=dict(size=12, color=C["ink_2"])),
    )
    try:
        fig.update_layout(barcornerradius=4)
    except Exception:
        pass
    return fig


def piramide_poblacion(
    perfil: pd.DataFrame,
    en_pct: bool,
    referencia: pd.DataFrame | None = None,
    nombre_referencia: str = "",
) -> go.Figure:
    """Pirámide de población: hombres a la izquierda, mujeres a la derecha,
    cada lado dividido en nacidos en España (tono base, pegado al eje) y
    nacidos en el extranjero (tono claro, apilado hacia fuera).

    perfil      : DataFrame grupo_edad x [H_esp, H_ext, M_esp, M_ext].
    en_pct      : si True, se expresa como % sobre la población total mostrada.
    referencia  : perfil grupo_edad x [Hombres, Mujeres] de otro territorio
                  (provincia / España) para superponer como silueta. Se reescala
                  a la población del municipio, de forma que la comparación sea
                  de FORMA, no de tamaño.
    """
    total = perfil.to_numpy().sum()
    escala = 100 / total if (en_pct and total) else 1.0
    edades = list(perfil.index)

    sufijo = " %" if en_pct else ""
    f = (lambda v: fmt_dec(v, 2)) if en_pct else (lambda v: fmt_int(v))

    fig = go.Figure()

    # Un trazo por segmento. Los hombres van al lado negativo del eje; el valor
    # real viaja en customdata para que el tooltip muestre cifras positivas.
    for clave, _sexo, lado, color, etiqueta in SEGMENTOS:
        v = perfil[clave].to_numpy() * escala
        fig.add_trace(go.Bar(
            y=edades, x=v * lado, orientation="h", name=etiqueta,
            marker=dict(
                color=C[color],
                # Filete del color de la superficie: deja aire entre los dos
                # segmentos apilados en vez de dibujarles un borde.
                line=dict(color=C["surface"], width=1),
            ),
            customdata=[f(x) + sufijo for x in v],
            hovertemplate=f"<b>{etiqueta}</b><br>%{{y}}"
                          "<br>%{customdata}<extra></extra>",
        ))

    # El extremo de cada lado es la suma de sus dos segmentos.
    vmax = max(
        (perfil["H_esp"] + perfil["H_ext"]).max() * escala,
        (perfil["M_esp"] + perfil["M_ext"]).max() * escala,
    )

    # Silueta de referencia: misma distribución relativa, reescalada al tamaño
    # del municipio para poder compararla en cualquiera de los dos modos.
    if referencia is not None:
        ref_total = referencia.to_numpy().sum()
        if ref_total:
            factor = (100 / ref_total) if en_pct else (total / ref_total)
            rh = referencia["Hombres"].to_numpy() * factor
            rm = referencia["Mujeres"].to_numpy() * factor
            vmax = max(vmax, rh.max(initial=0), rm.max(initial=0))
            for i, (serie, lado) in enumerate(((rh, -1), (rm, 1))):
                fig.add_trace(go.Scatter(
                    y=edades, x=serie * lado, mode="lines",
                    name=f"Perfil de {nombre_referencia}",
                    legendgroup="ref", showlegend=(i == 0),
                    line=dict(color=C["ink"], width=2),
                    customdata=[f(v) + sufijo for v in serie],
                    hovertemplate=(f"<b>{nombre_referencia}</b> · %{{y}}"
                                   "<br>%{customdata}<extra></extra>"),
                ))

    limite = vmax * 1.08
    vals, txt = _ticks_simetricos(vmax, en_pct, limite)

    fig.update_layout(
        # "relative" apila los negativos hacia la izquierda y los positivos
        # hacia la derecha, que es justo la geometría de la pirámide.
        barmode="relative",
        bargap=0.42,                # marcas finas, con aire entre grupos
        height=706,
        hovermode="closest",
        # Sitio para los rótulos de edad ("100 y más años") a la izquierda y,
        # arriba, para la leyenda: el título vive en la cabecera de la tarjeta.
        margin=dict(l=124, r=24, t=52, b=56),
        xaxis=dict(
            tickvals=vals, ticktext=txt,
            range=[-limite, limite],
            title=dict(text="% sobre la población total" if en_pct else "Personas"),
        ),
        yaxis=dict(
            showgrid=False,
            categoryorder="array", categoryarray=edades,
            ticksuffix="   ",
        ),
    )
    # Separación central: una banda del color de la superficie deja 2 px de aire
    # entre el relleno azul y el naranja, en vez de dibujar un borde.
    fig.add_vline(x=0, line_width=3, line_color=C["surface"])

    # Extremos de barra redondeados (Plotly >= 5.22). Si la versión instalada
    # no lo soporta, el gráfico sale igual con esquinas rectas.
    try:
        fig.update_layout(barcornerradius=4)
    except Exception:
        pass

    return fig


# ==========================================================================
# §6  PANTALLA: FICHA MUNICIPAL
# ==========================================================================


def pantalla_ficha_municipal(
    muni: pd.DataFrame,
    ref: pd.DataFrame,
    act_muni: pd.DataFrame,
    sitprof_muni: pd.DataFrame,
    desgloses: dict[str, pd.DataFrame],
    renta_serie: pd.Series,
    renta_ctx: dict,
    civil_muni: pd.DataFrame,
) -> None:
    # ---- Cabecera -------------------------------------------------------
    st.markdown('<p class="dash-eyebrow">Censo de Población · INE</p>',
                unsafe_allow_html=True)
    st.markdown('<h1 class="dash-title">Ficha Municipal</h1>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="dash-sub">Perfil demográfico, laboral y de renta de cada '
        'municipio español.</p>',
        unsafe_allow_html=True,
    )
    st.markdown('<hr class="dash-rule">', unsafe_allow_html=True)

    # ---- Filtros: una sola fila por encima del contenido ----------------
    catalogo = (
        muni.drop_duplicates("cod")[["cod", "cod_prov", "nombre", "etiqueta"]]
        .sort_values("etiqueta")
        .reset_index(drop=True)
    )

    c1, c2, _hueco = st.columns([3, 2, 2], gap="medium")

    # Municipio que aparece al abrir la app.
    coincide = catalogo.index[catalogo["etiqueta"] == MUNICIPIO_INICIAL]
    idx_inicial = int(coincide[0]) if len(coincide) else 0

    with c1:
        etiqueta_sel = st.selectbox(
            "Municipio",
            options=catalogo["etiqueta"],
            index=idx_inicial,
            help=f"Escribe para buscar entre los {fmt_int(len(catalogo))} municipios.",
        )
    with c2:
        comparar_con = st.selectbox(
            "Comparar perfil con", options=["—", "Su provincia", "España"], index=2,
            help="Superpone la silueta del territorio elegido para comparar la "
                 "forma de la pirámide, no su tamaño: en modo absoluto se "
                 "reescala a la población del municipio y en porcentaje se "
                 "muestra su propio reparto por edades.",
        )

    fila = catalogo[catalogo["etiqueta"] == etiqueta_sel].iloc[0]
    df_muni = muni[muni["cod"] == fila["cod"]]

    # ---- KPIs a la izquierda, mapa del municipio a la derecha ------------
    st.write("")
    col_kpis, col_mapa = st.columns(2, gap="medium")

    with col_kpis:
        fila_kpis(calcular_kpis(df_muni))

    with col_mapa:
        with st.container(border=True):
            cabecera_grafico(
                "Población por rejilla de 1 km (2021)",
                f"{fila['nombre']} · celdas habitadas sobre el término municipal",
            )
            bloque_mapa(fila["cod"], fila["nombre"])

    st.write("")

    # ---- Pirámide -------------------------------------------------------
    perfil = perfil_por_origen(df_muni)

    perfil_ref, nombre_ref = None, ""
    if comparar_con == "Su provincia":
        prov = ref[ref["cod"] == fila["cod_prov"]]
        if not prov.empty:
            perfil_ref = perfil_edad_sexo(prov)
            nombre_ref = prov["nombre"].iloc[0]
    elif comparar_con == "España":
        nacional = ref[ref["cod"].isna()]
        if not nacional.empty:
            perfil_ref = perfil_edad_sexo(nacional)
            nombre_ref = "España"

    izq, der = st.columns([2, 1], gap="medium")

    with izq:
        en_pct = st.toggle(
            "Ver en porcentaje", value=False,
            help="Necesario para comparar municipios de tamaños muy distintos.",
        )
        with st.container(border=True):
            cabecera_grafico(
                "Pirámide de población (2025)",
                f"{fila['nombre']} · por sexo y lugar de nacimiento",
            )
            st.plotly_chart(
                piramide_poblacion(
                    perfil,
                    en_pct=en_pct,
                    referencia=perfil_ref,
                    nombre_referencia=nombre_ref,
                ),
                width="stretch",
                config={"displayModeBar": False},
            )

        # Los valores nunca deben estar accesibles solo por el tooltip.
        with st.expander("Ver tabla de datos"):
            tabla = perfil.rename(columns={
                "H_esp": "Hombres · España", "H_ext": "Hombres · Extranjero",
                "M_esp": "Mujeres · España", "M_ext": "Mujeres · Extranjero",
            })
            tabla["Total"] = tabla.sum(axis=1)
            st.dataframe(
                tabla.rename_axis("Grupo de edad").astype(int),
                width="stretch",
                height=460,
            )

    with der:
        # ---- Mercado laboral ---------------------------------------------
        fila_act = (act_muni.loc[fila["cod"]]
                    if fila["cod"] in act_muni.index else None)
        k_lab = calcular_kpis_laborales(fila_act,
                                       poblacion_en_edad_trabajar(df_muni))

        st.markdown("&nbsp;", unsafe_allow_html=True)
        fila_kpis_laborales(k_lab)
        st.write("")

        if fila_act is None:
            st.info(
                f"El INE no publica los datos laborales de "
                f"**{fila['nombre']}**. Omite los 569 municipios más pequeños "
                "por confidencialidad estadística.",
                icon="🔒",
            )
        else:
            with st.container(border=True):
                cabecera_grafico(
                    "Relación con la actividad (2024)",
                    f"{fila['nombre']} · población de 16 y más años "
                    f"({fmt_int(fila_act['Total'])})",
                )
                st.plotly_chart(
                    barras_categorias(fila_act, CAT_ACTIVIDAD),
                    width="stretch",
                    config={"displayModeBar": False},
                )

        # ---- Situación profesional ---------------------------------------
        # Misma cobertura que la hoja de actividad (idénticos 7.563
        # municipios), pero se comprueba por separado para no depender de eso.
        fila_sp = (sitprof_muni.loc[fila["cod"]]
                   if fila["cod"] in sitprof_muni.index else None)
        if fila_sp is not None:
            st.write("")
            with st.container(border=True):
                cabecera_grafico(
                    "Situación profesional (2023)",
                    f"{fila['nombre']} · población ocupada "
                    f"({fmt_int(fila_sp['Total'])})",
                )
                st.plotly_chart(
                    barras_categorias(fila_sp, CAT_SITUACION_PROF),
                    width="stretch",
                    config={"displayModeBar": False},
                )

        st.caption(
            f"Código INE del municipio: **{fila['cod']}** — es la clave para "
            "unir con los shapefiles del IGN en la pantalla de mapas."
        )

    # ---- Desglose seleccionable, a todo el ancho -------------------------
    st.write("")
    with st.container(border=True):
        # El desplegable se pinta antes que el título aunque quede a su
        # derecha: así el título ya sabe qué se ha seleccionado.
        c_tit, c_sel = st.columns([3, 1], gap="medium",
                                  vertical_alignment="center")
        with c_sel:
            desglose_sel = st.selectbox(
                "Desglose", options=list(DESGLOSES), index=0,
                key="desglose_sel",
                help="Cambia el gráfico entre los tres desgloses del censo.",
            )
        cfg = DESGLOSES[desglose_sel]
        tabla_desglose = desgloses[cfg["hoja"]]
        fila_desglose = (tabla_desglose.loc[fila["cod"]]
                         if fila["cod"] in tabla_desglose.index else None)

        with c_tit:
            if fila_desglose is None:
                cabecera_grafico(cfg["titulo"], fila["nombre"])
            else:
                total_mostrado = sum(
                    0.0 if pd.isna(fila_desglose[col]) else float(fila_desglose[col])
                    for col, _e, _c in cfg["categorias"]
                )
                cabecera_grafico(
                    cfg["titulo"],
                    f"{fila['nombre']} · {cfg['base'].lower()} "
                    f"({fmt_int(total_mostrado)})",
                )

        if fila_desglose is None:
            st.info(
                f"El INE no publica este desglose de **{fila['nombre']}**. "
                "Omite los 569 municipios más pequeños por confidencialidad "
                "estadística.",
                icon="🔒",
            )
        else:
            st.plotly_chart(
                barras_verticales_categorias(fila_desglose, cfg["categorias"]),
                width="stretch",
                config={"displayModeBar": False},
            )

    # ---- Renta y estado civil, uno al lado del otro ----------------------
    st.write("")
    col_renta, col_civil = st.columns(2, gap="medium")

    with col_renta:
        with st.container(border=True):
            cabecera_grafico(
                "Renta (2023)",
                f"{fila['nombre']} · renta bruta media por persona",
            )
            valor_renta = renta_serie.get(fila["cod"])

            if valor_renta is None or pd.isna(valor_renta):
                st.info(
                    f"El INE no publica la renta de **{fila['nombre']}**: "
                    f"la marca con un punto, como en otros "
                    f"{fmt_int(renta_ctx['n_sin_dato'] - 1)} municipios.",
                    icon="🔒",
                )
            else:
                valor_renta = float(valor_renta)
                q = quintil_renta(valor_renta, renta_ctx["cortes"])
                st.plotly_chart(
                    barra_renta(valor_renta, renta_ctx),
                    width="stretch",
                    config={"displayModeBar": False},
                )
                k1, k2 = st.columns(2, gap="small")
                k1.metric(
                    "Quintil de renta", f"{q} de 5", border=True,
                    help="1 = quintil de renta más baja, 5 = más alta. "
                         "Calculado sobre los "
                         f"{fmt_int(renta_ctx['n_con_dato'])} municipios con "
                         "dato publicado.",
                )
                k2.metric(
                    "Mediana nacional",
                    fmt_int(renta_ctx["mediana"]) + " €", border=True,
                    help="Mediana de la renta bruta media por persona entre "
                         "los municipios españoles.",
                )

    with col_civil:
        with st.container(border=True):
            fila_civil = (civil_muni.loc[fila["cod"]]
                          if fila["cod"] in civil_muni.index else None)
            cabecera_grafico(
                "Estado civil (2024)",
                f"{fila['nombre']} · población de 16 y más años",
            )
            if fila_civil is None:
                st.info(
                    f"El INE no publica el estado civil de "
                    f"**{fila['nombre']}**. Omite los 569 municipios más "
                    "pequeños por confidencialidad estadística.",
                    icon="🔒",
                )
            else:
                st.plotly_chart(
                    donut_estado_civil(fila_civil),
                    width="stretch",
                    config={"displayModeBar": False},
                )


# ==========================================================================
# §7  PANTALLA: MAPAS
# ==========================================================================


def _capa_categoria(clave_mapa: str, mtime: float,
                    catalogo: pd.DataFrame) -> dict:
    """Capa de un mapa de categoría predominante.

    Sirve para el sector económico y para el nivel educativo: cambian la hoja,
    las categorías, los colores y el método, y todo eso vive en
    MAPAS_CATEGORIA.
    """
    cfg = MAPAS_CATEGORIA[clave_mapa]
    relativo = cfg["metodo"] == "relativo"
    tabla, medias = categoria_predominante(clave_mapa, mtime)
    etiquetas = [etq for etq, _cols, _color in cfg["grupos"]]
    indice = {etq: i for i, etq in enumerate(etiquetas)}

    nombres = catalogo.set_index("cod")["nombre"]

    def tooltip(fila) -> str:
        lineas = [f"<b>{nombres.get(fila.name, fila.name)}</b>",
                  f"Predomina: {fila['ganador']}"]
        for etq in etiquetas:
            linea = f"{etq}: {fmt_dec(fila['pct_' + etq], 1)} %"
            # El ratio solo significa algo si el mapa se decide con él.
            if relativo:
                linea += f" (x{fmt_dec(fila['ratio_' + etq], 2)})"
            lineas.append(linea)
        return "<br>".join(lineas)

    if relativo:
        referencia = ", ".join(f"{etq} {fmt_dec(medias[etq], 1)} %"
                               for etq in etiquetas)
        pie = ("Gana la categoría que más se desvía de lo normal, no la que "
               "más gente tiene. El porcentaje de cada categoría en el "
               "municipio se divide por la media de los porcentajes "
               f"municipales de España ({referencia}) y se queda el ratio "
               "mayor. El tooltip trae todos los porcentajes con sus ratios.")
    else:
        pie = ("Gana la categoría con mayor porcentaje en el municipio. Como "
               "las cuatro comparten denominador, es lo mismo que quedarse con "
               "la de más gente. El tooltip trae los cuatro porcentajes.")

    return {
        "categoria": tabla["ganador"].map(indice),
        "etiquetas": etiquetas,
        "colores": [color for _etq, _cols, color in cfg["grupos"]],
        "texto": tabla.apply(tooltip, axis=1),
        "titulo_leyenda": cfg["titulo_leyenda"],
        "pie": pie,
    }


def _capa_renta(renta_serie: pd.Series, renta_ctx: dict,
                catalogo: pd.DataFrame) -> dict:
    """Datos del mapa de quintiles de renta."""
    cortes = renta_ctx["cortes"]
    validos = renta_serie.dropna()
    quintiles = validos.map(lambda v: quintil_renta(float(v), cortes))

    nombres = catalogo.set_index("cod")["nombre"]
    texto = pd.Series({
        cod: (f"<b>{nombres.get(cod, cod)}</b><br>"
              f"{fmt_int(validos[cod])} € por persona<br>"
              f"Quintil {quintiles[cod]} de 5")
        for cod in validos.index
    })
    return {
        # Orden descendente: en una variable secuencial el valor más alto va
        # arriba en la leyenda. Q5 -> índice 0, Q1 -> índice 4.
        "categoria": 5 - quintiles,
        "etiquetas": [f"Q{q}" for q in range(5, 0, -1)],
        "colores": list(reversed(COLORES_QUINTIL)),
        "texto": texto,
        "titulo_leyenda": "Quintil de renta",
        "pie": (
            "Renta bruta media por persona, en quintiles: Q1 es el 20 % de "
            "municipios con la renta más baja y Q5 el 20 % más alta. Cortes en "
            + ", ".join(fmt_int(c) + " €" for c in cortes)
            + f". Calculado sobre los {fmt_int(renta_ctx['n_con_dato'])} "
            "municipios con dato publicado."
        ),
    }


def _capa_tasa(serie: pd.Series, catalogo: pd.DataFrame, colores: list[str],
               titulo_leyenda: str, plantilla_tooltip: str, pie: str,
               n_clases: int = 5) -> dict:
    """Capa de una tasa continua, repartida en clases de igual tamaño.

    Se agrupa en tramos y no se pinta como degradado continuo por dos razones:
    con 8.132 municipios los tramos se leen mucho mejor que un gradiente, y así
    los municipios sin dato caben como una categoría más en la leyenda en lugar
    de dejar agujeros en el mapa.

    El índice 0 es el tramo de valor MÁS ALTO, porque la leyenda se lee de
    arriba abajo; `colores` va en ese mismo orden.
    """
    valores = serie.dropna()
    clases, bordes = pd.qcut(valores, n_clases, labels=False,
                             retbins=True, duplicates="drop")
    n = int(clases.max()) + 1
    etiquetas_ascendentes = etiquetas_de_rangos(bordes)

    nombres = catalogo.set_index("cod")["nombre"]
    texto = pd.Series({
        cod: plantilla_tooltip.format(
            nombre=nombres.get(cod, cod), valor=fmt_dec(valores[cod], 1))
        for cod in valores.index
    })
    return {
        "categoria": (n - 1) - clases,
        "etiquetas": list(reversed(etiquetas_ascendentes)),
        "colores": colores,
        "texto": texto,
        "titulo_leyenda": titulo_leyenda,
        "pie": pie,
    }


def _capa_tasa_paro(mtime: float, catalogo: pd.DataFrame) -> dict:
    tasas = tasas_laborales_municipales(mtime)
    return _capa_tasa(
        tasas["paro"], catalogo,
        # Más paro arriba y en rojo; menos paro abajo y en verde.
        colores=COLORES_QUINTIL,
        titulo_leyenda="Tasa de paro",
        plantilla_tooltip="<b>{nombre}</b><br>Tasa de paro: {valor} %",
        pie=("Parados entre la suma de parados y ocupados, la misma definición "
             "que en la ficha municipal. Repartido en cinco tramos con el mismo "
             "número de municipios cada uno."),
    )


def _capa_tasa_empleo(mtime: float, catalogo: pd.DataFrame) -> dict:
    tasas = tasas_laborales_municipales(mtime)
    por_encima = int((tasas["empleo"] > 100).sum())
    return _capa_tasa(
        tasas["empleo"], catalogo,
        # Más empleo arriba y en verde; menos empleo abajo y en rojo.
        colores=list(reversed(COLORES_QUINTIL)),
        titulo_leyenda="Tasa de empleo",
        plantilla_tooltip="<b>{nombre}</b><br>Tasa de empleo: {valor} %",
        pie=("Ocupados entre la población de 15 a 64 años, la misma definición "
             "que en la ficha municipal. Repartido en cinco tramos con el mismo "
             f"número de municipios cada uno. En {por_encima} municipios pasa "
             "del 100 %: el numerador incluye a los ocupados de 65 y más años, "
             "que el denominador no cuenta."),
    )


# Opciones del desplegable de la pantalla de mapas, con la clave interna que
# decide cómo se construye la capa.
CAPAS_MAPA = {
    "Actividad principal": "cat:actividad",
    "Nivel educativo más común": "cat:nivel_estudios",
    "Renta media": "renta",
    "Tasa de paro (%)": "paro",
    "Tasa de empleo (%)": "empleo",
}


def pantalla_mapas(muni: pd.DataFrame, renta_serie: pd.Series,
                   renta_ctx: dict, mtime: float) -> None:
    st.markdown('<p class="dash-eyebrow">Censo de Población · INE</p>',
                unsafe_allow_html=True)
    st.markdown('<h1 class="dash-title">Mapas municipales</h1>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="dash-sub">Los 8.132 municipios de España, con Canarias en '
        'el recuadro inferior.</p>',
        unsafe_allow_html=True,
    )
    st.markdown('<hr class="dash-rule">', unsafe_allow_html=True)

    if not (GEOJSON_PENINBAL.exists() and GEOJSON_CANARIAS.exists()):
        st.info(
            "Faltan las geometrías simplificadas. Genéralas una vez con "
            "`py -3.10 preprocessGrids.py`.",
            icon="🧩",
        )
        return

    c_sel, _hueco = st.columns([2, 5], gap="medium")
    with c_sel:
        variable = st.selectbox(
            "Información a mostrar",
            options=list(CAPAS_MAPA), index=0, key="mapa_variable",
        )

    catalogo = muni.drop_duplicates("cod")[["cod", "nombre"]]
    clave = CAPAS_MAPA[variable]
    if clave.startswith("cat:"):
        capa = _capa_categoria(clave[4:], mtime, catalogo)
    elif clave == "renta":
        capa = _capa_renta(renta_serie, renta_ctx, catalogo)
    elif clave == "paro":
        capa = _capa_tasa_paro(mtime, catalogo)
    else:
        capa = _capa_tasa_empleo(mtime, catalogo)

    st.write("")
    with st.container(border=True):
        st.plotly_chart(
            mapa_nacional(
                capa["categoria"], capa["etiquetas"], capa["colores"],
                capa["texto"], capa["titulo_leyenda"],
            ),
            width="stretch",
            config={"displayModeBar": False, "scrollZoom": True},
        )

    n_con = int(capa["categoria"].notna().sum())
    st.caption(capa["pie"])
    st.caption(
        f"{fmt_int(n_con)} municipios con dato · "
        f"{fmt_int(len(catalogo) - n_con)} sin dato publicado."
    )


# ==========================================================================
# §8  MAIN
# ==========================================================================


def main() -> None:
    st.set_page_config(
        page_title="Dashboard Demográfico",
        page_icon="🗺️",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    registrar_plantilla()
    st.markdown(CSS, unsafe_allow_html=True)

    if not EXCEL_PATH.exists():
        st.error(f"No encuentro el fichero de datos: {EXCEL_PATH}")
        st.stop()

    mtime = EXCEL_PATH.stat().st_mtime
    muni, ref = cargar_poblacion(mtime)
    renta_serie, renta_ctx = cargar_renta(mtime)

    # Selector de pantalla, arriba del todo. Se usa un selector y no st.tabs
    # porque st.tabs ejecuta el contenido de TODAS las pestañas: el mapa
    # nacional se calcularía también estando en la ficha municipal.
    pantalla = st.segmented_control(
        "Pantalla", options=PANTALLAS, default=PANTALLAS[0],
        key="pantalla", label_visibility="collapsed",
    ) or PANTALLAS[0]

    if pantalla == "Mapas":
        pantalla_mapas(muni, renta_serie, renta_ctx, mtime)
        return

    act_muni = cargar_tabla_municipal("relacion_actividad", mtime)
    sitprof_muni = cargar_tabla_municipal("situacion_prof", mtime)
    # Las tres hojas del desplegable inferior, cacheadas por separado.
    desgloses = {
        cfg["hoja"]: cargar_tabla_municipal(cfg["hoja"], mtime)
        for cfg in DESGLOSES.values()
    }
    civil_muni = cargar_tabla_municipal("estado_civil", mtime)
    pantalla_ficha_municipal(muni, ref, act_muni, sitprof_muni, desgloses,
                             renta_serie, renta_ctx, civil_muni)


if __name__ == "__main__":
    main()
