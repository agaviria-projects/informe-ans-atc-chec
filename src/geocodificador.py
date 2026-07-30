import logging
import re
import unicodedata

from datetime import datetime
from typing import Callable

import pandas as pd

from geopy.extra.rate_limiter import RateLimiter
from geopy.geocoders import Nominatim
from geopy.location import Location

from src.config import (
    ESPERA_GEOCODIFICACION_SEGUNDOS,
    RUTA_CACHE_GEOCODIFICACION,
)
from src.normalizador_direcciones import (
    DireccionNormalizada,
    normalizar_direccion,
)


logger = logging.getLogger(__name__)


# ==========================================================
# VERSIÓN DEL PROCESO DE NORMALIZACIÓN
# ==========================================================

# Al cambiar esta versión, las búsquedas antiguas de la caché
# no se reutilizan automáticamente.
VERSION_NORMALIZADOR = "V3"


# ==========================================================
# ESTRUCTURA OFICIAL DE LA CACHÉ
# ==========================================================

COLUMNAS_CACHE = [
    "CLAVE_DIRECCION",
    "DIRECCION_ORIGINAL",
    "MUNICIPIO",
    "TIPO_DIRECCION",
    "DIRECCION_CONSULTADA",
    "LATITUD",
    "LONGITUD",
    "RESULTADO",
    "MUNICIPIO_RESPUESTA",
    "DEPARTAMENTO_RESPUESTA",
    "VALIDACION_MUNICIPIO",
    "INTENTOS",
    "DETALLE",
    "FECHA_CONSULTA",
]


class ErrorGeocodificacion(Exception):
    """
    Error controlado durante la geocodificación.
    """


# ==========================================================
# NORMALIZACIÓN DE TEXTOS
# ==========================================================

def limpiar_texto(
    valor: object,
) -> str:
    """
    Convierte un valor en texto limpio.
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""

    except (TypeError, ValueError):
        pass

    return re.sub(
        r"\s+",
        " ",
        str(valor).strip(),
    )


def normalizar_para_comparacion(
    valor: object,
) -> str:
    """
    Normaliza un texto para realizar comparaciones internas.
    """

    texto = limpiar_texto(
        valor
    ).upper()

    texto = unicodedata.normalize(
        "NFKD",
        texto,
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    texto = re.sub(
        r"[^A-Z0-9]+",
        " ",
        texto,
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto,
    )

    return texto.strip()


def construir_clave(
    direccion: object,
    municipio: object,
) -> str:
    """
    Construye la clave única de la caché.

    La versión evita reutilizar automáticamente resultados
    obtenidos con una lógica anterior de normalización.
    """

    direccion_normalizada = normalizar_para_comparacion(
        direccion
    )

    municipio_normalizado = normalizar_para_comparacion(
        municipio
    )

    return (
        f"{VERSION_NORMALIZADOR}|"
        f"{direccion_normalizada}|"
        f"{municipio_normalizado}"
    )


# ==========================================================
# CACHÉ DE GEOCODIFICACIÓN
# ==========================================================

def crear_cache_vacia() -> pd.DataFrame:
    """
    Crea una caché vacía con la estructura oficial.
    """

    return pd.DataFrame(
        columns=COLUMNAS_CACHE
    )


def cargar_cache() -> pd.DataFrame:
    """
    Lee la caché de geocodificación existente.
    """

    if not RUTA_CACHE_GEOCODIFICACION.exists():
        logger.info(
            "La caché de geocodificación todavía no existe."
        )

        return crear_cache_vacia()

    try:
        cache = pd.read_excel(
            RUTA_CACHE_GEOCODIFICACION,
            dtype=object,
            engine="openpyxl",
        )

    except PermissionError as error:
        logger.exception(
            "La caché de geocodificación está abierta."
        )

        raise ErrorGeocodificacion(
            "No fue posible leer CACHE_GEOCODIFICACION.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a ejecutar."
        ) from error

    except Exception as error:
        logger.exception(
            "No fue posible leer la caché."
        )

        raise ErrorGeocodificacion(
            "No fue posible leer CACHE_GEOCODIFICACION.xlsx."
        ) from error

    for columna in COLUMNAS_CACHE:
        if columna not in cache.columns:
            cache[columna] = pd.NA

    cache = cache[
        COLUMNAS_CACHE
    ].copy()

    cache["CLAVE_DIRECCION"] = (
        cache["CLAVE_DIRECCION"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    cache = cache[
        cache["CLAVE_DIRECCION"].ne("")
    ].copy()

    cache = cache.drop_duplicates(
        subset=["CLAVE_DIRECCION"],
        keep="last",
    )

    return cache.reset_index(
        drop=True
    )


def guardar_cache(
    cache: pd.DataFrame,
) -> None:
    """
    Guarda la caché de geocodificación.
    """

    try:
        RUTA_CACHE_GEOCODIFICACION.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        resultado = cache.copy()

        for columna in COLUMNAS_CACHE:
            if columna not in resultado.columns:
                resultado[columna] = pd.NA

        resultado = resultado[
            COLUMNAS_CACHE
        ]

        resultado = resultado.drop_duplicates(
            subset=["CLAVE_DIRECCION"],
            keep="last",
        )

        resultado.to_excel(
            RUTA_CACHE_GEOCODIFICACION,
            index=False,
            engine="openpyxl",
        )

        logger.info(
            "Caché actualizada: %s registros.",
            len(resultado),
        )

    except PermissionError as error:
        logger.exception(
            "No se pudo guardar la caché porque está abierta."
        )

        raise ErrorGeocodificacion(
            "No fue posible actualizar "
            "CACHE_GEOCODIFICACION.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a ejecutar."
        ) from error

    except Exception as error:
        logger.exception(
            "No fue posible guardar la caché."
        )

        raise ErrorGeocodificacion(
            "No fue posible guardar "
            "CACHE_GEOCODIFICACION.xlsx."
        ) from error


# ==========================================================
# CLIENTE DE GEOCODIFICACIÓN
# ==========================================================

def crear_geocodificador() -> Callable:
    """
    Crea Nominatim con control de frecuencia.
    """

    cliente = Nominatim(
        user_agent="informe_ans_atc_chec_1_0",
        timeout=20,
    )

    return RateLimiter(
        cliente.geocode,
        min_delay_seconds=(
            ESPERA_GEOCODIFICACION_SEGUNDOS
        ),
        max_retries=2,
        error_wait_seconds=2,
        swallow_exceptions=True,
    )


def consultar_direccion(
    geocodificar: Callable,
    direccion_consulta: str,
) -> list[Location]:
    """
    Consulta hasta cinco resultados para una dirección.

    Se devuelven varios candidatos porque posteriormente deben
    validarse contra el municipio esperado.
    """

    respuesta = geocodificar(
        direccion_consulta,
        exactly_one=False,
        limit=5,
        addressdetails=True,
        country_codes="co",
        language="es",
    )

    if respuesta is None:
        return []

    if isinstance(
        respuesta,
        Location,
    ):
        return [respuesta]

    return list(
        respuesta
    )


# ==========================================================
# VALIDACIÓN DE COORDENADAS
# ==========================================================

def coordenadas_validas(
    latitud: object,
    longitud: object,
) -> bool:
    """
    Valida que las coordenadas sean numéricas y plausibles.
    """

    latitud_numerica = pd.to_numeric(
        pd.Series([latitud]),
        errors="coerce",
    ).iloc[0]

    longitud_numerica = pd.to_numeric(
        pd.Series([longitud]),
        errors="coerce",
    ).iloc[0]

    if pd.isna(
        latitud_numerica
    ) or pd.isna(
        longitud_numerica
    ):
        return False

    return bool(
        -90 <= float(latitud_numerica) <= 90
        and -180 <= float(longitud_numerica) <= 180
    )


def extraer_datos_respuesta(
    ubicacion: Location,
) -> dict:
    """
    Extrae municipio y departamento de una respuesta Nominatim.
    """

    datos_crudos = (
        ubicacion.raw
        if isinstance(
            ubicacion.raw,
            dict,
        )
        else {}
    )

    direccion = datos_crudos.get(
        "address",
        {},
    )

    if not isinstance(
        direccion,
        dict,
    ):
        direccion = {}

    campos_municipio = (
        "city",
        "town",
        "village",
        "municipality",
        "county",
        "city_district",
    )

    valores_municipio = [
        limpiar_texto(
            direccion.get(campo)
        )
        for campo in campos_municipio
        if limpiar_texto(
            direccion.get(campo)
        )
    ]

    municipio_respuesta = (
        valores_municipio[0]
        if valores_municipio
        else ""
    )

    departamento_respuesta = limpiar_texto(
        direccion.get("state")
        or direccion.get("region")
    )

    pais_respuesta = limpiar_texto(
        direccion.get("country")
    )

    nombre_completo = limpiar_texto(
        datos_crudos.get("display_name")
        or ubicacion.address
    )

    return {
        "MUNICIPIOS_CANDIDATOS": valores_municipio,
        "MUNICIPIO_RESPUESTA": municipio_respuesta,
        "DEPARTAMENTO_RESPUESTA": departamento_respuesta,
        "PAIS_RESPUESTA": pais_respuesta,
        "NOMBRE_COMPLETO": nombre_completo,
    }


def validar_municipio_respuesta(
    ubicacion: Location,
    municipio_esperado: str,
    departamento_esperado: str = "Caldas",
) -> tuple[bool, dict]:
    """
    Valida que la coordenada corresponda al municipio esperado.

    No se acepta una coordenada únicamente porque Nominatim haya
    encontrado algún resultado.
    """

    datos = extraer_datos_respuesta(
        ubicacion
    )

    municipio_normalizado = normalizar_para_comparacion(
        municipio_esperado
    )

    departamento_normalizado = normalizar_para_comparacion(
        departamento_esperado
    )

    nombre_completo_normalizado = (
        normalizar_para_comparacion(
            datos["NOMBRE_COMPLETO"]
        )
    )

    municipios_candidatos_normalizados = {
        normalizar_para_comparacion(
            municipio
        )
        for municipio in datos[
            "MUNICIPIOS_CANDIDATOS"
        ]
        if municipio
    }

    municipio_coincide = (
        municipio_normalizado
        in municipios_candidatos_normalizados
        or municipio_normalizado
        in nombre_completo_normalizado
    )

    departamento_respuesta_normalizado = (
        normalizar_para_comparacion(
            datos["DEPARTAMENTO_RESPUESTA"]
        )
    )

    departamento_coincide = (
        not departamento_respuesta_normalizado
        or departamento_normalizado
        == departamento_respuesta_normalizado
        or departamento_normalizado
        in nombre_completo_normalizado
    )

    coordenada_valida = coordenadas_validas(
        ubicacion.latitude,
        ubicacion.longitude,
    )

    validacion_correcta = bool(
        municipio_coincide
        and departamento_coincide
        and coordenada_valida
    )

    datos["VALIDACION_MUNICIPIO"] = (
        "VALIDADA"
        if validacion_correcta
        else "RECHAZADA"
    )

    return validacion_correcta, datos


def seleccionar_ubicacion_valida(
    ubicaciones: list[Location],
    municipio_esperado: str,
) -> tuple[Location | None, dict]:
    """
    Selecciona el primer resultado que corresponda al municipio.
    """

    ultimo_detalle = {
        "MUNICIPIO_RESPUESTA": "",
        "DEPARTAMENTO_RESPUESTA": "",
        "VALIDACION_MUNICIPIO": "SIN RESULTADOS",
        "NOMBRE_COMPLETO": "",
    }

    for ubicacion in ubicaciones:
        es_valida, detalle = validar_municipio_respuesta(
            ubicacion=ubicacion,
            municipio_esperado=municipio_esperado,
            departamento_esperado="Caldas",
        )

        ultimo_detalle = detalle

        if es_valida:
            return ubicacion, detalle

    return None, ultimo_detalle


# ==========================================================
# VALIDACIÓN DE ENTRADA
# ==========================================================

def validar_columnas_geocodificacion(
    dataframe: pd.DataFrame,
) -> None:
    """
    Valida las columnas requeridas para geocodificar.
    """

    columnas_requeridas = {
        "DIRECCION",
        "DESC_MUNICIPIO",
    }

    faltantes = columnas_requeridas.difference(
        dataframe.columns
    )

    if faltantes:
        detalle = "\n".join(
            f"- {columna}"
            for columna in sorted(
                faltantes
            )
        )

        raise ErrorGeocodificacion(
            "Faltan columnas para geocodificar:\n\n"
            f"{detalle}"
        )


# ==========================================================
# CONSULTA PROGRESIVA DE UNA DIRECCIÓN
# ==========================================================

def geocodificar_direccion_normalizada(
    geocodificar: Callable,
    direccion_normalizada: DireccionNormalizada,
) -> dict:
    """
    Ejecuta las consultas progresivas hasta encontrar una
    coordenada validada dentro del municipio esperado.
    """

    intentos = 0

    ultima_consulta = ""
    ultimo_detalle = {
        "MUNICIPIO_RESPUESTA": "",
        "DEPARTAMENTO_RESPUESTA": "",
        "VALIDACION_MUNICIPIO": "SIN RESULTADOS",
        "NOMBRE_COMPLETO": "",
    }

    for consulta in direccion_normalizada.consultas:

        intentos += 1
        ultima_consulta = consulta

        logger.info(
            "Intento %s | Dirección: %s",
            intentos,
            consulta,
        )

        ubicaciones = consultar_direccion(
            geocodificar=geocodificar,
            direccion_consulta=consulta,
        )

        ubicacion_valida, detalle = (
            seleccionar_ubicacion_valida(
                ubicaciones=ubicaciones,
                municipio_esperado=(
                    direccion_normalizada.municipio
                ),
            )
        )

        ultimo_detalle = detalle

        if ubicacion_valida is not None:
            return {
                "ENCONTRADA": True,
                "LATITUD": float(
                    ubicacion_valida.latitude
                ),
                "LONGITUD": float(
                    ubicacion_valida.longitude
                ),
                "DIRECCION_CONSULTADA": consulta,
                "MUNICIPIO_RESPUESTA": detalle.get(
                    "MUNICIPIO_RESPUESTA",
                    "",
                ),
                "DEPARTAMENTO_RESPUESTA": detalle.get(
                    "DEPARTAMENTO_RESPUESTA",
                    "",
                ),
                "VALIDACION_MUNICIPIO": "VALIDADA",
                "INTENTOS": intentos,
                "DETALLE": (
                    detalle.get(
                        "NOMBRE_COMPLETO",
                        "",
                    )
                ),
            }

    return {
        "ENCONTRADA": False,
        "LATITUD": pd.NA,
        "LONGITUD": pd.NA,
        "DIRECCION_CONSULTADA": ultima_consulta,
        "MUNICIPIO_RESPUESTA": ultimo_detalle.get(
            "MUNICIPIO_RESPUESTA",
            "",
        ),
        "DEPARTAMENTO_RESPUESTA": ultimo_detalle.get(
            "DEPARTAMENTO_RESPUESTA",
            "",
        ),
        "VALIDACION_MUNICIPIO": ultimo_detalle.get(
            "VALIDACION_MUNICIPIO",
            "NO VALIDADA",
        ),
        "INTENTOS": intentos,
        "DETALLE": (
            "No se encontró una coordenada que coincidiera "
            "con el municipio esperado."
        ),
    }


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def geocodificar_dataframe(
    dataframe: pd.DataFrame,
    limite_consultas_nuevas: int = 10,
) -> tuple[pd.DataFrame, dict]:
    """
    Geocodifica pedidos mediante normalización y validación.

    El límite corresponde al número de direcciones nuevas
    procesadas, no al número de intentos individuales.
    """

    validar_columnas_geocodificacion(
        dataframe
    )

    if limite_consultas_nuevas < 1:
        raise ErrorGeocodificacion(
            "El límite de consultas debe ser mayor que cero."
        )

    resultado = dataframe.copy()

    resultado["LATITUD"] = pd.NA
    resultado["LONGITUD"] = pd.NA
    resultado["ESTADO_GEOCODIFICACION"] = ""
    resultado["DIRECCION_CONSULTADA"] = ""
    resultado["VALIDACION_MUNICIPIO"] = ""

    cache = cargar_cache()

    cache_por_clave = {
        limpiar_texto(
            fila.get("CLAVE_DIRECCION")
        ): fila
        for _, fila in cache.iterrows()
        if limpiar_texto(
            fila.get("CLAVE_DIRECCION")
        )
    }

    geocodificar = crear_geocodificador()

    nuevos_cache: list[dict] = []

    direcciones_procesadas = 0
    intentos_geocodificacion = 0
    encontrados = 0
    no_encontrados = 0
    reutilizados = 0
    direcciones_vacias = 0
    no_geocodificables = 0
    omitidos_por_limite = 0
    rechazados_municipio = 0

    for indice, fila in resultado.iterrows():

        direccion = limpiar_texto(
            fila.get("DIRECCION")
        )

        municipio = limpiar_texto(
            fila.get("DESC_MUNICIPIO")
        )

        clave = construir_clave(
            direccion=direccion,
            municipio=municipio,
        )

        registro_cache = cache_por_clave.get(
            clave
        )

        if registro_cache is not None:

            latitud_cache = registro_cache.get(
                "LATITUD"
            )

            longitud_cache = registro_cache.get(
                "LONGITUD"
            )

            estado_cache = limpiar_texto(
                registro_cache.get("RESULTADO")
            )

            resultado.at[
                indice,
                "LATITUD",
            ] = latitud_cache

            resultado.at[
                indice,
                "LONGITUD",
            ] = longitud_cache

            resultado.at[
                indice,
                "ESTADO_GEOCODIFICACION",
            ] = estado_cache

            resultado.at[
                indice,
                "DIRECCION_CONSULTADA",
            ] = limpiar_texto(
                registro_cache.get(
                    "DIRECCION_CONSULTADA"
                )
            )

            resultado.at[
                indice,
                "VALIDACION_MUNICIPIO",
            ] = limpiar_texto(
                registro_cache.get(
                    "VALIDACION_MUNICIPIO"
                )
            )

            if (
                estado_cache == "ENCONTRADA"
                and coordenadas_validas(
                    latitud_cache,
                    longitud_cache,
                )
            ):
                reutilizados += 1

            else:
                no_encontrados += 1

            continue

        if direcciones_procesadas >= limite_consultas_nuevas:
            resultado.at[
                indice,
                "ESTADO_GEOCODIFICACION",
            ] = "PENDIENTE"

            omitidos_por_limite += 1
            continue

        direccion_normalizada = normalizar_direccion(
            direccion=direccion,
            municipio=municipio,
            departamento="Caldas",
        )

        if not direccion_normalizada.es_geocodificable:

            estado = (
                "DIRECCION VACIA"
                if direccion_normalizada.tipo
                == "DIRECCION_VACIA"
                else "NO GEOCODIFICABLE"
            )

            resultado.at[
                indice,
                "ESTADO_GEOCODIFICACION",
            ] = estado

            if estado == "DIRECCION VACIA":
                direcciones_vacias += 1

            else:
                no_geocodificables += 1

            nuevo_registro = {
                "CLAVE_DIRECCION": clave,
                "DIRECCION_ORIGINAL": direccion,
                "MUNICIPIO": municipio,
                "TIPO_DIRECCION": (
                    direccion_normalizada.tipo
                ),
                "DIRECCION_CONSULTADA": "",
                "LATITUD": pd.NA,
                "LONGITUD": pd.NA,
                "RESULTADO": estado,
                "MUNICIPIO_RESPUESTA": "",
                "DEPARTAMENTO_RESPUESTA": "",
                "VALIDACION_MUNICIPIO": "NO APLICA",
                "INTENTOS": 0,
                "DETALLE": direccion_normalizada.detalle,
                "FECHA_CONSULTA": datetime.now(),
            }

            nuevos_cache.append(
                nuevo_registro
            )

            cache_por_clave[
                clave
            ] = pd.Series(
                nuevo_registro
            )

            direcciones_procesadas += 1
            continue

        respuesta = geocodificar_direccion_normalizada(
            geocodificar=geocodificar,
            direccion_normalizada=direccion_normalizada,
        )

        direcciones_procesadas += 1

        intentos_geocodificacion += int(
            respuesta["INTENTOS"]
        )

        if respuesta["ENCONTRADA"]:
            estado = "ENCONTRADA"
            encontrados += 1

        else:
            estado = "NO ENCONTRADA"
            no_encontrados += 1

            if (
                respuesta["VALIDACION_MUNICIPIO"]
                == "RECHAZADA"
            ):
                rechazados_municipio += 1

        resultado.at[
            indice,
            "LATITUD",
        ] = respuesta["LATITUD"]

        resultado.at[
            indice,
            "LONGITUD",
        ] = respuesta["LONGITUD"]

        resultado.at[
            indice,
            "ESTADO_GEOCODIFICACION",
        ] = estado

        resultado.at[
            indice,
            "DIRECCION_CONSULTADA",
        ] = respuesta[
            "DIRECCION_CONSULTADA"
        ]

        resultado.at[
            indice,
            "VALIDACION_MUNICIPIO",
        ] = respuesta[
            "VALIDACION_MUNICIPIO"
        ]

        nuevo_registro = {
            "CLAVE_DIRECCION": clave,
            "DIRECCION_ORIGINAL": direccion,
            "MUNICIPIO": municipio,
            "TIPO_DIRECCION": (
                direccion_normalizada.tipo
            ),
            "DIRECCION_CONSULTADA": respuesta[
                "DIRECCION_CONSULTADA"
            ],
            "LATITUD": respuesta["LATITUD"],
            "LONGITUD": respuesta["LONGITUD"],
            "RESULTADO": estado,
            "MUNICIPIO_RESPUESTA": respuesta[
                "MUNICIPIO_RESPUESTA"
            ],
            "DEPARTAMENTO_RESPUESTA": respuesta[
                "DEPARTAMENTO_RESPUESTA"
            ],
            "VALIDACION_MUNICIPIO": respuesta[
                "VALIDACION_MUNICIPIO"
            ],
            "INTENTOS": respuesta["INTENTOS"],
            "DETALLE": respuesta["DETALLE"],
            "FECHA_CONSULTA": datetime.now(),
        }

        nuevos_cache.append(
            nuevo_registro
        )

        cache_por_clave[
            clave
        ] = pd.Series(
            nuevo_registro
        )

    if nuevos_cache:
        cache_actualizada = pd.concat(
            [
                cache,
                pd.DataFrame(
                    nuevos_cache
                ),
            ],
            ignore_index=True,
        )

        guardar_cache(
            cache_actualizada
        )

    control = {
        "CONSULTAS_REALIZADAS": direcciones_procesadas,
        "INTENTOS_GEOCODIFICACION": intentos_geocodificacion,
        "COORDENADAS_ENCONTRADAS": encontrados,
        "DIRECCIONES_NO_ENCONTRADAS": no_encontrados,
        "COORDENADAS_REUTILIZADAS": reutilizados,
        "DIRECCIONES_VACIAS": direcciones_vacias,
        "NO_GEOCODIFICABLES": no_geocodificables,
        "RECHAZADAS_POR_MUNICIPIO": rechazados_municipio,
        "PENDIENTES_POR_LIMITE": omitidos_por_limite,
    }

    logger.info(
        "Geocodificación finalizada: %s",
        control,
    )

    return resultado, control