import logging
import re
import unicodedata

from datetime import date, datetime, timedelta
from pathlib import Path

import holidays
import pandas as pd

from src.config import (
    HOJA_FESTIVOS_ADICIONALES,
    HOJA_PARAMETROS_ANS,
    HOJA_REGLAS_NEGOCIO,
    HOJA_REGLAS_PRIORIDAD,
    RUTA_DIAS_CONTRACTUALES,
)


logger = logging.getLogger(__name__)


# ==========================================================
# PARÁMETROS OBLIGATORIOS DEL ARCHIVO DE CONFIGURACIÓN
# ==========================================================

PARAMETROS_OBLIGATORIOS = {
    "DIAS_INICIO_ALERTA",
    "EXCLUIR_SABADOS",
    "EXCLUIR_DOMINGOS",
    "EXCLUIR_FESTIVOS_COLOMBIA",
    "EXCLUIR_FESTIVOS_ADICIONALES",
}


class ErrorCalculoANS(Exception):
    """
    Error controlado durante el cálculo de los indicadores ANS.
    """


# ==========================================================
# UTILIDADES DE NORMALIZACIÓN
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


def normalizar_clave(
    valor: object,
) -> str:
    """
    Normaliza nombres y parámetros para compararlos.

    Se eliminan:

    - tildes;
    - espacios;
    - guiones;
    - signos;
    - diferencias entre mayúsculas y minúsculas.

    Esto permite considerar equivalentes:

    VILLA MARÍA
    VILLAMARIA
    Villa María
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
        "",
        texto,
    )

    return texto


def convertir_si_no(
    valor: object,
    nombre_parametro: str,
) -> bool:
    """
    Convierte valores SI/NO del Excel a booleanos.
    """

    texto = normalizar_clave(
        valor
    )

    if texto in {
        "SI",
        "S",
        "TRUE",
        "VERDADERO",
        "1",
    }:
        return True

    if texto in {
        "NO",
        "N",
        "FALSE",
        "FALSO",
        "0",
    }:
        return False

    raise ErrorCalculoANS(
        f"El parámetro {nombre_parametro} debe contener "
        "SI o NO."
    )


def convertir_entero_positivo(
    valor: object,
    nombre_campo: str,
    permitir_cero: bool = False,
) -> int:
    """
    Convierte un valor numérico del Excel en entero.
    """

    try:
        numero = int(
            float(valor)
        )

    except (
        TypeError,
        ValueError,
    ) as error:

        raise ErrorCalculoANS(
            f"El valor de {nombre_campo} debe ser numérico."
        ) from error

    minimo = (
        0
        if permitir_cero
        else 1
    )

    if numero < minimo:
        raise ErrorCalculoANS(
            f"El valor de {nombre_campo} debe ser mayor "
            f"o igual a {minimo}."
        )

    return numero


# ==========================================================
# LECTURA DEL ARCHIVO DE CONFIGURACIÓN
# ==========================================================

def validar_archivo_configuracion() -> None:
    """
    Valida que exista el archivo de reglas contractuales.
    """

    if not RUTA_DIAS_CONTRACTUALES.exists():
        raise ErrorCalculoANS(
            "No se encontró el archivo de configuración:\n\n"
            f"{RUTA_DIAS_CONTRACTUALES}\n\n"
            "Ubique DIAS_CONTRACTUALES.xlsx dentro de la "
            "carpeta config."
        )


def leer_hoja_excel(
    ruta_archivo: Path,
    nombre_hoja: str,
) -> pd.DataFrame:
    """
    Lee una hoja del archivo de configuración.
    """

    try:
        return pd.read_excel(
            ruta_archivo,
            sheet_name=nombre_hoja,
            dtype=object,
            engine="openpyxl",
        )

    except PermissionError as error:
        raise ErrorCalculoANS(
            "No fue posible leer DIAS_CONTRACTUALES.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a ejecutar."
        ) from error

    except ValueError as error:
        raise ErrorCalculoANS(
            f"No se encontró la hoja {nombre_hoja} dentro de "
            "DIAS_CONTRACTUALES.xlsx."
        ) from error

    except Exception as error:
        logger.exception(
            "No fue posible leer la hoja %s.",
            nombre_hoja,
        )

        raise ErrorCalculoANS(
            f"No fue posible leer la hoja {nombre_hoja}."
        ) from error


def cargar_reglas_municipios() -> dict[str, int]:
    """
    Lee REGLAS_DE_NEGOCIO y construye:

        municipio normalizado -> días pactados
    """

    dataframe = leer_hoja_excel(
        RUTA_DIAS_CONTRACTUALES,
        HOJA_REGLAS_NEGOCIO,
    )

    columnas_requeridas = {
        "Municipio",
        "Dias",
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

        raise ErrorCalculoANS(
            "Faltan columnas en REGLAS_DE_NEGOCIO:\n\n"
            f"{detalle}"
        )

    dataframe = dataframe[
        [
            "Municipio",
            "Dias",
        ]
    ].copy()

    dataframe["CLAVE_MUNICIPIO"] = (
        dataframe["Municipio"]
        .apply(
            normalizar_clave
        )
    )

    dataframe = dataframe[
        dataframe["CLAVE_MUNICIPIO"].ne("")
    ].copy()

    if dataframe.empty:
        raise ErrorCalculoANS(
            "La hoja REGLAS_DE_NEGOCIO no contiene reglas."
        )

    duplicados = dataframe[
        dataframe["CLAVE_MUNICIPIO"].duplicated(
            keep=False
        )
    ]

    if not duplicados.empty:
        municipios = sorted(
            {
                limpiar_texto(valor)
                for valor in duplicados["Municipio"]
            }
        )

        raise ErrorCalculoANS(
            "Existen municipios duplicados en "
            "REGLAS_DE_NEGOCIO:\n\n"
            + "\n".join(
                f"- {municipio}"
                for municipio in municipios
            )
        )

    reglas: dict[str, int] = {}

    for _, fila in dataframe.iterrows():

        clave = fila["CLAVE_MUNICIPIO"]

        dias = convertir_entero_positivo(
            fila["Dias"],
            (
                "Dias del municipio "
                f"{limpiar_texto(fila['Municipio'])}"
            ),
        )

        reglas[
            clave
        ] = dias

    return reglas

def cargar_reglas_prioridad() -> list[dict]:
    """
    Lee la hoja REGLAS_PRIORIDAD y construye las reglas
    especiales basadas en el contenido de OBSERVACION.

    Cada regla contiene:

    - PALABRA_CLAVE
    - TIPO
    - DIAS
    """

    dataframe = leer_hoja_excel(
        RUTA_DIAS_CONTRACTUALES,
        HOJA_REGLAS_PRIORIDAD,
    )

    columnas_requeridas = {
        "PALABRA_CLAVE",
        "TIPO",
        "DIAS",
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

        raise ErrorCalculoANS(
            "Faltan columnas en REGLAS_PRIORIDAD:\n\n"
            f"{detalle}"
        )

    dataframe = dataframe[
        [
            "PALABRA_CLAVE",
            "TIPO",
            "DIAS",
        ]
    ].copy()

    dataframe["CLAVE_PRIORIDAD"] = (
        dataframe["PALABRA_CLAVE"]
        .apply(
            normalizar_clave
        )
    )

    dataframe["TIPO_NORMALIZADO"] = (
        dataframe["TIPO"]
        .apply(
            limpiar_texto
        )
        .str.upper()
    )

    dataframe = dataframe[
        dataframe["CLAVE_PRIORIDAD"].ne("")
    ].copy()

    if dataframe.empty:
        raise ErrorCalculoANS(
            "La hoja REGLAS_PRIORIDAD no contiene reglas."
        )

    duplicados = dataframe[
        dataframe["CLAVE_PRIORIDAD"].duplicated(
            keep=False
        )
    ]

    if not duplicados.empty:
        palabras = sorted(
            {
                limpiar_texto(valor)
                for valor in duplicados[
                    "PALABRA_CLAVE"
                ]
            }
        )

        raise ErrorCalculoANS(
            "Existen palabras clave duplicadas en "
            "REGLAS_PRIORIDAD:\n\n"
            + "\n".join(
                f"- {palabra}"
                for palabra in palabras
            )
        )

    reglas: list[dict] = []

    for _, fila in dataframe.iterrows():

        palabra_clave = limpiar_texto(
            fila["PALABRA_CLAVE"]
        )

        clave_prioridad = fila[
            "CLAVE_PRIORIDAD"
        ]

        tipo = limpiar_texto(
            fila["TIPO_NORMALIZADO"]
        )

        dias = convertir_entero_positivo(
            fila["DIAS"],
            (
                "DIAS de la regla "
                f"{palabra_clave}"
            ),
            permitir_cero=True,
        )

        reglas.append(
            {
                "PALABRA_CLAVE": palabra_clave,
                "CLAVE_PRIORIDAD": clave_prioridad,
                "TIPO": tipo,
                "DIAS": dias,
            }
        )

    return reglas
def buscar_regla_prioridad(
    observacion: object,
    reglas_prioridad: list[dict],
) -> dict | None:
    """
    Busca una palabra clave dentro de OBSERVACION.

    Retorna la regla encontrada o None.
    """

    observacion_normalizada = normalizar_clave(
        observacion
    )

    if not observacion_normalizada:
        return None

    for regla in reglas_prioridad:

        if (
            regla["CLAVE_PRIORIDAD"]
            in observacion_normalizada
        ):
            return regla

    return None


def cargar_parametros() -> dict[str, object]:
    """
    Lee la hoja PARAMETROS.

    Retorna un diccionario:

        PARAMETRO -> VALOR
    """

    dataframe = leer_hoja_excel(
        RUTA_DIAS_CONTRACTUALES,
        HOJA_PARAMETROS_ANS,
    )

    columnas_requeridas = {
        "PARAMETRO",
        "VALOR",
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

        raise ErrorCalculoANS(
            "Faltan columnas en PARAMETROS:\n\n"
            f"{detalle}"
        )

    parametros: dict[str, object] = {}

    for _, fila in dataframe.iterrows():

        clave = limpiar_texto(
            fila.get("PARAMETRO")
        ).upper()

        if not clave:
            continue

        if clave in parametros:
            raise ErrorCalculoANS(
                f"El parámetro {clave} está duplicado."
            )

        parametros[
            clave
        ] = fila.get(
            "VALOR"
        )

    faltantes_parametros = (
        PARAMETROS_OBLIGATORIOS
        .difference(
            parametros
        )
    )

    if faltantes_parametros:
        detalle = "\n".join(
            f"- {parametro}"
            for parametro in sorted(
                faltantes_parametros
            )
        )

        raise ErrorCalculoANS(
            "Faltan parámetros obligatorios:\n\n"
            f"{detalle}"
        )

    return {
        "DIAS_INICIO_ALERTA": (
            convertir_entero_positivo(
                parametros[
                    "DIAS_INICIO_ALERTA"
                ],
                "DIAS_INICIO_ALERTA",
                permitir_cero=True,
            )
        ),
        "EXCLUIR_SABADOS": convertir_si_no(
            parametros["EXCLUIR_SABADOS"],
            "EXCLUIR_SABADOS",
        ),
        "EXCLUIR_DOMINGOS": convertir_si_no(
            parametros["EXCLUIR_DOMINGOS"],
            "EXCLUIR_DOMINGOS",
        ),
        "EXCLUIR_FESTIVOS_COLOMBIA": convertir_si_no(
            parametros["EXCLUIR_FESTIVOS_COLOMBIA"],
            "EXCLUIR_FESTIVOS_COLOMBIA",
        ),
        "EXCLUIR_FESTIVOS_ADICIONALES": convertir_si_no(
            parametros["EXCLUIR_FESTIVOS_ADICIONALES"],
            "EXCLUIR_FESTIVOS_ADICIONALES",
        ),
    }


def cargar_festivos_adicionales() -> set[date]:
    """
    Lee FESTIVOS_ADICIONALES.

    La hoja puede estar vacía, pero debe contener:

        FECHA
        DESCRIPCION
        ACTIVO
    """

    dataframe = leer_hoja_excel(
        RUTA_DIAS_CONTRACTUALES,
        HOJA_FESTIVOS_ADICIONALES,
    )

    columnas_requeridas = {
        "FECHA",
        "DESCRIPCION",
        "ACTIVO",
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

        raise ErrorCalculoANS(
            "Faltan columnas en FESTIVOS_ADICIONALES:\n\n"
            f"{detalle}"
        )

    festivos: set[date] = set()

    for numero_fila, fila in dataframe.iterrows():

        fecha_valor = fila.get(
            "FECHA"
        )

        activo_valor = fila.get(
            "ACTIVO"
        )

        if (
            pd.isna(fecha_valor)
            and pd.isna(activo_valor)
        ):
            continue

        activo = convertir_si_no(
            activo_valor,
            (
                "ACTIVO de FESTIVOS_ADICIONALES "
                f"en la fila {numero_fila + 2}"
            ),
        )

        if not activo:
            continue

        fecha = pd.to_datetime(
            fecha_valor,
            errors="coerce",
        )

        if pd.isna(fecha):
            raise ErrorCalculoANS(
                "Existe una fecha inválida en "
                "FESTIVOS_ADICIONALES, fila "
                f"{numero_fila + 2}."
            )

        festivos.add(
            fecha.date()
        )

    return festivos


# ==========================================================
# CALENDARIO HÁBIL
# ==========================================================

def construir_festivos_colombia(
    fecha_minima: date,
    fecha_maxima: date,
) -> set[date]:
    """
    Construye los festivos oficiales colombianos requeridos
    para el rango de fechas del informe.
    """

    anos = range(
        fecha_minima.year,
        fecha_maxima.year + 1,
    )

    # La librería holidays genera el calendario de festivos
    # oficiales de Colombia para los años requeridos.
    #
    # "CO" corresponde al código del país Colombia.
    calendario = holidays.country_holidays(
        "CO",
        years=anos,
    )

    # Se devuelven únicamente las fechas de los festivos
    # en un conjunto set[date], para facilitar la validación:
    #
    # if fecha in festivos:
    #     return False
    return set(
        calendario.keys()
    )


def es_dia_habil(
    fecha: date,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
) -> bool:
    """
    Determina si una fecha cuenta como día contractual.
    """

    if (
        excluir_sabados
        and fecha.weekday() == 5
    ):
        return False

    if (
        excluir_domingos
        and fecha.weekday() == 6
    ):
        return False

    if fecha in festivos:
        return False

    return True


def sumar_dias_habiles(
    fecha_inicio: date,
    dias_pactados: int,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
) -> date:
    """
    Calcula FECHA_LIMITE_ANS.

    La FECHA_ORDEN no cuenta.

    El conteo comienza desde el día siguiente y solamente
    incrementa cuando la fecha es hábil.
    """

    fecha_actual = fecha_inicio
    dias_contados = 0

    while dias_contados < dias_pactados:

        fecha_actual += timedelta(
            days=1
        )

        if es_dia_habil(
            fecha_actual,
            excluir_sabados,
            excluir_domingos,
            festivos,
        ):
            dias_contados += 1

    return fecha_actual


def contar_dias_habiles(
    fecha_inicial_exclusiva: date,
    fecha_final_inclusiva: date,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
) -> int:
    """
    Cuenta días hábiles después de la fecha inicial y hasta la
    fecha final, incluyendo la fecha final cuando sea hábil.

    Ejemplo:

        fecha inicial: jueves
        fecha final: martes

        viernes = 1
        lunes   = 2
        martes  = 3
    """

    if fecha_final_inclusiva <= fecha_inicial_exclusiva:
        return 0

    contador = 0
    fecha_actual = fecha_inicial_exclusiva

    while fecha_actual < fecha_final_inclusiva:

        fecha_actual += timedelta(
            days=1
        )

        if es_dia_habil(
            fecha_actual,
            excluir_sabados,
            excluir_domingos,
            festivos,
        ):
            contador += 1

    return contador


def calcular_dias_restantes(
    fecha_actual: date,
    fecha_limite: date,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
) -> int:
    """
    Calcula días hábiles restantes.

    Reglas:

    - Si hoy es la fecha límite: 0.
    - Si la fecha límite es futura: valor positivo.
    - Si la fecha límite ya pasó: valor negativo.
    """

    if fecha_actual == fecha_limite:
        return 0

    if fecha_actual < fecha_limite:
        return contar_dias_habiles(
            fecha_inicial_exclusiva=fecha_actual,
            fecha_final_inclusiva=fecha_limite,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        )

    dias_vencidos = contar_dias_habiles(
        fecha_inicial_exclusiva=fecha_limite,
        fecha_final_inclusiva=fecha_actual,
        excluir_sabados=excluir_sabados,
        excluir_domingos=excluir_domingos,
        festivos=festivos,
    )

    return -dias_vencidos


def determinar_estado(
    dias_restantes: int,
    umbral_alerta: int,
) -> str:
    """
    Clasifica el estado contractual.
    """

    if dias_restantes < 0:
        return "VENCIDO"

    if dias_restantes <= umbral_alerta:
        return "ALERTA"

    return "A TIEMPO"


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def aplicar_calculos_ans(
    dataframe: pd.DataFrame,
    fecha_corte: date | datetime | pd.Timestamp | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Asigna el tipo de regla, los días pactados y calcula
    todos los indicadores ANS.

    Prioridad:

    1. Busca reglas especiales en OBSERVACION.
    2. Si no encuentra coincidencia, usa la regla del municipio.
    """

    validar_archivo_configuracion()

    columnas_requeridas = {
        "FECHA_ORDEN",
        "DESC_MUNICIPIO",
        "OBSERVACION",
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

        raise ErrorCalculoANS(
            "Faltan columnas para calcular el ANS:\n\n"
            f"{detalle}"
        )

    # ======================================================
    # CARGAR REGLAS Y PARÁMETROS
    # ======================================================

    reglas_municipios = cargar_reglas_municipios()

    reglas_prioridad = cargar_reglas_prioridad()

    parametros = cargar_parametros()

    # ======================================================
    # CARGAR FESTIVOS ADICIONALES
    # ======================================================

    festivos_adicionales = (
        cargar_festivos_adicionales()
        if parametros[
            "EXCLUIR_FESTIVOS_ADICIONALES"
        ]
        else set()
    )

    resultado = dataframe.copy()

    resultado["FECHA_ORDEN"] = pd.to_datetime(
        resultado["FECHA_ORDEN"],
        errors="coerce",
    ).dt.normalize()

    fechas_validas = resultado[
        "FECHA_ORDEN"
    ].dropna()

    # ======================================================
    # FECHA DE CORTE
    # ======================================================

    if fecha_corte is None:

        fecha_actual = (
            pd.Timestamp.today()
            .normalize()
            .date()
        )

    else:

        fecha_actual_convertida = pd.to_datetime(
            fecha_corte,
            errors="coerce",
        )

        if pd.isna(
            fecha_actual_convertida
        ):
            raise ErrorCalculoANS(
                "La fecha de corte indicada no es válida."
            )

        fecha_actual = (
            fecha_actual_convertida
            .normalize()
            .date()
        )

    if fechas_validas.empty:

        fecha_minima = fecha_actual

    else:

        fecha_minima = min(
            fechas_validas.min().date(),
            fecha_actual,
        )

    # ======================================================
    # RANGO DEL CALENDARIO
    # ======================================================

    dias_municipios = list(
        reglas_municipios.values()
    )

    dias_prioridad = [
        int(regla["DIAS"])
        for regla in reglas_prioridad
    ]

    dias_maximos = max(
        dias_municipios
        + dias_prioridad
    )

    fecha_maxima = max(
        (
            fechas_validas.max().date()
            if not fechas_validas.empty
            else fecha_actual
        ),
        fecha_actual,
    ) + timedelta(
        days=(dias_maximos * 4) + 370
    )

    # ======================================================
    # CALENDARIO DE FESTIVOS
    # ======================================================

    festivos: set[date] = set(
        festivos_adicionales
    )

    if parametros[
        "EXCLUIR_FESTIVOS_COLOMBIA"
    ]:
        festivos.update(
            construir_festivos_colombia(
                fecha_minima=fecha_minima,
                fecha_maxima=fecha_maxima,
            )
        )

    dias_inicio_alerta = parametros[
        "DIAS_INICIO_ALERTA"
    ]

    excluir_sabados = parametros[
        "EXCLUIR_SABADOS"
    ]

    excluir_domingos = parametros[
        "EXCLUIR_DOMINGOS"
    ]

    # ======================================================
    # ACUMULADORES
    # ======================================================

    fechas_invalidas = 0

    tipos_resultado: list[str] = []
    dias_pactados_resultado: list[object] = []
    fechas_limite_resultado: list[object] = []
    dias_transcurridos_resultado: list[object] = []
    dias_restantes_resultado: list[object] = []
    estados_resultado: list[str] = []

    # ======================================================
    # PROCESAR REGISTROS
    # ======================================================

    for _, fila in resultado.iterrows():

        fecha_orden = fila.get(
            "FECHA_ORDEN"
        )

        municipio = limpiar_texto(
            fila.get(
                "DESC_MUNICIPIO"
            )
        )

        observacion = limpiar_texto(
            fila.get(
                "OBSERVACION"
            )
        )

        # --------------------------------------------------
        # BUSCAR REGLA PRIORITARIA
        # --------------------------------------------------

        regla_prioridad = buscar_regla_prioridad(
            observacion=observacion,
            reglas_prioridad=reglas_prioridad,
        )

        if regla_prioridad is not None:

            tipo = regla_prioridad[
                "TIPO"
            ]

            dias_pactados = regla_prioridad[
                "DIAS"
            ]

        else:

            tipo = "MUNICIPIO"

            clave_municipio = normalizar_clave(
                municipio
            )

            if (
                clave_municipio
                not in reglas_municipios
            ):
                raise ErrorCalculoANS(
                    "El municipio no está configurado en "
                    "REGLAS_DE_NEGOCIO:\n\n"
                    f"- {municipio}"
                )

            dias_pactados = reglas_municipios[
                clave_municipio
            ]

        tipos_resultado.append(
            tipo
        )

        dias_pactados_resultado.append(
            dias_pactados
        )

        # --------------------------------------------------
        # VALIDAR FECHA DE ORDEN
        # --------------------------------------------------

        if pd.isna(
            fecha_orden
        ):
            fechas_invalidas += 1

            fechas_limite_resultado.append(
                pd.NaT
            )

            dias_transcurridos_resultado.append(
                pd.NA
            )

            dias_restantes_resultado.append(
                pd.NA
            )

            estados_resultado.append(
                "SIN FECHA"
            )

            continue

        fecha_orden_date = (
            pd.Timestamp(
                fecha_orden
            )
            .date()
        )

        # --------------------------------------------------
        # FECHA LÍMITE
        # --------------------------------------------------

        fecha_limite = sumar_dias_habiles(
            fecha_inicio=fecha_orden_date,
            dias_pactados=dias_pactados,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        )

        # --------------------------------------------------
        # DÍAS TRANSCURRIDOS
        # --------------------------------------------------

        dias_transcurridos = contar_dias_habiles(
            fecha_inicial_exclusiva=fecha_orden_date,
            fecha_final_inclusiva=fecha_actual,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        )

        # --------------------------------------------------
        # DÍAS RESTANTES
        # --------------------------------------------------

        dias_restantes = calcular_dias_restantes(
            fecha_actual=fecha_actual,
            fecha_limite=fecha_limite,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        )

        # --------------------------------------------------
        # ESTADO
        # --------------------------------------------------

        if tipo in {
            "HV",
            "FACTIBILIDAD",
            "INMEDIATO",
        }:

            estado = tipo

        else:

            estado = determinar_estado(
                dias_restantes=dias_restantes,
                umbral_alerta=dias_inicio_alerta,
            )

        fechas_limite_resultado.append(
            pd.Timestamp(
                fecha_limite
            )
        )

        dias_transcurridos_resultado.append(
            dias_transcurridos
        )

        dias_restantes_resultado.append(
            dias_restantes
        )

        estados_resultado.append(
            estado
        )

    # ======================================================
    # ASIGNAR RESULTADOS
    # ======================================================

    resultado["TIPO"] = tipos_resultado

    resultado["DIAS_PACTADOS"] = pd.array(
        dias_pactados_resultado,
        dtype="Int64",
    )

    resultado["FECHA_LIMITE_ANS"] = pd.to_datetime(
        fechas_limite_resultado,
        errors="coerce",
    )

    resultado["DIAS_TRANSCURRIDOS"] = pd.array(
        dias_transcurridos_resultado,
        dtype="Int64",
    )

    resultado["DIAS_RESTANTES"] = pd.array(
        dias_restantes_resultado,
        dtype="Int64",
    )

    resultado["ESTADO"] = estados_resultado

    # ======================================================
    # CONTROL
    # ======================================================

    control = {
        "FECHA_CORTE_ANS": fecha_actual.strftime(
            "%d/%m/%Y"
        ),
        "REGLAS_MUNICIPIOS_CARGADAS": len(
            reglas_municipios
        ),
        "REGLAS_PRIORIDAD_CARGADAS": len(
            reglas_prioridad
        ),
        "DIAS_INICIO_ALERTA": dias_inicio_alerta,
        "FESTIVOS_CONSIDERADOS": len(
            festivos
        ),
        "FECHAS_SIN_CALCULO_ANS": fechas_invalidas,
        "PEDIDOS_VENCIDOS": int(
            resultado["ESTADO"]
            .eq("VENCIDO")
            .sum()
        ),
        "PEDIDOS_ALERTA": int(
            resultado["ESTADO"]
            .eq("ALERTA")
            .sum()
        ),
        "PEDIDOS_A_TIEMPO": int(
            resultado["ESTADO"]
            .eq("A TIEMPO")
            .sum()
        ),
        "PEDIDOS_INMEDIATOS": int(
            resultado["ESTADO"]
            .eq("INMEDIATO")
            .sum()
        ),
        "PEDIDOS_HV": int(
            resultado["ESTADO"]
            .eq("HV")
            .sum()
        ),
        "PEDIDOS_FACTIBILIDAD": int(
            resultado["ESTADO"]
            .eq("FACTIBILIDAD")
            .sum()
        ),
    }

    logger.info(
        "Cálculos ANS finalizados | %s",
        control,
    )

    return resultado, control