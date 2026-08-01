import logging

import pandas as pd


logger = logging.getLogger(__name__)


# ==========================================================
# COLUMNAS DEL ARCHIVO DE ORIGEN
# ==========================================================

COLUMNAS_ORIGEN_REQUERIDAS = [
    "ID_ORDEN",
    "FECHA_ORDEN",
    "DIRECCION",
    "PROPIETARIO",
    "ZONA",
    "MUNICIPIO",
    "DESC_MUNICIPIO",
    "REGION_ORIGEN",
    "OBSERVACION",
]


# ==========================================================
# COLUMNAS CREADAS POR EL SISTEMA
# ==========================================================

COLUMNAS_CALCULADAS = [
    "TIPO",
    "DIAS_PACTADOS",
    "FECHA_LIMITE_ANS",
    "DIAS_TRANSCURRIDOS",
    "DIAS_RESTANTES",
    "ESTADO",
]


# ==========================================================
# ORDEN FINAL DEL INFORME
# ==========================================================

ORDEN_COLUMNAS_SALIDA = [
    "ID_ORDEN",
    "FECHA_ORDEN",
    "DIRECCION",
    "PROPIETARIO",
    "ZONA",
    "MUNICIPIO",
    "DESC_MUNICIPIO",
    "REGION_ORIGEN",
    "TIPO",
    "DIAS_PACTADOS",
    "FECHA_LIMITE_ANS",
    "DIAS_TRANSCURRIDOS",
    "DIAS_RESTANTES",
    "ESTADO",
    "OBSERVACION",
]


class ErrorTransformacion(Exception):
    """
    Error controlado durante la transformación del informe.
    """


def validar_columnas_requeridas(
    dataframe: pd.DataFrame,
) -> None:
    """
    Valida que el consolidado contenga todas las columnas
    necesarias para construir el informe.
    """

    columnas_faltantes = [
        columna
        for columna in COLUMNAS_ORIGEN_REQUERIDAS
        if columna not in dataframe.columns
    ]

    if columnas_faltantes:
        detalle = "\n".join(
            f"- {columna}"
            for columna in columnas_faltantes
        )

        raise ErrorTransformacion(
            "Faltan columnas requeridas en los exportes:\n\n"
            f"{detalle}"
        )


def limpiar_identificador(
    serie: pd.Series,
) -> pd.Series:
    """
    Convierte identificadores a texto y elimina decimales
    artificiales generados por Excel, por ejemplo 12345.0.
    """

    resultado = (
        serie
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True,
        )
    )

    resultado = resultado.replace(
        {
            "nan": "",
            "NaN": "",
            "None": "",
            "<NA>": "",
        }
    )

    return resultado


def normalizar_zona(
    serie: pd.Series,
) -> pd.Series:
    """
    Normaliza la zona conservando las convenciones:

    U -> Urbano
    R -> Rural
    """

    resultado = (
        serie
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    equivalencias = {
        "URBANO": "U",
        "RURAL": "R",
    }

    return resultado.replace(
        equivalencias
    )


def normalizar_fechas(
    serie: pd.Series,
) -> pd.Series:
    """
    Convierte FECHA_ORDEN a fecha y elimina cualquier componente
    de hora que pueda venir en el archivo.
    """

    return pd.to_datetime(
        serie,
        errors="coerce",
    ).dt.normalize()


def limpiar_columnas_texto(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Limpia espacios en las columnas de texto.
    """

    resultado = dataframe.copy()

    columnas_texto = [
        "DIRECCION",
        "PROPIETARIO",
        "DESC_MUNICIPIO",
        "REGION_ORIGEN",
        "OBSERVACION",
    ]

    for columna in columnas_texto:
        resultado[columna] = (
            resultado[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    return resultado


def preparar_columnas_sistema(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Agrega las columnas que posteriormente serán calculadas
    mediante la tabla de días contractuales.
    """

    resultado = dataframe.copy()

    resultado["TIPO"] = "MUNICIPIO"
    resultado["DIAS_PACTADOS"] = pd.NA
    resultado["FECHA_LIMITE_ANS"] = pd.NaT
    resultado["DIAS_TRANSCURRIDOS"] = pd.NA
    resultado["DIAS_RESTANTES"] = pd.NA

    resultado["ESTADO"] = (
        "PENDIENTE CONFIGURACIÓN"
    )

    return resultado


def detectar_duplicados(
    dataframe: pd.DataFrame,
) -> int:
    """
    Cuenta los registros cuyo ID_ORDEN está repetido.

    Los duplicados se reportan, pero no se eliminan.
    """

    return int(
        dataframe["ID_ORDEN"]
        .duplicated(
            keep=False
        )
        .sum()
    )


def transformar_consolidado(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, dict]:
    """
    Ejecuta la transformación final del consolidado regional.

    Procesos realizados:

    - Valida columnas requeridas.
    - Conserva únicamente las columnas necesarias.
    - Limpia identificadores.
    - Normaliza zona.
    - Convierte FECHA_ORDEN.
    - Agrega columnas del sistema.
    - Ordena las columnas.
    - Genera métricas de control.
    """

    validar_columnas_requeridas(
        dataframe
    )

    resultado = dataframe[
        COLUMNAS_ORIGEN_REQUERIDAS
    ].copy()

    resultado["ID_ORDEN"] = limpiar_identificador(
        resultado["ID_ORDEN"]
    )

    resultado["MUNICIPIO"] = limpiar_identificador(
        resultado["MUNICIPIO"]
    )

    resultado["ZONA"] = normalizar_zona(
        resultado["ZONA"]
    )

    resultado["FECHA_ORDEN"] = normalizar_fechas(
        resultado["FECHA_ORDEN"]
    )

    resultado = limpiar_columnas_texto(
        resultado
    )

    resultado = preparar_columnas_sistema(
        resultado
    )

    resultado = resultado[
        ORDEN_COLUMNAS_SALIDA
    ].copy()

    resultado = resultado.reset_index(
        drop=True
    )

    control = {
        "REGISTROS_CONSOLIDADOS": len(resultado),
        "ID_ORDEN_DUPLICADOS": detectar_duplicados(
            resultado
        ),
        "FECHAS_INVALIDAS": int(
            resultado["FECHA_ORDEN"]
            .isna()
            .sum()
        ),
        "DIRECCIONES_VACIAS": int(
            resultado["DIRECCION"]
            .fillna("")
            .astype(str)
            .str.strip()
            .eq("")
            .sum()
        ),
    }

    logger.info(
        "Transformación finalizada | %s",
        control,
    )

    return resultado, control