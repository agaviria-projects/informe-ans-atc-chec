import logging
from dataclasses import dataclass, field

import pandas as pd


logger = logging.getLogger(__name__)


@dataclass
class ResultadoValidacion:
    """
    Resultado estructurado de la validación del CSV.
    """

    es_valido: bool = True
    errores: list[str] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)


def validar_dataframe(
    dataframe: pd.DataFrame,
) -> ResultadoValidacion:
    """
    Ejecuta validaciones estructurales iniciales.

    Todavía no valida columnas específicas de ANS porque sus nombres
    reales serán definidos después de analizar el CSV de ATC CHEC.
    """

    resultado = ResultadoValidacion()

    if dataframe.empty:
        resultado.es_valido = False
        resultado.errores.append(
            "El archivo CSV no contiene registros."
        )

        return resultado

    if len(dataframe.columns) == 0:
        resultado.es_valido = False
        resultado.errores.append(
            "El archivo CSV no contiene encabezados."
        )

        return resultado

    columnas = [
        str(columna).strip()
        for columna in dataframe.columns
    ]

    columnas_vacias = [
        columna
        for columna in columnas
        if not columna
        or columna.upper().startswith("UNNAMED")
    ]

    if columnas_vacias:
        resultado.advertencias.append(
            "Se encontraron columnas sin nombre o columnas tipo UNNAMED."
        )

    columnas_duplicadas = (
        pd.Index(columnas)
        .duplicated()
    )

    if columnas_duplicadas.any():
        duplicadas = sorted(
            {
                columnas[indice]
                for indice, duplicada
                in enumerate(columnas_duplicadas)
                if duplicada
            }
        )

        resultado.es_valido = False
        resultado.errores.append(
            "Existen nombres de columnas duplicados: "
            + ", ".join(duplicadas)
        )

    filas_duplicadas = int(
        dataframe.duplicated().sum()
    )

    if filas_duplicadas > 0:
        resultado.advertencias.append(
            f"Se detectaron {filas_duplicadas} filas completamente duplicadas."
        )

    columnas_totalmente_vacias = [
        columna
        for columna in dataframe.columns
        if (
            dataframe[columna]
            .astype(str)
            .str.strip()
            .eq("")
            .all()
        )
    ]

    if columnas_totalmente_vacias:
        resultado.advertencias.append(
            "Existen columnas completamente vacías: "
            + ", ".join(
                map(str, columnas_totalmente_vacias)
            )
        )

    logger.info(
        "Validación finalizada | Válido: %s | "
        "Errores: %s | Advertencias: %s",
        resultado.es_valido,
        len(resultado.errores),
        len(resultado.advertencias),
    )

    return resultado