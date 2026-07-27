import csv
import logging
from pathlib import Path

import pandas as pd

from src.config import (
    CODIFICACIONES_CSV,
    ENTRADA_DIR,
    EXTENSIONES_CSV_VALIDAS,
)


logger = logging.getLogger(__name__)


class ErrorLecturaCSV(Exception):
    """Error controlado durante la lectura del archivo CSV."""


def buscar_archivos_csv() -> list[Path]:
    """
    Busca archivos CSV dentro de la carpeta de entrada.

    Returns:
        Lista ordenada de archivos CSV encontrados.
    """

    archivos = sorted(
        archivo
        for archivo in ENTRADA_DIR.iterdir()
        if (
            archivo.is_file()
            and archivo.suffix.lower() in EXTENSIONES_CSV_VALIDAS
        )
    )

    logger.info(
        "Archivos CSV encontrados en entrada: %s",
        len(archivos),
    )

    return archivos


def seleccionar_archivo_csv() -> Path:
    """
    Selecciona el archivo CSV que será procesado.

    Returns:
        Ruta del único archivo CSV disponible.

    Raises:
        FileNotFoundError:
            Cuando no existe ningún archivo CSV.

        ErrorLecturaCSV:
            Cuando existen varios archivos CSV y no es posible
            determinar cuál debe procesarse.
    """

    archivos = buscar_archivos_csv()

    if not archivos:
        raise FileNotFoundError(
            "No se encontró ningún archivo CSV en la carpeta entrada."
        )

    if len(archivos) > 1:
        nombres = "\n".join(
            f"- {archivo.name}"
            for archivo in archivos
        )

        raise ErrorLecturaCSV(
            "Se encontraron varios archivos CSV en la carpeta entrada.\n\n"
            "Deje solamente el archivo que desea procesar:\n\n"
            f"{nombres}"
        )

    return archivos[0]


def detectar_separador(
    ruta_csv: Path,
    codificacion: str,
) -> str:
    """
    Detecta el separador utilizado por el archivo CSV.
    """

    with ruta_csv.open(
        mode="r",
        encoding=codificacion,
        errors="strict",
        newline="",
    ) as archivo:

        muestra = archivo.read(10_000)

    if not muestra.strip():
        raise ErrorLecturaCSV(
            f"El archivo {ruta_csv.name} está vacío."
        )

    try:
        dialecto = csv.Sniffer().sniff(
            muestra,
            delimiters=";,|\t",
        )

        return dialecto.delimiter

    except csv.Error:
        logger.warning(
            "No fue posible detectar automáticamente el separador. "
            "Se utilizará punto y coma."
        )

        return ";"


def leer_csv(ruta_csv: Path) -> pd.DataFrame:
    """
    Lee un archivo CSV intentando diferentes codificaciones.

    Todas las columnas se cargan inicialmente como texto para evitar:

    - Pérdida de ceros a la izquierda.
    - Notación científica en pedidos.
    - Conversiones automáticas incorrectas.
    """

    ultimo_error: Exception | None = None

    for codificacion in CODIFICACIONES_CSV:

        try:
            separador = detectar_separador(
                ruta_csv=ruta_csv,
                codificacion=codificacion,
            )

            dataframe = pd.read_csv(
                ruta_csv,
                sep=separador,
                encoding=codificacion,
                dtype=str,
                keep_default_na=False,
                na_filter=False,
                low_memory=False,
            )

            logger.info(
                "Archivo leído correctamente | Archivo: %s | "
                "Codificación: %s | Separador: %r | "
                "Filas: %s | Columnas: %s",
                ruta_csv.name,
                codificacion,
                separador,
                len(dataframe),
                len(dataframe.columns),
            )

            return dataframe

        except UnicodeDecodeError as error:
            ultimo_error = error

        except pd.errors.EmptyDataError as error:
            raise ErrorLecturaCSV(
                f"El archivo {ruta_csv.name} no contiene información."
            ) from error

        except pd.errors.ParserError as error:
            ultimo_error = error

        except OSError as error:
            raise ErrorLecturaCSV(
                f"No fue posible abrir el archivo {ruta_csv.name}."
            ) from error

    raise ErrorLecturaCSV(
        "No fue posible leer el archivo CSV con las codificaciones "
        "y separadores configurados."
    ) from ultimo_error


def cargar_archivo_entrada() -> tuple[Path, pd.DataFrame]:
    """
    Localiza y carga el archivo CSV disponible en entrada.
    """

    ruta_csv = seleccionar_archivo_csv()
    dataframe = leer_csv(ruta_csv)

    return ruta_csv, dataframe