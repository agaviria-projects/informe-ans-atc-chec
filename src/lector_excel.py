import logging
import re
import unicodedata
from pathlib import Path

import pandas as pd

from src.config import ENTRADA_DIR


logger = logging.getLogger(__name__)


EXTENSIONES_EXCEL = {
    ".xlsx",
    ".xlsm",
}


COLUMNAS_REFERENCIA_ENCABEZADO = {
    "ID_ORDEN",
    "FECHA_ORDEN",
    "DIRECCION",
}


class ErrorLecturaExcel(Exception):
    """Error controlado durante la lectura de los exportes."""


def normalizar_texto(valor: object) -> str:
    """
    Normaliza un texto para realizar comparaciones internas.
    """

    if valor is None:
        return ""

    texto = str(valor).strip().upper()

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
        r"[\s\-]+",
        "_",
        texto,
    )

    texto = re.sub(
        r"_+",
        "_",
        texto,
    )

    return texto.strip("_")


def buscar_archivos_excel() -> list[Path]:
    """
    Busca los exportes Excel disponibles en entrada.
    """

    archivos = sorted(
        archivo
        for archivo in ENTRADA_DIR.iterdir()
        if (
            archivo.is_file()
            and archivo.suffix.lower() in EXTENSIONES_EXCEL
            and not archivo.name.startswith("~$")
        )
    )

    logger.info(
        "Archivos Excel encontrados: %s",
        len(archivos),
    )

    return archivos


def validar_cantidad_archivos(
    archivos: list[Path],
) -> None:
    """
    Valida que existan los dos archivos regionales.
    """

    if not archivos:
        raise FileNotFoundError(
            "No se encontraron archivos Excel en la carpeta entrada."
        )

    if len(archivos) < 2:
        raise ErrorLecturaExcel(
            "Se requiere un archivo para Región 1 y otro para Región 2."
        )

    if len(archivos) > 2:
        nombres = "\n".join(
            f"- {archivo.name}"
            for archivo in archivos
        )

        raise ErrorLecturaExcel(
            "Se encontraron más de dos archivos Excel.\n\n"
            "Deje únicamente los exportes de Región 1 y Región 2:\n\n"
            f"{nombres}"
        )


def detectar_region_origen(
    ruta_archivo: Path,
    indice: int,
) -> str:
    """
    Determina la región desde el nombre del archivo.
    """

    nombre = normalizar_texto(
        ruta_archivo.stem
    )

    if re.search(r"(REGION_?1|R1)", nombre):
        return "REGION 1"

    if re.search(r"(REGION_?2|R2)", nombre):
        return "REGION 2"

    return f"REGION {indice}"


def detectar_fila_encabezados(
    ruta_archivo: Path,
    nombre_hoja: str,
    filas_busqueda: int = 30,
) -> int:
    """
    Busca automáticamente la fila que contiene los encabezados.

    Returns:
        Índice de fila compatible con pandas, iniciando en cero.
    """

    vista_previa = pd.read_excel(
        ruta_archivo,
        sheet_name=nombre_hoja,
        header=None,
        nrows=filas_busqueda,
        dtype=object,
        engine="openpyxl",
    )

    for indice, fila in vista_previa.iterrows():

        valores = {
            normalizar_texto(valor)
            for valor in fila.tolist()
            if pd.notna(valor)
        }

        if COLUMNAS_REFERENCIA_ENCABEZADO.issubset(
            valores
        ):
            logger.info(
                "Encabezados detectados | Archivo: %s | "
                "Hoja: %s | Fila Excel: %s",
                ruta_archivo.name,
                nombre_hoja,
                indice + 1,
            )

            return int(indice)

    raise ErrorLecturaExcel(
        f"No se encontró la fila de encabezados en "
        f"{ruta_archivo.name}."
    )


def seleccionar_hoja_datos(
    ruta_archivo: Path,
) -> tuple[str, int]:
    """
    Busca la hoja que contiene la tabla de órdenes.
    """

    libro = pd.ExcelFile(
        ruta_archivo,
        engine="openpyxl",
    )

    errores: list[str] = []

    for nombre_hoja in libro.sheet_names:
        try:
            fila_encabezados = detectar_fila_encabezados(
                ruta_archivo,
                nombre_hoja,
            )

            return nombre_hoja, fila_encabezados

        except ErrorLecturaExcel as error:
            errores.append(
                str(error)
            )

    raise ErrorLecturaExcel(
        f"No se encontró una hoja válida en {ruta_archivo.name}."
    )


def normalizar_encabezados(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Estandariza internamente los nombres de columnas.
    """

    resultado = dataframe.copy()

    resultado.columns = [
        normalizar_texto(columna)
        for columna in resultado.columns
    ]

    columnas_sin_nombre = [
        columna
        for columna in resultado.columns
        if not columna
        or columna.startswith("UNNAMED")
    ]

    if columnas_sin_nombre:
        resultado = resultado.drop(
            columns=columnas_sin_nombre,
            errors="ignore",
        )

    return resultado


def limpiar_filas_vacias(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """
    Elimina filas completamente vacías o compuestas por espacios.
    """

    resultado = dataframe.copy()

    filas_antes = len(resultado)

    resultado = resultado.replace(
        r"^\s*$",
        pd.NA,
        regex=True,
    )

    resultado = resultado.dropna(
        how="all"
    )

    resultado = resultado.reset_index(
        drop=True
    )

    filas_eliminadas = (
        filas_antes
        - len(resultado)
    )

    return resultado, filas_eliminadas


def leer_archivo_region(
    ruta_archivo: Path,
    region_origen: str,
) -> tuple[pd.DataFrame, dict]:
    """
    Lee y limpia un archivo regional.
    """

    nombre_hoja, fila_encabezados = seleccionar_hoja_datos(
        ruta_archivo
    )

    dataframe = pd.read_excel(
        ruta_archivo,
        sheet_name=nombre_hoja,
        header=fila_encabezados,
        dtype=object,
        engine="openpyxl",
    )

    dataframe = normalizar_encabezados(
        dataframe
    )

    dataframe, filas_vacias = limpiar_filas_vacias(
        dataframe
    )

    registros_antes_id = len(dataframe)

    if "ID_ORDEN" not in dataframe.columns:
        raise ErrorLecturaExcel(
            f"El archivo {ruta_archivo.name} no contiene "
            "la columna ID_ORDEN."
        )

    dataframe = dataframe[
        dataframe["ID_ORDEN"].notna()
    ].copy()

    dataframe = dataframe[
        dataframe["ID_ORDEN"]
        .astype(str)
        .str.strip()
        .ne("")
    ].copy()

    filas_sin_id = (
        registros_antes_id
        - len(dataframe)
    )

    dataframe["REGION_ORIGEN"] = (
        region_origen
    )

    # ==========================================================
    # NORMALIZAR NOMBRES ESPECÍFICOS DE MUNICIPIOS
    # ==========================================================

    if "DESC_MUNICIPIO" in dataframe.columns:

        dataframe["DESC_MUNICIPIO"] = (
            dataframe["DESC_MUNICIPIO"]
            .replace({
                "SANTA ROSA DE CABAL": "Santa Rosa de Cabal",
            })
        )
        
    control = {
        "ARCHIVO": ruta_archivo.name,
        "REGION": region_origen,
        "HOJA": nombre_hoja,
        "FILA_ENCABEZADOS": fila_encabezados + 1,
        "FILAS_VACIAS_ELIMINADAS": filas_vacias,
        "FILAS_SIN_ID_ORDEN": filas_sin_id,
        "REGISTROS_VALIDOS": len(dataframe),
    }

    logger.info(
        "Archivo regional procesado | %s",
        control,
    )

    return dataframe, control


def cargar_regiones() -> tuple[pd.DataFrame, list[dict]]:
    """
    Lee Región 1 y Región 2 y las unifica.
    """

    archivos = buscar_archivos_excel()

    validar_cantidad_archivos(
        archivos
    )

    dataframes: list[pd.DataFrame] = []
    controles: list[dict] = []

    for indice, archivo in enumerate(
        archivos,
        start=1,
    ):
        region = detectar_region_origen(
            archivo,
            indice,
        )

        dataframe, control = leer_archivo_region(
            archivo,
            region,
        )

        dataframes.append(
            dataframe
        )

        controles.append(
            control
        )

    consolidado = pd.concat(
        dataframes,
        ignore_index=True,
        sort=False,
    )

    logger.info(
        "Regiones unificadas | Registros consolidados: %s",
        len(consolidado),
    )

    return consolidado, controles