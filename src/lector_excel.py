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


VALOR_PDA_PENDIENTE = "PENDIENTE POR PROGRAMAR"


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


def normalizar_id_orden(
    serie: pd.Series,
) -> pd.Series:
    """
    Estandariza ID_ORDEN como texto.
    """

    return (
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


def normalizar_pda_numero(
    serie: pd.Series,
) -> pd.Series:
    """
    Estandariza PDA_NUMERO como texto.
    """

    return (
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


def buscar_archivos_excel() -> list[Path]:
    """
    Busca los cuatro exportes Excel disponibles en entrada.
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


def es_archivo_pda(
    ruta_archivo: Path,
) -> bool:
    """
    Indica si el archivo corresponde a pedidos programados
    y contiene la columna PDA_NUMERO.
    """

    nombre = normalizar_texto(
        ruta_archivo.stem
    )

    return "PDA" in nombre


def validar_cantidad_archivos(
    archivos: list[Path],
) -> None:
    """
    Valida que existan exactamente cuatro archivos:

    - Región 1 pendientes por programar.
    - Región 2 pendientes por programar.
    - Región 1 programados con PDA.
    - Región 2 programados con PDA.
    """

    if not archivos:
        raise FileNotFoundError(
            "No se encontraron archivos Excel "
            "en la carpeta entrada."
        )

    if len(archivos) != 4:

        nombres = "\n".join(
            f"- {archivo.name}"
            for archivo in archivos
        )

        raise ErrorLecturaExcel(
            "Se requieren exactamente cuatro archivos Excel:\n\n"
            "- Región 1 pendientes por programar\n"
            "- Región 2 pendientes por programar\n"
            "- Región 1 programados con PDA\n"
            "- Región 2 programados con PDA\n\n"
            "Archivos encontrados:\n"
            f"{nombres or '- Ninguno'}"
        )

    archivos_pda = [
        archivo
        for archivo in archivos
        if es_archivo_pda(
            archivo
        )
    ]

    archivos_pendientes = [
        archivo
        for archivo in archivos
        if not es_archivo_pda(
            archivo
        )
    ]

    if len(archivos_pda) != 2:
        raise ErrorLecturaExcel(
            "Se requieren exactamente dos archivos cuyo nombre "
            "contenga PDA:\n\n"
            "- PDA Región 1\n"
            "- PDA Región 2"
        )

    if len(archivos_pendientes) != 2:
        raise ErrorLecturaExcel(
            "Se requieren exactamente dos archivos de pedidos "
            "pendientes por programar:\n\n"
            "- Región 1\n"
            "- Región 2"
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

    if re.search(
        r"(REGION_?1|R1)",
        nombre,
    ):
        return "REGION 1"

    if re.search(
        r"(REGION_?2|R2)",
        nombre,
    ):
        return "REGION 2"

    return f"REGION {indice}"


def validar_combinaciones_archivos(
    archivos: list[Path],
) -> None:
    """
    Valida que exista una combinación para cada región:

    - Región 1 pendiente.
    - Región 2 pendiente.
    - Región 1 PDA.
    - Región 2 PDA.
    """

    combinaciones: list[tuple[str, str]] = []

    for indice, archivo in enumerate(
        archivos,
        start=1,
    ):

        region = detectar_region_origen(
            archivo,
            indice,
        )

        tipo = (
            "PDA"
            if es_archivo_pda(
                archivo
            )
            else "PENDIENTE"
        )

        combinaciones.append(
            (
                region,
                tipo,
            )
        )

    esperadas = {
        ("REGION 1", "PENDIENTE"),
        ("REGION 2", "PENDIENTE"),
        ("REGION 1", "PDA"),
        ("REGION 2", "PDA"),
    }

    encontradas = set(
        combinaciones
    )

    if encontradas != esperadas:

        detalle = "\n".join(
            f"- {archivo.name}: {region} / {tipo}"
            for archivo, (
                region,
                tipo,
            ) in zip(
                archivos,
                combinaciones,
            )
        )

        raise ErrorLecturaExcel(
            "No fue posible identificar correctamente los cuatro "
            "exportes requeridos:\n\n"
            f"{detalle}\n\n"
            "Los nombres deben permitir identificar R1, R2 y PDA."
        )


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
            normalizar_texto(
                valor
            )
            for valor in fila.tolist()
            if pd.notna(
                valor
            )
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

            return int(
                indice
            )

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

    for nombre_hoja in libro.sheet_names:

        try:
            fila_encabezados = detectar_fila_encabezados(
                ruta_archivo,
                nombre_hoja,
            )

            return nombre_hoja, fila_encabezados

        except ErrorLecturaExcel:
            continue

    raise ErrorLecturaExcel(
        f"No se encontró una hoja válida en "
        f"{ruta_archivo.name}."
    )


def normalizar_encabezados(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Estandariza internamente los nombres de columnas.
    """

    resultado = dataframe.copy()

    resultado.columns = [
        normalizar_texto(
            columna
        )
        for columna in resultado.columns
    ]

    columnas_sin_nombre = [
        columna
        for columna in resultado.columns
        if (
            not columna
            or columna.startswith(
                "UNNAMED"
            )
        )
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

    filas_antes = len(
        resultado
    )

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
    Lee y limpia uno de los cuatro exportes.

    Los archivos PDA conservan su PDA_NUMERO.
    Los archivos pendientes reciben el valor
    PENDIENTE POR PROGRAMAR.
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

    registros_antes_id = len(
        dataframe
    )

    if "ID_ORDEN" not in dataframe.columns:
        raise ErrorLecturaExcel(
            f"El archivo {ruta_archivo.name} no contiene "
            "la columna ID_ORDEN."
        )

    dataframe["ID_ORDEN"] = normalizar_id_orden(
        dataframe["ID_ORDEN"]
    )

    dataframe = dataframe[
        dataframe["ID_ORDEN"].ne("")
    ].copy()

    filas_sin_id = (
        registros_antes_id
        - len(dataframe)
    )

    dataframe["REGION_ORIGEN"] = (
        region_origen
    )

    archivo_pda = es_archivo_pda(
        ruta_archivo
    )

    if archivo_pda:

        if "PDA_NUMERO" not in dataframe.columns:
            raise ErrorLecturaExcel(
                f"El archivo {ruta_archivo.name} corresponde a PDA, "
                "pero no contiene la columna PDA_NUMERO."
            )

        dataframe["PDA_NUMERO"] = normalizar_pda_numero(
            dataframe["PDA_NUMERO"]
        )

        dataframe["PDA_NUMERO"] = (
            dataframe["PDA_NUMERO"]
            .replace(
                "",
                pd.NA,
            )
            .fillna(
                VALOR_PDA_PENDIENTE
            )
        )

        tipo_archivo = "PROGRAMADO"

    else:

        dataframe["PDA_NUMERO"] = (
            VALOR_PDA_PENDIENTE
        )

        tipo_archivo = "PENDIENTE"

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
        "TIPO_ARCHIVO": tipo_archivo,
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


def validar_ordenes_duplicadas(
    consolidado: pd.DataFrame,
) -> None:
    """
    Valida que una misma orden no aparezca repetida
    entre los cuatro archivos.

    La validación evita duplicar pedidos en el informe final.
    """

    duplicados = consolidado[
        consolidado["ID_ORDEN"].duplicated(
            keep=False
        )
    ]

    if duplicados.empty:
        return

    ordenes = sorted(
        set(
            duplicados["ID_ORDEN"]
            .astype(str)
            .tolist()
        )
    )

    detalle = "\n".join(
        f"- {orden}"
        for orden in ordenes[:30]
    )

    complemento = ""

    if len(ordenes) > 30:
        complemento = (
            f"\n- ... y {len(ordenes) - 30} órdenes adicionales"
        )

    raise ErrorLecturaExcel(
        "Se encontraron ID_ORDEN repetidos entre los cuatro "
        "archivos:\n\n"
        f"{detalle}"
        f"{complemento}\n\n"
        "Revise que un pedido no esté incluido al mismo tiempo "
        "como pendiente y como programado."
    )


def cargar_regiones() -> tuple[pd.DataFrame, list[dict]]:
    """
    Lee y unifica los cuatro archivos completos.

    - Región 1 pendiente.
    - Región 2 pendiente.
    - Región 1 con PDA.
    - Región 2 con PDA.

    Únicamente agrega o completa PDA_NUMERO.
    No modifica las reglas de negocio ni los cálculos ANS.
    """

    archivos = buscar_archivos_excel()

    validar_cantidad_archivos(
        archivos
    )

    validar_combinaciones_archivos(
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
            ruta_archivo=archivo,
            region_origen=region,
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

    validar_ordenes_duplicadas(
        consolidado
    )

    programados = int(
        consolidado["PDA_NUMERO"]
        .ne(
            VALOR_PDA_PENDIENTE
        )
        .sum()
    )

    pendientes = int(
        consolidado["PDA_NUMERO"]
        .eq(
            VALOR_PDA_PENDIENTE
        )
        .sum()
    )

    logger.info(
        "Cuatro archivos unificados | Registros: %s | "
        "Programados: %s | Pendientes por programar: %s",
        len(consolidado),
        programados,
        pendientes,
    )

    return consolidado, controles
