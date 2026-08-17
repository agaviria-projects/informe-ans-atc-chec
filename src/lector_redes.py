from __future__ import annotations

from pathlib import Path
from typing import Tuple

import pandas as pd

from src.config import BASE_DIR


# ============================================================
# RUTAS DEL PROYECTO
# ============================================================

CARPETA_ENTRADA_REDES = BASE_DIR / "entrada_redes"

EXTENSIONES_PERMITIDAS = {".xlsx", ".xlsm"}

# Encabezados mínimos que permiten identificar la fila real
# de títulos dentro del exporte CHEC.
ENCABEZADOS_CLAVE = {
    "NUMERO_PROCESO",
    "PROCESO",
    "REV_ESTADO",
    "REV_RESPONSABLE",
}


# ============================================================
# NORMALIZACIÓN
# ============================================================

def _normalizar_encabezado(valor: object) -> str:
    """
    Normaliza un encabezado para facilitar su comparación.
    No modifica los datos del archivo, solo los nombres de columnas.
    """
    if valor is None:
        return ""

    texto = str(valor).strip()
    texto = " ".join(texto.split())

    return texto.upper()


# ============================================================
# LOCALIZAR ARCHIVO
# ============================================================

def localizar_archivo_entrada_redes() -> Path:
    """
    Busca automáticamente el archivo Excel ubicado en entrada_redes.

    Reglas:
    - El nombre del archivo NO importa.
    - Se ignoran archivos temporales de Excel (~$...).
    - Debe existir exactamente un archivo Excel válido.
    """

    if not CARPETA_ENTRADA_REDES.exists():
        raise FileNotFoundError(
            f"No existe la carpeta de entrada de Redes:\n"
            f"{CARPETA_ENTRADA_REDES}"
        )

    archivos = sorted(
        archivo
        for archivo in CARPETA_ENTRADA_REDES.iterdir()
        if (
            archivo.is_file()
            and archivo.suffix.lower() in EXTENSIONES_PERMITIDAS
            and not archivo.name.startswith("~$")
        )
    )

    if not archivos:
        raise FileNotFoundError(
            "No se encontró ningún archivo Excel en 'entrada_redes'.\n"
            "Coloque un archivo .xlsx o .xlsm y vuelva a ejecutar el proceso."
        )

    if len(archivos) > 1:
        nombres = "\n".join(f" - {archivo.name}" for archivo in archivos)
        raise RuntimeError(
            "Se encontró más de un archivo Excel en 'entrada_redes'.\n"
            "Debe dejar únicamente el exporte que desea procesar.\n\n"
            f"Archivos encontrados:\n{nombres}"
        )

    return archivos[0]


# ============================================================
# DETECTAR FILA DE ENCABEZADOS
# ============================================================

def detectar_fila_encabezados(
    ruta_archivo: Path,
    max_filas_busqueda: int = 30,
) -> int:
    """
    Detecta automáticamente la fila donde comienzan los encabezados reales
    del exporte CHEC.

    Retorna la fila usando numeración de Excel:
        1, 2, 3, ...

    Ejemplo del exporte actual:
        encabezados en fila 6
    """

    vista_previa = pd.read_excel(
        ruta_archivo,
        header=None,
        nrows=max_filas_busqueda,
        engine="openpyxl",
        dtype=object,
    )

    for indice, fila in vista_previa.iterrows():
        valores = {
            _normalizar_encabezado(valor)
            for valor in fila.tolist()
            if pd.notna(valor)
        }

        if ENCABEZADOS_CLAVE.issubset(valores):
            return indice + 1

    raise ValueError(
        "No fue posible identificar la fila de encabezados del exporte de Redes.\n"
        "Se esperaban, como mínimo, estas columnas:\n"
        + "\n".join(f" - {columna}" for columna in sorted(ENCABEZADOS_CLAVE))
    )


# ============================================================
# LEER EXPORTE REDES
# ============================================================

def leer_exporte_redes() -> Tuple[pd.DataFrame, Path]:
    """
    Localiza y lee el exporte de Redes.

    Retorna:
        (dataframe, ruta_archivo)

    Esta función:
    - no depende del nombre del archivo;
    - detecta automáticamente la fila de encabezados;
    - elimina filas completamente vacías;
    - limpia espacios/saltos de línea en nombres de columnas;
    - NO aplica todavía filtros de negocio.
    """

    ruta_archivo = localizar_archivo_entrada_redes()
    fila_encabezados = detectar_fila_encabezados(ruta_archivo)

    df = pd.read_excel(
        ruta_archivo,
        header=fila_encabezados - 1,
        engine="openpyxl",
        dtype=object,
    )

    # Normalizar únicamente nombres de columnas.
    df.columns = [
        _normalizar_encabezado(columna)
        for columna in df.columns
    ]

    # Eliminar filas completamente vacías.
    df = df.dropna(how="all").reset_index(drop=True)

    # Eliminar columnas sin encabezado real si llegaran a existir.
    columnas_validas = [
        columna
        for columna in df.columns
        if columna and not columna.startswith("UNNAMED:")
    ]
    df = df[columnas_validas]

    return df, ruta_archivo


# ============================================================
# PRUEBA MANUAL
# ============================================================

if __name__ == "__main__":
    try:
        datos, archivo = leer_exporte_redes()

        print("=" * 70)
        print("LECTURA ANS REDES CORRECTA")
        print("=" * 70)
        print(f"Archivo detectado : {archivo.name}")
        print(f"Ruta              : {archivo}")
        print(f"Registros leídos  : {len(datos):,}")
        print(f"Columnas leídas   : {len(datos.columns)}")
        print()
        print("Columnas detectadas:")
        for columna in datos.columns:
            print(f" - {columna}")

    except Exception as error:
        print("=" * 70)
        print("ERROR AL LEER EL EXPORTE DE ANS REDES")
        print("=" * 70)
        print(error)
