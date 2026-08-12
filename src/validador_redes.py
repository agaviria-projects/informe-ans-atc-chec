from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterable

import pandas as pd


# ============================================================
# COMPATIBILIDAD DE IMPORTACIÓN
# ============================================================

DIRECTORIO_SRC = Path(__file__).resolve().parent
if str(DIRECTORIO_SRC) not in sys.path:
    sys.path.insert(0, str(DIRECTORIO_SRC))

from lector_redes import leer_exporte_redes


# ============================================================
# COLUMNAS OBLIGATORIAS DEL INFORME ANS REDES
# ============================================================

COLUMNAS_REQUERIDAS_REDES = [
    "NUMERO_PROCESO",
    "PROCESO",
    "D_PROCESO",
    "PRO_CLASIFICACION",
    "PRO_D_CLASIFICACION",
    "PRO_FECHA_SISTEMA_CREACION",
    "PRO_FECHA_VENCIMIENTO",
    "CLIENTE_ID",
    "CLI_NOMBRE",
    "CLI_DIRECCION",
    "CLI_D_CODIGO_AREA",
    "CLI_D_MUNICIPIO",
    "CLI_D_BARRIO",
    "REV_D_TIPO",
    "REV_ESTADO",
    "REV_RESPONSABLE",
    "REV_COMENTARIO",
    "CODIGO_UBIC_TRANSFORMADOR",
]


# ============================================================
# FUNCIONES DE VALIDACIÓN
# ============================================================

def validar_columnas_requeridas(
    df: pd.DataFrame,
    columnas_requeridas: Iterable[str] = COLUMNAS_REQUERIDAS_REDES,
) -> None:
    """
    Valida que el DataFrame tenga todas las columnas obligatorias
    para construir el informe ANS REDES.

    Si falta al menos una columna, detiene el proceso con un mensaje
    claro indicando cuáles columnas no fueron encontradas.
    """

    columnas_disponibles = {str(col).strip().upper() for col in df.columns}
    requeridas = [str(col).strip().upper() for col in columnas_requeridas]

    faltantes = [
        columna
        for columna in requeridas
        if columna not in columnas_disponibles
    ]

    if faltantes:
        detalle = "\n".join(f" - {columna}" for columna in faltantes)

        raise ValueError(
            "El exporte de ANS REDES no contiene todas las columnas requeridas.\n\n"
            f"Columnas faltantes:\n{detalle}\n\n"
            "No se continuará con el procesamiento para evitar generar "
            "un informe incompleto o incorrecto."
        )


def validar_datos_minimos(df: pd.DataFrame) -> None:
    """
    Validaciones mínimas adicionales antes de procesar.
    """

    if df.empty:
        raise ValueError(
            "El exporte de ANS REDES fue leído correctamente, "
            "pero no contiene registros para procesar."
        )

    validar_columnas_requeridas(df)


def validar_exporte_redes() -> tuple[pd.DataFrame, Path]:
    """
    Lee el archivo desde entrada_redes y ejecuta todas las
    validaciones estructurales requeridas.

    Retorna:
        (dataframe_validado, ruta_archivo)
    """

    df, ruta_archivo = leer_exporte_redes()
    validar_datos_minimos(df)

    return df, ruta_archivo


# ============================================================
# PRUEBA MANUAL
# ============================================================

if __name__ == "__main__":
    try:
        datos, archivo = validar_exporte_redes()

        print("=" * 70)
        print("VALIDACIÓN ANS REDES CORRECTA")
        print("=" * 70)
        print(f"Archivo validado   : {archivo.name}")
        print(f"Registros leídos   : {len(datos):,}")
        print(f"Columnas requeridas: {len(COLUMNAS_REQUERIDAS_REDES)}")
        print()
        print("Todas las columnas obligatorias fueron encontradas:")
        for columna in COLUMNAS_REQUERIDAS_REDES:
            print(f" - {columna}")

    except Exception as error:
        print("=" * 70)
        print("ERROR DE VALIDACIÓN ANS REDES")
        print("=" * 70)
        print(error)
        raise SystemExit(1)
