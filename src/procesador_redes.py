from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd


# ============================================================
# COMPATIBILIDAD DE IMPORTACIÓN
# ============================================================

DIRECTORIO_SRC = Path(__file__).resolve().parent
if str(DIRECTORIO_SRC) not in sys.path:
    sys.path.insert(0, str(DIRECTORIO_SRC))

from validador_redes import (
    COLUMNAS_REQUERIDAS_REDES,
    validar_exporte_redes,
)


# ============================================================
# REGLAS DE FILTRADO ANS REDES
# ============================================================

PROCESOS_PERMITIDOS = {
    "4109",
    "4133",
    "4166",
    "4167",
}

REV_ESTADO_PERMITIDO = "P"

REV_RESPONSABLES_PERMITIDOS = {
    r"CHEC\AHERRERV",
    r"CHEC\BQUINTES",
    r"CHEC\JGILFLOR",
    r"CHEC\MCARVAAB",
    r"CHEC\NSANCHAG",
    r"CHEC\SAGUDELJ",
    r"CHEC\SUCHIMAT",
}


# ============================================================
# COLUMNAS ANS PENDIENTES DE REGLA CONTRACTUAL
# ============================================================

COLUMNAS_ANS_PENDIENTES = [
    "DIAS_CONTRACTUALES",
    "DIAS_TRANSCURRIDOS",
    "DIAS_RESTANTES",
    "ESTADO_ANS",
]


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def _texto_serie(serie: pd.Series) -> pd.Series:
    """
    Convierte una Serie a texto de forma segura, elimina espacios
    laterales y normaliza valores vacíos.

    No modifica el contenido funcional del resto de columnas.
    """
    return (
        serie.fillna("")
        .astype(str)
        .str.strip()
    )


def _normalizar_proceso(serie: pd.Series) -> pd.Series:
    """
    Normaliza PROCESO para evitar problemas cuando Excel lo entrega
    como número (4109) o como decimal textual (4109.0).
    """
    texto = _texto_serie(serie)

    return texto.str.replace(r"\.0$", "", regex=True)


# ============================================================
# PROCESAMIENTO
# ============================================================

def procesar_redes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica exclusivamente los filtros definidos para ANS REDES:

    1. PROCESO:
       4109, 4133, 4166, 4167

    2. REV_ESTADO:
       P

    3. REV_RESPONSABLE:
       responsables CHEC autorizados

    IMPORTANTE:
    CODIGO_UBIC_TRANSFORMADOR NO es un filtro.
    Solo se conserva como columna del informe final.
    """

    df_trabajo = df.copy()

    proceso = _normalizar_proceso(df_trabajo["PROCESO"])

    rev_estado = (
        _texto_serie(df_trabajo["REV_ESTADO"])
        .str.upper()
    )

    rev_responsable = (
        _texto_serie(df_trabajo["REV_RESPONSABLE"])
        .str.upper()
    )

    responsables_permitidos = {
        responsable.upper()
        for responsable in REV_RESPONSABLES_PERMITIDOS
    }

    mascara = (
        proceso.isin(PROCESOS_PERMITIDOS)
        & rev_estado.eq(REV_ESTADO_PERMITIDO)
        & rev_responsable.isin(responsables_permitidos)
    )

    df_filtrado = df_trabajo.loc[
        mascara,
        COLUMNAS_REQUERIDAS_REDES,
    ].copy()

    # --------------------------------------------------------
    # COLUMNAS ANS:
    # Se crean desde ahora para dejar preparado el diseño final,
    # pero permanecen vacías hasta recibir los días contractuales
    # y las reglas oficiales del cálculo.
    # --------------------------------------------------------

    for columna in COLUMNAS_ANS_PENDIENTES:
        df_filtrado[columna] = pd.NA

    df_filtrado = df_filtrado.reset_index(drop=True)

    return df_filtrado


def procesar_exporte_redes() -> tuple[pd.DataFrame, Path]:
    """
    Flujo completo de esta fase:

        entrada_redes
            ↓
        lector_redes
            ↓
        validador_redes
            ↓
        filtros ANS REDES
            ↓
        columnas requeridas
            ↓
        columnas ANS vacías

    Retorna:
        (dataframe_procesado, ruta_archivo_origen)
    """

    df, ruta_archivo = validar_exporte_redes()
    resultado = procesar_redes(df)

    if resultado.empty:
        raise ValueError(
            "El archivo fue leído y validado correctamente, pero después "
            "de aplicar los filtros de PROCESO, REV_ESTADO y "
            "REV_RESPONSABLE no quedaron registros para el informe ANS REDES."
        )

    return resultado, ruta_archivo


# ============================================================
# PRUEBA MANUAL
# ============================================================

if __name__ == "__main__":
    try:
        resultado, archivo = procesar_exporte_redes()

        print("=" * 70)
        print("PROCESAMIENTO ANS REDES CORRECTO")
        print("=" * 70)
        print(f"Archivo procesado   : {archivo.name}")
        print(f"Registros finales   : {len(resultado):,}")
        print(f"Columnas finales    : {len(resultado.columns)}")
        print()
        print("Filtros aplicados:")
        print(" - PROCESO: 4109, 4133, 4166, 4167")
        print(" - REV_ESTADO: P")
        print(" - REV_RESPONSABLE: usuarios CHEC autorizados")
        print()
        print("IMPORTANTE:")
        print(" - CODIGO_UBIC_TRANSFORMADOR se conserva, pero NO se usa como filtro.")
        print(" - Las columnas ANS permanecen vacías hasta definir días contractuales.")
        print()
        print("Columnas finales:")
        for columna in resultado.columns:
            print(f" - {columna}")

    except Exception as error:
        print("=" * 70)
        print("ERROR EN PROCESAMIENTO ANS REDES")
        print("=" * 70)
        print(error)
        raise SystemExit(1)
