from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from src.calculador_ans_redes import aplicar_calculos_ans_redes
from src.config import BASE_DIR
from src.validador_redes import (
    COLUMNAS_REQUERIDAS_REDES,
    validar_exporte_redes,
)


# ============================================================
# RUTAS Y CONFIGURACIÓN
# ============================================================

RUTA_CONFIG_REDES = (
    BASE_DIR
    / "config"
    / "FILTROS_ANS_REDES.xlsx"
)

HOJA_PARAMETROS_REDES = "PARAMETROS_REDES"

CAMPOS_FILTRO_PERMITIDOS = {
    "PROCESO",
    "REV_ESTADO",
    "REV_RESPONSABLE",
}


def _texto_serie(serie: pd.Series) -> pd.Series:
    return serie.fillna("").astype(str).str.strip()


def _normalizar_proceso(serie: pd.Series) -> pd.Series:
    return _texto_serie(serie).str.replace(r"\.0$", "", regex=True)


def _normalizar_valor_proceso(valor: object) -> str:
    return re.sub(r"\.0$", "", str(valor).strip())


def _es_activo(valor: object) -> bool:
    texto = str(valor).strip().upper()

    if texto in {"SI", "S", "TRUE", "VERDADERO", "1"}:
        return True

    if texto in {"NO", "N", "FALSE", "FALSO", "0"}:
        return False

    raise ValueError(
        f"El valor ACTIVO debe ser SI o NO. Valor recibido: {valor}"
    )


def cargar_filtros_redes() -> dict[str, set[str]]:
    """
    Lee FILTROS_ANS_REDES.xlsx / PARAMETROS_REDES.

    Columnas:
    CAMPO | VALOR | ACTIVO
    """

    if not RUTA_CONFIG_REDES.exists():
        raise FileNotFoundError(
            "No se encontró el archivo de configuración ANS Redes:\n\n"
            f"{RUTA_CONFIG_REDES}"
        )

    try:
        df = pd.read_excel(
            RUTA_CONFIG_REDES,
            sheet_name=HOJA_PARAMETROS_REDES,
            dtype=object,
            engine="openpyxl",
        )
    except PermissionError as error:
        raise PermissionError(
            "No fue posible leer FILTROS_ANS_REDES.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a ejecutar."
        ) from error
    except ValueError as error:
        raise ValueError(
            f"No se encontró la hoja {HOJA_PARAMETROS_REDES} "
            "en FILTROS_ANS_REDES.xlsx."
        ) from error

    df.columns = [str(c).strip().upper() for c in df.columns]

    requeridas = {"CAMPO", "VALOR", "ACTIVO"}
    faltantes = requeridas.difference(df.columns)

    if faltantes:
        raise ValueError(
            "Faltan columnas en PARAMETROS_REDES:\n\n"
            + "\n".join(f"- {c}" for c in sorted(faltantes))
        )

    filtros = {
        "PROCESO": set(),
        "REV_ESTADO": set(),
        "REV_RESPONSABLE": set(),
    }

    for numero_fila, fila in df.iterrows():

        activo = fila.get("ACTIVO")

        if pd.isna(activo):
            continue

        if not _es_activo(activo):
            continue

        campo = str(fila.get("CAMPO", "")).strip().upper()

        if not campo:
            continue

        if campo not in CAMPOS_FILTRO_PERMITIDOS:
            raise ValueError(
                f"Campo de filtro no reconocido en la fila "
                f"{numero_fila + 2}: {campo}"
            )

        valor = fila.get("VALOR")

        if pd.isna(valor):
            continue

        valor = str(valor).strip()

        if not valor:
            continue

        if campo == "PROCESO":
            valor = _normalizar_valor_proceso(valor)
        else:
            valor = valor.upper()

        filtros[campo].add(valor)

    for campo, valores in filtros.items():
        if not valores:
            raise ValueError(
                f"No existen valores activos configurados para {campo} "
                "en PARAMETROS_REDES."
            )

    return filtros


def procesar_redes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Lee filtros desde Excel, filtra el exporte y después
    llama calculador_ans_redes.py.
    """

    filtros = cargar_filtros_redes()
    df_trabajo = df.copy()

    proceso = _normalizar_proceso(df_trabajo["PROCESO"])
    rev_estado = _texto_serie(df_trabajo["REV_ESTADO"]).str.upper()
    rev_responsable = _texto_serie(
        df_trabajo["REV_RESPONSABLE"]
    ).str.upper()

    mascara = (
        proceso.isin(filtros["PROCESO"])
        & rev_estado.isin(filtros["REV_ESTADO"])
        & rev_responsable.isin(filtros["REV_RESPONSABLE"])
    )

    df_filtrado = df_trabajo.loc[
        mascara,
        COLUMNAS_REQUERIDAS_REDES,
    ].copy()

    if df_filtrado.empty:
        raise ValueError(
            "Después de aplicar los filtros configurados en "
            "PARAMETROS_REDES no quedaron registros."
        )

    # ========================================================
    # NORMALIZACIÓN VISUAL PARA DASHBOARD
    # ========================================================

    columnas_mayusculas = [
        "D_PROCESO",
        "PRO_D_CLASIFICACION",
        "REV_RESPONSABLE",
        "CLI_D_MUNICIPIO",
    ]

    for columna in columnas_mayusculas:
        if columna in df_filtrado.columns:
            df_filtrado[columna] = (
                df_filtrado[columna]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

    # Aquí se aplican las reglas contractuales y de calendario
    # leídas por calculador_ans_redes.py desde FILTROS_ANS_REDES.xlsx.
    df_filtrado = aplicar_calculos_ans_redes(df_filtrado)

    # ESTADO se crea después del cálculo.
    if "ESTADO" in df_filtrado.columns:
        df_filtrado["ESTADO"] = (
            df_filtrado["ESTADO"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    return df_filtrado.reset_index(drop=True)


def procesar_exporte_redes() -> tuple[pd.DataFrame, Path]:
    df, ruta_archivo = validar_exporte_redes()
    resultado = procesar_redes(df)
    return resultado, ruta_archivo


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
        print("Filtros leídos desde PARAMETROS_REDES:")
        print(" - PROCESO")
        print(" - REV_ESTADO")
        print(" - REV_RESPONSABLE")
        print()
        print("Cálculos ANS aplicados desde calculador_ans_redes.py:")
        print(" - DIAS_CONTRACTUALES")
        print(" - FECHA_LIMITE_ANS")
        print(" - DIAS_TRANSCURRIDOS")
        print(" - DIAS_PARA_INICIAR_ALERTA")
        print(" - DIAS_RESTANTES")
        print(" - ESTADO")

    except Exception as error:
        print("=" * 70)
        print("ERROR EN PROCESAMIENTO ANS REDES")
        print("=" * 70)
        print(error)
        raise SystemExit(1)
