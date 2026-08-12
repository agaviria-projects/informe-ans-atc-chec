from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ============================================================
# COMPATIBILIDAD DE IMPORTACIÓN
# ============================================================

DIRECTORIO_SRC = Path(__file__).resolve().parent
if str(DIRECTORIO_SRC) not in sys.path:
    sys.path.insert(0, str(DIRECTORIO_SRC))

from procesador_redes import procesar_exporte_redes


# ============================================================
# RUTAS
# ============================================================

RAIZ_PROYECTO = Path(__file__).resolve().parents[1]
CARPETA_SALIDA_REDES = RAIZ_PROYECTO / "salida_redes"

NOMBRE_ARCHIVO_SALIDA = "INFORME_ANS_REDES.xlsx"
NOMBRE_HOJA = "DATOS_ANS_REDES"


# ============================================================
# COLUMNAS DEL INFORME
# ============================================================

COLUMNAS_SALIDA = [
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
    "DIAS_CONTRACTUALES",
    "DIAS_TRANSCURRIDOS",
    "DIAS_RESTANTES",
    "ESTADO_ANS",
]


# ============================================================
# ANCHOS FIJOS
# ============================================================

ANCHOS_COLUMNAS = {
    "A": 18,
    "B": 12,
    "C": 32,
    "D": 20,
    "E": 22,
    "F": 26,
    "G": 22,
    "H": 16,
    "I": 28,
    "J": 35,
    "K": 24,
    "L": 24,
    "M": 24,
    "N": 22,
    "O": 14,
    "P": 24,
    "Q": 40,
    "R": 30,
    "S": 20,
    "T": 20,
    "U": 18,
    "V": 18,
}


# ============================================================
# ESTILOS
# ============================================================

RELLENO_ENCABEZADO = PatternFill(
    fill_type="solid",
    fgColor="008D46",
)

FUENTE_ENCABEZADO = Font(
    color="FFFFFF",
    bold=True,
)

BORDE_FINO = Side(
    style="thin",
    color="D9E0E5",
)

BORDE_CELDA = Border(
    left=BORDE_FINO,
    right=BORDE_FINO,
    top=BORDE_FINO,
    bottom=BORDE_FINO,
)


# ============================================================
# PREPARACIÓN DE DATOS
# ============================================================

def _preparar_fechas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte explícitamente las columnas de fecha antes de exportar.

    PRO_FECHA_SISTEMA_CREACION:
        conserva fecha y hora.

    PRO_FECHA_VENCIMIENTO:
        conserva solo la fecha, eliminando 00:00:00.
    """
    df = df.copy()

    df["PRO_FECHA_SISTEMA_CREACION"] = pd.to_datetime(
        df["PRO_FECHA_SISTEMA_CREACION"],
        errors="coerce",
        dayfirst=True,
    )

    df["PRO_FECHA_VENCIMIENTO"] = pd.to_datetime(
        df["PRO_FECHA_VENCIMIENTO"],
        errors="coerce",
        dayfirst=True,
    ).dt.normalize()

    return df


# ============================================================
# FORMATO EXCEL
# ============================================================

def _aplicar_formato(ruta_archivo: Path) -> None:
    wb = load_workbook(ruta_archivo)
    ws = wb[NOMBRE_HOJA]

    # Encabezados
    for celda in ws[1]:
        celda.fill = RELLENO_ENCABEZADO
        celda.font = FUENTE_ENCABEZADO
        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=False,
        )
        celda.border = BORDE_CELDA

    ws.row_dimensions[1].height = 22

    # Datos
    for fila in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column,
    ):
        for celda in fila:
            celda.alignment = Alignment(
                vertical="center",
                wrap_text=False,
            )
            celda.border = BORDE_CELDA

    # Anchos fijos
    for letra, ancho in ANCHOS_COLUMNAS.items():
        ws.column_dimensions[letra].width = ancho

    # Formatos de fecha
    encabezados = {
        ws.cell(row=1, column=columna).value: columna
        for columna in range(1, ws.max_column + 1)
    }

    col_creacion = encabezados.get("PRO_FECHA_SISTEMA_CREACION")
    if col_creacion:
        for fila in range(2, ws.max_row + 1):
            ws.cell(fila, col_creacion).number_format = "dd/mm/yyyy hh:mm:ss"

    col_vencimiento = encabezados.get("PRO_FECHA_VENCIMIENTO")
    if col_vencimiento:
        for fila in range(2, ws.max_row + 1):
            celda = ws.cell(fila, col_vencimiento)

            # Si Excel recibió datetime, dejarlo como fecha visual corta.
            celda.number_format = "dd/mm/yyyy"

    # Vista
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(ruta_archivo)
    wb.close()


# ============================================================
# GENERACIÓN
# ============================================================

def generar_informe_ans_redes() -> Path:
    CARPETA_SALIDA_REDES.mkdir(
        parents=True,
        exist_ok=True,
    )

    df, archivo_origen = procesar_exporte_redes()

    faltantes = [
        columna
        for columna in COLUMNAS_SALIDA
        if columna not in df.columns
    ]

    if faltantes:
        detalle = "\n".join(f" - {c}" for c in faltantes)
        raise ValueError(
            "No se puede generar INFORME_ANS_REDES.xlsx porque "
            f"faltan columnas:\n{detalle}"
        )

    df_salida = df[COLUMNAS_SALIDA].copy()
    df_salida = _preparar_fechas(df_salida)

    ruta_salida = CARPETA_SALIDA_REDES / NOMBRE_ARCHIVO_SALIDA

    try:
        with pd.ExcelWriter(
            ruta_salida,
            engine="openpyxl",
            mode="w",
            datetime_format="dd/mm/yyyy hh:mm:ss",
            date_format="dd/mm/yyyy",
        ) as writer:
            df_salida.to_excel(
                writer,
                sheet_name=NOMBRE_HOJA,
                index=False,
            )

    except PermissionError as error:
        raise PermissionError(
            f"No fue posible actualizar:\n{ruta_salida}\n\n"
            "Cierre INFORME_ANS_REDES.xlsx en Excel y vuelva a ejecutar."
        ) from error

    _aplicar_formato(ruta_salida)

    print("=" * 70)
    print("INFORME ANS REDES GENERADO CORRECTAMENTE")
    print("=" * 70)
    print(f"Archivo origen      : {archivo_origen.name}")
    print(f"Registros generados : {len(df_salida):,}")
    print(f"Columnas generadas  : {len(df_salida.columns)}")
    print(f"Hoja                : {NOMBRE_HOJA}")
    print(f"Archivo de salida   : {ruta_salida}")
    print()
    print("FORMATO:")
    print(" - Columnas con anchos fijos y legibles.")
    print(" - Encabezados en una sola línea.")
    print(" - PRO_FECHA_SISTEMA_CREACION: fecha + hora.")
    print(" - PRO_FECHA_VENCIMIENTO: fecha corta.")

    return ruta_salida


# ============================================================
# EJECUCIÓN MANUAL
# ============================================================

if __name__ == "__main__":
    try:
        generar_informe_ans_redes()

    except Exception as error:
        print("=" * 70)
        print("ERROR AL GENERAR INFORME ANS REDES")
        print("=" * 70)
        print(error)
        raise SystemExit(1)
