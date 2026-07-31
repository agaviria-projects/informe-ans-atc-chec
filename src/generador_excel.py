import logging
from pathlib import Path

import pandas as pd

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.worksheet.table import (
    Table,
    TableStyleInfo,
)

from src.config import (
    HOJA_CONTROL,
    HOJA_DATOS_SALIDA,
    RUTA_INFORME_EXCEL,
    SALIDA_DIR,
)


logger = logging.getLogger(__name__)


COLOR_VERDE_OSCURO = "1E8449"
COLOR_VERDE_CLARO = "E9F7EF"
COLOR_BLANCO = "FFFFFF"
COLOR_BORDE = "AAB7B8"

# ==========================================================
# COLORES DE ESTADOS ANS
# ==========================================================

COLOR_VENCIDO = "FF1F1F"       # Rojo encendido
COLOR_ALERTA = "FFD426"        # Amarillo intenso
COLOR_A_TIEMPO = "8BCF4A"      # Verde claro intenso
COLOR_PENDIENTE = "647687"     # Gris azulado

COLOR_TEXTO_ESTADO = "000000"
COLOR_TEXTO_PENDIENTE = "FFFFFF"


ANCHOS_COLUMNAS = {
    "A": 16,  # ID_ORDEN
    "B": 15,  # FECHA_ORDEN
    "C": 45,  # DIRECCION
    "D": 33,  # PROPIETARIO
    "E": 10,  # ZONA
    "F": 15,  # MUNICIPIO
    "G": 24,  # DESC_MUNICIPIO
    "H": 17,  # REGION_ORIGEN
    "I": 16,  # DIAS_PACTADOS
    "J": 20,  # FECHA_LIMITE_ANS
    "K": 22,  # DIAS_TRANSCURRIDOS
    "L": 18,  # DIAS_RESTANTES
    "M": 25,  # ESTADO
    "N": 110,  # OBSERVACION
}


class ErrorGeneracionExcel(Exception):
    """Error controlado al generar el informe Excel."""


def crear_dataframe_control(
    controles_archivos: list[dict],
    control_transformacion: dict,
) -> pd.DataFrame:
    """
    Construye una hoja de trazabilidad del proceso.
    """

    registros: list[dict] = []

    for control in controles_archivos:

        registros.append(
            {
                "TIPO": "ARCHIVO",
                "ELEMENTO": control["REGION"],
                "DETALLE": (
                    f"{control['REGISTROS_VALIDOS']} "
                    f"registros procesados correctamente"
                ),
            }
        )

    registros.append(
        {
            "TIPO": "RESUMEN",
            "ELEMENTO": "Total consolidado",
            "DETALLE": (
                f"{control_transformacion['REGISTROS_CONSOLIDADOS']} registros"
            ),
        }
    )

    return pd.DataFrame(
        registros
    )


def aplicar_diseno_hoja_datos(
    ruta_archivo: Path,
    ) -> None:
    """
    Convierte DATOS_ANS en una tabla profesional.
    """

    libro = load_workbook(
        ruta_archivo
    )

    hoja = libro[
        HOJA_DATOS_SALIDA
    ]

    hoja.freeze_panes = "A2"
    hoja.sheet_view.showGridLines = False

    borde_fino = Side(
        style="thin",
        color=COLOR_BORDE,
    )

    for celda in hoja[1]:
        celda.fill = PatternFill(
            fill_type="solid",
            fgColor=COLOR_VERDE_OSCURO,
        )

        celda.font = Font(
            color=COLOR_BLANCO,
            bold=True,
        )

        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        celda.border = Border(
            left=borde_fino,
            right=borde_fino,
            top=borde_fino,
            bottom=borde_fino,
        )

    hoja.row_dimensions[1].height = 30

    ultima_fila = hoja.max_row
    ultima_columna = hoja.max_column

    if ultima_fila >= 2:

        # ------------------------------------------------------
        # Si la tabla ya existe, eliminarla antes de crearla.
        # ------------------------------------------------------

        if "TablaDatosANS" in hoja.tables:
            del hoja.tables["TablaDatosANS"]

        referencia = (
            f"A1:"
            f"{hoja.cell(1, ultima_columna).column_letter}"
            f"{ultima_fila}"
        )

        tabla = Table(
            displayName="TablaDatosANS",
            ref=referencia,
        )

        estilo_tabla = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

        tabla.tableStyleInfo = estilo_tabla

        hoja.add_table(
            tabla
        )

    for columna, ancho in ANCHOS_COLUMNAS.items():
        hoja.column_dimensions[columna].width = ancho

    # ======================================================
    # FORMATO DE LAS FILAS DEL INFORME
    # ======================================================

    for fila in range(
        2,
        ultima_fila + 1,
    ):

        # ----------------------------------------------
        # FORMATOS NUMÉRICOS Y DE FECHA
        # ----------------------------------------------

        hoja.cell(
            row=fila,
            column=1,
        ).number_format = "@"

        hoja.cell(
            row=fila,
            column=2,
        ).number_format = "dd/mm/yyyy"

        hoja.cell(
            row=fila,
            column=6,
        ).number_format = "@"

        hoja.cell(
            row=fila,
            column=10,
        ).number_format = "dd/mm/yyyy"

        # ----------------------------------------------
        # ALINEACIÓN GENERAL
        # ----------------------------------------------

        for columna in range(
            1,
            ultima_columna + 1,
        ):

            celda = hoja.cell(
                row=fila,
                column=columna,
            )

            # Columnas de texto descriptivo.
            if columna in {
                3,   # DIRECCION
                4,   # PROPIETARIO
                14,  # OBSERVACION
            }:
                celda.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=False,
                )

            # Columnas numéricas, fechas y estados.
            elif columna in {
                2,   # FECHA_ORDEN
                5,   # ZONA
                6,   # MUNICIPIO
                8,   # REGION_ORIGEN
                9,   # DIAS_PACTADOS
                10,  # FECHA_LIMITE_ANS
                11,  # DIAS_TRANSCURRIDOS
                12,  # DIAS_RESTANTES
                13,  # ESTADO
            }:
                celda.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=False,
                )

            else:
                celda.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=False,
                )

        # Mantiene todas las filas compactas y uniformes.
        hoja.row_dimensions[fila].height = 22

        hoja.cell(
            fila,
            1,
        ).number_format = "@"

        hoja.cell(
            fila,
            2,
        ).number_format = "dd/mm/yyyy"

        hoja.cell(
            fila,
            6,
        ).number_format = "@"

        hoja.cell(
            fila,
            10,
        ).number_format = "dd/mm/yyyy"

        for columna in range(
            1,
            ultima_columna + 1,
        ):
            hoja.cell(
                fila,
                columna,
            ).alignment = Alignment(
                vertical="top",
                wrap_text=columna in {
                    3,
                    4,
                    14,
                },
            )

        hoja.cell(
            fila,
            13,
        ).alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    # ======================================================
    # COLORES DIRECTOS PARA LOS ESTADOS ANS
    # ======================================================

    for fila in range(
        2,
        ultima_fila + 1,
    ):

        celda_estado = hoja.cell(
            row=fila,
            column=13,
        )

        estado = str(
            celda_estado.value or ""
        ).strip().upper()

        if estado == "VENCIDO":

            celda_estado.fill = PatternFill(
                fill_type="solid",
                fgColor=COLOR_VENCIDO,
            )

            celda_estado.font = Font(
                color=COLOR_TEXTO_ESTADO,
                bold=True,
            )

        elif estado == "ALERTA":

            celda_estado.fill = PatternFill(
                fill_type="solid",
                fgColor=COLOR_ALERTA,
            )

            celda_estado.font = Font(
                color=COLOR_TEXTO_ESTADO,
                bold=True,
            )

        elif estado == "A TIEMPO":

            celda_estado.fill = PatternFill(
                fill_type="solid",
                fgColor=COLOR_A_TIEMPO,
            )

            celda_estado.font = Font(
                color=COLOR_TEXTO_ESTADO,
                bold=True,
            )

        elif estado in {
            "SIN FECHA",
            "PENDIENTE CONFIGURACIÓN",
        }:

            celda_estado.fill = PatternFill(
                fill_type="solid",
                fgColor=COLOR_PENDIENTE,
            )

            celda_estado.font = Font(
                color=COLOR_TEXTO_PENDIENTE,
                bold=True,
            )

        celda_estado.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=False,
        )
        
    libro.save(
        ruta_archivo
    )

def aplicar_diseno_hoja_control(
    ruta_archivo: Path,
) -> None:
    """
    Aplica formato a la hoja de auditoría.
    """

    libro = load_workbook(
        ruta_archivo
    )

    hoja = libro[
        HOJA_CONTROL
    ]

    hoja.freeze_panes = "A2"
    hoja.sheet_view.showGridLines = False

    for celda in hoja[1]:
        celda.fill = PatternFill(
            fill_type="solid",
            fgColor=COLOR_VERDE_OSCURO,
        )

        celda.font = Font(
            color=COLOR_BLANCO,
            bold=True,
        )

        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    hoja.column_dimensions["A"].width = 18
    hoja.column_dimensions["B"].width = 35
    hoja.column_dimensions["C"].width = 100

    for fila in hoja.iter_rows(
        min_row=2,
    ):
        for celda in fila:
            celda.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    libro.save(
        ruta_archivo
    )


def generar_informe_excel(
    dataframe: pd.DataFrame,
    controles_archivos: list[dict],
    control_transformacion: dict,
) -> Path:
    """
    Genera el único informe unificado.
    """

    try:
        SALIDA_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        control_dataframe = crear_dataframe_control(
            controles_archivos,
            control_transformacion,
        )

        with pd.ExcelWriter(
            RUTA_INFORME_EXCEL,
            engine="openpyxl",
            mode="w",
        ) as escritor:

            dataframe.to_excel(
                escritor,
                sheet_name=HOJA_DATOS_SALIDA,
                index=False,
            )

            control_dataframe.to_excel(
                escritor,
                sheet_name=HOJA_CONTROL,
                index=False,
            )

        aplicar_diseno_hoja_datos(
            RUTA_INFORME_EXCEL
        )

        aplicar_diseno_hoja_control(
            RUTA_INFORME_EXCEL
        )

        logger.info(
            "Informe Excel generado: %s",
            RUTA_INFORME_EXCEL,
        )

        return RUTA_INFORME_EXCEL

    except Exception as error:
        logger.exception(
            "No fue posible generar el informe Excel."
        )

        raise ErrorGeneracionExcel(
            f"No fue posible generar el informe: {error}"
        ) from error