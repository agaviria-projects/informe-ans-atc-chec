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

COLOR_VENCIDO = "F4CCCC"
COLOR_ALERTA = "FFF2CC"
COLOR_A_TIEMPO = "D9EAD3"
COLOR_PENDIENTE = "D9EAF7"


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
                "ELEMENTO": control["ARCHIVO"],
                "DETALLE": (
                    f"Región: {control['REGION']} | "
                    f"Hoja: {control['HOJA']} | "
                    f"Encabezados: fila "
                    f"{control['FILA_ENCABEZADOS']} | "
                    f"Filas vacías eliminadas: "
                    f"{control['FILAS_VACIAS_ELIMINADAS']} | "
                    f"Filas sin ID_ORDEN: "
                    f"{control['FILAS_SIN_ID_ORDEN']} | "
                    f"Registros válidos: "
                    f"{control['REGISTROS_VALIDOS']}"
                ),
            }
        )

    for clave, valor in control_transformacion.items():
        registros.append(
            {
                "TIPO": "CONSOLIDADO",
                "ELEMENTO": clave,
                "DETALLE": valor,
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


    # ==========================================================
    # VALIDAR ENCABEZADOS PARA LA TABLA ESTRUCTURADA
    # ==========================================================

    encabezados_utilizados: set[str] = set()

    for numero_columna in range(
        1,
        ultima_columna + 1,
    ):
        celda_encabezado = hoja.cell(
            row=1,
            column=numero_columna,
        )

        encabezado = str(
            celda_encabezado.value or ""
        ).strip()

        if not encabezado:
            encabezado = f"COLUMNA_{numero_columna}"

        encabezado_base = encabezado
        consecutivo = 2

        while encabezado in encabezados_utilizados:
            encabezado = (
                f"{encabezado_base}_{consecutivo}"
            )

            consecutivo += 1

        celda_encabezado.value = encabezado

        encabezados_utilizados.add(
            encabezado
        )


    # ==========================================================
    # CREAR TABLA ESTRUCTURADA
    # ==========================================================

    if ultima_fila >= 2 and ultima_columna >= 1:

        ultima_letra = hoja.cell(
            row=1,
            column=ultima_columna,
        ).column_letter

        referencia_tabla = (
            f"A1:{ultima_letra}{ultima_fila}"
        )

        tabla = Table(
            displayName="TablaDatosANS",
            ref=referencia_tabla,
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

    for fila in range(
        2,
        ultima_fila + 1,
    ):
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
            celda = hoja.cell(
                fila,
                columna,
            )

            if columna == 14:
                # OBSERVACION:
                # columna amplia y texto en una sola línea.
                celda.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=False,
                )

            else:
                celda.alignment = Alignment(
                    vertical="center",
                    wrap_text=columna in {
                        3,
                        4,
                    },
                )

        # Mantiene una altura uniforme y compacta.
        hoja.row_dimensions[fila].height = 20

    rango_estado = (
        f"M2:M{ultima_fila}"
    )

    hoja.conditional_formatting.add(
        rango_estado,
        FormulaRule(
            formula=['$M2="VENCIDO"'],
            fill=PatternFill(
                fill_type="solid",
                fgColor=COLOR_VENCIDO,
            ),
        ),
    )

    hoja.conditional_formatting.add(
        rango_estado,
        FormulaRule(
            formula=['$M2="ALERTA"'],
            fill=PatternFill(
                fill_type="solid",
                fgColor=COLOR_ALERTA,
            ),
        ),
    )

    hoja.conditional_formatting.add(
        rango_estado,
        FormulaRule(
            formula=['$M2="A TIEMPO"'],
            fill=PatternFill(
                fill_type="solid",
                fgColor=COLOR_A_TIEMPO,
            ),
        ),
    )

    hoja.conditional_formatting.add(
        rango_estado,
        FormulaRule(
            formula=['$M2="PENDIENTE CONFIGURACIÓN"'],
            fill=PatternFill(
                fill_type="solid",
                fgColor=COLOR_PENDIENTE,
            ),
        ),
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