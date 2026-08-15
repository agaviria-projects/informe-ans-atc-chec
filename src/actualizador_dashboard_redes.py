from __future__ import annotations

import logging
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
import pythoncom
import win32com.client as win32


logger = logging.getLogger(__name__)


# ============================================================
# RUTAS
# ============================================================

BASE_DIR = Path(__file__).resolve().parents[1]

RUTA_INFORME_REDES = (
    BASE_DIR
    / "salida_redes"
    / "INFORME_ANS_REDES.xlsx"
)

RUTA_DASHBOARD_REDES = (
    BASE_DIR
    / "dashboard"
    / "INFORME ANS-REDES.xlsb"
)

HOJA_ORIGEN = "DATOS_ANS_REDES"

HOJA_DESTINO = "DATOS_ANS"

NOMBRE_TABLA_DESTINO = "REDES"


# ============================================================
# COLUMNAS DEL DASHBOARD REDES
# ============================================================

COLUMNAS_DASHBOARD_REDES = [
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
    "FECHA_LIMITE_ANS",
    "DIAS_TRANSCURRIDOS",
    "DIAS_PARA_INICIAR_ALERTA",
    "DIAS_RESTANTES",
    "ESTADO",
]


# ============================================================
# COLORES
# ============================================================

COLOR_ESTADOS = {
    "ALERTA": (255, 212, 38),
    "VENCIDO": (255, 31, 31),
    "ALERTA 0 DIAS": (243, 156, 18),
    "A TIEMPO": (139, 207, 74),
    "SIN FECHA": (100, 118, 135),
}


def color_excel(
    rojo: int,
    verde: int,
    azul: int,
) -> int:
    """
    Convierte RGB al entero utilizado por Excel COM.
    """
    return rojo + (verde * 256) + (azul * 65536)


class ErrorActualizacionDashboardRedes(Exception):
    """
    Error controlado durante la actualización del dashboard ANS REDES.
    """


# ============================================================
# VALIDACIONES
# ============================================================

def validar_archivos() -> None:
    if not RUTA_INFORME_REDES.exists():
        raise ErrorActualizacionDashboardRedes(
            "No se encontró el informe ANS Redes generado:\n\n"
            f"{RUTA_INFORME_REDES}\n\n"
            "Primero genere el Informe ANS Redes."
        )

    if not RUTA_DASHBOARD_REDES.exists():
        raise ErrorActualizacionDashboardRedes(
            "No se encontró el dashboard ANS Redes:\n\n"
            f"{RUTA_DASHBOARD_REDES}"
        )


def validar_columnas(
    dataframe: pd.DataFrame,
) -> None:
    faltantes = [
        columna
        for columna in COLUMNAS_DASHBOARD_REDES
        if columna not in dataframe.columns
    ]

    if faltantes:
        raise ErrorActualizacionDashboardRedes(
            "Faltan columnas en INFORME_ANS_REDES.xlsx:\n\n"
            + "\n".join(
                f"- {columna}"
                for columna in faltantes
            )
        )


# ============================================================
# CARGA DEL INFORME REDES
# ============================================================

def cargar_datos_informe_redes() -> pd.DataFrame:
    try:
        dataframe = pd.read_excel(
            RUTA_INFORME_REDES,
            sheet_name=HOJA_ORIGEN,
            dtype=object,
            engine="openpyxl",
        )

    except PermissionError as error:
        raise ErrorActualizacionDashboardRedes(
            "No fue posible leer INFORME_ANS_REDES.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a ejecutar."
        ) from error

    except ValueError as error:
        raise ErrorActualizacionDashboardRedes(
            f"No se encontró la hoja {HOJA_ORIGEN} "
            "en INFORME_ANS_REDES.xlsx."
        ) from error

    validar_columnas(
        dataframe
    )

    dataframe = dataframe[
        COLUMNAS_DASHBOARD_REDES
    ].copy()

    # NUMERO_PROCESO identifica los registros reales.
    dataframe["NUMERO_PROCESO"] = (
        dataframe["NUMERO_PROCESO"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True,
        )
    )

    dataframe = dataframe[
        dataframe["NUMERO_PROCESO"].ne("")
    ].copy()

    # Fechas.
    for columna in [
        "PRO_FECHA_SISTEMA_CREACION",
        "PRO_FECHA_VENCIMIENTO",
        "FECHA_LIMITE_ANS",
    ]:
        dataframe[columna] = pd.to_datetime(
            dataframe[columna],
            errors="coerce",
            dayfirst=True,
        )

    return dataframe.reset_index(
        drop=True
    )


# ============================================================
# CONVERSIÓN PARA EXCEL COM
# ============================================================

def convertir_valor_excel(
    valor: Any,
) -> Any:
    if valor is None:
        return None

    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(
        valor,
        pd.Timestamp,
    ):
        return valor.to_pydatetime()

    if isinstance(
        valor,
        datetime,
    ):
        return valor

    if isinstance(
        valor,
        date,
    ):
        return datetime.combine(
            valor,
            datetime.min.time(),
        )

    if hasattr(
        valor,
        "item",
    ):
        try:
            return valor.item()
        except (
            ValueError,
            AttributeError,
        ):
            pass

    return valor


def construir_matriz_excel(
    dataframe: pd.DataFrame,
) -> tuple[tuple[Any, ...], ...]:
    return tuple(
        tuple(
            convertir_valor_excel(valor)
            for valor in fila
        )
        for fila in dataframe.itertuples(
            index=False,
            name=None,
        )
    )


# ============================================================
# FORMATO DE DATOS_ANS
# ============================================================

def aplicar_formato_datos_ans_redes(
    hoja: Any,
    cantidad_registros: int,
) -> None:
    if cantidad_registros <= 0:
        return

    ultima_fila = cantidad_registros + 1
    ultima_columna = len(
        COLUMNAS_DASHBOARD_REDES
    )

    # --------------------------------------------------------
    # ENCABEZADOS
    # --------------------------------------------------------

    rango_encabezados = hoja.Range(
        hoja.Cells(1, 1),
        hoja.Cells(1, ultima_columna),
    )

    rango_encabezados.FormatConditions.Delete()
    rango_encabezados.Interior.Pattern = 1
    rango_encabezados.Interior.Color = color_excel(
        0,
        141,
        70,
    )
    rango_encabezados.Font.Color = color_excel(
        255,
        255,
        255,
    )
    rango_encabezados.Font.Bold = True
    rango_encabezados.HorizontalAlignment = -4108
    rango_encabezados.VerticalAlignment = -4108
    rango_encabezados.WrapText = True

    # --------------------------------------------------------
    # FORMATOS DE FECHA
    # --------------------------------------------------------

    # F = PRO_FECHA_SISTEMA_CREACION
    hoja.Range(
        hoja.Cells(2, 6),
        hoja.Cells(ultima_fila, 6),
    ).NumberFormat = "dd/mm/yyyy hh:mm:ss"

    # G = PRO_FECHA_VENCIMIENTO
    hoja.Range(
        hoja.Cells(2, 7),
        hoja.Cells(ultima_fila, 7),
    ).NumberFormat = "dd/mm/yyyy"

    # T = FECHA_LIMITE_ANS
    hoja.Range(
        hoja.Cells(2, 20),
        hoja.Cells(ultima_fila, 20),
    ).NumberFormat = "dd/mm/yyyy hh:mm:ss"

    # --------------------------------------------------------
    # ENTEROS
    # --------------------------------------------------------

    for numero_columna in (
        19,  # DIAS_CONTRACTUALES
        21,  # DIAS_TRANSCURRIDOS
        22,  # DIAS_PARA_INICIAR_ALERTA
        23,  # DIAS_RESTANTES
    ):
        hoja.Range(
            hoja.Cells(2, numero_columna),
            hoja.Cells(ultima_fila, numero_columna),
        ).NumberFormat = "0"

    # --------------------------------------------------------
    # ALINEACIÓN
    # --------------------------------------------------------

    rango_datos = hoja.Range(
        hoja.Cells(2, 1),
        hoja.Cells(ultima_fila, ultima_columna),
    )

    rango_datos.VerticalAlignment = -4108
    rango_datos.WrapText = False

    for numero_columna in (
        2, 4, 5, 6, 7, 11, 12, 13, 14, 15,
        18, 19, 20, 21, 22, 23, 24,
    ):
        hoja.Range(
            hoja.Cells(2, numero_columna),
            hoja.Cells(ultima_fila, numero_columna),
        ).HorizontalAlignment = -4108

    # --------------------------------------------------------
    # ESTADO - COLUMNA X
    # --------------------------------------------------------

    rango_estado = hoja.Range(
        hoja.Cells(2, 24),
        hoja.Cells(ultima_fila, 24),
    )

    rango_estado.FormatConditions.Delete()
    rango_estado.Interior.Pattern = -4142
    rango_estado.Font.Bold = True
    rango_estado.Font.Color = color_excel(
        0,
        0,
        0,
    )

    for numero_fila in range(
        2,
        ultima_fila + 1,
    ):
        celda = hoja.Cells(
            numero_fila,
            24,
        )

        estado = str(
            celda.Value or ""
        ).strip().upper()

        color = COLOR_ESTADOS.get(
            estado
        )

        if color is None:
            celda.Interior.Pattern = -4142
            continue

        celda.Interior.Color = color_excel(
            *color
        )

        if estado == "SIN FECHA":
            celda.Font.Color = color_excel(
                255,
                255,
                255,
            )
        else:
            celda.Font.Color = color_excel(
                0,
                0,
                0,
            )

    # --------------------------------------------------------
    # ANCHOS PRINCIPALES
    # --------------------------------------------------------

    anchos = {
        1: 18,
        2: 12,
        3: 30,
        4: 20,
        5: 22,
        6: 24,
        7: 20,
        8: 16,
        9: 28,
        10: 34,
        11: 22,
        12: 22,
        13: 22,
        14: 20,
        15: 14,
        16: 24,
        17: 40,
        18: 28,
        19: 20,
        20: 24,
        21: 20,
        22: 24,
        23: 18,
        24: 18,
    }

    for numero_columna, ancho in anchos.items():
        hoja.Columns(
            numero_columna
        ).ColumnWidth = ancho


# ============================================================
# ACTUALIZACIÓN DE ELEMENTOS DEL DASHBOARD
# ============================================================

def esperar_actualizacion_excel(
    aplicacion_excel: Any,
    segundos_maximos: int = 120,
) -> None:
    try:
        aplicacion_excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass

    tiempo_inicio = time.time()

    while (
        time.time() - tiempo_inicio
        < segundos_maximos
    ):
        try:
            if aplicacion_excel.CalculationState == 0:
                return
        except Exception:
            return

        time.sleep(
            0.5
        )


def actualizar_elementos_dashboard(
    libro_excel: Any,
    aplicacion_excel: Any,
) -> None:
    try:
        libro_excel.RefreshAll()
    except Exception:
        logger.warning(
            "No fue posible ejecutar RefreshAll.",
            exc_info=True,
        )

    esperar_actualizacion_excel(
        aplicacion_excel
    )

    try:
        cantidad_cache = libro_excel.PivotCaches().Count

        for indice in range(
            1,
            cantidad_cache + 1,
        ):
            try:
                libro_excel.PivotCaches().Item(
                    indice
                ).Refresh()
            except Exception:
                logger.warning(
                    "No fue posible actualizar una caché de tabla dinámica.",
                    exc_info=True,
                )

    except Exception:
        pass

    try:
        aplicacion_excel.CalculateFull()
    except Exception:
        pass


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def actualizar_dashboard_redes(
    refrescar_dashboard: bool = True,
) -> dict:
    """
    Transfiere los datos más recientes de INFORME_ANS_REDES.xlsx
    al dashboard INFORME ANS-REDES.xlsb.

    Destino:
        Hoja  : DATOS_ANS
        Tabla : REDES
    """

    excel = None
    libro = None
    com_inicializado = False

    try:
        pythoncom.CoInitialize()
        com_inicializado = True

        validar_archivos()

        dataframe = (
            cargar_datos_informe_redes()
        )

        cantidad_registros = len(
            dataframe
        )

        cantidad_columnas = len(
            COLUMNAS_DASHBOARD_REDES
        )

        excel = win32.DispatchEx(
            "Excel.Application"
        )

        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False

        libro = excel.Workbooks.Open(
            str(
                RUTA_DASHBOARD_REDES.resolve()
            ),
            UpdateLinks=0,
            ReadOnly=False,
        )

        # ----------------------------------------------------
        # HOJA DESTINO
        # ----------------------------------------------------

        try:
            hoja = libro.Worksheets(
                HOJA_DESTINO
            )
        except Exception as error:
            raise ErrorActualizacionDashboardRedes(
                f"No se encontró la hoja {HOJA_DESTINO} "
                "en INFORME ANS-REDES.xlsb."
            ) from error

        # ----------------------------------------------------
        # TABLA DESTINO
        # ----------------------------------------------------

        try:
            tabla = hoja.ListObjects(
                NOMBRE_TABLA_DESTINO
            )
        except Exception as error:
            raise ErrorActualizacionDashboardRedes(
                f"No se encontró la tabla {NOMBRE_TABLA_DESTINO} "
                f"en la hoja {HOJA_DESTINO}."
            ) from error

        # ----------------------------------------------------
        # ENCABEZADOS
        # ----------------------------------------------------

        hoja.Range(
            hoja.Cells(1, 1),
            hoja.Cells(
                1,
                cantidad_columnas,
            ),
        ).Value = (
            tuple(
                COLUMNAS_DASHBOARD_REDES
            ),
        )

        # ----------------------------------------------------
        # LIMPIAR DATOS ANTERIORES
        # ----------------------------------------------------

        if tabla.DataBodyRange is not None:
            tabla.DataBodyRange.ClearContents()

        ultima_fila_usada = (
            hoja.UsedRange.Rows.Count
        )

        if ultima_fila_usada >= 2:
            hoja.Range(
                hoja.Cells(2, 1),
                hoja.Cells(
                    ultima_fila_usada,
                    cantidad_columnas,
                ),
            ).ClearContents()

        # ----------------------------------------------------
        # REDIMENSIONAR TABLA
        # ----------------------------------------------------

        ultima_fila_tabla = (
            cantidad_registros + 1
            if cantidad_registros > 0
            else 1
        )

        tabla.Resize(
            hoja.Range(
                hoja.Cells(1, 1),
                hoja.Cells(
                    ultima_fila_tabla,
                    cantidad_columnas,
                ),
            )
        )

        # ----------------------------------------------------
        # COPIAR DATOS
        # ----------------------------------------------------

        if cantidad_registros > 0:
            matriz = construir_matriz_excel(
                dataframe
            )

            hoja.Range(
                hoja.Cells(2, 1),
                hoja.Cells(
                    cantidad_registros + 1,
                    cantidad_columnas,
                ),
            ).Value = matriz

            aplicar_formato_datos_ans_redes(
                hoja=hoja,
                cantidad_registros=(
                    cantidad_registros
                ),
            )

        # ----------------------------------------------------
        # ACTUALIZAR TABLAS DINÁMICAS Y GRÁFICOS
        # ----------------------------------------------------

        if refrescar_dashboard:
            actualizar_elementos_dashboard(
                libro_excel=libro,
                aplicacion_excel=excel,
            )

        libro.Save()

        resumen = {
            "ARCHIVO_ORIGEN": str(
                RUTA_INFORME_REDES
            ),
            "ARCHIVO_DESTINO": str(
                RUTA_DASHBOARD_REDES
            ),
            "HOJA_DESTINO": HOJA_DESTINO,
            "TABLA_DESTINO": NOMBRE_TABLA_DESTINO,
            "REGISTROS_TRANSFERIDOS": (
                cantidad_registros
            ),
            "COLUMNAS_TRANSFERIDAS": (
                cantidad_columnas
            ),
            "DASHBOARD_ACTUALIZADO": (
                refrescar_dashboard
            ),
        }

        logger.info(
            "Dashboard ANS Redes actualizado correctamente | %s",
            resumen,
        )

        return resumen

    except PermissionError as error:
        raise ErrorActualizacionDashboardRedes(
            "No fue posible actualizar INFORME ANS-REDES.xlsb.\n\n"
            "Cierre el archivo del dashboard en Excel "
            "y vuelva a ejecutar."
        ) from error

    except ErrorActualizacionDashboardRedes:
        raise

    except Exception as error:
        logger.exception(
            "No fue posible actualizar el dashboard ANS Redes."
        )

        raise ErrorActualizacionDashboardRedes(
            "No fue posible actualizar INFORME ANS-REDES.xlsb.\n\n"
            f"Detalle: {error}"
        ) from error

    finally:
        if libro is not None:
            try:
                libro.Close(
                    SaveChanges=False
                )
            except Exception:
                pass

        if excel is not None:
            try:
                excel.EnableEvents = True
                excel.DisplayAlerts = True
                excel.ScreenUpdating = True
                excel.Quit()
            except Exception:
                pass

        if com_inicializado:
            pythoncom.CoUninitialize()


# ============================================================
# PRUEBA DIRECTA
# ============================================================

if __name__ == "__main__":
    try:
        resultado = actualizar_dashboard_redes(
            refrescar_dashboard=True
        )

        print("=" * 70)
        print("DASHBOARD ANS REDES ACTUALIZADO CORRECTAMENTE")
        print("=" * 70)
        print(
            f"Registros transferidos: "
            f"{resultado['REGISTROS_TRANSFERIDOS']}"
        )
        print(
            f"Columnas transferidas: "
            f"{resultado['COLUMNAS_TRANSFERIDAS']}"
        )
        print(
            f"Hoja destino: "
            f"{resultado['HOJA_DESTINO']}"
        )
        print(
            f"Tabla destino: "
            f"{resultado['TABLA_DESTINO']}"
        )

    except ErrorActualizacionDashboardRedes as error:
        print()
        print("ERROR AL ACTUALIZAR EL DASHBOARD ANS REDES")
        print()
        print(error)
        raise SystemExit(1)
