import logging
import time

from datetime import date, datetime
from typing import Any

import pandas as pd
import pythoncom
import win32com.client as win32

from src.config import (
    BASE_DIR,
    HOJA_DATOS_SALIDA,
    RUTA_INFORME_EXCEL,
)


logger = logging.getLogger(__name__)


# ==========================================================
# CONFIGURACIÓN DEL DASHBOARD
# ==========================================================

DASHBOARD_DIR = BASE_DIR / "dashboard"

RUTA_DASHBOARD = (
    DASHBOARD_DIR / "INFORME_ANS.xlsb"
)

HOJA_DESTINO = "DATOS_ANS"
NOMBRE_TABLA_DESTINO = "DATOS"


# ==========================================================
# COLORES Y FORMATO DE ESTADOS EN DATOS_ANS
# ==========================================================

COLOR_ESTADOS = {
    "VENCIDO": (255, 31, 31),
    "ALERTA": (255, 212, 38),
    "A TIEMPO": (139, 207, 74),
    "INMEDIATO": (243, 156, 18),
    "HV": (52, 152, 219),
    "FACTIBILIDAD": (23, 165, 137),
    "SIN FECHA": (100, 118, 135),
    "PENDIENTE CONFIGURACIÓN": (100, 118, 135),
}


def color_excel(rojo: int, verde: int, azul: int) -> int:
    """Convierte RGB al valor entero utilizado por Excel COM."""

    return rojo + (verde * 256) + (azul * 65536)


# ==========================================================
# COLUMNAS QUE SE TRANSFIEREN
# ==========================================================

COLUMNAS_DASHBOARD = [
    "ID_ORDEN",
    "FECHA_ORDEN",
    "DIRECCION",
    "PROPIETARIO",
    "ZONA",
    "MUNICIPIO",
    "DESC_MUNICIPIO",
    "REGION_ORIGEN",
    "PDA_NUMERO",
    "TIPO",
    "DIAS_PACTADOS",
    "FECHA_LIMITE_ANS",
    "DIAS_TRANSCURRIDOS",
    "DIAS_RESTANTES",
    "ESTADO",
    "OBSERVACION",
]


class ErrorActualizacionDashboard(Exception):
    """
    Error controlado durante la actualización del dashboard.
    """


# ==========================================================
# VALIDACIONES
# ==========================================================

def validar_archivos() -> None:
    """
    Valida la existencia del informe generado y del dashboard.
    """

    if not RUTA_INFORME_EXCEL.exists():
        raise ErrorActualizacionDashboard(
            "No se encontró el informe generado:\n\n"
            f"{RUTA_INFORME_EXCEL}\n\n"
            "Primero genere el Informe ANS Conexiones."
        )

    if not RUTA_DASHBOARD.exists():
        raise ErrorActualizacionDashboard(
            "No se encontró el archivo del dashboard:\n\n"
            f"{RUTA_DASHBOARD}"
        )


def validar_columnas(
    dataframe: pd.DataFrame,
) -> None:
    """
    Valida que el informe tenga las 16 columnas requeridas.
    """

    columnas_faltantes = [
        columna
        for columna in COLUMNAS_DASHBOARD
        if columna not in dataframe.columns
    ]

    if columnas_faltantes:
        detalle = "\n".join(
            f"- {columna}"
            for columna in columnas_faltantes
        )

        raise ErrorActualizacionDashboard(
            "Faltan columnas en el informe generado:\n\n"
            f"{detalle}"
        )


# ==========================================================
# LECTURA DEL INFORME GENERADO
# ==========================================================

def cargar_datos_informe() -> pd.DataFrame:
    """
    Lee DATOS_ANS y conserva únicamente registros reales.

    Se consideran válidas las filas que tengan ID_ORDEN.
    """

    try:
        dataframe = pd.read_excel(
            RUTA_INFORME_EXCEL,
            sheet_name=HOJA_DATOS_SALIDA,
            dtype=object,
            engine="openpyxl",
        )

    except PermissionError as error:
        raise ErrorActualizacionDashboard(
            "No fue posible leer Informe_ANS_ELITE.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a ejecutar."
        ) from error

    except ValueError as error:
        raise ErrorActualizacionDashboard(
            f"No se encontró la hoja {HOJA_DATOS_SALIDA} "
            "en Informe_ANS_ELITE.xlsx."
        ) from error

    except Exception as error:
        logger.exception(
            "No fue posible leer el informe generado."
        )

        raise ErrorActualizacionDashboard(
            "No fue posible leer Informe_ANS_ELITE.xlsx."
        ) from error

    validar_columnas(
        dataframe
    )

    dataframe = dataframe[
        COLUMNAS_DASHBOARD
    ].copy()

    dataframe["ID_ORDEN"] = (
        dataframe["ID_ORDEN"]
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
        dataframe["ID_ORDEN"].ne("")
    ].copy()

    dataframe = dataframe.reset_index(
        drop=True
    )

    return dataframe


# ==========================================================
# CONVERSIÓN DE VALORES PARA EXCEL COM
# ==========================================================

def convertir_valor_excel(
    valor: Any,
) -> Any:
    """
    Convierte valores de pandas/numpy en valores compatibles
    con Excel mediante win32com.
    """

    if valor is None:
        return None

    try:
        if pd.isna(valor):
            return None

    except (
        TypeError,
        ValueError,
    ):
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
    """
    Convierte el DataFrame en una matriz compatible con Excel.
    """

    registros: list[tuple[Any, ...]] = []

    for fila in dataframe.itertuples(
        index=False,
        name=None,
    ):

        registros.append(
            tuple(
                convertir_valor_excel(valor)
                for valor in fila
            )
        )

    return tuple(
        registros
    )


# ==========================================================
# FORMATO DE LA HOJA DATOS_ANS
# ==========================================================

def aplicar_formato_datos_ans(
    hoja: Any,
    cantidad_registros: int,
) -> None:
    """
    Aplica formatos sobre las 16 columnas de DATOS_ANS.

    No modifica el diseño de la hoja DASHBOARD.
    """

    if cantidad_registros <= 0:
        return

    ultima_fila = cantidad_registros + 1

    # Anchos de columnas.
    anchos = {
        1: 16,   # ID_ORDEN
        2: 15,   # FECHA_ORDEN
        3: 45,   # DIRECCION
        4: 33,   # PROPIETARIO
        5: 10,   # ZONA
        6: 15,   # MUNICIPIO
        7: 24,   # DESC_MUNICIPIO
        8: 17,   # REGION_ORIGEN
        9: 24,   # PDA_NUMERO
        10: 18,  # TIPO
        11: 16,  # DIAS_PACTADOS
        12: 20,  # FECHA_LIMITE_ANS
        13: 22,  # DIAS_TRANSCURRIDOS
        14: 18,  # DIAS_RESTANTES
        15: 25,  # ESTADO
        16: 110, # OBSERVACION
    }

    for numero_columna, ancho in anchos.items():
        hoja.Columns(numero_columna).ColumnWidth = ancho

    # Formatos numéricos y de fecha.
    hoja.Range(
        hoja.Cells(2, 1),
        hoja.Cells(ultima_fila, 1),
    ).NumberFormat = "@"

    hoja.Range(
        hoja.Cells(2, 2),
        hoja.Cells(ultima_fila, 2),
    ).NumberFormat = "dd/mm/yyyy"

    hoja.Range(
        hoja.Cells(2, 6),
        hoja.Cells(ultima_fila, 6),
    ).NumberFormat = "@"

    hoja.Range(
        hoja.Cells(2, 9),
        hoja.Cells(ultima_fila, 9),
    ).NumberFormat = "@"

    hoja.Range(
        hoja.Cells(2, 11),
        hoja.Cells(ultima_fila, 11),
    ).NumberFormat = "0"

    hoja.Range(
        hoja.Cells(2, 12),
        hoja.Cells(ultima_fila, 12),
    ).NumberFormat = "dd/mm/yyyy"

    hoja.Range(
        hoja.Cells(2, 13),
        hoja.Cells(ultima_fila, 14),
    ).NumberFormat = "0"

    # Alineación general.
    rango_datos = hoja.Range(
        hoja.Cells(2, 1),
        hoja.Cells(ultima_fila, 16),
    )

    rango_datos.VerticalAlignment = -4108
    rango_datos.WrapText = False

    # Columnas centradas.
    for numero_columna in (
        2, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15
    ):
        rango = hoja.Range(
            hoja.Cells(2, numero_columna),
            hoja.Cells(ultima_fila, numero_columna),
        )
        rango.HorizontalAlignment = -4108
        rango.VerticalAlignment = -4108

    # Observación en P.
    rango_observacion = hoja.Range(
        hoja.Cells(2, 16),
        hoja.Cells(ultima_fila, 16),
    )

    rango_observacion.WrapText = True
    rango_observacion.HorizontalAlignment = -4131
    rango_observacion.VerticalAlignment = -4160

    hoja.Rows(
        f"2:{ultima_fila}"
    ).RowHeight = 36

    # Colores directos sobre ESTADO en O.
    rango_estado = hoja.Range(
        hoja.Cells(2, 15),
        hoja.Cells(ultima_fila, 15),
    )

    rango_estado.FormatConditions.Delete()
    rango_estado.Interior.Pattern = -4142
    rango_estado.Font.Bold = True
    rango_estado.Font.Color = color_excel(0, 0, 0)

    for numero_fila in range(
        2,
        ultima_fila + 1,
    ):
        celda = hoja.Cells(
            numero_fila,
            15,
        )

        estado = str(
            celda.Value or ""
        ).strip().upper()

        color = COLOR_ESTADOS.get(
            estado
        )

        if color is None:
            continue

        celda.Interior.Color = color_excel(
            *color
        )

        if estado in {
            "HV",
            "FACTIBILIDAD",
            "SIN FECHA",
            "PENDIENTE CONFIGURACIÓN",
        }:
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


# ==========================================================
# ACTUALIZACIÓN DE TABLAS DINÁMICAS
# ==========================================================

def esperar_actualizacion_excel(
    aplicacion_excel: Any,
    segundos_maximos: int = 120,
) -> None:
    """
    Espera hasta que Excel termine sus cálculos y consultas.
    """

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

    logger.warning(
        "Excel superó el tiempo máximo de espera "
        "durante la actualización."
    )


def actualizar_elementos_dashboard(
    libro_excel: Any,
    aplicacion_excel: Any,
) -> None:
    """
    Actualiza conexiones, tablas dinámicas, gráficos y cálculos.
    """

    try:
        libro_excel.RefreshAll()

    except Exception:
        logger.warning(
            "No fue posible ejecutar RefreshAll en el dashboard.",
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
                    "No fue posible actualizar una caché "
                    "de tabla dinámica.",
                    exc_info=True,
                )

    except Exception:
        pass

    try:
        aplicacion_excel.CalculateFull()

    except Exception:
        pass


# ==========================================================
# PROCESO PRINCIPAL
# ==========================================================

def actualizar_dashboard(
    refrescar_dashboard: bool = True,
) -> dict:
    """
    Reemplaza los datos del dashboard con el informe más reciente.

    Esta función puede ejecutarse desde un hilo secundario de
    Tkinter porque inicializa y libera COM explícitamente.
    """

    excel = None
    libro = None
    com_inicializado = False

    try:
        # Cada hilo que utiliza win32com debe inicializar COM.
        pythoncom.CoInitialize()
        com_inicializado = True

        validar_archivos()

        dataframe = cargar_datos_informe()

        cantidad_registros = len(
            dataframe
        )

        cantidad_columnas = len(
            COLUMNAS_DASHBOARD
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
                RUTA_DASHBOARD.resolve()
            ),
            UpdateLinks=0,
            ReadOnly=False,
        )

        try:
            hoja = libro.Worksheets(
                HOJA_DESTINO
            )

        except Exception as error:
            raise ErrorActualizacionDashboard(
                f"No se encontró la hoja {HOJA_DESTINO} "
                "en INFORME_ANS.xlsb."
            ) from error

        try:
            tabla = hoja.ListObjects(
                NOMBRE_TABLA_DESTINO
            )

        except Exception as error:
            raise ErrorActualizacionDashboard(
                f"No se encontró la tabla "
                f"{NOMBRE_TABLA_DESTINO} en la hoja "
                f"{HOJA_DESTINO}."
            ) from error

        # --------------------------------------------------
        # ENCABEZADOS
        # --------------------------------------------------

        rango_encabezados = hoja.Range(
            hoja.Cells(
                1,
                1,
            ),
            hoja.Cells(
                1,
                cantidad_columnas,
            ),
        )

        rango_encabezados.Value = (
            tuple(
                COLUMNAS_DASHBOARD
            ),
        )

        # --------------------------------------------------
        # LIMPIAR DATOS ANTERIORES
        # --------------------------------------------------

        if tabla.DataBodyRange is not None:
            tabla.DataBodyRange.ClearContents()

        ultima_fila_usada = hoja.UsedRange.Rows.Count

        if ultima_fila_usada >= 2:
            hoja.Range(
                hoja.Cells(
                    2,
                    1,
                ),
                hoja.Cells(
                    ultima_fila_usada,
                    cantidad_columnas,
                ),
            ).ClearContents()

        # --------------------------------------------------
        # REDIMENSIONAR LA TABLA
        # --------------------------------------------------

        ultima_fila_tabla = (
            cantidad_registros + 1
            if cantidad_registros > 0
            else 1
        )

        tabla.Resize(
            hoja.Range(
                hoja.Cells(
                    1,
                    1,
                ),
                hoja.Cells(
                    ultima_fila_tabla,
                    cantidad_columnas,
                ),
            )
        )

        # --------------------------------------------------
        # COPIAR DATOS NUEVOS
        # --------------------------------------------------

        if cantidad_registros > 0:

            matriz = construir_matriz_excel(
                dataframe
            )

            rango_destino = hoja.Range(
                hoja.Cells(
                    2,
                    1,
                ),
                hoja.Cells(
                    cantidad_registros + 1,
                    cantidad_columnas,
                ),
            )

            rango_destino.Value = matriz

            aplicar_formato_datos_ans(
                hoja=hoja,
                cantidad_registros=cantidad_registros,
            )

        # --------------------------------------------------
        # ACTUALIZAR DASHBOARD
        # --------------------------------------------------

        if refrescar_dashboard:
            actualizar_elementos_dashboard(
                libro_excel=libro,
                aplicacion_excel=excel,
            )

        libro.Save()

        resumen = {
            "ARCHIVO_ORIGEN": str(
                RUTA_INFORME_EXCEL
            ),
            "ARCHIVO_DESTINO": str(
                RUTA_DASHBOARD
            ),
            "HOJA_DESTINO": HOJA_DESTINO,
            "TABLA_DESTINO": NOMBRE_TABLA_DESTINO,
            "REGISTROS_TRANSFERIDOS": cantidad_registros,
            "COLUMNAS_TRANSFERIDAS": cantidad_columnas,
            "DASHBOARD_ACTUALIZADO": refrescar_dashboard,
        }

        logger.info(
            "Dashboard actualizado correctamente | %s",
            resumen,
        )

        return resumen

    except PermissionError as error:
        raise ErrorActualizacionDashboard(
            "No fue posible actualizar INFORME_ANS.xlsb.\n\n"
            "Cierre el archivo del dashboard en Excel "
            "y vuelva a ejecutar."
        ) from error

    except ErrorActualizacionDashboard:
        raise

    except Exception as error:
        logger.exception(
            "No fue posible actualizar el dashboard."
        )

        raise ErrorActualizacionDashboard(
            "No fue posible actualizar INFORME_ANS.xlsb.\n\n"
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


# ==========================================================
# PRUEBA DIRECTA
# ==========================================================

if __name__ == "__main__":

    try:
        resultado = actualizar_dashboard(
            refrescar_dashboard=True
        )

        print(
            "DASHBOARD ACTUALIZADO CORRECTAMENTE"
        )

        print(
            f"Registros transferidos: "
            f"{resultado['REGISTROS_TRANSFERIDOS']}"
        )

        print(
            f"Columnas transferidas: "
            f"{resultado['COLUMNAS_TRANSFERIDAS']}"
        )

    except ErrorActualizacionDashboard as error:
        print(
            "\nERROR AL ACTUALIZAR EL DASHBOARD\n"
        )

        print(
            error
        )
