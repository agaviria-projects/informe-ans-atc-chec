from pathlib import Path


# ==========================================================
# RUTA RAÍZ DEL PROYECTO
# ==========================================================

# config.py se encuentra en:
# Informe_ANS_ATC_CHEC/src/config.py
#
# parent        -> src
# parent.parent -> raíz del proyecto

BASE_DIR = Path(__file__).resolve().parent.parent


# ==========================================================
# CARPETAS DEL PROYECTO
# ==========================================================

ENTRADA_DIR = BASE_DIR / "entrada"
SALIDA_DIR = BASE_DIR / "salida"
LOGS_DIR = BASE_DIR / "logs"
ASSETS_DIR = BASE_DIR / "assets"
CONFIG_DIR = BASE_DIR / "config"
MAPAS_DIR = BASE_DIR / "mapas"


# ==========================================================
# RECURSOS VISUALES
# ==========================================================

RUTA_LOGO_EMPRESA = ASSETS_DIR / "elite.png"


# ==========================================================
# ARCHIVOS DE CONFIGURACIÓN
# ==========================================================

RUTA_DIAS_CONTRACTUALES = (
    CONFIG_DIR / "DIAS_CONTRACTUALES.xlsx"
)

RUTA_CACHE_GEOCODIFICACION = (
    CONFIG_DIR / "CACHE_GEOCODIFICACION.xlsx"
)


# ==========================================================
# ARCHIVOS DE SALIDA
# ==========================================================

NOMBRE_INFORME_EXCEL = "Informe_ANS_ELITE.xlsx"
NOMBRE_MAPA_HTML = "Mapa_ANS_ELITE.html"

RUTA_INFORME_EXCEL = (
    SALIDA_DIR / NOMBRE_INFORME_EXCEL
)

RUTA_MAPA_HTML = (
    MAPAS_DIR / NOMBRE_MAPA_HTML
)

RUTA_LOG = (
    LOGS_DIR / "Informe_ANS_ELITE.log"
)


# ==========================================================
# HOJAS DEL INFORME
# ==========================================================

HOJA_DATOS_SALIDA = "DATOS_ANS"
HOJA_CONTROL = "CONTROL_PROCESO"
HOJA_DIAS_CONTRACTUALES = "DIAS_CONTRACTUALES"

# Se agregarán posteriormente.
HOJA_RESUMEN = "RESUMEN"
HOJA_DASHBOARD = "DASHBOARD_ANS"


# ==========================================================
# CONFIGURACIÓN DE ARCHIVOS DE ENTRADA
# ==========================================================

EXTENSIONES_EXCEL_VALIDAS = {
    ".xlsx",
    ".xlsm",
}

# Se conserva soporte CSV para una posible evolución futura.
EXTENSIONES_CSV_VALIDAS = {
    ".csv",
}

CODIFICACIONES_CSV = (
    "utf-8-sig",
    "utf-8",
    "cp1252",
    "latin-1",
)

SEPARADORES_CSV = (
    ";",
    ",",
    "|",
    "\t",
)


# ==========================================================
# CONFIGURACIÓN DEL PROCESO
# ==========================================================

CANTIDAD_ARCHIVOS_REGIONALES = 2

NOMBRE_REGION_1 = "REGION 1"
NOMBRE_REGION_2 = "REGION 2"

# Valor provisional hasta confirmar la regla contractual.
UMBRAL_ALERTA_DIAS = 2


# ==========================================================
# CONFIGURACIÓN DE GEOCODIFICACIÓN Y MAPA
# ==========================================================

# Primera etapa controlada:
# se consultan como máximo 10 direcciones nuevas por ejecución.
#
# Las coordenadas ya guardadas en caché no cuentan dentro
# de este límite y se reutilizan automáticamente.
LIMITE_DIRECCIONES_MAPA = 140

# Tiempo mínimo entre consultas al servicio geográfico.
# Evita realizar solicitudes demasiado rápidas.
ESPERA_GEOCODIFICACION_SEGUNDOS = 1.1


# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

NOMBRE_APLICACION = "Informe ANS Elite"
VERSION_APLICACION = "1.0.0"


# ==========================================================
# PALETA VISUAL
# ==========================================================

COLOR_PRINCIPAL = "#1E8449"
COLOR_PRINCIPAL_HOVER = "#239B56"

COLOR_VERDE_EMPRESA = "#00843D"
COLOR_VERDE_EMPRESA_CLARO = "#8DC63F"

COLOR_FONDO = "#F4F6F7"
COLOR_FONDO_BANNER = "#EAEDED"

COLOR_TEXTO = "#1B2631"
COLOR_TEXTO_SECUNDARIO = "#566573"

COLOR_BORDE = "#AAB7B8"
COLOR_BLANCO = "#FFFFFF"


# ==========================================================
# CONFIGURACIÓN DE INTERFAZ
# ==========================================================

ANCHO_VENTANA = 760
ALTO_VENTANA = 650

TITULO_PRINCIPAL = "INFORME ANS"

SUBTITULO_PRINCIPAL = (
    "Generación de informes ANS y mapas geográficos"
)


# ==========================================================
# CREACIÓN DE CARPETAS
# ==========================================================

def crear_directorios() -> None:
    """
    Crea las carpetas esenciales cuando no existen.

    Todas las rutas se construyen desde BASE_DIR, por lo que
    el proyecto puede copiarse y ejecutarse en otro computador
    sin modificar rutas dentro del código.
    """

    directorios = (
        ENTRADA_DIR,
        SALIDA_DIR,
        LOGS_DIR,
        ASSETS_DIR,
        CONFIG_DIR,
        MAPAS_DIR,
    )

    for directorio in directorios:
        directorio.mkdir(
            parents=True,
            exist_ok=True,
        )


# ==========================================================
# VALIDACIÓN DE RECURSOS
# ==========================================================

def validar_recursos_visuales() -> list[str]:
    """
    Valida los recursos visuales requeridos por la aplicación.

    Returns:
        Lista de advertencias. Estará vacía cuando todos los
        recursos visuales existan correctamente.
    """

    advertencias: list[str] = []

    if not RUTA_LOGO_EMPRESA.exists():
        advertencias.append(
            "No se encontró el logo corporativo: "
            f"{RUTA_LOGO_EMPRESA.name}"
        )

    return advertencias