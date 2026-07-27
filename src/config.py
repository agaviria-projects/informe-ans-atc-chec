from pathlib import Path


# ==========================================================
# RUTA RAÍZ DEL PROYECTO
# ==========================================================

# config.py está ubicado en:
# Informe_ANS_ATC_CHEC/src/config.py
#
# parent        -> src
# parent.parent -> Informe_ANS_ATC_CHEC

BASE_DIR = Path(__file__).resolve().parent.parent


# ==========================================================
# CARPETAS DEL PROYECTO
# ==========================================================

ENTRADA_DIR = BASE_DIR / "entrada"
SALIDA_DIR = BASE_DIR / "salida"
LOGS_DIR = BASE_DIR / "logs"


# ==========================================================
# ARCHIVOS DE SALIDA
# ==========================================================

NOMBRE_INFORME_EXCEL = "Informe_ANS_ATC_CHEC.xlsx"
NOMBRE_MAPA_HTML = "Mapa_ANS_ATC_CHEC.html"

RUTA_INFORME_EXCEL = SALIDA_DIR / NOMBRE_INFORME_EXCEL
RUTA_MAPA_HTML = SALIDA_DIR / NOMBRE_MAPA_HTML
RUTA_LOG = LOGS_DIR / "Informe_ANS_ATC_CHEC.log"


# ==========================================================
# CONFIGURACIÓN DE ARCHIVOS DE ENTRADA
# ==========================================================

EXTENSIONES_CSV_VALIDAS = {".csv"}

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
# CONFIGURACIÓN GENERAL
# ==========================================================

NOMBRE_APLICACION = "Informe ANS ATC CHEC"
VERSION_APLICACION = "1.0.0"

COLOR_PRINCIPAL = "#1E8449"
COLOR_PRINCIPAL_HOVER = "#239B56"
COLOR_FONDO = "#F4F6F7"
COLOR_TEXTO = "#1B2631"


# ==========================================================
# CREACIÓN DE CARPETAS
# ==========================================================

def crear_directorios() -> None:
    """
    Crea las carpetas esenciales del proyecto cuando no existen.

    Las rutas se construyen dinámicamente desde BASE_DIR,
    por lo que el proyecto puede copiarse a cualquier equipo.
    """

    directorios = (
        ENTRADA_DIR,
        SALIDA_DIR,
        LOGS_DIR,
    )

    for directorio in directorios:
        directorio.mkdir(parents=True, exist_ok=True)