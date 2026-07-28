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
ASSETS_DIR = BASE_DIR / "assets"


# ==========================================================
# RECURSOS VISUALES
# ==========================================================

RUTA_LOGO_EMPRESA = ASSETS_DIR / "elite.png"


# ==========================================================
# ARCHIVOS DE SALIDA
# ==========================================================

NOMBRE_INFORME_EXCEL = "Informe_ANS_ELITE.xlsx"
NOMBRE_MAPA_HTML = "Mapa_ANS_ELITE.html"

RUTA_INFORME_EXCEL = SALIDA_DIR / NOMBRE_INFORME_EXCEL
RUTA_MAPA_HTML = SALIDA_DIR / NOMBRE_MAPA_HTML
RUTA_LOG = LOGS_DIR / "Informe_ANS_ELITE.log"


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
    Crea las carpetas esenciales del proyecto cuando no existen.

    Todas las rutas se generan dinámicamente a partir de BASE_DIR,
    por lo que el proyecto puede ejecutarse desde cualquier equipo
    sin utilizar rutas absolutas.
    """

    directorios = (
        ENTRADA_DIR,
        SALIDA_DIR,
        LOGS_DIR,
        ASSETS_DIR,
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
        Lista con los mensajes de advertencia encontrados.
        La lista estará vacía cuando todos los recursos existan.
    """

    advertencias: list[str] = []

    if not RUTA_LOGO_EMPRESA.exists():
        advertencias.append(
            f"No se encontró el logo corporativo: "
            f"{RUTA_LOGO_EMPRESA.name}"
        )

    return advertencias