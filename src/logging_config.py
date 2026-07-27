import logging
from logging.handlers import RotatingFileHandler

from src.config import RUTA_LOG, crear_directorios


FORMATO_LOG = (
    "%(asctime)s | "
    "%(levelname)-8s | "
    "%(name)s | "
    "%(message)s"
)


def configurar_logging() -> None:
    """
    Configura el sistema centralizado de logging.

    El archivo de log rota automáticamente cuando alcanza
    aproximadamente 2 MB, conservando hasta cinco respaldos.
    """

    crear_directorios()

    logger_principal = logging.getLogger()

    if logger_principal.handlers:
        return

    logger_principal.setLevel(logging.INFO)

    formateador = logging.Formatter(
        FORMATO_LOG,
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    manejador_archivo = RotatingFileHandler(
        filename=RUTA_LOG,
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8",
    )

    manejador_archivo.setLevel(logging.INFO)
    manejador_archivo.setFormatter(formateador)

    logger_principal.addHandler(manejador_archivo)