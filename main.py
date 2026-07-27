import logging
import sys
import tkinter as tk
from tkinter import messagebox

from src.config import (
    NOMBRE_APLICACION,
    crear_directorios,
)
from src.interfaz import iniciar_interfaz
from src.logging_config import configurar_logging


logger = logging.getLogger(__name__)


def main() -> None:
    """
    Punto de entrada principal de Informe ANS ATC CHEC.
    """

    crear_directorios()
    configurar_logging()

    logger.info(
        "Iniciando aplicación %s.",
        NOMBRE_APLICACION,
    )

    try:
        iniciar_interfaz()

    except tk.TclError as error:
        logger.exception(
            "No fue posible iniciar la interfaz gráfica."
        )

        raise RuntimeError(
            "No fue posible iniciar Tkinter."
        ) from error

    except Exception as error:
        logger.exception(
            "La aplicación finalizó por un error no controlado."
        )

        root = tk.Tk()
        root.withdraw()

        messagebox.showerror(
            NOMBRE_APLICACION,
            "La aplicación no pudo iniciarse correctamente.\n\n"
            "Revise el archivo de log para conocer el detalle.",
        )

        root.destroy()

        sys.exit(1)


if __name__ == "__main__":
    main()