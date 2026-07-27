import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, scrolledtext, ttk

from src.config import (
    COLOR_FONDO,
    COLOR_PRINCIPAL,
    COLOR_PRINCIPAL_HOVER,
    COLOR_TEXTO,
    ENTRADA_DIR,
    NOMBRE_APLICACION,
    SALIDA_DIR,
    VERSION_APLICACION,
)
from src.lector_csv import (
    ErrorLecturaCSV,
    cargar_archivo_entrada,
)
from src.validador import validar_dataframe


logger = logging.getLogger(__name__)


class AplicacionANS:
    """
    Ventana principal del proyecto Informe ANS ATC CHEC.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.proceso_activo = False

        self.configurar_ventana()
        self.construir_interfaz()

        logger.info("Interfaz principal iniciada.")

    def configurar_ventana(self) -> None:
        self.root.title(
            f"{NOMBRE_APLICACION} - Versión {VERSION_APLICACION}"
        )

        self.root.configure(
            bg=COLOR_FONDO
        )

        ancho = 760
        alto = 650

        pantalla_ancho = self.root.winfo_screenwidth()
        pantalla_alto = self.root.winfo_screenheight()

        posicion_x = (
            pantalla_ancho // 2
            - ancho // 2
        )

        posicion_y = (
            pantalla_alto // 2
            - alto // 2
        )

        self.root.geometry(
            f"{ancho}x{alto}+{posicion_x}+{posicion_y}"
        )

        self.root.minsize(
            ancho,
            alto,
        )

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.cerrar_aplicacion,
        )

    def construir_interfaz(self) -> None:
        self.construir_encabezado()
        self.construir_botones()
        self.construir_progreso()
        self.construir_area_log()
        self.construir_pie_pagina()

    def construir_encabezado(self) -> None:
        frame = tk.Frame(
            self.root,
            bg=COLOR_PRINCIPAL,
            height=130,
        )

        frame.pack(
            fill="x"
        )

        frame.pack_propagate(False)

        titulo = tk.Label(
            frame,
            text="INFORME ANS ATC CHEC",
            bg=COLOR_PRINCIPAL,
            fg="white",
            font=(
                "Segoe UI",
                20,
                "bold",
            ),
        )

        titulo.pack(
            pady=(24, 4)
        )

        subtitulo = tk.Label(
            frame,
            text=(
                "Generación de informes ANS "
                "y visor geográfico"
            ),
            bg=COLOR_PRINCIPAL,
            fg="white",
            font=(
                "Segoe UI",
                11,
            ),
        )

        subtitulo.pack()

    def construir_botones(self) -> None:
        frame = tk.Frame(
            self.root,
            bg=COLOR_FONDO,
        )

        frame.pack(
            fill="x",
            padx=30,
            pady=(22, 12),
        )

        frame.columnconfigure(
            (0, 1),
            weight=1,
        )

        self.btn_informe = self.crear_boton(
            contenedor=frame,
            texto="GENERAR\nINFORME ANS",
            comando=self.iniciar_validacion_csv,
        )

        self.btn_informe.grid(
            row=0,
            column=0,
            padx=8,
            pady=6,
            sticky="ew",
        )

        self.btn_mapa = self.crear_boton(
            contenedor=frame,
            texto="GENERAR\nMAPA",
            comando=self.mostrar_mapa_pendiente,
        )

        self.btn_mapa.grid(
            row=0,
            column=1,
            padx=8,
            pady=6,
            sticky="ew",
        )

        self.btn_salida = self.crear_boton(
            contenedor=frame,
            texto="ABRIR CARPETA\nDE SALIDA",
            comando=self.abrir_carpeta_salida,
        )

        self.btn_salida.grid(
            row=1,
            column=0,
            padx=8,
            pady=6,
            sticky="ew",
        )

        self.btn_salir = self.crear_boton(
            contenedor=frame,
            texto="SALIR",
            comando=self.cerrar_aplicacion,
        )

        self.btn_salir.grid(
            row=1,
            column=1,
            padx=8,
            pady=6,
            sticky="ew",
        )

    def crear_boton(
        self,
        contenedor: tk.Widget,
        texto: str,
        comando,
    ) -> tk.Button:

        boton = tk.Button(
            contenedor,
            text=texto,
            command=comando,
            bg=COLOR_PRINCIPAL,
            fg="white",
            activebackground=COLOR_PRINCIPAL_HOVER,
            activeforeground="white",
            font=(
                "Segoe UI",
                10,
                "bold",
            ),
            height=2,
            relief="flat",
            cursor="hand2",
        )

        boton.bind(
            "<Enter>",
            lambda evento: boton.configure(
                bg=COLOR_PRINCIPAL_HOVER
            ),
        )

        boton.bind(
            "<Leave>",
            lambda evento: boton.configure(
                bg=COLOR_PRINCIPAL
            ),
        )

        return boton

    def construir_progreso(self) -> None:
        self.barra_progreso = ttk.Progressbar(
            self.root,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )

        self.barra_progreso.pack(
            fill="x",
            padx=38,
            pady=(5, 10),
        )

    def construir_area_log(self) -> None:
        frame = tk.Frame(
            self.root,
            bg=COLOR_FONDO,
        )

        frame.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=(0, 10),
        )

        titulo = tk.Label(
            frame,
            text="Resultado del proceso",
            bg=COLOR_FONDO,
            fg=COLOR_TEXTO,
            font=(
                "Segoe UI",
                10,
                "bold",
            ),
            anchor="w",
        )

        titulo.pack(
            fill="x",
            pady=(0, 4),
        )

        self.area_mensajes = (
            scrolledtext.ScrolledText(
                frame,
                wrap=tk.WORD,
                font=(
                    "Consolas",
                    9,
                ),
                bg="white",
                fg=COLOR_TEXTO,
                height=12,
                state="disabled",
            )
        )

        self.area_mensajes.pack(
            fill="both",
            expand=True,
        )

        self.area_mensajes.tag_config(
            "info",
            foreground="#2471A3",
        )

        self.area_mensajes.tag_config(
            "correcto",
            foreground="#1E8449",
        )

        self.area_mensajes.tag_config(
            "advertencia",
            foreground="#B9770E",
        )

        self.area_mensajes.tag_config(
            "error",
            foreground="#C0392B",
        )

        self.agregar_mensaje(
            "Aplicación iniciada correctamente.",
            "correcto",
        )

        self.agregar_mensaje(
            f"Carpeta de entrada: {ENTRADA_DIR}",
            "info",
        )

    def construir_pie_pagina(self) -> None:
        self.estado = tk.Label(
            self.root,
            text="Esperando acción del usuario...",
            bg=COLOR_FONDO,
            fg=COLOR_TEXTO,
            font=(
                "Segoe UI",
                9,
                "italic",
            ),
            anchor="w",
        )

        self.estado.pack(
            fill="x",
            padx=30,
            pady=(0, 16),
        )

    def agregar_mensaje(
        self,
        mensaje: str,
        etiqueta: str = "info",
    ) -> None:

        hora = datetime.now().strftime(
            "%H:%M:%S"
        )

        self.area_mensajes.configure(
            state="normal"
        )

        self.area_mensajes.insert(
            tk.END,
            f"[{hora}] {mensaje}\n",
            etiqueta,
        )

        self.area_mensajes.configure(
            state="disabled"
        )

        self.area_mensajes.see(
            tk.END
        )

    def cambiar_estado(
        self,
        mensaje: str,
    ) -> None:

        self.estado.configure(
            text=mensaje
        )

    def bloquear_botones(
        self,
        bloquear: bool,
    ) -> None:

        estado = (
            tk.DISABLED
            if bloquear
            else tk.NORMAL
        )

        self.btn_informe.configure(
            state=estado
        )

        self.btn_mapa.configure(
            state=estado
        )

        self.proceso_activo = bloquear

    def iniciar_validacion_csv(self) -> None:
        if self.proceso_activo:
            return

        self.bloquear_botones(True)

        hilo = threading.Thread(
            target=self.validar_archivo_csv,
            daemon=True,
        )

        hilo.start()

    def validar_archivo_csv(self) -> None:
        try:
            self.root.after(
                0,
                self.preparar_proceso,
            )

            ruta_csv, dataframe = (
                cargar_archivo_entrada()
            )

            resultado = validar_dataframe(
                dataframe
            )

            self.root.after(
                0,
                lambda: self.mostrar_resultado_validacion(
                    ruta_csv.name,
                    dataframe,
                    resultado,
                ),
            )

        except (
            FileNotFoundError,
            ErrorLecturaCSV,
        ) as error:

            logger.warning(
                "Validación detenida: %s",
                error,
            )

            self.root.after(
                0,
                lambda mensaje=str(error): (
                    self.mostrar_error(mensaje)
                ),
            )

        except Exception as error:
            logger.exception(
                "Error inesperado durante la validación."
            )

            self.root.after(
                0,
                lambda mensaje=str(error): (
                    self.mostrar_error(
                        "Ocurrió un error inesperado.\n\n"
                        f"Detalle: {mensaje}"
                    )
                ),
            )

        finally:
            self.root.after(
                0,
                self.finalizar_proceso,
            )

    def preparar_proceso(self) -> None:
        self.barra_progreso.configure(
            mode="indeterminate"
        )

        self.barra_progreso.start(15)

        self.cambiar_estado(
            "Validando archivo CSV..."
        )

        self.agregar_mensaje(
            "Iniciando validación del archivo CSV.",
            "info",
        )

    def finalizar_proceso(self) -> None:
        self.barra_progreso.stop()

        self.barra_progreso.configure(
            mode="determinate"
        )

        self.bloquear_botones(False)

        self.cambiar_estado(
            "Esperando acción del usuario..."
        )

    def mostrar_resultado_validacion(
        self,
        nombre_archivo,
        dataframe,
        resultado,
    ) -> None:

        self.barra_progreso.configure(
            mode="determinate"
        )

        self.barra_progreso["value"] = 100

        self.agregar_mensaje(
            f"Archivo detectado: {nombre_archivo}",
            "correcto",
        )

        self.agregar_mensaje(
            f"Registros encontrados: {len(dataframe):,}",
            "info",
        )

        self.agregar_mensaje(
            f"Columnas encontradas: {len(dataframe.columns)}",
            "info",
        )

        self.agregar_mensaje(
            "Encabezados detectados:",
            "info",
        )

        for columna in dataframe.columns:
            self.agregar_mensaje(
                f"  - {columna}",
                "info",
            )

        for advertencia in resultado.advertencias:
            self.agregar_mensaje(
                advertencia,
                "advertencia",
            )

        for error in resultado.errores:
            self.agregar_mensaje(
                error,
                "error",
            )

        if resultado.es_valido:
            self.agregar_mensaje(
                "La estructura básica del CSV es válida.",
                "correcto",
            )

            self.agregar_mensaje(
                "El cálculo ANS será incorporado cuando se "
                "confirmen las columnas y reglas contractuales.",
                "advertencia",
            )

            messagebox.showinfo(
                NOMBRE_APLICACION,
                "El archivo CSV fue leído correctamente.\n\n"
                f"Registros: {len(dataframe):,}\n"
                f"Columnas: {len(dataframe.columns)}\n\n"
                "La estructura está lista para ser analizada.",
            )

        else:
            messagebox.showerror(
                NOMBRE_APLICACION,
                "El archivo presenta errores estructurales.\n\n"
                "Revise el área de resultados y el archivo de log.",
            )

        self.root.after(
            1500,
            lambda: self.barra_progreso.configure(
                value=0
            ),
        )

    def mostrar_error(
        self,
        mensaje: str,
    ) -> None:

        self.agregar_mensaje(
            mensaje,
            "error",
        )

        messagebox.showerror(
            NOMBRE_APLICACION,
            mensaje,
        )

    def mostrar_mapa_pendiente(self) -> None:
        mensaje = (
            "La generación del mapa se habilitará después de "
            "confirmar si el archivo contiene coordenadas o "
            "si será necesario geocodificar las direcciones."
        )

        self.agregar_mensaje(
            mensaje,
            "advertencia",
        )

        messagebox.showinfo(
            "Generar mapa",
            mensaje,
        )

    def abrir_carpeta_salida(self) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(SALIDA_DIR)

            elif sys.platform == "darwin":
                subprocess.Popen(
                    ["open", str(SALIDA_DIR)]
                )

            else:
                subprocess.Popen(
                    ["xdg-open", str(SALIDA_DIR)]
                )

            self.agregar_mensaje(
                "Carpeta de salida abierta correctamente.",
                "correcto",
            )

        except OSError as error:
            logger.exception(
                "No fue posible abrir la carpeta de salida."
            )

            self.mostrar_error(
                "No fue posible abrir la carpeta de salida.\n\n"
                f"Detalle: {error}"
            )

    def cerrar_aplicacion(self) -> None:
        if self.proceso_activo:
            confirmar = messagebox.askyesno(
                "Proceso activo",
                "Existe un proceso en ejecución.\n\n"
                "¿Desea cerrar la aplicación?",
            )

            if not confirmar:
                return

        logger.info(
            "Aplicación cerrada por el usuario."
        )

        self.root.destroy()


def iniciar_interfaz() -> None:
    root = tk.Tk()
    AplicacionANS(root)
    root.mainloop()