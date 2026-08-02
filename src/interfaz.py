import logging
import os
import subprocess
import sys
import threading
import tkinter as tk

import pythoncom
import win32com.client as win32

from datetime import datetime
from tkinter import messagebox, scrolledtext, ttk
from typing import Callable

from PIL import Image, ImageTk


from src.actualizador_dashboard import (
    ErrorActualizacionDashboard,
    actualizar_dashboard,
)

from src.config import (
    ALTO_VENTANA,
    BASE_DIR,
    ANCHO_VENTANA,
    COLOR_BLANCO,
    COLOR_BORDE,
    COLOR_FONDO,
    COLOR_FONDO_BANNER,
    COLOR_PRINCIPAL,
    COLOR_PRINCIPAL_HOVER,
    COLOR_TEXTO,
    COLOR_TEXTO_SECUNDARIO,
    COLOR_VERDE_EMPRESA,
    COLOR_VERDE_EMPRESA_CLARO,
    ENTRADA_DIR,
    NOMBRE_APLICACION,
    RUTA_ICONO_APLICACION,
    RUTA_LOGO_EMPRESA,
    SALIDA_DIR,
    SUBTITULO_PRINCIPAL,
    TITULO_PRINCIPAL,
    VERSION_APLICACION,
)
from src.generador_excel import ErrorGeneracionExcel
from src.generador_mapa import (
    ErrorGeneracionMapa,
    generar_mapa_desde_informe,
)

from src.lector_excel import ErrorLecturaExcel
from src.procesador_informe import procesar_informe_ans
from src.transformador import ErrorTransformacion


logger = logging.getLogger(__name__)


class AplicacionANS:
    """
    Ventana principal del proyecto Informe ANS.

    La interfaz administra la interacción con el usuario.
    La lectura, transformación y generación de archivos se delegan
    a los módulos especializados del proyecto.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.proceso_activo = False
        self.logo_empresa: ImageTk.PhotoImage | None = None

        self.configurar_ventana()
        self.configurar_estilos()
        self.construir_interfaz()

        logger.info("Interfaz principal iniciada.")

    # ==========================================================
    # CONFIGURACIÓN GENERAL
    # ==========================================================

    def configurar_ventana(self) -> None:
        """
        Configura título, tamaño, posición y comportamiento.
        """

        self.root.title(
            f"{NOMBRE_APLICACION} - Versión {VERSION_APLICACION}"
        )

        try:
            self.root.iconbitmap(
                str(
                    RUTA_ICONO_APLICACION
                )
            )

        except tk.TclError:
            logger.warning(
                "No fue posible cargar el ícono de la aplicación."
            )

        self.root.configure(
            bg=COLOR_FONDO
        )

        pantalla_ancho = self.root.winfo_screenwidth()
        pantalla_alto = self.root.winfo_screenheight()

        posicion_x = (
            pantalla_ancho // 2
            - ANCHO_VENTANA // 2
        )

        posicion_y = (
            pantalla_alto // 2
            - ALTO_VENTANA // 2
        )

        self.root.geometry(
            f"{ANCHO_VENTANA}x{ALTO_VENTANA}"
            f"+{posicion_x}+{posicion_y}"
        )

        self.root.minsize(
            ANCHO_VENTANA,
            ALTO_VENTANA,
        )

        self.root.resizable(
            False,
            False,
        )

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.cerrar_aplicacion,
        )

    def configurar_estilos(self) -> None:
        """
        Configura los estilos visuales de componentes ttk.
        """

        estilo = ttk.Style()

        try:
            estilo.theme_use("clam")

        except tk.TclError:
            logger.warning(
                "No fue posible aplicar el tema ttk 'clam'."
            )

        estilo.configure(
            "ANS.Horizontal.TProgressbar",
            troughcolor="#D5D8DC",
            background=COLOR_VERDE_EMPRESA_CLARO,
            bordercolor=COLOR_BORDE,
            lightcolor=COLOR_VERDE_EMPRESA_CLARO,
            darkcolor=COLOR_VERDE_EMPRESA,
            thickness=18,
        )

        estilo.configure(
            "ANS.TSeparator",
            background=COLOR_BORDE,
        )

    # ==========================================================
    # CONSTRUCCIÓN DE INTERFAZ
    # ==========================================================

    def construir_interfaz(self) -> None:
        """
        Construye todos los componentes principales.
        """

        self.construir_barra_superior()
        self.construir_encabezado()
        self.construir_botones()
        self.construir_progreso()
        self.construir_area_log()
        self.construir_pie_pagina()

    def construir_barra_superior(self) -> None:
        """
        Construye la barra superior con reloj.
        """

        barra_superior = tk.Frame(
            self.root,
            bg=COLOR_VERDE_EMPRESA,
            height=26,
        )

        barra_superior.pack(
            fill="x"
        )

        barra_superior.pack_propagate(
            False
        )

        self.reloj = tk.Label(
            barra_superior,
            text="",
            bg=COLOR_VERDE_EMPRESA,
            fg=COLOR_BLANCO,
            font=("Segoe UI", 9, "bold"),
            anchor="e",
        )

        self.reloj.pack(
            side="right",
            padx=14,
        )

        self.actualizar_reloj()

    def construir_encabezado(self) -> None:
        """
        Construye el encabezado corporativo.
        """

        frame_banner = tk.Frame(
            self.root,
            bg=COLOR_FONDO_BANNER,
            height=170,
        )

        frame_banner.pack(
            fill="x"
        )

        frame_banner.pack_propagate(
            False
        )

        contenido_banner = tk.Frame(
            frame_banner,
            bg=COLOR_FONDO_BANNER,
        )

        contenido_banner.pack(
            expand=True
        )

        self.construir_logo(
            contenido_banner
        )

        titulo = tk.Label(
            contenido_banner,
            text=TITULO_PRINCIPAL,
            bg=COLOR_FONDO_BANNER,
            fg=COLOR_TEXTO,
            font=("Segoe UI", 18, "bold"),
        )

        titulo.pack(
            pady=(0, 2)
        )

        subtitulo = tk.Label(
            contenido_banner,
            text=SUBTITULO_PRINCIPAL,
            bg=COLOR_FONDO_BANNER,
            fg=COLOR_VERDE_EMPRESA,
            font=("Segoe UI", 9, "bold"),
        )

        subtitulo.pack(
            pady=(0, 2)
        )

        ttk.Separator(
            self.root,
            orient="horizontal",
            style="ANS.TSeparator",
        ).pack(
            fill="x",
            pady=(0, 10),
        )

    def construir_logo(
        self,
        contenedor: tk.Widget,
    ) -> None:
        """
        Carga y muestra el logo corporativo.
        """

        try:
            if not RUTA_LOGO_EMPRESA.exists():
                raise FileNotFoundError(
                    f"No se encontró el logo: {RUTA_LOGO_EMPRESA}"
                )

            imagen = Image.open(
                RUTA_LOGO_EMPRESA
            )

            imagen.thumbnail(
                (218, 100),
                Image.Resampling.LANCZOS,
            )

            self.logo_empresa = ImageTk.PhotoImage(
                imagen
            )

            etiqueta_logo = tk.Label(
                contenedor,
                image=self.logo_empresa,
                bg=COLOR_FONDO_BANNER,
                borderwidth=0,
                highlightthickness=0,
            )

            etiqueta_logo.pack(
                pady=(4, 2)
            )

            logger.info(
                "Logo corporativo cargado: %s",
                RUTA_LOGO_EMPRESA.name,
            )

        except Exception as error:
            logger.exception(
                "No fue posible cargar el logo corporativo."
            )

            etiqueta_logo = tk.Label(
                contenedor,
                text="ELITE Ingenieros",
                bg=COLOR_FONDO_BANNER,
                fg=COLOR_VERDE_EMPRESA,
                font=("Segoe UI", 22, "bold"),
            )

            etiqueta_logo.pack(
                pady=(6, 2)
            )

            logger.warning(
                "Se utilizó texto alternativo para el logo: %s",
                error,
            )

    def construir_botones(self) -> None:
        """
        Construye el menú principal por módulos.

        La ventana principal presenta únicamente los módulos
        ANS CONEXIONES y ANS REDES. Las acciones específicas
        se administran dentro de cada subpanel.
        """

        contenedor = tk.Frame(
            self.root,
            bg=COLOR_FONDO,
        )

        contenedor.pack(
            fill="x",
            padx=34,
            pady=(5, 10),
        )

        contenedor.columnconfigure(
            (0, 1),
            weight=1,
            uniform="modulos",
        )

        self.btn_modulo_conexiones = self.crear_boton(
            contenedor=contenedor,
            texto="ANS\nCONEXIONES",
            comando=self.abrir_panel_conexiones,
            color=COLOR_VERDE_EMPRESA,
            color_hover=COLOR_PRINCIPAL_HOVER,
            color_texto=COLOR_BLANCO,
        )

        self.btn_modulo_conexiones.grid(
            row=0,
            column=0,
            padx=(8, 6),
            pady=4,
            sticky="ew",
        )

        self.btn_modulo_redes = self.crear_boton(
            contenedor=contenedor,
            texto="ANS\nREDES",
            comando=self.abrir_panel_redes,
            color="#7F8C8D",
            color_hover="#626D6F",
            color_texto=COLOR_BLANCO,
        )

        self.btn_modulo_redes.grid(
            row=0,
            column=1,
            padx=(6, 8),
            pady=4,
            sticky="ew",
        )

        self.btn_salir = self.crear_boton(
            contenedor=contenedor,
            texto="SALIR DEL PANEL",
            comando=self.cerrar_aplicacion,
            color="#5D6D7E",
            color_hover="#34495E",
            color_texto=COLOR_BLANCO,
        )

        self.btn_salir.grid(
            row=1,
            column=0,
            columnspan=2,
            padx=180,
            pady=(6, 0),
            sticky="ew",
        )

    def abrir_panel_conexiones(self) -> None:
        """
        Abre el subpanel operativo de ANS Conexiones.

        La ventana presenta las acciones organizadas por función
        e incluye el estado y el resultado del proceso.
        """

        if self.proceso_activo:
            return

        ventana_existente = getattr(
            self,
            "ventana_conexiones",
            None,
        )

        if (
            ventana_existente is not None
            and ventana_existente.winfo_exists()
        ):
            ventana_existente.lift()
            ventana_existente.focus_force()
            return

        # ==========================================================
        # PALETA VISUAL DEL SUBPANEL
        # ==========================================================

        color_fondo_modal = "#F4F6F8"
        color_tarjeta = "#FFFFFF"
        color_texto_principal = "#1F2937"
        color_texto_secundario = "#64748B"
        color_borde_suave = "#D7DEE5"

        color_accion_principal = "#00875A"
        color_accion_principal_hover = "#006B47"

        color_proceso_secundario = "#0F766E"
        color_proceso_secundario_hover = "#0B5F59"

        color_consulta = "#475569"
        color_consulta_hover = "#334155"

        color_cerrar = "#E5E7EB"
        color_cerrar_hover = "#D1D5DB"
        color_texto_cerrar = "#334155"

        # ==========================================================
        # CREACIÓN Y CONFIGURACIÓN DE LA VENTANA
        # ==========================================================

        ventana = tk.Toplevel(
            self.root
        )

        self.ventana_conexiones = ventana

        ventana.title(
            "ANS Conexiones"
        )

        try:
            ventana.iconbitmap(
                str(
                    RUTA_ICONO_APLICACION
                )
            )

        except tk.TclError:
            logger.warning(
                "No fue posible cargar el ícono "
                "del subpanel Conexiones."
            )

        ventana.configure(
            bg=color_fondo_modal
        )

        # ==========================================================
        # MISMO TAMAÑO Y POSICIÓN DE LA VENTANA PRINCIPAL
        # ==========================================================

        ancho = ANCHO_VENTANA
        alto = ALTO_VENTANA

        self.root.update_idletasks()

        posicion_x = self.root.winfo_x()
        posicion_y = self.root.winfo_y()

        ventana.geometry(
            f"{ancho}x{alto}+{posicion_x}+{posicion_y}"
        )

        ventana.minsize(
            ancho,
            alto,
        )

        ventana.resizable(
            False,
            False,
        )

        ventana.transient(
            self.root
        )

        ventana.grab_set()

        # ==========================================================
        # ENCABEZADO
        # ==========================================================

        encabezado = tk.Frame(
            ventana,
            bg=COLOR_VERDE_EMPRESA,
            height=82,
        )

        encabezado.pack(
            fill="x"
        )

        encabezado.pack_propagate(
            False
        )

        contenido_encabezado = tk.Frame(
            encabezado,
            bg=COLOR_VERDE_EMPRESA,
        )

        contenido_encabezado.pack(
            fill="both",
            expand=True,
            padx=30,
        )

        tk.Label(
            contenido_encabezado,
            text="ANS CONEXIONES",
            bg=COLOR_VERDE_EMPRESA,
            fg=COLOR_BLANCO,
            font=("Segoe UI", 17, "bold"),
            anchor="w",
        ).pack(
            fill="x",
            pady=(16, 0),
        )

        tk.Label(
            contenido_encabezado,
            text=(
                "Generación, actualización y consulta "
                "del informe operativo."
            ),
            bg=COLOR_VERDE_EMPRESA,
            fg="#DFF3EA",
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(
            fill="x",
            pady=(2, 14),
        )

        # ==========================================================
        # CONTENIDO PRINCIPAL
        # ==========================================================

        contenido = tk.Frame(
            ventana,
            bg=color_fondo_modal,
        )

        contenido.pack(
            fill="both",
            expand=True,
            padx=28,
            pady=(14, 18),
        )

        contenido.columnconfigure(
            0,
            weight=1,
        )

        contenido.rowconfigure(
            3,
            weight=1,
        )

        # ==========================================================
        # ACCIÓN PRINCIPAL
        # ==========================================================

        tk.Label(
            contenido,
            text="Proceso principal",
            bg=color_fondo_modal,
            fg=color_texto_principal,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 7),
        )

        self.btn_informe = self.crear_boton(
            contenedor=contenido,
            texto="GENERAR INFORME ANS CONEXIONES",
            comando=self.iniciar_generacion_informe,
            color=color_accion_principal,
            color_hover=color_accion_principal_hover,
            color_texto=COLOR_BLANCO,
        )

        self.btn_informe.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 14),
        )

        # ==========================================================
        # ACCIONES SECUNDARIAS Y DE CONSULTA
        # ==========================================================

        tarjeta_acciones = tk.Frame(
            contenido,
            bg=color_tarjeta,
            highlightbackground=color_borde_suave,
            highlightthickness=1,
        )

        tarjeta_acciones.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 14),
        )

        tarjeta_acciones.columnconfigure(
            (0, 1),
            weight=1,
            uniform="acciones_conexiones",
        )

        tk.Label(
            tarjeta_acciones,
            text="Procesamiento complementario",
            bg=color_tarjeta,
            fg=color_texto_principal,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=16,
            pady=(13, 7),
        )

        self.btn_dashboard = self.crear_boton(
            contenedor=tarjeta_acciones,
            texto="ACTUALIZAR DASHBOARD",
            comando=self.iniciar_actualizacion_dashboard,
            color=color_proceso_secundario,
            color_hover=color_proceso_secundario_hover,
            color_texto=COLOR_BLANCO,
        )

        self.btn_dashboard.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=(16, 6),
            pady=(0, 12),
        )

        self.btn_mapa = self.crear_boton(
            contenedor=tarjeta_acciones,
            texto="GENERAR MAPA",
            comando=self.generar_mapa_ans,
            color=color_proceso_secundario,
            color_hover=color_proceso_secundario_hover,
            color_texto=COLOR_BLANCO,
        )

        self.btn_mapa.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(6, 16),
            pady=(0, 12),
        )

        ttk.Separator(
            tarjeta_acciones,
            orient="horizontal",
        ).grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=16,
            pady=(0, 10),
        )

        tk.Label(
            tarjeta_acciones,
            text="Abrir y consultar resultados",
            bg=color_tarjeta,
            fg=color_texto_principal,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).grid(
            row=3,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=16,
            pady=(0, 7),
        )

        self.btn_abrir_informe = self.crear_boton(
            contenedor=tarjeta_acciones,
            texto="ABRIR INFORME",
            comando=self.abrir_informe_conexiones,
            color=color_consulta,
            color_hover=color_consulta_hover,
            color_texto=COLOR_BLANCO,
        )

        self.btn_abrir_informe.grid(
            row=4,
            column=0,
            sticky="ew",
            padx=(16, 6),
            pady=(0, 14),
        )

        self.btn_abrir_dashboard = self.crear_boton(
            contenedor=tarjeta_acciones,
            texto="ABRIR DASHBOARD",
            comando=self.abrir_archivo_dashboard_conexiones,
            color=color_consulta,
            color_hover=color_consulta_hover,
            color_texto=COLOR_BLANCO,
        )

        self.btn_abrir_dashboard.grid(
            row=4,
            column=1,
            sticky="ew",
            padx=(6, 16),
            pady=(0, 14),
        )

        # ==========================================================
        # ESTADO Y RESULTADO DEL PROCESO
        # ==========================================================

        tarjeta_resultado = tk.Frame(
            contenido,
            bg=color_tarjeta,
            highlightbackground=color_borde_suave,
            highlightthickness=1,
        )

        tarjeta_resultado.grid(
            row=3,
            column=0,
            sticky="nsew",
            pady=(0, 14),
        )

        tarjeta_resultado.columnconfigure(
            0,
            weight=1,
        )

        tarjeta_resultado.rowconfigure(
            3,
            weight=1,
        )

        tk.Label(
            tarjeta_resultado,
            text="Estado y resultado del proceso",
            bg=color_tarjeta,
            fg=color_texto_principal,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=16,
            pady=(13, 4),
        )

        self.etiqueta_estado_conexiones = tk.Label(
            tarjeta_resultado,
            text="Esperando acción del usuario...",
            bg=color_tarjeta,
            fg=color_texto_secundario,
            font=("Segoe UI", 9),
            anchor="w",
        )

        self.etiqueta_estado_conexiones.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=16,
            pady=(0, 6),
        )

        self.barra_progreso_conexiones = ttk.Progressbar(
            tarjeta_resultado,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            style="ANS.Horizontal.TProgressbar",
        )

        self.barra_progreso_conexiones.grid(
            row=2,
            column=0,
            sticky="ew",
            padx=16,
            pady=(0, 10),
        )

        self.area_mensajes_conexiones = scrolledtext.ScrolledText(
            tarjeta_resultado,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#F8FAFC",
            fg=color_texto_principal,
            insertbackground=color_texto_principal,
            relief="solid",
            borderwidth=1,
            height=5,
            state="disabled",
        )

        self.area_mensajes_conexiones.grid(
            row=3,
            column=0,
            sticky="nsew",
            padx=16,
            pady=(0, 14),
        )

        self.area_mensajes_conexiones.tag_config(
            "info",
            foreground="#2563A6",
        )

        self.area_mensajes_conexiones.tag_config(
            "correcto",
            foreground=color_accion_principal,
        )

        self.area_mensajes_conexiones.tag_config(
            "advertencia",
            foreground="#B9770E",
        )

        self.area_mensajes_conexiones.tag_config(
            "error",
            foreground="#C0392B",
        )

        hora = datetime.now().strftime(
            "%H:%M:%S"
        )

        self.area_mensajes_conexiones.configure(
            state="normal"
        )

        self.area_mensajes_conexiones.insert(
            tk.END,
            f"[{hora}] Panel ANS Conexiones preparado.\n",
            "correcto",
        )

        self.area_mensajes_conexiones.configure(
            state="disabled"
        )

        # ==========================================================
        # BOTÓN CERRAR
        # ==========================================================

        frame_cierre = tk.Frame(
            contenido,
            bg=color_fondo_modal,
            height=48,
        )

        frame_cierre.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(0, 4),
        )

        frame_cierre.grid_propagate(
            False
        )

        frame_cierre.columnconfigure(
            0,
            weight=1,
        )

        self.btn_cerrar_conexiones = self.crear_boton(
            contenedor=frame_cierre,
            texto="CERRAR",
            comando=self.cerrar_panel_conexiones,
            color=color_cerrar,
            color_hover=color_cerrar_hover,
            color_texto=color_texto_cerrar,
        )

        self.btn_cerrar_conexiones.grid(
            row=0,
            column=0,
            sticky="e",
            padx=(0, 2),
            pady=2,
            ipadx=24,
        )
        

        ventana.protocol(
            "WM_DELETE_WINDOW",
            self.cerrar_panel_conexiones,
        )

    def cerrar_panel_conexiones(self) -> None:
        """
        Cierra correctamente el subpanel ANS Conexiones.
        """

        ventana = getattr(
            self,
            "ventana_conexiones",
            None,
        )

        if (
            ventana is not None
            and ventana.winfo_exists()
        ):
            try:
                ventana.grab_release()

            except tk.TclError:
                pass

            ventana.destroy()

        self.ventana_conexiones = None   

    def abrir_panel_redes(self) -> None:
        """
        Abre el subpanel reservado para el futuro módulo ANS Redes.
        """

        if self.proceso_activo:
            return

        ventana_existente = getattr(
            self,
            "ventana_redes",
            None,
        )

        if (
            ventana_existente is not None
            and ventana_existente.winfo_exists()
        ):
            ventana_existente.lift()
            ventana_existente.focus_force()
            return

        ventana = tk.Toplevel(
            self.root
        )

        self.ventana_redes = ventana

        ventana.title(
            "ANS Redes"
        )
        try:
            ventana.iconbitmap(
                str(
                    RUTA_ICONO_APLICACION
                )
            )

        except tk.TclError:
            logger.warning(
                "No fue posible cargar el ícono del subpanel Redes."
            )
            
        ventana.configure(
            bg=COLOR_FONDO
        )

        ancho = 500
        alto = 300

        self.root.update_idletasks()

        posicion_x = (
            self.root.winfo_x()
            + (self.root.winfo_width() - ancho) // 2
        )

        posicion_y = (
            self.root.winfo_y()
            + (self.root.winfo_height() - alto) // 2
        )

        ventana.geometry(
            f"{ancho}x{alto}+{posicion_x}+{posicion_y}"
        )

        ventana.resizable(
            False,
            False,
        )

        ventana.transient(
            self.root
        )

        ventana.grab_set()

        encabezado = tk.Frame(
            ventana,
            bg="#7F8C8D",
            height=58,
        )

        encabezado.pack(
            fill="x"
        )

        encabezado.pack_propagate(
            False
        )

        tk.Label(
            encabezado,
            text="ANS REDES",
            bg="#7F8C8D",
            fg=COLOR_BLANCO,
            font=("Segoe UI", 16, "bold"),
        ).pack(
            expand=True
        )

        tk.Label(
            ventana,
            text=(
                "Módulo reservado para el próximo desarrollo.\n\n"
                "Las reglas de negocio, el informe, el mapa y el "
                "dashboard se habilitarán cuando sean definidos."
            ),
            bg=COLOR_FONDO,
            fg=COLOR_TEXTO,
            font=("Segoe UI", 10),
            justify="center",
            wraplength=410,
        ).pack(
            expand=True,
            padx=30,
            pady=24,
        )

        boton_cerrar = tk.Button(
            ventana,
            text="CERRAR",
            command=ventana.destroy,
            bg="#5D6D7E",
            fg=COLOR_BLANCO,
            activebackground="#34495E",
            activeforeground=COLOR_BLANCO,
            font=("Segoe UI", 10, "bold"),
            relief="ridge",
            borderwidth=3,
            cursor="hand2",
            height=2,
        )

        boton_cerrar.pack(
            fill="x",
            padx=120,
            pady=(0, 24),
        )

        ventana.protocol(
            "WM_DELETE_WINDOW",
            ventana.destroy,
        )

    def crear_boton(
        self,
        contenedor: tk.Widget,
        texto: str,
        comando: Callable[[], None],
        color: str,
        color_hover: str,
        color_texto: str,
    ) -> tk.Button:
        """
        Crea un botón corporativo reutilizable.
        """

        boton = tk.Button(
            contenedor,
            text=texto,
            command=comando,
            bg=color,
            fg=color_texto,
            activebackground=color_hover,
            activeforeground=color_texto,
            disabledforeground="#D5D8DC",
            font=("Segoe UI", 10, "bold"),
            height=2,
            relief="ridge",
            borderwidth=3,
            cursor="hand2",
            highlightthickness=0,
        )

        boton.bind(
            "<Enter>",
            lambda evento: self.aplicar_hover(
                boton,
                color_hover,
            ),
        )

        boton.bind(
            "<Leave>",
            lambda evento: self.retirar_hover(
                boton,
                color,
            ),
        )

        return boton

    def aplicar_hover(
        self,
        boton: tk.Button,
        color_hover: str,
    ) -> None:
        """
        Aplica el color hover cuando el botón está habilitado.
        """

        if str(boton["state"]) != tk.DISABLED:
            boton.configure(
                bg=color_hover
            )

    def retirar_hover(
        self,
        boton: tk.Button,
        color_normal: str,
    ) -> None:
        """
        Restaura el color normal del botón.
        """

        if str(boton["state"]) != tk.DISABLED:
            boton.configure(
                bg=color_normal
            )

    def construir_progreso(self) -> None:
        """
        Construye la barra de progreso.
        """

        frame_progreso = tk.Frame(
            self.root,
            bg=COLOR_FONDO,
        )

        frame_progreso.pack(
            fill="x",
            padx=42,
            pady=(2, 10),
        )

        self.etiqueta_progreso = tk.Label(
            frame_progreso,
            text="Estado del proceso",
            bg=COLOR_FONDO,
            fg=COLOR_TEXTO_SECUNDARIO,
            font=("Segoe UI", 9, "bold"),
            anchor="w",
        )

        self.etiqueta_progreso.pack(
            fill="x",
            pady=(0, 4),
        )

        self.barra_progreso = ttk.Progressbar(
            frame_progreso,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            style="ANS.Horizontal.TProgressbar",
        )

        self.barra_progreso.pack(
            fill="x"
        )

    def construir_area_log(self) -> None:
        """
        Construye el área de resultados.
        """

        frame = tk.Frame(
            self.root,
            bg=COLOR_FONDO,
        )

        frame.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=(0, 8),
        )

        titulo = tk.Label(
            frame,
            text="Resultado del proceso",
            bg=COLOR_FONDO,
            fg=COLOR_TEXTO,
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        )

        titulo.pack(
            fill="x",
            pady=(0, 5),
        )

        self.area_mensajes = scrolledtext.ScrolledText(
            frame,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg=COLOR_BLANCO,
            fg=COLOR_TEXTO,
            insertbackground=COLOR_TEXTO,
            relief="solid",
            borderwidth=1,
            height=12,
            state="disabled",
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
            foreground=COLOR_VERDE_EMPRESA,
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
            "Directorio de entrada verificado.",
            "correcto",
        )

        logger.info(
            "Carpeta de entrada: %s",
            ENTRADA_DIR,
        )

    def construir_pie_pagina(self) -> None:
        """
        Construye el pie de página.
        """

        ttk.Separator(
            self.root,
            orient="horizontal",
            style="ANS.TSeparator",
        ).pack(
            fill="x",
            padx=25,
            pady=(4, 5),
        )

        frame_footer = tk.Frame(
            self.root,
            bg=COLOR_FONDO,
        )

        frame_footer.pack(
            fill="x",
            padx=30,
            pady=(0, 14),
        )

        self.estado = tk.Label(
            frame_footer,
            text="Esperando acción del usuario...",
            bg=COLOR_FONDO,
            fg=COLOR_TEXTO_SECUNDARIO,
            font=("Segoe UI", 9, "italic"),
            anchor="w",
        )

        self.estado.pack(
            side="left"
        )

        version = tk.Label(
            frame_footer,
            text=f"Versión {VERSION_APLICACION}",
            bg=COLOR_FONDO,
            fg=COLOR_TEXTO_SECUNDARIO,
            font=("Segoe UI", 8),
            anchor="e",
        )

        version.pack(
            side="right"
        )

    # ==========================================================
    # UTILIDADES
    # ==========================================================

    def actualizar_reloj(self) -> None:
        """
        Actualiza el reloj cada segundo.
        """

        hora_actual = datetime.now().strftime(
            "%I:%M:%S %p"
        )

        self.reloj.configure(
            text=hora_actual
        )

        self.root.after(
            1000,
            self.actualizar_reloj,
        )

    def agregar_mensaje(
        self,
        mensaje: str,
        etiqueta: str = "info",
    ) -> None:
        """
        Agrega un mensaje al área principal y, cuando está abierto,
        al panel visual de ANS Conexiones.
        """

        hora = datetime.now().strftime(
            "%H:%M:%S"
        )

        texto_mensaje = (
            f"[{hora}] {mensaje}\n"
        )

        # Área principal.
        self.area_mensajes.configure(
            state="normal"
        )

        self.area_mensajes.insert(
            tk.END,
            texto_mensaje,
            etiqueta,
        )

        self.area_mensajes.configure(
            state="disabled"
        )

        self.area_mensajes.see(
            tk.END
        )

        # Área del subpanel ANS Conexiones.
        area_conexiones = getattr(
            self,
            "area_mensajes_conexiones",
            None,
        )

        if (
            area_conexiones is not None
            and area_conexiones.winfo_exists()
        ):
            area_conexiones.configure(
                state="normal"
            )

            area_conexiones.insert(
                tk.END,
                texto_mensaje,
                etiqueta,
            )

            area_conexiones.configure(
                state="disabled"
            )

            area_conexiones.see(
                tk.END
            )

    def cambiar_estado(
        self,
        mensaje: str,
    ) -> None:
        """
        Actualiza el estado inferior y el estado del subpanel.
        """

        self.estado.configure(
            text=mensaje
        )

        estado_conexiones = getattr(
            self,
            "etiqueta_estado_conexiones",
            None,
        )

        if (
            estado_conexiones is not None
            and estado_conexiones.winfo_exists()
        ):
            estado_conexiones.configure(
                text=mensaje
            )

    def cambiar_etiqueta_progreso(
        self,
        mensaje: str,
    ) -> None:
        """
        Actualiza el texto de la barra de progreso.
        """

        self.etiqueta_progreso.configure(
            text=mensaje
        )

    def bloquear_botones(
        self,
        bloquear: bool,
    ) -> None:
        """
        Habilita o deshabilita los controles disponibles.
        """

        estado = (
            tk.DISABLED
            if bloquear
            else tk.NORMAL
        )

        self.btn_modulo_conexiones.configure(
            state=estado
        )

        self.btn_modulo_redes.configure(
            state=estado
        )

        controles_subpanel = (
            "btn_informe",
            "btn_dashboard",
            "btn_abrir_informe",
            "btn_abrir_dashboard",
            "btn_mapa",
            "btn_cerrar_conexiones",
        )

        for nombre_control in controles_subpanel:
            control = getattr(
                self,
                nombre_control,
                None,
            )

            if (
                control is not None
                and control.winfo_exists()
            ):
                control.configure(
                    state=estado
                )

        self.proceso_activo = bloquear

    def preparar_proceso(
        self,
        estado: str,
        etiqueta: str,
        mensaje_log: str,
    ) -> None:
        """
        Prepara visualmente un proceso en la ventana principal
        y en el panel ANS Conexiones.
        """

        self.barra_progreso.configure(
            mode="indeterminate",
            value=0,
        )

        self.barra_progreso.start(
            15
        )

        barra_conexiones = getattr(
            self,
            "barra_progreso_conexiones",
            None,
        )

        if (
            barra_conexiones is not None
            and barra_conexiones.winfo_exists()
        ):
            barra_conexiones.configure(
                mode="indeterminate",
                value=0,
            )

            barra_conexiones.start(
                15
            )

        self.cambiar_estado(
            estado
        )

        self.cambiar_etiqueta_progreso(
            etiqueta
        )

        self.agregar_mensaje(
            mensaje_log,
            "info",
        )

    def finalizar_proceso(self) -> None:
        """
        Restablece la interfaz al finalizar un proceso.
        """

        self.barra_progreso.stop()

        self.barra_progreso.configure(
            mode="determinate",
            value=0,
        )

        barra_conexiones = getattr(
            self,
            "barra_progreso_conexiones",
            None,
        )

        if (
            barra_conexiones is not None
            and barra_conexiones.winfo_exists()
        ):
            barra_conexiones.stop()

            barra_conexiones.configure(
                mode="determinate",
                value=0,
            )

        self.bloquear_botones(
            False
        )

        self.cambiar_estado(
            "Esperando acción del usuario..."
        )

        self.cambiar_etiqueta_progreso(
            "Estado del proceso"
        )

    def mostrar_error(
        self,
        mensaje: str,
    ) -> None:
        """
        Muestra un error controlado.
        """

        self.agregar_mensaje(
            mensaje,
            "error",
        )

        messagebox.showerror(
            NOMBRE_APLICACION,
            mensaje,
        )

    # ==========================================================
    # GENERACIÓN DEL INFORME ANS
    # ==========================================================

    def iniciar_generacion_informe(self) -> None:
        """
        Inicia el ETL en un hilo secundario.
        """

        if self.proceso_activo:
            return

        self.bloquear_botones(
            True
        )

        hilo = threading.Thread(
            target=self.generar_informe_ans,
            daemon=True,
        )

        hilo.start()

    def generar_informe_ans(self) -> None:
        """
        Unifica Región 1 y Región 2 y genera el informe.
        """

        try:
            self.root.after(
                0,
                lambda: self.preparar_proceso(
                    estado="Unificando Región 1 y Región 2...",
                    etiqueta="Procesando archivos regionales...",
                    mensaje_log=(
                        "Iniciando ETL de Región 1 y Región 2."
                    ),
                ),
            )

            ruta_informe, resumen = (
                procesar_informe_ans()
            )

            self.root.after(
                0,
                lambda: self.mostrar_informe_generado(
                    ruta_informe,
                    resumen,
                ),
            )

        except (
            FileNotFoundError,
            ErrorLecturaExcel,
            ErrorTransformacion,
            ErrorGeneracionExcel,
        ) as error:

            logger.warning(
                "Generación detenida: %s",
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
                "Error inesperado al generar el informe."
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

    def mostrar_informe_generado(
        self,
        ruta_informe,
        resumen: dict,
    ) -> None:
        """
        Presenta el resultado del ETL.
        """

        self.barra_progreso.stop()

        self.barra_progreso.configure(
            mode="determinate",
            value=100,
        )

        self.agregar_mensaje(
            f"Archivos procesados: "
            f"{resumen['ARCHIVOS_PROCESADOS']}",
            "info",
        )

        self.agregar_mensaje(
            f"Registros consolidados: "
            f"{resumen['REGISTROS_GENERADOS']:,}",
            "correcto",
        )

        self.agregar_mensaje(
            f"ID_ORDEN duplicados detectados: "
            f"{resumen['ID_ORDEN_DUPLICADOS']:,}",
            "advertencia",
        )

        self.agregar_mensaje(
            f"Fechas inválidas: "
            f"{resumen['FECHAS_INVALIDAS']:,}",
            "advertencia",
        )

        self.agregar_mensaje(
            f"Direcciones vacías: "
            f"{resumen['DIRECCIONES_VACIAS']:,}",
            "advertencia",
        )

        self.agregar_mensaje(
            f"Informe generado: {ruta_informe.name}",
            "correcto",
        )

        self.agregar_mensaje(
            f"Pedidos vencidos: "
            f"{resumen['PEDIDOS_VENCIDOS']:,}",
            "advertencia",
        )

        self.agregar_mensaje(
            f"Pedidos en alerta: "
            f"{resumen['PEDIDOS_ALERTA']:,}",
            "advertencia",
        )

        self.agregar_mensaje(
            f"Pedidos a tiempo: "
            f"{resumen['PEDIDOS_A_TIEMPO']:,}",
            "correcto",
        )

        self.agregar_mensaje(
            f"Fecha de corte ANS: "
            f"{resumen['FECHA_CORTE_ANS']}",
            "info",
        )

        messagebox.showinfo(
            "Informe ANS",
            "Informe unificado generado correctamente.\n\n"
            f"Registros: {resumen['REGISTROS_GENERADOS']:,}\n"
            f"Archivo: {ruta_informe.name}",
        )

        self.root.after(
            1500,
            lambda: self.barra_progreso.configure(
                value=0
            ),
        )

    # ==========================================================
    # GENERACIÓN DEL MAPA
    # ==========================================================

    def generar_mapa_ans(self) -> None:
        """
        Genera y abre el mapa ANS con pedidos reales.
        """

        if self.proceso_activo:
            return

        try:
            self.bloquear_botones(
                True
            )

            self.preparar_proceso(
                estado="Generando mapa con pedidos reales...",
                etiqueta="Construyendo mapa ANS...",
                mensaje_log=(
                    "Iniciando generación del mapa ANS con pedidos reales."
                ),
            )

            ruta_mapa, resumen_mapa = generar_mapa_desde_informe(
                abrir_navegador=True
            )

            self.agregar_mensaje(
                f"Mapa generado correctamente: {ruta_mapa.name}",
                "correcto",
            )

            self.agregar_mensaje(
                f"Direcciones consultadas: "
                f"{resumen_mapa['CONSULTAS_REALIZADAS']}",
                "info",
            )

            self.agregar_mensaje(
                f"Ubicaciones encontradas: "
                f"{resumen_mapa['COORDENADAS_ENCONTRADAS']}",
                "correcto",
            )

            self.agregar_mensaje(
                f"Coordenadas reutilizadas desde caché: "
                f"{resumen_mapa['COORDENADAS_REUTILIZADAS']}",
                "info",
            )

            self.agregar_mensaje(
                f"Pedidos visibles en el mapa: "
                f"{resumen_mapa['REGISTROS_EN_EL_MAPA']}",
                "correcto",
            )

            self.agregar_mensaje(
                f"Direcciones pendientes: "
                f"{resumen_mapa['PENDIENTES_POR_LIMITE']}",
                "advertencia",
            )

            self.agregar_mensaje(
                f"Direcciones no encontradas: "
                f"{resumen_mapa['DIRECCIONES_NO_ENCONTRADAS']}",
                "advertencia",
            )

            self.agregar_mensaje(
                f"Direcciones rechazadas por municipio: "
                f"{resumen_mapa['RECHAZADAS_POR_MUNICIPIO']}",
                "advertencia",
            )

            self.agregar_mensaje(
                f"Intentos de geocodificación realizados: "
                f"{resumen_mapa['INTENTOS_GEOCODIFICACION']}",
                "info",
            )

        except ErrorGeneracionMapa as error:
            logger.warning(
                "No fue posible generar el mapa: %s",
                error,
            )

            self.mostrar_error(
                str(error)
            )

        except Exception as error:
            logger.exception(
                "Error inesperado durante la generación del mapa."
            )

            self.mostrar_error(
                "Ocurrió un error inesperado al generar el mapa.\n\n"
                f"Detalle: {error}"
            )

        finally:
            self.finalizar_proceso()

    # ==========================================================
    # ACCIONES COMPLEMENTARIAS
    # ==========================================================

    def abrir_carpeta_salida(self) -> None:
        """
        Abre la carpeta de salida.
        """

        try:
            SALIDA_DIR.mkdir(
                parents=True,
                exist_ok=True,
            )

            if sys.platform.startswith("win"):
                os.startfile(
                    SALIDA_DIR
                )

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

    # ==========================================================
    # APERTURA DIRECTA DE ARCHIVOS DE CONEXIONES
    # ==========================================================

    def abrir_informe_conexiones(self) -> None:
        """
        Abre Informe_ANS_ELITE.xlsx y posiciona Excel
        directamente en la hoja DATOS_ANS.
        """

        ruta_informe = (
            SALIDA_DIR
            / "Informe_ANS_ELITE.xlsx"
        )

        self.abrir_archivo_excel_en_hoja(
            ruta_archivo=ruta_informe,
            nombre_hoja="DATOS_ANS",
            descripcion="Informe ANS Conexiones",
        )

    def abrir_archivo_dashboard_conexiones(self) -> None:
        """
        Abre INFORME_ANS.xlsb y posiciona Excel
        directamente en la hoja DATOS_ANS.
        """

        ruta_dashboard = (
            BASE_DIR
            / "dashboard"
            / "INFORME_ANS.xlsb"
        )

        self.abrir_archivo_excel_en_hoja(
            ruta_archivo=ruta_dashboard,
            nombre_hoja="DATOS_ANS",
            descripcion="Dashboard ANS Conexiones",
        )

    def abrir_archivo_excel_en_hoja(
        self,
        ruta_archivo,
        nombre_hoja: str,
        descripcion: str,
    ) -> None:
        """
        Abre un archivo de Excel y activa una hoja específica.
        """

        if not ruta_archivo.exists():
            self.mostrar_error(
                f"No se encontró el archivo: {descripcion}.\n\n"
                f"Ruta esperada:\n{ruta_archivo}"
            )
            return

        excel = None
        com_inicializado = False

        try:
            pythoncom.CoInitialize()
            com_inicializado = True

            excel = win32.DispatchEx(
                "Excel.Application"
            )

            excel.Visible = True
            excel.DisplayAlerts = True

            libro = excel.Workbooks.Open(
                str(
                    ruta_archivo.resolve()
                ),
                UpdateLinks=0,
                ReadOnly=False,
            )

            libro.Activate()

            try:
                hoja = libro.Worksheets(
                    nombre_hoja
                )

            except Exception as error:
                libro.Close(
                    SaveChanges=False
                )

                raise RuntimeError(
                    f"No se encontró la hoja {nombre_hoja} "
                    f"en {ruta_archivo.name}."
                ) from error

            hoja.Activate()

            # Posiciona la vista en la primera celda.
            hoja.Range("A1").Select()

            # Lleva la hoja al inicio visible.
            excel.ActiveWindow.ScrollRow = 1
            excel.ActiveWindow.ScrollColumn = 1

            # Maximiza Excel.
            excel.WindowState = -4137

            # Excel queda visible y bajo control del usuario.
            excel.Visible = True
            excel.UserControl = True

            # ======================================================
            # CERRAR EL SUBPANEL ANTES DE MINIMIZAR
            # ======================================================

            self.cerrar_panel_conexiones() 

            # Minimiza solamente la ventana principal.
            self.root.iconify()

            # Lleva Excel al frente.
            libro.Activate()
            hoja.Activate()
            excel.ActiveWindow.Activate()

            self.agregar_mensaje(
                f"{descripcion} abierto en la hoja "
                f"{nombre_hoja}.",
                "correcto",
            )

        except Exception as error:
            logger.exception(
                "No fue posible abrir %s.",
                descripcion,
            )

            if excel is not None:
                try:
                    excel.Quit()

                except Exception:
                    pass

            self.mostrar_error(
                f"No fue posible abrir {descripcion}.\n\n"
                f"Detalle: {error}"
            )

        finally:
            if com_inicializado:
                pythoncom.CoUninitialize()

    # ==========================================================
    # ACTUALIZACIÓN DEL DASHBOARD ANS
    # ==========================================================

    def iniciar_actualizacion_dashboard(self) -> None:
        """
        Inicia la actualización del dashboard en un hilo secundario.
        """

        if self.proceso_activo:
            return

        self.bloquear_botones(
            True
        )

        hilo = threading.Thread(
            target=self.actualizar_dashboard_ans,
            daemon=True,
        )

        hilo.start()

    def actualizar_dashboard_ans(self) -> None:
        """
        Transfiere los datos del informe al archivo INFORME_ANS.xlsb
        y actualiza sus tablas dinámicas, gráficos y segmentadores.
        """

        try:
            self.root.after(
                0,
                lambda: self.preparar_proceso(
                    estado="Actualizando dashboard ANS...",
                    etiqueta="Transfiriendo datos al dashboard...",
                    mensaje_log=(
                        "Iniciando actualización del dashboard ANS."
                    ),
                ),
            )

            resumen = actualizar_dashboard(
                refrescar_dashboard=True
            )

            self.root.after(
                0,
                lambda: self.mostrar_dashboard_actualizado(
                    resumen
                ),
            )

        except ErrorActualizacionDashboard as error:
            logger.warning(
                "Actualización del dashboard detenida: %s",
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
                "Error inesperado al actualizar el dashboard."
            )

            self.root.after(
                0,
                lambda mensaje=str(error): (
                    self.mostrar_error(
                        "Ocurrió un error inesperado al actualizar "
                        "el dashboard.\n\n"
                        f"Detalle: {mensaje}"
                    )
                ),
            )

        finally:
            self.root.after(
                0,
                self.finalizar_proceso,
            )

    def mostrar_dashboard_actualizado(
        self,
        resumen: dict,
    ) -> None:
        """
        Presenta el resultado de la actualización del dashboard.
        """

        self.barra_progreso.stop()

        self.barra_progreso.configure(
            mode="determinate",
            value=100,
        )

        self.agregar_mensaje(
            "Dashboard ANS actualizado correctamente.",
            "correcto",
        )

        self.agregar_mensaje(
            f"Registros transferidos: "
            f"{resumen['REGISTROS_TRANSFERIDOS']:,}",
            "correcto",
        )

        self.agregar_mensaje(
            f"Columnas transferidas: "
            f"{resumen['COLUMNAS_TRANSFERIDAS']}",
            "info",
        )

        self.agregar_mensaje(
            f"Archivo actualizado: "
            f"{os.path.basename(resumen['ARCHIVO_DESTINO'])}",
            "correcto",
        )

        messagebox.showinfo(
            "Dashboard ANS",
            "Dashboard actualizado correctamente.\n\n"
            f"Registros transferidos: "
            f"{resumen['REGISTROS_TRANSFERIDOS']:,}\n"
            f"Archivo: "
            f"{os.path.basename(resumen['ARCHIVO_DESTINO'])}",
        )

        self.root.after(
            1500,
            lambda: self.barra_progreso.configure(
                value=0
            ),
        )

    def cerrar_aplicacion(self) -> None:
        """
        Cierra la aplicación de forma controlada.
        """

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
    """
    Inicia la ventana principal.
    """

    root = tk.Tk()
    AplicacionANS(root)
    root.mainloop()