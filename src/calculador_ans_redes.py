from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, time, timedelta
from pathlib import Path

import holidays
import pandas as pd

from src.config import BASE_DIR


# ============================================================
# RUTAS Y HOJAS DE CONFIGURACIÓN
# ============================================================

RUTA_CONFIG_REDES = (
    BASE_DIR
    / "config"
    / "FILTROS_ANS_REDES.xlsx"
)

HOJA_DIAS_CONTRACTUALES = "DIAS_CONTRACTUALES_REDES"
HOJA_CONFIGURACION_REDES = "CONFIGURACION_REDES"


# ============================================================
# PARÁMETROS OBLIGATORIOS
# ============================================================

PARAMETROS_CONFIGURACION_OBLIGATORIOS = {
    "EXCLUIR_SABADOS",
    "EXCLUIR_DOMINGOS",
    "EXCLUIR_FESTIVOS_COLOMBIA",
    "CONTEO_DESDE_MISMO_DIA",
    "CONSERVAR_HORA_INICIO",
}


class ErrorCalculoANSRedes(Exception):
    """
    Error controlado durante el cálculo ANS Redes.
    """


# ============================================================
# NORMALIZACIÓN
# ============================================================

def limpiar_texto(
    valor: object,
) -> str:
    """
    Convierte un valor en texto limpio.
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""

    except (TypeError, ValueError):
        pass

    return re.sub(
        r"\s+",
        " ",
        str(valor).strip(),
    )


def normalizar_clave(
    valor: object,
) -> str:
    """
    Normaliza textos para comparaciones robustas.

    Elimina:
    - tildes;
    - espacios;
    - signos;
    - diferencias entre mayúsculas/minúsculas.
    """

    texto = limpiar_texto(
        valor
    ).upper()

    texto = unicodedata.normalize(
        "NFKD",
        texto,
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    return re.sub(
        r"[^A-Z0-9]+",
        "",
        texto,
    )


def normalizar_proceso(
    valor: object,
) -> str:
    """
    Convierte 4109, 4109.0 y '4109' a '4109'.
    """

    return re.sub(
        r"\.0$",
        "",
        limpiar_texto(valor),
    )


def convertir_si_no(
    valor: object,
    nombre_parametro: str,
) -> bool:
    """
    Convierte SI/NO del Excel a booleano.
    """

    texto = normalizar_clave(
        valor
    )

    if texto in {
        "SI",
        "S",
        "TRUE",
        "VERDADERO",
        "1",
    }:
        return True

    if texto in {
        "NO",
        "N",
        "FALSE",
        "FALSO",
        "0",
    }:
        return False

    raise ErrorCalculoANSRedes(
        f"El parámetro {nombre_parametro} debe contener SI o NO."
    )


def convertir_porcentaje(
    valor: object,
) -> float:
    """
    Convierte PORCENTAJE_ALERTA a decimal.

    Acepta:
    - 70% en Excel -> 0.70
    - 0.70
    - 70
    - texto '70%'
    """

    if valor is None:
        raise ErrorCalculoANSRedes(
            "PORCENTAJE_ALERTA no puede estar vacío."
        )

    try:
        if pd.isna(valor):
            raise ErrorCalculoANSRedes(
                "PORCENTAJE_ALERTA no puede estar vacío."
            )

    except (TypeError, ValueError):
        pass

    texto = str(
        valor
    ).strip().replace(
        ",",
        ".",
    )

    tiene_signo_porcentaje = (
        texto.endswith("%")
    )

    if tiene_signo_porcentaje:
        texto = texto[:-1].strip()

    try:
        numero = float(
            texto
        )

    except (TypeError, ValueError) as error:
        raise ErrorCalculoANSRedes(
            f"PORCENTAJE_ALERTA inválido: {valor}"
        ) from error

    if tiene_signo_porcentaje:
        numero /= 100

    elif numero > 1:
        numero /= 100

    if (
        numero <= 0
        or numero > 1
    ):
        raise ErrorCalculoANSRedes(
            "PORCENTAJE_ALERTA debe estar entre 0% y 100%."
        )

    return numero


# ============================================================
# LECTURA DEL ARCHIVO DE CONFIGURACIÓN
# ============================================================

def validar_archivo_configuracion() -> None:
    """
    Valida que exista FILTROS_ANS_REDES.xlsx.
    """

    if not RUTA_CONFIG_REDES.exists():
        raise ErrorCalculoANSRedes(
            "No se encontró el archivo de configuración ANS Redes:\n\n"
            f"{RUTA_CONFIG_REDES}\n\n"
            "Ubique FILTROS_ANS_REDES.xlsx dentro de config."
        )


def leer_hoja_configuracion(
    nombre_hoja: str,
) -> pd.DataFrame:
    """
    Lee una hoja del archivo FILTROS_ANS_REDES.xlsx.
    """

    validar_archivo_configuracion()

    try:
        dataframe = pd.read_excel(
            RUTA_CONFIG_REDES,
            sheet_name=nombre_hoja,
            dtype=object,
            engine="openpyxl",
        )

    except PermissionError as error:
        raise ErrorCalculoANSRedes(
            "No fue posible leer FILTROS_ANS_REDES.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a ejecutar."
        ) from error

    except ValueError as error:
        raise ErrorCalculoANSRedes(
            f"No se encontró la hoja {nombre_hoja} dentro de "
            "FILTROS_ANS_REDES.xlsx."
        ) from error

    except Exception as error:
        raise ErrorCalculoANSRedes(
            f"No fue posible leer la hoja {nombre_hoja}."
        ) from error

    dataframe.columns = [
        limpiar_texto(
            columna
        ).upper()
        for columna in dataframe.columns
    ]

    return dataframe


# ============================================================
# CONFIGURACIÓN GENERAL DESDE EXCEL
# ============================================================

def cargar_configuracion_redes() -> dict[str, bool]:
    """
    Lee CONFIGURACION_REDES.

    Parámetros:
    - EXCLUIR_SABADOS
    - EXCLUIR_DOMINGOS
    - EXCLUIR_FESTIVOS_COLOMBIA
    - CONTEO_DESDE_MISMO_DIA
    - CONSERVAR_HORA_INICIO
    """

    dataframe = leer_hoja_configuracion(
        HOJA_CONFIGURACION_REDES
    )

    columnas_requeridas = {
        "PARAMETRO",
        "VALOR",
    }

    faltantes = columnas_requeridas.difference(
        dataframe.columns
    )

    if faltantes:
        raise ErrorCalculoANSRedes(
            "Faltan columnas en CONFIGURACION_REDES:\n\n"
            + "\n".join(
                f"- {columna}"
                for columna in sorted(faltantes)
            )
        )

    parametros: dict[str, object] = {}

    for _, fila in dataframe.iterrows():

        parametro = limpiar_texto(
            fila.get(
                "PARAMETRO"
            )
        ).upper()

        if not parametro:
            continue

        if parametro in parametros:
            raise ErrorCalculoANSRedes(
                f"El parámetro {parametro} está duplicado "
                "en CONFIGURACION_REDES."
            )

        parametros[
            parametro
        ] = fila.get(
            "VALOR"
        )

    faltantes_parametros = (
        PARAMETROS_CONFIGURACION_OBLIGATORIOS
        .difference(
            parametros
        )
    )

    if faltantes_parametros:
        raise ErrorCalculoANSRedes(
            "Faltan parámetros obligatorios en "
            "CONFIGURACION_REDES:\n\n"
            + "\n".join(
                f"- {parametro}"
                for parametro
                in sorted(faltantes_parametros)
            )
        )

    return {
        parametro: convertir_si_no(
            parametros[parametro],
            parametro,
        )
        for parametro
        in PARAMETROS_CONFIGURACION_OBLIGATORIOS
    }


# ============================================================
# REGLAS CONTRACTUALES DESDE EXCEL
# ============================================================

def cargar_reglas_contractuales_redes() -> list[dict]:
    """
    Lee DIAS_CONTRACTUALES_REDES.

    Columnas requeridas:
    - PROCESO
    - TIPO_ZONA
    - MUNICIPIO
    - DIAS_CONTRACTUALES
    - PORCENTAJE_ALERTA
    - ACTIVO
    """

    dataframe = leer_hoja_configuracion(
        HOJA_DIAS_CONTRACTUALES
    )

    columnas_requeridas = {
        "PROCESO",
        "TIPO_ZONA",
        "MUNICIPIO",
        "DIAS_CONTRACTUALES",
        "PORCENTAJE_ALERTA",
        "ACTIVO",
    }

    faltantes = columnas_requeridas.difference(
        dataframe.columns
    )

    if faltantes:
        raise ErrorCalculoANSRedes(
            "Faltan columnas en DIAS_CONTRACTUALES_REDES:\n\n"
            + "\n".join(
                f"- {columna}"
                for columna in sorted(faltantes)
            )
        )

    reglas: list[dict] = []

    for numero_fila, fila in dataframe.iterrows():

        activo = fila.get(
            "ACTIVO"
        )

        if pd.isna(
            activo
        ):
            continue

        if not convertir_si_no(
            activo,
            f"ACTIVO fila {numero_fila + 2}",
        ):
            continue

        proceso = normalizar_proceso(
            fila.get(
                "PROCESO"
            )
        )

        tipo_zona = normalizar_clave(
            fila.get(
                "TIPO_ZONA"
            )
        )

        municipio = normalizar_clave(
            fila.get(
                "MUNICIPIO"
            )
        )

        try:
            dias_contractuales = int(
                float(
                    fila.get(
                        "DIAS_CONTRACTUALES"
                    )
                )
            )

        except (TypeError, ValueError) as error:
            raise ErrorCalculoANSRedes(
                "DIAS_CONTRACTUALES inválido en la fila "
                f"{numero_fila + 2}."
            ) from error

        if dias_contractuales <= 0:
            raise ErrorCalculoANSRedes(
                "DIAS_CONTRACTUALES debe ser mayor que cero "
                f"en la fila {numero_fila + 2}."
            )

        porcentaje_alerta = convertir_porcentaje(
            fila.get(
                "PORCENTAJE_ALERTA"
            )
        )

        reglas.append(
            {
                "PROCESO": proceso,
                "TIPO_ZONA": tipo_zona,
                "MUNICIPIO": municipio,
                "DIAS_CONTRACTUALES": dias_contractuales,
                "PORCENTAJE_ALERTA": porcentaje_alerta,
            }
        )

    if not reglas:
        raise ErrorCalculoANSRedes(
            "No existen reglas contractuales activas "
            "en DIAS_CONTRACTUALES_REDES."
        )

    return reglas


def buscar_regla_contractual(
    proceso: object,
    tipo_zona: object,
    municipio: object,
    reglas: list[dict],
) -> dict:
    """
    Orden de prioridad:

    1. PROCESO + TIPO_ZONA + MUNICIPIO exactos.
    2. PROCESO + TODOS + TODOS.
    3. Para PROCESO 4109:
       RURAL + TODOS funciona como regla general de 23 días
       cuando no aplica uno de los municipios urbanos de 15 días.
    """

    proceso_normalizado = normalizar_proceso(
        proceso
    )

    zona_normalizada = normalizar_clave(
        tipo_zona
    )

    municipio_normalizado = normalizar_clave(
        municipio
    )

    reglas_proceso = [
        regla
        for regla in reglas
        if regla["PROCESO"] == proceso_normalizado
    ]

    # Coincidencia exacta.
    for regla in reglas_proceso:

        if (
            regla["TIPO_ZONA"] == zona_normalizada
            and regla["MUNICIPIO"] == municipio_normalizado
        ):
            return regla

    # Regla general del proceso.
    for regla in reglas_proceso:

        if (
            regla["TIPO_ZONA"] == "TODOS"
            and regla["MUNICIPIO"] == "TODOS"
        ):
            return regla

    # Regla general 4109: 23 días.
    if proceso_normalizado == "4109":

        for regla in reglas_proceso:

            if (
                regla["TIPO_ZONA"] == "RURAL"
                and regla["MUNICIPIO"] == "TODOS"
            ):
                return regla

    raise ErrorCalculoANSRedes(
        "No se encontró regla contractual para:\n\n"
        f"PROCESO: {proceso_normalizado}\n"
        f"TIPO_ZONA: {limpiar_texto(tipo_zona)}\n"
        f"MUNICIPIO: {limpiar_texto(municipio)}"
    )


# ============================================================
# FESTIVOS Y DÍAS HÁBILES
# ============================================================

def construir_festivos_colombia(
    fecha_minima: date,
    fecha_maxima: date,
) -> set[date]:
    """
    Construye los festivos oficiales de Colombia.
    """

    calendario = holidays.country_holidays(
        "CO",
        years=range(
            fecha_minima.year,
            fecha_maxima.year + 1,
        ),
    )

    return set(
        calendario.keys()
    )


def es_dia_habil(
    fecha: date,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
) -> bool:
    """
    Determina si la fecha cuenta para el ANS según Excel.
    """

    if (
        excluir_sabados
        and fecha.weekday() == 5
    ):
        return False

    if (
        excluir_domingos
        and fecha.weekday() == 6
    ):
        return False

    if fecha in festivos:
        return False

    return True


# ============================================================
# CÁLCULO DE FECHA LÍMITE
# ============================================================

def construir_datetime_limite(
    fecha: date,
    hora_inicio: time,
    conservar_hora_inicio: bool,
) -> datetime:
    """
    Si CONSERVAR_HORA_INICIO = SI:
        conserva HH:MM:SS.

    Si CONSERVAR_HORA_INICIO = NO:
        el vencimiento se controla únicamente por fecha;
        se utiliza el final del día para no vencer al inicio del día.
    """

    if conservar_hora_inicio:

        hora_limite = hora_inicio.replace(
            tzinfo=None
        )

    else:

        hora_limite = time(
            23,
            59,
            59,
        )

    return datetime.combine(
        fecha,
        hora_limite,
    )


def sumar_dias_habiles(
    fecha_inicio: datetime,
    dias_contractuales: int,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
    conteo_desde_mismo_dia: bool,
    conservar_hora_inicio: bool,
) -> datetime:
    """
    Calcula FECHA_LIMITE_ANS.

    CONTEO_DESDE_MISMO_DIA = SI:
        la fecha de creación es candidata a ser día 1.

    CONTEO_DESDE_MISMO_DIA = NO:
        el conteo comienza desde el día calendario siguiente.

    Sábados, domingos y festivos se excluyen según configuración.
    """

    fecha_actual = fecha_inicio.date()

    if not conteo_desde_mismo_dia:
        fecha_actual += timedelta(
            days=1
        )

    dias_contados = 0

    while dias_contados < dias_contractuales:

        if es_dia_habil(
            fecha=fecha_actual,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        ):
            dias_contados += 1

        if dias_contados < dias_contractuales:

            fecha_actual += timedelta(
                days=1
            )

    return construir_datetime_limite(
        fecha=fecha_actual,
        hora_inicio=fecha_inicio.time(),
        conservar_hora_inicio=conservar_hora_inicio,
    )


# ============================================================
# DÍAS TRANSCURRIDOS
# ============================================================

def contar_dias_transcurridos(
    fecha_inicio: datetime,
    fecha_actual: datetime,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
    conteo_desde_mismo_dia: bool,
    conservar_hora_inicio: bool,
) -> int:
    """
    Cuenta días contractuales alcanzados.

    Si se conserva la hora:
        cada nuevo día contractual se considera alcanzado
        a la misma HH:MM:SS de PRO_FECHA_SISTEMA_CREACION.
    """

    if fecha_actual < fecha_inicio:
        return 0

    fecha_cursor = fecha_inicio.date()

    if not conteo_desde_mismo_dia:
        fecha_cursor += timedelta(
            days=1
        )

    contador = 0

    while fecha_cursor <= fecha_actual.date():

        if es_dia_habil(
            fecha=fecha_cursor,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        ):

            if conservar_hora_inicio:

                hito = datetime.combine(
                    fecha_cursor,
                    fecha_inicio.time().replace(
                        tzinfo=None
                    ),
                )

                if hito <= fecha_actual:
                    contador += 1

            else:

                contador += 1

        fecha_cursor += timedelta(
            days=1
        )

    return contador


# ============================================================
# DÍAS RESTANTES
# ============================================================

def contar_dias_habiles_entre_fechas(
    fecha_inicial: date,
    fecha_final: date,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
) -> int:
    """
    Cuenta días hábiles inclusivos entre dos fechas.
    """

    if fecha_final < fecha_inicial:
        return 0

    contador = 0
    fecha_cursor = fecha_inicial

    while fecha_cursor <= fecha_final:

        if es_dia_habil(
            fecha=fecha_cursor,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        ):
            contador += 1

        fecha_cursor += timedelta(
            days=1
        )

    return contador


def calcular_dias_restantes(
    fecha_actual: datetime,
    fecha_limite: datetime,
    excluir_sabados: bool,
    excluir_domingos: bool,
    festivos: set[date],
) -> int:
    """
    - El día de la fecha límite devuelve 0.
    - Antes del límite devuelve positivos.
    - Después devuelve negativos desde el siguiente día hábil.
    """

    if (
        fecha_actual.date()
        == fecha_limite.date()
    ):
        return 0

    if fecha_actual < fecha_limite:

        return contar_dias_habiles_entre_fechas(
            fecha_inicial=(
                fecha_actual.date()
                + timedelta(days=1)
            ),
            fecha_final=fecha_limite.date(),
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
        )

    dias_vencidos = contar_dias_habiles_entre_fechas(
        fecha_inicial=(
            fecha_limite.date()
            + timedelta(days=1)
        ),
        fecha_final=fecha_actual.date(),
        excluir_sabados=excluir_sabados,
        excluir_domingos=excluir_domingos,
        festivos=festivos,
    )

    return -dias_vencidos


# ============================================================
# ALERTA Y ESTADO
# ============================================================

def calcular_dias_para_iniciar_alerta(
    dias_contractuales: int,
    porcentaje_alerta: float,
) -> int:
    """
    Respeta exactamente la regla entregada:

    ALERTA = 70 % DE LOS DÍAS

    15 días -> 10
    23 días -> 16
    45 días -> 31

    El resultado se trunca hacia abajo.
    """

    return int(
        dias_contractuales
        * porcentaje_alerta
    )


def determinar_estado(
    fecha_actual: datetime,
    fecha_limite: datetime,
    dias_transcurridos: int,
    dias_para_iniciar_alerta: int,
    dias_restantes: int,
) -> str:
    """
    Estados ANS Redes:

    fecha_actual > FECHA_LIMITE_ANS
        -> VENCIDO

    DIAS_RESTANTES = 0, sin superar todavía FECHA_LIMITE_ANS
        -> ALERTA 0 DIAS

    DIAS_TRANSCURRIDOS >= DIAS_PARA_INICIAR_ALERTA
        -> ALERTA

    resto
        -> A TIEMPO
    """

    if fecha_actual > fecha_limite:
        return "VENCIDO"

    if dias_restantes == 0:
        return "ALERTA 0 DIAS"

    if (
        dias_transcurridos
        >= dias_para_iniciar_alerta
    ):
        return "ALERTA"

    return "A TIEMPO"


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def aplicar_calculos_ans_redes(
    dataframe: pd.DataFrame,
    fecha_corte: datetime | pd.Timestamp | None = None,
) -> pd.DataFrame:
    """
    Aplica las reglas ANS Redes.

    Inicio:
        PRO_FECHA_SISTEMA_CREACION

    Configuración:
        CONFIGURACION_REDES

    Reglas contractuales:
        DIAS_CONTRACTUALES_REDES

    Salida:
        DIAS_CONTRACTUALES
        FECHA_LIMITE_ANS
        DIAS_TRANSCURRIDOS
        DIAS_PARA_INICIAR_ALERTA
        DIAS_RESTANTES
        ESTADO
    """

    columnas_requeridas = {
        "PROCESO",
        "PRO_D_CLASIFICACION",
        "CLI_D_MUNICIPIO",
        "PRO_FECHA_SISTEMA_CREACION",
    }

    faltantes = columnas_requeridas.difference(
        dataframe.columns
    )

    if faltantes:
        raise ErrorCalculoANSRedes(
            "Faltan columnas para calcular ANS Redes:\n\n"
            + "\n".join(
                f"- {columna}"
                for columna in sorted(faltantes)
            )
        )

    # --------------------------------------------------------
    # CARGAR CONFIGURACIÓN Y REGLAS
    # --------------------------------------------------------

    configuracion = (
        cargar_configuracion_redes()
    )

    reglas = (
        cargar_reglas_contractuales_redes()
    )

    excluir_sabados = configuracion[
        "EXCLUIR_SABADOS"
    ]

    excluir_domingos = configuracion[
        "EXCLUIR_DOMINGOS"
    ]

    excluir_festivos = configuracion[
        "EXCLUIR_FESTIVOS_COLOMBIA"
    ]

    conteo_desde_mismo_dia = configuracion[
        "CONTEO_DESDE_MISMO_DIA"
    ]

    conservar_hora_inicio = configuracion[
        "CONSERVAR_HORA_INICIO"
    ]

    resultado = dataframe.copy()

    resultado[
        "PRO_FECHA_SISTEMA_CREACION"
    ] = pd.to_datetime(
        resultado[
            "PRO_FECHA_SISTEMA_CREACION"
        ],
        errors="coerce",
        dayfirst=True,
    )

    # --------------------------------------------------------
    # FECHA/HORA DE CORTE
    # --------------------------------------------------------

    if fecha_corte is None:

        fecha_actual = datetime.now()

    else:

        fecha_convertida = pd.to_datetime(
            fecha_corte,
            errors="coerce",
        )

        if pd.isna(
            fecha_convertida
        ):
            raise ErrorCalculoANSRedes(
                "La fecha/hora de corte no es válida."
            )

        fecha_actual = pd.Timestamp(
            fecha_convertida
        ).to_pydatetime()

    # --------------------------------------------------------
    # RANGO PARA FESTIVOS
    # --------------------------------------------------------

    fechas_validas = resultado[
        "PRO_FECHA_SISTEMA_CREACION"
    ].dropna()

    if fechas_validas.empty:

        fecha_minima = (
            fecha_actual.date()
        )

        fecha_maxima_base = (
            fecha_actual.date()
        )

    else:

        fecha_minima = min(
            fechas_validas.min().date(),
            fecha_actual.date(),
        )

        fecha_maxima_base = max(
            fechas_validas.max().date(),
            fecha_actual.date(),
        )

    dias_maximos = max(
        regla[
            "DIAS_CONTRACTUALES"
        ]
        for regla in reglas
    )

    fecha_maxima = (
        fecha_maxima_base
        + timedelta(
            days=(dias_maximos * 4) + 370
        )
    )

    # --------------------------------------------------------
    # FESTIVOS
    # --------------------------------------------------------

    if excluir_festivos:

        festivos = construir_festivos_colombia(
            fecha_minima=fecha_minima,
            fecha_maxima=fecha_maxima,
        )

    else:

        festivos = set()

    # --------------------------------------------------------
    # ACUMULADORES
    # --------------------------------------------------------

    dias_contractuales_resultado = []
    fechas_limite_resultado = []
    dias_transcurridos_resultado = []
    dias_alerta_resultado = []
    dias_restantes_resultado = []
    estados_resultado = []

    # --------------------------------------------------------
    # PROCESAR REGISTROS
    # --------------------------------------------------------

    for _, fila in resultado.iterrows():

        regla = buscar_regla_contractual(
            proceso=fila.get(
                "PROCESO"
            ),
            tipo_zona=fila.get(
                "PRO_D_CLASIFICACION"
            ),
            municipio=fila.get(
                "CLI_D_MUNICIPIO"
            ),
            reglas=reglas,
        )

        dias_contractuales = regla[
            "DIAS_CONTRACTUALES"
        ]

        porcentaje_alerta = regla[
            "PORCENTAJE_ALERTA"
        ]

        dias_para_alerta = (
            calcular_dias_para_iniciar_alerta(
                dias_contractuales=(
                    dias_contractuales
                ),
                porcentaje_alerta=(
                    porcentaje_alerta
                ),
            )
        )

        fecha_inicio = fila.get(
            "PRO_FECHA_SISTEMA_CREACION"
        )

        dias_contractuales_resultado.append(
            dias_contractuales
        )

        dias_alerta_resultado.append(
            dias_para_alerta
        )

        # ----------------------------------------------------
        # FECHA INVÁLIDA
        # ----------------------------------------------------

        if pd.isna(
            fecha_inicio
        ):

            fechas_limite_resultado.append(
                pd.NaT
            )

            dias_transcurridos_resultado.append(
                pd.NA
            )

            dias_restantes_resultado.append(
                pd.NA
            )

            estados_resultado.append(
                "SIN FECHA"
            )

            continue

        fecha_inicio_datetime = pd.Timestamp(
            fecha_inicio
        ).to_pydatetime()

        # ----------------------------------------------------
        # FECHA LÍMITE
        # ----------------------------------------------------

        fecha_limite = sumar_dias_habiles(
            fecha_inicio=fecha_inicio_datetime,
            dias_contractuales=dias_contractuales,
            excluir_sabados=excluir_sabados,
            excluir_domingos=excluir_domingos,
            festivos=festivos,
            conteo_desde_mismo_dia=(
                conteo_desde_mismo_dia
            ),
            conservar_hora_inicio=(
                conservar_hora_inicio
            ),
        )

        # ----------------------------------------------------
        # DÍAS TRANSCURRIDOS
        # ----------------------------------------------------

        dias_transcurridos = (
            contar_dias_transcurridos(
                fecha_inicio=(
                    fecha_inicio_datetime
                ),
                fecha_actual=(
                    fecha_actual
                ),
                excluir_sabados=(
                    excluir_sabados
                ),
                excluir_domingos=(
                    excluir_domingos
                ),
                festivos=(
                    festivos
                ),
                conteo_desde_mismo_dia=(
                    conteo_desde_mismo_dia
                ),
                conservar_hora_inicio=(
                    conservar_hora_inicio
                ),
            )
        )

        # ----------------------------------------------------
        # DÍAS RESTANTES
        # ----------------------------------------------------

        dias_restantes = (
            calcular_dias_restantes(
                fecha_actual=fecha_actual,
                fecha_limite=fecha_limite,
                excluir_sabados=(
                    excluir_sabados
                ),
                excluir_domingos=(
                    excluir_domingos
                ),
                festivos=festivos,
            )
        )

        # ----------------------------------------------------
        # ESTADO
        # ----------------------------------------------------

        estado = determinar_estado(
            fecha_actual=fecha_actual,
            fecha_limite=fecha_limite,
            dias_transcurridos=(
                dias_transcurridos
            ),
            dias_para_iniciar_alerta=(
                dias_para_alerta
            ),
            dias_restantes=(
                dias_restantes
            ),
        )

        fechas_limite_resultado.append(
            pd.Timestamp(
                fecha_limite
            )
        )

        dias_transcurridos_resultado.append(
            dias_transcurridos
        )

        dias_restantes_resultado.append(
            dias_restantes
        )

        estados_resultado.append(
            estado
        )

    # ========================================================
    # ASIGNAR RESULTADOS
    # ========================================================

    resultado[
        "DIAS_CONTRACTUALES"
    ] = pd.array(
        dias_contractuales_resultado,
        dtype="Int64",
    )

    resultado[
        "FECHA_LIMITE_ANS"
    ] = pd.to_datetime(
        fechas_limite_resultado,
        errors="coerce",
    )

    resultado[
        "DIAS_TRANSCURRIDOS"
    ] = pd.array(
        dias_transcurridos_resultado,
        dtype="Int64",
    )

    resultado[
        "DIAS_PARA_INICIAR_ALERTA"
    ] = pd.array(
        dias_alerta_resultado,
        dtype="Int64",
    )

    resultado[
        "DIAS_RESTANTES"
    ] = pd.array(
        dias_restantes_resultado,
        dtype="Int64",
    )

    resultado[
        "ESTADO"
    ] = estados_resultado

    return resultado
