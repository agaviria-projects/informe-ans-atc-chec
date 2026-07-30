import re
import unicodedata

from dataclasses import dataclass

import pandas as pd


# ==========================================================
# RESULTADO DE NORMALIZACIÓN
# ==========================================================

@dataclass(frozen=True)
class DireccionNormalizada:
    """
    Representa una dirección preparada para geocodificación.

    consultas:
        Consultas ordenadas desde la más precisa hasta la menos
        específica.

    tipo:
        Clasificación operativa de la dirección.

    es_geocodificable:
        Indica si existe información suficiente para intentar
        encontrar una ubicación real.

    detalle:
        Información de auditoría del proceso.
    """

    direccion_original: str
    municipio: str
    departamento: str
    consultas: tuple[str, ...]
    tipo: str
    es_geocodificable: bool
    detalle: str


# ==========================================================
# CONFIGURACIÓN DE ABREVIATURAS
# ==========================================================

REEMPLAZOS_ABREVIATURAS = (
    (r"\bCLL\b", "CALLE"),
    (r"\bCL\b", "CALLE"),
    (r"\bCRA\b", "CARRERA"),
    (r"\bCR\b", "CARRERA"),
    (r"\bKRA\b", "CARRERA"),
    (r"\bKR\b", "CARRERA"),
    (r"\bAVDA\b", "AVENIDA"),
    (r"\bAV\b", "AVENIDA"),
    (r"\bDG\b", "DIAGONAL"),
    (r"\bDIAG\b", "DIAGONAL"),
    (r"\bTV\b", "TRANSVERSAL"),
    (r"\bTRANSV\b", "TRANSVERSAL"),
    (r"\bVDA\b", "VEREDA"),
    (r"\bVRD\b", "VEREDA"),
    (r"\bURB\b", "URBANIZACION"),
    (r"\bURBANIZ\b", "URBANIZACION"),
    (r"\bMZ\b", "MANZANA"),
    (r"\bMNZ\b", "MANZANA"),
    (r"\bMANZ\b", "MANZANA"),
    (r"\bLT\b", "LOTE"),
    (r"\bLTE\b", "LOTE"),
    (r"\bLOT\b", "LOTE"),
    (r"\bCAS\b", "CASA"),
    (r"\bPSO\b", "PISO"),
    (r"\bPIS\b", "PISO"),
    (r"\bP\b", "PISO"),
    (r"\bAPTO\b", "APARTAMENTO"),
    (r"\bAPT\b", "APARTAMENTO"),
    (r"\bBLQ\b", "BLOQUE"),
    (r"\bBL\b", "BLOQUE"),
    (r"\bBRR\b", "BARRIO"),
    (r"\bBR\b", "BARRIO"),
    (r"\bCONJ\b", "CONJUNTO"),
    (r"\bED\b", "EDIFICIO"),
    (r"\bEDIF\b", "EDIFICIO"),
    (r"\bTOR\b", "TORRE"),
    (r"\bINT\b", "INTERIOR"),
    (r"\bSEC\b", "SECTOR"),
)


# ==========================================================
# PALABRAS OPERATIVAS O RUIDO
# ==========================================================

PATRONES_RUIDO = (
    r"\bNOA\b",
    r"\bNOCITO\b",
    r"\bZONA\s+RURAL\b",
    r"\bBARRIO\s+SIN\s+DEFINIR\b",
    r"\bBARRIO\s+SIN\s+DEFIN\b",
    r"\bSECTOR\s+SIN\s+DEFINIR\b",
    r"\bSECTOR\s+SIN\s+DEFIN\b",
    r"\bSIN\s+DEFINIR\b",
    r"\bSIN\s+DEFIN\b",
    r"\bSIN\s+DATOS\b",
)


PREFIJOS_VIA = (
    "CALLE",
    "CARRERA",
    "AVENIDA",
    "DIAGONAL",
    "TRANSVERSAL",
)


PALABRAS_TERRITORIALES = (
    "BARRIO",
    "URBANIZACION",
    "SECTOR",
    "VEREDA",
    "CORREGIMIENTO",
    "CONJUNTO",
    "FINCA",
    "HACIENDA",
)


# ==========================================================
# UTILIDADES DE TEXTO
# ==========================================================

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

    texto = str(valor).strip()

    texto = unicodedata.normalize(
        "NFKC",
        texto,
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto,
    )

    return texto.strip()


def eliminar_tildes(
    valor: object,
) -> str:
    """
    Elimina tildes para comparaciones internas.
    """

    texto = limpiar_texto(
        valor
    )

    texto = unicodedata.normalize(
        "NFKD",
        texto,
    )

    return "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )


def normalizar_comparacion(
    valor: object,
) -> str:
    """
    Normaliza texto para comparar valores sin diferencias
    por tildes, mayúsculas o signos.
    """

    texto = eliminar_tildes(
        valor
    ).upper()

    texto = re.sub(
        r"[^A-Z0-9]+",
        " ",
        texto,
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto,
    )

    return texto.strip()


def limpiar_signos(
    texto: str,
) -> str:
    """
    Normaliza separadores conservando la información útil.
    """

    resultado = texto.replace(
        "_",
        " ",
    )

    resultado = resultado.replace(
        ";",
        " ",
    )

    resultado = re.sub(
        r"\s*-\s*",
        " - ",
        resultado,
    )

    resultado = re.sub(
        r"\s+",
        " ",
        resultado,
    )

    return resultado.strip(
        " ,;-"
    )


def expandir_abreviaturas(
    direccion: object,
) -> str:
    """
    Expande abreviaturas frecuentes del archivo fuente.
    """

    resultado = limpiar_texto(
        direccion
    ).upper()

    resultado = limpiar_signos(
        resultado
    )

    for patron, reemplazo in REEMPLAZOS_ABREVIATURAS:
        resultado = re.sub(
            patron,
            reemplazo,
            resultado,
        )

    resultado = re.sub(
        r"\s+",
        " ",
        resultado,
    )

    return resultado.strip()


def eliminar_ruido_operativo(
    direccion: str,
) -> str:
    """
    Elimina términos que no aportan información geográfica.
    """

    resultado = direccion

    for patron in PATRONES_RUIDO:
        resultado = re.sub(
            patron,
            " ",
            resultado,
        )

    resultado = re.sub(
        r"\s+",
        " ",
        resultado,
    )

    return resultado.strip(
        " ,;-"
    )


def eliminar_repeticiones_consecutivas(
    texto: str,
) -> str:
    """
    Elimina palabras o bloques consecutivos repetidos.

    Ejemplo:
        LA MAGDALENA LA MAGDALENA
        VEREDA EL YUNQUE VEREDA EL YUNQUE
    """

    palabras = limpiar_texto(
        texto
    ).split()

    if not palabras:
        return ""

    longitud = len(
        palabras
    )

    for tamano_bloque in range(
        longitud // 2,
        0,
        -1,
    ):
        if longitud < tamano_bloque * 2:
            continue

        primer_bloque = palabras[
            :tamano_bloque
        ]

        segundo_bloque = palabras[
            tamano_bloque:tamano_bloque * 2
        ]

        if (
            normalizar_comparacion(
                " ".join(primer_bloque)
            )
            == normalizar_comparacion(
                " ".join(segundo_bloque)
            )
        ):
            palabras = (
                primer_bloque
                + palabras[tamano_bloque * 2:]
            )

            break

    return " ".join(
        palabras
    ).strip()


def unir_sin_repetir(
    valores: list[str],
) -> str:
    """
    Une referencias evitando valores repetidos.
    """

    resultado: list[str] = []
    claves: set[str] = set()

    for valor in valores:
        texto = limpiar_texto(
            valor
        )

        if not texto:
            continue

        clave = normalizar_comparacion(
            texto
        )

        if not clave:
            continue

        if clave in claves:
            continue

        resultado.append(
            texto
        )

        claves.add(
            clave
        )

    return " ".join(
        resultado
    ).strip()


# ==========================================================
# CLASIFICACIÓN
# ==========================================================

def contiene_via_urbana(
    direccion: str,
) -> bool:
    """
    Identifica direcciones con nomenclatura urbana.
    """

    return direccion.startswith(
        PREFIJOS_VIA
    )


def contiene_referencia_rural(
    direccion: str,
) -> bool:
    """
    Identifica direcciones rurales.
    """

    texto = normalizar_comparacion(
        direccion
    )

    return bool(
        re.search(
            r"\b("
            r"VEREDA|"
            r"FINCA|"
            r"HACIENDA|"
            r"CORREGIMIENTO"
            r")\b",
            texto,
        )
    )


def contiene_sector_localizable(
    direccion: str,
) -> bool:
    """
    Identifica barrios, urbanizaciones, sectores o conjuntos.
    """

    texto = normalizar_comparacion(
        direccion
    )

    expresiones = "|".join(
        PALABRAS_TERRITORIALES
    )

    return bool(
        re.search(
            rf"\b({expresiones})\b",
            texto,
        )
    )


def contiene_referencia_residencial(
    direccion: str,
) -> bool:
    """
    Detecta manzana, casa, lote, bloque, torre o conjunto.
    """

    texto = normalizar_comparacion(
        direccion
    )

    return bool(
        re.search(
            r"\b("
            r"MANZANA|"
            r"CASA|"
            r"LOTE|"
            r"BLOQUE|"
            r"TORRE|"
            r"INTERIOR|"
            r"CONJUNTO|"
            r"URBANIZACION"
            r")\b",
            texto,
        )
    )


# ==========================================================
# PARSEO DE DIRECCIONES URBANAS
# ==========================================================

def compactar_alfanumericos(
    texto: str,
) -> str:
    """
    Compacta números y letras de nomenclatura.

    Ejemplos:
        48 H -> 48H
        5 B  -> 5B
        5 A  -> 5A
    """

    resultado = texto

    resultado = re.sub(
        r"\b(\d+)\s+([A-Z])\b",
        r"\1\2",
        resultado,
    )

    resultado = re.sub(
        r"\b([A-Z])\s+(\d+)\b",
        r"\1\2",
        resultado,
    )

    return resultado


def separar_referencia(
    direccion: str,
) -> tuple[str, str]:
    """
    Separa la parte principal de una referencia posterior.
    """

    partes = [
        parte.strip()
        for parte in direccion.split(
            " - ",
            maxsplit=1,
        )
    ]

    if len(partes) == 1:
        return partes[0], ""

    return partes[0], partes[1]


def limpiar_complemento_urbano(
    complemento: str,
) -> str:
    """
    Retira detalles internos que no mejoran la búsqueda.
    """

    resultado = complemento

    patrones_eliminar = (
        r"\bPISO\s+[A-Z0-9]+\b",
        r"\bAPARTAMENTO\s+[A-Z0-9\-]+\b",
        r"\bINTERIOR\s+[A-Z0-9\-]+\b",
        r"\bTORRE\s+[A-Z0-9\-]+\b",
    )

    for patron in patrones_eliminar:
        resultado = re.sub(
            patron,
            " ",
            resultado,
        )

    resultado = re.sub(
        r"\s+",
        " ",
        resultado,
    )

    resultado = eliminar_repeticiones_consecutivas(
        resultado
    )

    return resultado.strip(
        " ,-"
    )


def extraer_direccion_urbana(
    direccion: str,
) -> tuple[str, str]:
    """
    Intenta convertir una dirección urbana al formato:

        CALLE 48H # 5B-23

    Returns:
        Dirección principal y complemento.
    """

    texto = compactar_alfanumericos(
        direccion
    )

    patron = re.compile(
        r"^(CALLE|CARRERA|AVENIDA|DIAGONAL|TRANSVERSAL)"
        r"\s+([0-9]+[A-Z]?)"
        r"(?:\s+|#)"
        r"([0-9]+[A-Z]?)"
        r"\s+"
        r"([0-9]+)"
        r"(.*)$"
    )

    coincidencia = patron.match(
        texto
    )

    if not coincidencia:
        return "", ""

    tipo_via = coincidencia.group(1)
    numero_via = coincidencia.group(2)
    numero_cruce = coincidencia.group(3)
    numero_predio = coincidencia.group(4)

    complemento = limpiar_complemento_urbano(
        coincidencia.group(5)
    )

    direccion_principal = (
        f"{tipo_via} {numero_via} "
        f"# {numero_cruce}-{numero_predio}"
    )

    return direccion_principal, complemento


def extraer_referencia_urbana(
    direccion: str,
) -> str:
    """
    Intenta extraer barrio, urbanización, conjunto o sector.
    """

    patrones = (
        r"\bBARRIO\s+(.+)$",
        r"\bURBANIZACION\s+(.+)$",
        r"\bCONJUNTO\s+(.+)$",
        r"\bSECTOR\s+(.+)$",
    )

    for patron in patrones:
        coincidencia = re.search(
            patron,
            direccion,
        )

        if coincidencia:
            valor = coincidencia.group(
                1
            ).strip()

            return limpiar_complemento_urbano(
                valor
            )

    return ""


# ==========================================================
# PARSEO DE REFERENCIAS RESIDENCIALES
# ==========================================================

def extraer_nombre_residencial(
    direccion: str,
) -> str:
    """
    Obtiene el nombre principal de una urbanización, conjunto,
    barrio o sector, eliminando detalles internos como lotes,
    casas, manzanas, pisos y bloques.

    Ejemplo:
        URBANIZACION VILLA SANDRA LOTE 10 Y 11
        -> URBANIZACION VILLA SANDRA
    """

    resultado = direccion

    patrones_eliminar = (
        r"\bMANZANA\s+[A-Z0-9\-]+(?:\s+Y\s+[A-Z0-9\-]+)?\b",
        r"\bCASA\s+[A-Z0-9\-]+(?:\s+Y\s+[A-Z0-9\-]+)?\b",
        r"\bLOTE\s+[A-Z0-9\-]+(?:\s+Y\s+[A-Z0-9\-]+)?\b",
        r"\bPISO\s+[A-Z0-9\-]+\b",
        r"\bBLOQUE\s+[A-Z0-9\-]+\b",
        r"\bTORRE\s+[A-Z0-9\-]+\b",
        r"\bINTERIOR\s+[A-Z0-9\-]+\b",
        r"\bAPARTAMENTO\s+[A-Z0-9\-]+\b",
    )

    for patron in patrones_eliminar:
        resultado = re.sub(
            patron,
            " ",
            resultado,
        )

    # Elimina residuos como "Y 11" si quedaron aislados.
    resultado = re.sub(
        r"\bY\s+[A-Z0-9\-]+\b",
        " ",
        resultado,
    )

    resultado = re.sub(
        r"\s+",
        " ",
        resultado,
    )

    resultado = eliminar_repeticiones_consecutivas(
        resultado
    )

    return resultado.strip(
        " ,-"
    )

# ==========================================================
# CONSTRUCCIÓN DE CONSULTAS
# ==========================================================

def agregar_consulta_unica(
    consultas: list[str],
    consulta: str,
) -> None:
    """
    Agrega una consulta evitando duplicados.
    """

    consulta_limpia = limpiar_texto(
        consulta
    )

    if not consulta_limpia:
        return

    clave_nueva = normalizar_comparacion(
        consulta_limpia
    )

    if not clave_nueva:
        return

    claves_existentes = {
        normalizar_comparacion(
            valor
        )
        for valor in consultas
    }

    if clave_nueva not in claves_existentes:
        consultas.append(
            consulta_limpia
        )


def construir_consultas_urbanas(
    direccion: str,
    municipio: str,
    departamento: str,
) -> list[str]:
    """
    Construye consultas progresivas para una dirección urbana.
    """

    consultas: list[str] = []

    principal, referencia_separada = separar_referencia(
        direccion
    )

    direccion_principal, complemento = extraer_direccion_urbana(
        principal
    )

    if not direccion_principal:
        direccion_principal, complemento = extraer_direccion_urbana(
            direccion
        )

    referencia_extraida = extraer_referencia_urbana(
        direccion
    )

    referencia = unir_sin_repetir(
        [
            referencia_separada,
            complemento,
            referencia_extraida,
        ]
    )

    referencia = limpiar_complemento_urbano(
        referencia
    )

    if direccion_principal and referencia:
        agregar_consulta_unica(
            consultas,
            (
                f"{direccion_principal}, "
                f"{referencia}, "
                f"{municipio}, "
                f"{departamento}, Colombia"
            ),
        )

    if direccion_principal:
        agregar_consulta_unica(
            consultas,
            (
                f"{direccion_principal}, "
                f"{municipio}, "
                f"{departamento}, Colombia"
            ),
        )

    if referencia:
        agregar_consulta_unica(
            consultas,
            (
                f"{referencia}, "
                f"{municipio}, "
                f"{departamento}, Colombia"
            ),
        )

    if direccion_principal:
        via_sin_predio = re.sub(
            r"\s+#\s+.+$",
            "",
            direccion_principal,
        )

        agregar_consulta_unica(
            consultas,
            (
                f"{via_sin_predio}, "
                f"{municipio}, "
                f"{departamento}, Colombia"
            ),
        )

    return consultas


def construir_consultas_rurales(
    direccion: str,
    municipio: str,
    departamento: str,
) -> list[str]:
    """
    Construye consultas para veredas y sectores rurales.

    Evita repetir referencias equivalentes y nunca utiliza
    únicamente el centro del municipio.
    """

    consultas: list[str] = []

    partes = [
        parte.strip()
        for parte in direccion.split(" - ")
        if parte.strip()
    ]

    partes_unicas: list[str] = []
    claves_partes: set[str] = set()

    for parte in partes:

        clave = normalizar_comparacion(
            parte
        )

        if not clave:
            continue

        if clave in claves_partes:
            continue

        partes_unicas.append(
            parte
        )

        claves_partes.add(
            clave
        )

    # Si ambas partes representan la misma vereda,
    # se conserva únicamente una.
    referencias_filtradas: list[str] = []

    for parte in partes_unicas:

        clave_parte = normalizar_comparacion(
            parte
        )

        es_repetida = any(
            clave_parte in normalizar_comparacion(
                existente
            )
            or normalizar_comparacion(
                existente
            ) in clave_parte
            for existente in referencias_filtradas
        )

        if not es_repetida:
            referencias_filtradas.append(
                parte
            )

    direccion_limpia = " - ".join(
        referencias_filtradas
    )

    direccion_limpia = eliminar_repeticiones_consecutivas(
        direccion_limpia
    )

    if direccion_limpia:
        agregar_consulta_unica(
            consultas,
            (
                f"{direccion_limpia}, "
                f"{municipio}, "
                f"{departamento}, Colombia"
            ),
        )

    for parte in referencias_filtradas:

        if (
            contiene_referencia_rural(parte)
            or contiene_sector_localizable(parte)
        ):
            agregar_consulta_unica(
                consultas,
                (
                    f"{parte}, "
                    f"{municipio}, "
                    f"{departamento}, Colombia"
                ),
            )

    return consultas


def construir_consultas_residenciales(
    direccion: str,
    municipio: str,
    departamento: str,
) -> list[str]:
    """
    Construye consultas para manzanas, lotes, casas,
    urbanizaciones y conjuntos.
    """

    consultas: list[str] = []

    direccion_limpia = eliminar_repeticiones_consecutivas(
        direccion
    )

    nombre_residencial = extraer_nombre_residencial(
        direccion_limpia
    )

    agregar_consulta_unica(
        consultas,
        (
            f"{direccion_limpia}, "
            f"{municipio}, "
            f"{departamento}, Colombia"
        ),
    )

    if nombre_residencial:
        agregar_consulta_unica(
            consultas,
            (
                f"{nombre_residencial}, "
                f"{municipio}, "
                f"{departamento}, Colombia"
            ),
        )

    return consultas


def construir_consultas_referencia(
    direccion: str,
    municipio: str,
    departamento: str,
) -> list[str]:
    """
    Construye consultas para barrios, sectores o urbanizaciones.
    """

    consultas: list[str] = []

    direccion_limpia = eliminar_repeticiones_consecutivas(
        direccion
    )

    agregar_consulta_unica(
        consultas,
        (
            f"{direccion_limpia}, "
            f"{municipio}, "
            f"{departamento}, Colombia"
        ),
    )

    partes = [
        parte.strip()
        for parte in direccion_limpia.split(
            " - "
        )
        if parte.strip()
    ]

    for parte in partes:
        agregar_consulta_unica(
            consultas,
            (
                f"{parte}, "
                f"{municipio}, "
                f"{departamento}, Colombia"
            ),
        )

    return consultas


# ==========================================================
# FUNCIÓN PRINCIPAL
# ==========================================================

def normalizar_direccion(
    direccion: object,
    municipio: object,
    departamento: str = "Caldas",
) -> DireccionNormalizada:
    """
    Normaliza una dirección y construye búsquedas progresivas.

    Nunca genera una consulta que contenga únicamente el
    municipio, porque eso produciría una ubicación aproximada
    y no la localización real del pedido.
    """

    direccion_original = limpiar_texto(
        direccion
    )

    municipio_limpio = limpiar_texto(
        municipio
    )

    departamento_limpio = (
        limpiar_texto(
            departamento
        )
        or "Caldas"
    )

    if not direccion_original:
        return DireccionNormalizada(
            direccion_original="",
            municipio=municipio_limpio,
            departamento=departamento_limpio,
            consultas=(),
            tipo="DIRECCION_VACIA",
            es_geocodificable=False,
            detalle="La dirección está vacía.",
        )

    if not municipio_limpio:
        return DireccionNormalizada(
            direccion_original=direccion_original,
            municipio="",
            departamento=departamento_limpio,
            consultas=(),
            tipo="MUNICIPIO_VACIO",
            es_geocodificable=False,
            detalle=(
                "No se puede validar la dirección porque "
                "el municipio está vacío."
            ),
        )

    direccion_expandida = expandir_abreviaturas(
        direccion_original
    )

    direccion_expandida = eliminar_ruido_operativo(
        direccion_expandida
    )

    direccion_expandida = eliminar_repeticiones_consecutivas(
        direccion_expandida
    )

    if not direccion_expandida:
        return DireccionNormalizada(
            direccion_original=direccion_original,
            municipio=municipio_limpio,
            departamento=departamento_limpio,
            consultas=(),
            tipo="SIN_INFORMACION_UTIL",
            es_geocodificable=False,
            detalle=(
                "La dirección solo contenía expresiones "
                "operativas sin información geográfica."
            ),
        )

    if contiene_via_urbana(
        direccion_expandida
    ):
        tipo = "URBANA"

        consultas = construir_consultas_urbanas(
            direccion=direccion_expandida,
            municipio=municipio_limpio,
            departamento=departamento_limpio,
        )

    elif contiene_referencia_rural(
        direccion_expandida
    ):
        tipo = "RURAL"

        consultas = construir_consultas_rurales(
            direccion=direccion_expandida,
            municipio=municipio_limpio,
            departamento=departamento_limpio,
        )

    elif contiene_referencia_residencial(
        direccion_expandida
    ):
        tipo = "RESIDENCIAL"

        consultas = construir_consultas_residenciales(
            direccion=direccion_expandida,
            municipio=municipio_limpio,
            departamento=departamento_limpio,
        )

    elif contiene_sector_localizable(
        direccion_expandida
    ):
        tipo = "REFERENCIA_TERRITORIAL"

        consultas = construir_consultas_referencia(
            direccion=direccion_expandida,
            municipio=municipio_limpio,
            departamento=departamento_limpio,
        )

    else:
        tipo = "NO_CLASIFICADA"
        consultas = []

        agregar_consulta_unica(
            consultas,
            (
                f"{direccion_expandida}, "
                f"{municipio_limpio}, "
                f"{departamento_limpio}, Colombia"
            ),
        )

    consultas_validas = tuple(
        consulta
        for consulta in consultas
        if consulta
    )

    if not consultas_validas:
        return DireccionNormalizada(
            direccion_original=direccion_original,
            municipio=municipio_limpio,
            departamento=departamento_limpio,
            consultas=(),
            tipo=tipo,
            es_geocodificable=False,
            detalle=(
                "No fue posible construir una consulta "
                "geográfica suficientemente específica."
            ),
        )

    return DireccionNormalizada(
        direccion_original=direccion_original,
        municipio=municipio_limpio,
        departamento=departamento_limpio,
        consultas=consultas_validas,
        tipo=tipo,
        es_geocodificable=True,
        detalle=(
            f"Se generaron {len(consultas_validas)} "
            "consultas progresivas."
        ),
    )


# ==========================================================
# PRUEBA LOCAL
# ==========================================================

if __name__ == "__main__":

    ejemplos = [
        (
            "CLL 48 H 5 B 23 PSO 2 - BENGALA",
            "Manizales",
        ),
        (
            "CRA 9 5A 08 LA MAGDALENA - LA MAGDALENA",
            "La Dorada",
        ),
        (
            "VDA EL YUNQUE - Vereda El Yunque",
            "Neira",
        ),
        (
            "NOA MZ 1 CASA 5 PISO 2 VILLA DEL PRADO",
            "Chinchina",
        ),
        (
            "URB VILLA SANDRA LOT 10 Y 11",
            "Chinchina",
        ),
    ]

    for direccion_ejemplo, municipio_ejemplo in ejemplos:

        resultado = normalizar_direccion(
            direccion=direccion_ejemplo,
            municipio=municipio_ejemplo,
        )

        print(
            "\n"
            f"ORIGINAL: {resultado.direccion_original}\n"
            f"TIPO: {resultado.tipo}\n"
            f"GEOCODIFICABLE: {resultado.es_geocodificable}\n"
            f"DETALLE: {resultado.detalle}"
        )

        for numero, consulta in enumerate(
            resultado.consultas,
            start=1,
        ):
            print(
                f"  {numero}. {consulta}"
            )