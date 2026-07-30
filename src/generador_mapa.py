import json
import logging
import webbrowser

from datetime import datetime
from html import escape
from pathlib import Path

import pandas as pd

from src.config import (
    HOJA_DATOS_SALIDA,
    LIMITE_DIRECCIONES_MAPA,
    MAPAS_DIR,
    RUTA_INFORME_EXCEL,
    RUTA_MAPA_HTML,
)
from src.geocodificador import (
    ErrorGeocodificacion,
    geocodificar_dataframe,
)


logger = logging.getLogger(__name__)


# ==========================================================
# CENTRO INICIAL DEL MAPA
# ==========================================================

CENTRO_MANIZALES = {
    "latitud": 5.0689,
    "longitud": -75.5174,
    "zoom": 12,
}


# ==========================================================
# COLORES DE LOS ESTADOS ANS
# ==========================================================

COLORES_ESTADO = {
    "A TIEMPO": "#00B050",
    "ALERTA": "#F4D03F",
    "VENCIDO": "#E74C3C",
    "PENDIENTE CONFIGURACIÓN": "#5D6D7E",
}


class ErrorGeneracionMapa(Exception):
    """
    Error controlado durante la generación del mapa.
    """


# ==========================================================
# LIMPIEZA Y NORMALIZACIÓN
# ==========================================================

def limpiar_identificador(
    valor: object,
) -> str:
    """
    Convierte un identificador de Excel en texto limpio.

    También elimina el decimal artificial que puede aparecer
    al leer valores numéricos, por ejemplo: 1003031018.0.
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""

    except (TypeError, ValueError):
        pass

    texto = str(valor).strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def limpiar_valor(
    valor: object,
) -> str:
    """
    Convierte valores vacíos en texto vacío.
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""

    except (TypeError, ValueError):
        pass

    return str(valor).strip()


def normalizar_estado(
    estado: object,
) -> str:
    """
    Normaliza el estado ANS utilizado en el mapa.
    """

    valor = limpiar_valor(
        estado
    ).upper()

    equivalencias = {
        "VENCIDOS": "VENCIDO",
        "A_TIEMPO": "A TIEMPO",
        "PENDIENTE CONFIGURACION": (
            "PENDIENTE CONFIGURACIÓN"
        ),
    }

    valor = equivalencias.get(
        valor,
        valor,
    )

    if valor not in COLORES_ESTADO:
        return "PENDIENTE CONFIGURACIÓN"

    return valor


def interpretar_zona(
    zona: object,
) -> str:
    """
    Convierte el código U/R en su descripción operativa.
    """

    valor = limpiar_valor(
        zona
    ).upper()

    equivalencias = {
        "U": "URBANO",
        "URBANO": "URBANO",
        "R": "RURAL",
        "RURAL": "RURAL",
    }

    return equivalencias.get(
        valor,
        valor or "SIN DEFINIR",
    )


# ==========================================================
# LECTURA DEL INFORME
# ==========================================================

def cargar_pedidos_reales() -> pd.DataFrame:
    """
    Lee la hoja DATOS_ANS del informe consolidado.
    """

    if not RUTA_INFORME_EXCEL.exists():
        raise ErrorGeneracionMapa(
            "No existe Informe_ANS_ELITE.xlsx.\n\n"
            "Genere primero el informe ANS."
        )

    try:
        dataframe = pd.read_excel(
            RUTA_INFORME_EXCEL,
            sheet_name=HOJA_DATOS_SALIDA,
            dtype=object,
            engine="openpyxl",
        )

    except PermissionError as error:
        raise ErrorGeneracionMapa(
            "No fue posible leer Informe_ANS_ELITE.xlsx.\n\n"
            "Cierre el archivo en Excel y vuelva a intentar."
        ) from error

    except ValueError as error:
        raise ErrorGeneracionMapa(
            f"No se encontró la hoja {HOJA_DATOS_SALIDA} "
            "dentro del informe."
        ) from error

    except Exception as error:
        logger.exception(
            "No fue posible leer el informe para generar el mapa."
        )

        raise ErrorGeneracionMapa(
            "No fue posible leer Informe_ANS_ELITE.xlsx."
        ) from error

    columnas_requeridas = {
        "ID_ORDEN",
        "DIRECCION",
        "PROPIETARIO",
        "ZONA",
        "DESC_MUNICIPIO",
        "REGION_ORIGEN",
        "DIAS_RESTANTES",
        "ESTADO",
        "OBSERVACION",
    }

    columnas_faltantes = columnas_requeridas.difference(
        dataframe.columns
    )

    if columnas_faltantes:
        detalle = "\n".join(
            f"- {columna}"
            for columna in sorted(
                columnas_faltantes
            )
        )

        raise ErrorGeneracionMapa(
            "El informe no contiene las columnas necesarias "
            "para generar el mapa:\n\n"
            f"{detalle}"
        )

    dataframe = dataframe[
        dataframe["ID_ORDEN"].notna()
    ].copy()

    dataframe = dataframe[
        dataframe["DIRECCION"].notna()
    ].copy()

    dataframe["DIRECCION"] = (
        dataframe["DIRECCION"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    dataframe = dataframe[
        dataframe["DIRECCION"].ne("")
    ].copy()

    if dataframe.empty:
        raise ErrorGeneracionMapa(
            "El informe no contiene pedidos con dirección."
        )

    return dataframe.reset_index(
        drop=True
    )


# ==========================================================
# PREPARACIÓN DE DATOS PARA LEAFLET
# ==========================================================

def preparar_datos_reales() -> tuple[list[dict], dict]:
    """
    Geocodifica los pedidos reales y prepara los registros
    que serán enviados al mapa Leaflet.
    """

    dataframe = cargar_pedidos_reales()

    dataframe_geocodificado, control = (
        geocodificar_dataframe(
            dataframe=dataframe,
            limite_consultas_nuevas=(
                LIMITE_DIRECCIONES_MAPA
            ),
        )
    )

    dataframe_geocodificado["LATITUD"] = pd.to_numeric(
        dataframe_geocodificado["LATITUD"],
        errors="coerce",
    )

    dataframe_geocodificado["LONGITUD"] = pd.to_numeric(
        dataframe_geocodificado["LONGITUD"],
        errors="coerce",
    )

    ubicaciones_validas = (
        dataframe_geocodificado
        .dropna(
            subset=[
                "LATITUD",
                "LONGITUD",
            ]
        )
        .copy()
    )

    registros: list[dict] = []

    for _, fila in ubicaciones_validas.iterrows():

        latitud = float(
            fila["LATITUD"]
        )

        longitud = float(
            fila["LONGITUD"]
        )

        estado = normalizar_estado(
            fila.get("ESTADO")
        )

        dias_restantes = limpiar_valor(
            fila.get("DIAS_RESTANTES")
        )

        if not dias_restantes:
            dias_restantes = "PENDIENTE"

        registro = {
            "pedido": limpiar_identificador(
                fila.get("ID_ORDEN")
            ),
            "direccion": limpiar_valor(
                fila.get("DIRECCION")
            ),
            "municipio": limpiar_valor(
                fila.get("DESC_MUNICIPIO")
            ),
            "region": limpiar_valor(
                fila.get("REGION_ORIGEN")
            ),
            "zona": interpretar_zona(
                fila.get("ZONA")
            ),
            "propietario": limpiar_valor(
                fila.get("PROPIETARIO")
            ),
            "estado": estado,
            "dias_restantes": dias_restantes,
            "observacion": limpiar_valor(
                fila.get("OBSERVACION")
            ),
            "latitud": latitud,
            "longitud": longitud,
            "color": COLORES_ESTADO.get(
                estado,
                "#5D6D7E",
            ),
            "url_google_maps": (
                "https://www.google.com/maps/search/?api=1"
                f"&query={latitud},{longitud}"
            ),
        }

        registros.append(
            registro
        )

    control["REGISTROS_EN_EL_MAPA"] = len(
        registros
    )

    control["TOTAL_PEDIDOS_INFORME"] = len(
        dataframe
    )

    if not registros:
        raise ErrorGeneracionMapa(
            "No fue posible obtener coordenadas para las "
            "direcciones procesadas.\n\n"
            "Revise el archivo "
            "config/CACHE_GEOCODIFICACION.xlsx."
        )

    return registros, control


# ==========================================================
# CONSTRUCCIÓN DEL HTML
# ==========================================================

def construir_html_mapa(
    registros: list[dict],
) -> str:
    """
    Construye el HTML completo del visor Leaflet.
    """

    if not registros:
        raise ErrorGeneracionMapa(
            "No existen registros para generar el mapa."
        )

    datos_json = json.dumps(
        registros,
        ensure_ascii=False,
    )

    fecha_generacion = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )

    latitud_centro = CENTRO_MANIZALES[
        "latitud"
    ]

    longitud_centro = CENTRO_MANIZALES[
        "longitud"
    ]

    zoom_inicial = CENTRO_MANIZALES[
        "zoom"
    ]

    return f"""<!DOCTYPE html>
<html lang="es">

<head>

    <meta charset="UTF-8">

    <meta
        name="viewport"
        content="width=device-width, initial-scale=1.0"
    >

    <title>Mapa ANS ELITE</title>

    <link
        rel="stylesheet"
        href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
    >

    <script
        src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js">
    </script>

    <style>

        * {{
            box-sizing: border-box;
        }}

        html,
        body {{
            width: 100%;
            height: 100%;
            margin: 0;
            font-family: "Segoe UI", Arial, sans-serif;
            background: #f4f6f7;
        }}

        #map {{
            width: 100%;
            height: 100vh;
        }}

        .panel {{
            position: fixed;
            top: 18px;
            right: 18px;
            width: 340px;
            max-height: calc(100vh - 36px);
            overflow-y: auto;
            padding: 18px;
            border-radius: 14px;
            background: rgba(255, 255, 255, 0.97);
            box-shadow: 0 8px 28px rgba(0, 0, 0, 0.25);
            z-index: 9999;
        }}

        .panel h2 {{
            margin: 0;
            color: #1b2631;
            font-size: 22px;
        }}

        .subtitulo {{
            margin-top: 4px;
            margin-bottom: 16px;
            color: #00843d;
            font-size: 13px;
            font-weight: 600;
        }}

        .panel label {{
            display: block;
            margin-top: 12px;
            margin-bottom: 5px;
            color: #34495e;
            font-size: 13px;
            font-weight: 700;
        }}

        .campo,
        .selector {{
            width: 100%;
            padding: 10px;
            border: 1px solid #aab7b8;
            border-radius: 7px;
            background: #ffffff;
            font-size: 14px;
        }}

        .fila-botones {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 7px;
            margin-top: 8px;
        }}

        .boton {{
            width: 100%;
            padding: 10px 8px;
            border: none;
            border-radius: 7px;
            background: #1e8449;
            color: #ffffff;
            font-size: 13px;
            font-weight: 700;
            cursor: pointer;
        }}

        .boton:hover {{
            filter: brightness(1.08);
        }}

        .boton-secundario {{
            background: #5d6d7e;
        }}

        .boton-azul {{
            background: #2471a3;
        }}

        .boton-copiar {{
            background: #25d366;
        }}

        .estados {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 7px;
            margin-top: 8px;
        }}

        .estado {{
            padding: 9px 5px;
            border: none;
            border-radius: 7px;
            color: #ffffff;
            font-size: 12px;
            font-weight: 700;
            cursor: pointer;
        }}

        .estado-alerta {{
            color: #1b2631;
        }}

        .contador {{
            margin-top: 14px;
            padding: 10px;
            border-radius: 8px;
            background: #eaf2f8;
            color: #1b2631;
            font-weight: 700;
            text-align: center;
        }}

        .nota {{
            margin-top: 12px;
            padding: 9px;
            border-left: 4px solid #2471a3;
            background: #ebf5fb;
            color: #1f618d;
            font-size: 12px;
            line-height: 1.4;
        }}

        .fecha {{
            margin-top: 12px;
            color: #7b7d7d;
            font-size: 11px;
            text-align: center;
        }}

        .popup-ans {{
            min-width: 285px;
            max-width: 380px;
            line-height: 1.45;
        }}

        .popup-ans h3 {{
            margin-top: 0;
            margin-bottom: 10px;
            color: #1b2631;
        }}

        .estado-popup {{
            display: inline-block;
            margin-bottom: 9px;
            padding: 5px 9px;
            border-radius: 5px;
            color: #ffffff;
            font-weight: 700;
        }}

        .dato-popup {{
            margin-bottom: 4px;
        }}

        .observacion-popup {{
            margin-top: 8px;
            padding: 7px;
            border-radius: 5px;
            background: #f4f6f7;
            max-height: 100px;
            overflow-y: auto;
        }}

        .popup-ans a {{
            display: inline-block;
            margin-top: 10px;
            padding: 8px 10px;
            border-radius: 6px;
            background: #2471a3;
            color: #ffffff;
            text-decoration: none;
            font-weight: 700;
        }}

        #toast {{
            display: none;
            position: fixed;
            bottom: 28px;
            left: 50%;
            transform: translateX(-50%);
            padding: 12px 20px;
            border-radius: 8px;
            background: #1e8449;
            color: #ffffff;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.30);
            z-index: 10000;
            font-weight: 600;
        }}

        @media (max-width: 720px) {{

            .panel {{
                top: 10px;
                right: 10px;
                width: calc(100% - 20px);
                max-height: 55vh;
            }}
        }}

    </style>

</head>

<body>

    <div id="map"></div>

    <div class="panel">

        <h2>Mapa ANS</h2>

        <div class="subtitulo">
            ELITE Ingenieros
        </div>

        <label for="buscarPedido">
            Buscar pedido
        </label>

        <input
            id="buscarPedido"
            class="campo"
            type="text"
            placeholder="Ingrese ID_ORDEN"
        >

        <div class="fila-botones">

            <button
                class="boton"
                onclick="buscarPedido()"
            >
                Buscar
            </button>

            <button
                class="boton boton-secundario"
                onclick="limpiarBusqueda()"
            >
                Limpiar
            </button>

        </div>

        <label>
            Filtrar por estado
        </label>

        <div class="estados">

            <button
                class="estado"
                style="background:#00B050"
                onclick="filtrarEstado('A TIEMPO')"
            >
                A TIEMPO
            </button>

            <button
                class="estado estado-alerta"
                style="background:#F4D03F"
                onclick="filtrarEstado('ALERTA')"
            >
                ALERTA
            </button>

            <button
                class="estado"
                style="background:#E74C3C"
                onclick="filtrarEstado('VENCIDO')"
            >
                VENCIDO
            </button>

            <button
                class="estado"
                style="background:#5D6D7E"
                onclick="filtrarEstado(
                    'PENDIENTE CONFIGURACIÓN'
                )"
            >
                PENDIENTE
            </button>

            <button
                class="estado"
                style="
                    background:#34495E;
                    grid-column:1 / span 2;
                "
                onclick="mostrarTodos()"
            >
                MOSTRAR TODOS
            </button>

        </div>

        <label for="regionFiltro">
            Región
        </label>

        <select
            id="regionFiltro"
            class="selector"
            onchange="filtrarRegion()"
        >
            <option value="">
                Todas las regiones
            </option>
        </select>

        <label for="zonaFiltro">
            Zona
        </label>

        <select
            id="zonaFiltro"
            class="selector"
            onchange="filtrarZona()"
        >
            <option value="">
                Todas las zonas
            </option>
        </select>

        <label for="municipioFiltro">
            Municipio
        </label>

        <select
            id="municipioFiltro"
            class="selector"
            onchange="filtrarMunicipio()"
        >
            <option value="">
                Todos los municipios
            </option>
        </select>

        <div class="fila-botones">

            <button
                class="boton boton-azul"
                onclick="centrarMapa()"
            >
                Centrar mapa
            </button>

            <button
                class="boton boton-copiar"
                onclick="copiarPedidoSeleccionado()"
            >
                Copiar datos
            </button>

        </div>

        <div
            id="contador"
            class="contador"
        >
            Registros visibles: 0
        </div>

        <div class="nota">
            El visor muestra únicamente los pedidos que cuentan
            con coordenadas obtenidas a partir de la dirección.
        </div>

        <div class="fecha">
            Generado: {escape(fecha_generacion)}
        </div>

    </div>

    <div id="toast"></div>

    <script>

        const registros = {datos_json};

        const centroInicial = [
            {latitud_centro},
            {longitud_centro}
        ];

        const zoomInicial = {zoom_inicial};

        const map = L.map(
            "map"
        ).setView(
            centroInicial,
            zoomInicial
        );

        L.tileLayer(
            "https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png",
            {{
                maxZoom: 19,
                attribution:
                    "&copy; OpenStreetMap contributors"
            }}
        ).addTo(map);

        let marcadores = [];
        let marcadorSeleccionado = null;

        let estadoActivo = "";
        let regionActiva = "";
        let zonaActiva = "";
        let municipioActivo = "";

        function crearIcono(color) {{

            return L.divIcon({{
                className: "",
                html: `
                    <div style="
                        width:24px;
                        height:24px;
                        border-radius:50% 50% 50% 0;
                        transform:rotate(-45deg);
                        background:${{color}};
                        border:3px solid #ffffff;
                        box-shadow:0 2px 7px rgba(0,0,0,.45);
                    ">
                        <div style="
                            width:7px;
                            height:7px;
                            margin:6px;
                            border-radius:50%;
                            background:#ffffff;
                        ">
                        </div>
                    </div>
                `,
                iconSize: [30, 30],
                iconAnchor: [15, 30],
                popupAnchor: [0, -28]
            }});
        }}

        function escaparHtml(valor) {{

            return String(valor ?? "")
                .replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;")
                .replaceAll('"', "&quot;")
                .replaceAll("'", "&#039;");
        }}

        function construirPopup(registro) {{

            return `
                <div class="popup-ans">

                    <h3>
                        Pedido ${{escaparHtml(registro.pedido)}}
                    </h3>

                    <span
                        class="estado-popup"
                        style="background:${{registro.color}}"
                    >
                        ${{escaparHtml(registro.estado)}}
                    </span>

                    <div class="dato-popup">
                        <b>Dirección:</b>
                        ${{escaparHtml(registro.direccion)}}
                    </div>

                    <div class="dato-popup">
                        <b>Municipio:</b>
                        ${{escaparHtml(registro.municipio)}}
                    </div>

                    <div class="dato-popup">
                        <b>Región:</b>
                        ${{escaparHtml(registro.region)}}
                    </div>

                    <div class="dato-popup">
                        <b>Zona:</b>
                        ${{escaparHtml(registro.zona)}}
                    </div>

                    <div class="dato-popup">
                        <b>Propietario:</b>
                        ${{escaparHtml(registro.propietario)}}
                    </div>

                    <div class="dato-popup">
                        <b>Días restantes:</b>
                        ${{escaparHtml(registro.dias_restantes)}}
                    </div>

                    <div class="observacion-popup">
                        <b>Observación:</b><br>
                        ${{escaparHtml(registro.observacion)}}
                    </div>

                    <a
                        href="${{registro.url_google_maps}}"
                        target="_blank"
                        rel="noopener noreferrer"
                    >
                        Abrir en Google Maps
                    </a>

                </div>
            `;
        }}

        function agregarOpcion(
            selectorId,
            valor
        ) {{

            const selector = document.getElementById(
                selectorId
            );

            const opcion = document.createElement(
                "option"
            );

            opcion.value = valor;
            opcion.textContent = valor;

            selector.appendChild(
                opcion
            );
        }}

        function crearMarcadores() {{

            const regiones = new Set();
            const zonas = new Set();
            const municipios = new Set();

            registros.forEach(registro => {{

                const marcador = L.marker(
                    [
                        registro.latitud,
                        registro.longitud
                    ],
                    {{
                        icon: crearIcono(
                            registro.color
                        )
                    }}
                );

                marcador.registro = registro;

                marcador.bindPopup(
                    construirPopup(registro)
                );

                marcador.on(
                    "click",
                    () => {{
                        marcadorSeleccionado = marcador;
                    }}
                );

                marcador.addTo(map);

                marcadores.push(
                    marcador
                );

                if (registro.region) {{
                    regiones.add(
                        registro.region
                    );
                }}

                if (registro.zona) {{
                    zonas.add(
                        registro.zona
                    );
                }}

                if (registro.municipio) {{
                    municipios.add(
                        registro.municipio
                    );
                }}
            }});

            Array.from(regiones)
                .sort()
                .forEach(valor => {{
                    agregarOpcion(
                        "regionFiltro",
                        valor
                    );
                }});

            Array.from(zonas)
                .sort()
                .forEach(valor => {{
                    agregarOpcion(
                        "zonaFiltro",
                        valor
                    );
                }});

            Array.from(municipios)
                .sort()
                .forEach(valor => {{
                    agregarOpcion(
                        "municipioFiltro",
                        valor
                    );
                }});

            actualizarContador();
        }}

        function cumpleFiltros(marcador) {{

            const registro = marcador.registro;

            const cumpleEstado =
                !estadoActivo
                || registro.estado === estadoActivo;

            const cumpleRegion =
                !regionActiva
                || registro.region === regionActiva;

            const cumpleZona =
                !zonaActiva
                || registro.zona === zonaActiva;

            const cumpleMunicipio =
                !municipioActivo
                || registro.municipio === municipioActivo;

            return (
                cumpleEstado
                && cumpleRegion
                && cumpleZona
                && cumpleMunicipio
            );
        }}

        function aplicarFiltros() {{

            marcadores.forEach(marcador => {{

                if (cumpleFiltros(marcador)) {{

                    if (!map.hasLayer(marcador)) {{
                        marcador.addTo(map);
                    }}

                }} else {{

                    if (map.hasLayer(marcador)) {{
                        map.removeLayer(marcador);
                    }}
                }}
            }});

            actualizarContador();
        }}

        function filtrarEstado(estado) {{

            estadoActivo = estado;
            aplicarFiltros();
        }}

        function filtrarRegion() {{

            regionActiva = document
                .getElementById("regionFiltro")
                .value;

            aplicarFiltros();
        }}

        function filtrarZona() {{

            zonaActiva = document
                .getElementById("zonaFiltro")
                .value;

            aplicarFiltros();
        }}

        function filtrarMunicipio() {{

            municipioActivo = document
                .getElementById("municipioFiltro")
                .value;

            aplicarFiltros();
        }}

        function mostrarTodos() {{

            estadoActivo = "";
            regionActiva = "";
            zonaActiva = "";
            municipioActivo = "";

            document.getElementById(
                "regionFiltro"
            ).value = "";

            document.getElementById(
                "zonaFiltro"
            ).value = "";

            document.getElementById(
                "municipioFiltro"
            ).value = "";

            marcadores.forEach(marcador => {{

                if (!map.hasLayer(marcador)) {{
                    marcador.addTo(map);
                }}
            }});

            actualizarContador();
            centrarMapa();
        }}

        function buscarPedido() {{

            const valor = document
                .getElementById("buscarPedido")
                .value
                .trim()
                .toUpperCase();

            if (!valor) {{

                mostrarToast(
                    "Ingrese un ID_ORDEN."
                );

                return;
            }}

            const encontrado = marcadores.find(
                marcador =>
                    marcador.registro.pedido
                        .toUpperCase() === valor
            );

            if (!encontrado) {{

                mostrarToast(
                    "El pedido no está disponible en el mapa."
                );

                return;
            }}

            marcadores.forEach(marcador => {{

                if (
                    marcador !== encontrado
                    && map.hasLayer(marcador)
                ) {{
                    map.removeLayer(
                        marcador
                    );
                }}
            }});

            if (!map.hasLayer(encontrado)) {{
                encontrado.addTo(map);
            }}

            marcadorSeleccionado = encontrado;

            map.setView(
                encontrado.getLatLng(),
                17
            );

            encontrado.openPopup();

            actualizarContador();
        }}

        function limpiarBusqueda() {{

            document.getElementById(
                "buscarPedido"
            ).value = "";

            marcadorSeleccionado = null;

            mostrarTodos();
        }}

        function centrarMapa() {{

            const visibles = marcadores.filter(
                marcador => map.hasLayer(marcador)
            );

            if (visibles.length === 0) {{

                map.setView(
                    centroInicial,
                    zoomInicial
                );

                return;
            }}

            if (visibles.length === 1) {{

                map.setView(
                    visibles[0].getLatLng(),
                    16
                );

                return;
            }}

            const grupo = L.featureGroup(
                visibles
            );

            map.fitBounds(
                grupo.getBounds().pad(0.20)
            );
        }}

        function actualizarContador() {{

            const visibles = marcadores.filter(
                marcador => map.hasLayer(marcador)
            ).length;

            document.getElementById(
                "contador"
            ).textContent =
                `Registros visibles: ${{visibles}}`;
        }}

        function copiarPedidoSeleccionado() {{

            if (!marcadorSeleccionado) {{

                mostrarToast(
                    "Seleccione primero un marcador."
                );

                return;
            }}

            const registro = (
                marcadorSeleccionado.registro
            );

            const texto = [
                `Pedido: ${{registro.pedido}}`,
                `Estado: ${{registro.estado}}`,
                `Dirección: ${{registro.direccion}}`,
                `Municipio: ${{registro.municipio}}`,
                `Región: ${{registro.region}}`,
                `Zona: ${{registro.zona}}`,
                `Propietario: ${{registro.propietario}}`,
                `Días restantes: ${{registro.dias_restantes}}`,
                `Observación: ${{registro.observacion}}`,
                `Ubicación: ${{registro.url_google_maps}}`
            ].join("\\n");

            navigator.clipboard
                .writeText(texto)
                .then(() => {{

                    mostrarToast(
                        "Información copiada correctamente."
                    );
                }})
                .catch(() => {{

                    mostrarToast(
                        "No fue posible copiar la información."
                    );
                }});
        }}

        function mostrarToast(mensaje) {{

            const toast = document.getElementById(
                "toast"
            );

            toast.textContent = mensaje;
            toast.style.display = "block";

            setTimeout(
                () => {{
                    toast.style.display = "none";
                }},
                2500
            );
        }}

        crearMarcadores();

        setTimeout(
            centrarMapa,
            300
        );

    </script>

</body>

</html>
"""


# ==========================================================
# GENERACIÓN DEL MAPA
# ==========================================================

def generar_mapa_desde_informe(
    abrir_navegador: bool = True,
) -> tuple[Path, dict]:
    """
    Genera el mapa utilizando los pedidos reales del informe.
    """

    try:
        MAPAS_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        registros, control = preparar_datos_reales()

        contenido_html = construir_html_mapa(
            registros
        )

        RUTA_MAPA_HTML.write_text(
            contenido_html,
            encoding="utf-8",
        )

        logger.info(
            "Mapa ANS generado correctamente: %s",
            RUTA_MAPA_HTML,
        )

        if abrir_navegador:
            webbrowser.open(
                RUTA_MAPA_HTML.resolve().as_uri()
            )

        return RUTA_MAPA_HTML, control

    except (
        ErrorGeocodificacion,
        ErrorGeneracionMapa,
    ):
        raise

    except PermissionError as error:
        logger.exception(
            "No fue posible guardar el mapa."
        )

        raise ErrorGeneracionMapa(
            "No fue posible guardar Mapa_ANS_ELITE.html.\n\n"
            "Cierre el archivo o el navegador y vuelva a intentar."
        ) from error

    except OSError as error:
        logger.exception(
            "Error del sistema al guardar el mapa."
        )

        raise ErrorGeneracionMapa(
            "No fue posible guardar el mapa HTML."
        ) from error

    except Exception as error:
        logger.exception(
            "Error inesperado al generar el mapa."
        )

        raise ErrorGeneracionMapa(
            f"No fue posible generar el mapa: {error}"
        ) from error


if __name__ == "__main__":

    ruta_generada, resumen = (
        generar_mapa_desde_informe(
            abrir_navegador=True
        )
    )

    print(
        f"Mapa generado correctamente: {ruta_generada}"
    )

    print(
        resumen
    )