import json
import logging
import webbrowser
from datetime import datetime
from html import escape
from pathlib import Path

from src.config import RUTA_MAPA_HTML, SALIDA_DIR


logger = logging.getLogger(__name__)


# ==========================================================
# CENTRO INICIAL DEL MAPA
# ==========================================================

CENTRO_MANIZALES = {
    "latitud": 5.0689,
    "longitud": -75.5174,
    "zoom": 13,
}


# ==========================================================
# DATOS PROVISIONALES DE DEMOSTRACIÓN
# ==========================================================

DATOS_MAPA_DEMO = [
    {
        "pedido": "DEMO-001",
        "direccion": "Sector Centro - Dirección demostrativa",
        "municipio": "MANIZALES",
        "actividad": "INSTALACIÓN",
        "estado": "A TIEMPO",
        "cliente": "CLIENTE DEMOSTRACIÓN 1",
        "telefono": "3000000001",
        "dias_restantes": "5 días",
        "latitud": 5.0689,
        "longitud": -75.5174,
    },
    {
        "pedido": "DEMO-002",
        "direccion": "Sector Palogrande - Dirección demostrativa",
        "municipio": "MANIZALES",
        "actividad": "REVISIÓN",
        "estado": "ALERTA",
        "cliente": "CLIENTE DEMOSTRACIÓN 2",
        "telefono": "3000000002",
        "dias_restantes": "2 días",
        "latitud": 5.0567,
        "longitud": -75.4918,
    },
    {
        "pedido": "DEMO-003",
        "direccion": "Sector Chipre - Dirección demostrativa",
        "municipio": "MANIZALES",
        "actividad": "MANTENIMIENTO",
        "estado": "ALERTA 0 DÍAS",
        "cliente": "CLIENTE DEMOSTRACIÓN 3",
        "telefono": "3000000003",
        "dias_restantes": "0 días",
        "latitud": 5.0738,
        "longitud": -75.5334,
    },
    {
        "pedido": "DEMO-004",
        "direccion": "Sector La Enea - Dirección demostrativa",
        "municipio": "MANIZALES",
        "actividad": "INSTALACIÓN",
        "estado": "VENCIDO",
        "cliente": "CLIENTE DEMOSTRACIÓN 4",
        "telefono": "3000000004",
        "dias_restantes": "VENCIDO",
        "latitud": 5.0358,
        "longitud": -75.4696,
    },
    {
        "pedido": "DEMO-005",
        "direccion": "Sector Villamaría - Dirección demostrativa",
        "municipio": "VILLAMARÍA",
        "actividad": "REVISIÓN",
        "estado": "SIN FECHA",
        "cliente": "CLIENTE DEMOSTRACIÓN 5",
        "telefono": "3000000005",
        "dias_restantes": "SIN FECHA",
        "latitud": 5.0446,
        "longitud": -75.5143,
    },
]


# ==========================================================
# COLORES DE ESTADO
# ==========================================================

COLORES_ESTADO = {
    "A TIEMPO": "#00B050",
    "ALERTA": "#F4D03F",
    "ALERTA 0 DÍAS": "#F39C12",
    "VENCIDO": "#E74C3C",
    "SIN FECHA": "#5D6D7E",
}


class ErrorGeneracionMapa(Exception):
    """Error controlado durante la generación del mapa."""


def normalizar_estado(estado: str) -> str:
    """
    Normaliza el nombre del estado para filtros y colores.
    """

    valor = str(estado).strip().upper()

    equivalencias = {
        "ALERTA_0 DIAS": "ALERTA 0 DÍAS",
        "ALERTA 0 DIAS": "ALERTA 0 DÍAS",
        "ALERTA_0 DÍAS": "ALERTA 0 DÍAS",
    }

    return equivalencias.get(valor, valor)


def preparar_datos_demo() -> list[dict]:
    """
    Prepara los datos demostrativos para enviarlos al mapa.
    """

    datos_preparados: list[dict] = []

    for registro in DATOS_MAPA_DEMO:
        item = registro.copy()

        item["estado"] = normalizar_estado(
            item.get("estado", "")
        )

        item["color"] = COLORES_ESTADO.get(
            item["estado"],
            "#5D6D7E",
        )

        item["url_google_maps"] = (
            "https://www.google.com/maps/search/?api=1"
            f"&query={item['latitud']},{item['longitud']}"
        )

        datos_preparados.append(item)

    return datos_preparados


def construir_html_mapa(
    registros: list[dict],
) -> str:
    """
    Construye el HTML completo del mapa Leaflet.
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

    latitud_centro = CENTRO_MANIZALES["latitud"]
    longitud_centro = CENTRO_MANIZALES["longitud"]
    zoom_inicial = CENTRO_MANIZALES["zoom"]

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
            width: 330px;
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

        .panel .subtitulo {{
            margin-top: 4px;
            margin-bottom: 16px;
            color: #00843d;
            font-weight: 600;
            font-size: 13px;
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

        .boton-whatsapp {{
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

        .nota-demo {{
            margin-top: 12px;
            padding: 9px;
            border-left: 4px solid #f39c12;
            background: #fef5e7;
            color: #7d6608;
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
            min-width: 260px;
            line-height: 1.45;
        }}

        .popup-ans h3 {{
            margin-top: 0;
            margin-bottom: 10px;
            color: #1b2631;
        }}

        .popup-ans .estado-popup {{
            display: inline-block;
            margin-bottom: 9px;
            padding: 5px 9px;
            border-radius: 5px;
            color: #ffffff;
            font-weight: 700;
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
                max-height: 52vh;
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
            placeholder="Ejemplo: DEMO-001"
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
                style="background:#F39C12"
                onclick="filtrarEstado('ALERTA 0 DÍAS')"
            >
                ALERTA 0 DÍAS
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
                onclick="filtrarEstado('SIN FECHA')"
            >
                SIN FECHA
            </button>

            <button
                class="estado"
                style="background:#34495E"
                onclick="mostrarTodos()"
            >
                TODOS
            </button>

        </div>

        <label for="actividadFiltro">
            Actividad
        </label>

        <select
            id="actividadFiltro"
            class="selector"
            onchange="filtrarActividad()"
        >
            <option value="">
                Todas las actividades
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
                class="boton boton-whatsapp"
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

        <div class="nota-demo">
            Este mapa utiliza cinco ubicaciones demostrativas.
            Las direcciones y coordenadas deberán reemplazarse
            cuando se reciba el CSV oficial.
        </div>

        <div class="fecha">
            Generado: {escape(fecha_generacion)}
        </div>

    </div>

    <div id="toast">
        Información copiada correctamente.
    </div>

    <script>

        const registros = {datos_json};

        const centroInicial = [
            {latitud_centro},
            {longitud_centro}
        ];

        const zoomInicial = {zoom_inicial};

        const map = L.map("map").setView(
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
        let actividadActiva = "";

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
                        "></div>
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

                    <div>
                        <b>Dirección:</b>
                        ${{escaparHtml(registro.direccion)}}
                    </div>

                    <div>
                        <b>Municipio:</b>
                        ${{escaparHtml(registro.municipio)}}
                    </div>

                    <div>
                        <b>Actividad:</b>
                        ${{escaparHtml(registro.actividad)}}
                    </div>

                    <div>
                        <b>Cliente:</b>
                        ${{escaparHtml(registro.cliente)}}
                    </div>

                    <div>
                        <b>Teléfono:</b>
                        ${{escaparHtml(registro.telefono)}}
                    </div>

                    <div>
                        <b>Días restantes:</b>
                        ${{escaparHtml(registro.dias_restantes)}}
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

        function crearMarcadores() {{

            const actividades = new Set();

            registros.forEach(registro => {{

                const marcador = L.marker(
                    [
                        registro.latitud,
                        registro.longitud
                    ],
                    {{
                        icon: crearIcono(registro.color)
                    }}
                );

                marcador.registro = registro;

                marcador.bindPopup(
                    construirPopup(registro)
                );

                marcador.on("click", () => {{
                    marcadorSeleccionado = marcador;
                }});

                marcador.addTo(map);

                marcadores.push(marcador);

                actividades.add(
                    registro.actividad
                );
            }});

            const selector = document.getElementById(
                "actividadFiltro"
            );

            Array.from(actividades)
                .sort()
                .forEach(actividad => {{

                    const opcion = document.createElement(
                        "option"
                    );

                    opcion.value = actividad;
                    opcion.textContent = actividad;

                    selector.appendChild(opcion);
                }});

            actualizarContador();
        }}

        function cumpleFiltros(marcador) {{

            const registro = marcador.registro;

            const cumpleEstado =
                !estadoActivo
                || registro.estado === estadoActivo;

            const cumpleActividad =
                !actividadActiva
                || registro.actividad === actividadActiva;

            return cumpleEstado && cumpleActividad;
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

        function filtrarActividad() {{

            actividadActiva = document
                .getElementById("actividadFiltro")
                .value;

            aplicarFiltros();
        }}

        function mostrarTodos() {{

            estadoActivo = "";
            actividadActiva = "";

            document.getElementById(
                "actividadFiltro"
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
                    "Ingrese un número de pedido."
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
                    "Pedido no encontrado."
                );

                return;
            }}

            marcadores.forEach(marcador => {{

                if (marcador !== encontrado) {{

                    if (map.hasLayer(marcador)) {{
                        map.removeLayer(marcador);
                    }}
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

            const registro =
                marcadorSeleccionado.registro;

            const texto = [
                `Pedido: ${{registro.pedido}}`,
                `Estado: ${{registro.estado}}`,
                `Dirección: ${{registro.direccion}}`,
                `Municipio: ${{registro.municipio}}`,
                `Actividad: ${{registro.actividad}}`,
                `Cliente: ${{registro.cliente}}`,
                `Teléfono: ${{registro.telefono}}`,
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

            setTimeout(() => {{
                toast.style.display = "none";
            }}, 2500);
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


def generar_mapa_demo(
    abrir_navegador: bool = True,
) -> Path:
    """
    Genera el mapa provisional con cinco registros demostrativos.

    Args:
        abrir_navegador:
            Abre automáticamente el mapa al finalizar.

    Returns:
        Ruta del archivo HTML generado.
    """

    try:
        SALIDA_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        registros = preparar_datos_demo()

        contenido_html = construir_html_mapa(
            registros
        )

        RUTA_MAPA_HTML.write_text(
            contenido_html,
            encoding="utf-8",
        )

        logger.info(
            "Mapa demostrativo generado correctamente: %s",
            RUTA_MAPA_HTML,
        )

        if abrir_navegador:
            webbrowser.open(
                RUTA_MAPA_HTML.resolve().as_uri()
            )

        return RUTA_MAPA_HTML

    except OSError as error:
        logger.exception(
            "Error al guardar el mapa demostrativo."
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
    ruta = generar_mapa_demo(
        abrir_navegador=True
    )

    print(
        f"Mapa generado correctamente: {ruta}"
    )