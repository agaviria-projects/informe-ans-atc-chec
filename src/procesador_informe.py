import logging
from pathlib import Path

from src.generador_excel import generar_informe_excel
from src.lector_excel import cargar_regiones
from src.transformador import transformar_consolidado


logger = logging.getLogger(__name__)


def procesar_informe_ans() -> tuple[Path, dict]:
    """
    Ejecuta el ETL completo de Región 1 y Región 2.
    """

    dataframe_origen, controles_archivos = (
        cargar_regiones()
    )

    dataframe_final, control_transformacion = (
        transformar_consolidado(
            dataframe_origen
        )
    )

    ruta_informe = generar_informe_excel(
        dataframe=dataframe_final,
        controles_archivos=controles_archivos,
        control_transformacion=control_transformacion,
    )

    resumen = {
        "ARCHIVOS_PROCESADOS": len(
            controles_archivos
        ),
        "REGISTROS_GENERADOS": len(
            dataframe_final
        ),
        **control_transformacion,
    }

    logger.info(
        "ETL finalizado correctamente | %s",
        resumen,
    )

    return ruta_informe, resumen