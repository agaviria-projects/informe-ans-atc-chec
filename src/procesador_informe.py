import logging

from pathlib import Path

from src.calculador_ans import aplicar_calculos_ans
from src.generador_excel import generar_informe_excel
from src.lector_excel import cargar_regiones
from src.transformador import transformar_consolidado


logger = logging.getLogger(__name__)


def procesar_informe_ans() -> tuple[Path, dict]:
    """
    Ejecuta el proceso completo del Informe ANS.

    Flujo:

    1. Lee Región 1 y Región 2.
    2. Consolida y limpia los registros.
    3. Aplica las reglas contractuales.
    4. Calcula fechas y estados ANS.
    5. Genera el archivo Excel final.
    """

    # ======================================================
    # 1. LECTURA DE LOS ARCHIVOS REGIONALES
    # ======================================================

    dataframe_origen, controles_archivos = (
        cargar_regiones()
    )

    # ======================================================
    # 2. LIMPIEZA Y TRANSFORMACIÓN
    # ======================================================

    dataframe_transformado, control_transformacion = (
        transformar_consolidado(
            dataframe_origen
        )
    )

    # ======================================================
    # 3. CÁLCULOS ANS
    # ======================================================

    dataframe_final, control_ans = (
        aplicar_calculos_ans(
            dataframe_transformado
        )
    )

    # ======================================================
    # 4. CONTROL COMPLETO DEL PROCESO
    # ======================================================

    control_completo = {
        **control_transformacion,
        **control_ans,
    }

    # ======================================================
    # 5. GENERACIÓN DEL INFORME EXCEL
    # ======================================================

    ruta_informe = generar_informe_excel(
        dataframe=dataframe_final,
        controles_archivos=controles_archivos,
        control_transformacion=control_completo,
    )

    # ======================================================
    # 6. RESUMEN PARA LA INTERFAZ
    # ======================================================

    resumen = {
        "ARCHIVOS_PROCESADOS": len(
            controles_archivos
        ),
        "REGISTROS_GENERADOS": len(
            dataframe_final
        ),
        **control_completo,
    }

    logger.info(
        "ETL y cálculo ANS finalizados correctamente | %s",
        resumen,
    )

    return ruta_informe, resumen