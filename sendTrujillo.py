#sendTrujillo.py
import pandas as pd
import os
import logging
from conexion import ManejadorConexionSQL
# Asumo que este archivo existe
from consultas import QUERY_RESULTADOS, QUERY_ENCABEZADOS
from Pendientes import Pendientes
from generador_excel import crear_excel_trujillo, crear_excel_olmos
#from manejador_correo import enviar_correo_con_adjunto, crear_cuerpo_html_correo
from manejador_correo import enviar_correo_con_adjunto_yagmail, crear_cuerpo_html_correo # Cambiado aquí


from configuracion_correo import (
    LOGS_DIR, LOG_EXITOS_FILE, LOG_ERRORES_FILE, LOG_FORMAT,
    DESTINATARIO_TO_POR_DEFECTO, DESTINATARIO_CC_POR_DEFECTO, DESTINATARIO_BCC_POR_DEFECTO,
    ASUNTO_PLANTILLA
)
from datetime import datetime


def formatear_fecha_mejorado(valor_fecha_original, formato_salida_deseado):
    """
    Parsea una cadena de fecha (DD/MM/YYYY o DD/MM/YYYY HH:MM:SS) y la reformatea.
    Devuelve una cadena vacía si la entrada es None o una cadena vacía.
    Devuelve la cadena original si no se puede parsear.
    """
    if valor_fecha_original is None:
        # print(f"DEBUG: formatear_fecha - Valor es None. Devolviendo ''.")
        return ''

    # Asegurarse que es una cadena para trabajar con ella
    # Convertir a str y quitar espacios extra
    valor_fecha_str = str(valor_fecha_original).strip()

    if not valor_fecha_str:  # Si después de strip queda vacía
        # print(f"DEBUG: formatear_fecha - Valor es cadena vacía. Devolviendo ''.")
        return ''

    # print(f"DEBUG: formatear_fecha - Procesando: '{valor_fecha_str}', Formato Salida: '{formato_salida_deseado}'")

    fecha_parte_str = valor_fecha_str
    # Intentar obtener solo la parte de la fecha si hay un espacio (indicando hora)
    if ' ' in valor_fecha_str:
        fecha_parte_str = valor_fecha_str.split(' ')[0]

    try:
        # Intentar parsear como DD/MM/YYYY
        dt_obj = datetime.strptime(fecha_parte_str, '%d/%m/%Y')
        resultado_formateado = dt_obj.strftime(formato_salida_deseado)
        # print(f"DEBUG: formatear_fecha - Parseo exitoso. Formateado a: '{resultado_formateado}'")
        return resultado_formateado
    except ValueError as e:
        print(
            f"⚠️ No se pudo parsear la fecha '{fecha_parte_str}' (original: '{valor_fecha_original}') como DD/MM/YYYY. Error: {e}. Devolviendo original.")
        return valor_fecha_original  # Devolver el valor original si el parseo falla
    except Exception as ex:  # Otros errores inesperados
        print(
            f"⚠️ Error inesperado en formatear_fecha con valor '{valor_fecha_original}': {ex}. Devolviendo original.")
        return valor_fecha_original


def configurar_logging():
    """Configura el sistema de logging para escribir en archivos."""
    if not os.path.exists(LOGS_DIR):
        os.makedirs(LOGS_DIR)

    # Configuración básica del logger raíz
    logging.basicConfig(
        level=logging.INFO,
        format=LOG_FORMAT,
        handlers=[
            logging.StreamHandler()  # Para ver logs en consola también
        ]
    )

    # Handler para archivo de éxitos
    exitos_handler = logging.FileHandler(
        os.path.join(LOGS_DIR, LOG_EXITOS_FILE), mode='a')
    exitos_handler.setLevel(logging.INFO)  # Solo logs INFO y superiores
    exitos_formatter = logging.Formatter(LOG_FORMAT)
    exitos_handler.setFormatter(exitos_formatter)
    # Filtro para que solo los mensajes de éxito vayan a este archivo (opcional, si quieres ser más granular)
    # class InfoFilter(logging.Filter):
    #     def filter(self, record):
    #         return record.levelno == logging.INFO
    # exitos_handler.addFilter(InfoFilter())
    logging.getLogger().addHandler(exitos_handler)

    # Handler para archivo de errores
    errores_handler = logging.FileHandler(
        os.path.join(LOGS_DIR, LOG_ERRORES_FILE), mode='a')
    errores_handler.setLevel(logging.WARNING)  # Logs WARNING, ERROR, CRITICAL
    errores_formatter = logging.Formatter(LOG_FORMAT)
    errores_handler.setFormatter(errores_formatter)
    logging.getLogger().addHandler(errores_handler)

    logging.info("Sistema de Logging configurado.")


def procesar_un_pendiente(manejador_db, cdamostra_actual):
    """
    Procesa un único cdamostra, obteniendo sus resultados y encabezado.
    Retorna un diccionario con los datos o None si hay error crítico.
    """
    logging.info(
        f"--- Iniciando procesamiento para CDAMOSTRA: {cdamostra_actual} ---")
    datos_procesados = {
        "cdamostra": cdamostra_actual,
        "encabezado": None,
        "resultados": []
    }

    # Obtener resultados
    resultados_data = manejador_db.ejecutar_consulta(
        QUERY_RESULTADOS, (cdamostra_actual,))
    if resultados_data is not None:  # Puede ser lista vacía si no hay resultados, no es error
        logging.info(
            f"Resultados obtenidos para {cdamostra_actual}: {len(resultados_data)} filas.")
        datos_procesados["resultados"] = resultados_data
    else:
        logging.error(f"Error al obtener resultados para {cdamostra_actual}.")
        # Podrías decidir si continuar sin resultados o retornar None aquí
        # return None

    # Obtener encabezado
    encabezado_data_lista = manejador_db.ejecutar_consulta(
        QUERY_ENCABEZADOS, (cdamostra_actual,))
    if encabezado_data_lista and len(encabezado_data_lista) == 1:
        logging.info(f"Encabezado obtenido para {cdamostra_actual}.")
        datos_procesados["encabezado"] = encabezado_data_lista[0]
    elif encabezado_data_lista:  # Más de 1 o 0 filas
        logging.warning(
            f"Se esperaba 1 fila para el encabezado de {cdamostra_actual}, pero se obtuvieron {len(encabezado_data_lista)}. Usando la primera si existe.")
        datos_procesados["encabezado"] = encabezado_data_lista[0]
    else:  # None o lista vacía
        logging.error(f"Error al obtener encabezado para {cdamostra_actual}.")
        # Si el encabezado es crucial, podrías retornar None
        # return None

    # Formatear fechas en el encabezado para el cuerpo del correo
    if datos_procesados["encabezado"]:
        datos_procesados["encabezado"]["datacoleta_formateada"] = formatear_fecha_mejorado(
            datos_procesados["encabezado"].get('datacoleta'), '%d/%m/%Y')
        datos_procesados["encabezado"]["datachegada_formateada"] = formatear_fecha_mejorado(
            datos_procesados["encabezado"].get('datachegada'), '%d/%m/%Y')

    return datos_procesados


if __name__ == "__main__":
    configurar_logging()  # Configurar logging al inicio
    logging.info(
        "********** INICIO DEL PROCESO DE ENVÍO DE INFORMES **********")

    # --- Parámetros para el correo (puedes obtenerlos de otro lado si es necesario) ---
    # Estos sobrescribirán los valores por defecto de configuracion_correo.py si se proporcionan


    # Flag para decidir qué Excel generar (1 para Trujillo, otro valor para Olmos)
    tipo_informe_flag = 1  # 1: Trujillo, Otro: Olmos (ej. 2)

    manejador_mylims = ManejadorConexionSQL("myLIMS_Novo_conn")
    conexion_activa = manejador_mylims.conectar()

    if conexion_activa:
        gestor_pendientes = Pendientes(manejador_mylims)
        cdamostras_pendientes = gestor_pendientes.obtener_pendientes()

        if cdamostras_pendientes:
            logging.info(
                f"Se encontraron {len(cdamostras_pendientes)} CDAMOSTRAS pendientes para procesar.")

            for cdamostra_id_pendiente in cdamostras_pendientes:
                datos_un_pendiente = procesar_un_pendiente(
                    manejador_mylims, cdamostra_id_pendiente)

                if not datos_un_pendiente or (not datos_un_pendiente.get("encabezado") and not datos_un_pendiente.get("resultados")):
                    logging.error(
                        f"No se pudieron obtener datos suficientes para CDAMOSTRA {cdamostra_id_pendiente}. Se omite este pendiente.")
                    continue

                ruta_excel_adjuntar = None
                nombre_excel_base = ""

                if tipo_informe_flag == 1:
                    # Directorio donde se guardarán los excels (ej. subcarpeta 'informes_excel')
                    directorio_salida_excel = "informes_generados"
                    nombre_excel_base = f"Trujillo_Muestra_{datos_un_pendiente['cdamostra']}.xlsx"
                    ruta_completa_excel_t = os.path.join(
                        directorio_salida_excel, nombre_excel_base)

                    ruta_excel_adjuntar = crear_excel_trujillo(
                        datos_un_pendiente.get("encabezado"),
                        datos_un_pendiente.get("resultados", []),
                        ruta_completa_excel_t
                    )
                else:  # Asumimos Olmos
                    directorio_salida_excel = "informes_generados"
                    nombre_excel_base = f"Olmos_Muestra_{datos_un_pendiente['cdamostra']}.xlsx"
                    ruta_completa_excel_o = os.path.join(
                        directorio_salida_excel, nombre_excel_base)

                    ruta_excel_adjuntar = crear_excel_olmos(
                        datos_un_pendiente.get("encabezado"),
                        datos_un_pendiente.get("resultados", []),
                        ruta_completa_excel_o
                    )

                if ruta_excel_adjuntar:
                    logging.info(
                        f"Excel '{ruta_excel_adjuntar}' generado exitosamente para CDAMOSTRA {datos_un_pendiente['cdamostra']}.")

                    # Preparar y enviar correo
                    asunto_correo = ASUNTO_PLANTILLA.format(
                        cdamostra=datos_un_pendiente['cdamostra'])

                    cuerpo_html_email = crear_cuerpo_html_correo(
                        datos_un_pendiente['cdamostra'],
                        # Pasar dict vacío si es None
                        datos_un_pendiente.get("encabezado", {}),
                        datos_un_pendiente.get("resultados", [])
                    )

                    envio_exitoso = enviar_correo_con_adjunto_yagmail(
                        destinatarios_to= DESTINATARIO_TO_POR_DEFECTO,
                        asunto=asunto_correo,
                        cuerpo_html=cuerpo_html_email,
                        ruta_archivo_adjunto=ruta_excel_adjuntar,
                        destinatarios_cc= DESTINATARIO_CC_POR_DEFECTO,
                        destinatarios_bcc=DESTINATARIO_BCC_POR_DEFECTO
                    )
                    if envio_exitoso:
                        logging.info(
                            f"Correo para CDAMOSTRA {datos_un_pendiente['cdamostra']} enviado con adjunto '{nombre_excel_base}'.")
                    else:
                        logging.error(
                            f"Fallo al enviar correo para CDAMOSTRA {datos_un_pendiente['cdamostra']} con adjunto '{nombre_excel_base}'.")
                else:
                    logging.error(
                        f"No se pudo generar el archivo Excel para CDAMOSTRA {datos_un_pendiente['cdamostra']}. No se enviará correo.")

            logging.info(
                "--- Proceso completado para todos los pendientes ---")
        else:
            logging.info(
                "No se encontraron CDAMOSTRAS pendientes según los criterios definidos.")

        manejador_mylims.cerrar()
    else:
        logging.critical(
            "No se pudo establecer la conexión a myLIMS_Novo. El proceso no puede continuar.")

    logging.info("********** FIN DEL PROCESO DE ENVÍO DE INFORMES **********")
