# sendTrujillo.py
import pandas as pd
import os
import logging
from conexion import ManejadorConexionSQL
# Asumo que este archivo existe
from consultas import QUERY_RESULTADOS, QUERY_ENCABEZADOS, QUERY_ENVIADOS_BDCLINK, QUERY_INSERT_ENVIADOS_BDCLINK
from Pendientes import Pendientes
from generador_excel import crear_excel_trujillo, crear_excel_olmos
# from manejador_correo import enviar_correo_con_adjunto, crear_cuerpo_html_correo
from manejador_correo import enviar_correo_con_adjunto, crear_cuerpo_html_correo


from configuracion_correo import (
    LOGS_DIR, LOG_EXITOS_FILE, LOG_ERRORES_FILE, LOG_FORMAT,
    DESTINATARIO_TO_POR_DEFECTO, DESTINATARIO_CC_POR_DEFECTO, DESTINATARIO_BCC_POR_DEFECTO,
    ASUNTO_PLANTILLA,
    DESTINATARIO_TO_TRUJILLO,
    DESTINATARIO_CC_TRUJILLO,
    DESTINATARIO_BCC_TRUJILLO,

    DESTINATARIO_TO_OLMOS,
    DESTINATARIO_CC_OLMOS,
    DESTINATARIO_BCC_OLMOS
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
    logging.getLogger().addHandler(exitos_handler)

    # Handler para archivo de errores
    errores_handler = logging.FileHandler(
        os.path.join(LOGS_DIR, LOG_ERRORES_FILE), mode='a')
    errores_handler.setLevel(logging.WARNING)  # Logs WARNING, ERROR, CRITICAL
    errores_formatter = logging.Formatter(LOG_FORMAT)
    errores_handler.setFormatter(errores_formatter)
    logging.getLogger().addHandler(errores_handler)

    logging.info("Sistema de Logging configurado.")


def procesar_un_pendiente(cdamostra_actual):
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

    # REABRIR LA CONEXIÓN A MYLIMS_NOVO AQUÍ
    manejador_mylims_local = ManejadorConexionSQL("myLIMS_Novo_conn")
    conexion_activa_local = manejador_mylims_local.conectar()
    
    if not conexion_activa_local:
        logging.error(f"No se pudo establecer conexión a myLIMS_Novo para procesar {cdamostra_actual}. Saltando este pendiente.")
        return None # Retorna None si la conexión falla aquí

    try:
        # Obtener resultados
        resultados_data = manejador_mylims_local.ejecutar_consulta(
            QUERY_RESULTADOS, (cdamostra_actual,))
        if resultados_data is not None:
            logging.info(
                f"Resultados obtenidos para {cdamostra_actual}: {len(resultados_data)} filas.")
            datos_procesados["resultados"] = resultados_data
        else:
            logging.error(f"Error al obtener resultados para {cdamostra_actual}.")

        # Obtener encabezado
        encabezado_data_lista = manejador_mylims_local.ejecutar_consulta(
            QUERY_ENCABEZADOS, (cdamostra_actual,))

        if encabezado_data_lista and len(encabezado_data_lista) == 1:
            logging.info(f"Encabezado obtenido para {cdamostra_actual}.")
            datos_procesados["encabezado"] = encabezado_data_lista[0]
        elif encabezado_data_lista:
            logging.warning(
                f"Se esperaba 1 fila para el encabezado de {cdamostra_actual}, pero se obtuvieron {len(encabezado_data_lista)}. Usando la primera si existe.")
            datos_procesados["encabezado"] = encabezado_data_lista[0]
        else:
            logging.error(f"Error al obtener encabezado para {cdamostra_actual}.")

        # Formatear fechas en el encabezado para el cuerpo del correo
        if datos_procesados["encabezado"]:
            datos_procesados["encabezado"]["datacoleta_formateada"] = formatear_fecha_mejorado(
                datos_procesados["encabezado"].get('datacoleta'), '%d/%m/%Y')
            datos_procesados["encabezado"]["datachegada_formateada"] = formatear_fecha_mejorado(
                datos_procesados["encabezado"].get('datachegada'), '%d/%m/%Y')
            
        return datos_procesados
    finally:
        # ASEGURAR CIERRE DE LA CONEXIÓN LOCAL
        if conexion_activa_local:
            manejador_mylims_local.cerrar()


if __name__ == "__main__":
    configurar_logging()  # Configurar logging al inicio
    logging.info(
        "********** INICIO DEL PROCESO DE ENVÍO DE INFORMES **********")

    # Flag para decidir qué Excel generar (1 para Trujillo, otro valor para Olmos)
    tipo_informe_flag = 2  # 1: Trujillo, Otro: Olmos (ej. 2)

    # CONEXION PARA BDCLINK CREAR UNA CONEXION NUEVA
    manejador_bdclink = ManejadorConexionSQL("BDClink_conn")
    conexionactiva_BDCLINK = manejador_bdclink.conectar()
    resultadosEnviados_raw = None

    if conexionactiva_BDCLINK:
        resultadosEnviados_raw = manejador_bdclink.ejecutar_consulta(
            QUERY_ENVIADOS_BDCLINK)
        manejador_bdclink.cerrar()

    # Extrae el valor de 'CDAMOSTRA' de cada diccionario y crea un conjunto con ellos.
    if resultadosEnviados_raw is not None:
        resultadosEnviados_set = {fila['CDAMOSTRA']
                                  for fila in resultadosEnviados_raw}
    else:
        resultadosEnviados_set = set()

    manejador_mylims = ManejadorConexionSQL("myLIMS_Novo_conn")
    conexion_activa = manejador_mylims.conectar()

    if conexion_activa:
        gestor_pendientes = Pendientes(
            manejador_mylims)
        cdamostras_pendientes_set = gestor_pendientes.obtener_pendientes()
        manejador_mylims.cerrar()

        cdamostras_pendientes_final = cdamostras_pendientes_set - resultadosEnviados_set

        cdamostras_pendientes = list(cdamostras_pendientes_final)

        if cdamostras_pendientes:
            logging.info(
                f"Se encontraron {len(cdamostras_pendientes)} CDAMOSTRAS pendientes para procesar.")

            for cdamostra_id_pendiente in cdamostras_pendientes:
                datos_un_pendiente = procesar_un_pendiente(
                    cdamostra_id_pendiente) # Ya no se pasa manejador_db

                if not datos_un_pendiente or (not datos_un_pendiente.get("encabezado") and not datos_un_pendiente.get("resultados")):
                    logging.error(
                        f"No se pudieron obtener datos suficientes para CDAMOSTRA {cdamostra_id_pendiente}. Se omite este pendiente.")
                    continue

                ruta_excel_adjuntar = None
                nombre_excel_base = ""
                cadena_Unidad = ""
                #tipo_informe_flag se le asigna si es olmos o trujillo
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
                    cadena_Unidad = "Trujillo"
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
                    cadena_Unidad = "Olmos"

                if ruta_excel_adjuntar:
                    logging.info(
                        f"Excel '{ruta_excel_adjuntar}' generado exitosamente para CDAMOSTRA {datos_un_pendiente['cdamostra']}.")

                    # Preparar y enviar correo
                    asunto_correo = ASUNTO_PLANTILLA.format(
                        cdamostra=datos_un_pendiente['encabezado']['numero_base'])+' - '+cadena_Unidad

                    cuerpo_html_email = crear_cuerpo_html_correo(
                        datos_un_pendiente['cdamostra'],
                        # Pasar dict vacío si es None
                        datos_un_pendiente.get("encabezado", {}),
                        datos_un_pendiente.get("resultados", [])
                    )
                    if tipo_informe_flag == 1:
                        DESTINATARIO_TO_POR_DEFECTO = DESTINATARIO_TO_TRUJILLO
                        DESTINATARIO_CC_POR_DEFECTO = DESTINATARIO_CC_TRUJILLO
                        DESTINATARIO_BCC_POR_DEFECTO = DESTINATARIO_BCC_TRUJILLO
                    else:
                        DESTINATARIO_TO_POR_DEFECTO = DESTINATARIO_TO_OLMOS
                        DESTINATARIO_CC_POR_DEFECTO = DESTINATARIO_CC_OLMOS
                        DESTINATARIO_BCC_POR_DEFECTO = DESTINATARIO_BCC_OLMOS

                    envio_exitoso = enviar_correo_con_adjunto(
                        destinatarios_to=DESTINATARIO_TO_POR_DEFECTO,
                        asunto=asunto_correo,
                        cuerpo_html=cuerpo_html_email,
                        ruta_archivo_adjunto=ruta_excel_adjuntar,
                        destinatarios_cc=DESTINATARIO_CC_POR_DEFECTO,
                        destinatarios_bcc=DESTINATARIO_BCC_POR_DEFECTO
                    )

                    manejador_bdclink = ManejadorConexionSQL("BDClink_conn")
                    conexionactiva_BDCLINK = manejador_bdclink.conectar()
                    if conexionactiva_BDCLINK:
                        try:
                            manejador_bdclink.ejecutar_consulta(
                                QUERY_INSERT_ENVIADOS_BDCLINK,
                                (datos_un_pendiente['cdamostra'],
                                 cadena_Unidad)
                            )

                            logging.info(
                                f"✅ Registro de envío exitoso para CDAMOSTRA '{datos_un_pendiente['cdamostra']}' en '{cadena_Unidad}' en BDClink.")
                        except Exception as e:
                            logging.error(
                                f"❌ Error al insertar registro de envío para CDAMOSTRA '{datos_un_pendiente['cdamostra']}' en BDClink: {e}", exc_info=True)
                        finally:
                            manejador_bdclink.cerrar()
                    else:
                        logging.error(
                            f"❌ No se pudo establecer conexión a BDClink para registrar el envío de CDAMOSTRA '{datos_un_pendiente['cdamostra']}'.")
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

        # La conexión manejador_mylims principal no necesita cerrarse aquí,
        # ya que cada procesamiento individual de pendiente abre y cierra su propia conexión.
        # manejador_mylims.cerrar() # Esta línea se puede comentar o eliminar
    else:
        logging.critical(
            "No se pudo establecer la conexión a myLIMS_Novo. El proceso no puede continuar.")

    logging.info("********** FIN DEL PROCESO DE ENVÍO DE INFORMES **********")