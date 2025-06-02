# generador_excel.py
import pandas as pd
from datetime import datetime
import os  # Necesario para os.path y os.makedirs
import logging  # Para el logging
from unidecode import unidecode
# La configuración del logger se hará en el script principal (sendTrujillo.py)
# Pero podemos obtener el logger aquí para usarlo.
logger = logging.getLogger(__name__)


def formatear_fecha_mejorado(valor_fecha_original, formato_salida_deseado):
    """
    Parsea una cadena de fecha (DD/MM/YYYY o DD/MM/YYYY HH:MM:SS) y la reformatea.
    Devuelve una cadena vacía si la entrada es None o una cadena vacía.
    Devuelve la cadena original si no se puede parsear.
    """
    if valor_fecha_original is None:
        return ''

    valor_fecha_str = str(valor_fecha_original).strip()

    if not valor_fecha_str:
        return ''

    fecha_parte_str = valor_fecha_str
    if ' ' in valor_fecha_str:
        fecha_parte_str = valor_fecha_str.split(' ')[0]

    try:
        dt_obj = datetime.strptime(fecha_parte_str, '%d/%m/%Y')
        resultado_formateado = dt_obj.strftime(formato_salida_deseado)
        return resultado_formateado
    except ValueError as e:
        logger.warning(
            f"No se pudo parsear la fecha '{fecha_parte_str}' (original: '{valor_fecha_original}') como DD/MM/YYYY. Error: {e}. Devolviendo original.")
        return valor_fecha_original
    except Exception as ex:
        logger.error(
            f"Error inesperado en formatear_fecha con valor '{valor_fecha_original}': {ex}. Devolviendo original.", exc_info=True)
        return valor_fecha_original


def MatricesEquivalencias(valorEntrante):
    # aqui incluiriamos equivalencias, array1 seria como viene el valor entrante, valor 2 es como sale, tener un diccionario con las equivalencias
    equivalencias = {
        'Arándanos': 'arandano',
        'Frambuesa': 'frambuesa',
        'Zarzamora': 'zarzamora',
        # Agregar más equivalencias según sea necesario
    }
    return equivalencias.get(valorEntrante, valorEntrante)


def crear_excel_trujillo(datos_encabezado_dict, lista_resultados_dicts, nombre_archivo_excel_completo):
    """
    Crea un archivo Excel para Trujillo.
    Retorna la ruta completa del archivo si se crea exitosamente, None si falla.
    """
    headers_trujillo = [
        'Nombre', 'SS', 'OT', 'Informe', 'N Muestra', 'Matriz', 'Variedad',
        'Codigo Prod', 'Nombre Prod', 'Cuartel/Lote', 'Parcela/Sector',
        'Turno/Equipo', 'Fecha muestra', 'Fecha ingreso', 'Fecha emisión',
        'Analisis/A', 'Resultado'
    ]

    filas_para_df = []
    # Obtener cdamostra para logs, priorizando 'identificacao' del encabezado
    cdamostra_log = "Desconocida"
    if datos_encabezado_dict and datos_encabezado_dict.get('identificacao'):
        cdamostra_log = datos_encabezado_dict.get('identificacao')
    # Fallback al primer resultado si no está en encabezado
    elif lista_resultados_dicts and lista_resultados_dicts[0].get('CDAMOSTRA'):
        cdamostra_log = lista_resultados_dicts[0].get('CDAMOSTRA')

    if not lista_resultados_dicts:
        logger.info(
            f"No hay resultados analíticos para la muestra {cdamostra_log} (archivo '{nombre_archivo_excel_completo}'). No se creará Excel de Trujillo.")
        return None

    for res_dict in lista_resultados_dicts:
        fila_actual = {}

        if datos_encabezado_dict:
            fila_actual['Nombre'] = 'HORTIFRUT - PERU S.A.C.'
            fila_actual['OT'] = datos_encabezado_dict.get('NMPROCESSO', '')
            fila_actual['Informe'] = datos_encabezado_dict.get(
                'numero_base', '')

            # N Muestra: Priorizar 'identificacao' del encabezado general
            fila_actual['N Muestra'] = datos_encabezado_dict.get(
                'identificacao', '')

            fila_actual['Matriz'] = MatricesEquivalencias(
                datos_encabezado_dict.get('desc_amostra', 'N/A'))
            '''quitar_tildes(
                datos_encabezado_dict.get('desc_amostra', ''))'''
            fila_actual['Variedad'] = datos_encabezado_dict.get('variedad', '')
            fila_actual['Codigo Prod'] = datos_encabezado_dict.get(
                'cod_productor', '')
            fila_actual['Nombre Prod'] = datos_encabezado_dict.get(
                'productor', '')
            fila_actual['Cuartel/Lote'] = datos_encabezado_dict.get(
                'huerto', '')
            fila_actual['Parcela/Sector'] = 'N/A'
            fila_actual['Turno/Equipo'] = 'N/A'

            fila_actual['Fecha muestra'] = formatear_fecha_mejorado(
                datos_encabezado_dict.get('datacoleta'), '%Y-%m-%d')
            fila_actual['Fecha ingreso'] = formatear_fecha_mejorado(
                datos_encabezado_dict.get('datachegada'), '%Y/%m/%d')
            fila_actual['Fecha emisión'] = formatear_fecha_mejorado(
                datos_encabezado_dict.get('data_emissao'), '%Y/%m/%d')
        else:
            logger.warning(
                f"No se proporcionó datos_encabezado_dict para la muestra asociada al archivo '{nombre_archivo_excel_completo}'. Columnas de encabezado general podrían quedar vacías o con N/D.")
            for col in ['Nombre', 'OT', 'Informe', 'N Muestra', 'Matriz', 'Variedad', 'Codigo Prod', 'Nombre Prod',
                        'Cuartel/Lote', 'Parcela/Sector', 'Turno/Equipo', 'Fecha muestra',
                        'Fecha ingreso', 'Fecha emisión']:
                # Valor por defecto si no hay encabezado
                fila_actual[col] = 'N/D'

        fila_actual['SS'] = res_dict.get('ref', '')
        fila_actual['Analisis/A'] = res_dict.get('parametro', '')
        fila_actual['Resultado'] = res_dict.get('resultado', '')

        # Fallback para 'N Muestra' si no se obtuvo del encabezado
        if not fila_actual.get('N Muestra'):  # Si está vacío o no se asignó
            # CDAMOSTRA debe estar en res_dict (QUERY_RESULTADOS)
            fila_actual['N Muestra'] = res_dict.get('CDAMOSTRA', '')

        filas_para_df.append(fila_actual)

    if filas_para_df:
        df = pd.DataFrame(filas_para_df, columns=headers_trujillo)
        try:
            directorio_excel = os.path.dirname(nombre_archivo_excel_completo)
            if directorio_excel and not os.path.exists(directorio_excel):
                os.makedirs(directorio_excel, exist_ok=True)
                logger.info(
                    f"Directorio para Excel creado: {directorio_excel}")

            df.to_excel(nombre_archivo_excel_completo,
                        index=False, engine='openpyxl')
            logger.info(
                f" Excel para Trujillo '{nombre_archivo_excel_completo}' creado exitosamente para muestra {cdamostra_log}.")
            # Retorna la ruta del archivo si es exitoso
            return nombre_archivo_excel_completo
        except Exception as e:
            logger.error(
                f"❌ Error al crear Excel para Trujillo '{nombre_archivo_excel_completo}' para muestra {cdamostra_log}: {e}", exc_info=True)
            return None  # Retorna None si hay error
    else:
        logger.warning(
            f"No se generaron filas para el Excel de Trujillo '{nombre_archivo_excel_completo}' (muestra {cdamostra_log}).")
        return None  # Retorna None si no se generaron filas


def crear_excel_olmos(datos_encabezado_dict, lista_resultados_dicts, nombre_archivo_excel_completo):
    """
    Crea un archivo Excel para Olmos.
    Retorna la ruta completa del archivo si se crea exitosamente, None si falla.
    """
    headers_olmos = [
        'Proyecto', 'N Muestra', 'Especie', 'Variedad', 'Productor',
        'Nombre Pro', 'Nombre Pa', 'Parcela/ Se', 'Turno/Equ', 'Nombre Tu',
        'Fecha Ing', 'Fecha M', 'Fecha Em',
        'Analisis/A', 'Resultado', 'N° Analitos D'
    ]
    filas_para_df = []
    num_analitos_d = len(
        lista_resultados_dicts) if lista_resultados_dicts else 0

    cdamostra_log = "Desconocida"
    if datos_encabezado_dict and datos_encabezado_dict.get('identificacao'):
        cdamostra_log = datos_encabezado_dict.get('identificacao')
    elif lista_resultados_dicts and lista_resultados_dicts[0].get('CDAMOSTRA'):
        cdamostra_log = lista_resultados_dicts[0].get('CDAMOSTRA')

    if not lista_resultados_dicts:
        logger.info(
            f"No hay resultados analíticos para la muestra {cdamostra_log} (archivo '{nombre_archivo_excel_completo}'). No se creará Excel de Olmos.")
        return None

    for res_dict in lista_resultados_dicts:
        fila_actual = {}
        if datos_encabezado_dict:
            # Asumiendo que 'idprocesso' es el proyecto para Olmos
            fila_actual['Proyecto'] = datos_encabezado_dict.get(
                'idprocesso', '')
            fila_actual['N Muestra'] = datos_encabezado_dict.get(
                'identificacao', '')
            # enviar datos_encabezado_dict.get('desc_amostra' a la fucnuin MatricesEquivalencias que enviaremos desc_amostra
            fila_actual['Especie'] = MatricesEquivalencias(
                datos_encabezado_dict.get('desc_amostra', 'N/A'))
            '''fila_actual['Especie'] = quitar_tildes(
                datos_encabezado_dict.get('desc_amostra', 'N/A'))'''
            fila_actual['Variedad'] = datos_encabezado_dict.get(
                'variedad', 'N/A')
            fila_actual['Productor'] = datos_encabezado_dict.get(
                'cod_productor', 'N/A')
            fila_actual['Nombre Pro'] = datos_encabezado_dict.get(
                'productor', 'N/A')
            fila_actual['Nombre Pa'] = datos_encabezado_dict.get(
                'huerto', 'N/A')
            fila_actual['Parcela/ Se'] = 'N/A'
            fila_actual['Turno/Equ'] = 'N/A'
            fila_actual['Nombre Tu'] = 'N/A'
            fila_actual['Fecha M'] = formatear_fecha_mejorado(
                datos_encabezado_dict.get('datacoleta'), '%Y-%m-%d')
            fila_actual['Fecha Ing'] = formatear_fecha_mejorado(
                datos_encabezado_dict.get('datachegada'), '%Y/%m/%d')
            fila_actual['Fecha Em'] = formatear_fecha_mejorado(
                datos_encabezado_dict.get('data_emissao'), '%Y/%m/%d')
        else:
            logger.warning(
                f"No se proporcionó datos_encabezado_dict para la muestra asociada al archivo '{nombre_archivo_excel_completo}' (Olmos).")
            for col in headers_olmos:
                if col not in ['Analisis/A', 'Resultado', 'N° Analitos D', 'N Muestra', 'Proyecto']:
                    fila_actual[col] = 'N/D'

        # Fallbacks si no vienen del encabezado
        if not fila_actual.get('Proyecto'):
            # Asumiendo que 'ref' es el proyecto para Olmos
            fila_actual['Proyecto'] = res_dict.get('ref', '')
        if not fila_actual.get('N Muestra'):
            fila_actual['N Muestra'] = res_dict.get('CDAMOSTRA', '')

        fila_actual['Analisis/A'] = res_dict.get('parametro', '')
        fila_actual['Resultado'] = res_dict.get('resultado', '')
        fila_actual['N° Analitos D'] = num_analitos_d

        filas_para_df.append(fila_actual)

    if filas_para_df:
        df = pd.DataFrame(filas_para_df, columns=headers_olmos)
        try:
            directorio_excel = os.path.dirname(nombre_archivo_excel_completo)
            if directorio_excel and not os.path.exists(directorio_excel):
                os.makedirs(directorio_excel, exist_ok=True)
                logger.info(
                    f"Directorio para Excel (Olmos) creado: {directorio_excel}")

            df.to_excel(nombre_archivo_excel_completo,
                        index=False, engine='openpyxl')
            logger.info(
                f"✅ Excel para Olmos '{nombre_archivo_excel_completo}' creado exitosamente para muestra {cdamostra_log}.")
            return nombre_archivo_excel_completo
        except Exception as e:
            logger.error(
                f"❌ Error al crear Excel para Olmos '{nombre_archivo_excel_completo}' para muestra {cdamostra_log}: {e}", exc_info=True)
            return None
    else:
        logger.warning(
            f"No se generaron filas para el Excel de Olmos '{nombre_archivo_excel_completo}' (muestra {cdamostra_log}).")
        return None


def quitar_tildes(texto):
    return unidecode(texto)
