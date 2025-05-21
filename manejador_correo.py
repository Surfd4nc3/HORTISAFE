# manejador_correo.py
import yagmail
import smtplib # <--- AÑADE ESTA LÍNEA
import os
import logging
import re
from configuracion_correo import (
    SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD,
    REMITENTE_NOMBRE, REMITENTE_EMAIL
)
# CORRECCIÓN 1: Importar excepciones específicas de yagmail correctamente
import yagmail
logger = logging.getLogger(__name__)

def is_valid_email(email_str):
    """Valida un formato de dirección de correo electrónico."""
    if not email_str or not isinstance(email_str, str):
        return False
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if re.match(regex, email_str):
        return True
    logger.warning(f"Formato de email inválido detectado: {email_str}")
    return False

def crear_cuerpo_html_correo(cdamostra, datos_encabezado, datos_resultados):
    """Crea un cuerpo de correo HTML formateado."""
    cliente = datos_encabezado.get('solicitante', 'N/D') if datos_encabezado else 'N/D'
    matriz = datos_encabezado.get('matriz', 'N/D') if datos_encabezado else 'N/D'
    n_muestra = datos_encabezado.get('identificacao', cdamostra) if datos_encabezado else cdamostra
    fecha_muestreo = datos_encabezado.get('datacoleta_formateada', 'N/D') if datos_encabezado else 'N/D'
    fecha_recepcion = datos_encabezado.get('datachegada_formateada', 'N/D') if datos_encabezado else 'N/D'
    num_analitos = len(datos_resultados) if datos_resultados else 0

    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
            .container {{ margin: 20px; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }}
            .header {{ background-color: #f4f4f4; padding: 10px; text-align: center; }}
            .content p {{ margin-bottom: 10px; }}
            .footer {{ margin-top: 20px; font-size: 0.9em; color: #777; }}
            table {{ width: 100%; border-collapse: collapse; margin-top:15px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>Resultados de Análisis - Muestra {n_muestra}</h2>
            </div>
            <div class="content">
                <p>Estimados,</p>
                <p>Adjunto encontrarán el informe de resultados correspondiente a la muestra <strong>{n_muestra}</strong>.</p>
                <h3>Detalles de la Muestra:</h3>
                <ul>
                    <li><strong>Cliente:</strong> {cliente}</li>
                    <li><strong>N° Muestra:</strong> {n_muestra}</li>
                    <li><strong>Matriz:</strong> {matriz}</li>
                    <li><strong>Fecha de Muestreo:</strong> {fecha_muestreo}</li>
                    <li><strong>Fecha de Recepción en Laboratorio:</strong> {fecha_recepcion}</li>
                </ul>
                <h3>Resumen de Análisis:</h3>
                <p>Se han analizado <strong>{num_analitos}</strong> parámetros/analitos.</p>
                <p>Para visualizar el detalle completo, por favor revise el archivo Excel adjunto.</p>
            </div>
            <div class="footer">
                <p>Atentamente,</p>
                <p>{REMITENTE_NOMBRE}</p>
                <p>Laboratorio Ceimic Perú</p>
            </div>
        </div>
    </body>
    </html>
    """
    return html

def enviar_correo_con_adjunto_yagmail(destinatarios_to, asunto, cuerpo_html, ruta_archivo_adjunto=None,
                               destinatarios_cc=None, destinatarios_bcc=None, reply_to_email=None):
    if not SMTP_SERVER or not SMTP_USER or not SMTP_PASSWORD:
        logger.error("Configuración SMTP incompleta. No se puede enviar correo.")
        return False

    def procesar_lista_correos(dest_str_o_lista):
        validos = []
        if isinstance(dest_str_o_lista, str):
            for email in dest_str_o_lista.split(';'):
                email_limpio = email.strip()
                if is_valid_email(email_limpio):
                    validos.append(email_limpio)
        elif isinstance(dest_str_o_lista, list):
            for email in dest_str_o_lista:
                if is_valid_email(email.strip()):
                    validos.append(email.strip())
        return validos

    lista_to_validos = procesar_lista_correos(destinatarios_to)
    if not lista_to_validos:
        logger.error("No hay destinatarios válidos en 'To'. No se enviará el correo.")
        return False

    lista_cc_validos = procesar_lista_correos(destinatarios_cc) if destinatarios_cc else []
    lista_bcc_validos = procesar_lista_correos(destinatarios_bcc) if destinatarios_bcc else []
    
    adjuntos_finales = []
    if ruta_archivo_adjunto and os.path.exists(ruta_archivo_adjunto):
        adjuntos_finales.append(ruta_archivo_adjunto)
        logger.info(f"Archivo '{os.path.basename(ruta_archivo_adjunto)}' preparado para adjuntar.")
    elif ruta_archivo_adjunto:
        logger.warning(f"El archivo adjunto especificado no existe: {ruta_archivo_adjunto}.")

# (Esto va dentro de la función enviar_correo_con_adjunto_yagmail)
    try:
        logger.info(f"Inicializando Yagmail para {SMTP_USER} en {SMTP_SERVER}:{SMTP_PORT}")
        yag = yagmail.SMTP(
            user={SMTP_USER: REMITENTE_NOMBRE},
            password=SMTP_PASSWORD,
            host=SMTP_SERVER,
            port=SMTP_PORT,
            smtp_starttls=True,
            smtp_ssl=False,
            timeout=30
        )

        headers = {}
        if reply_to_email and is_valid_email(reply_to_email.strip()):
            headers['Reply-To'] = reply_to_email.strip()
            logger.info(f"Estableciendo Reply-To a: {reply_to_email.strip()}")
        elif reply_to_email:
            logger.warning(f"Dirección de Reply-To inválida: {reply_to_email}.")

        logger.info(f"Intentando enviar correo a: {lista_to_validos}, Cc: {lista_cc_validos}, Bcc: {lista_bcc_validos}")

        yag.send(
            to=lista_to_validos,
            subject=asunto,
            contents=[cuerpo_html] + adjuntos_finales,
            cc=lista_cc_validos if lista_cc_validos else None,
            bcc=lista_bcc_validos if lista_bcc_validos else None,
            headers=headers if headers else None
        )

        logger.info(f"Correo enviado exitosamente usando Yagmail a: {', '.join(lista_to_validos)}")
        return True

    except smtplib.SMTPAuthenticationError as e_auth: # CAPTURAMOS EL ERROR DE AUTENTICACIÓN DE SMTPLIB
        logger.error(f"Error de AUTENTICACIÓN SMTP (desde yagmail/smtplib): {e_auth}")
        logger.error("CAUSA MÁS PROBABLE: Credenciales incorrectas o la cuenta requiere una 'Contraseña de Aplicación' debido a MFA en Office 365.")
        return False
    except Exception as e:  # Captura cualquier otra excepción
        logger.error(f"Error general al enviar el correo con Yagmail: {e}", exc_info=True)
        if "WRONG_VERSION_NUMBER" in str(e):
            logger.error("El error 'WRONG_VERSION_NUMBER' persistió. Problema con SSL/TLS.")
        # Aquí podrías añadir más lógica para interpretar 'e' si fuera necesario
        return False