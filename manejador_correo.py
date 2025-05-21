#manejador_correo.py
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import logging
import re # Para validación de email con expresiones regulares
from configuracion_correo import (
    SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_USA_TLS,
    REMITENTE_NOMBRE, REMITENTE_EMAIL
)

# Configuración del logger (se asume que se configura en el script principal)
logger = logging.getLogger(__name__)

def is_valid_email(email_str):
    """Valida un formato de dirección de correo electrónico."""
    if not email_str or not isinstance(email_str, str):
        return False
    # Expresión regular para validar emails (simplificada, puedes usar una más compleja si es necesario)
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
    # Asume que las fechas ya vienen formateadas o las formateas aquí si es necesario
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

def enviar_correo_con_adjunto(destinatarios_to, asunto, cuerpo_html, ruta_archivo_adjunto=None, 
                               destinatarios_cc=None, destinatarios_bcc=None, reply_to_email=None):
    """
    Envía un correo electrónico con un archivo adjunto.
    - destinatarios_to, cc, bcc: Cadenas de correos separados por ';' o listas de correos.
    - reply_to_email: Dirección de correo para el campo Reply-To.
    """
    if not SMTP_SERVER or not SMTP_USER or not SMTP_PASSWORD:
        logger.error("Configuración SMTP incompleta (servidor, usuario o contraseña no definidos). No se puede enviar correo.")
        return False

    msg = MIMEMultipart('alternative')
    msg['From'] = f"{REMITENTE_NOMBRE} <{REMITENTE_EMAIL}>"
    
    # Procesar destinatarios TO
    lista_to_validos = []
    if isinstance(destinatarios_to, str):
        for email in destinatarios_to.split(';'):
            email_limpio = email.strip()
            if is_valid_email(email_limpio):
                lista_to_validos.append(email_limpio)
    elif isinstance(destinatarios_to, list):
        for email in destinatarios_to:
            if is_valid_email(email.strip()):
                lista_to_validos.append(email.strip())
    
    if not lista_to_validos:
        logger.error("No hay destinatarios válidos en 'To'. No se enviará el correo.")
        return False
    msg['To'] = ", ".join(lista_to_validos)

    # Procesar destinatarios CC
    lista_cc_validos = []
    if destinatarios_cc:
        if isinstance(destinatarios_cc, str):
            for email in destinatarios_cc.split(';'):
                email_limpio = email.strip()
                if is_valid_email(email_limpio):
                    lista_cc_validos.append(email_limpio)
        elif isinstance(destinatarios_cc, list):
            for email in destinatarios_cc:
                if is_valid_email(email.strip()):
                    lista_cc_validos.append(email.strip())
        if lista_cc_validos:
            msg['Cc'] = ", ".join(lista_cc_validos)

    # Procesar destinatarios BCC (no se añaden al header, se pasan a sendmail)
    lista_bcc_validos = []
    if destinatarios_bcc:
        if isinstance(destinatarios_bcc, str):
            for email in destinatarios_bcc.split(';'):
                email_limpio = email.strip()
                if is_valid_email(email_limpio):
                    lista_bcc_validos.append(email_limpio)
        elif isinstance(destinatarios_bcc, list):
            for email in destinatarios_bcc:
                if is_valid_email(email.strip()):
                    lista_bcc_validos.append(email.strip())

    msg['Subject'] = asunto

    # Añadir Reply-To si se proporciona y es válido
    if reply_to_email and is_valid_email(reply_to_email.strip()):
        msg.add_header('Reply-To', reply_to_email.strip())
        logger.info(f"Estableciendo Reply-To a: {reply_to_email.strip()}")
    elif reply_to_email:
        logger.warning(f"Dirección de Reply-To inválida: {reply_to_email}. No se establecerá.")
    
    msg.attach(MIMEText(cuerpo_html, 'html', 'utf-8')) # Especificar utf-8

    if ruta_archivo_adjunto and os.path.exists(ruta_archivo_adjunto):
        try:
            nombre_archivo = os.path.basename(ruta_archivo_adjunto)
            with open(ruta_archivo_adjunto, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {nombre_archivo}",
            )
            msg.attach(part)
            logger.info(f"Archivo '{nombre_archivo}' adjuntado al correo.")
        except Exception as e:
            logger.error(f"Error al adjuntar el archivo '{ruta_archivo_adjunto}': {e}", exc_info=True)
            return False 
    elif ruta_archivo_adjunto:
        logger.warning(f"El archivo adjunto especificado no existe: {ruta_archivo_adjunto}. El correo se enviará sin adjunto.")

    todos_los_destinatarios_finales = lista_to_validos + lista_cc_validos + lista_bcc_validos
    if not todos_los_destinatarios_finales:
        logger.error("No hay destinatarios válidos (To, Cc, o Bcc) después de la validación. No se enviará el correo.")
        return False

    try:
        server = None
        if SMTP_USA_TLS or SMTP_PORT == 587:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) # Añadido timeout
            server.ehlo()
            server.starttls()
            server.ehlo()
        elif SMTP_PORT == 465:
             server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=30) # Añadido timeout
        else:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) # Añadido timeout
        
        logger.info(f"Intentando login SMTP a {SMTP_SERVER}:{SMTP_PORT} con usuario {SMTP_USER}")
        server.login(SMTP_USER, SMTP_PASSWORD)
        logger.info("Login SMTP exitoso.")
        
        server.sendmail(REMITENTE_EMAIL, todos_los_destinatarios_finales, msg.as_string())
        server.quit()
        logger.info(f"Correo enviado exitosamente a: {', '.join(lista_to_validos)}")
        return True
    except smtplib.SMTPAuthenticationError as e:
        logger.error(f"Error de autenticación SMTP: {e}. Verifica usuario/contraseña y configuración del servidor.", exc_info=True)
        return False
    except smtplib.SMTPConnectError as e:
        logger.error(f"Error de conexión SMTP al servidor {SMTP_SERVER}:{SMTP_PORT}. Error: {e}", exc_info=True)
        return False
    except smtplib.SMTPServerDisconnected as e:
        logger.error(f"Servidor SMTP desconectado inesperadamente. Error: {e}", exc_info=True)
        return False
    except Exception as e:
        logger.error(f"Error general al enviar el correo: {e}", exc_info=True)
        return False