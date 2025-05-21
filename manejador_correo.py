# manejador_correo.py
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import logging
import re
from configuracion_correo import (
    SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_USA_TLS,
    REMITENTE_NOMBRE, REMITENTE_EMAIL
)
# Pynliner ya no es necesario si volvemos a smtplib y el cliente de correo ignora los <style>
# from pynliner import Pynliner 

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
    """Crea un cuerpo de correo HTML formateado y con mejor estética."""
    cliente = datos_encabezado.get('solicitante', 'N/D') if datos_encabezado else 'N/D'
    matriz = datos_encabezado.get('matriz', 'N/D') if datos_encabezado else 'N/D'
    #n_muestra = datos_encabezado.get('identificacao', cdamostra) if datos_encabezado else cdamostra
    n_muestra = datos_encabezado.get('numero_base', 'N/A') 
    fecha_muestreo = datos_encabezado.get('datacoleta_formateada', 'N/D') if datos_encabezado else 'N/D'
    fecha_recepcion = datos_encabezado.get('datachegada_formateada', 'N/D') if datos_encabezado else 'N/D'
    num_analitos = len(datos_resultados) if datos_resultados else 0

    # Estilos CSS (se intentarán aplicar, pero la compatibilidad depende del cliente de correo)
    # Para máxima compatibilidad, los estilos críticos deberían estar "en línea" en las etiquetas HTML.
    # Sin embargo, mantenerlos aquí es más fácil de gestionar.
    html_estilos = """
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif, 'Apple Color Emoji', 'Segoe UI Emoji', 'Segoe UI Symbol';
            line-height: 1.6;
            color: #333333;
            background-color: #f4f5f7;
            margin: 0;
            padding: 20px;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }
        .email-container {
            max-width: 680px;
            margin: 0 auto;
            background-color: #ffffff;
            padding: 35px;
            border-radius: 12px;
            box-shadow: 0 6px 25px rgba(0,0,0,0.08);
            border: 1px solid #dee2e6;
        }
        .email-header {
            background-color: #004a99; /* Azul corporativo */
            color: #ffffff;
            padding: 25px;
            border-radius: 8px 8px 0 0;
            text-align: center;
        }
        .email-header h2 {
            margin: 0;
            font-size: 26px;
            font-weight: 600;
        }
        .email-content {
            padding: 20px 0;
        }
        .email-content h3 {
            font-size: 18px;
            color: #004a99;
            margin-top: 30px;
            margin-bottom: 15px;
            border-bottom: 2px solid #e9ecef;
            padding-bottom: 8px;
            font-weight: 600;
        }
        .email-content p {
            margin-bottom: 15px;
            font-size: 15px;
            color: #555555;
        }
        .details-list {
            list-style-type: none;
            padding-left: 0;
            font-size: 15px;
            margin-bottom: 20px;
        }
        .details-list li {
            margin-bottom: 10px;
            padding: 8px 12px;
            background-color: #f8f9fa;
            border-left: 3px solid #004a99;
            border-radius: 4px;
        }
        .details-list li strong {
            color: #333333;
            font-weight: 600;
            margin-right: 8px;
        }
        .email-footer {
            margin-top: 30px;
            padding-top: 20px;
            font-size: 13px;
            color: #777777;
            text-align: center;
            border-top: 1px solid #e9ecef;
        }
        .email-footer p {
            margin: 5px 0;
        }
    </style>
    """

    # Cuerpo HTML del correo
    html_cuerpo = f"""
    <div class="email-container">
        <div class="email-header">
            <h2>Resultados de Análisis - Informe {n_muestra}</h2>
        </div>
        <div class="email-content">
            <p>Estimados,</p>
            <p>Adjunto encontrarán el informe de resultados correspondiente al informe <strong>{n_muestra}</strong>.</p>
            
            <h3>Detalles de la Muestra:</h3>
            <ul class="details-list">
                <li><strong>Cliente:</strong> {cliente}</li>
                <li><strong>N° Informe:</strong> {n_muestra}</li>
                <li><strong>Matriz:</strong> {matriz}</li>
                <li><strong>Fecha de Muestreo:</strong> {fecha_muestreo}</li>
                <li><strong>Fecha de Recepción en Laboratorio:</strong> {fecha_recepcion}</li>
            </ul>
            
            <h3>Resumen de Análisis:</h3>
            <p>Se han analizado <strong>{num_analitos}</strong> parámetros/analitos.</p>
            <p>Para visualizar el detalle completo, por favor revise el archivo Excel adjunto.</p>
        </div>
        <div class="email-footer">
            <p>Atentamente,</p>
            <p>{REMITENTE_NOMBRE}</p>
            <p>Laboratorio Ceimic Perú</p>
        </div>
    </div>
    """

    # HTML completo con la declaración DOCTYPE y la etiqueta <style> en <head>
    # Aunque muchos clientes de correo ignoran los estilos en <head>,
    # algunos más modernos pueden usarlos. El CSS en línea sigue siendo lo más seguro.
    # Por ahora, mantendremos los estilos en <head> para ver si smtplib lo maneja diferente.
    html_final = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Resultados de Análisis - Muestra {n_muestra}</title>
        {html_estilos}
    </head>
    <body>
        {html_cuerpo}
    </body>
    </html>
    """
    return html_final


# Esta es tu función original para enviar con smtplib, la reactivamos.
def enviar_correo_con_adjunto(destinatarios_to, asunto, cuerpo_html, ruta_archivo_adjunto=None, 
                              destinatarios_cc=None, destinatarios_bcc=None, reply_to_email=None):
    if not SMTP_SERVER or not SMTP_USER or not SMTP_PASSWORD:
        logger.error("Configuración SMTP incompleta (servidor, usuario o contraseña no definidos). No se puede enviar correo.")
        return False

    msg = MIMEMultipart('alternative') # 'alternative' permite incluir texto plano y HTML
    msg['From'] = f"{REMITENTE_NOMBRE} <{REMITENTE_EMAIL}>"
    
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

    if reply_to_email and is_valid_email(reply_to_email.strip()):
        msg.add_header('Reply-To', reply_to_email.strip())
        logger.info(f"Estableciendo Reply-To a: {reply_to_email.strip()}")
    elif reply_to_email:
        logger.warning(f"Dirección de Reply-To inválida: {reply_to_email}. No se establecerá.")
    
    # Adjuntar la parte HTML. Es importante establecer el subtipo a 'html'.
    msg.attach(MIMEText(cuerpo_html, 'html', 'utf-8'))

    if ruta_archivo_adjunto and os.path.exists(ruta_archivo_adjunto):
        try:
            nombre_archivo = os.path.basename(ruta_archivo_adjunto)
            with open(ruta_archivo_adjunto, "rb") as attachment_file:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment_file.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {nombre_archivo}",
            )
            msg.attach(part)
            logger.info(f"Archivo '{nombre_archivo}' adjuntado al correo.")
        except Exception as e:
            logger.error(f"Error al adjuntar el archivo '{ruta_archivo_adjunto}': {e}", exc_info=True)
            # Considerar si se debe retornar False aquí o enviar el correo sin adjunto.
            # Por ahora, el correo se intentará enviar incluso si falla el adjunto.
    elif ruta_archivo_adjunto:
        logger.warning(f"El archivo adjunto especificado no existe: {ruta_archivo_adjunto}. El correo se enviará sin adjunto.")

    todos_los_destinatarios_finales = lista_to_validos + lista_cc_validos + lista_bcc_validos
    if not todos_los_destinatarios_finales: # Doble chequeo, aunque el primero en lista_to_validos debería ser suficiente
        logger.error("No hay destinatarios válidos (To, Cc, o Bcc) después de la validación. No se enviará el correo.")
        return False

    try:
        server = None
        logger.info(f"Intentando conexión SMTP a {SMTP_SERVER}:{SMTP_PORT}")
        if SMTP_USA_TLS or SMTP_PORT == 587: # Para STARTTLS
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
            server.ehlo() # Saludo al servidor
            server.starttls() # Iniciar conexión TLS
            server.ehlo() # Saludo de nuevo después de TLS
        elif SMTP_PORT == 465: # Para SSL directo
             server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=30)
             server.ehlo() # Saludo al servidor (opcional para SMTP_SSL pero no daña)
        else: # Conexión sin encriptar (no recomendado)
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
            server.ehlo()
        
        logger.info(f"Intentando login SMTP con usuario {SMTP_USER}")
        server.login(SMTP_USER, SMTP_PASSWORD)
        logger.info("Login SMTP exitoso.")
        
        server.sendmail(REMITENTE_EMAIL, todos_los_destinatarios_finales, msg.as_string())
        server.quit()
        logger.info(f"Correo enviado exitosamente con smtplib a: {', '.join(lista_to_validos)}")
        return True
        
    except smtplib.SMTPAuthenticationError as e_auth:
        logger.error(f"Error de AUTENTICACIÓN SMTP con smtplib: {e_auth}", exc_info=True)
        logger.error("CAUSA MÁS PROBABLE: Credenciales incorrectas o la cuenta requiere una 'Contraseña de Aplicación' debido a MFA en Office 365.")
        return False
    except smtplib.SMTPConnectError as e_conn:
        logger.error(f"Error de CONEXIÓN SMTP al servidor {SMTP_SERVER}:{SMTP_PORT}. Error: {e_conn}", exc_info=True)
        return False
    except smtplib.SMTPServerDisconnected as e_disconn:
        logger.error(f"Servidor SMTP desconectado inesperadamente. Error: {e_disconn}", exc_info=True)
        return False
    except Exception as e:
        logger.error(f"Error general al enviar el correo con smtplib: {e}", exc_info=True)
        return False

# La función que usaba yagmail queda aquí por si quieres volver a probarla,
# pero no la usaremos si reactivamos la de smtplib.
# def enviar_correo_con_adjunto_yagmail(...):
#     pass