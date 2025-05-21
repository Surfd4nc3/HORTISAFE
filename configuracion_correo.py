#configuracion_correo.py
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USA_TLS = True

SMTP_USER = "resultados@ceimic.com"
SMTP_PASSWORD = "*123myLIMS4"

REMITENTE_NOMBRE = "Ceimic Perú - Resultados de Análisis"
REMITENTE_EMAIL = SMTP_USER


DESTINATARIO_TO_POR_DEFECTO = "Eder.ortega@ceimic.com;joel.zarate@ceimic.com"
DESTINATARIO_CC_POR_DEFECTO = "herrick.davis@ceimic.com"
DESTINATARIO_BCC_POR_DEFECTO = "herrick.davis@ceimic.com"

ASUNTO_PLANTILLA = "Resultados de Análisis - Muestra {cdamostra}"

LOGS_DIR = "logs"
LOG_EXITOS_FILE = "exitos.log"
LOG_ERRORES_FILE = "errores.log"
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'
