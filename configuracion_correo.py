#configuracion_correo.py
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USA_TLS = True

SMTP_USER = "noreply.new@ceimic.com"
SMTP_PASSWORD = "P)645998820497uc"

REMITENTE_NOMBRE = "Ceimic Perú - Resultados de Análisis"
REMITENTE_EMAIL = SMTP_USER


DESTINATARIO_TO_POR_DEFECTO = "Eder.ortega@ceimic.com;joel.zarate@ceimic.com"
DESTINATARIO_CC_POR_DEFECTO = "herrick.davis@ceimic.com"
DESTINATARIO_BCC_POR_DEFECTO = "herrick.davis@ceimic.com;rita.noriega@ceimic.com;Eder.ortega@ceimic.com;joel.zarate@ceimic.com"

DESTINATARIO_TO_TRUJILLO="rcevallos@hortifrut.com;rimunoz@hortifrut.com"
DESTINATARIO_CC_TRUJILLO="herrick.davis@ceimic.com"
DESTINATARIO_BCC_TRUJILLO="herrick.davis@ceimic.com;rita.noriega@ceimic.com;Eder.ortega@ceimic.com;joel.zarate@ceimic.com"

DESTINATARIO_TO_OLMOS="nayda.berruc@hfeberries.com;rimunoz@hortifrut.com"
DESTINATARIO_CC_OLMOS="herrick.davis@ceimic.com"
DESTINATARIO_BCC_OLMOS="herrick.davis@ceimic.com;rita.noriega@ceimic.com;Eder.ortega@ceimic.com;joel.zarate@ceimic.com"

ASUNTO_PLANTILLA = "Resultado de ensayo {cdamostra}"

LOGS_DIR = "logs"
LOG_EXITOS_FILE = "exitos.log"
LOG_ERRORES_FILE = "errores.log"
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'
