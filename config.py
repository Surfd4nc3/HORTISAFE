# config_bd.py
DB_CONFIG = {
    "myLIMS_Novo_conn": {
        "server": "CMCCLDDB01",
        "database": "myLIMS_Novo",
        "username": "SI$MYLIMS",
        "password": "@M4qBjLs",
        # Importante: Asegúrate de tener el driver correcto aquí o que {SQL Server} funcione.
        # Si {SQL Server} no funciona, prueba con {ODBC Driver 17 for SQL Server} u otro específico.
        "driver": "{SQL Server}" # O "{ODBC Driver 17 for SQL Server}"
    },
    "BDClink_conn": {
        "server": "CMCCLDDB01", # O "CMCPEW1031\\SQLEXPRESS" si es diferente
        "database": "BDClink",
        "username": "SI$MYLIMS", # O usa "integrated_security": True si aplica
        "password": "@M4qBjLs", # O elimina si "integrated_security": True
        "driver": "{SQL Server}" # O "{ODBC Driver 17 for SQL Server}"
        # "integrated_security": True # Descomenta esta línea y comenta username/password si usas autenticación de Windows
    }
    # Puedes añadir más configuraciones de conexión aquí
}