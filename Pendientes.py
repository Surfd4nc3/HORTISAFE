# Pendientes.py
import pandas as pd # pandas no se usa directamente en esta clase para obtener pendientes
# from conexion import ManejadorConexionSQL # Esta importación es correcta si conexion.py está en el mismo directorio

class Pendientes:
    def __init__(self, manejador_db):
        self.manejador_db = manejador_db
        # La conexión se establece al llamar a manejador_db.conectar()
        # El cursor se obtiene dentro de ejecutar_consulta

    def obtener_pendientes(self):
        """Obtiene los CDAMOSTRAS pendientes de la base de datos."""
        query = """
        SELECT DISTINCT 
            a.cdamostra,
            a.cdunidadeneg,
            0 as hidro,
            aga.cdgrpamostra as cdgrpamostra,
            a.cdunidade
        FROM amostra a
        INNER JOIN amostrasgrpamostra aga ON aga.cdamostra = a.cdamostra
        INNER JOIN histsitamostra h ON h.cdamostra = a.cdamostra AND h.cdsitamostra = 4
        INNER JOIN tipoamostra t ON t.cdtipoamostra = a.cdtipoamostra
        INNER JOIN AMOSTRASITENSPRO AIP ON AIP.CDAMOSTRA = a.CDAMOSTRA
        INNER JOIN PROCESSO P ON P.CDPROCESSO = AIP.CDPROCESSO
        WHERE 
            a.CDCLASSEAMOSTRA in (1)  
            AND a.flativo = 'S'
            AND P.CDPROCESSO = 99262
            AND a.dtcoleta is not null
        """
        # La conexión se maneja dentro de ejecutar_consulta
        # ejecutar_consulta ahora devuelve una lista de diccionarios o None
        resultados_lista_de_dicts = self.manejador_db.ejecutar_consulta(query)
        
        if resultados_lista_de_dicts: # Si la lista no es None y no está vacía
            # --- CAMBIO IMPORTANTE AQUÍ ---
            # 'fila' es un diccionario, accedemos por la clave 'cdamostra'
            lista_cdamostras = []
            for fila_dict in resultados_lista_de_dicts:
                if isinstance(fila_dict, dict) and 'cdamostra' in fila_dict:
                    lista_cdamostras.append(fila_dict['cdamostra'])
                else:
                    # Esto podría pasar si una fila no es un diccionario o no tiene la clave
                    print(f"⚠️ Fila inesperada o sin 'cdamostra': {fila_dict}")
            
            if not lista_cdamostras and resultados_lista_de_dicts:
                 print("⚠️ La consulta de pendientes devolvió datos, pero no se pudo extraer 'cdamostra' de ninguna fila.")
            return lista_cdamostras
        
        elif resultados_lista_de_dicts == []: # Lista vacía, significa que no hay pendientes
            print("ℹ️ No se encontraron CDAMOSTRAS pendientes que cumplan los criterios.")
            return []
        else: # resultados_lista_de_dicts es None, lo que indica un error en la consulta
            print("❌ Error al obtener CDAMOSTRAS pendientes (la consulta devolvió None).")
            return []
