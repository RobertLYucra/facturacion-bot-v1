import os
import shutil
from datetime import datetime

def sincronizar_excel_onedrive(ruta_origen, ruta_destino, nombre_archivo):
    """
    Copia un archivo desde una carpeta origen a una carpeta destino,
    reemplazando el archivo si ya existe en el destino.
    
    Args:
        ruta_origen (str): Ruta de la carpeta donde se encuentra el archivo original
        ruta_destino (str): Ruta donde se copiará el archivo
        nombre_archivo (str): Nombre del archivo a sincronizar
    
    Returns:
        bool: True si la copia fue exitosa, False en caso contrario
    """
    try:
        # Ruta completa del archivo origen
        archivo_origen = os.path.join(ruta_origen, nombre_archivo)
        
        # Ruta completa del archivo destino (manteniendo el mismo nombre)
        archivo_destino = os.path.join(ruta_destino, nombre_archivo)
        
        # Verificar si el archivo existe en la ubicación de origen
        if not os.path.exists(archivo_origen):
            print(f"Error: El archivo '{nombre_archivo}' no existe en la ruta de origen.")
            return False
        
        # Crear la carpeta de destino si no existe
        if not os.path.exists(ruta_destino):
            os.makedirs(ruta_destino)
            print(f"Carpeta de destino creada: {ruta_destino}")
        
        # Copiar el archivo desde el origen (shutil.copy2 reemplazará automáticamente 
        # si el archivo ya existe en el destino)
        shutil.copy2(archivo_origen, archivo_destino)
        print(f"Archivo copiado y reemplazado en destino: {nombre_archivo}")
        
        # Verificar que se haya copiado correctamente
        if os.path.exists(archivo_destino):
            print(f"Verificación exitosa: Archivo encontrado en destino")
            return True
        else:
            print(f"Error: No se encontró el archivo en destino después de copiarlo")
            return False
        
    except Exception as e:
        print(f"Error durante la copia del archivo: {str(e)}")
        return False

# Ejecución única
if __name__ == "__main__":
    # Configura estas rutas según tu sistema
    RUTA_ORIGEN = r"C:\Users\Administrator\OneDrive - Indra (1)\Facturas\Carpeta Archivos Adjuntos\Registro Historico"
    RUTA_DESTINO = r"D:\facturas_bot\Registros Historico"
    NOMBRE_ARCHIVO = "3.Historicov2.xlsx"
    
    print(f"Copiando archivo a las {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    sincronizar_excel_onedrive(RUTA_ORIGEN, RUTA_DESTINO, NOMBRE_ARCHIVO)
    print("Copia completada.")