import os
import shutil
from datetime import datetime
import pandas as pd

def sincronizar_excel_onedrive(ruta_origen, ruta_destino, nombre_archivo):
    """
    Sincroniza un archivo Excel desde una carpeta de OneDrive local a otra ubicación.
    
    Args:
        ruta_origen (str): Ruta de la carpeta OneDrive donde se encuentra el archivo
        ruta_destino (str): Ruta donde se copiará el archivo
        nombre_archivo (str): Nombre del archivo Excel a sincronizar
    
    Returns:
        bool: True si la sincronización fue exitosa, False en caso contrario
    """
    try:
        # Ruta completa del archivo origen
        archivo_origen = os.path.join(ruta_origen, nombre_archivo)
        
        # Ruta completa del archivo destino
        archivo_destino = os.path.join(ruta_destino, nombre_archivo)
        
        # Verificar si el archivo existe en la ubicación de origen
        if not os.path.exists(archivo_origen):
            print(f"Error: El archivo '{nombre_archivo}' no existe en la ruta de origen.")
            return False
        
        # Crear la carpeta de destino si no existe
        if not os.path.exists(ruta_destino):
            os.makedirs(ruta_destino)
            print(f"Carpeta de destino creada: {ruta_destino}")
        
        # Verificar si el archivo ya existe en el destino
        if os.path.exists(archivo_destino):
            # Obtener las fechas de modificación
            fecha_origen = os.path.getmtime(archivo_origen)
            fecha_destino = os.path.getmtime(archivo_destino)
            
            # Si el archivo origen es más reciente, actualizar
            if fecha_origen > fecha_destino:
                shutil.copy2(archivo_origen, archivo_destino)
                print(f"Archivo actualizado: {nombre_archivo}")
                print(f"Fecha anterior: {datetime.fromtimestamp(fecha_destino)}")
                print(f"Nueva fecha: {datetime.fromtimestamp(fecha_origen)}")
            else:
                print(f"El archivo ya está actualizado: {nombre_archivo}")
        else:
            # Si el archivo no existe en el destino, copiarlo
            shutil.copy2(archivo_origen, archivo_destino)
            print(f"Archivo copiado por primera vez: {nombre_archivo}")
        
        # Opcional: Cargar el archivo Excel para verificar
        try:
            df = pd.read_excel(archivo_destino)
            print(f"Archivo Excel verificado. Contiene {len(df)} filas y {len(df.columns)} columnas.")
        except Exception as e:
            print(f"Advertencia: El archivo se copió pero no se pudo verificar el contenido: {str(e)}")
        
        return True
        
    except Exception as e:
        print(f"Error durante la sincronización: {str(e)}")
        return False

# Ejecución única
if __name__ == "__main__":
    # Configura estas rutas según tu sistema
    RUTA_ONEDRIVE = r"C:\Users\Administrator\OneDrive - Indra (1)\Facturas\Carpeta Archivos Adjuntos"
    RUTA_DESTINO = r"D:\facturas_bot\Maestra"
    NOMBRE_ARCHIVO = "Robot 2_Estructura Carpetas Factura SSFF 03.03.2025.xlsx"
    
    print(f"Sincronizando a las {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    sincronizar_excel_onedrive(RUTA_ONEDRIVE, RUTA_DESTINO, NOMBRE_ARCHIVO)
    print("Sincronización completada.")