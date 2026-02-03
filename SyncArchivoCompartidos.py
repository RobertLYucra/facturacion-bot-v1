import os
import shutil
from datetime import datetime

def sincronizar_directorio(ruta_origen, ruta_destino):
    """
    Sincroniza un directorio completo desde una ubicación a otra.
    
    Args:
        ruta_origen (str): Ruta del directorio origen a sincronizar
        ruta_destino (str): Ruta del directorio destino donde se copiarán los archivos
    
    Returns:
        bool: True si la sincronización fue exitosa, False en caso contrario
    """
    try:
        # Verificar si el directorio origen existe
        if not os.path.exists(ruta_origen):
            print(f"Error: El directorio origen '{ruta_origen}' no existe.")
            return False
        
        # Crear el directorio destino si no existe
        if not os.path.exists(ruta_destino):
            os.makedirs(ruta_destino)
            print(f"Directorio destino creado: {ruta_destino}")
        
        # Contador de archivos sincronizados
        archivos_actualizados = 0
        archivos_nuevos = 0
        archivos_ignorados = 0
        
        # Recorrer todos los archivos y subdirectorios en la ruta de origen
        for raiz, dirs, archivos in os.walk(ruta_origen):
            # Calcular la ruta relativa para usar en el destino
            ruta_relativa = os.path.relpath(raiz, ruta_origen)
            ruta_destino_actual = os.path.join(ruta_destino, ruta_relativa)
            
            # Crear subdirectorios en destino si no existen
            if not os.path.exists(ruta_destino_actual):
                os.makedirs(ruta_destino_actual)
                print(f"Creado subdirectorio: {ruta_destino_actual}")
            
            # Procesar cada archivo en el directorio actual
            for archivo in archivos:
                archivo_origen = os.path.join(raiz, archivo)
                archivo_destino = os.path.join(ruta_destino_actual, archivo)
                
                # Verificar si el archivo ya existe en el destino
                if os.path.exists(archivo_destino):
                    # Obtener las fechas de modificación
                    fecha_origen = os.path.getmtime(archivo_origen)
                    fecha_destino = os.path.getmtime(archivo_destino)
                    
                    # Si el archivo origen es más reciente, actualizar
                    if fecha_origen > fecha_destino:
                        shutil.copy2(archivo_origen, archivo_destino)
                        archivos_actualizados += 1
                        print(f"Archivo actualizado: {os.path.join(ruta_relativa, archivo)}")
                    else:
                        archivos_ignorados += 1
                else:
                    # Si el archivo no existe en el destino, copiarlo
                    shutil.copy2(archivo_origen, archivo_destino)
                    archivos_nuevos += 1
                    print(f"Archivo nuevo copiado: {os.path.join(ruta_relativa, archivo)}")
        
        # Informe final
        print(f"\nSincronización completada:")
        print(f"- Archivos nuevos: {archivos_nuevos}")
        print(f"- Archivos actualizados: {archivos_actualizados}")
        print(f"- Archivos sin cambios: {archivos_ignorados}")
        print(f"- Total procesado: {archivos_nuevos + archivos_actualizados + archivos_ignorados}")
        
        return True
        
    except Exception as e:
        print(f"Error durante la sincronización del directorio: {str(e)}")
        return False

# Ejecución única
if __name__ == "__main__":
    # Configura estas rutas según tu sistema
    RUTA_ONEDRIVE = r"C:\Users\Administrator\OneDrive - Indra (1)\Facturas\Carpeta Archivos Adjuntos\Documentos de Facturación SSFF"
    RUTA_DESTINO = r"D:\facturas_bot\Documentos de Facturación SSFF"
    
    print(f"Sincronizando directorio a las {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    sincronizar_directorio(RUTA_ONEDRIVE, RUTA_DESTINO)
    print("Sincronización completada.")