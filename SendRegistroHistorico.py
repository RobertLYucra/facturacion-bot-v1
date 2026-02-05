import os
import shutil
import time
import logging
from datetime import datetime

# Crear directorio LOGS/send_registro_historico si no existe
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGS_DIR = os.path.join(SCRIPT_DIR, "LOGS", "send_registro_historico")
os.makedirs(LOGS_DIR, exist_ok=True)

# Configurar logging
logging.basicConfig(filename=os.path.join(LOGS_DIR, 'file_copy.log'), level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Rutas de los archivos
ARCHIVO_ORIGEN = "Registros Historico/3.Historicov2.xlsx"
ARCHIVO_DESTINO = r"C:\Users\Administrator\OneDrive - Indra (1)\Facturas\Carpeta Archivos Adjuntos\Registro Historico\3.Historicov2.xlsx"

def copiar_archivo(origen, destino):
    """
    Copia un archivo de origen a destino, reemplazando si ya existe.
    
    Args:
        origen (str): Ruta del archivo de origen
        destino (str): Ruta del archivo de destino
    
    Returns:
        bool: True si la copia fue exitosa, False en caso contrario
    """
    try:
        # Crear la carpeta de destino si no existe
        directorio_destino = os.path.dirname(destino)
        if not os.path.exists(directorio_destino):
            os.makedirs(directorio_destino)
            print(f"Carpeta de destino creada: {directorio_destino}")
            logging.info(f"Carpeta de destino creada: {directorio_destino}")
        
        # Verificar si el archivo de origen existe
        if not os.path.exists(origen):
            mensaje = f"Error: El archivo de origen no existe: {origen}"
            print(mensaje)
            logging.error(mensaje)
            return False
        
        # Copiar el archivo (reemplazando si ya existe)
        shutil.copy2(origen, destino)
        
        # Verificar que se haya copiado correctamente
        if os.path.exists(destino):
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            mensaje = f" [{timestamp}] Archivo copiado exitosamente de {origen} a {destino}"
            print(mensaje)
            logging.info(f"Archivo copiado exitosamente de {origen} a {destino}")
            return True
        else:
            mensaje = f" Error: No se pudo verificar la existencia del archivo en destino"
            print(mensaje)
            logging.error(mensaje)
            return False
            
    except Exception as e:
        mensaje = f" Error al copiar el archivo: {str(e)}"
        print(mensaje)
        logging.error(mensaje)
        return False

def ejecucion_periodica(intervalo_minutos=60):
    """
    Ejecuta la copia de archivo periódicamente
    
    Args:
        intervalo_minutos (int): Intervalo en minutos entre copias
    """
    try:
        print(f"Iniciando proceso de copia periódica cada {intervalo_minutos} minutos")
        print("Presione Ctrl+C para detener")
        
        while True:
            main()
            print(f"\nPróxima copia en {intervalo_minutos} minutos...")
            print("=" * 50)
            time.sleep(intervalo_minutos * 60)
            
    except KeyboardInterrupt:
        print("\nProceso detenido por el usuario")
    except Exception as e:
        print(f"\nError en el proceso periódico: {str(e)}")
        logging.error(f"Error en el proceso periódico: {str(e)}")

def main():
    print(f"Iniciando copia de archivo: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"De: {ARCHIVO_ORIGEN}")
    print(f"A: {ARCHIVO_DESTINO}")
    
    # Obtener rutas absolutas
    origen_abs = os.path.abspath(ARCHIVO_ORIGEN)
    
    resultado = copiar_archivo(origen_abs, ARCHIVO_DESTINO)
    
    if resultado:
        print(" Proceso completado con éxito")
    else:
        print(" Hubo errores durante el proceso")

if __name__ == "__main__":
    # Para una sola copia:
    main()
    
    # Para copias periódicas (cada 5 minutos):
    # ejecucion_periodica(5)