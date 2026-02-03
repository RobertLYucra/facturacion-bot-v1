import subprocess
import os
import logging
import time
import sys

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='ejecucion_automatizada.log',
    filemode='a'
)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
logging.getLogger().addHandler(console_handler)

logger = logging.getLogger('script_principal')

def ejecutar_script(nombre_script):
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(base_dir, nombre_script)
        
        if not os.path.exists(script_path):
            logger.error(f"El archivo no existe: {script_path}")
            return False

        logger.info(f"Iniciando ejecución de {nombre_script}")
        
        resultado = subprocess.run(
            [sys.executable, script_path], 
            capture_output=True, 
            text=True, 
            check=True,
            cwd=base_dir 
        )
        
        logger.info(f"Script {nombre_script} ejecutado exitosamente")
        return True

    except subprocess.CalledProcessError as e:
        logger.error(f"Error en {nombre_script}. Código: {e.returncode}")
        logger.error(f"Salida Error: {e.stderr}")
        return False
    except Exception as e:
        logger.error(f"Excepción: {str(e)}")
        return False

def main():
    scripts_en_orden = [
        "SyncMaestra.py",
        "SyncArchivoCompartidos.py",
        "SyncHistorico.py",
        "SendEmail.py",
        "SendRegistroHistorico.py"
    ]
    
    logger.info("=== INICIANDO EJECUCIÓN ===")
    
    for script in scripts_en_orden:
        exito = ejecutar_script(script)
        
        if not exito:
            break
        
        time.sleep(2)
    
    logger.info("=== FINALIZADO ===")

if __name__ == "__main__":
    main()

