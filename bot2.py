import subprocess
import os
import time
import sys
import imaplib
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from src.utils.log_manager import LogManager

# Configuración de correo (igual que en 1.inbox.py)
EMAIL = "alertasflm@indra.es"
PASSWORD = "es8EaB63"
SERVER = "imap.indra.es"
PORT_EMAIL = 993
FILE_EMAIL_PENDIENTE = "BOT2-PENDIENTES"

# Inicializar LogManager
log_manager = LogManager()
logger = log_manager.global_logger

def verificar_correos_pendientes():
    """
    Verifica si hay correos en la carpeta sin procesarlos.
    
    Returns:
        tuple: (hay_correos: bool, cantidad: int)
    """
    try:
        print(f"\n>>> Conectando al servidor de correo...")
        logger.info(f"Verificando correos en carpeta: {FILE_EMAIL_PENDIENTE}")
        
        # Conectar al servidor
        mail = imaplib.IMAP4_SSL(SERVER, PORT_EMAIL)
        mail.login(EMAIL, PASSWORD)
        
        # Seleccionar carpeta
        try:
            mail.select(FILE_EMAIL_PENDIENTE)
        except Exception:
            try:
                mail.select(FILE_EMAIL_PENDIENTE.encode('utf-8'))
            except Exception:
                mail.select(f"{EMAIL}/{FILE_EMAIL_PENDIENTE}")
        
        # Buscar todos los correos
        status, data = mail.search(None, "ALL")
        mail_ids = data[0].split()
        cantidad = len(mail_ids)
        
        # Cerrar conexión
        mail.logout()
        
        hay_correos = cantidad > 0
        
        if hay_correos:
            print(f"✅ Se encontraron {cantidad} correo(s) pendiente(s)")
            logger.info(f"Se encontraron {cantidad} correo(s) en la carpeta {FILE_EMAIL_PENDIENTE}")
        else:
            print(f"ℹ️  No hay correos pendientes en la carpeta '{FILE_EMAIL_PENDIENTE}'")
            logger.info(f"No hay correos en la carpeta {FILE_EMAIL_PENDIENTE}")
        
        return hay_correos, cantidad
        
    except Exception as e:
        error_msg = f"Error al verificar correos: {str(e)}"
        print(f"❌ {error_msg}")
        logger.error(error_msg)
        logger.exception("Detalle del error:")
        return False, 0

def ejecutar_script(nombre_script, argumentos=None):
    """Ejecuta un script de Python y registra el resultado.
    
    Args:
        nombre_script: Nombre del script a ejecutar
        argumentos: Lista de argumentos para pasar al script
        
    Returns:
        tuple: (éxito, stdout)
    """
    try:
        script_logger = log_manager.get_logger_for_email(nombre_script, "ejecucion")
        
        print(f"\n>>> EJECUTANDO: {nombre_script}")
        script_logger.info(f"Iniciando ejecución de {nombre_script}")
        logger.info(f"Iniciando ejecución de {nombre_script}")
        
        # Obtener la ruta completa al script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(script_dir, nombre_script)
        
        # Verificar si el script existe
        if not os.path.exists(script_path):
            error_msg = f"El script {nombre_script} no existe en {script_path}"
            print(f"ERROR: {error_msg}")
            script_logger.error(error_msg)
            logger.error(error_msg)
            return False, ""
        
        # Preparar el comando
        comando = [sys.executable, script_path]  # Usa sys.executable
        if argumentos:
            comando.extend(argumentos)
            
        # Mostrar el comando completo
        cmd_str = " ".join(comando)
        print(f">>> Ejecutando: {cmd_str}")
        print(f">>> Directorio de trabajo: {script_dir}")
        script_logger.info(f"Comando: {cmd_str}")
        script_logger.info(f"Directorio de trabajo: {script_dir}")
        
        # Configurar las variables de entorno para forzar UTF-8
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        
        # Ejecutar el script con encoding UTF-8
        proceso = subprocess.Popen(
            comando, 
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8',  # <-- CRÍTICO: Forzar UTF-8
            errors='replace',  # <-- CRÍTICO: Reemplazar caracteres problemáticos
            cwd=script_dir,
            env=env
        )
        
        # Capturar la salida
        stdout, stderr = proceso.communicate()
        
        # Asegurar que stdout y stderr nunca sean None
        stdout = stdout if stdout is not None else ""
        stderr = stderr if stderr is not None else ""
        
        # Guardar la salida en archivos
        output_dir = os.path.join(log_manager.base_dir, 'outputs')
        os.makedirs(output_dir, exist_ok=True)
        
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        stdout_filename = os.path.join(output_dir, f"{nombre_script.replace('.py', '')}_{timestamp}_stdout.log")
        stderr_filename = os.path.join(output_dir, f"{nombre_script.replace('.py', '')}_{timestamp}_stderr.log")
        
        # Guardar salida estándar
        with open(stdout_filename, 'w', encoding='utf-8') as f:  # <-- UTF-8
            f.write(stdout)
        script_logger.info(f"Salida estándar guardada en: {stdout_filename}")
        
        # Guardar salida de error si existe
        if stderr:
            with open(stderr_filename, 'w', encoding='utf-8') as f:  # <-- UTF-8
                f.write(stderr)
            script_logger.info(f"Salida de error guardada en: {stderr_filename}")
        
        # Mostrar la salida estándar
        if stdout:
            print(stdout.strip())
        
        # Verificar si hubo error
        if proceso.returncode != 0:
            error_msg = f"Error al ejecutar {nombre_script}: código de salida {proceso.returncode}"
            print(f"\n{'='*60}")
            print(f"❌ ERROR en {nombre_script}")
            print(f"{'='*60}")
            print(f"CÓDIGO DE SALIDA: {proceso.returncode}")
            if stdout:
                print(f"\n--- STDOUT ---")
                print(stdout)
            if stderr:
                print(f"\n--- STDERR ---")
                print(stderr)
            print(f"{'='*60}\n")
            
            script_logger.error(error_msg)
            script_logger.error(f"Salida de error: {stderr}")
            logger.error(error_msg)
            return False, stdout
            
        # Verificar si hay mensajes de error en la salida estándar
        if "error:" in stdout.lower() or "exception:" in stdout.lower():
            error_msg = f"Se detectaron mensajes de error en la salida de {nombre_script}"
            print(f"ERROR: {error_msg}")
            script_logger.error(error_msg)
            logger.error(error_msg)
            return False, stdout
            
        print(f">>> COMPLETADO: {nombre_script}")
        script_logger.info(f"Script {nombre_script} ejecutado exitosamente")
        logger.info(f"Script {nombre_script} ejecutado exitosamente")
        return True, stdout
        
    except Exception as e:
        error_msg = f"Excepción al ejecutar {nombre_script}: {str(e)}"
        print(f"EXCEPCIÓN: {error_msg}")
        logger.error(error_msg)
        logger.exception("Detalle del error:")
        return False, ""

def main():
    """Función principal que ejecuta el flujo de procesamiento de facturas"""
    
    ejecucion_logger = log_manager.get_logger_for_email("orquestador_facturas", "principal")
    
    print("\n" + "="*60)
    print("   ORQUESTADOR DE PROCESAMIENTO DE FACTURAS")
    print("="*60)
    ejecucion_logger.info("=== INICIANDO ORQUESTADOR DE FACTURAS ===")
    logger.info("=== INICIANDO ORQUESTADOR DE FACTURAS ===")
    
    # PASO 1: Verificar si hay correos pendientes (sin procesarlos)
    print("\n[PASO 1] Verificando si hay correos pendientes...")
    ejecucion_logger.info("PASO 1: Verificando correos pendientes en el inbox")
    
    hay_correos, cantidad = verificar_correos_pendientes()
    
    # Si NO hay correos, terminar aquí
    if not hay_correos:
        print("\n" + "="*60)
        print("ℹ️  NO HAY CORREOS PENDIENTES")
        print("   No hay nada que procesar.")
        print("   Proceso finalizado.")
        print("="*60)
        ejecucion_logger.info("No hay correos pendientes. Proceso finalizado.")
        logger.info("No hay correos pendientes. Proceso finalizado.")
        return
    
    # PASO 2: Si HAY correos, ejecutar scripts de sincronización primero
    print("\n" + "="*60)
    print(f"✅ HAY {cantidad} CORREO(S) PENDIENTE(S)")
    print("   Sincronizando archivos antes de procesar...")
    print("="*60)
    ejecucion_logger.info(f"Hay {cantidad} correo(s) pendiente(s). Iniciando sincronización.")
    logger.info(f"Hay {cantidad} correo(s) pendiente(s). Iniciando sincronización.")
    
    scripts_sync = [
        "src/services/SyncMaestra.py",
        "src/services/SyncHistorico.py",
        "src/services/SyncArchivoCompartidos.py"
    ]
    
    for i, script in enumerate(scripts_sync, start=2):
        print(f"\n[PASO {i}] Ejecutando {script}...")
        ejecucion_logger.info(f"PASO {i}: Ejecutando {script}")
        
        exito, stdout = ejecutar_script(script)
        
        if not exito:
            print(f"\n❌ Error al ejecutar {script}. Deteniendo proceso.")
            ejecucion_logger.error(f"Error al ejecutar {script}. Proceso detenido.")
            logger.error(f"Error al ejecutar {script}. Proceso detenido.")
            return  # Detener todo si falla la sincronización
        
        # Esperar entre scripts
        print(f">>> Esperando 2 segundos antes del siguiente script...")
        time.sleep(2)
    
    # PASO 3: Ahora sí, ejecutar 1.inbox.py para procesar los correos
    print("\n" + "="*60)
    print("✅ SINCRONIZACIÓN COMPLETADA")
    print("   Procesando correos...")
    print("="*60)
    print(f"\n[PASO {i+1}] Ejecutando 1.inbox.py para procesar correos...")
    ejecucion_logger.info(f"PASO {i+1}: Ejecutando 1.inbox.py")
    
    exito, stdout = ejecutar_script("src/processors/1.inbox.py")
    
    if not exito:
        print("\n❌ Error al procesar correos con 1.inbox.py")
        ejecucion_logger.error("Error al ejecutar 1.inbox.py")
        logger.error("Error al ejecutar 1.inbox.py")
        return
    
    print("\n" + "="*60)
    print("✅ PROCESO COMPLETO FINALIZADO")
    print("="*60)
    ejecucion_logger.info("=== PROCESO COMPLETO FINALIZADO ===")
    logger.info("=== PROCESO COMPLETO FINALIZADO ===")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.critical(f"ERROR CRÍTICO NO CONTROLADO: {str(e)}")
        logger.exception("Detalle del error crítico:")
        print(f"\n❌ ERROR CRÍTICO: {str(e)}")