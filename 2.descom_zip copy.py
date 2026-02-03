import os
import zipfile
import rarfile
import sys
import re
import logging
from email_log_module import actualizar_estado_log
from datetime import datetime
from registro_errores import registrar_log_detallado

# Configurar logging
log_directory = "logs"
os.makedirs(log_directory, exist_ok=True)
log_file = os.path.join(log_directory, f"descompresion_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

# Configurar el logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("descompresion")

# Log de inicio
logger.info("==================== INICIANDO SCRIPT DE DESCOMPRESIÓN ====================")
logger.info(f"Argumentos recibidos: {sys.argv}")

# Verificar si se proporcionó un directorio como argumento
if len(sys.argv) > 1:
    DIRECTORIO_A_ANALIZAR = sys.argv[1]
    logger.info(f"Usando directorio proporcionado: {DIRECTORIO_A_ANALIZAR}")
    print(f"Usando directorio proporcionado: {DIRECTORIO_A_ANALIZAR}")
else:
    # Directorio por defecto (solo como respaldo si no hay argumento)
    DIRECTORIO_A_ANALIZAR = "/Volumes/diskZ/INDRA/facturas_bot/inboxFacturas/RV_ Facturación Perú 19.03.2025"
    logger.warning(f"No se proporcionó directorio, usando directorio por defecto: {DIRECTORIO_A_ANALIZAR}")
    print(f"No se proporcionó directorio, usando directorio por defecto: {DIRECTORIO_A_ANALIZAR}")

# Verificar si se proporcionó la fila del Excel como segundo argumento
fila_excel = None
if len(sys.argv) >= 3:
    try:
        fila_excel = int(sys.argv[2])
        logger.info(f"RECIBIDO - Procesando correo de la fila {fila_excel} del Excel")
        print(f"RECIBIDO - Procesando correo de la fila {fila_excel} del Excel")
    except ValueError:
        logger.error(f"Error: Valor inválido para fila_excel: {sys.argv[2]}")
        print(f"Error: Valor inválido para fila_excel: {sys.argv[2]}")
else:
    logger.warning("No se proporcionó número de fila del Excel")

# Verificar existencia del directorio
if not os.path.exists(DIRECTORIO_A_ANALIZAR):
    logger.critical(f"ERROR CRÍTICO: El directorio {DIRECTORIO_A_ANALIZAR} no existe!")
    print(f"ERROR CRÍTICO: El directorio {DIRECTORIO_A_ANALIZAR} no existe!")
else:
    logger.info(f"Verificado: Directorio {DIRECTORIO_A_ANALIZAR} existe")
    # Mostrar el contenido del directorio
    archivos = os.listdir(DIRECTORIO_A_ANALIZAR)
    logger.info(f"Contenido del directorio (total: {len(archivos)} elementos):")
    for archivo in archivos:
        ruta_completa = os.path.join(DIRECTORIO_A_ANALIZAR, archivo)
        tipo = "Directorio" if os.path.isdir(ruta_completa) else "Archivo"
        tamano = os.path.getsize(ruta_completa) if os.path.isfile(ruta_completa) else 0
        logger.info(f"  - {tipo}: {archivo} ({tamano} bytes)")

def descomprimir_archivos(directorio,asunto_correo):
    logger.info(f"Iniciando descompresión en: {directorio}")
    
    if not os.path.isdir(directorio):
        descomprimir_archivos_msg = f"Error: El directorio '{directorio}' no existe."
        logger.error(descomprimir_archivos_msg)
        print(descomprimir_archivos_msg)
        registrar_log_detallado(asunto_correo, "2.Descompresion", "Error", descomprimir_archivos_msg)
        return False  # Retornar False para indicar error

    archivos_procesados = 0
    carpetas_creadas = []
    
    archivos_en_directorio = os.listdir(directorio)
    logger.info(f"Total de archivos en el directorio: {len(archivos_en_directorio)}")
    
    # Registrar todos los archivos zip/rar encontrados
    archivos_comprimidos = [a for a in archivos_en_directorio if a.lower().endswith(('.zip', '.rar'))]
    logger.info(f"Archivos comprimidos encontrados: {len(archivos_comprimidos)}")
    for archivo in archivos_comprimidos:
        logger.info(f"  - Archivo comprimido: {archivo}")
    
    for archivo in archivos_en_directorio:
        ruta_completa = os.path.join(directorio, archivo)

        if archivo.lower().endswith('.zip'):
            logger.info(f"Procesando archivo ZIP: {archivo}")
            carpeta_destino = os.path.join(directorio, archivo[:-4])  # Quita la extensión
            os.makedirs(carpeta_destino, exist_ok=True)

            try:
                with zipfile.ZipFile(ruta_completa, 'r') as zip_ref:
                    archivos_en_zip = zip_ref.namelist()
                    logger.info(f"Contenido del ZIP ({len(archivos_en_zip)} archivos):")
                    for arch in archivos_en_zip[:10]:  # Limitar a 10 para no llenar el log
                        logger.info(f"  - {arch}")
                    if len(archivos_en_zip) > 10:
                        logger.info(f"  ... y {len(archivos_en_zip) - 10} archivos más")
                    
                    zip_ref.extractall(carpeta_destino)
                
                logger.info(f"✔ Archivo ZIP '{archivo}' descomprimido en '{carpeta_destino}'")
                print(f"✔ Archivo ZIP '{archivo}' descomprimido en '{carpeta_destino}'")
                archivos_procesados += 1
                carpetas_creadas.append(carpeta_destino)
            except zipfile.BadZipFile:
                error_msg = f"Error: '{archivo}' no es un archivo ZIP válido."
                logger.error(error_msg)
                print(error_msg)
                registrar_log_detallado(asunto_correo, "2.Descompresion", "Error", error_msg)

        elif archivo.lower().endswith('.rar'):
            logger.info(f"Procesando archivo RAR: {archivo}")
            carpeta_destino = os.path.join(directorio, archivo[:-4])  # Quita la extensión
            os.makedirs(carpeta_destino, exist_ok=True)

            try:
                with rarfile.RarFile(ruta_completa, 'r') as rar_ref:
                    archivos_en_rar = rar_ref.namelist()
                    logger.info(f"Contenido del RAR ({len(archivos_en_rar)} archivos):")
                    for arch in archivos_en_rar[:10]:  # Limitar a 10 para no llenar el log
                        logger.info(f"  - {arch}")
                    if len(archivos_en_rar) > 10:
                        logger.info(f"  ... y {len(archivos_en_rar) - 10} archivos más")
                    
                    rar_ref.extractall(carpeta_destino)
                
                logger.info(f"✔ Archivo RAR '{archivo}' descomprimido en '{carpeta_destino}'")
                print(f"✔ Archivo RAR '{archivo}' descomprimido en '{carpeta_destino}'")
                archivos_procesados += 1
                carpetas_creadas.append(carpeta_destino)
            except rarfile.BadRarFile:
                error_msg = f"Error: '{archivo}' no es un archivo RAR válido."
                logger.error(error_msg)
                print(error_msg)
                registrar_log_detallado(asunto_correo, "2.Descompresion", "Error", error_msg)
            except rarfile.NotRarFile:
                error_msg = f"❌ Error: '{archivo}' no es un archivo RAR."
                logger.error(error_msg)
                print(error_msg)
                registrar_log_detallado(asunto_correo, "2.Descompresion", "Error", error_msg)
    
    logger.info(f"\nResumen: {archivos_procesados} archivos descomprimidos en {directorio}")
    print(f"\nResumen: {archivos_procesados} archivos descomprimidos en {directorio}")
    
    if archivos_procesados > 0:
        registrar_log_detallado(asunto_correo, "2.Descompresion", "Éxito", "Archivos ZIP descomprimidos correctamente.")
    else:
        registrar_log_detallado(asunto_correo, "2.Descompresion", "Error", "No se descomprimió ningún archivo ZIP.")


    return carpetas_creadas  # Retorna la lista de carpetas creadas

def renombrar_carpetas_sin_fechas(carpetas):
    logger.info(f"Iniciando renombrado de {len(carpetas)} carpetas")
    carpetas_renombradas = 0

    for carpeta in carpetas:
        nombre_carpeta = os.path.basename(carpeta)
        directorio_padre = os.path.dirname(carpeta)

        logger.info(f"Procesando carpeta: {nombre_carpeta}")

        # Paso 1: Validar si empieza con "comprobantes"
        if not nombre_carpeta.lower().startswith("comprobantes"):
            nuevo_nombre_carpeta = f"comprobantes_{nombre_carpeta}"
            nueva_ruta_comprobantes = os.path.join(directorio_padre, nuevo_nombre_carpeta)

            try:
                os.rename(carpeta, nueva_ruta_comprobantes)
                logger.info(f"✅ Prefijo agregado: '{nombre_carpeta}' -> '{nuevo_nombre_carpeta}'")
                print(f"✅ Prefijo agregado: '{nombre_carpeta}' -> '{nuevo_nombre_carpeta}'")
                carpeta = nueva_ruta_comprobantes
                nombre_carpeta = nuevo_nombre_carpeta
            except Exception as e:
                logger.error(f"❌ Error al agregar 'comprobantes' a '{nombre_carpeta}': {str(e)}")
                print(f"❌ Error al agregar 'comprobantes' a '{nombre_carpeta}': {str(e)}")
                continue  # Saltar esta carpeta si no pudo renombrarse

        # Paso 2: Renombrar manteniendo solo los primeros dos bloques
        partes = nombre_carpeta.split('_')
        if len(partes) > 2:
            nuevo_nombre = '_'.join(partes[:2])
            logger.info(f"Partes detectadas: {partes}, usando las primeras 2: {partes[:2]}")
        else:
            nuevo_nombre = nombre_carpeta
            logger.info(f"Insuficientes partes para renombrar ({len(partes)} ≤ 2), manteniendo nombre original")

        nuevo_nombre = nuevo_nombre.strip()
        nueva_ruta = os.path.join(directorio_padre, nuevo_nombre)

        # Si la ruta destino existe, evitar conflicto
        if nuevo_nombre != nombre_carpeta:
            if os.path.exists(nueva_ruta):
                contador = 1
                while os.path.exists(f"{nueva_ruta}_{contador}"):
                    contador += 1
                nueva_ruta = f"{nueva_ruta}_{contador}"
                logger.info(f"Destino ya existe, añadiendo sufijo: '{nuevo_nombre}_{contador}'")

            try:
                os.rename(carpeta, nueva_ruta)
                logger.info(f"✅ Carpeta renombrada: '{nombre_carpeta}' -> '{os.path.basename(nueva_ruta)}'")
                print(f"✅ Carpeta renombrada: '{nombre_carpeta}' -> '{os.path.basename(nueva_ruta)}'")
                carpetas_renombradas += 1
            except Exception as e:
                logger.error(f"⚠️ No se pudo renombrar '{nombre_carpeta}': {str(e)}")
                print(f"⚠️ No se pudo renombrar '{nombre_carpeta}': {str(e)}")
        else:
            logger.info(f"No es necesario renombrar (nombre actual es igual al nuevo nombre)")

    logger.info(f"\nResumen: {carpetas_renombradas} carpetas renombradas eliminando lo posterior al segundo '_'")
    print(f"\nResumen: {carpetas_renombradas} carpetas renombradas eliminando lo posterior al segundo '_'")
    return carpetas_renombradas

def actualizar_estado_en_excel(directorio):
    """
    Actualiza el estado de la etapa "2.Descompresion" en el Excel.
    """
    try:
        # Obtener el asunto del correo (nombre del directorio)
        asunto = os.path.basename(directorio)
        logger.info(f"Actualizando estado en Excel para correo con asunto: {asunto}")
        
        # Obtener el directorio base donde está el Excel
        directorio_base = os.path.dirname(os.path.dirname(directorio))
        logger.info(f"Directorio base del Excel: {directorio_base}")
        
        # Actualizar el estado en el Excel
        resultado = actualizar_estado_log(asunto, "2.Descompresion", "Completo", directorio_base)
        if resultado:
            logger.info(f"✅ Estado actualizado en Excel: 2.Descompresion -> Completo")
            print(f"✅ Estado actualizado en Excel: 2.Descompresion -> Completo")
        else:
            logger.warning("⚠️ actualizar_estado_log devolvió False - posible error")
            print("⚠️ actualizar_estado_log devolvió False - posible error")
        return resultado
    except Exception as e:
        logger.error(f"❌ Error al actualizar estado en Excel: {str(e)}", exc_info=True)
        print(f"❌ Error al actualizar estado en Excel: {str(e)}")
        return False

if __name__ == "__main__":
    logger.info("\n========== INICIANDO DESCOMPRESIÓN DE ARCHIVOS ==========")
    print("\n========== INICIANDO DESCOMPRESIÓN DE ARCHIVOS ==========")
    
    # Registrar valores importantes
    logger.info(f"Directorio a analizar: {DIRECTORIO_A_ANALIZAR}")
    logger.info(f"Fila Excel: {fila_excel}")
    
    # Ejecutar descompresión
    asunto_correo = os.environ.get("ASUNTO_CORREO", "Asunto desconocido")
    carpetas_creadas = descomprimir_archivos(DIRECTORIO_A_ANALIZAR,asunto_correo)
    
    if carpetas_creadas:
        logger.info(f"Carpetas creadas durante la descompresión: {carpetas_creadas}")
        
        # Renombrar carpetas
        renombrar_carpetas_sin_fechas(carpetas_creadas)
        
        # Actualizar estado en Excel
        if fila_excel is not None:
            logger.info(f"Actualizando estado en Excel para fila {fila_excel}")
            actualizar_estado_en_excel(DIRECTORIO_A_ANALIZAR)
            logger.info(f"Actualizado estado en Excel para el correo en la fila {fila_excel}")
            print(f"Actualizado estado en Excel para el correo en la fila {fila_excel}")
        else:
            logger.warning("No se actualizó el estado en Excel porque no se proporcionó número de fila")
        
        logger.info("\n========== PROCESAMIENTO COMPLETADO CON ÉXITO ==========")
        print("\n========== PROCESAMIENTO COMPLETADO CON ÉXITO ==========")
        
        # Imprimir mensaje especial para que el script principal capture el directorio
        logger.info(f"OUTPUT_DIRECTORY={DIRECTORIO_A_ANALIZAR}")
        print(f"OUTPUT_DIRECTORY={DIRECTORIO_A_ANALIZAR}")
    else:
        logger.info("\n========== PROCESAMIENTO COMPLETADO SIN CAMBIOS ==========")
        print("\n========== PROCESAMIENTO COMPLETADO SIN CAMBIOS ==========")
        # Igualmente devolvemos el directorio
        logger.info(f"OUTPUT_DIRECTORY={DIRECTORIO_A_ANALIZAR}")
        print(f"OUTPUT_DIRECTORY={DIRECTORIO_A_ANALIZAR}")
        
    # Log de cierre
    logger.info("==================== FINALIZADO SCRIPT DE DESCOMPRESIÓN ====================")