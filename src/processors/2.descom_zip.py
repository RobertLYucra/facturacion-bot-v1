import os
import zipfile
import rarfile
import sys
import re
import logging
import shutil

# Agregar raíz del proyecto al path
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, PROJECT_ROOT)

from src.utils.email_log_module import actualizar_estado_log
from datetime import datetime
import io

# ==================== CONFIGURACIÓN UTF-8 PARA WINDOWS ====================
# Forzar UTF-8 en stdout y stderr para evitar errores de codificación
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
# ==========================================================================

# Configurar logging (relativo a raíz del proyecto)
log_directory = os.path.join(PROJECT_ROOT, "LOGS", "descom_zip")
os.makedirs(log_directory, exist_ok=True)
log_file = os.path.join(log_directory, f"descompresion_organizacion_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

# Configurar el logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("descompresion_organizacion")

# Log de inicio
logger.info("==================== INICIANDO SCRIPT HÍBRIDO: DESCOMPRESIÓN/ORGANIZACIÓN ====================")
logger.info(f"Argumentos recibidos: {sys.argv}")

# Verificar si se proporcionó un directorio como argumento
if len(sys.argv) > 1:
    DIRECTORIO_A_ANALIZAR = sys.argv[1]
    logger.info(f"Usando directorio proporcionado: {DIRECTORIO_A_ANALIZAR}")
    print(f"Usando directorio proporcionado: {DIRECTORIO_A_ANALIZAR}")
else:
    # Directorio por defecto (solo como respaldo si no hay argumento)
    DIRECTORIO_A_ANALIZAR = os.path.join(os.getcwd(), "inboxFacturas")
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
    sys.exit(1)
else:
    logger.info(f"Verificado: Directorio {DIRECTORIO_A_ANALIZAR} existe")


def analizar_contenido_directorio(directorio):
    """
    Analiza el contenido del directorio para determinar si hay archivos ZIP o archivos sueltos.
    Retorna un diccionario con el análisis.
    """
    archivos = os.listdir(directorio)
    analisis = {
        'archivos_zip': [],
        'archivos_rar': [],
        'archivos_sueltos': [],
        'otros_archivos': [],
        'directorios': []
    }
    
    logger.info(f"Analizando contenido del directorio (total: {len(archivos)} elementos):")
    
    for item in archivos:
        ruta_completa = os.path.join(directorio, item)
        
        if os.path.isdir(ruta_completa):
            analisis['directorios'].append(item)
        elif os.path.isfile(ruta_completa):
            # Excluir archivos del sistema y logs
            if (item.startswith('.') or 
                item.startswith('contenido_email') or
                item.startswith('tabla_') or
                item.startswith('debug_') or
                item.endswith('.log') or
                item.endswith('.txt')):
                analisis['otros_archivos'].append(item)
            elif item.lower().endswith('.zip'):
                analisis['archivos_zip'].append(item)
            elif item.lower().endswith('.rar'):
                analisis['archivos_rar'].append(item)
            else:
                # Archivos que podrían ser comprobantes sueltos
                analisis['archivos_sueltos'].append(item)
        
        # Logging detallado
        tipo = "Directorio" if os.path.isdir(ruta_completa) else "Archivo"
        tamano = os.path.getsize(ruta_completa) if os.path.isfile(ruta_completa) else 0
        logger.info(f"  - {tipo}: {item} ({tamano} bytes)")
    
    return analisis


def identificar_tipo_archivo(nombre_archivo):
    """
    Identifica el tipo de comprobante basado en patrones específicos:
    - CDR: archivos que inician con "R_" y tienen extensión .xml
    - PDF: archivos que contienen "F00" y tienen extensión .pdf
    - XML: archivos que contienen "F00" y tienen extensión .xml
    """
    nombre_upper = nombre_archivo.upper()
    
    # CDR: archivos que inician con "R_" y son .xml
    if nombre_archivo.startswith('R-') and nombre_archivo.lower().endswith('.xml'):
        return 'CDR'
    
    # PDF: archivos que contienen "F00" y son .pdf
    elif 'F00' in nombre_upper and nombre_archivo.lower().endswith('.pdf'):
        return 'PDF'
    
    # XML: archivos que contienen "F00" y son .xml
    elif 'F00' in nombre_upper and nombre_archivo.lower().endswith('.xml'):
        return 'XML'
    
    # Si no cumple ningún criterio, clasificar como OTROS
    return 'OTROS'


def descomprimir_archivos_zip(directorio, archivos_zip):
    """
    Descomprime archivos ZIP y los organiza por tipo de comprobante usando los nuevos criterios.
    """
    logger.info(f"=== MODO DESCOMPRESIÓN: Procesando {len(archivos_zip)} archivos ZIP ===")
    
    # Crear carpetas destino para cada tipo
    carpetas_tipo = {
        "CDR": os.path.join(directorio, "comprobantes_CDR"),
        "PDF": os.path.join(directorio, "comprobantes_PDF"),
        "XML": os.path.join(directorio, "comprobantes_XML"),
        "OTROS": os.path.join(directorio, "comprobantes_OTROS")
    }
    for ruta in carpetas_tipo.values():
        os.makedirs(ruta, exist_ok=True)

    archivos_procesados = 0
    archivos_organizados = {tipo: 0 for tipo in carpetas_tipo.keys()}

    for archivo in archivos_zip:
        ruta_zip = os.path.join(directorio, archivo)
        logger.info(f" Procesando ZIP: {archivo}")

        try:
            with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
                for member in zip_ref.infolist():
                    if member.is_dir():
                        continue

                    nombre_archivo = os.path.basename(member.filename)
                    
                    # Usar la nueva lógica de identificación de tipos
                    tipo = identificar_tipo_archivo(nombre_archivo)
                    carpeta_destino = carpetas_tipo[tipo]
                    
                    logger.info(f"  - Archivo '{nombre_archivo}' clasificado como: {tipo}")

                    destino_final = os.path.join(carpeta_destino, nombre_archivo)

                    # Evitar sobrescritura
                    contador = 1
                    while os.path.exists(destino_final):
                        nombre_base, ext = os.path.splitext(nombre_archivo)
                        destino_final = os.path.join(carpeta_destino, f"{nombre_base}_{contador}{ext}")
                        contador += 1

                    with zip_ref.open(member) as fuente, open(destino_final, 'wb') as destino:
                        destino.write(fuente.read())
                        logger.info(f"    OK  {nombre_archivo} -> {os.path.relpath(destino_final, directorio)}")
                        
                    archivos_organizados[tipo] += 1

            archivos_procesados += 1
        except zipfile.BadZipFile:
            logger.error(f" Archivo ZIP inválido: {archivo}")
        except Exception as e:
            logger.error(f" Error al extraer '{archivo}': {str(e)}")

    # Mostrar resumen de descompresión
    logger.info(f"RESUMEN DE DESCOMPRESIÓN:")
    for tipo, cantidad in archivos_organizados.items():
        if cantidad > 0:
            logger.info(f"  - {tipo}: {cantidad} archivos")

    # Limpiar carpetas vacías
    carpetas_creadas = []
    for tipo, ruta_carpeta in carpetas_tipo.items():
        if archivos_organizados[tipo] > 0:
            carpetas_creadas.append(ruta_carpeta)
        else:
            # Eliminar carpeta si está vacía
            try:
                if os.path.exists(ruta_carpeta) and len(os.listdir(ruta_carpeta)) == 0:
                    os.rmdir(ruta_carpeta)
                    logger.info(f"🗑️ Carpeta vacía eliminada: {os.path.basename(ruta_carpeta)}")
            except Exception as e:
                logger.warning(f"No se pudo eliminar carpeta vacía {ruta_carpeta}: {str(e)}")

    logger.info(f" {archivos_procesados} ZIPs procesados correctamente.")
    return carpetas_creadas


def organizar_archivos_sueltos(directorio, archivos_sueltos):
    """
    Organiza archivos sueltos (no comprimidos) en carpetas por tipo de comprobante.
    """
    logger.info(f"=== MODO ORGANIZACIÓN: Procesando {len(archivos_sueltos)} archivos sueltos ===")
    
    # Crear carpetas destino para cada tipo
    carpetas_tipo = {
        "CDR": os.path.join(directorio, "comprobantes_CDR"),
        "PDF": os.path.join(directorio, "comprobantes_PDF"),
        "XML": os.path.join(directorio, "comprobantes_XML"),
        "OTROS": os.path.join(directorio, "comprobantes_OTROS")
    }
    
    # Crear las carpetas si no existen
    for ruta in carpetas_tipo.values():
        os.makedirs(ruta, exist_ok=True)

    archivos_procesados = 0
    archivos_organizados = {tipo: 0 for tipo in carpetas_tipo.keys()}
    
    for archivo in archivos_sueltos:
        ruta_archivo = os.path.join(directorio, archivo)
        
        # Identificar tipo de archivo
        tipo = identificar_tipo_archivo(archivo)
        carpeta_destino = carpetas_tipo[tipo]
        
        # Mostrar detalle de la clasificación
        detalle_clasificacion = ""
        if tipo == 'CDR':
            detalle_clasificacion = f"(inicia con 'R_' y es .xml)"
        elif tipo == 'PDF':
            detalle_clasificacion = f"(contiene 'F00' y es .pdf)"
        elif tipo == 'XML':
            detalle_clasificacion = f"(contiene 'F00' y es .xml)"
        else:
            detalle_clasificacion = f"(no cumple criterios específicos)"
        
        logger.info(f" Procesando archivo '{archivo}' como tipo {tipo} {detalle_clasificacion}")
        print(f" Procesando archivo '{archivo}' como tipo {tipo} {detalle_clasificacion}")

        try:
            # Crear ruta de destino
            destino_final = os.path.join(carpeta_destino, archivo)
            
            # Evitar sobrescritura si el archivo ya existe en destino
            contador = 1
            nombre_base, extension = os.path.splitext(archivo)
            while os.path.exists(destino_final):
                nuevo_nombre = f"{nombre_base}_{contador}{extension}"
                destino_final = os.path.join(carpeta_destino, nuevo_nombre)
                contador += 1

            # Mover el archivo
            shutil.move(ruta_archivo, destino_final)
            logger.info(f"OK  {archivo} -> {os.path.relpath(destino_final, directorio)}")
            print(f"OK  {archivo} -> {os.path.relpath(destino_final, directorio)}")
            
            archivos_procesados += 1
            archivos_organizados[tipo] += 1
            
        except Exception as e:
            logger.error(f" Error al mover '{archivo}': {str(e)}")
            print(f" Error al mover '{archivo}': {str(e)}")

    # Mostrar resumen
    logger.info(f"RESUMEN DE ORGANIZACIÓN:")
    for tipo, cantidad in archivos_organizados.items():
        if cantidad > 0:
            logger.info(f"  - {tipo}: {cantidad} archivos")
    
    # Limpiar carpetas vacías
    carpetas_creadas = []
    for tipo, ruta_carpeta in carpetas_tipo.items():
        if archivos_organizados[tipo] > 0:
            carpetas_creadas.append(ruta_carpeta)
        else:
            # Eliminar carpeta si está vacía
            try:
                if os.path.exists(ruta_carpeta) and len(os.listdir(ruta_carpeta)) == 0:
                    os.rmdir(ruta_carpeta)
                    logger.info(f"Carpeta vacía eliminada: {os.path.basename(ruta_carpeta)}")
            except Exception as e:
                logger.warning(f"No se pudo eliminar carpeta vacía {ruta_carpeta}: {str(e)}")

    logger.info(f" {archivos_procesados} archivos sueltos organizados correctamente.")
    return carpetas_creadas


def renombrar_carpetas_sin_fechas(carpetas):
    """
    Renombra carpetas para mantener consistencia (elimina partes innecesarias del nombre).
    """
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
                logger.info(f" Prefijo agregado: '{nombre_carpeta}' -> '{nuevo_nombre_carpeta}'")
                carpeta = nueva_ruta_comprobantes
                nombre_carpeta = nuevo_nombre_carpeta
            except Exception as e:
                logger.error(f"Error al agregar 'comprobantes' a '{nombre_carpeta}': {str(e)}")
                continue

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
                logger.info(f" Carpeta renombrada: '{nombre_carpeta}' -> '{os.path.basename(nueva_ruta)}'")
                carpetas_renombradas += 1
            except Exception as e:
                logger.error(f" No se pudo renombrar '{nombre_carpeta}': {str(e)}")

    logger.info(f"\nResumen: {carpetas_renombradas} carpetas renombradas")
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
        directorio_base = os.path.dirname(directorio)
        logger.info(f"Directorio base del Excel: {directorio_base}")
        
        # Actualizar el estado en el Excel
        resultado = actualizar_estado_log(asunto, "2.Descompresion", "Completo", directorio_base)
        if resultado:
            logger.info(f" Estado actualizado en Excel: 2.Descompresion -> Completo")
            print(f" Estado actualizado en Excel: 2.Descompresion -> Completo")
        else:
            logger.warning(" actualizar_estado_log devolvió False - posible error")
            print(" actualizar_estado_log devolvió False - posible error")
        return resultado
    except Exception as e:
        logger.error(f" Error al actualizar estado en Excel: {str(e)}", exc_info=True)
        print(f" Error al actualizar estado en Excel: {str(e)}")
        return False


if __name__ == "__main__":
    logger.info("\n========== INICIANDO PROCESAMIENTO HÍBRIDO ==========")
    print("\n========== INICIANDO PROCESAMIENTO HÍBRIDO ==========")
    
    # Registrar valores importantes
    logger.info(f"Directorio a analizar: {DIRECTORIO_A_ANALIZAR}")
    logger.info(f"Fila Excel: {fila_excel}")
    
    # PASO 1: Analizar contenido del directorio
    analisis = analizar_contenido_directorio(DIRECTORIO_A_ANALIZAR)
    
    # PASO 2: Determinar modo de operación
    carpetas_creadas = []
    
    if analisis['archivos_zip']:
        logger.info(f" DETECCIÓN: {len(analisis['archivos_zip'])} archivos ZIP encontrados")
        logger.info(f" Archivos ZIP: {analisis['archivos_zip']}")
        print(f" DETECCIÓN: {len(analisis['archivos_zip'])} archivos ZIP encontrados")
        
        # Ejecutar descompresión
        carpetas_creadas = descomprimir_archivos_zip(DIRECTORIO_A_ANALIZAR, analisis['archivos_zip'])
        
    if analisis['archivos_sueltos']:
        logger.info(f" DETECCIÓN: {len(analisis['archivos_sueltos'])} archivos sueltos encontrados")
        logger.info(f" Archivos sueltos: {analisis['archivos_sueltos']}")
        print(f" DETECCIÓN: {len(analisis['archivos_sueltos'])} archivos sueltos encontrados")
        
        # Ejecutar organización de archivos sueltos
        carpetas_organizacion = organizar_archivos_sueltos(DIRECTORIO_A_ANALIZAR, analisis['archivos_sueltos'])
        carpetas_creadas.extend(carpetas_organizacion)
    
    if not analisis['archivos_zip'] and not analisis['archivos_sueltos']:
        logger.info(" No se encontraron archivos ZIP ni archivos sueltos para procesar")
        print(" No se encontraron archivos ZIP ni archivos sueltos para procesar")
    
    # PASO 3: Renombrar carpetas creadas
    if carpetas_creadas:
        logger.info(f"Carpetas creadas: {carpetas_creadas}")
        renombrar_carpetas_sin_fechas(carpetas_creadas)
        
        # PASO 4: Actualizar estado en Excel
        if fila_excel is not None:
            logger.info(f"Actualizando estado en Excel para fila {fila_excel}")
            actualizar_estado_en_excel(DIRECTORIO_A_ANALIZAR)
            logger.info(f"Actualizado estado en Excel para el correo en la fila {fila_excel}")
            print(f"Actualizado estado en Excel para el correo en la fila {fila_excel}")
        else:
            logger.warning("No se actualizó el estado en Excel porque no se proporcionó número de fila")
        
        logger.info("\n========== PROCESAMIENTO COMPLETADO CON ÉXITO ==========")
        print("\n========== PROCESAMIENTO COMPLETADO CON ÉXITO ==========")
        
    else:
        logger.info("No se crearon carpetas - posiblemente no había archivos para procesar")
        print("No se crearon carpetas - posiblemente no había archivos para procesar")
        
        # Si no hay archivos para procesar, aún así marcar como completado
        if fila_excel is not None:
            actualizar_estado_en_excel(DIRECTORIO_A_ANALIZAR)
        
        logger.info("\n========== PROCESAMIENTO COMPLETADO SIN CAMBIOS ==========")
        print("\n========== PROCESAMIENTO COMPLETADO SIN CAMBIOS ==========")
    
    # Siempre devolver el directorio para continuidad del flujo
    logger.info(f"OUTPUT_DIRECTORY={DIRECTORIO_A_ANALIZAR}")
    print(f"OUTPUT_DIRECTORY={DIRECTORIO_A_ANALIZAR}")
        
    # Log de cierre
    logger.info("==================== FINALIZADO SCRIPT HÍBRIDO ====================")