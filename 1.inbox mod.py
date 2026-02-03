import imaplib
import email
import sys
import io
import re
import html
import base64
import quopri
import os
import csv
import datetime
from email.header import decode_header
from bs4 import BeautifulSoup  # Necesitarás instalar BeautifulSoup4: pip install beautifulsoup4
from email_log_module import inicializar_log_excel, registrar_correo_log, actualizar_estado_log, obtener_cuerpo_correo
from log_manager import LogManager
import pandas as pd
import subprocess
import time
from registro_errores import registrar_log_detallado



sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

EMAIL = "alertasflm@indra.es"
PASSWORD = "LKo9J94e"

FILE_EMAIL_PENDIENTE = "BOT2-PRUEBAS" #"BOT2-PENDIENTES"
FILE_EMAIL_PROCESADOS = "BOT2-PROCESADOS"
FILE_DOWNLOAD_ATTACHMENT = "inboxFacturas"
SERVER = "imap.indra.es"
PORT_EMAIL = 993
LOG_FILE = "envio_fe_log.txt"  # Nombre del archivo de log para "Envío FE"
DEBUG_HTML_FILE = "debug_html_original.html"  # Para guardar HTML original para depuración
DIRECTORY_ONEDRIVE = "/Users/adrianvela/Library/CloudStorage/OneDrive-Indra/Facturas/Carpeta Archivos Adjuntos/BOT3 Estructura de Carpetas"

adjuntos = []
imagenes = []
todos_los_textos = []
directorio_descarga = ""
envio_fe_encontrado = False  # Variable global para rastrear si encontramos "Envío FE"

log_manager = LogManager()

def decode_mime_header(header):
    if not header:
        return "Desconocido"
    try:
        decoded_header = decode_header(header)
        decoded_string = ''.join([str(part[0], part[1] or 'utf-8') if isinstance(part[0], bytes) else part[0] for part in decoded_header])
        return decoded_string
    except Exception as e:
        print(f"Error al decodificar encabezado: {str(e)}")
        return str(header)

def limpiar_nombre_carpeta(nombre):
    if not nombre:
        return "sin_asunto"
    try:
        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "_", nombre)
        nombre_limpio = nombre_limpio.strip()[:100]  # Limitar a 100 caracteres
        # Si está vacío, usar sin_asunto
        if not nombre_limpio:
            nombre_limpio = "sin_asunto"
        return nombre_limpio
    except Exception:
        return "sin_asunto"

def html_a_texto(contenido_html): 
    if not contenido_html:
        return ""
    try:
        texto = html.unescape(contenido_html)
        texto = re.sub(r'<br[^>]*>', '\n', texto, flags=re.IGNORECASE)
        texto = re.sub(r'<p[^>]*>', '\n\n', texto, flags=re.IGNORECASE)
        texto = re.sub(r'</p>', '', texto, flags=re.IGNORECASE)
        texto = re.sub(r'<div[^>]*>', '\n', texto, flags=re.IGNORECASE)
        texto = re.sub(r'</div>', '', texto, flags=re.IGNORECASE)
        texto = re.sub(r'<tr[^>]*>', '\n', texto, flags=re.IGNORECASE)
        texto = re.sub(r'<td[^>]*>', '\t', texto, flags=re.IGNORECASE)
        texto = re.sub(r'<li[^>]*>', '\n- ', texto, flags=re.IGNORECASE)
        texto = re.sub(r'<[^>]+>', '', texto)
        texto = re.sub(r'\n\s*\n', '\n\n', texto)
        texto = re.sub(r'[ \t]+', ' ', texto)
        return texto.strip()
    except Exception as e:
        print(f"Error al convertir HTML a texto: {str(e)}")
        return ""

# Función para guardar el HTML original para depuración
def guardar_html_debug(contenido_html, directorio, nombre="debug_html_original.html"):
    try:
        ruta_debug = os.path.join(directorio, nombre)
        with open(ruta_debug, 'w', encoding='utf-8') as f:
            f.write(contenido_html)
        print(f"HTML guardado en: {ruta_debug}")
        return True
    except Exception as e:
        print(f"Error al guardar HTML: {str(e)}")
        return False

def detectar_envio_fe(texto):
    """
    Detecta si el texto contiene la palabra "Envío FE" (insensible a mayúsculas/minúsculas)
    """
    if not texto:
        return False
    # Búsqueda insensible a mayúsculas/minúsculas usando re.IGNORECASE
    return bool(re.search(r'envío\s+fe', texto, re.IGNORECASE))

def registrar_envio_fe(email_message, asunto, texto_encontrado, contexto):
    """
    Crea una entrada en el archivo de log cuando se detecta "Envío FE"
    """
    try:
        # Determinar la ruta del archivo de log
        log_path = os.path.join(os.getcwd(), LOG_FILE)
        
        # Obtener información relevante del correo
        fecha = email_message.get('Date', 'Desconocida')
        remitente = email_message.get('From', 'Desconocido')
        
        # Obtener la fecha y hora actual para el registro
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Extraer un fragmento de texto que contiene "Envío FE" (hasta 100 caracteres antes y después)
        match = re.search(r'.{0,100}envío\s+fe.{0,100}', texto_encontrado, re.IGNORECASE | re.DOTALL)
        fragmento = match.group(0) if match else "No se pudo extraer el fragmento"
        
        # Crear el mensaje de log
        log_message = (
            f"=== DETECCIÓN 'ENVÍO FE' - {timestamp} ===\n"
            f"Fecha del correo: {fecha}\n"
            f"De: {remitente}\n"
            f"Asunto: {asunto}\n"
            f"Encontrado en: {contexto}\n"
            f"Fragmento: {fragmento.strip()}\n"
            f"Ruta: {directorio_descarga}\n"
            f"{'=' * 50}\n\n"
        )
        
        # Añadir al archivo de log (modo append)
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(log_message)
            
        print(f"\n🔍 ¡DETECCIÓN! 'Envío FE' encontrado en correo: {asunto}")
        print(f"✓ Registrado en el log: {log_path}\n")
        
        return True
    except Exception as e:
        print(f"Error al registrar 'Envío FE' en el log: {str(e)}")
        return False

def extraer_tablas_html(contenido_html, directorio=None, email_message=None, asunto=None):
    """
    Extrae todas las tablas del contenido HTML con múltiples métodos para asegurar
    que se capturan incluso tablas con estructuras complejas.
    También detecta si alguna tabla contiene "Envío FE" y registra esta información.
    """
    global envio_fe_encontrado
    
    if not contenido_html:
        return []
    
    # Guardar copia del HTML original para depuración si se proporciona un directorio
    if directorio:
        guardar_html_debug(contenido_html, directorio)
    
    try:
        soup = BeautifulSoup(contenido_html, 'html.parser')
        tablas = []
        
        print(f"Analizando HTML para extracción de tablas (longitud: {len(contenido_html)} caracteres)")
        
        # VERIFICAR SI EL HTML COMPLETO CONTIENE "ENVÍO FE"
        if detectar_envio_fe(contenido_html) and not envio_fe_encontrado:
            print("\n🔍 ¡VERIFICACIÓN GLOBAL! El HTML contiene 'Envío FE'")
            if email_message and asunto:
                texto_convertido = html_a_texto(contenido_html)
                registrar_envio_fe(email_message, asunto, texto_convertido, "HTML COMPLETO")
                envio_fe_encontrado = True
        
        # BUSCAR ESPECÍFICAMENTE ELEMENTOS CON "ENVÍO FE"
        print("\nBuscando elementos específicos con 'Envío FE'...")
        elementos_envio_fe = soup.find_all(string=re.compile(r'Envío\s+FE|envío\s+fe|ENVÍO\s+FE', re.IGNORECASE))
        
        if elementos_envio_fe:
            print(f"🔍 ¡ENCONTRADO! Se hallaron {len(elementos_envio_fe)} elementos con 'Envío FE'")
            
            # Para cada elemento "Envío FE", buscar tablas cercanas
            for i, elemento in enumerate(elementos_envio_fe):
                elemento_texto = elemento.strip()
                print(f"  - Elemento #{i+1}: '{elemento_texto}'")
                
                # Buscar la primera tabla después de este elemento
                elemento_padre = elemento.parent
                tabla_siguiente = None
                
                # Método 1: Buscar en hermanos posteriores
                print("    Buscando tablas en hermanos posteriores...")
                for hermano in elemento_padre.next_siblings:
                    if hermano and hermano.name == 'table':
                        tabla_siguiente = hermano
                        print(f"    ✓ Tabla encontrada como hermano directo!")
                        break
                
                # Método 2: Si no encontramos, ir subiendo al padre y buscando en sus hermanos
                if not tabla_siguiente:
                    print("    Buscando tablas en ancestros y sus hermanos...")
                    ancestro = elemento_padre
                    for nivel in range(5):  # Subir hasta 5 niveles de padres
                        if not ancestro or not ancestro.parent:
                            break
                        ancestro = ancestro.parent
                        print(f"    - Subiendo al nivel {nivel+1} ({ancestro.name if ancestro.name else 'sin nombre'})")
                        
                        # Buscar en hermanos de este ancestro
                        for hermano in ancestro.next_siblings:
                            if hermano and hermano.name == 'table':
                                tabla_siguiente = hermano
                                print(f"    ✓ Tabla encontrada en hermano de ancestro nivel {nivel+1}!")
                                break
                        
                        # También buscar tablas dentro del ancestro
                        if not tabla_siguiente and ancestro:
                            tablas_en_ancestro = ancestro.find_all('table')
                            if tablas_en_ancestro:
                                print(f"    Encontradas {len(tablas_en_ancestro)} tablas dentro del ancestro nivel {nivel+1}")
                                tabla_siguiente = tablas_en_ancestro[0]  # Tomar la primera
                                print(f"    ✓ Seleccionada primera tabla dentro del ancestro!")
                                break
                        
                        if tabla_siguiente:
                            break
                
                # Si encontramos una tabla, procesarla
                if tabla_siguiente:
                    print(f"    ✓ Se procesará la tabla encontrada después de 'Envío FE'")
                    tabla_datos = []
                    
                    # Guardar la tabla completa en un archivo HTML separado para depuración
                    if directorio:
                        try:
                            ruta_tabla_html = os.path.join(directorio, f"tabla_envio_fe_{i+1}.html")
                            with open(ruta_tabla_html, 'w', encoding='utf-8') as f:
                                f.write(str(tabla_siguiente))
                            print(f"    ✓ Tabla HTML guardada en: {ruta_tabla_html}")
                        except Exception as e:
                            print(f"    ✗ Error al guardar tabla HTML: {str(e)}")
                    
                    # Procesar filas
                    filas = tabla_siguiente.find_all('tr')
                    print(f"    - La tabla tiene {len(filas)} filas")
                    
                    for fila in filas:
                        celdas = []
                        celdas_elementos = fila.find_all(['th', 'td'])
                        
                        for celda in celdas_elementos:
                            texto_celda = celda.get_text(strip=True)
                            celdas.append(texto_celda)
                            
                            # Verificar si esta celda contiene "Envío FE"
                            if detectar_envio_fe(texto_celda) and not envio_fe_encontrado and email_message and asunto:
                                print(f"    🔍 ¡DETECCIÓN! 'Envío FE' encontrado en celda: '{texto_celda}'")
                                registrar_envio_fe(email_message, asunto, texto_celda, f"CELDA DE TABLA (después de '{elemento_texto}')")
                                envio_fe_encontrado = True
                        
                        if any(celda for celda in celdas if celda.strip()):
                            tabla_datos.append(celdas)
                            
                            # Verificar si esta fila contiene "Envío FE"
                            fila_texto = " ".join(celdas)
                            if detectar_envio_fe(fila_texto) and not envio_fe_encontrado and email_message and asunto:
                                print(f"    🔍 ¡DETECCIÓN! 'Envío FE' encontrado en fila de tabla")
                                registrar_envio_fe(email_message, asunto, fila_texto, f"FILA DE TABLA (después de '{elemento_texto}')")
                                envio_fe_encontrado = True
                    
                    if tabla_datos:
                        print(f"    ✓ Tabla después de 'Envío FE' añadida con {len(tabla_datos)} filas")
                        tablas.append(tabla_datos)
                else:
                    print("    ✗ No se encontró ninguna tabla después de 'Envío FE'")
        
        # MÉTODO ESTÁNDAR: Buscar todas las tablas HTML
        print("\nBuscando todas las tablas en el documento...")
        tablas_html = soup.find_all('table')
        print(f"  - Encontradas {len(tablas_html)} tablas HTML estándar")
        
        # Procesar tablas encontradas (las que no hayan sido procesadas antes)
        for idx, tabla_html in enumerate(tablas_html):
            # Verificar si esta tabla ya fue procesada anteriormente
            ya_procesada = False
            for tabla_anterior in tablas:
                # Convertir ambas tablas a texto para comparación
                tabla_html_texto = "".join(celda.get_text() for fila in tabla_html.find_all('tr') for celda in fila.find_all(['th', 'td']))
                tabla_anterior_texto = "".join(" ".join(celda for celda in fila) for fila in tabla_anterior)
                
                if tabla_html_texto == tabla_anterior_texto:
                    ya_procesada = True
                    break
            
            if ya_procesada:
                print(f"  - Tabla #{idx+1} ya fue procesada anteriormente, omitiendo...")
                continue
                
            tabla_datos = []
            filas = tabla_html.find_all('tr')
            
            print(f"  - Procesando tabla #{idx+1} con {len(filas)} filas")
            
            # Verificar si la tabla contiene "Envío FE"
            tabla_texto = tabla_html.get_text(strip=True)
            if detectar_envio_fe(tabla_texto) and not envio_fe_encontrado and email_message and asunto:
                print(f"  🔍 ¡DETECCIÓN! 'Envío FE' encontrado en tabla #{idx+1}")
                registrar_envio_fe(email_message, asunto, tabla_texto, f"TABLA #{idx+1}")
                envio_fe_encontrado = True
                
                # Guardar esta tabla específica con "Envío FE" en un archivo HTML
                if directorio:
                    try:
                        ruta_tabla_html = os.path.join(directorio, f"tabla_con_envio_fe_{idx+1}.html")
                        with open(ruta_tabla_html, 'w', encoding='utf-8') as f:
                            f.write(str(tabla_html))
                        print(f"  ✓ Tabla con 'Envío FE' guardada en: {ruta_tabla_html}")
                    except Exception as e:
                        print(f"  ✗ Error al guardar tabla con 'Envío FE': {str(e)}")
            
            for fila in filas:
                celdas = []
                for celda in fila.find_all(['th', 'td']):
                    texto_celda = celda.get_text(strip=True)
                    celdas.append(texto_celda)
                    
                    # Verificar si esta celda contiene "Envío FE"
                    if detectar_envio_fe(texto_celda) and not envio_fe_encontrado and email_message and asunto:
                        print(f"  🔍 ¡DETECCIÓN! 'Envío FE' encontrado en celda de tabla #{idx+1}: '{texto_celda}'")
                        registrar_envio_fe(email_message, asunto, texto_celda, f"CELDA DE TABLA #{idx+1}")
                        envio_fe_encontrado = True
                
                if any(celda for celda in celdas if celda.strip()):
                    tabla_datos.append(celdas)
                    
                    # Verificar si esta fila contiene "Envío FE"
                    fila_texto = " ".join(celdas)
                    if detectar_envio_fe(fila_texto) and not envio_fe_encontrado and email_message and asunto:
                        print(f"  🔍 ¡DETECCIÓN! 'Envío FE' encontrado en fila de tabla #{idx+1}")
                        registrar_envio_fe(email_message, asunto, fila_texto, f"FILA DE TABLA #{idx+1}")
                        envio_fe_encontrado = True
            
            if tabla_datos:
                print(f"  ✓ Tabla #{idx+1} procesada con {len(tabla_datos)} filas")
                tablas.append(tabla_datos)
        
        # MÉTODO PARA ESTRUCTURAS DIV QUE SIMULAN TABLAS
        print("\nBuscando estructuras div que simulan tablas...")
        
        # Buscar divs con clases que sugieren tabla
        div_tablas = soup.find_all('div', class_=lambda c: c and any(term in (c.lower() or '') for term in ['table', 'grid', 'row', 'column', 'data-table']))
        print(f"  - Encontrados {len(div_tablas)} divs que podrían ser tablas")
        
        for idx, div_tabla in enumerate(div_tablas):
            # Verificar si este div contiene "Envío FE"
            div_texto = div_tabla.get_text(strip=True)
            if detectar_envio_fe(div_texto) and not envio_fe_encontrado and email_message and asunto:
                print(f"  🔍 ¡DETECCIÓN! 'Envío FE' encontrado en estructura div-tabla #{idx+1}")
                registrar_envio_fe(email_message, asunto, div_texto, f"ESTRUCTURA DIV-TABLA #{idx+1}")
                envio_fe_encontrado = True
                
                # Guardar este div con "Envío FE" en un archivo HTML
                if directorio:
                    try:
                        ruta_div_html = os.path.join(directorio, f"div_tabla_con_envio_fe_{idx+1}.html")
                        with open(ruta_div_html, 'w', encoding='utf-8') as f:
                            f.write(str(div_tabla))
                        print(f"  ✓ Div-tabla con 'Envío FE' guardado en: {ruta_div_html}")
                    except Exception as e:
                        print(f"  ✗ Error al guardar div-tabla con 'Envío FE': {str(e)}")
            
            # Buscar filas dentro del div (pueden ser divs con class="row" o similares)
            filas_divs = div_tabla.find_all('div', class_=lambda c: c and ('row' in (c.lower() or '')))
            
            if not filas_divs and len(div_tabla.find_all('div')) > 0:
                # Si no hay divs específicos para filas, tomar los divs hijos directos
                filas_divs = div_tabla.find_all('div', recursive=False)
            
            print(f"  - Estructura div #{idx+1} tiene {len(filas_divs)} posibles filas")
            
            if filas_divs:
                tabla_datos = []
                
                for fila_div in filas_divs:
                    # Buscar celdas (pueden ser divs con class="cell" o similares)
                    celdas_divs = fila_div.find_all('div', class_=lambda c: c and any(term in (c.lower() or '') for term in ['cell', 'col', 'column']))
                    
                    if not celdas_divs:
                        # Si no hay divs específicos para celdas, tomar los divs hijos directos
                        celdas_divs = fila_div.find_all('div', recursive=False)
                    
                    # Si aún no hay celdas, usar todos los elementos de texto
                    if not celdas_divs:
                        celdas_divs = fila_div.find_all(text=True, recursive=False)
                    
                    celdas = []
                    for celda_div in celdas_divs:
                        if hasattr(celda_div, 'get_text'):
                            texto_celda = celda_div.get_text(strip=True)
                        else:
                            texto_celda = str(celda_div).strip()
                        
                        if texto_celda:
                            celdas.append(texto_celda)
                            
                            # Verificar si esta celda contiene "Envío FE"
                            if detectar_envio_fe(texto_celda) and not envio_fe_encontrado and email_message and asunto:
                                print(f"  🔍 ¡DETECCIÓN! 'Envío FE' encontrado en celda de div-tabla #{idx+1}: '{texto_celda}'")
                                registrar_envio_fe(email_message, asunto, texto_celda, f"CELDA DE DIV-TABLA #{idx+1}")
                                envio_fe_encontrado = True
                    
                    if celdas:
                        tabla_datos.append(celdas)
                        
                        # Verificar si esta fila contiene "Envío FE"
                        fila_texto = " ".join(celdas)
                        if detectar_envio_fe(fila_texto) and not envio_fe_encontrado and email_message and asunto:
                            print(f"  🔍 ¡DETECCIÓN! 'Envío FE' encontrado en fila de div-tabla #{idx+1}")
                            registrar_envio_fe(email_message, asunto, fila_texto, f"FILA DE DIV-TABLA #{idx+1}")
                            envio_fe_encontrado = True
                
                if tabla_datos and len(tabla_datos) > 1:  # Al menos dos filas para considerarla tabla
                    print(f"  ✓ Estructura div #{idx+1} procesada como tabla con {len(tabla_datos)} filas")
                    tablas.append(tabla_datos)
        
        # Eliminar tablas duplicadas (misma estructura)
        tablas_unicas = []
        for tabla in tablas:
            # Convertir la tabla a una representación de string para comparación
            tabla_str = str(tabla)
            if not any(str(t) == tabla_str for t in tablas_unicas):
                tablas_unicas.append(tabla)
        
        print(f"\nResultado final: {len(tablas_unicas)} tablas únicas extraídas")
        
        if envio_fe_encontrado:
            print("✓ 'Envío FE' fue detectado y registrado durante el procesamiento de tablas")
        else:
            print("✗ No se encontró 'Envío FE' en ninguna tabla o estructura")
        
        return tablas_unicas
        
    except Exception as e:
        print(f"Error al extraer tablas HTML: {str(e)}")
        return []

def guardar_tablas_csv(tablas, directorio):
    """
    Guarda las tablas extraídas en archivos CSV con delimitador '|'.
    """
    archivos_guardados = []
    
    if not tablas:
        return archivos_guardados
    
    # Si solo hay una tabla, guardarla como tabla.csv
    if len(tablas) == 1:
        archivo_csv = os.path.join(directorio, "tabla_1.csv")
        try:
            with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter='|')  # Usando delimitador '|'
                for fila in tablas[0]:
                    writer.writerow(fila)
            archivos_guardados.append(archivo_csv)
            print(f"Tabla guardada en: {archivo_csv}")
        except Exception as e:
            print(f"Error al guardar tabla CSV: {str(e)}")
    else:
        # Si hay múltiples tablas, numerarlas
        for i, tabla in enumerate(tablas, 1):
            archivo_csv = os.path.join(directorio, f"tabla_{i}.csv")
            try:
                with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter='|')  # Usando delimitador '|'
                    for fila in tabla:
                        writer.writerow(fila)
                archivos_guardados.append(archivo_csv)
                print(f"Tabla {i} guardada en: {archivo_csv}")
            except Exception as e:
                print(f"Error al guardar tabla {i} CSV: {str(e)}")
    
    return archivos_guardados

def decodificar_filename(filename):
    if not filename:
        return None
        
    try:
        # Intentar decodificar si está codificado en base64 o quoted-printable
        if "=?" in filename and "?=" in filename:
            decoded_parts = decode_header(filename)
            filename_parts = []
            for content, encoding in decoded_parts:
                if isinstance(content, bytes):
                    if encoding:
                        content = content.decode(encoding)
                    else:
                        content = content.decode('utf-8', errors='replace')
                filename_parts.append(content)
            return ''.join(filename_parts)
        return filename
    except Exception:
        return filename

def conectar_correo(email_addr, password, server, port, carpeta):
    try:
        mail = imaplib.IMAP4_SSL(server, port)
        mail.login(email_addr, password)
        
        try:
            mail.select(carpeta)
        except Exception:
            try:
                # Probar con codificación UTF-8
                mail.select(carpeta.encode('utf-8'))
            except Exception:
                # Intentar con el formato que aparece en la imagen
                mail.select(f"{email_addr}/{carpeta}")
        
        return mail
    except Exception as e:
        print(f"Error al conectar al correo: {str(e)}")
        return None

def listar_asuntos_correos(mail):
    """
    Muestra una lista de todos los asuntos de correos en la bandeja actual
    con sus IDs.
    """
    print("\n" + "="*50)
    print("LISTA DE ASUNTOS DE CORREOS")
    print("="*50)
    
    try:
        # Obtener todos los IDs de correo
        status, data = mail.search(None, "ALL")
        mail_ids = data[0].split()
        
        if len(mail_ids) == 0:
            print("No se encontraron correos en la bandeja.")
            return
        
        print(f"Total de correos: {len(mail_ids)}")
        print(f"# | ID | Asunto")
        print("-" * 50)
        
        # Para cada ID, obtener asunto
        for i, email_id in enumerate(mail_ids):
            # Obtener encabezados para cada correo
            status, msg_data = mail.fetch(email_id, "(BODY[HEADER.FIELDS (SUBJECT)])")
            raw_headers = msg_data[0][1].decode('utf-8', errors='replace')
            
            # Extraer asunto
            subject_match = re.search(r'Subject: (.*)', raw_headers)
            asunto_raw = subject_match.group(1).strip() if subject_match else "Sin asunto"
            
            # Decodificar asunto si es necesario
            asunto = decode_mime_header(asunto_raw)
            
            
            print(f"{i+1} | {email_id.decode()} | {asunto}")
        
        print("="*50)
        
    except Exception as e:
        print(f"Error al listar asuntos: {str(e)}")


def procesar_mensaje(mail, msg_id=None, carpeta_destino=None):
    global adjuntos, imagenes, todos_los_textos, directorio_descarga, envio_fe_encontrado
    
    # Reiniciar variables
    adjuntos = []
    imagenes = []
    todos_los_textos = []
    tablas_encontradas = []
    envio_fe_encontrado = False
    
    # Definir directorio base de descarga
    directorio_base = carpeta_destino if carpeta_destino else os.path.join(os.getcwd(), FILE_DOWNLOAD_ATTACHMENT)
    os.makedirs(directorio_base, exist_ok=True)
    
    try:
        listar_asuntos_correos(mail)
        # Si no se especifica un ID, obtener el último correo
        if not msg_id:
            status, data = mail.search(None, "ALL")
            mail_ids = data[0].split()
            
            if len(mail_ids) == 0:
                log_manager.ejecucion_logger.warning("No se encontraron correos en la bandeja.")
                print("No se encontraron correos en la bandeja.")
                return False
                
            msg_id = mail_ids[-1]  # Último correo
            log_manager.ejecucion_logger.info(f"Procesando último correo (ID: {msg_id})")
            print(f"Procesando último correo (ID: {msg_id})")
        
        # Obtener el correo
        log_manager.ejecucion_logger.info(f"Obteniendo correo con ID: {msg_id}")
        print(f"Obteniendo correo con ID: {msg_id}")
        
        status, data = mail.fetch(msg_id, "(RFC822)")
        raw_email = data[0][1]
        email_message = email.message_from_bytes(raw_email)
        
        # Información básica del correo
        asunto = decode_mime_header(email_message.get('Subject', 'Sin asunto'))
        remitente = decode_mime_header(email_message.get('From', 'Desconocido'))
        fecha = email_message.get('Date', 'Desconocida')
        
        # Obtener logger específico para este correo
        correo_logger = log_manager.get_correo_logger(asunto)
        correo_logger.info(f"PROCESANDO CORREO - ID: {msg_id}")
        correo_logger.info(f"Asunto: {asunto}")
        correo_logger.info(f"De: {remitente}")
        correo_logger.info(f"Fecha: {fecha}")
        
        print(f"\nPROCESANDO CORREO:")
        print(f"Asunto: {asunto}")
        print(f"De: {remitente}")
        print(f"Fecha: {fecha}")

        os.environ["ASUNTO_CORREO"] = asunto
        
        # Definir directorio específico para este correo
        nombre_carpeta = limpiar_nombre_carpeta(asunto)
        directorio_descarga = os.path.join(directorio_base, nombre_carpeta).replace(".","")
        os.makedirs(directorio_descarga, exist_ok=True)
        
        log_manager.registrar_etapa(correo_logger, "creación de directorio", 
                                    f"Directorio creado: {directorio_descarga}")
        print(f"Creada carpeta para este correo: {directorio_descarga}")
        
        # Iniciar proceso de registro en el log de correos
        log_manager.registrar_etapa(correo_logger, "registro", 
                                   "Registrando correo en el log de seguimiento")
        print("\nRegistrando correo en el log de seguimiento...")
        
        # Almacenar todo el contenido HTML y texto plano
        html_partes = []
        texto_plano_partes = []
        
        # IMPORTANTE: Extraer TODAS las partes HTML y texto plano
        log_manager.registrar_etapa(correo_logger, "extracción de contenido", 
                                   "Extrayendo contenido del mensaje...")
        print("\nExtrayendo contenido del mensaje...")
        
        # Primera pasada para extraer solo texto y HTML (no adjuntos)
        for part in email_message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            
            # Omitir cualquier adjunto
            if "attachment" in content_disposition or part.get_filename():
                continue
            
            # Procesar contenido de texto y HTML
            if content_type == "text/plain" and "attachment" not in content_disposition:
                try:
                    charset = part.get_content_charset() or 'utf-8'
                    payload = part.get_payload(decode=True)
                    if payload:
                        texto = payload.decode(charset, errors='replace')
                        texto_plano_partes.append(texto)
                        todos_los_textos.append(("TEXTO PLANO", texto))
                        correo_logger.info("Parte de texto plano extraída")
                        print("✓ Parte de texto plano extraída")
                except Exception as e:
                    correo_logger.error(f"Error al extraer texto plano: {str(e)}")
                    correo_logger.exception("Detalle del error:")
                    print(f"Error al extraer texto plano: {str(e)}")
            
            elif content_type == "text/html" and "attachment" not in content_disposition:
                try:
                    charset = part.get_content_charset() or 'utf-8'
                    payload = part.get_payload(decode=True)
                    if payload:
                        html = payload.decode(charset, errors='replace')
                        html_partes.append(html)
                        texto_convertido = html_a_texto(html)
                        todos_los_textos.append(("HTML CONVERTIDO", texto_convertido))
                        correo_logger.info("Parte HTML extraída")
                        print("✓ Parte HTML extraída")
                except Exception as e:
                    correo_logger.error(f"Error al extraer HTML: {str(e)}")
                    correo_logger.exception("Detalle del error:")
                    print(f"Error al extraer HTML: {str(e)}")
        
        # Si no hay partes, podría ser un mensaje de una sola parte
        if not texto_plano_partes and not html_partes and email_message.get_content_type() in ["text/plain", "text/html"]:
            try:
                charset = email_message.get_content_charset() or 'utf-8'
                payload = email_message.get_payload(decode=True)
                if payload:
                    contenido = payload.decode(charset, errors='replace')
                    if email_message.get_content_type() == "text/plain":
                        texto_plano_partes.append(contenido)
                        todos_los_textos.append(("TEXTO PLANO", contenido))
                        correo_logger.info("Texto plano extraído del mensaje principal")
                        print("✓ Texto plano extraído del mensaje principal")
                    else:
                        html_partes.append(contenido)
                        texto_convertido = html_a_texto(contenido)
                        todos_los_textos.append(("HTML CONVERTIDO", texto_convertido))
                        correo_logger.info("HTML extraído del mensaje principal")
                        print("✓ HTML extraído del mensaje principal")
            except Exception as e:
                correo_logger.error(f"Error al extraer contenido del mensaje principal: {str(e)}")
                correo_logger.exception("Detalle del error:")
                print(f"Error al extraer contenido del mensaje principal: {str(e)}")
        
        # Obtener cuerpo del correo para el registro
        cuerpo_correo = obtener_cuerpo_correo(texto_plano_partes, html_partes)
        
        # Registrar correo en el log con información completa
        # Aquí creamos un diccionario con toda la info necesaria
        datos_correo = {
            "email_message": email_message,
            "asunto": asunto,
            "remitentes": remitente,
            "destinatarios": None,  # Se extraerá en la función
            "cuerpo": cuerpo_correo,
            "directorio_base": os.path.dirname(directorio_descarga)  # Directorio padre del directorio_descarga
        }

        # Registrar el correo en el log y marcar como "Completado" la primera etapa
        fila_excel = None  # Inicializar variable para la fila del Excel
        try:
            # Inicializar el archivo de registro si no existe
            inicializar_log_excel(datos_correo["directorio_base"])
            
            # Registrar este correo en el log y obtener la fila donde se registró
            fila_excel = registrar_correo_log(email_message, datos_correo["directorio_base"])
            
            # Marcar como completada la primera etapa: Lectura del correo
            actualizar_estado_log(asunto, "1.Lectura Correo", "Pendiente", datos_correo["directorio_base"])
            
            correo_logger.info(f"Correo registrado en log de Excel en la fila {fila_excel} y marcada etapa de lectura como completada")
            print(f"✓ Correo registrado en log en la fila {fila_excel} y marcada etapa de lectura como completada")
        except Exception as e:
            correo_logger.error(f"Error al registrar correo en log Excel: {str(e)}")
            correo_logger.exception("Detalle del error:")
            print(f"Error al registrar correo en log: {str(e)}")
        
        # Procesar cada parte HTML para buscar "Envío FE" y extraer tablas
        log_manager.registrar_etapa(correo_logger, "búsqueda de Envío FE",
                                  "Procesando partes HTML para buscar 'Envío FE' y tablas...")
        print("\nProcesando partes HTML para buscar 'Envío FE' y tablas...")
        todas_las_tablas = []
        
        for idx, html in enumerate(html_partes, 1):
            correo_logger.info(f"Analizando parte HTML #{idx}...")
            print(f"\nAnalizando parte HTML #{idx}...")
            
            # Verificar si contiene "Envío FE"
            if detectar_envio_fe(html) and not envio_fe_encontrado:
                correo_logger.warning(f"¡DETECCIÓN! 'Envío FE' encontrado en parte HTML #{idx}")
                print(f"🔍 ¡DETECCIÓN! 'Envío FE' encontrado en parte HTML #{idx}")
                texto_convertido = html_a_texto(html)
                registrar_envio_fe(email_message, asunto, texto_convertido, f"PARTE HTML #{idx}")
                envio_fe_encontrado = True
            
            # Extraer tablas de esta parte HTML
            try:
                tablas = extraer_tablas_html(html, directorio_descarga, email_message, asunto)
                if tablas:
                    todas_las_tablas.extend(tablas)
                    correo_logger.info(f"Se extrajeron {len(tablas)} tablas de la parte HTML #{idx}")
                    print(f"✓ Se extrajeron {len(tablas)} tablas de la parte HTML #{idx}")
                    
                    # Guardar esta parte HTML específica para referencia
                    try:
                        with open(os.path.join(directorio_descarga, f"parte_html_{idx}.html"), 'w', encoding='utf-8') as f:
                            f.write(html)
                        correo_logger.info(f"Parte HTML #{idx} guardada para referencia")
                        print(f"✓ Parte HTML #{idx} guardada para referencia")
                    except Exception as e:
                        correo_logger.error(f"Error al guardar parte HTML #{idx}: {str(e)}")
                        correo_logger.exception("Detalle del error:")
                        print(f"Error al guardar parte HTML #{idx}: {str(e)}")
            except Exception as e:
                correo_logger.error(f"Error al extraer tablas de la parte HTML #{idx}: {str(e)}")
                correo_logger.exception("Detalle del error:")
                print(f"Error al extraer tablas de la parte HTML #{idx}: {str(e)}")
        
        # También verificar el texto plano
        for idx, texto in enumerate(texto_plano_partes, 1):
            if detectar_envio_fe(texto) and not envio_fe_encontrado:
                correo_logger.warning(f"¡DETECCIÓN! 'Envío FE' encontrado en parte de texto plano #{idx}")
                print(f"🔍 ¡DETECCIÓN! 'Envío FE' encontrado en parte de texto plano #{idx}")
                registrar_envio_fe(email_message, asunto, texto, f"PARTE TEXTO PLANO #{idx}")
                envio_fe_encontrado = True
        
        # Guardar todas las tablas encontradas
        if todas_las_tablas:
            log_manager.registrar_etapa(correo_logger, "guardar tablas", 
                                      f"Guardando {len(todas_las_tablas)} tablas encontradas...")
            print(f"\nGuardando {len(todas_las_tablas)} tablas encontradas...")
            try:
                archivos_tablas = guardar_tablas_csv(todas_las_tablas, directorio_descarga)
                correo_logger.info(f"Se guardaron {len(archivos_tablas)} archivos CSV de tablas")
                print(f"✓ Se guardaron {len(archivos_tablas)} archivos CSV de tablas")
            except Exception as e:
                correo_logger.error(f"Error al guardar tablas CSV: {str(e)}")
                correo_logger.exception("Detalle del error:")
                print(f"Error al guardar tablas CSV: {str(e)}")
        else:
            correo_logger.warning("No se encontraron tablas en el contenido HTML")
            print("\n✗ No se encontraron tablas en el contenido HTML")
        
        # Guardar contenido completo en un archivo de texto
        log_manager.registrar_etapa(correo_logger, "guardar contenido", 
                                   "Guardando contenido completo del correo...")
        try:
            with open(os.path.join(directorio_descarga, "contenido_email.txt"), "w", encoding="utf-8") as f:
                f.write(f"Asunto: {asunto}\n")
                f.write(f"De: {remitente}\n")
                f.write(f"Fecha: {fecha}\n\n")
                f.write("="*60 + "\n")
                f.write("CONTENIDO DEL MENSAJE EXTRAÍDO:\n")
                f.write("="*60 + "\n\n")
                
                cuerpo_correo = ""
                for tipo, texto in todos_los_textos:
                    cuerpo_correo += f"--- {tipo} ---\n{texto}\n\n{'-' * 40}\n\n"
                    f.write(f"--- {tipo} ---\n")
                    f.write(texto)
                    f.write("\n\n" + "-" * 40 + "\n\n")
                
            correo_logger.info(f"Contenido del correo guardado en: {os.path.join(directorio_descarga, 'contenido_email.txt')}")
            print(f"✓ Contenido del correo guardado en: {os.path.join(directorio_descarga, 'contenido_email.txt')}")

            # Marcar etapa de organización como completada
            try:
                actualizar_estado_log(asunto, "Cuerpo del Correo", cuerpo_correo, datos_correo["directorio_base"])
                correo_logger.info("Etapa de organización marcada como completada en el log")
                print(f"✓ Etapa de organización marcada como completada en el log")
            except Exception as e:
                correo_logger.error(f"Error al actualizar estado de organización en log: {str(e)}")
                correo_logger.exception("Detalle del error:")
                print(f"Error al actualizar estado de organización en log: {str(e)}")
        except Exception as e:
            correo_logger.error(f"Error al guardar contenido_email.txt: {str(e)}")
            correo_logger.exception("Detalle del error:")
            print(f"Error al guardar contenido_email.txt: {str(e)}")
        
        # Segunda pasada para procesar adjuntos
        log_manager.registrar_etapa(correo_logger, "procesar adjuntos", 
                                   "Procesando archivos adjuntos...")
        print("\nProcesando archivos adjuntos...")
        for part in email_message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            
            # Procesamos solo adjuntos (no correos adjuntos)
            if "attachment" in content_disposition or part.get_filename():
                filename = part.get_filename()
                if filename:
                    # Decodificar nombre
                    filename = decodificar_filename(filename)
                    
                    # Verificar si es un tipo de archivo que queremos excluir (correos adjuntos)
                    if (filename.lower().endswith(('.msg', '.eml')) or 
                        "Elemento de Outlook" in content_type or 
                        "message/rfc822" in content_type or
                        content_type == "application/ms-tnef"):
                        correo_logger.info(f"Ignorando correo adjunto: {filename}")
                        print(f"➖ Ignorando correo adjunto: {filename}")
                        continue  # Saltar este adjunto y pasar al siguiente
                    
                    # Crear ruta
                    filepath = os.path.join(directorio_descarga, filename)
                    
                    # Guardar archivo
                    try:
                        with open(filepath, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        correo_logger.info(f"Archivo adjunto guardado: {filename}")
                        print(f"✓ Archivo adjunto guardado: {filename}")
                        
                        # Registrar archivo
                        if content_type.startswith('image/'):
                            imagenes.append((filename, filepath))
                        else:
                            adjuntos.append((filename, filepath))
                    except Exception as e:
                        correo_logger.error(f"Error al guardar adjunto {filename}: {str(e)}")
                        correo_logger.exception("Detalle del error:")
                        print(f"Error al guardar adjunto {filename}: {str(e)}")
        
        # Mostrar resumen final
        log_manager.registrar_etapa(correo_logger, "resumen", "RESUMEN DEL PROCESAMIENTO:")
        print("\nRESUMEN DEL PROCESAMIENTO:")
        
        if envio_fe_encontrado:
            correo_logger.warning("'Envío FE' fue detectado y registrado")
            print("✅ 'Envío FE' fue detectado y registrado")
        else:
            correo_logger.info("No se encontró 'Envío FE' en el contenido del correo")
            print("❌ No se encontró 'Envío FE' en el contenido del correo")
        
        if todas_las_tablas:
            correo_logger.info(f"Se extrajeron {len(todas_las_tablas)} tablas")
            print(f"✅ Se extrajeron {len(todas_las_tablas)} tablas")
        else:
            correo_logger.info("No se encontraron tablas para extraer")
            print("❌ No se encontraron tablas para extraer")
            
        if adjuntos:
            correo_logger.info(f"Se guardaron {len(adjuntos)} archivos adjuntos")
            print(f"✅ Se guardaron {len(adjuntos)} archivos adjuntos")
        
        if imagenes:
            correo_logger.info(f"Se guardaron {len(imagenes)} imágenes adjuntas")
            print(f"✅ Se guardaron {len(imagenes)} imágenes adjuntas")
        
        # Marcar la etapa final como completada
        log_manager.registrar_etapa(correo_logger, "actualización de estados", 
                                  "Actualizando estados en log Excel...")
        try:
            # Aquí iría el código para enviar a OneDrive, pero para nuestro ejemplo solo actualizamos el estado
            actualizar_estado_log(asunto, "1.Lectura Correo", "Completo", datos_correo["directorio_base"])
            actualizar_estado_log(asunto, "2.Descompresion", "Pendiente", datos_correo["directorio_base"])
            actualizar_estado_log(asunto, "3.Lectura XML", "Pendiente", datos_correo["directorio_base"])
            actualizar_estado_log(asunto, "4.Organizacion", "Pendiente", datos_correo["directorio_base"])
            actualizar_estado_log(asunto, "5.EnvioOneDrive", "Pendiente", datos_correo["directorio_base"])

            correo_logger.info("Etapa de envío a OneDrive marcada como 'Pendiente' en el log")
            print(f"✓ Etapa de envío a OneDrive marcada como 'Pendiente' en el log")
        except Exception as e:
            correo_logger.error(f"Error al actualizar estado de envío a OneDrive en log: {str(e)}")
            correo_logger.exception("Detalle del error:")
            print(f"Error al actualizar estado de envío a OneDrive en log: {str(e)}")
        
        # Finalizar el log con un resumen
        resumen = {
            "Envío FE": "Detectado" if envio_fe_encontrado else "No detectado",
            "Tablas": f"{len(todas_las_tablas)} extraídas" if todas_las_tablas else "Ninguna",
            "Adjuntos": f"{len(adjuntos)} archivos" if adjuntos else "Ninguno",
            "Imágenes": f"{len(imagenes)} imágenes" if imagenes else "Ninguna",
            "Directorio": directorio_descarga
        }
        
        log_manager.finalizar_correo(correo_logger, "PROCESAMIENTO COMPLETO", resumen)
        
        # # AQUÍ ES DONDE SE AÑADE EL CÓDIGO PARA LLAMAR A LOS DEMÁS SCRIPTS
        # # ================================================================
        # # Ejecutar el resto de scripts en orden para este correo
        # if directorio_descarga:
        #     print("\n====== CONTINUANDO CON EL FLUJO DE PROCESAMIENTO ======")
        #     correo_logger.info("Continuando con el flujo de procesamiento para los siguientes pasos")
            
        #     # Mostrar información de la fila del Excel
        #     if fila_excel:
        #         print(f">>> Correo registrado en la fila {fila_excel} del Excel")
        #         correo_logger.info(f"Correo registrado en la fila {fila_excel} del Excel")
        #     else:
        #         print(">>> ADVERTENCIA: No se pudo determinar la fila del Excel para este correo")
        #         correo_logger.warning("No se pudo determinar la fila del Excel para este correo")
            
        #     # Lista de scripts a ejecutar en orden
        #     scripts_siguientes = [
        #         "2.descom_zip.py",
        #         "3.readXML.py",
        #         "4.org_directorios.py",
        #         "5.SendOnedrive.py",
        #         #"SendRegistroHistorico.py"
        #     ]
            
        #     # Ejecutar cada script en secuencia
        #     for script in scripts_siguientes:
        #         correo_logger.info(f"Iniciando ejecución de {script}")
        #         exito = False;
        #         if("5.SendOnedrive.py" == script):
        #             exito = ejecutar_script_siguiente(script, directorio_descarga, DIRECTORY_ONEDRIVE)
        #         else:
        #             exito = ejecutar_script_siguiente(script, directorio_descarga, fila_excel)
                
        #         # Si hubo error, pasar al siguiente correo
        #         if not exito:
        #             correo_logger.error(f"Error al ejecutar {script}, pasando al siguiente correo")
        #             print(f"\n❌ Error al ejecutar {script}, pasando al siguiente correo")
        #             break
                    
        #         correo_logger.info(f"Completada ejecución de {script}")
                
        #         # Esperar un poco entre scripts
        #         print(f">>> Esperando 2 segundos antes del siguiente script...")
        #         time.sleep(2)
                
        #     if mover_correo_a_carpeta(mail, msg_id, FILE_EMAIL_PENDIENTE, FILE_EMAIL_PROCESADOS):
        #         correo_logger.info(f"Correo movido exitosamente a '{FILE_EMAIL_PROCESADOS}'")
        #         print(f"✅ Correo movido exitosamente a '{FILE_EMAIL_PROCESADOS}'")
        #     else:
        #         correo_logger.error(f"Error al mover el correo a '{FILE_EMAIL_PROCESADOS}'")
        #         print(f"❌ Error al mover el correo a '{FILE_EMAIL_PROCESADOS}'")

            
        #     correo_logger.info("Finalizado el flujo completo de procesamiento para este correo")
        #     print("\n✅ FINALIZADO EL FLUJO COMPLETO DE PROCESAMIENTO PARA ESTE CORREO")
        
        return True
        
    except Exception as e:
        if 'correo_logger' in locals():
            error_msg = f"Error al procesar el correo: {str(e)}"
            correo_logger.error()
            correo_logger.exception("Detalle del error:")
            resumen_error = {
                "Error": str(e),
                "Etapa": "desconocida",
                "Directorio": directorio_descarga if 'directorio_descarga' in locals() else "No creado"
            }
            log_manager.finalizar_correo(correo_logger, "ERROR", resumen_error)
            
        else:
            # Si falló antes de crear el logger
            error_msg = f"Error al procesar el correo: {str(e)}"
            log_manager.ejecucion_logger.error(error_msg)
            log_manager.ejecucion_logger.exception("Detalle del error:")
            
        print(f"Error al procesar el correo: {str(e)}")
        return False

def mover_correo_a_carpeta(mail, msg_id, carpeta_origen, carpeta_destino):
    """
    Mueve un correo de una carpeta a otra usando la secuencia COPY + DELETE
    
    Args:
        mail: Objeto de conexión IMAP
        msg_id: ID del mensaje a mover
        carpeta_origen: Carpeta de origen (actual)
        carpeta_destino: Carpeta de destino
    
    Returns:
        bool: True si se movió correctamente, False en caso contrario
    """
    try:
        print(f"Intentando mover correo ID {msg_id} de '{carpeta_origen}' a '{carpeta_destino}'...")
        
        # Asegurarse de que estamos en la carpeta correcta
        mail.select(carpeta_origen)
        
        # Intentar con MOVE (IMAP4rev1)
        try:
            typ, data = mail.uid('MOVE', msg_id, carpeta_destino)
            if typ == 'OK':
                print(f"✅ Correo movido exitosamente usando MOVE")
                return True
        except Exception as e:
            print(f"Comando MOVE no disponible: {e}. Intentando con COPY+DELETE...")
        
        # Método tradicional: COPY + DELETE
        # Paso 1: Copiar el mensaje
        typ, data = mail.copy(msg_id, carpeta_destino)
        if typ != 'OK':
            print(f"❌ Error al copiar el correo: {typ} {data}")
            return False
            
        print("✓ Correo copiado exitosamente")
        
        # Paso 2: Marcar el mensaje original como eliminado
        typ, data = mail.store(msg_id, '+FLAGS', '\\Deleted')
        if typ != 'OK':
            print(f"❌ Error al marcar el correo como eliminado: {typ} {data}")
            return False
            
        print("✓ Correo marcado para eliminación")
        
        # Paso 3: Expurgar para eliminar físicamente
        typ, data = mail.expunge()
        if typ != 'OK':
            print(f"❌ Error al expurgar: {typ} {data}")
            return False
            
        print("✓ Expurge completado")
        
        print(f"✅ Correo movido exitosamente usando COPY+DELETE")
        return True
        
    except Exception as e:
        print(f"❌ Error al mover correo: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
         
def procesar_email_completo(email_addr=EMAIL, password=PASSWORD, server=SERVER, port=PORT_EMAIL, carpeta=FILE_EMAIL_PENDIENTE, carpeta_destino=None, msg_id=None, procesar_todos=False):
    """
    Procesa uno o múltiples emails y retorna el directorio donde se descargó el contenido
    """
    global directorio_descarga
    
    # Obtener logger global
    global_logger = log_manager.global_logger
    global_logger.info(f"INICIANDO PROCESAMIENTO DE CORREO(S) - Bandeja: {carpeta}")
    global_logger.info(f"Procesamiento múltiple: {'Sí' if procesar_todos else 'No'}")
    
    print("\n" + "="*50)
    print("INICIANDO PROCESAMIENTO DE CORREO(S)")
    print("="*50)
    
    # Inicializar directorio_descarga como None por si falla la conexión
    directorio_descarga = None
    
    try:
        mail = conectar_correo(email_addr, password, server, port, carpeta)
        if not mail:
            global_logger.error("No se pudo conectar al servidor de correo.")
            print("✗ No se pudo conectar al servidor de correo.")
            return False, None
        
        global_logger.info("Conexión establecida al servidor de correo")
        print("✓ Conexión establecida al servidor de correo")
        
        # Nuevo: Si no se especifica un ID concreto y se quiere procesar todos
        if not msg_id and procesar_todos:
            status, data = mail.search(None, "ALL")
            mail_ids = data[0].split()
            
            if len(mail_ids) == 0:
                global_logger.warning("No se encontraron correos en la bandeja.")
                print("No se encontraron correos en la bandeja.")
                return False, None
                
            global_logger.info(f"Se procesarán {len(mail_ids)} correos.")
            print(f"Se procesarán {len(mail_ids)} correos.")
            directorios_procesados = []
            
            for i, email_id in enumerate(mail_ids):
                global_logger.info(f"Procesando correo {i+1} de {len(mail_ids)} (ID: {email_id})")
                print(f"\n{'='*50}\nProcesando correo {i+1} de {len(mail_ids)} (ID: {email_id})\n{'='*50}")
                resultado = procesar_mensaje(mail, email_id, carpeta_destino)
                if resultado and directorio_descarga:
                    directorios_procesados.append(directorio_descarga)
            
            # Cerrar conexión
            try:
                mail.logout()
                global_logger.info("Conexión cerrada correctamente")
                print("✓ Conexión cerrada correctamente")
            except Exception as e:
                global_logger.error(f"Error al cerrar la conexión: {str(e)}")
                global_logger.exception("Detalle del error:")
                print("✗ Error al cerrar la conexión")
            
            global_logger.info(f"FIN DEL PROCESAMIENTO MÚLTIPLE - Total correos procesados: {len(directorios_procesados)}")
            print("="*50)
            print("FIN DEL PROCESAMIENTO MÚLTIPLE")
            print(f"Total de correos procesados: {len(directorios_procesados)}")
            print("="*50)
            
            return True, directorios_procesados
        else:
            # Comportamiento original (procesar un solo correo)
            resultado = procesar_mensaje(mail, msg_id, carpeta_destino)
            
            # Cerrar conexión
            try:
                mail.logout()
                global_logger.info("Conexión cerrada correctamente")
                print("✓ Conexión cerrada correctamente")
            except Exception as e:
                global_logger.error(f"Error al cerrar la conexión: {str(e)}")
                global_logger.exception("Detalle del error:")
                print("✗ Error al cerrar la conexión")
            
            global_logger.info(f"FIN DEL PROCESAMIENTO - Directorio: {directorio_descarga}")
            print("="*50)
            print("FIN DEL PROCESAMIENTO")
            print(f"Directorio de descarga: {directorio_descarga}")
            print("="*50)
            
            return resultado, directorio_descarga
            
    except Exception as e:
        global_logger.error(f"ERROR GRAVE en el procesamiento: {type(e).__name__}: {str(e)}")
        global_logger.exception("Detalle del error crítico:")
        print(f"ERROR GRAVE: {type(e).__name__}: {str(e)}")
        return False, None
    

def ejecutar_script_siguiente(script_name, directorio, fila_excel=None):
    """
    Ejecuta el siguiente script en la cadena de procesamiento.
    
    Args:
        script_name: Nombre del script a ejecutar
        directorio: Directorio donde se encuentran los archivos descargados
        fila_excel: Número de fila en el Excel registro_correos.xlsx
    """
    try:
        print(f"\n>>> EJECUTANDO SIGUIENTE PASO: {script_name}")
        
        # Obtener la ruta completa al script
        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_name)
        
        # Preparar el comando con los argumentos
        comando = ['python', script_path, directorio]
        
        # Si se proporciona la fila del Excel, añadirla como parámetro adicional
        if fila_excel is not None:
            comando.append(str(fila_excel))
            
        # Mostrar el comando completo
        cmd_str = " ".join(comando)
        print(f">>> Ejecutando: {cmd_str}")
        
        # Ejecutar el script
        proceso = subprocess.run(comando, capture_output=True, text=True)
        
        # Mostrar la salida
        print(proceso.stdout)
        
        # Verificar si hubo error
        if proceso.returncode != 0:
            print(f"ERROR en {script_name}:")
            print(proceso.stderr)
            return False
            
        print(f">>> COMPLETADO: {script_name}")
        return True
    except Exception as e:
        print(f"EXCEPCIÓN al ejecutar {script_name}: {str(e)}")
        return False
    
if __name__ == "__main__":
    try:
        # Ejemplo de uso:
        print("="*50)
        print("PROCESADOR DE CORREOS - DETECTOR DE 'ENVÍO FE'")
        print("="*50)
        print("Se extraerán tablas del cuerpo principal del correo, ignorando correos adjuntos")
        print("="*50)
        
        # Para procesar múltiples correos, descomenta esta línea:
        # resultado, paths_descarga = procesar_email_completo(procesar_todos=True)
        
        # Para procesar un solo correo (comportamiento original):
        resultado, path_descarga = procesar_email_completo(procesar_todos=True)
        
        # Imprimir de forma destacada el directorio de descarga
        if resultado:
            if isinstance(path_descarga, list):
                # Caso de múltiples correos
                print("\n" + "="*50)
                print(f"📁 DIRECTORIOS DE DESCARGA:")
                for i, path in enumerate(path_descarga):
                    print(f"   {i+1}. {path}")
                print("="*50)
                
                # Este print es especial para poder capturar esta salida desde otro script
                print(f"OUTPUT_DIRECTORIES={','.join(path_descarga)}")
            else:
                # Caso de un solo correo
                print("\n" + "="*50)
                print(f"📁 DIRECTORIO DE DESCARGA: {path_descarga}")
                print("="*50)
                
                # Este print es especial para poder capturar esta salida desde otro script
                print(f"OUTPUT_DIRECTORY={path_descarga}")
        else:
            print("\n❌ No se pudo procesar el correo correctamente")
    except Exception as e:
        print(f"ERROR GRAVE: {type(e).__name__}: {str(e)}")