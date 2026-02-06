import os
import pandas as pd
import datetime
from email.header import decode_header

# Constantes para el registro de logs
LOG_EXCEL_FILE = "registro_correos.xlsx"
LOG_COLUMNS = [
    "Asunto", 
    "Fecha Procesamiento", 
    "Remitentes", 
    "Destinatarios", 
    "Cuerpo del Correo", 
    "1.Lectura Correo", 
    "2.Descompresion", 
    "3.Lectura XML", 
    "4.Organizacion", 
    "5.EnvioOneDrive"
]
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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

def inicializar_log_excel(directorio=None):
    directorio = SCRIPT_DIR
    
    ruta_excel = os.path.join(directorio, LOG_EXCEL_FILE)
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_excel):
        try:
            # Intentar leer el archivo existente para verificar su estructura
            df = pd.read_excel(ruta_excel)
            # Verificar que tiene las columnas necesarias
            columnas_faltantes = set(LOG_COLUMNS) - set(df.columns)
            if columnas_faltantes:
                # Si faltan columnas, añadirlas
                for col in columnas_faltantes:
                    df[col] = ""
                # Guardar con las nuevas columnas
                df.to_excel(ruta_excel, index=False)
                print(f"✓ Actualizadas las columnas faltantes en el archivo de registro: {ruta_excel}")
        except Exception as e:
            print(f"Error al leer el archivo de registro existente: {str(e)}")
            print("Creando un nuevo archivo de registro...")
            # Crear un DataFrame vacío con las columnas necesarias
            df = pd.DataFrame(columns=LOG_COLUMNS)
            df.to_excel(ruta_excel, index=False)
    else:
        # Crear un nuevo archivo Excel
        df = pd.DataFrame(columns=LOG_COLUMNS)
        df.to_excel(ruta_excel, index=False)
        print(f"✓ Creado nuevo archivo de registro: {ruta_excel}")
    
    return ruta_excel

def extraer_destinatarios(email_message):
    """
    Extrae los destinatarios de un mensaje de correo.
    
    Args:
        email_message: Objeto email.message que contiene el correo
        
    Returns:
        str: Cadena con todos los destinatarios concatenados
    """
    destinatarios = []
    
    # Extraer de campos To, CC y BCC
    for field in ['To', 'Cc', 'Bcc']:
        value = email_message.get(field)
        if value:
            try:
                decoded_value = decode_mime_header(value)
                destinatarios.append(f"{field}: {decoded_value}")
            except Exception as e:
                print(f"Error al decodificar campo {field}: {str(e)}")
                destinatarios.append(f"{field}: {value}")
    
    return "; ".join(destinatarios)

def obtener_cuerpo_correo(texto_plano_partes, html_partes):
    """
    Obtiene el cuerpo del correo a partir de las partes de texto o HTML.
    
    Args:
        texto_plano_partes (list): Lista con las partes de texto plano
        html_partes (list): Lista con las partes HTML
    
    Returns:
        str: Contenido del cuerpo del correo (preferiblemente texto plano)
    """
    # Preferir texto plano si está disponible
    if texto_plano_partes and len(texto_plano_partes) > 0:
        # Unir todas las partes de texto plano
        return "\n\n".join(texto_plano_partes)
    
    # Si no hay texto plano pero hay HTML, convertir HTML a texto
    elif html_partes and len(html_partes) > 0:
        from bs4 import BeautifulSoup
        texto = ""
        for html in html_partes:
            try:
                soup = BeautifulSoup(html, 'html.parser')
                texto += soup.get_text(separator="\n") + "\n\n"
            except Exception as e:
                print(f"Error al convertir HTML a texto para el log: {str(e)}")
                texto += "Error al procesar HTML"
        return texto
    
    # Si no hay nada, devolver mensaje de error
    return "No se pudo extraer el cuerpo del correo"

def registrar_correo_log(email_message, directorio_base=None):
    """
    Registra la información de un correo en el archivo Excel de logs.
    Siempre añade una nueva entrada, independientemente de si el correo ya existe.
    
    Args:
        email_message: Objeto email.message que contiene el correo
        directorio_base (str, optional): Directorio base donde se encuentra o 
                                        se creará el archivo Excel
    
    Returns:
        bool: True si se registró correctamente, False en caso contrario
    """
    try:
        # Inicializar o cargar el archivo Excel
        ruta_excel = inicializar_log_excel(directorio_base)
        
        # Leer el DataFrame existente
        df = pd.read_excel(ruta_excel)
        
        # Extraer información del correo
        asunto = decode_mime_header(email_message.get('Subject', 'Sin asunto'))
        remitentes = decode_mime_header(email_message.get('From', 'Desconocido'))
        destinatarios = extraer_destinatarios(email_message)
        fecha_procesamiento = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Extraer el cuerpo del correo
        cuerpo = "Este campo debe ser extraído del email_message o pasado como parámetro"
        
        # Intentar extraer el cuerpo directamente si es un mensaje simple
        if email_message.get_content_type() in ["text/plain", "text/html"]:
            try:
                charset = email_message.get_content_charset() or 'utf-8'
                payload = email_message.get_payload(decode=True)
                if payload:
                    contenido = payload.decode(charset, errors='replace')
                    cuerpo = contenido[:5000]  # Limitar a 5000 caracteres para no sobrecargar el Excel
            except Exception as e:
                print(f"Error al extraer cuerpo directo: {str(e)}")
        
        # Crear una nueva fila para el DataFrame
        nueva_fila = {
            "Asunto": asunto,
            "Fecha Procesamiento": fecha_procesamiento,
            "Remitentes": remitentes,
            "Destinatarios": destinatarios,
            "Cuerpo del Correo": cuerpo,
            "1.Lectura Correo": "",
            "2.Descompresion": "",
            "3.Lectura XML": "",
            "4.Organizacion": "",
            "5.EnvioOneDrive": ""
        }
        
        # Siempre añadir una nueva fila, sin verificar si ya existe
        df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
        print(f"✓ Registrado correo en el log: {asunto}")
        
        # Guardar el DataFrame actualizado
        df.to_excel(ruta_excel, index=False)
        print(f"✓ Archivo de registro actualizado: {ruta_excel}")
        
        return True
    
    except Exception as e:
        print(f"Error al registrar correo en el log: {str(e)}")
        return False

def actualizar_estado_log(asunto, columna, estado, directorio_base=None):
    """
    Actualiza el estado de una columna específica para un correo identificado por su asunto.
    Si hay múltiples correos con el mismo asunto, actualiza el más reciente.
    
    Args:
        asunto (str): Asunto del correo que identifica la entrada en el log
        columna (str): Nombre de la columna a actualizar (debe empezar con el número y punto)
        estado (str): Nuevo estado a establecer
        directorio_base (str, optional): Directorio base donde se encuentra el archivo Excel
    
    Returns:
        bool: True si se actualizó correctamente, False en caso contrario
    """
    try:
        # Verificar que la columna sea válida
        if not any(col == columna for col in LOG_COLUMNS):
            print(f"Error: Columna '{columna}' no válida. Debe ser una de: {', '.join([c for c in LOG_COLUMNS if c.startswith(tuple(['1.', '2.', '3.', '4.', '5.']))])}")
            return False
        
        # Inicializar o cargar el archivo Excel
        ruta_excel = inicializar_log_excel(directorio_base)
        
        # Leer el DataFrame existente
        df = pd.read_excel(ruta_excel)
        
        # Buscar el correo por asunto
        df_filtrado = df[df["Asunto"] == asunto]
        
        if len(df_filtrado) == 0:
            print(f"Error: No se encontró ningún correo con asunto '{asunto}' en el registro")
            return False
        
        # Si hay múltiples correos con el mismo asunto, actualizar el más reciente
        # Convertir la columna de fecha a datetime para poder ordenar
        df_filtrado['Fecha Procesamiento'] = pd.to_datetime(df_filtrado['Fecha Procesamiento'])
        # Ordenar por fecha en orden descendente
        df_filtrado = df_filtrado.sort_values('Fecha Procesamiento', ascending=False)
        # Obtener el índice del correo más reciente
        indice_correo = df_filtrado.index[0]
        
        # Actualizar el estado
        df.loc[indice_correo, columna] = estado
        
        # Guardar el DataFrame actualizado
        df.to_excel(ruta_excel, index=False)
        print(f"✓ Actualizado estado de '{columna}' a '{estado}' para correo: {asunto}")
        
        return True
    
    except Exception as e:
        print(f"Error al actualizar estado en el log: {str(e)}")
        return False