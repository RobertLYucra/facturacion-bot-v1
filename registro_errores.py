import os
import pandas as pd
from datetime import datetime

# Ruta del archivo de log detallado
LOG_DETALLADO = "log_detallado.xlsx"

def registrar_log_detallado(asunto, etapa, estado, descripcion, directorio_base=None):
    """
    Registra un log detallado con fecha, asunto, etapa, estado y descripción del error.
    
    Args:
        asunto (str): Asunto del correo electrónico procesado.
        etapa (str): Etapa del proceso donde ocurrió el evento (lectura, descompresión, etc.).
        estado (str): Estado del procesamiento ('Éxito' o 'Error').
        descripcion (str): Descripción detallada del evento o error.
        directorio_base (str, optional): Directorio base para el archivo de log.
    """
    directorio_base = directorio_base or os.path.dirname(os.path.abspath(__file__))
    ruta_log = os.path.join(directorio_base, LOG_DETALLADO)

    # Crear estructura del log si no existe
    columnas_log = ["Fecha y Hora", "Asunto correo", "Etapa del proceso", "Estado", "Descripción del error"]

    registro = {
        "Fecha y Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Asunto correo": asunto,
        "Etapa del proceso": etapa,
        "Estado": estado,
        "Descripción del error": descripcion
    }

    try:
        if os.path.exists(ruta_log):
            df_log = pd.read_excel(ruta_log)
        else:
            df_log = pd.DataFrame(columns=columnas_log)

        df_log = pd.concat([df_log, pd.DataFrame([registro])], ignore_index=True)
        df_log.to_excel(ruta_log, index=False)
        print(f"✅ Registrado en log detallado: {registro}")

    except Exception as e:
        print(f"❌ Error al registrar en log detallado: {e}")
