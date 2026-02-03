import os
import shutil
import sys
from registro_errores import registrar_log_detallado

# Parámetros por defecto
DEFAULT_ORIGEN = "inboxFacturas/RV_ Facturación Perú 19.03.2025/Organizado"
DEFAULT_DESTINO = r"C:\Users\Administrator\OneDrive - Indra (1)\Facturas\Carpeta Archivos Adjuntos\BOT3 Estructura de Carpetas"

# def copiar_solo_carpetas(origen, destino):
#     """Copia solo las carpetas (y su contenido) de una ubicación a otra"""
#     # Verificar que la carpeta origen existe
#     if not os.path.exists(origen):
#         print(f"Error: La carpeta origen '{origen}' no existe.")
#         return False
        
#     # Verificar que la carpeta origen es realmente una carpeta
#     if not os.path.isdir(origen):
#         print(f"Error: '{origen}' no es una carpeta.")
#         return False
        
#     # Crear la carpeta destino si no existe
#     if not os.path.exists(destino):
#         print(f"Creando carpeta destino '{destino}'...")
#         os.makedirs(destino)
        
#     # Obtener la lista de elementos en la carpeta origen
#     todos_elementos = os.listdir(origen)
    
#     # Filtrar para quedarnos solo con las carpetas
#     carpetas = [elemento for elemento in todos_elementos if os.path.isdir(os.path.join(origen, elemento))]
#     total_carpetas = len(carpetas)
    
#     print(f"Se encontraron {total_carpetas} carpetas para copiar.")
    
#     if total_carpetas == 0:
#         print("No hay carpetas para copiar.")
#         return True
    
#     # Copiar cada carpeta
#     for i, carpeta in enumerate(carpetas, 1):
#         ruta_origen = os.path.join(origen, carpeta)
#         ruta_destino = os.path.join(destino, carpeta)
        
#         try:
#             print(f"[{i}/{total_carpetas}] Copiando carpeta: {carpeta}...")
#             shutil.copytree(ruta_origen, ruta_destino, dirs_exist_ok=True)
#         except Exception as e:
#             print(f"  ❌ Error al copiar la carpeta '{carpeta}': {str(e)}")
    
#     print(f"\n✅ Copia finalizada. Se copiaron {total_carpetas} carpetas de '{origen}' a '{destino}'.")
#     return True

def copiar_solo_carpetas(origen, destino, asunto_correo):
    if not os.path.exists(origen):
        error_msg = f"La carpeta origen '{origen}' no existe."
        print(error_msg)
        registrar_log_detallado(asunto_correo, "4.Organizacion", "Error", error_msg)
        return False
        
    if not os.path.isdir(origen):
        error_msg = f"'{origen}' no es una carpeta."
        print(error_msg)
        registrar_log_detallado(asunto_correo, "4.Organizacion", "Error", error_msg)
        return False
        
    if not os.path.exists(destino):
        try:
            os.makedirs(destino)
            print(f"Creada carpeta destino '{destino}'.")
        except Exception as e:
            error_msg = f"No se pudo crear la carpeta destino '{destino}': {str(e)}"
            print(error_msg)
            registrar_log_detallado(asunto_correo, "4.Organizacion", "Error", error_msg)
            return False
    
    todos_elementos = os.listdir(origen)
    carpetas = [e for e in todos_elementos if os.path.isdir(os.path.join(origen, e))]
    
    if not carpetas:
        msg = "No hay carpetas para copiar."
        print(msg)
        registrar_log_detallado(asunto_correo, "4.Organizacion", "Éxito", msg)
        return True
    
    errores = []
    for carpeta in carpetas:
        ruta_origen = os.path.join(origen, carpeta)
        ruta_destino = os.path.join(destino, carpeta)
        
        try:
            shutil.copytree(ruta_origen, ruta_destino, dirs_exist_ok=True)
            print(f"Carpeta '{carpeta}' copiada exitosamente.")
        except Exception as e:
            error_msg = f"Error al copiar '{carpeta}': {str(e)}"
            print(error_msg)
            errores.append(error_msg)
    
    if errores:
        registrar_log_detallado(asunto_correo, "4.Organizacion", "Error", " ; ".join(errores))
        return False
    else:
        registrar_log_detallado(asunto_correo, "4.Organizacion", "Éxito", "Carpetas copiadas correctamente.")
        return True
    
def main():
    # Determinar origen según los argumentos
    if len(sys.argv) > 1:
        # Si hay al menos un argumento, usarlo como carpeta origen
        base_dir = sys.argv[1]
        origen = os.path.join(base_dir, "Organizado")
        print(f"Usando directorio proporcionado: {base_dir}")
        print(f"Carpeta origen calculada: {origen}")
    else:
        # Si no hay argumentos, usar el valor predeterminado
        origen = DEFAULT_ORIGEN
        print(f"No se proporcionó directorio. Usando origen predeterminado: {origen}")
    
    # Determinar destino (puede ser el segundo argumento o el valor predeterminado)
    if len(sys.argv) > 2:
        destino = sys.argv[2]
        print(f"Usando destino proporcionado: {destino}")
    else:
        destino = DEFAULT_DESTINO
        print(f"Usando destino predeterminado: {destino}")
    
    print(f"Copiando carpetas de '{origen}' a '{destino}'...")
    asunto_correo = os.environ.get("ASUNTO_CORREO", "Asunto desconocido")
    resultado = copiar_solo_carpetas(origen, destino,asunto_correo)
    
    # Imprimir un mensaje especial para el script automatizado
    if resultado:
        print(f"OUTPUT_DIRECTORY={origen}")
        # Salir con código de éxito
        sys.exit(0)
    else:
        # Salir con código de error
        sys.exit(1)

if __name__ == "__main__":
    main()