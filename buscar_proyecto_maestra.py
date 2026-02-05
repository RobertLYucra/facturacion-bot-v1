import pandas as pd
import sys

def cargar_excel(ruta_archivo):
    """
    Carga el archivo Excel y devuelve un DataFrame con sus datos.
    """
    try:
        # Leer la primera hoja del Excel
        df = pd.read_excel(ruta_archivo, sheet_name="Carpeta Organización Fact")
        
        # Renombrar las columnas duplicadas (hay dos columnas RUC)
        df.columns = ["Cliente", "RUC_Cliente", "Sociedad", "RUC_Sociedad", 
                     "Proyecto", "Tipo_Documento", "Fecha", "Destinatarios", 
                     "Documentos_Adjuntar"]
        
        # Limpiar los nombres de proyectos (eliminar espacios en blanco)
        if "Proyecto" in df.columns:
            df["Proyecto"] = df["Proyecto"].astype(str).str.strip()
        
        return df
    except Exception as e:
        print(f"Error al cargar el archivo Excel: {str(e)}")
        return None

def buscar_proyecto(codigo_proyecto, ruta_excel="Maestra/Robot 2_Estructura Carpetas Factura SSFF 03.03.2025.xlsx"):
    """
    Busca un proyecto específico en el DataFrame y devuelve su información.
    
    Parámetros:
    - codigo_proyecto: Código del proyecto a buscar
    - ruta_excel: Ruta del archivo Excel (por defecto usa la ruta predeterminada)
    
    Retorna:
    - DataFrame con los resultados de la búsqueda
    """
    # Imprimir información de búsqueda
    print(f"Buscando proyecto con código: {codigo_proyecto}")
    print(f"Usando archivo Excel: {ruta_excel}")
    
    # Cargar el Excel
    datos = cargar_excel(ruta_excel)
    
    if datos is None:
        print("ERROR: No se pudo cargar el archivo Excel")
        return None
    
    print(f"Excel cargado exitosamente. Total de registros: {len(datos)}")
    
    # Mostrar los primeros valores de la columna Proyecto para depuración
    proyectos_muestra = datos["Proyecto"].head(5).tolist()
    print(f"Muestra de proyectos en el Excel: {proyectos_muestra}")
    
    # Convertir código de búsqueda a mayúsculas para comparación
    codigo_upper = codigo_proyecto.upper()
    print(f"Buscando coincidencia exacta para: {codigo_upper}")
    
    # Buscar registros que contengan el código de proyecto (ignorando mayúsculas/minúsculas)
    resultados = datos[datos["Proyecto"].str.upper() == codigo_upper]
    
    print(f"Coincidencias exactas encontradas: {len(resultados)}")
    
    if len(resultados) == 0:
        # Si no hay coincidencias exactas, buscar coincidencias parciales
        print(f"No se encontraron coincidencias exactas. Buscando coincidencias parciales...")
        resultados = datos[datos["Proyecto"].str.upper().str.contains(codigo_upper)]
        print(f"Coincidencias parciales encontradas: {len(resultados)}")
    
    # Mostrar los primeros resultados para verificación
    if len(resultados) > 0:
        primeros_resultados = resultados.head(min(3, len(resultados)))
        print(f"Primeros resultados encontrados:")
        for i, proyecto in enumerate(primeros_resultados["Proyecto"].tolist()):
            print(f"  {i+1}. {proyecto}")
    else:
        print(f"ADVERTENCIA: No se encontraron coincidencias para el código: {codigo_proyecto}")
    
    return resultados

def main():
    """ 
    Función principal para ejecución directa del script.
    """
    if len(sys.argv) < 2:
        print("Uso: python buscar_proyecto.py <codigo_proyecto>")
        print("  o  python buscar_proyecto.py --interactivo")
        return
    
    # Modo interactivo
    if sys.argv[1] == "--interactivo":
        while True:
            codigo_proyecto = input("\nIngrese el código del proyecto a buscar (o 'salir' para terminar): ")
            if codigo_proyecto.lower() == 'salir':
                break
            
            resultados = buscar_proyecto(codigo_proyecto)
            
            if resultados is not None and len(resultados) > 0:
                for _, proyecto in resultados.iterrows():
                    print("\nResultado:")
                    print(f"  Cliente: {proyecto['Cliente']}")
                    print(f"  Proyecto: {proyecto['Proyecto']}")
                    print(f"  Sociedad: {proyecto['Sociedad']}")
            else:
                print("No se encontraron resultados.")
    else:
        # Modo directo: usar el primer argumento como código de proyecto
        codigo_proyecto = sys.argv[1]
        resultados = buscar_proyecto(codigo_proyecto)
        
        if resultados is not None and len(resultados) > 0:
            for _, proyecto in resultados.iterrows():
                print("\nResultado:")
                print(f"  Cliente: {proyecto['Cliente']}")
                print(f"  Proyecto: {proyecto['Proyecto']}")
                print(f"  Sociedad: {proyecto['Sociedad']}")
        else:
            print("No se encontraron resultados.")
            sys.exit(1)

if __name__ == "__main__":
    main()