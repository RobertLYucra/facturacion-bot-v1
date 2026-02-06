# log_manager.py
import os
import logging
from datetime import datetime
import re

class LogManager:
    """
    Administrador centralizado de logs para el procesamiento de correos.
    Crea y gestiona logs individuales para cada correo.
    """
    
    def __init__(self, ejecucion_id=None):
        """
        Inicializa el administrador de logs simplificado.
        
        Args:
            ejecucion_id: Identificador único para esta ejecución.
                     Si es None, se generará automáticamente.
        """
        # Generar ID de ejecución si no se proporciona
        self.ejecucion_id = ejecucion_id or datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Definir directorio base para logs
        self.base_dir = os.path.join(os.getcwd(), "LOGS")
        
        # Crear subdirectorios
        self.ejecuciones_dir = os.path.join(self.base_dir, "ejecuciones")
        self.correos_dir = os.path.join(self.base_dir, "correos")
        
        # Crear directorios si no existen
        for dir_path in [self.base_dir, self.ejecuciones_dir, self.correos_dir]:
            os.makedirs(dir_path, exist_ok=True)
        
        # Logger de ejecución global
        self.global_logger = self._setup_global_logger()
        
        # Añadir ejecucion_logger como alias de global_logger para mantener
        # compatibilidad con código existente que lo usa
        self.ejecucion_logger = self.global_logger
    
    def _setup_global_logger(self):
        """Configura el logger global para todo el sistema."""
        logger = logging.getLogger('global')
        
        if not logger.handlers:
            # Archivo de log global
            log_file = os.path.join(self.base_dir, "procesamiento_completo.log")
            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            
            # Salida a consola
            console = logging.StreamHandler()
            console.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            
            # Configurar logger
            logger.setLevel(logging.INFO)
            logger.addHandler(file_handler)
            logger.addHandler(console)
            
            logger.info("Sistema de logs inicializado")
        
        return logger
    
    def get_correo_logger(self, asunto):
        """
        Obtiene un logger para un correo específico.
        Todo se registra en un único archivo de log por correo.
        
        Args:
            asunto: Asunto del correo
            
        Returns:
            logger: Logger configurado para este correo
        """
        # Limpiar asunto para nombre de archivo
        asunto_limpio = self.limpiar_nombre_archivo(asunto)
        
        # Timestamp específico para este correo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        correo_id = f"{asunto_limpio}_{timestamp}"
        
        # Crear nombre de logger único
        logger_name = f"correo_{correo_id}"
        
        # Verificar si el logger ya existe
        if logger_name in logging.root.manager.loggerDict:
            return logging.getLogger(logger_name)
        
        # Crear nuevo logger
        logger = logging.getLogger(logger_name)
        logger.setLevel(logging.INFO)
        logger.propagate = False  # No propagar a logger padre
        
        # Archivo de log único para este correo
        log_file = os.path.join(self.correos_dir, f"{correo_id}.log")
        file_handler = logging.FileHandler(log_file)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        
        # También lo mandamos a consola
        console = logging.StreamHandler()
        console.setFormatter(formatter)
        logger.addHandler(console)
        
        # Guardar el ID del correo para usarlo en otras funciones
        self.correo_id = correo_id
        
        # Registrar inicio
        logger.info(f"=== INICIO PROCESAMIENTO CORREO: {asunto} ===")
        logger.info(f"Fecha de procesamiento: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"ID de ejecución: {self.ejecucion_id}")
        
        # También registrar en el log de ejecución
        self.global_logger.info(f"Iniciado procesamiento del correo: {asunto}")
        
        return logger
    
    def get_logger_for_email(self, asunto, etapa):
        """
        Método de compatibilidad para mantener el funcionamiento con el código existente.
        Redirige a get_correo_logger y añade información de la etapa.
        
        Args:
            asunto: Asunto o identificador
            etapa: Nombre de la etapa
            
        Returns:
            logger: Configurado para este asunto/etapa
        """
        logger = self.get_correo_logger(asunto)
        logger.info(f"=== INICIANDO ETAPA: {etapa} ===")
        return logger
    
    def registrar_etapa(self, logger, etapa, mensaje=None):
        """
        Registra el inicio de una etapa en el log del correo.
        
        Args:
            logger: Logger del correo
            etapa: Nombre de la etapa (inbox, descompresion, etc.)
            mensaje: Mensaje adicional (opcional)
        """
        etapa_norm = etapa.lower().replace(".", "").strip()
        
        # Registrar la etapa con formato destacado
        logger.info(f"{'=' * 10} ETAPA: {etapa.upper()} {'=' * 10}")
        
        if mensaje:
            logger.info(mensaje)
        
        # También registrar en el log de ejecución
        self.global_logger.info(f"Iniciando etapa {etapa} para correo {self.correo_id if hasattr(self, 'correo_id') else 'desconocido'}")
    
    def registrar_resultado(self, logger, etapa, resultado, detalles=None):
        """
        Registra el resultado de una etapa en el log del correo.
        
        Args:
            logger: Logger del correo
            etapa: Nombre de la etapa completada
            resultado: Resultado (éxito, error, etc.)
            detalles: Detalles adicionales (opcional)
        """
        # Registrar resultado
        logger.info(f"{'=' * 10} RESULTADO {etapa.upper()}: {resultado} {'=' * 10}")
        
        if detalles:
            logger.info(f"Detalles: {detalles}")
        
        # También registrar en el log de ejecución
        self.global_logger.info(f"Etapa {etapa} para correo {self.correo_id if hasattr(self, 'correo_id') else 'desconocido'}: {resultado}")
    
    def registrar_salida_script(self, logger, nombre_script, stdout, stderr):
        """
        Registra la salida de un script en el log del correo.
        
        Args:
            logger: Logger del correo
            nombre_script: Nombre del script ejecutado
            stdout: Salida estándar del script
            stderr: Salida de error del script
        """
        # Limpiar nombre del script
        script_name = nombre_script.replace('.py', '')
        
        # Registrar en el log del correo
        logger.info(f"{'=' * 10} SALIDA DE {script_name.upper()} {'=' * 10}")
        
        # Registrar stdout (solo las partes importantes)
        if stdout:
            # Filtrar líneas relevantes para no saturar el log
            lineas_importantes = []
            for linea in stdout.splitlines():
                # Incluir líneas que contienen información importante
                if ("ERROR" in linea or "AVISO" in linea or "INFO" in linea or 
                    "✓" in linea or "❌" in linea or "DETECCIÓN" in linea):
                    lineas_importantes.append(linea)
            
            # Si hay muchas líneas, mostrar solo un resumen
            if len(lineas_importantes) > 20:
                logger.info(f"Salida (mostrando {min(20, len(lineas_importantes))} líneas de {len(lineas_importantes)}):")
                for linea in lineas_importantes[:20]:
                    logger.info(f"  | {linea}")
                logger.info(f"  | ... {len(lineas_importantes) - 20} líneas más ...")
            else:
                logger.info("Salida:")
                for linea in lineas_importantes:
                    logger.info(f"  | {linea}")
        
        # Registrar stderr (siempre es importante)
        if stderr:
            logger.warning("Errores:")
            for linea in stderr.splitlines():
                logger.warning(f"  ! {linea}")
        
        # También registrar en el log de ejecución
        self.global_logger.info(f"Registradas salidas del script {nombre_script} para correo {self.correo_id if hasattr(self, 'correo_id') else 'desconocido'}")
    
    def finalizar_correo(self, logger, resultado_final, detalles=None):
        """
        Finaliza el log de un correo con un resumen.
        
        Args:
            logger: Logger del correo
            resultado_final: Resultado final del procesamiento
            detalles: Detalles adicionales (opcional)
        """
        logger.info(f"{'=' * 20}")
        logger.info(f"RESULTADO FINAL: {resultado_final}")
        
        if detalles:
            if isinstance(detalles, dict):
                for clave, valor in detalles.items():
                    logger.info(f"{clave}: {valor}")
            else:
                logger.info(f"Detalles: {detalles}")
        
        logger.info(f"Fecha de finalización: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"{'=' * 20}")
        
        # También registrar en el log de ejecución
        self.global_logger.info(f"Finalizado procesamiento del correo {self.correo_id if hasattr(self, 'correo_id') else 'desconocido'}: {resultado_final}")
    
    def limpiar_nombre_archivo(self, texto):
        """Limpia un texto para usarlo como nombre de archivo."""
        if not texto:
            return "sin_asunto"
        # Eliminar caracteres no permitidos
        texto_limpio = re.sub(r'[\\/*?:"<>|]', "_", texto)
        # Limitar longitud
        return texto_limpio.strip()[:50]