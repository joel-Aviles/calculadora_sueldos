import os
import time
import logging

logging.basicConfig(
    filename='calculadora.log',  # Nombre del archivo de log
    level=logging.INFO,      # Nivel de log mínimo (INFO en adelante)
    format='%(asctime)s - %(levelname)s - %(message)s'  # Formato del mensaje
)

BASE_DIRECTORY = "C:\\caluladora\\pensiones"  # Ruta donde guardarás los archivos procesados.
TIME_LIMIT = 24 * 60 * 60  # 24 horas en segundos

def is_valid_file_path(file_path: str) -> bool:
    # Evitar rutas relativas o manipulación de la ruta
    full_path = os.path.abspath(os.path.join(BASE_DIRECTORY, file_path))

    # Verificar que el archivo esté dentro del directorio base y no se haya intentado acceder a fuera
    if not full_path.startswith(BASE_DIRECTORY):
        return False
    
    # Comprobar que el archivo existe y es un archivo regular
    return os.path.exists(full_path) and os.path.isfile(full_path)

def clean_old_files(directory: str):
    logging.info("clean_old_files activado")
    # Verificar si la ruta es válida
    if not os.path.isdir(directory):
        logging.info(f"Directorio no válido: {directory}")
        return

    # Obtener el tiempo actual
    current_time = time.time()
    
    # Iterar sobre los archivos del directorio
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)

        # Verificar si es un archivo y no una carpeta
        if os.path.isfile(file_path):
            # Verificar la fecha de última modificación del archivo
            file_mod_time = os.path.getmtime(file_path)

            # Imprimir la fecha y hora de la última modificación
            logging.info(f"{filename} modificado en: {time.ctime(file_mod_time)}")

            # Si el archivo es más antiguo que el tiempo límite, eliminarlo
            if current_time - file_mod_time > TIME_LIMIT:
                try:
                    os.remove(file_path)
                    logging.info(f"Archivo eliminado: {file_path}")
                except Exception as e:
                    logging.error(f"Error al intentar eliminar el archivo {file_path}: {str(e)}")
    